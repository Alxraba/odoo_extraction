#!/usr/bin/env python3
"""
Extraction Odoo + upload SharePoint — version GitHub Actions.

Différences vs la version Mac :
  - Plus de Trousseau : tous les secrets viennent de variables d'environnement
  - Plus de fallback OneDrive : upload direct vers SharePoint via Microsoft Graph
"""

import io
import os
import sys
import xmlrpc.client

import msal
import requests
from openpyxl import Workbook

# ─── SECRETS depuis l'environnement (GitHub Secrets) ──────────────
ODOO_URL     = os.environ["ODOO_URL"]
ODOO_DB      = os.environ["ODOO_DB"]
ODOO_USER    = os.environ["ODOO_USER"]
ODOO_API_KEY = os.environ["ODOO_API_KEY"]

AZ_TENANT = os.environ["AZURE_TENANT_ID"]
AZ_CLIENT = os.environ["AZURE_CLIENT_ID"]
AZ_SECRET = os.environ["AZURE_CLIENT_SECRET"]

SP_HOST   = os.environ["SP_SITE_HOSTNAME"]   # ex : carb0n642.sharepoint.com
SP_PATH   = os.environ["SP_SITE_PATH"]       # ex : /sites/TemplateTools
SP_DRIVE  = os.environ.get("SP_DRIVE_NAME", "Documents")  # nom bibliothèque
SP_FOLDER = os.environ["SP_FOLDER_PATH"]     # ex : /1_Extracts_Odoo

# ─── Templates à extraire (idem version Mac) ──────────────────────
EXPORTS = [
    ("Workload project 1", "project.project", [], "Project.xlsx"),
    ("Workload 3",         "sale.order",      [], "Sales_Order.xlsx"),
    ("Workload 7",         "project.task",    [], "Task.xlsx"),
]


# ═════════════════════════════════════════════════════════════════
# ODOO
# ═════════════════════════════════════════════════════════════════
def odoo_connect():
    common = xmlrpc.client.ServerProxy(f"{ODOO_URL}/xmlrpc/2/common")
    uid = common.authenticate(ODOO_DB, ODOO_USER, ODOO_API_KEY, {})
    if not uid:
        sys.exit("❌ Auth Odoo échouée")
    return uid, xmlrpc.client.ServerProxy(f"{ODOO_URL}/xmlrpc/2/object")


def odoo_call(models, uid, model, method, args, kwargs=None):
    return models.execute_kw(ODOO_DB, uid, ODOO_API_KEY, model, method, args, kwargs or {})


_fields_cache = {}
def get_model_fields(models, uid, model_name):
    if model_name not in _fields_cache:
        _fields_cache[model_name] = odoo_call(
            models, uid, model_name, "fields_get",
            [], {"attributes": ["string", "relation"]},
        )
    return _fields_cache[model_name]


def resolve_label(models, uid, model_name, field_path):
    parts, labels, current = field_path.split("/"), [], model_name
    for p in parts:
        info = get_model_fields(models, uid, current).get(p)
        if not info:
            labels.append(p); break
        labels.append(info.get("string") or p)
        if info.get("relation"):
            current = info["relation"]
        else:
            break
    return "/".join(labels)


def get_export_fields(models, uid, template_name, model_name):
    ids = odoo_call(models, uid, "ir.exports", "search",
                    [[("name", "=ilike", template_name),
                      ("resource", "=", model_name)]])
    if not ids:
        ids = odoo_call(models, uid, "ir.exports", "search",
                        [[("name", "=ilike", template_name)]])
    if not ids:
        raise RuntimeError(f"Template introuvable : {template_name}")
    lines = odoo_call(models, uid, "ir.exports.line", "search_read",
                      [[("export_id", "=", ids[0])]],
                      {"fields": ["name"], "order": "id"})
    return [l["name"] for l in lines]


def extract_rows(models, uid, model_name, domain, fields):
    ids = odoo_call(models, uid, model_name, "search", [domain])
    if not ids:
        return []
    return odoo_call(models, uid, model_name, "export_data",
                     [ids, fields]).get("datas", [])


def make_xlsx_bytes(headers, rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(headers)
    for r in rows:
        ws.append(["" if v is False else v for v in r])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ═════════════════════════════════════════════════════════════════
# MICROSOFT GRAPH
# ═════════════════════════════════════════════════════════════════
def graph_token():
    app = msal.ConfidentialClientApplication(
        AZ_CLIENT,
        authority=f"https://login.microsoftonline.com/{AZ_TENANT}",
        client_credential=AZ_SECRET,
    )
    res = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in res:
        sys.exit(f"❌ Auth Azure échouée : {res.get('error_description', res)}")
    return res["access_token"]


def find_drive_id(token):
    h = {"Authorization": f"Bearer {token}"}

    # 1) Site ID
    r = requests.get(f"https://graph.microsoft.com/v1.0/sites/{SP_HOST}:{SP_PATH}",
                     headers=h, timeout=30)
    r.raise_for_status()
    site_id = r.json()["id"]

    # 2) Drive ID (bibliothèque de documents)
    r = requests.get(f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives",
                     headers=h, timeout=30)
    r.raise_for_status()
    drives = r.json()["value"]
    for d in drives:
        if d["name"] == SP_DRIVE:
            return d["id"]
    sys.exit(
        f"❌ Bibliothèque '{SP_DRIVE}' introuvable. "
        f"Disponibles : {[d['name'] for d in drives]}"
    )


def upload_to_sharepoint(token, drive_id, filename, content):
    """Simple PUT (fichiers < 4 MB — largement suffisant ici)."""
    folder = SP_FOLDER.strip("/")
    url = (
        f"https://graph.microsoft.com/v1.0/drives/{drive_id}/"
        f"root:/{folder}/{filename}:/content"
    )
    r = requests.put(
        url,
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": (
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
        },
        data=content,
        timeout=60,
    )
    r.raise_for_status()
    return r.json().get("webUrl", "(uploaded)")


# ═════════════════════════════════════════════════════════════════
# MAIN
# ═════════════════════════════════════════════════════════════════
def main():
    uid, models = odoo_connect()
    print(f"✓ Odoo connecté (UID={uid})")

    token = graph_token()
    drive_id = find_drive_id(token)
    print(f"✓ SharePoint connecté (drive_id={drive_id[:20]}...)")

    had_error = False
    for template, model, domain, filename in EXPORTS:
        print(f"\n→ {filename}")
        try:
            fields  = get_export_fields(models, uid, template, model)
            headers = [resolve_label(models, uid, model, f) for f in fields]
            rows    = extract_rows(models, uid, model, domain, fields)
            xlsx    = make_xlsx_bytes(headers, rows)
            upload_to_sharepoint(token, drive_id, filename, xlsx)
            print(f"  ✓ {len(rows)} lignes uploadées ({len(xlsx) // 1024} KB)")
        except Exception as e:
            had_error = True
            print(f"  ✗ Erreur : {e}", file=sys.stderr)

    sys.exit(1 if had_error else 0)


if __name__ == "__main__":
    main()
