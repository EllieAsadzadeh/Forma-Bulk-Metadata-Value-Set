"""
Push metadata from the Files sheet to Autodesk Construction Cloud (ACC) / Forma
custom attributes on the matching file version only (not every version of that file).

Requires an APS (Autodesk Platform Services) app with data:read + data:write scopes,
added as a project member with document-edit permission.

Credentials: put APS_CLIENT_ID and APS_CLIENT_SECRET in a local ".env" file next to this
script. They are loaded automatically at startup and never leave your machine.
"""
import os
import re
import time
from urllib.parse import quote

import openpyxl
import requests
from dotenv import load_dotenv

load_dotenv()

# Data region the ACC/BIM 360 project's data lives in.
# Set APS_REGION in .env if it's anything other than US: US | EMEA | AUS | CAN | DEU | GBR | IND | JPN
REGION = os.environ.get("APS_REGION", "US").upper()

EXCEL_PATH = "Files_log_report-202609301420_updated.xlsx"
SHEET = "Files"

# Extra Excel header aliases -> ACC custom attribute name.
# The Files log "Description" column is the text for ACC "File Description".
# When this alias is present it wins over the "File Description" column, including
# a blank cell (blank Description clears File Description in Forma).
ATTRIBUTE_ALIASES = {
    "Description": "File Description",
}

AUTH_URL = "https://developer.api.autodesk.com/authentication/v2/token"
DM_BASE = "https://developer.api.autodesk.com/data/v1"
PROJECT_BASE = "https://developer.api.autodesk.com/project/v1"
DOCS_BASE = "https://developer.api.autodesk.com/bim360/docs/v1"

UUID_RE = re.compile(
    r"^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$"
)

_session = requests.Session()
_hub_by_project = {}
_folder_cache = {}  # (dm_proj, folder_id) -> list of contents data
_item_cache = {}  # (dm_proj, folder_path) -> (folder_id, {file_name: item})
_attr_defs_cache = {}  # (docs_project_id, folder_id) -> {name: definition}


def get_token():
    client_id = os.environ.get("APS_CLIENT_ID")
    client_secret = os.environ.get("APS_CLIENT_SECRET")
    if not client_id or not client_secret:
        raise SystemExit(
            "Missing APS_CLIENT_ID / APS_CLIENT_SECRET.\n"
            "Create a '.env' file next to this script (see .env.example) with:\n"
            "  APS_CLIENT_ID=your_client_id\n"
            "  APS_CLIENT_SECRET=your_client_secret"
        )
    resp = requests.post(
        AUTH_URL,
        data={
            "grant_type": "client_credentials",
            "client_id": client_id,
            "client_secret": client_secret,
            "scope": "data:read data:write",
        },
    )
    resp.raise_for_status()
    return resp.json()["access_token"]


def dm_project_id(bim360_project_id):
    # Data Management API needs the "b." prefix; BIM 360/ACC docs API does not.
    return bim360_project_id if bim360_project_id.startswith("b.") else f"b.{bim360_project_id}"


def docs_project_id(bim360_project_id):
    return bim360_project_id.replace("b.", "", 1) if bim360_project_id else bim360_project_id


def is_uuid(value):
    return bool(value) and bool(UUID_RE.match(str(value).strip()))


def auth_headers(token, json_body=False):
    headers = {"Authorization": f"Bearer {token}"}
    if REGION != "US":
        headers["x-ads-region"] = REGION
    if json_body:
        headers["Content-Type"] = "application/json"
    return headers


def api_request(method, url, token, **kwargs):
    """GET/POST with region header, pagination-friendly, and light retries."""
    headers = kwargs.pop("headers", None) or auth_headers(token, json_body=method.upper() == "POST")
    last_error = None
    for attempt in range(5):
        resp = _session.request(method, url, headers=headers, timeout=60, **kwargs)
        if resp.status_code in (429, 500, 502, 503, 504):
            last_error = resp
            time.sleep(1.5 * (attempt + 1))
            continue
        return resp
    return last_error


def get_json(method, url, token, **kwargs):
    resp = api_request(method, url, token, **kwargs)
    resp.raise_for_status()
    if not resp.content:
        return {}
    return resp.json()


def paginate_data(url, token, params=None):
    """Yield JSON:API `data` items following `links.next`."""
    params = dict(params or {})
    params.setdefault("page[limit]", 200)
    while url:
        body = get_json("GET", url, token, params=params)
        for item in body.get("data", []):
            yield item
        next_link = (body.get("links") or {}).get("next") or {}
        url = next_link.get("href") if isinstance(next_link, dict) else None
        params = None  # next href already includes query params


def parse_version_number(label):
    """Turn Excel 'V1' / 'v1' / 1 into the integer ACC versionNumber."""
    if label is None or label == "":
        return None
    text = str(label).strip().upper()
    if text.startswith("V"):
        text = text[1:]
    try:
        return int(text)
    except ValueError:
        return None


def version_number_from_urn(version_id):
    if not version_id or "?version=" not in str(version_id):
        return None
    try:
        return int(str(version_id).rsplit("?version=", 1)[-1])
    except ValueError:
        return None


def normalize_project_ids(ws, header, project_col):
    """
    Excel fill-handle often increments a pasted GUID (…c3f9, …c3f10, …).
    ACC project IDs are the same UUID on every row of a single-project export.
    """
    valid = []
    invalid_rows = []
    for row in range(2, ws.max_row + 1):
        raw = ws.cell(row=row, column=project_col).value
        if not raw or "PASTE-PROJECT-ID" in str(raw):
            continue
        pid = str(raw).strip()
        if is_uuid(pid):
            valid.append(pid)
        else:
            invalid_rows.append((row, pid))

    if not invalid_rows:
        return

    canonical = valid[0] if valid else None
    if not canonical:
        print("[warn] no valid Project ID UUID found; leaving cells as-is")
        return

    unique_valid = sorted(set(valid))
    if len(unique_valid) > 1:
        print(f"[warn] multiple valid Project IDs in the sheet: {unique_valid}")
        return

    print(
        f"[fix] {len(invalid_rows)} Project ID(s) look auto-filled; "
        f"using {canonical} for every row"
    )
    for row in range(2, ws.max_row + 1):
        name = ws.cell(row=row, column=header["Name"]).value
        if not name:
            continue
        ws.cell(row=row, column=project_col, value=canonical)


def list_hubs(token):
    params = {} if REGION == "US" else {"region": REGION}
    return list(paginate_data(f"{PROJECT_BASE}/hubs", token, params=params))


def find_hub_id(token, dm_proj):
    if dm_proj in _hub_by_project:
        return _hub_by_project[dm_proj]
    for hub in list_hubs(token):
        hub_id = hub["id"]
        url = f"{PROJECT_BASE}/hubs/{hub_id}/projects"
        for project in paginate_data(url, token):
            if project["id"] == dm_proj:
                _hub_by_project[dm_proj] = hub_id
                return hub_id
    raise RuntimeError(f"Project {dm_proj} was not found in any accessible hub")


def folder_contents(token, dm_proj, folder_id):
    key = (dm_proj, folder_id)
    if key not in _folder_cache:
        url = f"{DM_BASE}/projects/{dm_proj}/folders/{quote(folder_id, safe='')}/contents"
        _folder_cache[key] = list(paginate_data(url, token))
    return _folder_cache[key]


def walk_to_folder(token, dm_proj, folder_path):
    """Return the destination folder URN for an ACC folder path, or None."""
    parts = [p for p in str(folder_path).replace("\\", "/").split("/") if p]
    if not parts:
        return None

    hub_id = find_hub_id(token, dm_proj)
    top = get_json(
        "GET",
        f"{PROJECT_BASE}/hubs/{hub_id}/projects/{dm_proj}/topFolders",
        token,
    )
    folders = {f["attributes"]["name"]: f["id"] for f in top.get("data", [])}
    if parts[0] not in folders:
        return None

    current_id = folders[parts[0]]
    for part in parts[1:]:
        match = next(
            (
                d
                for d in folder_contents(token, dm_proj, current_id)
                if d["type"] == "folders" and d["attributes"].get("displayName") == part
            ),
            None,
        )
        if not match:
            return None
        current_id = match["id"]
    return current_id


def items_in_folder_path(token, dm_proj, folder_path):
    key = (dm_proj, folder_path)
    if key not in _item_cache:
        folder_id = walk_to_folder(token, dm_proj, folder_path)
        if not folder_id:
            _item_cache[key] = (None, {})
        else:
            items = {
                d["attributes"].get("displayName"): d
                for d in folder_contents(token, dm_proj, folder_id)
                if d["type"] == "items"
            }
            _item_cache[key] = (folder_id, items)
    return _item_cache[key]


def find_version(token, dm_proj, item, wanted_number):
    """Return the version URN whose versionNumber matches the Excel Version number."""
    if wanted_number is None:
        return None

    tip_id = (item.get("relationships") or {}).get("tip", {}).get("data", {}).get("id")
    if version_number_from_urn(tip_id) == wanted_number:
        return tip_id

    item_id = item["id"]
    url = (
        f"{DM_BASE}/projects/{dm_proj}/items/{quote(item_id, safe='')}/versions"
        f"?filter[versionNumber]={wanted_number}"
    )
    versions = list(paginate_data(url, token))
    match = next(
        (
            v
            for v in versions
            if v.get("attributes", {}).get("versionNumber") == wanted_number
        ),
        None,
    )
    if match:
        return match["id"]

    # Fallback: revision display label (Docs UI "V1") can differ from versionNumber.
    url = f"{DM_BASE}/projects/{dm_proj}/items/{quote(item_id, safe='')}/versions"
    for version in paginate_data(url, token):
        attrs = version.get("attributes") or {}
        label = ((attrs.get("extension") or {}).get("data") or {}).get("revisionDisplayLabel")
        if parse_version_number(label) == wanted_number:
            return version["id"]
        if parse_version_number(attrs.get("displayName")) == wanted_number:
            return version["id"]
    return None


def get_attribute_definitions(token, project_id, folder_id):
    """GET folder custom-attribute-definitions -> {name: definition}."""
    key = (project_id, folder_id)
    if key not in _attr_defs_cache:
        url = (
            f"{DOCS_BASE}/projects/{project_id}/folders/"
            f"{quote(folder_id, safe='')}/custom-attribute-definitions"
        )
        body = get_json("GET", url, token)
        results = body.get("results", body if isinstance(body, list) else [])
        _attr_defs_cache[key] = {d["name"]: d for d in results}
    return _attr_defs_cache[key]


def is_blank(value):
    if value is None:
        return True
    if isinstance(value, str) and value.strip() == "":
        return True
    return False


def values_from_row(ws, row, header, attr_defs):
    """
    Read every Excel cell that corresponds to a Forma custom attribute.
    Blank cells are kept as None so they can clear the value in Forma.
    """
    values = {}
    for attr_name in attr_defs:
        if attr_name in header:
            values[attr_name] = ws.cell(row=row, column=header[attr_name]).value
    for excel_col, attr_name in ATTRIBUTE_ALIASES.items():
        if excel_col in header and attr_name in attr_defs:
            values[attr_name] = ws.cell(row=row, column=header[excel_col]).value
    return values


def coerce_attribute_value(definition, value):
    if is_blank(value):
        return None, None  # JSON null clears the Forma value
    attr_type = definition.get("type")
    if attr_type == "array":
        text = str(value).strip()
        allowed = definition.get("arrayValues") or []
        if text not in allowed:
            return None, f"{definition['name']} value {text!r} is not an allowed option"
        return text, None
    if isinstance(value, str):
        return value.strip(), None
    return value, None


def push_attributes(token, project_id, version_id, attr_defs, values):
    payload = []
    warnings = []
    cleared = 0
    for name, value in values.items():
        if name not in attr_defs:
            continue
        coerced, warn = coerce_attribute_value(attr_defs[name], value)
        if warn:
            warnings.append(warn)
            continue
        if is_blank(value):
            payload.append({"id": attr_defs[name]["id"], "value": None})
            cleared += 1
        else:
            payload.append({"id": attr_defs[name]["id"], "value": coerced})

    if not payload:
        note = "nothing to push"
        if warnings:
            note += " (" + "; ".join(warnings) + ")"
        return True, note

    url = (
        f"{DOCS_BASE}/projects/{project_id}/versions/"
        f"{quote(version_id, safe='')}/custom-attributes:batch-update"
    )
    resp = api_request("POST", url, token, json=payload)
    if not resp.ok:
        return False, resp.text
    set_count = len(payload) - cleared
    note = f"pushed ({set_count} set, {cleared} cleared)"
    if warnings:
        note += " | " + "; ".join(warnings)
    return True, note


def main():
    token = get_token()
    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws = wb[SHEET]
    header = {cell.value: cell.column for cell in ws[1] if cell.value}

    project_col = header["Project ID"]
    item_col = header["Item ID (URN)"]
    version_col = header["Version ID (URN)"]
    folder_col = header["Folder name and path"]
    name_col = header["Name"]
    version_number_col = header.get("Version number")

    normalize_project_ids(ws, header, project_col)

    counts = {"ok": 0, "fail": 0, "skip": 0}

    for row in range(2, ws.max_row + 1):
        file_name = ws.cell(row=row, column=name_col).value
        if not file_name:
            continue

        bim360_project_id = ws.cell(row=row, column=project_col).value
        if not bim360_project_id or "PASTE-PROJECT-ID" in str(bim360_project_id):
            print(f"[skip] row {row}: Project ID not set")
            counts["skip"] += 1
            continue
        bim360_project_id = str(bim360_project_id).strip()
        if not is_uuid(docs_project_id(bim360_project_id)):
            print(f"[skip] row {row}: invalid Project ID {bim360_project_id}")
            counts["skip"] += 1
            continue

        version_label = (
            ws.cell(row=row, column=version_number_col).value if version_number_col else None
        )
        wanted_number = parse_version_number(version_label)
        if wanted_number is None:
            print(
                f"[skip] row {row}: {file_name} has no Version number "
                f"(refusing to push onto every/latest version)"
            )
            counts["skip"] += 1
            continue

        dm_proj = dm_project_id(bim360_project_id)
        docs_pid = docs_project_id(bim360_project_id)
        folder_path = ws.cell(row=row, column=folder_col).value
        cached_item_id = ws.cell(row=row, column=item_col).value
        cached_version_id = ws.cell(row=row, column=version_col).value

        folder_id, items = items_in_folder_path(token, dm_proj, folder_path)
        item = items.get(file_name)
        if not item:
            print(f"[not found] row {row}: {folder_path}/{file_name}")
            counts["skip"] += 1
            continue

        item_id = item["id"]
        version_id = None
        if (
            cached_item_id == item_id
            and cached_version_id
            and version_number_from_urn(cached_version_id) == wanted_number
        ):
            version_id = cached_version_id
        else:
            version_id = find_version(token, dm_proj, item, wanted_number)

        if not version_id:
            print(
                f"[not found] row {row}: {file_name} version {version_label} "
                f"(will not fall back to other versions)"
            )
            counts["skip"] += 1
            continue

        ws.cell(row=row, column=item_col, value=item_id)
        ws.cell(row=row, column=version_col, value=version_id)

        attr_defs = get_attribute_definitions(token, docs_pid, folder_id)
        values = values_from_row(ws, row, header, attr_defs)

        ok, info = push_attributes(token, docs_pid, version_id, attr_defs, values)
        status = "ok" if ok else "FAIL"
        counts["ok" if ok else "fail"] += 1
        print(f"[{status}] row {row}: {file_name} {version_label} -> {info}")
        time.sleep(0.2)

    wb.save(EXCEL_PATH)
    print(
        f"done: {counts['ok']} pushed, {counts['fail']} failed, {counts['skip']} skipped"
    )


if __name__ == "__main__":
    main()
