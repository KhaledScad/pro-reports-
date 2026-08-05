"""
Excel ingestion. Ported from the original tool's `v17ReadWorkbook` /
`v17WorksheetToRows` logic: detect the real header row even if the sheet
has title/notes rows above it, by scanning for a Frame ID-looking cell
(survey files) or a Username/Password pair (Account file).
"""
from __future__ import annotations

import io
from openpyxl import load_workbook

from .calc_common import normalize_header

FRAME_ID_TOKENS = {
    "frameid", "frame", "id", "establishmentid", "establishmentnumber",
    "رقمالاطار", "الاطار", "رقمالمنشاة", "رقمالمنشأة",
}


def _is_frame_header_cell(value) -> bool:
    n = normalize_header(value)
    if not n:
        return False
    if n in FRAME_ID_TOKENS:
        return True
    return "frameid" in n or "رقمالاطار" in n


def _unique_headers(headers: list) -> list[str]:
    seen: dict[str, int] = {}
    out = []
    for i, h in enumerate(headers):
        base = str(h).strip() if h not in (None, "") else f"Column_{i + 1}"
        if base not in seen:
            seen[base] = 1
            out.append(base)
        else:
            seen[base] += 1
            out.append(f"{base}_{seen[base]}")
    return out


def _find_header_row(grid: list, is_header_cell) -> int | None:
    for i, row in enumerate(grid):
        if any(is_header_cell(c) for c in row):
            return i
    for i, row in enumerate(grid):
        non_empty = sum(1 for c in row if c is not None and str(c).strip() != "")
        if non_empty >= 2:
            return i
    return None


def _sheet_to_rows(ws, is_header_cell=_is_frame_header_cell) -> list[dict]:
    grid = list(ws.iter_rows(values_only=True))
    if not grid:
        return []
    header_index = _find_header_row(grid, is_header_cell)
    if header_index is None:
        return []
    headers = _unique_headers(list(grid[header_index]))
    rows = []
    for line in grid[header_index + 1:]:
        if not any(c is not None and str(c).strip() != "" for c in line):
            continue
        obj = {headers[c]: ("" if (c >= len(line) or line[c] is None) else line[c]) for c in range(len(headers))}
        rows.append(obj)
    return rows


def sheet_names(file_bytes: bytes) -> list[str]:
    wb = load_workbook(io.BytesIO(file_bytes), read_only=True)
    names = list(wb.sheetnames)
    wb.close()
    return names


def read_workbook(file_bytes: bytes, only_sheet: str | None = None) -> dict[str, list[dict]]:
    """Return {sheet_name: [row_dict, ...]}. If `only_sheet` is given (and
    not the sentinel "__ALL__"), restrict to that one sheet."""
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)
    names = [only_sheet] if only_sheet and only_sheet != "__ALL__" else wb.sheetnames
    sheets = {name: _sheet_to_rows(wb[name]) for name in names if name in wb.sheetnames}
    wb.close()
    return sheets


def get_frame_id(row: dict) -> str:
    for key, value in row.items():
        if _is_frame_header_cell(key):
            return str(value).strip()
    if row:
        first_key = next(iter(row))
        return str(row[first_key]).strip()
    return ""


def rows_for_survey(file_bytes: bytes, sheet_name: str | None = None) -> list[dict]:
    """Flatten sheet(s) into a list of {frame_id, data} dicts, skipping rows
    without a Frame ID."""
    sheets = read_workbook(file_bytes, sheet_name)
    out = []
    for sname, rows in sheets.items():
        for row in rows:
            if not row:
                continue
            frame_id = get_frame_id(row)
            if not frame_id:
                continue
            data = dict(row)
            data["__sheet"] = sname
            out.append({"frame_id": frame_id, "data": data})
    return out


def _is_account_header_cell(value) -> bool:
    n = normalize_header(value)
    return n in ("username", "password")


def parse_account_sheet(file_bytes: bytes, sheet_name: str | None = None) -> list[dict]:
    """Parse the Account sheet used to seed users. If `sheet_name` is not
    given, uses a sheet literally named 'Account' if present, otherwise the
    first sheet that contains Username/Password columns anywhere in it."""
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)
    try:
        if sheet_name and sheet_name != "__ALL__":
            candidates = [sheet_name] if sheet_name in wb.sheetnames else []
        else:
            named = next((n for n in wb.sheetnames if n.strip().lower() == "account"), None)
            candidates = [named] + list(wb.sheetnames) if named else list(wb.sheetnames)

        rows = []
        for name in candidates:
            if name is None:
                continue
            rows = _sheet_to_rows(wb[name], _is_account_header_cell)
            if rows and any(normalize_header(k) in ("username", "password") for k in rows[0]):
                break
        else:
            rows = []
    finally:
        wb.close()

    if not rows:
        raise ValueError("Could not find Username/Password columns in this file. "
                          "Use the sheet-selection option to pick the correct sheet.")

    users = []
    for row in rows:
        norm = {normalize_header(k): v for k, v in row.items()}
        username = str(norm.get("username", "") or "").strip()
        password = str(norm.get("password", "") or "").strip()
        role_raw = str(norm.get("authorization", "") or "").strip().upper()
        role = "SUPERVISOR" if role_raw == "SUPERVISOR" else ("ADMIN" if role_raw == "ADMIN" else "USER")
        name = str(norm.get("name", "") or username).strip()
        job = str(norm.get("job", "") or "").strip()
        location = str(norm.get("location", "") or "").strip()
        if username and password:
            users.append({
                "username": username, "password": password, "authorization": role,
                "name": name, "job": job, "location": location,
            })
    return users
