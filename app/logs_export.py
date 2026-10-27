"""Admin-only downloadable XLSX exports: the full activity log and the user
accounts list."""
from __future__ import annotations

import io

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .database import uae_time

HEADER_FILL = PatternFill(start_color="1F2937", end_color="1F2937", fill_type="solid")
HEADER_FONT = Font(color="FFFFFF", bold=True)
WRAP = Alignment(wrap_text=True, vertical="top")

# (column title, width, record key)
LOG_COLUMNS = [("Time (UAE)", 22, "time"), ("Username", 20, "username"), ("Name", 28, "name"), ("Role", 16, "role"),
               ("Action", 30, "action"), ("Details", 60, "details"), ("IP address", 16, "ip_address")]
# Passwords are deliberately left out, as on the Admin page.
USER_COLUMNS = [("Username", 22, "username"), ("Authorization", 16, "authorization"), ("Name", 32, "name"),
                ("Job", 32, "job"), ("Location", 20, "location"), ("Status", 14, "status"), ("Employees", 40, "employees")]


def _build_workbook(sheet_title: str, columns: list[tuple[str, int, str]], records: list[dict]) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_title

    for c, (title, width, _) in enumerate(columns, start=1):
        ws.column_dimensions[get_column_letter(c)].width = width
        cell = ws.cell(row=1, column=c, value=title)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = WRAP
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{max(len(records) + 1, 1)}"

    for r, record in enumerate(records, start=2):
        for c, (_, _, key) in enumerate(columns, start=1):
            cell = ws.cell(row=r, column=c, value=record.get(key))
            cell.alignment = WRAP

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def build_logs_workbook(logs: list[dict]) -> bytes:
    return _build_workbook("Activity log", LOG_COLUMNS, [{**log, "time": uae_time(log["time"])} for log in logs])


ASSIGNMENT_COLUMNS = [("Concat_Frame ID", 80, "frame_ids"), ("Auditor Name", 36, "name"),
                      ("Auditor_User_Name", 22, "username"), ("Count", 10, "count")]


def build_assignments_workbook(auditors: list[dict]) -> bytes:
    """A survey's auditor list in the layout it is uploaded in, so it can be
    edited and uploaded again: one row per auditor, Frame IDs comma-separated."""
    rows = [{"frame_ids": ",".join(sorted(a["frames"])), "name": a["name"], "username": a["username"],
             "count": len(a["frames"])} for a in auditors]
    return _build_workbook("Auditor assignments", ASSIGNMENT_COLUMNS, rows)


def build_users_workbook(users: list[dict]) -> bytes:
    return _build_workbook("Users", USER_COLUMNS, users)
