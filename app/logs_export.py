"""Admin-only downloadable XLSX export of the full activity log."""
from __future__ import annotations

import io

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

HEADER_FILL = PatternFill(start_color="1F2937", end_color="1F2937", fill_type="solid")
HEADER_FONT = Font(color="FFFFFF", bold=True)
WRAP = Alignment(wrap_text=True, vertical="top")

COLUMNS = ["Time (UTC)", "Username", "Role", "Action", "Details", "IP address"]
COLUMN_WIDTHS = [22, 20, 16, 30, 60, 16]


def build_logs_workbook(logs: list[dict]) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Activity log"

    for i, width in enumerate(COLUMN_WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = width

    for c, title in enumerate(COLUMNS, start=1):
        cell = ws.cell(row=1, column=c, value=title)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = WRAP
    ws.freeze_panes = "A2"

    for r, log in enumerate(logs, start=2):
        values = [log.get("time"), log.get("username"), log.get("role"),
                  log.get("action"), log.get("details"), log.get("ip_address")]
        for c, v in enumerate(values, start=1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.alignment = WRAP

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
