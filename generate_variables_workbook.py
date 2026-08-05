"""
Generates a reference XLSX documenting every variable/column used by the
QES and FISQ calculation + comparison engines in this app, so an admin
knows exactly what to put in the Excel files they upload.

Produces 4 sheets: QES Calculations, QES Comparisons, FISQ Calculations,
FISQ Comparisons.
"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

RED = "D61F2C"
DARK = "222222"
LIGHT_GRAY = "F7F7F8"
SECTION_FILL = "FDE9E9"

HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
HEADER_FILL = PatternFill("solid", fgColor=RED)
SECTION_FONT = Font(bold=True, color=DARK, size=11)
SECTION_FILL_STYLE = PatternFill("solid", fgColor=SECTION_FILL)
BORDER = Border(*(Side(style="thin", color="D9D9D9") for _ in range(4)))
WRAP = Alignment(wrap_text=True, vertical="top")
WRAP_CENTER = Alignment(wrap_text=True, vertical="top", horizontal="center")

COLUMNS = ["Variable / Column name", "Category", "Required?", "Data type",
           "Applies to (business type)", "Description", "Example value"]
WIDTHS = [34, 20, 12, 10, 24, 55, 16]


def write_sheet(ws, title, rows):
    ws.title = title
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A2"

    for c, (header, width) in enumerate(zip(COLUMNS, WIDTHS), start=1):
        cell = ws.cell(row=1, column=c, value=header)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = WRAP_CENTER
        cell.border = BORDER
        ws.column_dimensions[get_column_letter(c)].width = width
    ws.row_dimensions[1].height = 28

    r = 2
    for row in rows:
        if row[0] == "__SECTION__":
            cell = ws.cell(row=r, column=1, value=row[1])
            cell.font = SECTION_FONT
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=len(COLUMNS))
            for c in range(1, len(COLUMNS) + 1):
                ws.cell(row=r, column=c).fill = SECTION_FILL_STYLE
                ws.cell(row=r, column=c).border = BORDER
            ws.row_dimensions[r].height = 20
            r += 1
            continue
        for c, value in enumerate(row, start=1):
            cell = ws.cell(row=r, column=c, value=value)
            cell.border = BORDER
            cell.alignment = WRAP if c in (1, 6) else Alignment(vertical="top", wrap_text=True)
            if r % 2 == 0:
                cell.fill = PatternFill("solid", fgColor=LIGHT_GRAY)
        r += 1
    ws.auto_filter.ref = f"A1:{get_column_letter(len(COLUMNS))}{r - 1}"


def sec(title):
    return ("__SECTION__", title)


# ============================================================ QES CALC ===
qes_calc_rows = [
    sec("Identification & assignment (one value per Frame ID, no quarter prefix)"),
    ["FRAME_ID", "Identification", "Required", "Text", "All", "Unique establishment / project identifier. One row per Frame ID.", "F001"],
    ["ESTABLISHMENT_NAME", "Identification", "Optional", "Text", "All", "Business / establishment name shown on the record.", "Al Noor Trading LLC"],
    ["ACTIVITY_TYPE_EN / ACTIVITY_TYPE_AR", "Identification", "Optional", "Text", "All", "Business activity in English/Arabic. Used with ISIC4 to classify Trade / Industry / Construction / Transport.", "Trade"],
    ["ISIC4", "Identification", "Optional", "Text/Number", "All", "ISIC Rev.4 activity code. First 2 digits are used as a fallback classifier (45-47 Trade, 10-33 Industry, 41-43 Construction, 49-53 Transport).", "4610"],
    ["ISIC4_DESC_EN / ISIC4_DESC_AR", "Identification", "Optional", "Text", "All", "ISIC4 description text, also scanned for classification keywords.", "Wholesale trade"],
    ["AUDITOR_NAME_EN / AUDITOR_NAME_AR", "Assignment", "Optional", "Text", "All", "Auditor's display name.", "Fatima Auditor"],
    ["Auditor_User_Name", "Assignment", "Required for Auditor role", "Text", "All", "Login username of the Auditor. A user whose Job = 'Auditor' can only open records where this matches their username, in any quarter.", "auditor1"],
    ["ENUMERATOR_NAME_EN / AR", "Assignment", "Optional", "Text", "All", "Enumerator / Key Account Manager display name.", "Ahmed Enumerator"],
    ["_ENUMERATOR_User_name", "Assignment", "Required for Enumerator role", "Text", "All", "Login username of the Enumerator / Key Account Manager. Same per-quarter matching as Auditor.", "enum1"],
    ["Supervisor_USER_Name", "Assignment", "Required for Supervisor role", "Text", "All", "Login username of the Supervisor. A user whose Job contains 'Supervisor' only sees records where this matches.", "sup1"],

    sec("Per-quarter financial inputs — prefix each with the quarter, e.g. Q1_2026_Total_Revenue"),
    ["Total_Revenue", "Financial input", "Required", "Number", "All", "Quarter revenue / sales, excluding VAT.", "500000"],
    ["Total_Expenses", "Financial input", "Required", "Number", "All", "Total quarter expenses (used for Profit and as part of Intermediate Consumption).", "120000"],
    ["Operations_Cost", "Financial input", "Required", "Number", "All", "Operating / administrative expenses, part of Intermediate Consumption.", "60000"],
    ["RAW_Material", "Financial input", "Required for Trade & Industry", "Number", "Trade, Industry", "Raw materials used (Industry) or net purchases of goods for resale (Trade).", "30000"],
    ["Total_Finished_Goods_Begining_Of_period", "Financial input", "Required for Trade & Industry", "Number", "Trade, Industry", "Opening inventory value for the quarter.", "20000"],
    ["Total_Finished_Goods_End_of_period", "Financial input", "Required for Trade & Industry", "Number", "Trade, Industry", "Closing inventory value for the quarter. Compared to next quarter's opening value in the comparison page.", "25000"],
    ["IC", "Financial input (override)", "Optional", "Number", "All", "Intermediate Consumption supplied directly in the file. If present it REPLACES the calculated formula below.", "180000"],
    ["TOTAL_EMPLOYEES", "Financial input", "Required for Average Salary", "Number", "All", "Headcount for the quarter.", "12"],
    ["Employee_Salary", "Financial input", "Required for Average Salary", "Number", "All", "Total quarterly employee compensation (used to derive a monthly average).", "180000"],
    ["Comment_Revenue", "Note", "Optional", "Text", "All", "Free-text note on the revenue figure, shown alongside the record.", "Confirmed with client"],
    ["Comment_Expense", "Note", "Optional", "Text", "All", "Free-text note on the expenses figure.", ""],
    ["Comment_Inventory", "Note", "Optional", "Text", "All", "Free-text note on the inventory figures.", ""],

    sec("Calculated outputs (produced by the tool — do not upload these)"),
    ["Business classification", "Calculated", "—", "Text", "All", "Trade / Industry / Construction / Transport / Other, derived from ISIC4 and activity text.", "Trade"],
    ["Gross Output", "Calculated", "—", "Number", "All", "Trade: Revenue + End Inv. − Begin Inv. − Raw Material.  Industry/Other: Revenue + End Inv. − Begin Inv.  Construction/Transport: Revenue.", "511,000"],
    ["Intermediate Consumption (IC)", "Calculated", "—", "Number", "All", "Trade/Construction/Transport: Operations Cost + Expenses.  Industry/Other: Operations Cost + Expenses + Raw Material.  Uses the IC column instead if it was supplied.", "192,000"],
    ["Added Value", "Calculated", "—", "Number", "All", "Gross Output − Intermediate Consumption.", "319,000"],
    ["Profit", "Calculated", "—", "Number", "All", "Total Revenue − Total Expenses.", "410,000"],
    ["Average Salary", "Calculated", "—", "Number", "All", "Employee_Salary ÷ TOTAL_EMPLOYEES ÷ 3 (quarterly salary spread over 3 months).", "5,000"],
    ["Profit %", "Calculated", "—", "Percent", "All", "Profit ÷ |Revenue| × 100.", "75.9%"],
    ["Added Value %", "Calculated", "—", "Percent", "All", "Added Value ÷ |Gross Output| × 100.", "62.4%"],
    ["Intermediate Consumption %", "Calculated", "—", "Percent", "All", "Intermediate Consumption ÷ |Gross Output| × 100.", "37.6%"],
    ["Validation status", "Calculated", "—", "Text", "All", "'Positive added value' / 'Negative added value' / 'Zero added value'.", "Positive added value"],
]

# ============================================================ QES COMP ===
qes_comp_rows = [
    sec("How comparison works"),
    ["Frame ID", "Selection", "Required", "Text", "All", "One Frame ID is compared across 2 or more of its own quarters (same row, wide Q#_YYYY_ columns).", "F001"],
    ["Selected periods", "Selection", "Required", "List", "All", "Every quarter present as Q#_YYYY_ prefixed columns for that Frame ID is selectable; pick 2+ to compare.", "Q1_2026, Q2_2026"],

    sec("Table 1 — Financial comparison (one row per selected quarter)"),
    ["Revenue", "Financial", "Calculated", "Number", "All", "= Total_Revenue for that quarter.", "500,000 / 540,000"],
    ["Expenditure", "Financial", "Calculated", "Number", "All", "= Total_Expenses for that quarter.", "120,000 / 130,000"],
    ["Profit", "Financial", "Calculated", "Number", "All", "= Revenue − Expenditure.", "380,000 / 410,000"],
    ["Added value", "Financial", "Calculated", "Number", "All", "= Gross Output − Intermediate Consumption.", "295,000 / 319,000"],
    ["IC", "Financial", "Calculated", "Number", "All", "Intermediate Consumption for that quarter.", "180,000 / 192,000"],
    ["Profit %, AV %, IC %", "Financial", "Calculated", "Percent", "All", "Same ratios as the single-quarter calculation page.", "75.9% / 62.4% / 37.6%"],

    sec("Table 2 — Inventory continuity (one row per selected quarter)"),
    ["Inventory start", "Inventory", "Financial input", "Number", "Trade, Industry", "= Total_Finished_Goods_Begining_Of_period for that quarter.", "20,000"],
    ["Inventory end", "Inventory", "Financial input", "Number", "Trade, Industry", "= Total_Finished_Goods_End_of_period for that quarter.", "25,000"],
    ["Validation", "Inventory", "Calculated", "Text", "Trade, Industry", "Flags whether this quarter's Inventory start equals the PREVIOUS selected quarter's Inventory end.", "Start equals previous end"],

    sec("Table 3 — Employee comparison (one row per selected quarter)"),
    ["Employees", "Employee", "Financial input", "Number", "All", "= TOTAL_EMPLOYEES for that quarter.", "12 / 13"],
    ["Employee compensation", "Employee", "Financial input", "Number", "All", "= Employee_Salary for that quarter.", "180,000 / 195,000"],
    ["Average salary", "Employee", "Calculated", "Number", "All", "= Employee_Salary ÷ Employees ÷ 3.", "5,000"],
    ["Employee change count", "Employee", "Calculated", "Number", "All", "= Employees(this quarter) − Employees(previous quarter). First row shows 'Base period'.", "+1"],
    ["Employee change %", "Employee", "Calculated", "Percent", "All", "= Employee change count ÷ |previous quarter Employees| × 100.", "8.3%"],
]

# ============================================================ FISQ CALC ==
fisq_calc_rows = [
    sec("Identification & assignment (one value per Frame ID, no quarter prefix)"),
    ["FRAME_ID", "Identification", "Required", "Text", "All", "Unique establishment / project identifier. One row per Frame ID.", "F001"],
    ["ESTABLISHMENT_NAME", "Identification", "Optional", "Text", "All", "Business / establishment name.", "Al Noor Trading LLC"],
    ["AUDITOR_NAME_EN / AUDITOR_NAME_AR", "Assignment", "Optional", "Text", "All", "Auditor's display name.", "Fatima Auditor"],
    ["Auditor_User_Name", "Assignment", "Required for Auditor role", "Text", "All", "Login username of the Auditor; drives row-level access, checked per quarter.", "auditor1"],
    ["_ENUMERATOR_User_name", "Assignment", "Required for Enumerator role", "Text", "All", "Login username of the Enumerator / Key Account Manager.", "enum1"],
    ["Supervisor_USER_Name", "Assignment", "Required for Supervisor role", "Text", "All", "Login username of the Supervisor.", "sup1"],

    sec("Per-quarter asset equation inputs — prefix each with the quarter, e.g. Q1_2026_Current_Assets"),
    ["Current_Assets", "Financial input", "Required", "Number", "All", "Current assets value for the quarter.", "300,000"],
    ["Non_Current_Assets", "Financial input", "Required", "Number", "All", "Non-current assets value for the quarter.", "500,000"],
    ["Total_Assets", "Financial input (cross-check)", "Optional", "Number", "All", "Total assets as reported directly in the file; shown next to the calculated total for comparison, does not replace it.", "800,000"],
    ["Current_Liabilities", "Financial input", "Required", "Number", "All", "Current liabilities value for the quarter.", "150,000"],
    ["Non_Current_Liabilities", "Financial input", "Required", "Number", "All", "Non-current liabilities value for the quarter.", "200,000"],
    ["Total_Liabilities", "Financial input (cross-check)", "Optional", "Number", "All", "Total liabilities as reported directly in the file; shown next to the calculated total for comparison.", "350,000"],
    ["Equity_In_Form", "Financial input", "Required", "Number", "All", "Equity value as reported on the survey form; the anchor value everything else is checked against.", "450,000"],
    ["Equity_Calculated", "Financial input (cross-check)", "Optional", "Number", "All", "The file's own calculated equity value, if supplied, shown next to the tool's calculated equity.", "450,000"],
    ["Difference_In_Equation", "Financial input (cross-check)", "Optional", "Number", "All", "The file's own equation difference value, if supplied.", "0"],
    ["Equity_7_1", "Financial input", "Required", "Number", "All", "Equity detail line 7.1 (e.g. paid-up capital).", "100,000"],
    ["Equity_7_2", "Financial input", "Required", "Number", "All", "Equity detail line 7.2.", "200,000"],
    ["Equity_7_3", "Financial input", "Required", "Number", "All", "Equity detail line 7.3.", "150,000"],
    ["Equity_7_4", "Financial input", "Required", "Number", "All", "Equity detail line 7.4.", "0"],
    ["Equity_7_5", "Financial input", "Required", "Number", "All", "Equity detail line 7.5.", "0"],
    ["Equity_7_6", "Financial input", "Required", "Number", "All", "Equity detail line 7.6.", "0"],

    sec("Calculated outputs (produced by the tool — do not upload these)"),
    ["Total Assets", "Calculated", "—", "Number", "All", "= Current_Assets + Non_Current_Assets.", "800,000"],
    ["Total Liabilities", "Calculated", "—", "Number", "All", "= Current_Liabilities + Non_Current_Liabilities.", "350,000"],
    ["Calculated Equity", "Calculated", "—", "Number", "All", "= Total Assets − Total Liabilities.", "450,000"],
    ["Equation Difference", "Calculated", "—", "Number", "All", "= Calculated Equity − Equity_In_Form. Should be ~0.", "0"],
    ["Equity Details Total", "Calculated", "—", "Number", "All", "= Equity_7_1 + 7_2 + 7_3 + 7_4 + 7_5 + 7_6.", "450,000"],
    ["Equity Details Difference", "Calculated", "—", "Number", "All", "= Equity Details Total − Equity_In_Form.", "0"],
    ["7.2 + 7.3 Total", "Calculated", "—", "Number", "All", "= Equity_7_2 + Equity_7_3.", "350,000"],
    ["7.2 + 7.3 Difference / Status", "Calculated", "—", "Number/Text", "All", "= (7.2+7.3) − Equity_In_Form; status 'Correct' if within 0.01, else 'Flag / review'.", "Correct"],
    ["Status", "Calculated", "—", "Text", "All", "'Balanced' if both the Equation Difference and the 7.2+7.3 check are within 0.01, else 'Needs review'.", "Balanced"],
]

# ============================================================ FISQ COMP ==
fisq_comp_rows = [
    sec("How comparison works"),
    ["Frame ID", "Selection", "Required", "Text", "All", "One Frame ID is compared across 2 or more of its own quarters.", "F001"],
    ["Selected periods", "Selection", "Required", "List", "All", "Every quarter present as Q#_YYYY_ prefixed columns for that Frame ID is selectable; pick 2+ to compare.", "Q1_2026, Q2_2026"],

    sec("Table 1 — Asset equation comparison (one row per selected quarter)"),
    ["Current assets / Non-current assets / Total assets", "Asset equation", "Calculated", "Number", "All", "Same fields as the single-quarter calculation.", "300,000 / 500,000 / 800,000"],
    ["Current liabilities / Non-current liabilities / Total liabilities", "Asset equation", "Calculated", "Number", "All", "Same fields as the single-quarter calculation.", "150,000 / 200,000 / 350,000"],
    ["Calculated equity / Equity in form / Difference", "Asset equation", "Calculated", "Number", "All", "Same fields as the single-quarter calculation.", "450,000 / 450,000 / 0"],

    sec("Table 2 — Equity details comparison (one row per selected quarter)"),
    ["7.1 .. 7.6", "Equity details", "Financial input", "Number", "All", "The six equity detail lines for that quarter.", "100,000 / 200,000 / ..."],
    ["Total / 7.2+7.3 / Equity in form / Diff / Status", "Equity details", "Calculated", "Number/Text", "All", "Same checks as the single-quarter calculation.", "450,000 / 350,000 / 450,000 / 0 / Correct"],

    sec("Table 3 — Validation (one row per selected quarter)"),
    ["Assets: file vs calculated", "Validation", "Calculated", "Text", "All", "Shows Total_Assets (file) / Total Assets (calculated) side by side when a file value was supplied.", "800,000 / 800,000"],
    ["Liabilities: file vs calculated", "Validation", "Calculated", "Text", "All", "Shows Total_Liabilities (file) / Total Liabilities (calculated) side by side when a file value was supplied.", "350,000 / 350,000"],
    ["Equation difference / Status", "Validation", "Calculated", "Number/Text", "All", "Repeats the equation difference and overall Balanced / Needs review status.", "0 / Balanced"],
]

wb = Workbook()
write_sheet(wb.active, "QES Calculations", qes_calc_rows)
write_sheet(wb.create_sheet(), "QES Comparisons", qes_comp_rows)
write_sheet(wb.create_sheet(), "FISQ Calculations", fisq_calc_rows)
write_sheet(wb.create_sheet(), "FISQ Comparisons", fisq_comp_rows)

OUT_PATH = "Proactive_Reporting_QES_FISQ_Variables.xlsx"
wb.save(OUT_PATH)
print("Saved", OUT_PATH)
