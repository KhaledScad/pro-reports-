"""One-off helper: generates sample Account / QES / FISQ Excel files for
smoke-testing the app. Not part of the running application."""
from openpyxl import Workbook

OUT = "sample_data"
import os
os.makedirs(OUT, exist_ok=True)

# ---- Account.xlsx --------------------------------------------------------
wb = Workbook()
ws = wb.active
ws.title = "Account"
ws.append(["Username", "Password", "Authorization", "Name", "Job", "Location"])
ws.append(["admin", "Admin@123", "ADMIN", "Khaled", "Administrator", "Abu Dhabi"])
ws.append(["auditor1", "Aud@123", "USER", "Fatima Auditor", "Auditor", "Abu Dhabi"])
ws.append(["sup1", "Sup@123", "SUPERVISOR", "Sara Supervisor", "Project Supervisor", "Abu Dhabi"])
wb.save(f"{OUT}/Account.xlsx")

# ---- QES sample: one row per Frame ID, wide quarter columns --------------
wb = Workbook()
ws = wb.active
ws.title = "QES"
headers = ["FRAME_ID", "ESTABLISHMENT_NAME", "ACTIVITY_TYPE_EN", "ISIC4", "ISIC4_DESC_EN",
           "AUDITOR_NAME_EN", "Auditor_User_Name", "Supervisor_USER_Name"]
for q in ["Q1_2026", "Q2_2026"]:
    headers += [f"{q}_Total_Revenue", f"{q}_Total_Expenses", f"{q}_Operations_Cost", f"{q}_RAW_Material",
                f"{q}_Total_Finished_Goods_Begining_Of_period", f"{q}_Total_Finished_Goods_End_of_period",
                f"{q}_TOTAL_EMPLOYEES", f"{q}_Employee_Salary"]
ws.append(headers)
ws.append([
    "F001", "Al Noor Trading LLC", "Trade", "4610", "Wholesale trade",
    "Fatima Auditor", "auditor1", "sup1",
    # Q1_2026
    500000, 120000, 60000, 30000, 20000, 25000, 12, 180000,
    # Q2_2026
    540000, 130000, 62000, 32000, 25000, 28000, 13, 195000,
])
ws.append([
    "F002", "Gulf Manufacturing Co", "Manufacturing", "2599", "Fabricated metal products",
    "Fatima Auditor", "auditor1", "sup1",
    620000, 200000, 90000, 150000, 40000, 45000, 25, 300000,
    650000, 210000, 95000, 155000, 45000, 50000, 26, 312000,
])
wb.save(f"{OUT}/QES_sample.xlsx")

# ---- FISQ sample -----------------------------------------------------------
wb = Workbook()
ws = wb.active
ws.title = "FISQ"
headers = ["FRAME_ID", "ESTABLISHMENT_NAME", "AUDITOR_NAME_EN", "Auditor_User_Name", "Supervisor_USER_Name"]
for q in ["Q1_2026", "Q2_2026"]:
    headers += [f"{q}_Current_Assets", f"{q}_Non_Current_Assets", f"{q}_Current_Liabilities",
                f"{q}_Non_Current_Liabilities", f"{q}_Equity_In_Form",
                f"{q}_Equity_7_1", f"{q}_Equity_7_2", f"{q}_Equity_7_3",
                f"{q}_Equity_7_4", f"{q}_Equity_7_5", f"{q}_Equity_7_6"]
ws.append(headers)
ws.append([
    "F001", "Al Noor Trading LLC", "Fatima Auditor", "auditor1", "sup1",
    300000, 500000, 150000, 200000, 450000, 100000, 200000, 150000, 0, 0, 0,
    320000, 510000, 155000, 205000, 470000, 100000, 220000, 150000, 0, 0, 0,
])
wb.save(f"{OUT}/FISQ_sample.xlsx")

# ===========================================================================
# Dummy: 20 records x 4 quarters (current + previous 3), for QES and FISQ.
# Quarters chosen relative to "today" so they read as current + previous 3.
QUARTERS = ["Q4_2025", "Q1_2026", "Q2_2026", "Q3_2026"]
BUSINESS_TYPES = [
    ("Trade", "تجارة", "4610", "Wholesale trade", "تجارة الجملة"),
    ("Industry", "صناعة", "2599", "Fabricated metal products", "منتجات معدنية مشكلة"),
    ("Construction", "إنشاءات", "4290", "Construction of other civil works", "إنشاءات مدنية أخرى"),
    ("Transport", "نقل", "4923", "Freight transport by road", "نقل البضائع بالطرق البرية"),
]
AUDITORS = [("Fatima Auditor", "فاطمة المدقق", "auditor1"), ("Mohammed Auditor", "محمد المدقق", "auditor1")]
ENUMERATORS = [("Ahmed Enumerator", "أحمد الباحث", "enum1"), ("Layla Enumerator", "ليلى الباحثة", "enum1")]


def grow(base, i, q, jitter=0.0):
    """Deterministic, slightly wavy growth so quarter-over-quarter trend
    coloring (green/red) has both ups and downs to show."""
    wave = [1.00, 1.08, 0.95, 1.12][q]
    return round(base * (1 + i * 0.03) * wave * (1 + jitter))


# ---- QES dummy -------------------------------------------------------------
wb = Workbook()
ws = wb.active
ws.title = "QES"
headers = ["FRAME_ID", "ESTABLISHMENT_NAME", "ACTIVITY_TYPE_EN", "ACTIVITY_TYPE_AR", "ISIC4",
           "ISIC4_DESC_EN", "ISIC4_DESC_AR", "AUDITOR_NAME_EN", "AUDITOR_NAME_AR", "Auditor_User_Name",
           "ENUMERATOR_NAME_EN", "ENUMERATOR_NAME_AR", "_ENUMERATOR_User_name", "Supervisor_USER_Name"]
for q in QUARTERS:
    headers += [f"{q}_Total_Revenue", f"{q}_Total_Expenses", f"{q}_Operations_Cost", f"{q}_RAW_Material",
                f"{q}_Total_Finished_Goods_Begining_Of_period", f"{q}_Total_Finished_Goods_End_of_period",
                f"{q}_TOTAL_EMPLOYEES", f"{q}_Employee_Salary", f"{q}_Comment_Revenue"]
ws.append(headers)

for i in range(20):
    fid = f"F{101 + i}"
    biz_en, biz_ar, isic, isic_en, isic_ar = BUSINESS_TYPES[i % 4]
    aud_en, aud_ar, aud_user = AUDITORS[i % 2]
    enu_en, enu_ar, enu_user = ENUMERATORS[i % 2]
    row = [fid, f"Establishment {i + 1} LLC", biz_en, biz_ar, isic, isic_en, isic_ar,
           aud_en, aud_ar, aud_user, enu_en, enu_ar, enu_user, "sup1"]
    inv_end_prev = None
    for qi, q in enumerate(QUARTERS):
        revenue = grow(400000, i, qi)
        expenses = grow(110000, i, qi)
        ops = grow(55000, i, qi)
        raw = grow(28000, i, qi) if biz_en in ("Trade", "Industry") else 0
        inv_begin = inv_end_prev if inv_end_prev is not None else grow(18000, i, qi)
        # Break inventory continuity intentionally for 2 records to show a red validation row.
        if i in (5, 14) and qi == 2:
            inv_begin = inv_begin + 5000
        inv_end = grow(20000, i, qi) if biz_en in ("Trade", "Industry") else 0
        inv_end_prev = inv_end
        employees = 8 + (i % 10) + qi
        salary = grow(150000, i, qi)
        row += [revenue, expenses, ops, raw, inv_begin, inv_end, employees, salary, ""]
    ws.append(row)
wb.save(f"{OUT}/QES_dummy_20x4Q.xlsx")

# ---- FISQ dummy -------------------------------------------------------------
wb = Workbook()
ws = wb.active
ws.title = "FISQ"
headers = ["FRAME_ID", "ESTABLISHMENT_NAME", "AUDITOR_NAME_EN", "AUDITOR_NAME_AR", "Auditor_User_Name",
           "ENUMERATOR_NAME_EN", "ENUMERATOR_NAME_AR", "_ENUMERATOR_User_name", "Supervisor_USER_Name"]
for q in QUARTERS:
    headers += [f"{q}_Current_Assets", f"{q}_Non_Current_Assets", f"{q}_Current_Liabilities",
                f"{q}_Non_Current_Liabilities", f"{q}_Equity_In_Form",
                f"{q}_Equity_7_1", f"{q}_Equity_7_2", f"{q}_Equity_7_3",
                f"{q}_Equity_7_4", f"{q}_Equity_7_5", f"{q}_Equity_7_6"]
ws.append(headers)

for i in range(20):
    fid = f"F{101 + i}"
    aud_en, aud_ar, aud_user = AUDITORS[i % 2]
    enu_en, enu_ar, enu_user = ENUMERATORS[i % 2]
    row = [fid, f"Establishment {i + 1} LLC", aud_en, aud_ar, aud_user, enu_en, enu_ar, enu_user, "sup1"]
    for qi, q in enumerate(QUARTERS):
        current_assets = grow(280000, i, qi)
        non_current_assets = grow(480000, i, qi)
        current_liab = grow(140000, i, qi)
        non_current_liab = grow(190000, i, qi)
        equity = current_assets + non_current_assets - current_liab - non_current_liab
        e71, e72, e73 = round(equity * 0.3), round(equity * 0.45), round(equity * 0.25)
        # Intentionally unbalance 2 records to show a red "Needs review" status.
        equity_in_form = equity if i not in (3, 17) else equity - 15000
        row += [current_assets, non_current_assets, current_liab, non_current_liab, equity_in_form,
                e71, e72, e73, 0, 0, 0]
    ws.append(row)
wb.save(f"{OUT}/FISQ_dummy_20x4Q.xlsx")

print("Sample files written to", OUT)
print("Dummy files: QES_dummy_20x4Q.xlsx, FISQ_dummy_20x4Q.xlsx (20 Frame IDs, quarters:", ", ".join(QUARTERS), ")")
