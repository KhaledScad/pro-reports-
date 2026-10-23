"""
QES (Quarter Economic Survey) calculation engine.

Business type drives which formula applies (per SCAD's QES + FDI validation
rules: industrial establishments, commercial establishments, and "the rest
of the economic surveys" — services / construction / real estate / transport
& storage each have their own Profit/Loss rule):

  Industry:                 GO = Revenue + InvEnd - InvBegin
                             IC = IC (file) or Ops + BankCharges + RawMaterial
                             Profit = GO - Expenses
  Trade:                     GO = Gross margin = Revenue + InvEnd - InvBegin - RawMaterial
                             IC = IC (file) or Ops + Expenses
                             Profit = Gross margin - (Expenses - RawMaterial)
  Construction / Transport /
  Other (services, real      GO = Revenue
  estate, ...):               IC = IC (file) or Ops + Expenses
                             Profit = Revenue - Expenses

Trade computes a gross margin first: inventory movement minus net purchases
for resale (the "RawMaterial" column). Because the reported Total Expenses
figure already includes those purchases, they are backed out again before
being subtracted from the margin, so they are not counted twice. Industry's
raw material is used up in production, so it counts in intermediate
consumption (with operations cost and bank charges) instead of coming off
gross output; its Profit = GO - Expenses is the same figure as before. Every
other activity (no cost-of-goods concept in the survey) uses the simple
Revenue - Expenses rule. A non-zero IC column in the file is used as-is (in
the 20-10 data it always equals Ops + BankCharges + RawMaterial for Industry).

  Added Value      = GO - IC
  Average Salary   = Employee_Salary / Employees / 3   (quarterly salary -> monthly)

GO (key `gross_output`) is shown to users as "Total production". A quarter
whose compensation (Employee_Salary) is more than its Total Expenses is
flagged, since compensation is part of expenses.
"""
from __future__ import annotations

import re
from .calc_common import text_val, num_val, get_periods, period_label, fmt, find_key

LICENSE_NUMBER = ["Licence_Number", "License_Number", "LIC_NO"]

REVENUE = ["Total_Revenue"]
EXPENSES = ["Total_Expenses"]
OPS_COST = ["Operations_Cost"]
EMP_SALARY = ["Employee_Salary"]
RAW_MATERIAL = ["RAW_Material"]
BANK_CHARGES = ["Banking_Charges", "Bank_Charges", "T4_3_5_BANK_CHARGE"]
INV_BEGIN = ["Total_Finished_Goods_Begining_Of_period", "Finished_Goods_Begining_Of_period",
             "Finished_Goods_Beginning_Of_period", "Inventory_Start"]
INV_END = ["Total_Finished_Goods_End_of_period", "Finished_Goods_End_of_period", "Inventory_End"]
IC_FILE = ["IC"]
EMPLOYEES = ["TOTAL_EMPLOYEES"]
COMMENT_REV = ["Comment_Revenue"]
COMMENT_EXP = ["Comment_Expense"]
COMMENT_INV = ["Comment_Inventory"]

# A quarter's reported figures (not its names, codes or comments) — how the
# search tells which Frame ID reported a quarter when one licence number has
# several Frame IDs.
FIGURES = [REVENUE, EXPENSES, OPS_COST, EMP_SALARY, RAW_MATERIAL, BANK_CHARGES,
           INV_BEGIN, INV_END, IC_FILE, EMPLOYEES]

ACTIVITY_EN = ["ACTIVITY_TYPE_EN"]
ACTIVITY_AR = ["ACTIVITY_TYPE_AR"]
ISIC4 = ["ISIC4"]
ISIC4_DESC_EN = ["ISIC4_DESC_EN"]
ISIC4_DESC_AR = ["ISIC4_DESC_AR"]

TRADE_WORDS = re.compile(r"trade|تجار|تجارة|بيع", re.IGNORECASE)
INDUSTRY_WORDS = re.compile(r"industry|manufactur|industrial|صناع|تصنيع", re.IGNORECASE)
CONSTRUCTION_WORDS = re.compile(r"construction|تشييد|انشاء|إنشاء|مقاول", re.IGNORECASE)
TRANSPORT_WORDS = re.compile(r"transport|storage|نقل|تخزين", re.IGNORECASE)


def classify_business(row: dict, period: str) -> str:
    isic = text_val(row, period, ISIC4)
    two_digit = None
    if isic[:2].isdigit():
        two_digit = int(isic[:2])
    blob = " ".join([
        text_val(row, period, ACTIVITY_EN), text_val(row, period, ACTIVITY_AR),
        text_val(row, period, ISIC4_DESC_EN), text_val(row, period, ISIC4_DESC_AR),
    ])
    if TRADE_WORDS.search(blob) or (two_digit is not None and 45 <= two_digit <= 47):
        return "Trade"
    if INDUSTRY_WORDS.search(blob) or (two_digit is not None and 10 <= two_digit <= 33):
        return "Industry"
    if CONSTRUCTION_WORDS.search(blob) or (two_digit is not None and 41 <= two_digit <= 43):
        return "Construction"
    if TRANSPORT_WORDS.search(blob) or (two_digit is not None and 49 <= two_digit <= 53):
        return "Transport"
    return "Other"


def activity_label(row: dict, period: str, lang: str = "en") -> str:
    ar = text_val(row, period, ACTIVITY_AR)
    en = text_val(row, period, ACTIVITY_EN)
    return (ar or en) if lang == "ar" else (en or ar)


def calc_qes_period(row: dict, period: str, lang: str = "en") -> dict:
    rev = num_val(row, period, REVENUE)
    exp = num_val(row, period, EXPENSES)
    ops = num_val(row, period, OPS_COST)
    sal = num_val(row, period, EMP_SALARY)
    raw = num_val(row, period, RAW_MATERIAL)
    bank = num_val(row, period, BANK_CHARGES)
    inv_begin = num_val(row, period, INV_BEGIN)
    inv_end = num_val(row, period, INV_END)
    ic_file = num_val(row, period, IC_FILE)
    employees = num_val(row, period, EMPLOYEES)

    cls = classify_business(row, period)
    trade_margin = 0.0
    has_cost_of_goods = cls in ("Trade", "Industry")
    if cls == "Trade":
        trade_margin = rev + inv_end - inv_begin - raw
        gross_output = trade_margin
        ic = ic_file or (ops + exp)
        profit = trade_margin - (exp - raw)
    elif cls == "Industry":
        gross_output = rev + inv_end - inv_begin
        ic = ic_file or (ops + bank + raw)
        profit = gross_output - exp
    else:  # Construction, Transport, and everything else (services, real estate, ...)
        gross_output = rev
        ic = ic_file or (ops + exp)
        profit = rev - exp

    added_value = gross_output - ic
    average_salary = (sal / employees / 3) if (sal and employees) else 0.0

    return {
        "period": period,
        "business_class": cls,
        "has_cost_of_goods": has_cost_of_goods,
        "business_type": activity_label(row, period, lang) or cls,
        "total_revenue": rev,
        "total_expenses": exp,
        "operations_cost": ops,
        "employee_salary": sal,
        "raw_material": raw,
        "bank_charges": bank,
        "begin_inventory": inv_begin,
        "end_inventory": inv_end,
        "ic_from_file": ic_file,
        "employees": employees,
        "gross_output": gross_output,
        "intermediate_consumption": ic,
        "added_value": added_value,
        "profit": profit,
        "trade_margin": trade_margin,
        "average_salary": average_salary,
        "profit_pct": (profit / abs(rev) * 100) if rev else 0.0,
        "av_pct": (added_value / abs(gross_output) * 100) if gross_output else 0.0,
        "ic_pct": (ic / abs(gross_output) * 100) if gross_output else 0.0,
        "comment_revenue": text_val(row, period, COMMENT_REV),
        "comment_expense": text_val(row, period, COMMENT_EXP),
        "comment_inventory": text_val(row, period, COMMENT_INV),
        "validation_status": (
            "Positive added value" if added_value > 0 else
            "Negative added value" if added_value < 0 else
            "Zero added value"
        ),
        # Compensation is part of total expenses, so it can never be larger.
        "compensation_exceeds_expenditure": sal > exp,
        # Paying more than everything the quarter produced is unusual too
        # (only when there is compensation: a negative total production
        # alone is not this flag).
        "compensation_exceeds_production": sal > 0 and sal > gross_output,
    }


# ---------------------------------------------------------------------
# Detail line items from the real SCAD survey form's DB column codes.
# Display-only: none of this feeds Total_Revenue / Total_Expenses / etc.
# above — those keep being read exactly as before. Each entry is
# (label_en, label_ar, [technical codes to sum]); an item is shown only
# if at least one of its codes exists as a column for that record (any of
# its codes may be summed together, e.g. Water = two separate DB fields).
# ---------------------------------------------------------------------
REVENUE_DETAILS = [
    ("Main Revenues", "الإيرادات الرئيسية", ["T3_1_99_TOT"]),
    ("Government subsidies on products", "الإعانات الحكومية على المنتجات", ["T3_2_1_1_GOV_PRODUCTS_TOT"]),
    ("Rentals of Buildings", "إيجارات المباني", ["T3_2_2_RENT_BUL_TOT"]),
    ("Machineries, Equipments Rentals", "إيجارات الآلات والمعدات", ["T3_2_3_1_ME_RENT_TOT"]),
    ("Transportation Means Rentals", "إيجارات وسائل النقل", ["T3_2_3_2_TM_RENT_TOT"]),
    ("Industrial services provided to others", "خدمات صناعية مقدمة للغير", ["T3_2_6_IND_SERV_TOT"]),
    ("Other miscellaneous income", "إيرادات متنوعة أخرى", ["T3_2_9_OTH_MIS_TOT"]),
    ("Net trading income", "صافي إيرادات المتاجرة", ["T3_2_1_1_NET_PROFT_TOT", "T3_2_1_2_COMMISSION_TOT", "T3_2_1_3_NET_PROFT_TOT"]),
    ("Income property management", "إيرادات إدارة الممتلكات", ["T3_2_2_1_RENT_REC_TOT", "T3_2_2_2_RE_MANG_TOT"]),
]

EXPENSE_DETAILS = [
    ("Wages, salaries and Bonuses in cash", "الأجور والرواتب والمكافآت النقدية", ["T4_1_1_1_WAGES_TOT"]),
    ("Benefits granted to employees/workers", "المزايا الممنوحة للموظفين/العمال", ["T4_1_1_2_BENEFITS_TOT"]),
    ("Raw Materials", "المواد الخام", ["T4_1_2_RAW_TOTAL"]),
    ("Water", "المياه", ["T4_1_3_WATER_TOT", "T4_2_2_WATER_TOT"]),
    ("Electricity", "الكهرباء", ["T4_1_4_ELECT_TOT", "T4_2_3_ELECT_TOT"]),
    ("Fuel, Gas and Oils", "الوقود والغاز والزيوت", ["T4_1_5_FUEL_TOT", "T4_2_4_FUEL_TOT"]),
    ("Non Durable Spare Parts", "قطع الغيار غير المعمرة", ["T4_1_6_NON_DUR_TOT"]),
    ("Buildings Rents", "إيجارات المباني", ["T4_1_7_B_RENT_TOT", "T4_2_5_B_RENT_TOT"]),
    ("Machineries and Transportation Means Rentals", "إيجارات الآلات ووسائل النقل",
     ["T4_1_8_1_ME_RENT_TOT", "T4_1_8_2_TM_RENT_TOT", "T4_2_6_1_ME_RENT_TOT", "T4_2_6_2_TM_RENT_TOT"]),
    ("Current Maintenance of Transport Means", "الصيانة الجارية لوسائل النقل", ["T4_1_10_CMTM_TOT", "T4_2_8_CMTM_TOT"]),
    ("Current Maintenance of Machineries & Equipments", "الصيانة الجارية للآلات والمعدات", ["T4_1_11_CMME_TOT", "T4_2_9_CMME_TOT"]),
    ("Current Maintenance of Buildings & Constructions", "الصيانة الجارية للمباني والإنشاءات", ["T4_1_12_CMBC_TOT", "T4_2_10_CMBC_TOT"]),
    ("Transportation and Storage Services", "خدمات النقل والتخزين",
     ["T4_1_13_1_1_AIR_TOT", "T4_1_13_1_2_LAND_TOT", "T4_1_13_1_3_WATER_TOT", "T4_1_13_2_1_AIR_TOT",
      "T4_1_13_2_2_LAND_TOT", "T4_1_13_2_3_WATER_TOT", "T4_1_13_3_STORING_TOT", "T4_2_11_1_1_AIR_TOT",
      "T4_2_11_1_2_LAND_TOT", "T4_2_11_1_3_WATER_TOT", "T4_2_11_2_STORING_TOT"]),
    ("Computer and Information Systems Services", "خدمات الحاسوب ونظم المعلومات",
     ["T4_1_14_1_COMP_REPAIR_TOT", "T4_2_12_1_COMP_REPAIR_TOT", "T4_1_14_2_OTH_COMP_TOT", "T4_2_12_2_OTH_COMP_TOT"]),
    ("Wrapping & Packing Materials", "مواد التغليف والتعبئة", ["T4_1_15_WRAPPING_TOT"]),
    ("Telecommunication", "الاتصالات", ["T4_1_17_TELECO_TOT"]),
    ("Industrial services provided by others", "خدمات صناعية مقدمة من الغير", ["T4_1_19_INDUS_SER_TOTAL"]),
    ("Others (not included above)", "أخرى (غير مدرجة أعلاه)", ["T4_1_97_TOTAL_OTH", "T4_2_97_TOTAL_OTH"]),
    ("Auditors Fees", "أتعاب مراجعي الحسابات", ["T4_2_15_AUDIT_TOT"]),
    ("Engineering & Consultation Services", "الخدمات الهندسية والاستشارية", ["T4_2_16_ENGIN_TOT"]),
    ("Legal & Courts Expenses", "المصاريف القانونية والقضائية", ["T4_2_17_LEGAL_TOT"]),
    ("Advertisements and Promotions", "الإعلانات والترويج", ["T4_2_18_ADVERT_TOT"]),
    ("Banking Charges", "العمولات البنكية", ["T4_3_5_BANK_CHARGE", "Banking_Charges"]),
]

INVENTORY_PURCHASE_DETAILS = [
    ("Opening inventory for resale goods", "مخزون بداية المدة لبضاعة إعادة البيع", ["T4_1_13_BEG_INVT_TOTAL"]),
    ("Net purchases for resale goods", "صافي مشتريات بضاعة إعادة البيع", ["T4_1_21_NET_PURH_TOTAL"]),
    ("Closing inventory for resale goods", "مخزون نهاية المدة لبضاعة إعادة البيع", ["T4_1_31_END_INVT_TOT"]),
]

FIXED_ASSETS_DETAILS = [
    ("Purchases", "المشتريات", ["T6_PURCH_TOT_FXD_ASTS"]),
    ("Additions & Improvements", "الإضافات والتحسينات", ["T6_ADD_IMPR_TOT_FXD_ASTS"]),
    ("Transfer (produced within facility)", "التحويل (منتج داخل المنشأة)", ["T6_TRANSFER_TOT_FXD_ASTS"]),
    ("Sold", "المباع", ["T6_SOLD_TOT_FXD_ASTS"]),
    ("Damaged & Lost", "التالف والمفقود", ["T6_DMGD_LOST_TOT_FXD_ASTS"]),
    ("Transfer (to other establishments)", "التحويل (إلى منشآت أخرى)", ["T6_TRNSFRS_TOT_FXD_ASTS"]),
    ("Charge for the Year", "استهلاك السنة", ["T6_CHARGE_YEAR_TOT_FXD_ASTS"]),
]

DETAIL_GROUPS = [
    ("Revenue details", "تفاصيل الإيرادات", REVENUE_DETAILS),
    ("Expense details", "تفاصيل المصروفات", EXPENSE_DETAILS),
    ("Inventory & purchases (resale goods)", "المخزون والمشتريات (بضاعة إعادة البيع)", INVENTORY_PURCHASE_DETAILS),
    ("Fixed assets", "الأصول الثابتة", FIXED_ASSETS_DETAILS),
]


def qes_detail_groups(row: dict, period: str, lang: str = "en") -> list[dict]:
    """Granular survey-form line items for display alongside the main
    calculation — read-only, not part of any formula above."""
    groups = []
    for title_en, title_ar, items in DETAIL_GROUPS:
        rows = []
        for label_en, label_ar, codes in items:
            if find_key(row, period, codes) is None:
                continue
            value = sum(num_val(row, period, [c]) for c in codes)
            rows.append({"label": label_ar if lang == "ar" else label_en, "value": value})
        if rows:
            groups.append({"title": title_ar if lang == "ar" else title_en, "rows": rows})
    return groups


def available_periods(row: dict) -> list[str]:
    return get_periods(row)


def calc_qes(row: dict, periods: list[str], mode: str = "single", lang: str = "en") -> dict:
    periods = [p for p in periods if p in get_periods(row)] or get_periods(row)
    calcs = [calc_qes_period(row, p, lang) for p in periods]
    if mode == "compare" and len(calcs) >= 2:
        return {"mode": "compare", "periods": periods, "calcs": calcs, "first": calcs[0], "last": calcs[-1]}
    return {"mode": "single", "periods": periods, "calcs": calcs, "first": calcs[0] if calcs else None}
