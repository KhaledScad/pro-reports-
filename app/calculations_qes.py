"""
QES (Quarter Economic Survey) calculation engine.

Ported from the final effective JS logic (the v43 patch, which is the last
one that touches QES formulas). Business type drives which formula applies:

  Industry:               GO = Revenue + InvEnd - InvBegin
                           IC = IC (file) or Ops + Expenses + RawMaterial
  Trade:                   GO = TradeMargin = Revenue + InvEnd - InvBegin - RawMaterial
                           IC = IC (file) or Ops + Expenses
  Construction / Transport: GO = Revenue
                           IC = IC (file) or Ops + Expenses
  Other (fallback):        same as Industry

  Added Value      = GO - IC
  Profit           = Revenue - Expenses
  Average Salary   = Employee_Salary / Employees / 3   (quarterly salary -> monthly)
"""
from __future__ import annotations

import re
from .calc_common import text_val, num_val, get_periods, period_label, fmt

REVENUE = ["Total_Revenue"]
EXPENSES = ["Total_Expenses"]
OPS_COST = ["Operations_Cost"]
EMP_SALARY = ["Employee_Salary"]
RAW_MATERIAL = ["RAW_Material"]
INV_BEGIN = ["Total_Finished_Goods_Begining_Of_period", "Finished_Goods_Begining_Of_period",
             "Finished_Goods_Beginning_Of_period", "Inventory_Start"]
INV_END = ["Total_Finished_Goods_End_of_period", "Finished_Goods_End_of_period", "Inventory_End"]
IC_FILE = ["IC"]
EMPLOYEES = ["TOTAL_EMPLOYEES"]
COMMENT_REV = ["Comment_Revenue"]
COMMENT_EXP = ["Comment_Expense"]
COMMENT_INV = ["Comment_Inventory"]

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
    inv_begin = num_val(row, period, INV_BEGIN)
    inv_end = num_val(row, period, INV_END)
    ic_file = num_val(row, period, IC_FILE)
    employees = num_val(row, period, EMPLOYEES)

    cls = classify_business(row, period)
    trade_margin = 0.0
    if cls == "Trade":
        trade_margin = rev + inv_end - inv_begin - raw
        gross_output = trade_margin
        ic = ic_file or (ops + exp)
    elif cls == "Construction" or cls == "Transport":
        gross_output = rev
        ic = ic_file or (ops + exp)
    else:  # Industry or Other (fallback uses the industry formula)
        gross_output = rev + inv_end - inv_begin
        ic = ic_file or (ops + exp + raw)

    added_value = gross_output - ic
    profit = rev - exp
    average_salary = (sal / employees / 3) if (sal and employees) else 0.0

    return {
        "period": period,
        "business_class": cls,
        "business_type": activity_label(row, period, lang) or cls,
        "total_revenue": rev,
        "total_expenses": exp,
        "operations_cost": ops,
        "employee_salary": sal,
        "raw_material": raw,
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
    }


def available_periods(row: dict) -> list[str]:
    return get_periods(row)


def calc_qes(row: dict, periods: list[str], mode: str = "single", lang: str = "en") -> dict:
    periods = [p for p in periods if p in get_periods(row)] or get_periods(row)
    calcs = [calc_qes_period(row, p, lang) for p in periods]
    if mode == "compare" and len(calcs) >= 2:
        first, last = calcs[0], calcs[-1]
        change = last["added_value"] - first["added_value"]
        percent = (change / abs(first["added_value"]) * 100) if first["added_value"] else 0.0
        return {"mode": "compare", "periods": periods, "calcs": calcs, "first": first, "last": last,
                "change": change, "percent": percent}
    return {"mode": "single", "periods": periods, "calcs": calcs, "first": calcs[0] if calcs else None}
