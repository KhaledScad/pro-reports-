"""
FISQ (Foreign Investment Quarter Survey) calculation engine — asset equation.

Ported from the final effective JS logic (the v45 patch, the last one that
touches FISQ formulas):

  Total Assets       = Current_Assets + Non_Current_Assets
  Total Liabilities  = Current_Liabilities + Non_Current_Liabilities
  Calculated Equity  = Total Assets - Total Liabilities
  Equation Diff      = Calculated Equity - Equity_In_Form
  Equity Details     = Equity_7_1 + ... + Equity_7_6
  7.2+7.3 check      = (Equity_7_2 + Equity_7_3) vs Equity_In_Form
"""
from __future__ import annotations

from .calc_common import text_val, num_val, get_periods

# Beginning-of-quarter / end-of-quarter asset-equation columns (بداية المدة /
# نهاية المدة): the same quarter's submission carries both an opening and a
# closing balance-sheet position, mirroring the "_Begining_Of_period" /
# "_End_of_period" suffix convention already used for QES's inventory fields.
# The main per-quarter calculation below reads the *_End_of_period column as
# its primary source (the closing/"as reported" position for the quarter),
# falling back to the old plain column name for any file that predates the
# begin/end split — so a quarter needs only the begin/end pair, not all three.
CURRENT_ASSETS_BEGIN = ["Current_Assets_Begining_Of_period", "Current_Assets_Beginning_Of_period", "Current_Assets_Start"]
CURRENT_ASSETS_END = ["Current_Assets_End_of_period", "Current_Assets_End"]
NON_CURRENT_ASSETS_BEGIN = ["Non_Current_Assets_Begining_Of_period", "Non_Current_Assets_Beginning_Of_period", "Non_Current_Assets_Start"]
NON_CURRENT_ASSETS_END = ["Non_Current_Assets_End_of_period", "Non_Current_Assets_End"]
TOTAL_ASSETS_FILE_BEGIN = ["Total_Assets_Begining_Of_period", "Total_Assets_Beginning_Of_period"]
TOTAL_ASSETS_FILE_END = ["Total_Assets_End_of_period"]
CURRENT_LIAB_BEGIN = ["Current_Liabilities_Begining_Of_period", "Current_Liabilities_Beginning_Of_period", "Current_Liabilities_Start"]
CURRENT_LIAB_END = ["Current_Liabilities_End_of_period", "Current_Liabilities_End"]
NON_CURRENT_LIAB_BEGIN = ["Non_Current_Liabilities_Begining_Of_period", "Non_Current_Liabilities_Beginning_Of_period",
                          "Total_of_Long_Term_Liabilities_Begining_Of_period", "Long_Term_Liabilities_Begining_Of_period",
                          "Non_Current_Liabilities_Start"]
NON_CURRENT_LIAB_END = ["Non_Current_Liabilities_End_of_period", "Non_Current_Liabilities_End",
                        "Total_of_Long_Term_Liabilities_End_of_period", "Long_Term_Liabilities_End_of_period"]
TOTAL_LIAB_FILE_BEGIN = ["Total_Liabilities_Begining_Of_period", "Total_Liabilities_Beginning_Of_period"]
TOTAL_LIAB_FILE_END = ["Total_Liabilities_End_of_period"]
EQUITY_IN_FORM_BEGIN = ["Equity_In_Form_Begining_Of_period", "Equity_In_Form_Beginning_Of_period",
                        "Equity_Form_Begining_Of_period", "Form_Equity_Begining_Of_period"]
EQUITY_IN_FORM_END = ["Equity_In_Form_End_of_period", "Equity_Form_End_of_period", "Form_Equity_End_of_period"]

# Primary per-quarter lookups (used by the main "Asset equation comparison"
# and "Equity details" tables): prefer the End-of-period column, then fall
# back to the old plain column name for files uploaded before the begin/end
# split existed.
CURRENT_ASSETS = CURRENT_ASSETS_END + ["Current_Assets"]
NON_CURRENT_ASSETS = NON_CURRENT_ASSETS_END + ["Non_Current_Assets"]
TOTAL_ASSETS_FILE = TOTAL_ASSETS_FILE_END + ["Total_Assets"]
CURRENT_LIAB = CURRENT_LIAB_END + ["Current_Liabilities"]
NON_CURRENT_LIAB = NON_CURRENT_LIAB_END + ["Non_Current_Liabilities", "Total_of_Long_Term_Liabilities",
                                            "Long_Term_Liabilities", "NonCurrent_Liabilities"]
TOTAL_LIAB_FILE = TOTAL_LIAB_FILE_END + ["Total_Liabilities"]
EQUITY_IN_FORM = EQUITY_IN_FORM_END + ["Equity_In_Form", "Equity_Form", "Form_Equity"]
EQUITY_CALCULATED_FILE = ["Equity_Calculated", "Calculated_Equity"]
DIFFERENCE_FILE = ["Difference_In_Equation"]
EQUITY_DETAIL_KEYS = [f"Equity_7_{i}" for i in range(1, 7)]


def calc_fisq_period(row: dict, period: str) -> dict:
    current_assets = num_val(row, period, CURRENT_ASSETS)
    non_current_assets = num_val(row, period, NON_CURRENT_ASSETS)
    total_assets_file = num_val(row, period, TOTAL_ASSETS_FILE)
    current_liab = num_val(row, period, CURRENT_LIAB)
    non_current_liab = num_val(row, period, NON_CURRENT_LIAB)
    total_liab_file = num_val(row, period, TOTAL_LIAB_FILE)

    total_assets = current_assets + non_current_assets
    total_liabilities = current_liab + non_current_liab
    calculated_equity = total_assets - total_liabilities

    equity_in_form = num_val(row, period, EQUITY_IN_FORM)
    file_equity = num_val(row, period, EQUITY_CALCULATED_FILE)
    equation_difference = calculated_equity - equity_in_form

    e = [num_val(row, period, [k]) for k in EQUITY_DETAIL_KEYS]
    equity_details_total = sum(e)
    equity_details_difference = equity_details_total - equity_in_form

    equity_723_total = e[1] + e[2]  # Equity_7_2 + Equity_7_3
    equity_723_difference = equity_723_total - equity_in_form

    eq_ok = abs(equation_difference) < 0.01
    e723_ok = abs(equity_723_difference) < 0.01

    return {
        "period": period,
        "current_assets": current_assets,
        "non_current_assets": non_current_assets,
        "total_assets": total_assets,
        "total_assets_file": total_assets_file,
        "current_liabilities": current_liab,
        "non_current_liabilities": non_current_liab,
        "total_liabilities": total_liabilities,
        "total_liabilities_file": total_liab_file,
        "calculated_equity": calculated_equity,
        "file_equity": file_equity,
        "equity_in_form": equity_in_form,
        "equation_difference": equation_difference,
        "equity_7_1": e[0], "equity_7_2": e[1], "equity_7_3": e[2],
        "equity_7_4": e[3], "equity_7_5": e[4], "equity_7_6": e[5],
        "equity_details_total": equity_details_total,
        "equity_details_difference": equity_details_difference,
        "equity_723_total": equity_723_total,
        "equity_723_difference": equity_723_difference,
        "equity_723_status": "Correct" if e723_ok else "Flag / review",
        "status": "Balanced" if (eq_ok and e723_ok) else "Needs review",
        "establishment_name": text_val(row, period, ["ESTABLISHMENT_NAME"]),
    }


def _calc_asset_equation(row: dict, period: str, current_assets_keys, non_current_assets_keys,
                          total_assets_file_keys, current_liab_keys, non_current_liab_keys,
                          total_liab_file_keys, equity_in_form_keys) -> dict:
    current_assets = num_val(row, period, current_assets_keys)
    non_current_assets = num_val(row, period, non_current_assets_keys)
    total_assets_file = num_val(row, period, total_assets_file_keys)
    current_liab = num_val(row, period, current_liab_keys)
    non_current_liab = num_val(row, period, non_current_liab_keys)
    total_liab_file = num_val(row, period, total_liab_file_keys)
    equity_in_form = num_val(row, period, equity_in_form_keys)

    total_assets = current_assets + non_current_assets
    total_liabilities = current_liab + non_current_liab
    calculated_equity = total_assets - total_liabilities
    equation_difference = calculated_equity - equity_in_form

    return {
        "current_assets": current_assets,
        "non_current_assets": non_current_assets,
        "total_assets": total_assets,
        "total_assets_file": total_assets_file,
        "current_liabilities": current_liab,
        "non_current_liabilities": non_current_liab,
        "total_liabilities": total_liabilities,
        "total_liabilities_file": total_liab_file,
        "equity_in_form": equity_in_form,
        "calculated_equity": calculated_equity,
        "equation_difference": equation_difference,
        "status": "balanced" if abs(equation_difference) < 0.01 else "needs_review",
    }


def calc_fisq_period_movement(row: dict, period: str) -> dict:
    """Beginning-of-quarter vs end-of-quarter asset-equation position for the
    same quarter (بداية المدة / نهاية المدة) — the within-quarter movement,
    as distinct from the quarter-over-quarter comparison in calc_fisq()."""
    begin = _calc_asset_equation(row, period, CURRENT_ASSETS_BEGIN, NON_CURRENT_ASSETS_BEGIN,
                                  TOTAL_ASSETS_FILE_BEGIN, CURRENT_LIAB_BEGIN, NON_CURRENT_LIAB_BEGIN,
                                  TOTAL_LIAB_FILE_BEGIN, EQUITY_IN_FORM_BEGIN)
    end = _calc_asset_equation(row, period, CURRENT_ASSETS_END, NON_CURRENT_ASSETS_END,
                                TOTAL_ASSETS_FILE_END, CURRENT_LIAB_END, NON_CURRENT_LIAB_END,
                                TOTAL_LIAB_FILE_END, EQUITY_IN_FORM_END)
    return {
        "period": period,
        "begin": begin,
        "end": end,
        "assets_change": end["total_assets"] - begin["total_assets"],
        "liabilities_change": end["total_liabilities"] - begin["total_liabilities"],
        "equity_change": end["calculated_equity"] - begin["calculated_equity"],
    }


def available_periods(row: dict) -> list[str]:
    return get_periods(row)


def calc_fisq(row: dict, periods: list[str], mode: str = "single") -> dict:
    periods = [p for p in periods if p in get_periods(row)] or get_periods(row)
    calcs = [calc_fisq_period(row, p) for p in periods]
    if mode == "compare" and len(calcs) >= 2:
        first, last = calcs[0], calcs[-1]
        change = last["calculated_equity"] - first["calculated_equity"]
        percent = (change / abs(first["calculated_equity"]) * 100) if first["calculated_equity"] else 0.0
        return {"mode": "compare", "periods": periods, "calcs": calcs, "first": first, "last": last,
                "change": change, "percent": percent}
    return {"mode": "single", "periods": periods, "calcs": calcs, "first": calcs[0] if calcs else None}
