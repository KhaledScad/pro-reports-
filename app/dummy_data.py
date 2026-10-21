"""
Deterministic dummy QES / FISQ records in the exact 20-10 upload layouts
(survey_layouts.py), for the variable reference download and the sample
upload files in sample_data/. The figures follow the tool's own checks, with
a few records deliberately off so the red flags have something to show:

  QES   F106, F115  inventory start differs from the previous quarter's end (Q2 2026)
  FISQ  F104, F118  equity in form is 15,000 off the asset equation
  both  F119, F120  share one licence number (the Record card lists both Frame IDs)
"""
from __future__ import annotations

import re

from .calc_common import normalize_header
from .excel_utils import QuarterLayout, _unique_headers
from .survey_layouts import FISQ_COLUMNS, QES_COLUMNS

# Quarter of each QES block in the dummy records. Blocks 2-4 of the 20-10
# layout are all labelled Q2_2026 (Q2_2026_ / Right_Q2_2026_ /
# Right_Right_Q2_2026_, merge leftovers); each record's own YEAR / QUARTER
# values in them say which quarter they hold.
QES_BLOCKS = {"Q3_2026_": "Q3_2026", "Q2_2026_": "Q2_2026",
              "Right_Q2_2026_": "Q1_2026", "Right_Right_Q2_2026_": "Q4_2025"}
CHRONOLOGICAL = {"QES": ["Q4_2025", "Q1_2026", "Q2_2026", "Q3_2026"],
                 "FISQ": ["Q3_2025", "Q1_2026", "Q2_2026", "Q3_2026"]}

BUSINESS_TYPES = [
    ("Trade", "تجارة", "4610", "Wholesale trade", "تجارة الجملة"),
    ("Industry", "صناعة", "2599", "Fabricated metal products", "منتجات معدنية مشكلة"),
    ("Construction", "إنشاءات", "4290", "Construction of other civil works", "إنشاءات مدنية أخرى"),
    ("Transport", "نقل", "4923", "Freight transport by road", "نقل البضائع بالطرق البرية"),
]
AUDITORS = [("Fatima Auditor", "فاطمة المدقق"), ("Mohammed Auditor", "محمد المدقق")]
ENUMERATORS = [("Ahmed Enumerator", "أحمد الباحث"), ("Layla Enumerator", "ليلى الباحثة")]


def _grow(base: float, i: int, ci: int) -> int:
    """Deterministic, slightly wavy growth so quarter-over-quarter trend
    colouring has both ups and downs to show."""
    return round(base * (1 + i * 0.03) * [1.00, 1.08, 0.95, 1.12][ci])


def _qes_figures(i: int, ci: int) -> dict:
    stock = BUSINESS_TYPES[i % 4][0] in ("Trade", "Industry")

    def inventory_end(c: int) -> int:
        return _grow(20000, i, c) if stock else 0

    inventory_begin = (inventory_end(ci - 1) if ci else _grow(18000, i, ci)) if stock else 0
    if i in (5, 14) and ci == 2:
        inventory_begin += 5000  # breaks inventory continuity on purpose
    return {
        "totalrevenue": _grow(400000, i, ci), "totalexpenses": _grow(110000, i, ci),
        "operationscost": _grow(55000, i, ci), "employeesalary": _grow(150000, i, ci),
        "bankingcharges": _grow(2500, i, ci), "rawmaterial": _grow(28000, i, ci) if stock else 0,
        "totalfinishedgoodsbeginingofperiod": inventory_begin,
        "totalfinishedgoodsendofperiod": inventory_end(ci),
        "totalemployees": 8 + i % 10 + ci, "fundforresearch": 0,
        "ic": "", "commentrevenue": "", "commentexpense": "", "commentinventory": "",
    }


def _fisq_figures(i: int, ci: int) -> dict:
    ca, nca = _grow(280000, i, ci), _grow(480000, i, ci)
    cl, ltl = _grow(140000, i, ci), _grow(190000, i, ci)
    equity = ca + nca - cl - ltl - (15000 if i in (3, 17) else 0)  # 2 records off on purpose
    paid_up = round(equity * 0.6)
    return {
        "totalofcurrentassets": ca, "totalofnoncurrentassets": nca,
        "totalofcurrentliabilities": cl, "totaloflongtermliabilities": ltl,
        "totalequity": equity, "72paidupcapitalstock": paid_up, "73otherequity": equity - paid_up,
        "currasq3d": round(ca * 0.35), "currasq3e": round(ca * 0.5),
        "ncurrasq4d": round(nca * 0.6), "ncurrasq4e": round(nca * 0.3),
        "currliabq5c": round(cl * 0.5), "currliabq5d": round(cl * 0.3),
        "lngliabq6c": round(ltl * 0.55), "lngliabq6d": round(ltl * 0.25),
    }


def _text(key: str, i: int) -> object:
    biz, aud, enu = BUSINESS_TYPES[i % 4], AUDITORS[i % 2], ENUMERATORS[i % 2]
    licence = f"CN-{1000101 + min(i, 18)}"  # F119 and F120 share a licence
    return {
        "frameid": f"F{101 + i}", "establishmentname": f"Establishment {i + 1} LLC",
        "licencenumber": licence, "licensenumber": licence,
        "stratum": 1 + i % 4, "regionid": 1 + i % 3, "address": f"Building {i + 1}, Abu Dhabi",
        "visitstatusen": "Active", "visitstatusar": "نشطة",
        "activitytypeen": biz[0], "activitytypear": biz[1], "isic4": biz[2],
        "isic4descen": biz[3], "isic4descar": biz[4],
        "auditorname": aud[0], "auditorname2": aud[0], "auditornameen": aud[0], "auditornamear": aud[1],
        "enumeratorname": enu[0], "enumeratornameen": enu[0], "enumeratornamear": enu[1],
        "auditorusername": "auditor1", "enumeratorusername": "enum1",
        "audsupervisorusername": "sup1", "enusupervisorusername": "sup1",
    }.get(key, "")


def dummy_rows(survey_type: str, count: int = 20) -> tuple[list[str], list[list]]:
    """The 20-10 header row for QES or FISQ, exactly as in the source file, and
    `count` dummy records under it (Frame IDs F101, F102, ...)."""
    layout = QES_COLUMNS if survey_type == "QES" else FISQ_COLUMNS
    headers = _unique_headers(layout)
    # Where the upload step puts each column. The Q2_2026-labelled blocks are
    # placed through their YEAR / QUARTER values, so the probe carries those.
    probe: dict = {h: None for h in headers}
    for prefix, quarter in QES_BLOCKS.items():
        if f"{prefix}YEAR" in probe:
            probe[f"{prefix}YEAR"], probe[f"{prefix}QUARTER"] = quarter[3:], quarter[:2]
    rename = QuarterLayout([probe]).renames(probe)
    order = CHRONOLOGICAL[survey_type]
    figures = _qes_figures if survey_type == "QES" else _fisq_figures

    rows = []
    for i in range(count):
        row = []
        for h in headers:
            m = re.match(r"^(Q[1-4]_\d{4})_(.+)$", rename.get(h) or h)
            quarter, var = (m.group(1), m.group(2)) if m else (None, h)
            key = normalize_header(re.sub(r"(?i)^((left|right)_)+", "", var))
            value = None
            if quarter in order:
                ci = order.index(quarter)
                numbers = figures(i, ci)
                if key == "year":
                    value = int(quarter[3:])
                elif key == "quarter":
                    value = quarter[:2]
                elif key in numbers:
                    value = numbers[key]
                elif key.endswith("prev") and key[:-4] in numbers:
                    base = key[:-4]
                    value = figures(i, ci - 1)[base] if ci else round(numbers[base] * 0.96)
            row.append(_text(key, i) if value is None else value)
        rows.append(row)
    return list(layout), rows
