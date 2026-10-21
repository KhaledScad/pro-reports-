"""
Shared helpers used by both the QES and FISQ calculation engines. This is a
direct port of the final (v43/v45/v47) header-matching logic from the
original HTML tool: columns can either be plain (`Total_Revenue`) or
period-prefixed wide columns (`Q1_2026_Total_Revenue`), and lookups fall
back from an exact prefixed match to a plain match.
"""
from __future__ import annotations

import re
from functools import lru_cache
from typing import Iterable

_ARABIC_NORMALIZE = str.maketrans({
    "إ": "ا", "أ": "ا", "آ": "ا",
    "ة": "ه",
    "ى": "ي",
})

_QUARTER_RE = re.compile(r"^(Q[1-4])_(\d{4})_", re.IGNORECASE)
# A quarter label, possibly behind merged-file prefixes (Right_Q2_2026_..., Left_Right_...).
_LABELLED_RE = re.compile(r"^(?:(?:left|right)_)*Q[1-4]_\d{4}_", re.IGNORECASE)
_QUARTER_ONLY_RE = re.compile(r"Q\s*([1-4])|([1-4])", re.IGNORECASE)


def clean(v) -> str:
    if v is None:
        return ""
    return str(v).replace("﻿", "").strip()


def normalize_header(v) -> str:
    return _normalize(clean(v))


@lru_cache(maxsize=8192)
def _normalize(s: str) -> str:
    # Cached: the same few hundred column headers are normalized for every
    # record on every page load.
    s = s.translate(_ARABIC_NORMALIZE).lower()
    return re.sub(r"[\s_\-.\\/()]+", "", s)


def to_num(v) -> float:
    s = clean(v)
    if s == "":
        return 0.0
    s = s.replace(",", "")
    neg = False
    if s.startswith("(") and s.endswith(")"):
        neg = True
        s = s[1:-1]
    s = re.sub(r"[^0-9.\-]", "", s)
    try:
        n = float(s)
    except ValueError:
        return 0.0
    return -n if neg else n


def fmt(n) -> str:
    try:
        n = float(n or 0)
    except (TypeError, ValueError):
        n = 0.0
    return f"{n:,.2f}".rstrip("0").rstrip(".") if n != int(n) else f"{int(n):,}"


def pct(n) -> str:
    return f"{fmt(n)}%"


def row_keys(row: dict) -> list[str]:
    return list(row.keys())


def _key_matches_suffix(key: str, suffix: str) -> bool:
    nk, ns = normalize_header(key), normalize_header(suffix)
    # "Current_Assets" must never pick up "Non-Current_Assets" (or liabilities alike).
    return nk == ns or (nk.endswith(ns) and not nk[:-len(ns)].endswith("non"))


def find_key(row: dict, period: str | None, suffixes: Iterable[str]) -> str | None:
    """Locate the column for `suffix` in `row`, preferring an exact
    `{period}_{suffix}` match, then a period-prefixed fuzzy match, then a
    bare (no-period) match."""
    suffixes = (suffixes,) if isinstance(suffixes, str) else tuple(suffixes)
    # Every record from one upload has the same columns, so the lookup is
    # cached per column set rather than repeated for each record.
    return _find_key(tuple(row_keys(row)), period, suffixes)


@lru_cache(maxsize=8192)
def _find_key(keys: tuple[str, ...], period: str | None, suffixes: tuple[str, ...]) -> str | None:
    if period:
        for suffix in suffixes:
            target = normalize_header(f"{period}_{suffix}")
            for k in keys:
                if normalize_header(k) == target:
                    return k
        prefix = normalize_header(period)
        for suffix in suffixes:
            for k in keys:
                if normalize_header(k).startswith(prefix) and _key_matches_suffix(k, suffix):
                    return k

    # Bare match: a plain (no-quarter) column serves any quarter that lacks its own,
    # but another quarter's column never stands in for the one asked for. An exact
    # name wins over a longer one that merely ends with it.
    plain = [k for k in keys if not (period and _LABELLED_RE.match(clean(k)))]
    for suffix in suffixes:
        hits = [k for k in plain if _key_matches_suffix(k, suffix)]
        if hits:
            ns = normalize_header(suffix)
            return next((k for k in hits if normalize_header(k) == ns), hits[0])
    return None


def text_val(row: dict, period: str | None, suffixes) -> str:
    k = find_key(row, period, suffixes)
    return clean(row[k]) if k else ""


def num_val(row: dict, period: str | None, suffixes) -> float:
    return to_num(text_val(row, period, suffixes))


def key_period(key) -> str | None:
    """The Q#_YYYY period a wide column belongs to ('Q1_2026_Total_Revenue'
    -> 'Q1_2026'), or None for a plain column."""
    m = _QUARTER_RE.match(clean(key))
    return f"{m.group(1).upper()}_{m.group(2)}" if m else None


def get_periods(row: dict) -> list[str]:
    """Extract every Q#_YYYY wide-column period present in this row's
    headers, e.g. {'Q1_2026_Total_Revenue': ...} -> 'Q1_2026'."""
    # Cached per column set, like find_key: called for every record on every page load.
    return list(_periods(tuple(row_keys(row))))


@lru_cache(maxsize=1024)
def _periods(keys: tuple[str, ...]) -> tuple[str, ...]:
    found = {p for k in keys if (p := key_period(k))}

    def sort_key(p: str):
        m = re.match(r"Q([1-4])_(\d{4})", p)
        return (int(m.group(2)), int(m.group(1))) if m else (0, 0)

    return tuple(sorted(found, key=sort_key))


_YEAR_ISIC_RE = re.compile(r"^(\d{4})_ISIC4$", re.IGNORECASE)


def annual_isic_series(row: dict) -> list[tuple[str, str]]:
    """Extract every {YYYY}_ISIC4 wide column (the annual-file counterpart
    of QES/FISQ's {Q#_YYYY}_... columns), sorted oldest-to-newest year."""
    out = []
    for k in row_keys(row):
        m = _YEAR_ISIC_RE.match(clean(k))
        if m:
            out.append((m.group(1), clean(row[k])))
    out.sort(key=lambda pair: pair[0])
    return out


def period_label(p: str, lang: str = "en") -> str:
    m = re.match(r"^(Q[1-4])_(\d{4})$", p or "")
    if not m:
        return clean(p)
    q, y = m.group(1), m.group(2)
    n = q[1]
    if lang == "ar":
        names = {"1": "الربع الأول", "2": "الربع الثاني", "3": "الربع الثالث", "4": "الربع الرابع"}
        return f"{names.get(n, q)} - {y}"
    return f"Quarter {n} - {y}"


def frame_id_of(row: dict) -> str:
    for k in row_keys(row):
        nk = normalize_header(k)
        if nk in ("frameid", "frame"):
            return clean(row[k])
    return ""


def latest_text(row: dict, suffixes) -> str:
    """First non-empty value for `suffixes`, newest quarter first, then the
    plain (no quarter prefix) column."""
    for p in [*reversed(get_periods(row)), None]:
        value = text_val(row, p, suffixes)
        if value:
            return value
    return ""


def record_meta(row: dict, lang: str = "en") -> dict:
    """Common identifying fields shown above any QES/FISQ calculation, each
    from the newest quarter that has it — older quarters' blocks in a merged
    file often leave names and assignments empty."""
    def pick(first: list[str], second: list[str] = ()) -> str:
        return latest_text(row, first) or (latest_text(row, second) if second else "")

    ar = lang == "ar"
    # FISQ carries plain AUDITOR_NAME / ENUMERATOR_NAME (no _EN/_AR pair), tried last.
    return {
        "frame_id": frame_id_of(row),
        "establishment_name": pick(["ESTABLISHMENT_NAME"]),
        "isic4": pick(["ISIC4"]),
        "isic4_desc": pick(["ISIC4_DESC_AR"], ["ISIC4_DESC_EN"]) if ar else pick(["ISIC4_DESC_EN"], ["ISIC4_DESC_AR"]),
        "activity_type": pick(["ACTIVITY_TYPE_AR"], ["ACTIVITY_TYPE_EN"]) if ar else pick(["ACTIVITY_TYPE_EN"], ["ACTIVITY_TYPE_AR"]),
        "auditor_name": pick(["AUDITOR_NAME_AR"], ["AUDITOR_NAME_EN", "AUDITOR_NAME"]) if ar
                        else pick(["AUDITOR_NAME_EN"], ["AUDITOR_NAME_AR", "AUDITOR_NAME"]),
        "enumerator_name": pick(["ENUMERATOR_NAME_AR"], ["ENUMERATOR_NAME_EN", "ENUMERATOR_NAME"]) if ar
                           else pick(["ENUMERATOR_NAME_EN"], ["ENUMERATOR_NAME_AR", "ENUMERATOR_NAME"]),
        "auditor_username": pick(["Auditor_User_Name", "Auditor_Username"]),
        "enumerator_username": pick(["_ENUMERATOR_User_name", "ENUMERATOR_User_name", "Enumerator_User_Name"]),
        "supervisor_username": pick(["AUD_Supervisor_USER_Name", "Supervisor_USER_Name",
                                     "Supervisor_Username", "ENU_Supervisor_User_Name"]),
        "stratum": pick(["STRATUM"]),
        "visit_status": pick(["VISIT_STATUS_AR"], ["VISIT_STATUS_EN"]) if ar else pick(["VISIT_STATUS_EN"], ["VISIT_STATUS_AR"]),
        "year": pick(["YEAR"]),
        "employee_number": pick(["Employee_Number", "TOTAL_EMPLOYEES"]),
    }
