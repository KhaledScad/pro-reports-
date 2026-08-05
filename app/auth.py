"""
Login, sessions, and role-based row-level access control.

Ported from the final effective JS logic (the v43 patch): a user's mode is
derived from their Job text (ADMIN sees everything; "supervisor" job sees
rows where a Supervisor_USER_Name column matches; "auditor" job sees rows
where an Auditor_User_Name column matches; "enumerator"/"key account" job
sees rows where an _ENUMERATOR_User_name column matches). Matching is done
per period (for QES/FISQ wide columns) so a user can see a record if they
are assigned to it in ANY quarter.
"""
from __future__ import annotations

import re
from fastapi import Request

from . import database as db
from .calc_common import normalize_header, get_periods, clean, row_keys

AUDITOR_COLS = ["Auditor_User_Name", "Auditor_Username"]
ENUMERATOR_COLS = ["_ENUMERATOR_User_name", "ENUMERATOR_User_name", "Enumerator_User_Name", "Enumerator_Username"]
SUPERVISOR_COLS = ["Supervisor_USER_Name", "Supervisor_Username", "ENU_Supervisor_User_Name",
                    "AUD_Supervisor_USER_Name", "Enumerator_Supervisor_User_Name", "Auditor_Supervisor_User_Name"]
AUDITOR_NAME_COLS = ["AUDITOR_NAME_EN", "AUDITOR_NAME_AR"]
ENUMERATOR_NAME_COLS = ["ENUMERATOR_NAME_EN", "ENUMERATOR_NAME_AR"]
# Generic fallback assignment columns, used for AES/FIS which don't have a
# finalized per-role column scheme yet.
GENERIC_ASSIGN_COLS = ["Assigned_To", "Assigned To", "Username", "Enumerator", "Employee_Username", "Owner"]


def norm_user(v) -> str:
    return re.sub(r"[\s_\-.]+", "", clean(v).lower())


def login(username: str, password: str) -> dict | None:
    user = db.get_user(username)
    if user and user["password"] == password:
        return user
    return None


def set_session_user(request: Request, user: dict) -> None:
    request.session["user"] = {
        "username": user["username"], "authorization": user["authorization"],
        "name": user["name"], "job": user["job"], "location": user["location"],
    }


def current_user(request: Request) -> dict | None:
    return request.session.get("user")


def is_admin(user: dict | None) -> bool:
    return bool(user) and str(user.get("authorization", "")).upper() == "ADMIN"


def role_mode(user: dict | None) -> str:
    if not user:
        return "anonymous"
    if is_admin(user):
        return "admin"
    job = clean(user.get("job", "")).lower()
    if re.search(r"supervisor|مشرف", job):
        return "supervisor"
    if re.search(r"auditor|مدقق", job):
        return "auditor"
    if re.search(r"key account|kay account|enumerator|enumerater|emnunater|باحث|عداد", job):
        return "enumerator"
    return "direct"


def user_tokens(user: dict) -> list[str]:
    return [norm_user(user.get(f)) for f in ("username", "name", "email") if user.get(f)]


def _value_matches(value, tokens: list[str]) -> bool:
    nv = norm_user(value)
    if not nv:
        return False
    for t in tokens:
        if nv == t or (len(t) > 2 and t in nv) or (len(nv) > 2 and nv in t):
            return True
    return False


def _has_match(row: dict, columns: list[str], tokens: list[str]) -> bool:
    wanted = {normalize_header(c) for c in columns}
    for k in row_keys(row):
        nk = normalize_header(k)
        if any(nk == w or nk.endswith(w) for w in wanted):
            if _value_matches(row[k], tokens):
                return True
    return False


def can_access_record(user: dict | None, row: dict) -> bool:
    """Row-level visibility check, mirroring the original tool's suffix
    matching (checks every quarter's assignment columns, not just one)."""
    if is_admin(user):
        return True
    if not user:
        return False
    tokens = user_tokens(user)
    if not tokens:
        return False
    mode = role_mode(user)
    if mode == "auditor":
        return _has_match(row, AUDITOR_COLS, tokens) or _has_match(row, AUDITOR_NAME_COLS, tokens)
    if mode == "enumerator":
        return _has_match(row, ENUMERATOR_COLS, tokens) or _has_match(row, ENUMERATOR_NAME_COLS, tokens)
    if mode == "supervisor":
        return _has_match(row, SUPERVISOR_COLS, tokens)
    # "direct" mode / generic USER accounts: allow if the username shows up
    # in any known assignment-style column at all.
    all_cols = AUDITOR_COLS + ENUMERATOR_COLS + SUPERVISOR_COLS + AUDITOR_NAME_COLS + ENUMERATOR_NAME_COLS + GENERIC_ASSIGN_COLS
    return _has_match(row, all_cols, tokens)


def visible_records(user: dict | None, survey_type: str) -> list[dict]:
    return [r for r in db.all_records(survey_type) if can_access_record(user, r["data"])]
