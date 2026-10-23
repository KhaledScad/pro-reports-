"""
Login, sessions, and row-level access control.

Only an ADMIN opens every Frame ID. Everyone else opens only the Frame IDs
whose assignment columns (auditor, enumerator or supervisor username / name)
name them in the current quarter, and nothing when none do
(visible_records); a CONTROLLER, the Frame IDs of the employees the admin
placed under them (team_tokens). role_mode (from the Job text, ported from
the original tool's JS) no longer limits which of those columns count.
A quarter that leaves someone's username blank but fills in their full name
still counts: other quarters, in either survey, show which username goes
with that name (learned_names).
"""
from __future__ import annotations

import re
from collections import Counter, defaultdict
from functools import lru_cache
from fastapi import Request

from . import database as db
from .calc_common import normalize_header, get_periods, clean, row_keys, key_period

AUDITOR_COLS = ["Auditor_User_Name", "Auditor_Username"]
ENUMERATOR_COLS = ["_ENUMERATOR_User_name", "ENUMERATOR_User_name", "Enumerator_User_Name", "Enumerator_Username"]
SUPERVISOR_COLS = ["Supervisor_USER_Name", "Supervisor_Username", "ENU_Supervisor_User_Name",
                    "AUD_Supervisor_USER_Name", "Enumerator_Supervisor_User_Name", "Auditor_Supervisor_User_Name"]
AUDITOR_NAME_COLS = ["AUDITOR_NAME_EN", "AUDITOR_NAME_AR", "AUDITOR_NAME"]
ENUMERATOR_NAME_COLS = ["ENUMERATOR_NAME_EN", "ENUMERATOR_NAME_AR", "ENUMERATOR_NAME"]
# Generic fallback assignment columns, used for AES/FIS which don't have a
# finalized per-role column scheme yet.
GENERIC_ASSIGN_COLS = ["Assigned_To", "Assigned To", "Username", "Enumerator", "Employee_Username", "Owner"]
ROLE_ASSIGN_COLS = AUDITOR_COLS + ENUMERATOR_COLS + SUPERVISOR_COLS + AUDITOR_NAME_COLS + ENUMERATOR_NAME_COLS
# A full-name column and the username column for the same person.
NAME_USERNAME_PAIRS = [(AUDITOR_NAME_COLS, AUDITOR_COLS), (ENUMERATOR_NAME_COLS, ENUMERATOR_COLS)]

# False: only an admin opens every Frame ID; everyone else opens only the
# Frame IDs assigned to them in the current quarter (visible_records). True
# would let every signed-in user open every Frame ID, as an admin does
# (a CONTROLLER stays limited to their team either way).
ALL_USERS_SEE_ALL_RECORDS = False


def norm_user(v) -> str:
    return re.sub(r"[\s_\-.]+", "", clean(v).lower())


def account_name(username: str) -> str:
    """Display name of the account with this username (matched loosely, like
    record assignment), or the username itself when no account has it."""
    key = norm_user(username)
    if not key:
        return ""
    for u in db.list_users():
        if norm_user(u.get("username")) == key:
            return clean(u.get("name")) or clean(username)
    return clean(username)


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
    """The signed-in user, re-read from the account list on every request, so
    a role change (e.g. to CONTROLLER) applies on their next click instead of
    their next login, and a removed or suspended account is signed out."""
    user = request.session.get("user")
    if not user:
        return None
    account = db.get_user(user["username"])
    if not account or account.get("suspended"):
        request.session.clear()
        return None
    fresh = {k: account[k] for k in ("username", "authorization", "name", "job", "location")}
    if fresh != user:
        request.session["user"] = fresh
    return fresh


def is_admin(user: dict | None) -> bool:
    return bool(user) and str(user.get("authorization", "")).upper() == "ADMIN"


def is_controller(user: dict | None) -> bool:
    return bool(user) and str(user.get("authorization", "")).upper() == "CONTROLLER"


def team_members(user: dict) -> list[dict]:
    """The employees the admin placed under a CONTROLLER, by name:
    {username, name, job, tokens}; tokens are what their records match on."""
    accounts = {norm_user(u["username"]): u for u in db.list_users()}
    members = []
    for member in db.get_team(user["username"]):
        account = accounts.get(norm_user(member))
        members.append({
            "username": account["username"] if account else member,
            "name": clean(account.get("name")) if account else "",
            "job": clean(account.get("job")) if account else "",
            "tokens": set(user_tokens(account) if account else [norm_user(member)]) - {""},
        })
    return sorted(members, key=lambda m: (m["name"] or m["username"]).lower())


def team_tokens(user: dict) -> set[str]:
    """What a CONTROLLER's records are matched on: the username and account
    name of each employee the admin placed under them, and their own."""
    tokens = set(user_tokens(user))
    for member in team_members(user):
        tokens |= member["tokens"]
    tokens.discard("")
    return tokens


def assigned_values(row: dict, period: str | None = None) -> set[str]:
    """Everyone assigned to a record (as auditor, enumerator, supervisor, …),
    in `period` if given, as matchable tokens."""
    keys = _matching_keys(tuple(row_keys(row)), tuple(ROLE_ASSIGN_COLS + GENERIC_ASSIGN_COLS), period)
    return {v for k in keys if (v := norm_user(row[k]))}


def role_mode(user: dict | None) -> str:
    if not user:
        return "anonymous"
    if is_admin(user):
        return "admin"
    job = clean(user.get("job", "")).lower()
    if str(user.get("authorization", "")).upper() == "SUPERVISOR" or re.search(r"supervisor|مشرف", job):
        return "supervisor"
    if re.search(r"auditor|مدقق", job):
        return "auditor"
    if re.search(r"key account|kay account|enumerator|enumerater|emnunater|باحث|عداد", job):
        return "enumerator"
    return "direct"


def user_tokens(user: dict) -> list[str]:
    return [norm_user(user.get(f)) for f in ("username", "name", "email") if user.get(f)]


def _value_matches(value, tokens: list[str]) -> bool:
    # Exact match only: substring matching leaks records between similar
    # usernames (e.g. "aalshamsi" / "mzaalshamsi").
    nv = norm_user(value)
    return bool(nv) and nv in tokens


def _has_match(row: dict, columns: list[str], tokens: list[str], period: str | None = None) -> bool:
    return any(_value_matches(row[k], tokens) for k in _matching_keys(tuple(row_keys(row)), tuple(columns), period))


@lru_cache(maxsize=1024)
def _matching_keys(keys: tuple[str, ...], columns: tuple[str, ...], period: str | None = None) -> tuple[str, ...]:
    """The keys among `keys` that are one of `columns` (or end with one) —
    cached per column set, since every record of an upload shares one. With
    `period`, only that quarter's keys count, plus plain (no-quarter) ones."""
    wanted = {normalize_header(c) for c in columns}
    out = []
    for k in keys:
        if period and key_period(k) not in (period, None):
            continue
        nk = normalize_header(k)
        if any(nk == w or nk.endswith(w) for w in wanted):
            out.append(k)
    return tuple(out)


def access_tokens(user: dict) -> set[str]:
    """The names a non-admin's Frame IDs are matched on, in any assignment
    column — auditor, enumerator or supervisor, whatever their Job says (the
    data names some "supervisor" jobs as enumerators): their username and
    account name; a CONTROLLER's also include their team's."""
    return team_tokens(user) if is_controller(user) else set(user_tokens(user)) - {""}


_names_cache: dict[str, tuple[list, dict[str, set[str]]]] = {}


def names_by_username(survey_type: str, records: list[dict]) -> dict[str, set[str]]:
    """Matchable username -> the full names the data gives that person (e.g.
    mssalrashdi -> MEERA SALEM SUHAIL TAHI ALRASHDI), learned from records
    where the name and the username sit side by side in the same quarter. A
    name counts only when it was seen with that username at least twice and
    never with another one — in FDI the name and username columns often
    name different people. Cached until the data changes."""
    cached = _names_cache.get(survey_type)
    if cached and cached[0] is records:
        return cached[1]
    seen: dict[str, Counter] = defaultdict(Counter)
    for r in records:
        row = r["data"]
        keys = tuple(row_keys(row))
        for p in get_periods(row):
            for name_cols, user_cols in NAME_USERNAME_PAIRS:
                users = {norm_user(row[k]) for k in _matching_keys(keys, tuple(user_cols), p) if key_period(k) == p} - {""}
                if len(users) != 1:
                    continue
                user = next(iter(users))
                for k in _matching_keys(keys, tuple(name_cols), p):
                    if key_period(k) == p and (name := norm_user(row[k])):
                        seen[name][user] += 1
    learned: dict[str, set[str]] = defaultdict(set)
    for name, users in seen.items():
        if len(users) == 1:
            (user, count), = users.items()
            if count >= 2:
                learned[user].add(name)
    _names_cache[survey_type] = (records, dict(learned))
    return _names_cache[survey_type][1]


def learned_names() -> dict[str, set[str]]:
    """names_by_username across every survey: a person's name is the same in
    QES and FDI, so a name learned in one counts in the other (FDI's Q3 rows
    leave usernames blank that only QES pairs with a name). A name two surveys
    give to different usernames is dropped."""
    owner: dict[str, str] = {}
    conflicts: set[str] = set()
    for st in db.SURVEY_TYPES:
        for user, names in names_by_username(st, db.all_records(st)).items():
            for name in names:
                if owner.setdefault(name, user) != user:
                    conflicts.add(name)
    combined: dict[str, set[str]] = defaultdict(set)
    for name, user in owner.items():
        if name not in conflicts:
            combined[user].add(name)
    return combined


def visible_records(user: dict | None, survey_type: str, period: str | None = None) -> list[dict]:
    """Records `user` can open. An admin opens everything (with `period`: the
    records someone is assigned to in that quarter). Everyone else opens only
    the Frame IDs assigned to them — a CONTROLLER: to their team — in the
    current quarter (or in `period`), so a Frame ID handed to someone else
    since an earlier quarter is no longer theirs, and with no assignment
    they open nothing. Every record of such a Frame ID comes along, so its
    history still shows."""
    if not user:
        return []
    records = db.all_records(survey_type)
    if is_admin(user):
        if not period:
            return list(records)
        return [r for r in records
                if any(clean(r["data"][k]) for k in _matching_keys(tuple(row_keys(r["data"])), tuple(ROLE_ASSIGN_COLS), period))]
    if ALL_USERS_SEE_ALL_RECORDS and not period and not is_controller(user):
        return list(records)
    tokens = access_tokens(user)  # once, not once per record
    if not tokens:
        return []
    # A quarter with only their full name filled in (username left blank) still names them.
    learned = learned_names()
    tokens = tokens | {name for t in tokens for name in learned.get(t, ())}
    when = period or db.current_period(survey_type)
    columns = ROLE_ASSIGN_COLS + GENERIC_ASSIGN_COLS
    frames = {r["frame_id"] for r in records if _has_match(r["data"], columns, tokens, when)}
    return [r for r in records if r["frame_id"] in frames]
