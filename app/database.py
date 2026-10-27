"""
SQLite storage layer.

Each survey type (QES, FDI) gets its own physical .db file so the
data for each project is completely separate, as requested. A shared
main.db holds user accounts and the action log.

The app itself lives on a network share so its code follows a user across
VDI sessions / workstations, but SQLite does not reliably support writes
over SMB (file locking over a network share can intermittently fail —
this surfaced in practice as "attempt to write a readonly database").
To stay both reliable and durable across VDI resets, every database has
two copies:

  - a LOCAL working copy, on this machine's own disk, that every read and
    write actually goes through — fast, and free of network-share locking
    issues.
  - a DURABLE copy next to the app itself (wherever it's installed — the
    network share, in the VDI case) that the local copy is hydrated from
    on startup and synced back to periodically and on shutdown.

A non-persistent VDI session wipes the local copy on reset, but the next
startup re-hydrates it from the durable copy, so at most the last sync
interval's worth of activity is ever at risk.
"""
from __future__ import annotations

import json
import os
import secrets
import sqlite3
import tempfile
import threading
from contextlib import contextmanager
from pathlib import Path
from datetime import datetime, timedelta, timezone

APP_DIR = Path(__file__).resolve().parent.parent
DURABLE_DATA_DIR = APP_DIR / "data"
DURABLE_DATA_DIR.mkdir(exist_ok=True)

LOCAL_DATA_DIR = Path(os.environ.get("LOCALAPPDATA") or tempfile.gettempdir()) / "ProactiveReporting" / "data"
LOCAL_DATA_DIR.mkdir(parents=True, exist_ok=True)

SURVEY_TYPES = ["QES", "FDI"]

# Minimal annual reference data (Frame ID + {YYYY}_ISIC4 columns only) used
# purely to cross-check ISIC classification stability on the QES/FDI
# comparison pages against the establishment's AES / FIS annual filings.
# These are deliberately NOT full survey types: no calculation engine, no
# /survey or /compare pages of their own, no dashboard "access" cards.
ANNUAL_ISIC_SOURCES = ["AES", "FIS"]

ALL_DB_NAMES = ["main"] + SURVEY_TYPES + ANNUAL_ISIC_SOURCES

# Surveys renamed after their data was first stored: new name -> the old
# name their database files still have. FDI was called FISQ until October 2026.
RENAMED_DBS = {"FDI": "FISQ"}

SYNC_INTERVAL_SECONDS = 60


def _local_path(name: str) -> Path:
    return LOCAL_DATA_DIR / f"{name.lower()}.db"


def _durable_path(name: str) -> Path:
    return DURABLE_DATA_DIR / f"{name.lower()}.db"


def _db_path(name: str) -> Path:
    return _local_path(name)


def _sqlite_copy(src_path: Path, dst_path: Path) -> None:
    """Transaction-safe copy via SQLite's own backup API — safe even while
    the source is open elsewhere. Used for both hydrate and sync."""
    dst_path.parent.mkdir(parents=True, exist_ok=True)
    src = sqlite3.connect(str(src_path))
    dst = sqlite3.connect(str(dst_path))
    try:
        with dst:
            src.backup(dst)
    finally:
        src.close()
        dst.close()


def _file_stamp(path: Path) -> tuple:
    """Changes whenever the database file does: its mtime, size and SQLite's
    own file change counter (header bytes 24-27, bumped on every commit)."""
    st = path.stat()
    with open(path, "rb") as f:
        counter = f.read(28)[24:28]
    return st.st_mtime_ns, st.st_size, counter


def _migrate_renamed_dbs() -> None:
    """On the first startup after a rename, copy the survey's data to its new
    file name. Like hydrating, this machine's working copy wins; otherwise
    the durable copy is used. The old files are left in place, untouched."""
    for new, old in RENAMED_DBS.items():
        local_new = _local_path(new)
        if local_new.exists():
            continue
        if _local_path(old).exists():
            _sqlite_copy(_local_path(old), local_new)
        elif not _durable_path(new).exists() and _durable_path(old).exists():
            _sqlite_copy(_durable_path(old), local_new)


def _hydrate_from_durable() -> None:
    """On startup, pull down any database that exists durably but not yet
    locally — a fresh/reset VDI session, or the first run on this machine."""
    for name in ALL_DB_NAMES:
        local, durable = _local_path(name), _durable_path(name)
        if not local.exists() and durable.exists():
            _sqlite_copy(durable, local)


_synced_stamps: dict[str, tuple] = {}


def sync_now() -> None:
    """Push the local working copies out to the durable location — only the
    ones that changed since the last push: a full copy of qes.db alone takes
    most of a second over the network share. Called periodically, after
    important writes, and on app shutdown."""
    for name in ALL_DB_NAMES:
        local, durable = _local_path(name), _durable_path(name)
        if not local.exists():
            continue
        stamp = _file_stamp(local)
        if _synced_stamps.get(name) == stamp:
            continue
        _sqlite_copy(local, durable)
        _synced_stamps[name] = stamp


def _start_background_sync() -> None:
    def loop():
        while True:
            threading.Event().wait(SYNC_INTERVAL_SECONDS)
            try:
                sync_now()
            except Exception:
                pass  # best-effort; the next tick (or shutdown) will retry

    threading.Thread(target=loop, daemon=True, name="db-durable-sync").start()


def get_or_create_session_secret() -> str:
    """The Starlette session-cookie signing key. Uses SESSION_SECRET_KEY if
    set; otherwise a random key is generated once and persisted to the
    durable data directory (gitignored, never committed) so sessions
    survive app restarts without a hardcoded secret in source."""
    env_secret = os.environ.get("SESSION_SECRET_KEY")
    if env_secret:
        return env_secret
    secret_path = DURABLE_DATA_DIR / "session_secret.txt"
    if secret_path.exists():
        return secret_path.read_text(encoding="utf-8").strip()
    secret = secrets.token_urlsafe(32)
    secret_path.write_text(secret, encoding="utf-8")
    return secret


@contextmanager
def get_conn(name: str):
    conn = sqlite3.connect(_db_path(name))
    conn.row_factory = sqlite3.Row
    try:
        yield conn
        conn.commit()
    finally:
        conn.close()


def init_db() -> None:
    _migrate_renamed_dbs()
    _hydrate_from_durable()
    # main.db: users + logs
    with get_conn("main") as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS users (
                username TEXT PRIMARY KEY,
                password TEXT NOT NULL,
                authorization TEXT NOT NULL,
                name TEXT,
                job TEXT,
                location TEXT
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS logs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                time TEXT NOT NULL,
                username TEXT,
                role TEXT,
                action TEXT,
                details TEXT
            )
        """)
        # Migration: add ip_address to any logs table created before this
        # column existed (SQLite has no "ADD COLUMN IF NOT EXISTS").
        existing_cols = {row["name"] for row in conn.execute("PRAGMA table_info(logs)").fetchall()}
        if "ip_address" not in existing_cols:
            conn.execute("ALTER TABLE logs ADD COLUMN ip_address TEXT")
        # Migration: a suspended account cannot sign in until it is reactivated.
        user_cols = {row["name"] for row in conn.execute("PRAGMA table_info(users)").fetchall()}
        if "suspended" not in user_cols:
            conn.execute("ALTER TABLE users ADD COLUMN suspended INTEGER NOT NULL DEFAULT 0")
        # Admin settings (the database connection and its queries — see db_import.py).
        conn.execute("CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, value TEXT)")
        # The admin's auditor list per survey (Admin -> Auditor assignments):
        # while a survey has one, it decides which Frame IDs a USER opens.
        conn.execute("""
            CREATE TABLE IF NOT EXISTS assignments (
                survey TEXT NOT NULL,
                frame_id TEXT NOT NULL,
                username TEXT NOT NULL,
                name TEXT
            )
        """)
        conn.execute("CREATE INDEX IF NOT EXISTS idx_assignments_survey ON assignments(survey)")
        # Each CONTROLLER's employees (see auth.team_tokens). Kept apart from
        # users so an Account-sheet upload, which replaces users, keeps them.
        conn.execute("""
            CREATE TABLE IF NOT EXISTS controller_members (
                controller TEXT NOT NULL COLLATE NOCASE,
                member TEXT NOT NULL COLLATE NOCASE,
                PRIMARY KEY (controller, member)
            )
        """)
        # Bootstrap admin so the app is reachable before any Account sheet
        # has been uploaded. No credential is hardcoded in source: set
        # BOOTSTRAP_ADMIN_USERNAME / BOOTSTRAP_ADMIN_PASSWORD to choose your
        # own, or leave unset and a random password is generated once and
        # written to data/bootstrap_admin_credentials.txt (gitignored, never
        # committed) plus printed to the console on first startup.
        row = conn.execute("SELECT COUNT(*) AS c FROM users").fetchone()
        if row["c"] == 0:
            username = os.environ.get("BOOTSTRAP_ADMIN_USERNAME", "admin")
            password = os.environ.get("BOOTSTRAP_ADMIN_PASSWORD") or secrets.token_urlsafe(12)
            conn.execute(
                "INSERT INTO users (username, password, authorization, name, job, location) "
                "VALUES (?, ?, ?, ?, ?, ?)",
                (username, password, "ADMIN", "Setup Admin", "Administrator", "Abu Dhabi"),
            )
            if "BOOTSTRAP_ADMIN_PASSWORD" not in os.environ:
                notice = (
                    f"First run: created bootstrap admin account.\n"
                    f"  username: {username}\n  password: {password}\n"
                    f"Log in once and upload a real Account sheet, or set BOOTSTRAP_ADMIN_USERNAME / "
                    f"BOOTSTRAP_ADMIN_PASSWORD to choose your own next time.\n"
                )
                print(notice)
                try:
                    (DURABLE_DATA_DIR / "bootstrap_admin_credentials.txt").write_text(notice, encoding="utf-8")
                except OSError:
                    pass

    # One records table per survey type (and per annual-ISIC source), each
    # in its own database file.
    for survey_type in SURVEY_TYPES + ANNUAL_ISIC_SOURCES:
        with get_conn(survey_type) as conn:
            conn.execute("""
                CREATE TABLE IF NOT EXISTS records (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    frame_id TEXT NOT NULL,
                    source_file TEXT,
                    uploaded_at TEXT,
                    data TEXT NOT NULL
                )
            """)
            conn.execute("CREATE INDEX IF NOT EXISTS idx_frame_id ON records(frame_id)")

    sync_now()  # make sure a freshly created/bootstrapped DB is durable right away
    _start_background_sync()


# ---------------------------------------------------------------- users ---

def list_users() -> list[dict]:
    with get_conn("main") as conn:
        rows = conn.execute("SELECT * FROM users").fetchall()
        return [dict(r) for r in rows]


def get_user(username: str) -> dict | None:
    with get_conn("main") as conn:
        row = conn.execute(
            "SELECT * FROM users WHERE username = ? COLLATE NOCASE", (username,)
        ).fetchone()
        return dict(row) if row else None


def replace_users(users: list[dict]) -> None:
    """Replace every account with `users` (an Account-sheet upload). A user's
    `suspended` is None when the sheet has no Status column; they then keep
    the suspension (or not) their account already had."""
    with get_conn("main") as conn:
        was_suspended = {r["username"].lower() for r in conn.execute("SELECT username FROM users WHERE suspended").fetchall()}
        rows = [{**u, "suspended": int(u["username"].lower() in was_suspended if u.get("suspended") is None else bool(u["suspended"]))}
                for u in users]
        if not any(r["authorization"] == "ADMIN" and not r["suspended"] for r in rows):
            raise ValueError("Include at least one active (not suspended) ADMIN account in the Account sheet.")
        conn.execute("DELETE FROM users")
        conn.executemany(
            "INSERT INTO users (username, password, authorization, name, job, location, suspended) "
            "VALUES (:username, :password, :authorization, :name, :job, :location, :suspended)",
            rows,
        )
    sync_now()


def _check_other_active_admin(conn, account: dict) -> None:
    """Raise ValueError when `account` is the only active ADMIN — suspending,
    removing or demoting it would leave nobody able to open the Admin page."""
    if str(account["authorization"]).upper() != "ADMIN":
        return
    others = conn.execute(
        "SELECT COUNT(*) AS c FROM users WHERE authorization = 'ADMIN' AND NOT suspended "
        "AND username != ? COLLATE NOCASE",
        (account["username"],),
    ).fetchone()
    if others["c"] == 0:
        raise ValueError("Keep at least one active ADMIN account.")


def set_user_suspended(username: str, suspended: bool) -> dict:
    """Suspend (or reactivate) an account; returns it as it was before. A
    suspended user cannot sign in, and is signed out on their next click."""
    with get_conn("main") as conn:
        row = conn.execute("SELECT * FROM users WHERE username = ? COLLATE NOCASE", (username,)).fetchone()
        if not row:
            raise ValueError("User not found.")
        account = dict(row)
        if suspended:
            _check_other_active_admin(conn, account)
        conn.execute("UPDATE users SET suspended = ? WHERE username = ?", (int(suspended), account["username"]))
    sync_now()
    return account


def delete_user(username: str) -> dict:
    """Remove an account, and take it out of every controller team (its own
    team, if a CONTROLLER, and any it was placed in); returns it as it was."""
    with get_conn("main") as conn:
        row = conn.execute("SELECT * FROM users WHERE username = ? COLLATE NOCASE", (username,)).fetchone()
        if not row:
            raise ValueError("User not found.")
        account = dict(row)
        if not account["suspended"]:
            _check_other_active_admin(conn, account)
        conn.execute("DELETE FROM users WHERE username = ?", (account["username"],))
        conn.execute("DELETE FROM controller_members WHERE controller = ? OR member = ?",
                     (account["username"], account["username"]))
    sync_now()
    return account


def save_user(user: dict, original_username: str = "") -> str:
    """Insert one account, or update the account identified by original_username.

    An empty password on update keeps the current password. Returns
    USER_CREATED or USER_UPDATED. Raises ValueError when the write would
    remove the last admin or collide with another username.
    """
    username = user["username"].strip()
    password = (user.get("password") or "").strip()
    original = (original_username or "").strip()
    with get_conn("main") as conn:
        current = None
        if original:
            row = conn.execute(
                "SELECT * FROM users WHERE username = ? COLLATE NOCASE", (original,)
            ).fetchone()
            if not row:
                raise ValueError("User not found.")
            current = dict(row)
        clash = conn.execute(
            "SELECT username FROM users WHERE username = ? COLLATE NOCASE", (username,)
        ).fetchone()
        if clash and (not current or clash["username"].lower() != current["username"].lower()):
            raise ValueError("That username is already in use.")
        if current and not current["suspended"] and user["authorization"] != "ADMIN":
            _check_other_active_admin(conn, current)
        fields = (
            username, user["authorization"], user["name"], user["job"], user["location"],
        )
        if current:
            if password:
                conn.execute(
                    "UPDATE users SET username = ?, password = ?, authorization = ?, name = ?, job = ?, location = ? "
                    "WHERE username = ? COLLATE NOCASE",
                    (username, password, *fields[1:], current["username"]),
                )
            else:
                conn.execute(
                    "UPDATE users SET username = ?, authorization = ?, name = ?, job = ?, location = ? "
                    "WHERE username = ? COLLATE NOCASE",
                    (*fields, current["username"]),
                )
            if username.lower() != current["username"].lower():  # renamed: carry their team memberships over
                conn.execute("UPDATE OR REPLACE controller_members SET controller = ? WHERE controller = ?", (username, current["username"]))
                conn.execute("UPDATE OR REPLACE controller_members SET member = ? WHERE member = ?", (username, current["username"]))
            action = "USER_UPDATED"
        else:
            if not password:
                raise ValueError("Password is required for a new user.")
            conn.execute(
                "INSERT INTO users (username, password, authorization, name, job, location) "
                "VALUES (?, ?, ?, ?, ?, ?)",
                (username, password, *fields[1:]),
            )
            action = "USER_CREATED"
    sync_now()
    return action


def all_teams() -> dict[str, list[str]]:
    """{controller username (lower-case): [employee usernames]} for every CONTROLLER."""
    teams: dict[str, list[str]] = {}
    with get_conn("main") as conn:
        for r in conn.execute("SELECT controller, member FROM controller_members ORDER BY member").fetchall():
            teams.setdefault(r["controller"].lower(), []).append(r["member"])
    return teams


def get_team(controller: str) -> list[str]:
    return all_teams().get(controller.lower(), [])


def set_team(controller: str, members: list[str]) -> list[str]:
    """Replace `controller`'s employees; returns them as saved (deduplicated,
    without the controller itself)."""
    seen: dict[str, str] = {}
    for m in members:
        m = m.strip()
        if m and m.lower() != controller.lower():
            seen.setdefault(m.lower(), m)
    with get_conn("main") as conn:
        conn.execute("DELETE FROM controller_members WHERE controller = ?", (controller,))
        conn.executemany("INSERT INTO controller_members (controller, member) VALUES (?, ?)",
                         [(controller, m) for m in seen.values()])
    sync_now()
    return sorted(seen.values(), key=str.lower)


# ------------------------------------------------------ auditor assignments ---
# The admin's list of which auditor has which Frame IDs in the current
# project, one per survey. While a survey has one, it — not the data's
# assignment columns — decides which Frame IDs a USER opens (auth.visible_records).

_assignments_cache: dict[str, list[dict]] = {}


def assignments(survey_type: str) -> list[dict] | None:
    """`survey_type`'s auditor list as {frame_id, username, name} rows, or
    None when it has none. Cached until the next upload or removal."""
    if survey_type not in _assignments_cache:
        with get_conn("main") as conn:
            rows = conn.execute("SELECT frame_id, username, name FROM assignments WHERE survey = ?", (survey_type,)).fetchall()
        _assignments_cache[survey_type] = [dict(r) for r in rows]
    rows = _assignments_cache[survey_type]
    return rows if rows else None


def replace_assignments(survey_type: str, entries: list[dict], source: dict) -> int:
    """Replace `survey_type`'s auditor list with `entries` ({username, name,
    frame_ids}); `source` (file, sheet, by) is kept for the Admin page.
    Returns the number of Frame ID assignments saved."""
    rows = {(fid, e["username"].strip(), (e.get("name") or "").strip())
            for e in entries for fid in e["frame_ids"] if fid.strip()}
    with get_conn("main") as conn:
        conn.execute("DELETE FROM assignments WHERE survey = ?", (survey_type,))
        conn.executemany("INSERT INTO assignments (survey, frame_id, username, name) VALUES (?, ?, ?, ?)",
                         [(survey_type, *r) for r in sorted(rows)])
        conn.execute("INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)", (
            f"assignments_{survey_type}",
            json.dumps({**source, "time": datetime.utcnow().isoformat(timespec="seconds") + "Z"})))
    _assignments_cache.pop(survey_type, None)
    sync_now()
    return len(rows)


def clear_assignments(survey_type: str) -> None:
    """Remove `survey_type`'s auditor list: the data's assignment columns decide again."""
    with get_conn("main") as conn:
        conn.execute("DELETE FROM assignments WHERE survey = ?", (survey_type,))
        conn.execute("DELETE FROM settings WHERE key = ?", (f"assignments_{survey_type}",))
    _assignments_cache.pop(survey_type, None)
    sync_now()


# ----------------------------------------------------------------- logs ---

# Log times are stored in UTC (the 7-day cutoff below compares them as text)
# and shown in UAE time, which is UTC+4 all year — the UAE has no daylight saving.
UAE_TZ = timezone(timedelta(hours=4))


def uae_time(utc_iso: str) -> str:
    """A stored log time ('2026-10-21T11:31:01Z') in UAE time ('2026-10-21 15:31:01')."""
    try:
        t = datetime.fromisoformat(str(utc_iso).replace("Z", "+00:00"))
    except ValueError:
        return utc_iso or ""
    if t.tzinfo is None:
        t = t.replace(tzinfo=timezone.utc)
    return t.astimezone(UAE_TZ).strftime("%Y-%m-%d %H:%M:%S")


def add_log(username: str, role: str, action: str, details: str = "", ip_address: str = "") -> None:
    with get_conn("main") as conn:
        conn.execute(
            "INSERT INTO logs (time, username, role, action, details, ip_address) VALUES (?, ?, ?, ?, ?, ?)",
            (datetime.utcnow().isoformat(timespec="seconds") + "Z", username, role, action, details, ip_address),
        )


def recent_logs(days: int = 7) -> list[dict]:
    """Every log entry from the last `days` days, newest first — for the
    Admin page. Times are stored as UTC ISO strings, so they compare as text."""
    cutoff = (datetime.utcnow() - timedelta(days=days)).isoformat(timespec="seconds") + "Z"
    with get_conn("main") as conn:
        rows = conn.execute(
            "SELECT * FROM logs WHERE time >= ? ORDER BY id DESC", (cutoff,)
        ).fetchall()
    return [dict(r) for r in rows]


def all_logs() -> list[dict]:
    """Every log entry ever recorded, oldest first — for the admin export."""
    with get_conn("main") as conn:
        rows = conn.execute("SELECT * FROM logs ORDER BY id ASC").fetchall()
        return [dict(r) for r in rows]


# ------------------------------------------------------------- settings ---

def get_settings() -> dict[str, str]:
    with get_conn("main") as conn:
        return {r["key"]: r["value"] for r in conn.execute("SELECT key, value FROM settings").fetchall()}


def save_settings(values: dict[str, str]) -> None:
    with get_conn("main") as conn:
        conn.executemany("INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)", list(values.items()))
    sync_now()


# -------------------------------------------------------------- records ---

# A survey's records can come in two parts: an Excel upload's, and those of
# Admin → Database connection (db_import), whose source_file starts with this.
DATABASE_SOURCE = "Database: "


def _part(database: bool | None) -> tuple[str, tuple]:
    """A WHERE condition for one part of a survey's records: the database
    import's (True), the Excel upload's (False) or both (None)."""
    if database is None:
        return "1", ()
    return (f"{'' if database else 'NOT '}substr(COALESCE(source_file, ''), 1, ?) = ?",
            (len(DATABASE_SOURCE), DATABASE_SOURCE))


def _from_database(record: dict) -> bool:
    return (record.get("source_file") or "").startswith(DATABASE_SOURCE)


def replace_records(survey_type: str, rows: list[dict], source_file: str, keep_database: bool = False) -> int:
    """Each upload fully replaces the previous dataset for that survey type,
    mirroring the original tool's `surveyRecords[type] = rowsOut` behavior —
    except that a database import (source_file starting with DATABASE_SOURCE)
    replaces only the database part, and an Excel upload with `keep_database`
    only the Excel part, so the two can hold different quarters."""
    uploaded_at = datetime.utcnow().isoformat(timespec="seconds") + "Z"
    database = source_file.startswith(DATABASE_SOURCE)
    where, params = _part(database) if database or keep_database else _part(None)
    with get_conn(survey_type) as conn:
        conn.execute(f"DELETE FROM records WHERE {where}", params)
        conn.executemany(
            "INSERT INTO records (frame_id, source_file, uploaded_at, data) VALUES (?, ?, ?, ?)",
            [(r["frame_id"], source_file, uploaded_at, json.dumps(r["data"], ensure_ascii=False)) for r in rows],
        )
    sync_now()
    return len(rows)


_records_cache: dict[str, tuple[tuple, list[dict]]] = {}


def all_records(survey_type: str) -> list[dict]:
    """Every record of `survey_type`, parsed and quarter-placed. Kept in
    memory until its database file changes (an upload), so a page load does
    not re-read and re-parse every record. The records are shared between
    requests: callers must not modify them. The database import's records
    come after the Excel upload's: where both give a Frame ID the same column
    (the quarter the database holds), the database's value counts, as pages
    merge a Frame ID's records in order (app.merged_row). For the quarter(s)
    the database holds, the Excel upload's figures are dropped altogether
    (_database_quarters_only)."""
    stamp = _file_stamp(_db_path(survey_type))
    cached = _records_cache.get(survey_type)
    if cached and cached[0] == stamp:
        return cached[1]
    with get_conn(survey_type) as conn:
        rows = conn.execute("SELECT * FROM records").fetchall()
    out = []
    for r in rows:
        d = dict(r)
        d["data"] = json.loads(d["data"])
        out.append(d)
    out.sort(key=_from_database)
    if survey_type in SURVEY_TYPES:
        _normalize_stored_columns(survey_type, out)
    _align_database_names(out)
    if survey_type in SURVEY_TYPES:
        _database_quarters_only(out)
    _records_cache[survey_type] = (stamp, out)
    return out


def _database_quarters_only(records: list[dict]) -> None:
    """Excel for the earlier quarters, the database for the current one. The
    current quarter is the newest quarter in the data. A database import
    counts only for it: a database row's columns for any earlier quarter are
    dropped (the Excel upload holds those — e.g. a view whose rows say Q1
    2026), and a database record left with no quarter at all is dropped.
    When the database holds the current quarter, it comes from the database
    alone: the Excel upload's figures for it are dropped, so a Frame ID the
    database does not return shows no figures for it instead of the Excel
    file's. Two kinds of the Excel quarter's columns stay, as they are not
    figures: the assignment columns (auditor / enumerator / supervisor
    usernames and names) — the database queries do not return usernames, and
    access to a Frame ID follows them — and the licence number, which groups
    Frame IDs (app.licence_group). `records` are sorted, Excel part first."""
    from .auth import GENERIC_ASSIGN_COLS, ROLE_ASSIGN_COLS, _matching_keys  # deferred: auth imports this module
    from .calc_common import key_period, get_periods
    from .calculations_qes import LICENSE_NUMBER

    if not any(_from_database(r) for r in records):
        return  # no database import: the Excel upload holds every quarter
    _refile_by_own_quarter([r for r in records if _from_database(r)])
    periods = get_periods(dict.fromkeys(k for r in records for k in r["data"]))
    current = periods[-1] if periods else None
    for r in records:
        if _from_database(r):
            r["data"] = {k: v for k, v in r["data"].items() if key_period(k) in (current, None)}
    records[:] = [r for r in records if not _from_database(r) or any(key_period(k) for k in r["data"])]
    if not any(_from_database(r) for r in records):
        return  # the database holds no current-quarter figures: the Excel upload holds every quarter
    assignment = tuple(ROLE_ASSIGN_COLS + GENERIC_ASSIGN_COLS + LICENSE_NUMBER)
    for r in records:
        if _from_database(r):
            continue
        keep = set(_matching_keys(tuple(r["data"]), assignment))
        r["data"] = {k: v for k, v in r["data"].items() if key_period(k) != current or k in keep}


def _refile_by_own_quarter(records: list[dict]) -> None:
    """Database rows stored under one quarter's prefix whose own YEAR /
    QUARTER cells name another quarter (imported before an import honoured
    them: Q3_2026_YEAR 2026, Q3_2026_QUARTER 'Q1') move to that quarter, as
    excel_utils.label_plain_rows now files them on import."""
    from .calc_common import key_period, normalize_header
    from .excel_utils import _cell_quarter

    for r in records:
        d = r["data"]
        for p in sorted({key_period(k) for k in d} - {None}):
            cells = {normalize_header(k[len(p) + 1:]): k for k in d if key_period(k) == p}
            own = _cell_quarter(d.get(cells["year"]), d.get(cells["quarter"])) if {"year", "quarter"} <= cells.keys() else None
            if own and own != p:
                r["data"] = d = {(own + k[len(p):] if key_period(k) == p else k): v for k, v in d.items()}


def _align_database_names(records: list[dict]) -> None:
    """Give the database part's columns the Excel part's spelling where the
    two name one column differently (Oracle's COMMENT_REVENUE, the file's
    Comment_Revenue; "Total_of _Non-Current_Assets"), so the database's value
    replaces the Excel file's when a Frame ID's records are merged instead of
    standing beside it. `records` are sorted, Excel part first."""
    from .calc_common import normalize_header

    spelling: dict[str, str] = {}
    for keys in {tuple(r["data"]) for r in records if not _from_database(r)}:
        for k in keys:
            spelling.setdefault(normalize_header(k), k)
    renames: dict[tuple, dict[str, str]] = {}
    for r in records:
        if not spelling or not _from_database(r):
            continue
        keys = tuple(r["data"])
        if keys not in renames:
            renames[keys] = {k: s for k in keys if (s := spelling.get(normalize_header(k), k)) != k}
        if renames[keys]:
            r["data"] = {renames[keys].get(k, k): v for k, v in r["data"].items()}


_quarter_layout_cache: dict[tuple, object] = {}


def _normalize_stored_columns(survey_type: str, records: list[dict]) -> None:
    """Data stored by an app version that predates upload-time quarter
    placement (excel_utils.normalize_quarter_columns) still carries the merged
    file's raw names — Right_Q2_2026_Total_Revenue, Right_Licence_Number, and
    a Q2_2026_ block that may really hold Q4 2025. Place them on read the way
    an upload now does; data placed at upload comes through unchanged."""
    from .excel_utils import QuarterLayout  # deferred: keeps openpyxl off the DB import path

    by_sheet: dict[tuple, list[dict]] = {}
    for r in records:
        key = (survey_type, r.get("uploaded_at"), r["data"].get("__sheet"), tuple(r["data"].keys()))
        by_sheet.setdefault(key, []).append(r)
    for key, group in by_sheet.items():
        if key not in _quarter_layout_cache:
            _quarter_layout_cache[key] = QuarterLayout([r["data"] for r in group])
        layout = _quarter_layout_cache[key]
        for r in group:
            r["data"] = layout.apply(r["data"])


_period_cache: dict[str, tuple[list, str | None]] = {}


def current_period(survey_type: str) -> str | None:
    """The newest quarter (Q#_YYYY) in `survey_type`'s data, or None if it
    has no quarter columns. Cached until the data changes."""
    from .calc_common import get_periods

    records = all_records(survey_type)
    cached = _period_cache.get(survey_type)
    if cached and cached[0] is records:
        return cached[1]
    periods = get_periods(dict.fromkeys(k for r in records for k in r["data"]))
    _period_cache[survey_type] = (records, periods[-1] if periods else None)
    return _period_cache[survey_type][1]


def records_by_frame(survey_type: str, frame_id: str) -> list[dict]:
    target = frame_id.strip().lower()
    out = []
    for r in all_records(survey_type):
        if r["frame_id"].strip().lower() == target:
            out.append(r)
    if not out:
        for r in all_records(survey_type):
            if target in r["frame_id"].strip().lower():
                out.append(r)
    return out


def distinct_frame_ids(survey_type: str, limit: int = 200) -> list[str]:
    with get_conn(survey_type) as conn:
        rows = conn.execute(
            "SELECT DISTINCT frame_id FROM records ORDER BY frame_id LIMIT ?", (limit,)
        ).fetchall()
    return [r["frame_id"] for r in rows]


def stored_records(survey_type: str, database: bool | None = None) -> list[dict]:
    """{frame_id, data} of `survey_type`'s rows exactly as stored — before
    quarter placement — of one part with `database` (see _part)."""
    where, params = _part(database)
    with get_conn(survey_type) as conn:
        rows = conn.execute(f"SELECT frame_id, data FROM records WHERE {where}", params).fetchall()
    return [{"frame_id": r["frame_id"], "data": json.loads(r["data"])} for r in rows]


def record_count(survey_type: str, database: bool | None = None) -> int:
    """Rows of `survey_type`; with `database`, of one part (see _part)."""
    where, params = _part(database)
    with get_conn(survey_type) as conn:
        return conn.execute(f"SELECT COUNT(*) AS c FROM records WHERE {where}", params).fetchone()["c"]


def last_load(survey_type: str, database: bool | None = None) -> tuple[str, str]:
    """(source_file, uploaded_at) of the data now loaded — of one part with
    `database` (see _part) — or ("", "") if none."""
    where, params = _part(database)
    with get_conn(survey_type) as conn:
        row = conn.execute(f"SELECT source_file, uploaded_at FROM records WHERE {where} LIMIT 1", params).fetchone()
    return (row["source_file"] or "", row["uploaded_at"] or "") if row else ("", "")


def annual_isic_by_frame(source: str, frame_id: str) -> dict:
    """Merged {YYYY}_ISIC4 wide-column data for one Frame ID from an annual
    ISIC source (AES/FIS) — empty dict if nothing was uploaded for it."""
    target = frame_id.strip().lower()
    matches = [r for r in all_records(source) if r["frame_id"].strip().lower() == target]
    if not matches:
        matches = [r for r in all_records(source) if target in r["frame_id"].strip().lower()]
    merged: dict = {}
    for r in matches:
        merged.update(r["data"])
    return merged
