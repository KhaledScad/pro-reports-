"""
SQLite storage layer.

Each survey type (QES, FISQ) gets its own physical .db file so the
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
from datetime import datetime

APP_DIR = Path(__file__).resolve().parent.parent
DURABLE_DATA_DIR = APP_DIR / "data"
DURABLE_DATA_DIR.mkdir(exist_ok=True)

LOCAL_DATA_DIR = Path(os.environ.get("LOCALAPPDATA") or tempfile.gettempdir()) / "ProactiveReporting" / "data"
LOCAL_DATA_DIR.mkdir(parents=True, exist_ok=True)

SURVEY_TYPES = ["QES", "FISQ"]

# Minimal annual reference data (Frame ID + {YYYY}_ISIC4 columns only) used
# purely to cross-check ISIC classification stability on the QES/FISQ
# comparison pages against the establishment's AES / FIS annual filings.
# These are deliberately NOT full survey types: no calculation engine, no
# /survey or /compare pages of their own, no dashboard "access" cards.
ANNUAL_ISIC_SOURCES = ["AES", "FIS"]

ALL_DB_NAMES = ["main"] + SURVEY_TYPES + ANNUAL_ISIC_SOURCES

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
    with get_conn("main") as conn:
        conn.execute("DELETE FROM users")
        conn.executemany(
            "INSERT INTO users (username, password, authorization, name, job, location) "
            "VALUES (:username, :password, :authorization, :name, :job, :location)",
            users,
        )
    sync_now()


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
        if current and str(current["authorization"]).upper() == "ADMIN" and user["authorization"] != "ADMIN":
            others = conn.execute(
                "SELECT COUNT(*) AS c FROM users WHERE authorization = 'ADMIN' "
                "AND username != ? COLLATE NOCASE",
                (current["username"],),
            ).fetchone()
            if others["c"] == 0:
                raise ValueError("Keep at least one ADMIN account.")
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


# ----------------------------------------------------------------- logs ---

def add_log(username: str, role: str, action: str, details: str = "", ip_address: str = "") -> None:
    with get_conn("main") as conn:
        conn.execute(
            "INSERT INTO logs (time, username, role, action, details, ip_address) VALUES (?, ?, ?, ?, ?, ?)",
            (datetime.utcnow().isoformat(timespec="seconds") + "Z", username, role, action, details, ip_address),
        )


def list_logs(limit: int = 500) -> list[dict]:
    with get_conn("main") as conn:
        rows = conn.execute(
            "SELECT * FROM logs ORDER BY id DESC LIMIT ?", (limit,)
        ).fetchall()
    return [dict(r) for r in rows]


def all_logs() -> list[dict]:
    """Every log entry ever recorded, oldest first — for the admin export."""
    with get_conn("main") as conn:
        rows = conn.execute("SELECT * FROM logs ORDER BY id ASC").fetchall()
        return [dict(r) for r in rows]


# -------------------------------------------------------------- records ---

def replace_records(survey_type: str, rows: list[dict], source_file: str) -> int:
    """Each upload fully replaces the previous dataset for that survey type,
    mirroring the original tool's `surveyRecords[type] = rowsOut` behavior."""
    uploaded_at = datetime.utcnow().isoformat(timespec="seconds") + "Z"
    with get_conn(survey_type) as conn:
        conn.execute("DELETE FROM records")
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
    requests: callers must not modify them."""
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
    if survey_type in SURVEY_TYPES:
        _normalize_stored_columns(survey_type, out)
    _records_cache[survey_type] = (stamp, out)
    return out


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


def record_count(survey_type: str) -> int:
    with get_conn(survey_type) as conn:
        return conn.execute("SELECT COUNT(*) AS c FROM records").fetchone()["c"]


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
