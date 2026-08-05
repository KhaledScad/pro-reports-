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
ALL_DB_NAMES = ["main"] + SURVEY_TYPES

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


def _hydrate_from_durable() -> None:
    """On startup, pull down any database that exists durably but not yet
    locally — a fresh/reset VDI session, or the first run on this machine."""
    for name in ALL_DB_NAMES:
        local, durable = _local_path(name), _durable_path(name)
        if not local.exists() and durable.exists():
            _sqlite_copy(durable, local)


def sync_now() -> None:
    """Push the local working copies out to the durable location. Called
    periodically, after important writes, and on app shutdown."""
    for name in ALL_DB_NAMES:
        local, durable = _local_path(name), _durable_path(name)
        if local.exists():
            _sqlite_copy(local, durable)


def _start_background_sync() -> None:
    def loop():
        while True:
            threading.Event().wait(SYNC_INTERVAL_SECONDS)
            try:
                sync_now()
            except Exception:
                pass  # best-effort; the next tick (or shutdown) will retry

    threading.Thread(target=loop, daemon=True, name="db-durable-sync").start()


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
        # Bootstrap admin so the app is reachable before any Account sheet
        # has been uploaded.
        row = conn.execute("SELECT COUNT(*) AS c FROM users").fetchone()
        if row["c"] == 0:
            conn.execute(
                "INSERT INTO users (username, password, authorization, name, job, location) "
                "VALUES (?, ?, ?, ?, ?, ?)",
                ("admin", "Admin@123", "ADMIN", "Setup Admin", "Administrator", "Abu Dhabi"),
            )

    # One records table per survey type, in its own database file.
    for survey_type in SURVEY_TYPES:
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
            "SELECT * FROM users WHERE username = ?", (username,)
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


# ----------------------------------------------------------------- logs ---

def add_log(username: str, role: str, action: str, details: str = "") -> None:
    with get_conn("main") as conn:
        conn.execute(
            "INSERT INTO logs (time, username, role, action, details) VALUES (?, ?, ?, ?, ?)",
            (datetime.utcnow().isoformat(timespec="seconds") + "Z", username, role, action, details),
        )


def list_logs(limit: int = 500) -> list[dict]:
    with get_conn("main") as conn:
        rows = conn.execute(
            "SELECT * FROM logs ORDER BY id DESC LIMIT ?", (limit,)
        ).fetchall()
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


def all_records(survey_type: str) -> list[dict]:
    with get_conn(survey_type) as conn:
        rows = conn.execute("SELECT * FROM records").fetchall()
    out = []
    for r in rows:
        d = dict(r)
        d["data"] = json.loads(d["data"])
        out.append(d)
    return out


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
