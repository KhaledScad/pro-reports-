"""
SQLite storage layer.

Each survey type (QES, FISQ) gets its own physical .db file so the
data for each project is completely separate, as requested. A shared
main.db holds user accounts and the action log.
"""
from __future__ import annotations

import json
import sqlite3
from contextlib import contextmanager
from pathlib import Path
from datetime import datetime

DATA_DIR = Path(__file__).resolve().parent.parent / "data"
DATA_DIR.mkdir(exist_ok=True)

SURVEY_TYPES = ["QES", "FISQ"]


def _db_path(name: str) -> Path:
    return DATA_DIR / f"{name.lower()}.db"


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
