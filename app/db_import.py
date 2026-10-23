"""Live survey data from a database (Admin → Database connection (JDBC)).

The admin enters the JDBC URL their database team gives out, a username and
password, and one SELECT query per dataset (QES, FDI, AES, FIS). The URL is
read here and opened with Python's own driver for that database, so no Java
runtime or driver .jar is needed on the server:

  jdbc:oracle:thin:@//host:1521/SERVICE      python-oracledb, thin mode (no Oracle client);
  jdbc:oracle:thin:@host:1521:SID              falls back to the installed Oracle Client
  jdbc:oracle:thin:@(DESCRIPTION=...)          for servers thin mode cannot reach
  jdbc:sqlserver://host:1433;databaseName=DB   pyodbc + the installed SQL Server ODBC driver

A query's rows go through the same steps as an Excel sheet
(excel_utils.records_from_rows), so imported data looks the same to every
page, and an import replaces that dataset like an upload does. Queries are
read-only: only SELECT / WITH statements run, nothing is ever committed, and
every connection is rolled back before it closes. With auto-refresh on,
datasets that have a query are re-imported every N minutes; an import whose
data has not changed rewrites nothing.
"""
from __future__ import annotations

import hashlib
import json
import re
import threading
import time
from datetime import date, datetime
from decimal import Decimal

from . import database as db
from . import excel_utils
from .calc_common import get_periods, period_label

DATASETS = db.SURVEY_TYPES + db.ANNUAL_ISIC_SOURCES
TIMEOUT_SECONDS = 15

_SELECT_RE = re.compile(r"^\s*(select|with)\b", re.IGNORECASE)
_import_lock = threading.Lock()


# ----------------------------------------------------------- connection ---

def parse_jdbc_url(url: str) -> dict:
    """The parts of a JDBC URL this module can open; ValueError otherwise."""
    url = (url or "").strip()
    low = url.lower()
    if low.startswith("jdbc:oracle:"):
        m = re.match(r"jdbc:oracle:(?:thin|oci):(.*)$", url, re.IGNORECASE | re.DOTALL)
        if not m:
            raise ValueError("Oracle JDBC URLs look like jdbc:oracle:thin:@//host:1521/SERVICE.")
        rest = m.group(1)
        user = password = ""
        at = rest.find("@")
        if at > 0 and not rest.startswith("("):  # jdbc:oracle:thin:user/password@host...
            user, _, password = rest[:at].partition("/")
        target = rest[at + 1:] if at >= 0 else rest
        if target.startswith("("):
            return {"engine": "oracle", "dsn": target, "label": "Oracle", "user": user, "password": password}
        if target.lower().startswith("ldap"):
            raise ValueError("LDAP Oracle URLs are not supported; use //host:port/SERVICE.")
        target = target.lstrip("/")
        m = re.match(r"^([^:/]+)(?::(\d+))?(?:/(.+)|:(.+))?$", target)
        if not m:
            raise ValueError(f"Could not read the Oracle address '{target}'.")
        host, port, service, sid = m.group(1), m.group(2) or "1521", m.group(3), m.group(4)
        if not (service or sid):
            raise ValueError("Add the service name (…:1521/SERVICE) or SID (…:1521:SID) to the URL.")
        return {"engine": "oracle", "host": host, "port": int(port), "service": service, "sid": sid,
                "label": f"Oracle {host}:{port}/{service or sid}", "user": user, "password": password}
    if low.startswith("jdbc:sqlserver://"):
        address, *props = url[len("jdbc:sqlserver://"):].split(";")
        options = {k.strip().lower(): v.strip() for k, _, v in (p.partition("=") for p in props) if k.strip()}
        host, _, port = address.partition(":")
        if not host:
            raise ValueError("SQL Server JDBC URLs look like jdbc:sqlserver://host:1433;databaseName=DB.")
        database = options.get("databasename") or options.get("database") or ""
        return {"engine": "sqlserver", "host": host, "port": port, "database": database, "options": options,
                "label": f"SQL Server {host}{':' + port if port else ''}/{database}",
                "user": options.get("user", ""), "password": options.get("password", "")}
    raise ValueError("Supported JDBC URLs: jdbc:oracle:thin:@//host:1521/SERVICE and "
                     "jdbc:sqlserver://host:1433;databaseName=DB.")


def connect(url: str, user: str, password: str):
    """An open DB-API connection for a JDBC URL (see parse_jdbc_url)."""
    cfg = parse_jdbc_url(url)
    user, password = user or cfg["user"], password or cfg["password"]
    if cfg["engine"] == "oracle":
        try:
            import oracledb
        except ImportError:
            raise RuntimeError("The Oracle driver is not installed: run  .venv\\Scripts\\pip install oracledb") from None
        oracledb.defaults.fetch_lobs = False  # CLOB / BLOB columns arrive as plain text / bytes
        dsn = cfg.get("dsn") or (oracledb.makedsn(cfg["host"], cfg["port"], sid=cfg["sid"]) if cfg["sid"]
                                 else f"{cfg['host']}:{cfg['port']}/{cfg['service']}")
        try:
            return oracledb.connect(user=user, password=password, dsn=dsn, tcp_connect_timeout=TIMEOUT_SECONDS)
        except oracledb.Error as exc:
            # Thin mode cannot reach some older servers / password types; the
            # installed Oracle Client (thick mode) can.
            if not any(code in str(exc) for code in ("DPY-3010", "DPY-3015", "DPY-4011")) or not oracledb.is_thin_mode():
                raise
            oracledb.init_oracle_client()
            return oracledb.connect(user=user, password=password, dsn=dsn)
    try:
        import pyodbc
    except ImportError:
        raise RuntimeError("The SQL Server driver is not installed: run  .venv\\Scripts\\pip install pyodbc") from None
    installed = pyodbc.drivers()
    driver = next((d for d in ("ODBC Driver 18 for SQL Server", "ODBC Driver 17 for SQL Server", "SQL Server") if d in installed), None)
    if not driver:
        raise RuntimeError("No SQL Server ODBC driver is installed on this machine.")
    options = cfg["options"]
    server = cfg["host"] + (f",{cfg['port']}" if cfg["port"] else "")
    parts = [f"DRIVER={{{driver}}}", f"SERVER={server}"]
    if cfg["database"]:
        parts.append(f"DATABASE={cfg['database']}")
    if options.get("integratedsecurity", "").lower() == "true":
        parts.append("Trusted_Connection=yes")
    else:
        escaped = password.replace("}", "}}")  # ODBC: a braced value doubles its closing braces
        parts += [f"UID={user}", f"PWD={{{escaped}}}"]
    for jdbc_key, odbc_key in (("encrypt", "Encrypt"), ("trustservercertificate", "TrustServerCertificate")):
        if jdbc_key in options:
            parts.append(f"{odbc_key}={'yes' if options[jdbc_key].lower() == 'true' else 'no'}")
    return pyodbc.connect(";".join(parts), timeout=TIMEOUT_SECONDS, autocommit=False)


def describe(url: str) -> str:
    """The database a URL points at, without any credentials."""
    try:
        return parse_jdbc_url(url)["label"]
    except ValueError:
        return "database"


# ---------------------------------------------------------------- query ---

def clean_query(sql: str) -> str:
    """`sql` as one read-only statement; ValueError for anything else."""
    sql = (sql or "").strip().rstrip(";").strip()
    if not sql:
        raise ValueError("Enter a SELECT query first.")
    if not _SELECT_RE.match(sql):
        raise ValueError("Only SELECT queries can run here — the tool never changes your database.")
    if ";" in sql:
        raise ValueError("Enter a single SELECT statement (no ';').")
    return sql


def _cell(value):
    """A database value as the app stores it (what an Excel cell would hold)."""
    if value is None:
        return ""
    if isinstance(value, Decimal):
        return int(value) if value == value.to_integral_value() else float(value)
    if isinstance(value, float) and value.is_integer():
        return int(value)
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    if isinstance(value, (bytes, bytearray)):
        return value.decode("utf-8", errors="replace")
    return value


def fetch_rows(settings: dict, sql: str) -> list[dict]:
    """Run one read-only query; its rows as {column: value} dicts."""
    sql = clean_query(sql)
    if not settings.get("db_url"):
        raise ValueError("Save the database connection first.")
    conn = connect(settings["db_url"], settings.get("db_user", ""), settings.get("db_password", ""))
    try:
        cur = conn.cursor()
        cur.arraysize = 1000
        cur.execute(sql)
        names = excel_utils._unique_headers([d[0] for d in cur.description])
        rows = [{n: _cell(v) for n, v in zip(names, row)} for row in cur.fetchall()]
        cur.close()
        return rows
    finally:
        try:
            conn.rollback()
        finally:
            conn.close()


def test_connection(settings: dict) -> str:
    """Connect and run a trivial query; a short success message, or raises."""
    url = settings.get("db_url", "")
    probe = "SELECT 1 FROM DUAL" if parse_jdbc_url(url)["engine"] == "oracle" else "SELECT 1"
    fetch_rows(settings, probe)
    return f"Connected to {describe(url)}."


def preview(kind: str) -> str:
    """What an import of `kind` would load, without loading it."""
    settings = db.get_settings()
    rows = fetch_rows(settings, settings.get(f"query_{kind}", ""))
    records, notes = excel_utils.records_from_rows(rows, "database")
    if not records:
        return (f"The query returned {len(rows)} rows, but none has a Frame ID — "
                "include a FRAME_ID column (or alias one, e.g. SELECT ESTABLISHMENT_ID AS FRAME_ID, …).")
    periods = sorted({p for r in records for p in get_periods(r["data"])}, key=lambda p: (p[3:], p[:2]))
    return (f"{len(rows)} rows, {len(rows[0])} columns → {len(records)} records for "
            f"{len({r['frame_id'] for r in records})} Frame IDs; quarters: "
            f"{', '.join(period_label(p) for p in periods) or 'none found'}" + "".join(f"; {n}" for n in notes)
            + ". Nothing was imported.")


# --------------------------------------------------------------- import ---

def _save_status(kind: str, ok: bool, message: str) -> None:
    db.save_settings({f"status_{kind}": json.dumps({
        "time": datetime.utcnow().isoformat(timespec="seconds") + "Z", "ok": ok, "message": message})})


def status(kind: str, settings: dict | None = None) -> dict | None:
    """The last import of `kind` from the database: {time, ok, message}."""
    raw = (settings if settings is not None else db.get_settings()).get(f"status_{kind}")
    return json.loads(raw) if raw else None


def import_dataset(kind: str, username: str = "SYSTEM", role: str = "SYSTEM", ip: str = "",
                   automatic: bool = False) -> tuple[bool, str]:
    """Run `kind`'s query and load its rows in place of the current dataset.
    Returns (ok, message). An automatic run that finds the same data as the
    last import rewrites nothing and logs nothing."""
    with _import_lock:
        settings = db.get_settings()
        try:
            rows = fetch_rows(settings, settings.get(f"query_{kind}", ""))
            records, notes = excel_utils.records_from_rows(rows, "database")
            if not records:
                raise ValueError(f"The query returned {len(rows)} rows, but none has a Frame ID.")
            digest = hashlib.sha1(json.dumps(records, sort_keys=True, default=str).encode()).hexdigest()
            if digest == settings.get(f"hash_{kind}") and db.last_load(kind)[1] == settings.get(f"loaded_at_{kind}"):
                message = f"{kind}: no changes since the last import ({len(records)} records)."
                _save_status(kind, True, message)
                if not automatic:
                    db.add_log(username, role, f"{kind}_DB_IMPORTED", message, ip)
                return True, message
            count = db.replace_records(kind, records, f"Database: {describe(settings.get('db_url', ''))}")
            db.save_settings({f"hash_{kind}": digest, f"loaded_at_{kind}": db.last_load(kind)[1]})
            message = f"{kind}: {count} records imported from {describe(settings.get('db_url', ''))}" + "".join(f"; {n}" for n in notes)
            _save_status(kind, True, message)
            db.add_log(username, role, f"{kind}_DB_IMPORTED", message + (" (auto-refresh)" if automatic else ""), ip)
            return True, message
        except Exception as exc:  # noqa: BLE001 — any driver / query error is reported, not raised
            message = f"{kind}: import failed — {exc}"
            previous = status(kind, settings)
            _save_status(kind, False, message)
            if not automatic or not previous or previous.get("message") != message:
                db.add_log(username, role, f"{kind}_DB_IMPORT_FAILED", message, ip)
            return False, message


def start_auto_refresh() -> None:
    """Re-import every dataset that has a query, every `refresh_minutes`
    (Admin setting; 0 = off). Checks the setting every 30 seconds, so a change
    takes effect without a restart."""
    def loop():
        last_run = 0.0
        while True:
            threading.Event().wait(30)
            try:
                settings = db.get_settings()
                minutes = int(settings.get("refresh_minutes") or 0)
                if minutes <= 0 or not settings.get("db_url") or time.monotonic() - last_run < minutes * 60:
                    continue
                last_run = time.monotonic()
                for kind in DATASETS:
                    if settings.get(f"query_{kind}", "").strip():
                        import_dataset(kind, automatic=True)
            except Exception:  # noqa: BLE001 — keep the loop alive; the next tick retries
                pass

    threading.Thread(target=loop, daemon=True, name="db-auto-refresh").start()
