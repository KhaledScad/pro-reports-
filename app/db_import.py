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
page. An import replaces only the dataset's database part: the Excel
upload's records stay beside it (database.replace_records), and for the
quarter both hold, the database's values replace the Excel file's — its
columns are also stored under the Excel names for that (to_excel_names). A QES /
FDI query whose columns carry no Q#_YYYY_ label gets the prefix of the
quarter chosen with the connection (the latest ended quarter unless set),
and Test query / Import now name the columns the pages read that it does not return.
Queries are read-only: only SELECT / WITH statements run, nothing is ever
committed, and every connection is rolled back before it closes. With auto-refresh on,
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

from . import calculations_fdi as fdi
from . import calculations_qes as qes
from . import database as db
from . import excel_utils
from .auth import AUDITOR_COLS, AUDITOR_NAME_COLS, ENUMERATOR_COLS, ENUMERATOR_NAME_COLS, SUPERVISOR_COLS
from .calc_common import find_key, get_periods, normalize_header, period_label
from .survey_layouts import FDI_COLUMNS, QES_COLUMNS

DATASETS = db.SURVEY_TYPES + db.ANNUAL_ISIC_SOURCES
TIMEOUT_SECONDS = 15
ROW_QUARTERS = "rows"  # db_quarter setting: each row's own YEAR / QUARTER, not one quarter for all

_SELECT_RE = re.compile(r"^\s*(select|with)\b", re.IGNORECASE)
_QUARTER_SETTING_RE = re.compile(r"^Q[1-4]_\d{4}$")
_import_lock = threading.Lock()

# The columns each survey's pages read: (the Excel file's name, every name it
# is read under). A query that returns none of a group's names leaves that
# field empty on every page, so Test query / Import now name it.
_ASSIGNED = [("Auditor_User_Name", AUDITOR_COLS), ("ENUMERATOR_User_name", ENUMERATOR_COLS),
             ("AUD_Supervisor_USER_Name", SUPERVISOR_COLS)]
READ_COLUMNS = {
    "QES": [(cols[0], cols) for cols in (
        ["ESTABLISHMENT_NAME"], qes.LICENSE_NUMBER, ["STRATUM"], qes.ISIC4, qes.ISIC4_DESC_EN, qes.ISIC4_DESC_AR,
        qes.ACTIVITY_EN, qes.ACTIVITY_AR, *qes.FIGURES, qes.COMMENT_REV, qes.COMMENT_EXP, qes.COMMENT_INV)]
    + [("AUDITOR_NAME_EN", AUDITOR_NAME_COLS), ("ENUMERATOR_NAME_EN", ENUMERATOR_NAME_COLS)] + _ASSIGNED,
    "FDI": [(cols[0], cols) for cols in (
        ["ESTABLISHMENT_NAME"], ["STRATUM"], ["VISIT_STATUS_EN"], ["VISIT_STATUS_AR"], ["ISIC4"],
        ["ISIC4_DESC_EN"], ["ISIC4_DESC_AR"], *fdi.FIGURES)]
    + [("LICENSE_NUMBER", qes.LICENSE_NUMBER), ("AUDITOR_NAME", AUDITOR_NAME_COLS),
       ("ENUMERATOR_NAME", ENUMERATOR_NAME_COLS)] + _ASSIGNED,
}


def _plain_name(column: str) -> str:
    """A column's name without merge prefixes or quarter label."""
    return re.sub(r"^(?:(?:left|right)_)*(?:Q[1-4]_\d{4}_)?", "", column, flags=re.IGNORECASE)


_EXCEL_NAMES = {kind: {normalize_header(_plain_name(c)) for c in columns}
                for kind, columns in (("QES", QES_COLUMNS), ("FDI", FDI_COLUMNS))}
# Each name a column is read under that the Excel file does not use -> the Excel name (LIC_NO -> Licence_Number).
_DATABASE_NAMES = {kind: {normalize_header(n): name for name, names in groups for n in names
                          if normalize_header(n) not in _EXCEL_NAMES[kind]}
                   for kind, groups in READ_COLUMNS.items()}
_LABEL_RE = re.compile(r"^(Q[1-4]_\d{4}_)?(.+)$", re.IGNORECASE)


def to_excel_names(kind: str, records: list[dict]) -> None:
    """Copy, in place, each column the database names differently from the
    Excel file under the Excel name too (Q3_2026_MAIN_REVENUE ->
    Q3_2026_Total_Revenue). For a quarter the Excel upload also holds, the
    database's value then replaces the Excel file's instead of standing beside
    it. The database's own name stays, for the lookups that read it."""
    names = _DATABASE_NAMES.get(kind, {})
    for r in records:
        added = {}
        for k, v in r["data"].items():
            label, plain = _LABEL_RE.match(k).groups()
            target = names.get(normalize_header(plain))
            if target and (new := f"{label or ''}{target}") not in r["data"] and new not in added:
                added[new] = v
        r["data"].update(added)


# -------------------------------------------------------------- quarter ---

def latest_quarter(today: date | None = None) -> str:
    """The newest quarter that has ended: Q3_2026 from October to December 2026."""
    today = today or date.today()
    ended = (today.month - 1) // 3
    return f"Q{ended}_{today.year}" if ended else f"Q4_{today.year - 1}"


def quarter_setting(settings: dict) -> str:
    """The saved db_quarter (Q#_YYYY or ROW_QUARTERS), else the latest quarter."""
    return settings.get("db_quarter") or latest_quarter()


def quarter_choices(saved: str = "", count: int = 8) -> list[str]:
    """The quarters the Admin page offers, newest first: the last `count`
    ended ones, plus `saved` when it is older."""
    q, year = int(latest_quarter()[1]), int(latest_quarter()[3:])
    out = []
    for _ in range(count):
        out.append(f"Q{q}_{year}")
        q, year = (q - 1, year) if q > 1 else (4, year - 1)
    if _QUARTER_SETTING_RE.match(saved or "") and saved not in out:
        out.append(saved)
    return out


def clean_quarter(value: str) -> str:
    """`value` as a db_quarter setting; ValueError for anything else."""
    value = (value or "").strip()
    if value != ROW_QUARTERS and not _QUARTER_SETTING_RE.match(value):
        raise ValueError("Pick the quarter the query's columns belong to.")
    return value


def _query_quarter(kind: str, settings: dict) -> str | None:
    """The quarter `kind`'s unlabelled query columns go under — none for the
    annual AES / FIS ({YYYY}_ISIC4 columns) or with ROW_QUARTERS."""
    quarter = quarter_setting(settings)
    return quarter if kind in db.SURVEY_TYPES and quarter != ROW_QUARTERS else None


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


def column_notes(kind: str, records: list[dict], renamed: bool = True) -> list[str]:
    """Which columns `kind`'s pages read come from a query column the Excel
    file names differently (MAIN_REVENUE as Total_Revenue; with `renamed`),
    and which are not in the query at all — checked in its newest quarter."""
    columns = dict.fromkeys(k for r in records for k in r["data"])
    periods = get_periods(columns)
    if kind not in READ_COLUMNS or not periods:
        return []
    found_as, missing = [], []
    for name, names in READ_COLUMNS[kind]:
        key = find_key(columns, periods[-1], names)
        if key is None:
            missing.append(name)
        elif normalize_header(plain := _plain_name(key)) not in _EXCEL_NAMES[kind]:
            found_as.append(f"{plain} as {name}")
    notes = [f"read from the database's names: {', '.join(found_as)}"] if renamed and found_as else []
    if missing:
        notes.append(f"not in the query, so empty on the pages ({period_label(periods[-1])}): {', '.join(missing)}")
    return notes


def preview(kind: str) -> str:
    """What an import of `kind` would load, without loading it."""
    settings = db.get_settings()
    rows = fetch_rows(settings, settings.get(f"query_{kind}", ""))
    records, notes = excel_utils.records_from_rows(rows, "database", _query_quarter(kind, settings))
    notes += column_notes(kind, records)
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


def record_changes(kind: str, records: list[dict]) -> dict[str, list[str]]:
    """The Frame IDs an import of `records` changes, compared with the
    database part loaded now: "updated" (any value differs), "new" (not loaded
    before) and "removed" (no longer returned by the query)."""
    def by_frame(recs: list[dict]) -> dict[str, list[str]]:
        out: dict[str, list[str]] = {}
        for r in recs:
            out.setdefault(str(r["frame_id"]), []).append(json.dumps(r["data"], sort_keys=True, default=str))
        return {f: sorted(v) for f, v in out.items()}
    old = by_frame(db.stored_records(kind, database=True))
    new = by_frame(records)
    return {"updated": sorted(f for f in new if f in old and new[f] != old[f]),
            "new": sorted(f for f in new if f not in old),
            "removed": sorted(f for f in old if f not in new)}


def _changes_text(changes: dict[str, list[str]]) -> str:
    return ", ".join(f"{len(changes[k])} {k}" for k in ("updated", "new", "removed"))


def _changed_frames_text(changes: dict[str, list[str]], limit: int = 30) -> str:
    """The changed Frame IDs for the activity log, at most `limit` of each kind."""
    parts = []
    for k in ("updated", "new", "removed"):
        if frames := changes[k]:
            more = f" (+{len(frames) - limit} more)" if len(frames) > limit else ""
            parts.append(f"{k} Frame IDs: {', '.join(frames[:limit])}{more}")
    return "".join(f"; {p}" for p in parts)


def import_dataset(kind: str, username: str = "SYSTEM", role: str = "SYSTEM", ip: str = "",
                   automatic: bool = False) -> tuple[bool, str]:
    """Run `kind`'s query and load its rows in place of the current dataset.
    Returns (ok, message). An automatic run that finds the same data as the
    last import rewrites nothing and logs nothing."""
    with _import_lock:
        settings = db.get_settings()
        try:
            rows = fetch_rows(settings, settings.get(f"query_{kind}", ""))
            records, notes = excel_utils.records_from_rows(rows, "database", _query_quarter(kind, settings))
            if not records:
                raise ValueError(f"The query returned {len(rows)} rows, but none has a Frame ID.")
            notes += column_notes(kind, records, renamed=False)
            to_excel_names(kind, records)
            digest = hashlib.sha1(json.dumps(records, sort_keys=True, default=str).encode()).hexdigest()
            if (digest == settings.get(f"hash_{kind}")
                    and db.last_load(kind, database=True)[1] == settings.get(f"loaded_at_{kind}")):
                message = f"{kind}: no changes since the last import ({len(records)} records, 0 updated)."
                _save_status(kind, True, message)
                if not automatic:
                    db.add_log(username, role, f"{kind}_DB_IMPORTED", message, ip)
                return True, message
            changes = record_changes(kind, records)  # before the old database part is replaced
            count = db.replace_records(kind, records, f"{db.DATABASE_SOURCE}{describe(settings.get('db_url', ''))}")
            db.save_settings({f"hash_{kind}": digest, f"loaded_at_{kind}": db.last_load(kind, database=True)[1]})
            if excel := db.record_count(kind, database=False):
                notes.append(f"the Excel upload's {excel} records are kept for the other quarters and the columns the query does not return")
            message = (f"{kind}: {count} records imported ({_changes_text(changes)}) from {describe(settings.get('db_url', ''))}"
                       + "".join(f"; {n}" for n in notes))
            _save_status(kind, True, message)
            db.add_log(username, role, f"{kind}_DB_IMPORTED",
                       message + (" (auto-refresh)" if automatic else "") + _changed_frames_text(changes), ip)
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
