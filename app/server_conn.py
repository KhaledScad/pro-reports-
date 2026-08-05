"""Test connectivity to an external database server the admin has configured.

Drivers are imported lazily per engine so the app still runs with none of
them installed — the admin only needs the driver for the engine they
actually pick, and "Test connection" reports exactly which package to add
if it's missing.
"""
from __future__ import annotations

ENGINE_LABELS = {
    "mssql": "Microsoft SQL Server",
    "postgres": "PostgreSQL",
    "mysql": "MySQL / MariaDB",
}

DEFAULT_PORTS = {"mssql": "1433", "postgres": "5432", "mysql": "3306"}


def test_connection(cfg: dict) -> tuple[bool, str]:
    engine = cfg.get("engine")
    host = cfg.get("host", "")
    port = cfg.get("port") or DEFAULT_PORTS.get(engine, "")
    database = cfg.get("database", "")
    username = cfg.get("username", "")
    password = cfg.get("password", "")

    try:
        if engine == "mssql":
            try:
                import pyodbc
            except ImportError:
                return False, "The 'pyodbc' package is required for SQL Server. Add it to requirements.txt and pip install it."
            conn_str = (
                "DRIVER={ODBC Driver 17 for SQL Server};"
                f"SERVER={host},{port};DATABASE={database};UID={username};PWD={password};"
                "Connection Timeout=5;"
            )
            conn = pyodbc.connect(conn_str, timeout=5)
            conn.close()
        elif engine == "postgres":
            try:
                import psycopg2
            except ImportError:
                return False, "The 'psycopg2-binary' package is required for PostgreSQL. Add it to requirements.txt and pip install it."
            conn = psycopg2.connect(
                host=host, port=port, dbname=database, user=username, password=password, connect_timeout=5,
            )
            conn.close()
        elif engine == "mysql":
            try:
                import pymysql
            except ImportError:
                return False, "The 'pymysql' package is required for MySQL/MariaDB. Add it to requirements.txt and pip install it."
            conn = pymysql.connect(
                host=host, port=int(port), database=database, user=username, password=password, connect_timeout=5,
            )
            conn.close()
        else:
            return False, f"Unknown engine: {engine}"
    except Exception as exc:  # noqa: BLE001
        return False, f"Connection failed: {exc}"

    return True, f"Connected to {ENGINE_LABELS.get(engine, engine)} at {host}:{port}/{database} successfully."
