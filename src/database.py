"""
SUPPLIER PERFORMANCE ANALYSIS — DATABASE CONNECTION MANAGER
Module: src/database.py
Provides robust, thread-safe SQLite connectivity with foreign key pragma enforcement.
"""

import sqlite3
import pandas as pd
from typing import Optional
from src.config import DB_PATH

def get_connection(db_path: Optional[str] = None) -> sqlite3.Connection:
    """
    Returns an active SQLite connection with Foreign Keys enabled.
    """
    path = db_path if db_path else str(DB_PATH)
    conn = sqlite3.connect(path, timeout=10.0)
    conn.execute("PRAGMA foreign_keys = ON;")
    conn.row_factory = sqlite3.Row
    return conn

def execute_query(query: str, params: tuple = (), db_path: Optional[str] = None) -> list:
    """
    Executes a read query and returns rows as dictionaries.
    """
    with get_connection(db_path) as conn:
        cursor = conn.cursor()
        cursor.execute(query, params)
        rows = cursor.fetchall()
        return [dict(row) for row in rows]

def query_to_dataframe(query: str, params: tuple = (), db_path: Optional[str] = None) -> pd.DataFrame:
    """
    Executes a SQL query and returns the results as a pandas DataFrame.
    """
    with get_connection(db_path) as conn:
        df = pd.read_sql_query(query, conn, params=params)
        return df

def check_db_health(db_path: Optional[str] = None) -> dict:
    """
    Performs quick database diagnostic checks (FK check, integrity check, row counts).
    """
    with get_connection(db_path) as conn:
        cursor = conn.cursor()
        fk_violations = cursor.execute("PRAGMA foreign_key_check;").fetchall()
        integrity = cursor.execute("PRAGMA integrity_check;").fetchone()[0]
        
        tables = ["suppliers", "categories", "products", "purchase_orders", "order_deliveries"]
        counts = {}
        for t in tables:
            try:
                cnt = cursor.execute(f"SELECT COUNT(*) FROM {t};").fetchone()[0]
                counts[t] = cnt
            except Exception:
                counts[t] = 0
                
        return {
            "integrity_check": integrity,
            "fk_violations_count": len(fk_violations),
            "table_counts": counts,
            "status": "HEALTHY" if integrity == "ok" and len(fk_violations) == 0 else "WARNING"
        }
