"""
SUPPLIER PERFORMANCE ANALYSIS — DATA LAYER TESTS
Module: tests/test_data.py
Covers: TC-T-01, TC-T-02, TC-T-03, TC-T-04, TC-T-05
"""

import os
import sqlite3
import pytest
import pandas as pd

DB_PATH = os.path.join('data', 'app.db')

@pytest.fixture(scope="module")
def db_conn():
    """Provides a SQLite connection with foreign keys enabled."""
    assert os.path.exists(DB_PATH), f"Database file missing at {DB_PATH}"
    conn = sqlite3.connect(DB_PATH)
    conn.execute("PRAGMA foreign_keys = ON;")
    yield conn
    conn.close()

def test_tc_t_01_database_connection(db_conn):
    """TC-T-01: Verify successful connection and foreign key pragma enabled."""
    cursor = db_conn.cursor()
    fk_status = cursor.execute("PRAGMA foreign_keys;").fetchone()[0]
    assert fk_status == 1, "Foreign keys are not enabled in SQLite"

def test_tc_t_02_table_row_and_column_counts(db_conn):
    """TC-T-02: Verify expected table row counts and schema dimensions."""
    cursor = db_conn.cursor()
    
    # Check minimum row counts
    suppliers_count = cursor.execute("SELECT COUNT(*) FROM suppliers").fetchone()[0]
    assert suppliers_count >= 25, f"Expected >= 25 suppliers, got {suppliers_count}"
    
    categories_count = cursor.execute("SELECT COUNT(*) FROM categories").fetchone()[0]
    assert categories_count == 6, f"Expected exactly 6 categories, got {categories_count}"
    
    products_count = cursor.execute("SELECT COUNT(*) FROM products").fetchone()[0]
    assert products_count >= 40, f"Expected >= 40 products, got {products_count}"
    
    orders_count = cursor.execute("SELECT COUNT(*) FROM purchase_orders").fetchone()[0]
    assert orders_count >= 300, f"Expected >= 300 purchase orders, got {orders_count}"
    
    deliveries_count = cursor.execute("SELECT COUNT(*) FROM order_deliveries").fetchone()[0]
    assert deliveries_count >= 300, f"Expected >= 300 order deliveries, got {deliveries_count}"

def test_tc_t_03_primary_key_uniqueness(db_conn):
    """TC-T-03: Verify primary key uniqueness across all tables."""
    cursor = db_conn.cursor()
    tables_pk = [
        ("suppliers", "supplier_id"),
        ("categories", "category_id"),
        ("products", "product_id"),
        ("purchase_orders", "order_id"),
        ("order_deliveries", "delivery_id")
    ]
    for table, pk in tables_pk:
        total = cursor.execute(f"SELECT COUNT(*) FROM {table}").fetchone()[0]
        distinct_pks = cursor.execute(f"SELECT COUNT(DISTINCT {pk}) FROM {table}").fetchone()[0]
        assert total == distinct_pks, f"Duplicate primary keys detected in {table} on {pk}"

def test_tc_t_04_foreign_key_referential_integrity(db_conn):
    """TC-T-04: Execute PRAGMA foreign_key_check to confirm zero orphan records."""
    cursor = db_conn.cursor()
    violations = cursor.execute("PRAGMA foreign_key_check;").fetchall()
    assert len(violations) == 0, f"Foreign key integrity violations detected: {violations}"

def test_tc_t_05_analytical_view_columns(db_conn):
    """TC-T-05: Verify unified analytical view exists and has 40+ columns."""
    cursor = db_conn.cursor()
    cols = cursor.execute("PRAGMA table_info(v_supplier_performance_analysis);").fetchall()
    assert len(cols) >= 40, f"Expected >= 40 columns in view, got {len(cols)}"
    
    row_count = cursor.execute("SELECT COUNT(*) FROM v_supplier_performance_analysis;").fetchone()[0]
    assert row_count >= 300, f"Expected >= 300 rows in analytical view, got {row_count}"
