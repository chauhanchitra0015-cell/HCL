"""
SUPPLIER PERFORMANCE ANALYSIS — DATA VALIDATION & HYGIENE ENGINE
Module: src/validator.py
Performs automated data validation, missing value handling, type casting, schema aliasing, and anomaly audits.
"""

import pandas as pd
import numpy as np

def validate_and_clean_dataset(df: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    """
    Cleans raw merged procurement analytical dataset:
    1. Casts date columns to datetime.
    2. Flags in-transit records where actual_delivery_date is null.
    3. Handles potential duplicate records.
    4. Enforces bidirectional schema aliases for seamless UI/Analytics mapping.
    5. Computes data hygiene metrics.
    
    Returns:
        cleaned_df: Cleaned pandas DataFrame.
        report: Dictionary summarizing data hygiene status.
    """
    df = df.copy()
    initial_rows = len(df)
    
    # 1. Date conversions
    date_cols = ['order_date', 'promised_delivery_date', 'actual_delivery_date']
    for col in date_cols:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')
            
    # 2. Duplicate checking
    duplicates_count = df.duplicated(subset=['delivery_id']).sum() if 'delivery_id' in df.columns else 0
    if duplicates_count > 0:
        df = df.drop_duplicates(subset=['delivery_id'], keep='first')
        
    # 3. Handle Null actual_delivery_dates (in-transit deliveries)
    null_actual_deliveries = df['actual_delivery_date'].isna().sum() if 'actual_delivery_date' in df.columns else 0
    if 'delay_days' in df.columns:
        df['delay_days'] = df['delay_days'].fillna(0)
        
    # 4. Bidirectional Schema Aliasing (Ensures DB names and UI shorthand names both exist)
    alias_map = [
        ('inspected_quantity', 'inspected_qty'),
        ('defective_quantity', 'defective_qty'),
        ('ordered_quantity', 'ordered_qty'),
        ('delivered_quantity', 'delivered_qty'),
        ('defect_reason_code', 'defect_reason'),
        ('line_total_invoiced_spend', 'invoiced_line_total'),
        ('line_total_contracted_spend', 'contracted_line_total'),
        ('price_variance_amount', 'cost_variance_amount'),
        ('order_status', 'delivery_status'),
        ('standard_lead_time_days', 'sla_lead_time_days')
    ]
    for c_db, c_alias in alias_map:
        if c_db in df.columns and c_alias not in df.columns:
            df[c_alias] = df[c_db]
        elif c_alias in df.columns and c_db not in df.columns:
            df[c_db] = df[c_alias]

    # Inspection QC status column
    if 'qc_inspection_status' not in df.columns:
        defect_col = 'defective_quantity' if 'defective_quantity' in df.columns else 'defective_qty'
        if defect_col in df.columns:
            df['qc_inspection_status'] = df[defect_col].apply(lambda x: 'REJECTED' if x > 0 else 'ACCEPTED')
        else:
            df['qc_inspection_status'] = 'ACCEPTED'

    # 5. Fill numeric nulls with safe defaults
    numeric_fills = {
        'defective_qty': 0,
        'defective_quantity': 0,
        'inspected_qty': 0,
        'inspected_quantity': 0,
        'ordered_qty': 0,
        'ordered_quantity': 0,
        'delivered_qty': 0,
        'delivered_quantity': 0,
        'defect_ppm': 0,
        'is_on_time': 0,
        'is_in_full': 0,
        'is_otif': 0,
        'price_variance_pct': 0.0,
        'cost_variance_amount': 0.0,
        'price_variance_amount': 0.0,
        'invoiced_line_total': 0.0,
        'line_total_invoiced_spend': 0.0,
        'contracted_line_total': 0.0,
        'line_total_contracted_spend': 0.0
    }
    for col, default_val in numeric_fills.items():
        if col in df.columns:
            df[col] = df[col].fillna(default_val)
            
    # 6. Build Hygiene Report
    report = {
        "initial_rows": initial_rows,
        "clean_rows": len(df),
        "duplicates_removed": int(duplicates_count),
        "in_transit_records": int(null_actual_deliveries),
        "columns_validated": len(df.columns),
        "missing_values_by_column": df.isna().sum().to_dict()
    }
    
    return df, report
