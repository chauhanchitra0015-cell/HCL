"""
SUPPLIER PERFORMANCE ANALYSIS — DATA VALIDATION & HYGIENE ENGINE
Module: src/validator.py
Performs automated data validation, missing value handling, type casting, and anomaly audits.
"""

import pandas as pd
import numpy as np

def validate_and_clean_dataset(df: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    """
    Cleans raw merged procurement analytical dataset:
    1. Casts date columns to datetime.
    2. Flags in-transit records where actual_delivery_date is null.
    3. Handles potential duplicate records.
    4. Computes data hygiene metrics.
    
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
    # For in-transit orders, delivery_lead_time_days is null; fill delay_days with 0 for KPI safety if null
    if 'delay_days' in df.columns:
        df['delay_days'] = df['delay_days'].fillna(0)
        
    # 4. Fill numeric nulls with safe defaults
    numeric_fills = {
        'defective_qty': 0,
        'inspected_qty': 0,
        'defect_ppm': 0,
        'is_on_time': 0,
        'is_in_full': 0,
        'is_otif': 0,
        'price_variance_pct': 0.0,
        'cost_variance_amount': 0.0
    }
    for col, default_val in numeric_fills.items():
        if col in df.columns:
            df[col] = df[col].fillna(default_val)
            
    # 5. Build Hygiene Report
    report = {
        "initial_rows": initial_rows,
        "clean_rows": len(df),
        "duplicates_removed": int(duplicates_count),
        "in_transit_records": int(null_actual_deliveries),
        "columns_validated": len(df.columns),
        "missing_values_by_column": df.isna().sum().to_dict()
    }
    
    return df, report
