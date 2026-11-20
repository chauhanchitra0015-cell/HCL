"""
SUPPLIER PERFORMANCE ANALYSIS — CACHED DATA LOADER PIPELINE
Module: src/data_loader.py
Implements cached data extraction pipelines with Streamlit caching.
"""

import pandas as pd
from typing import Tuple, Dict, List, Optional
try:
    import streamlit as st
    cache_decorator = st.cache_data
except Exception:
    # Fallback decorator for headless testing or non-Streamlit runtime
    def cache_decorator(func):
        return func

from src.database import query_to_dataframe
from src.validator import validate_and_clean_dataset

@cache_decorator
def load_analytical_data() -> Tuple[pd.DataFrame, Dict]:
    """
    Loads unified analytical dataset from SQLite view v_supplier_performance_analysis
    and applies cleaning and validation.
    """
    query = "SELECT * FROM v_supplier_performance_analysis;"
    raw_df = query_to_dataframe(query)
    clean_df, report = validate_and_clean_dataset(raw_df)
    return clean_df, report

@cache_decorator
def load_entity_tables() -> Dict[str, pd.DataFrame]:
    """
    Loads raw relational tables from SQLite.
    """
    tables = ["suppliers", "categories", "products", "purchase_orders", "order_deliveries"]
    data = {}
    for table in tables:
        data[table] = query_to_dataframe(f"SELECT * FROM {table};")
    return data

def get_filter_bounds(df: pd.DataFrame) -> Dict:
    """
    Extracts available unique filter bounds from DataFrame.
    """
    categories = sorted(df['category_name'].dropna().unique().tolist()) if 'category_name' in df.columns else []
    suppliers = sorted(df['supplier_name'].dropna().unique().tolist()) if 'supplier_name' in df.columns else []
    statuses = sorted(df['delivery_status'].dropna().unique().tolist()) if 'delivery_status' in df.columns else []
    
    min_date = df['order_date'].min() if 'order_date' in df.columns else None
    max_date = df['order_date'].max() if 'order_date' in df.columns else None
    
    return {
        "categories": categories,
        "suppliers": suppliers,
        "statuses": statuses,
        "min_date": min_date,
        "max_date": max_date
    }

def filter_dataframe(
    df: pd.DataFrame,
    categories: Optional[List[str]] = None,
    suppliers: Optional[List[str]] = None,
    statuses: Optional[List[str]] = None,
    date_range: Optional[Tuple] = None
) -> pd.DataFrame:
    """
    Applies unified multi-dimension filtering to analytical dataset.
    """
    filtered = df.copy()
    
    if categories:
        filtered = filtered[filtered['category_name'].isin(categories)]
        
    if suppliers:
        filtered = filtered[filtered['supplier_name'].isin(suppliers)]
        
    if statuses:
        filtered = filtered[filtered['delivery_status'].isin(statuses)]
        
    if date_range and len(date_range) == 2 and date_range[0] and date_range[1]:
        start_date, end_date = pd.to_datetime(date_range[0]), pd.to_datetime(date_range[1])
        filtered = filtered[(filtered['order_date'] >= start_date) & (filtered['order_date'] <= end_date)]
        
    return filtered
