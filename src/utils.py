"""
SUPPLIER PERFORMANCE ANALYSIS — UTILITIES & FORMATTING ENGINE
Module: src/utils.py
Provides formatting helpers, currency representations, and export serializers.
"""

import io
import pandas as pd

def format_inr(amount: float) -> str:
    """
    Formats a numeric amount into Indian Rupee representation (Lakhs / Crores / Standard).
    """
    if pd.isna(amount) or amount == 0:
        return "₹ 0"
    
    abs_amt = abs(amount)
    sign = "-" if amount < 0 else ""
    
    if abs_amt >= 10_000_000:
        return f"{sign}₹ {abs_amt / 10_000_000:.2f} Cr"
    elif abs_amt >= 100_000:
        return f"{sign}₹ {abs_amt / 100_000:.2f} L"
    else:
        return f"{sign}₹ {abs_amt:,.0f}"

def format_percent(val: float) -> str:
    """
    Formats a decimal or percentage value as a clean string.
    """
    if pd.isna(val):
        return "0.0%"
    return f"{val:.1f}%"

def to_csv_bytes(df: pd.DataFrame) -> bytes:
    """
    Serializes a DataFrame to CSV byte stream for Streamlit download button.
    """
    return df.to_csv(index=False).encode('utf-8')

def to_excel_bytes(df: pd.DataFrame, sheet_name: str = "Analytical_Data") -> bytes:
    """
    Serializes a DataFrame to Excel xlsx byte stream using openpyxl.
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
    return output.getvalue()
