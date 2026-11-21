"""
SUPPLIER PERFORMANCE ANALYSIS — DATA EXPLORER & EXPORT (UI-06)
Author: Chitra Chauhan (Roll No: 2500320140043, MCA 2nd Year DA, ABES EC)
Mentor: Mr. Dinesh
"""

import streamlit as st
import pandas as pd

from src.data_loader import load_analytical_data
from src.ui import render_tailwind_header
from src.utils import to_csv_bytes, to_excel_bytes

st.set_page_config(page_title="Data Explorer & Export", page_icon="🔍", layout="wide")

df_clean, _ = load_analytical_data()

render_tailwind_header(
    title="Interactive Procurement Data Explorer",
    subtitle="Unified 42-column analytical view with granular search, dimension filtering, and master data export",
    badge_text="UI-06 • DATA REPOSITORY"
)

st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)

# 1. Search and Column Selection
col_srch, col_cols = st.columns([1, 2])

with col_srch:
    search_query = st.text_input("🔍 Quick Text Search", placeholder="Supplier name, PO ID, or Product...")

with col_cols:
    default_selected = [
        'delivery_id', 'order_id', 'supplier_name', 'category_name', 'product_name',
        'order_date', 'promised_delivery_date', 'actual_delivery_date',
        'is_otif', 'defect_ppm', 'price_variance_pct', 'invoiced_line_total'
    ]
    # Filter available columns
    available_cols = list(df_clean.columns)
    valid_defaults = [c for c in default_selected if c in available_cols]
    
    selected_columns = st.multiselect(
        "Select Columns to Display",
        options=available_cols,
        default=valid_defaults
    )

# 2. Filter Application
filtered_df = df_clean.copy()

if search_query:
    q = search_query.lower()
    mask = (
        filtered_df['supplier_name'].astype(str).str.lower().str.contains(q) |
        filtered_df['order_id'].astype(str).str.lower().str.contains(q) |
        filtered_df['product_name'].astype(str).str.lower().str.contains(q) |
        filtered_df['category_name'].astype(str).str.lower().str.contains(q)
    )
    filtered_df = filtered_df[mask]

# Summary metric
st.caption(f"Showing **{len(filtered_df):,}** matching rows across **{len(selected_columns)}** selected columns.")

# 3. Data Table Display
if len(selected_columns) > 0 and len(filtered_df) > 0:
    st.dataframe(filtered_df[selected_columns], use_container_width=True, hide_index=True)
elif len(selected_columns) == 0:
    st.warning("Please select at least one column to display.")
else:
    st.info("No records match the active search criteria.")

# 4. Data Export Buttons
st.markdown("---")
st.subheader("📥 Export Filtered Analytical Records")

col_d1, col_d2, _ = st.columns([1, 1, 3])

with col_d1:
    st.download_button(
        label="Download as CSV",
        data=to_csv_bytes(filtered_df[selected_columns] if len(selected_columns) > 0 else filtered_df),
        file_name="procurement_analytical_data.csv",
        mime="text/csv"
    )

with col_d2:
    st.download_button(
        label="Download as Excel (.xlsx)",
        data=to_excel_bytes(filtered_df[selected_columns] if len(selected_columns) > 0 else filtered_df),
        file_name="procurement_analytical_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
