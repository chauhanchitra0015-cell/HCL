import sys
import os
sys.path.insert(0, os.path.abspath('.'))
import sqlite3
import pandas as pd
from src.data_loader import load_analytical_data
from src.analytics import (
    compute_kpi_summary,
    compute_supplier_cpi_table,
    compute_spend_at_risk,
    compute_defect_pareto,
    compute_lead_time_distribution,
    compute_defect_ppm
)

print("1. Testing load_analytical_data()...")
df, report = load_analytical_data()
print(f"Loaded {len(df)} rows and {len(df.columns)} columns.")

print("2. Testing Executive Overview KPIs...")
kpis = compute_kpi_summary(df)
print("Executive KPIs:", kpis)

print("3. Testing Scorecard & CPI table...")
cpi_df = compute_supplier_cpi_table(df)
print(f"Computed CPI for {len(cpi_df)} suppliers. Top score: {cpi_df.iloc[0]['cpi_score']}")

print("4. Testing Quality Pareto...")
pareto = compute_defect_pareto(df)
print(f"Pareto computed for {len(pareto)} defect categories.")

print("5. Testing Lead Time distribution...")
summary, outliers = compute_lead_time_distribution(df)
print(f"Lead time stats: {len(summary)} categories, {len(outliers)} outliers.")

print("6. Testing Quality screen data slicing...")
total_inspected = int(df['inspected_qty'].sum())
total_defective = int(df['defective_qty'].sum())
macro_ppm = compute_defect_ppm(total_inspected, total_defective)
quarantine_count = len(df[df['defective_qty'] > 0])
quarantined_orders = df[df['defective_qty'] > 0][
    ['delivery_id', 'order_id', 'supplier_name', 'product_name', 'delivered_qty', 'inspected_qty', 'defective_qty', 'defect_ppm', 'defect_reason', 'qc_inspection_status']
]
print(f"Quality screen slice successful! Quarantined orders count: {len(quarantined_orders)}")

print("ALL SCREENS LOGIC VALIDATED WITH ZERO ERRORS!")
