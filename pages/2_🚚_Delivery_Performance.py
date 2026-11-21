"""
SUPPLIER PERFORMANCE ANALYSIS — DELIVERY & LEAD TIME PERFORMANCE (UI-03)
Author: Chitra Chauhan (Roll No: 2500320140043, MCA 2nd Year DA, ABES EC)
Mentor: Mr. Dinesh
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

from src.data_loader import load_analytical_data
from src.analytics import compute_lead_time_distribution
from src.ui import render_tailwind_header, render_kpi_cards, render_insight_box
from src.config import TARGET_OTIF_PCT, SEVERE_DELAY_DAYS

st.set_page_config(page_title="Delivery Performance", page_icon="🚚", layout="wide")

df_clean, _ = load_analytical_data()

render_tailwind_header(
    title="Delivery Adherence & Lead Time Analytics",
    subtitle="Empirical tracking of OTIF punctuality, SLA compliance, and logistics distribution tails",
    badge_text="UI-03 • LOGISTICS FULFILLMENT"
)

st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)

# 1. Delivery KPIs
total_del = len(df_clean)
otif_pct = (df_clean['is_otif'].sum() / total_del * 100.0) if total_del > 0 else 0
ontime_pct = (df_clean['is_on_time'].sum() / total_del * 100.0) if total_del > 0 else 0
infull_pct = (df_clean['is_in_full'].sum() / total_del * 100.0) if total_del > 0 else 0
avg_lead = df_clean['delivery_lead_time_days'].dropna().mean() if 'delivery_lead_time_days' in df_clean else 0

delivery_kpis = [
    {
        "label": "Macro OTIF %",
        "value": f"{otif_pct:.1f}%",
        "benchmark": f"Target ≥ {TARGET_OTIF_PCT:.0f}%",
        "status": "success" if otif_pct >= TARGET_OTIF_PCT else "warning",
        "subtext": "On-Time AND In-Full standard"
    },
    {
        "label": "On-Time Rate %",
        "value": f"{ontime_pct:.1f}%",
        "benchmark": "Punctuality alone",
        "status": "success" if ontime_pct >= 95.0 else "warning",
        "subtext": f"{df_clean['is_on_time'].sum()} orders arrived on/before SLA"
    },
    {
        "label": "In-Full Rate %",
        "value": f"{infull_pct:.1f}%",
        "benchmark": "Quantity fulfillment",
        "status": "success" if infull_pct >= 98.0 else "neutral",
        "subtext": f"{df_clean['is_in_full'].sum()} orders delivered 100% quantity"
    },
    {
        "label": "Mean Lead Time",
        "value": f"{avg_lead:.1f} Days",
        "benchmark": "Contract SLA ~14d",
        "status": "neutral",
        "subtext": "From PO creation to dock receipt"
    }
]

render_kpi_cards(delivery_kpis)

st.markdown("<div style='height: 15px;'></div>", unsafe_allow_html=True)

# 2. Charts: OTIF Leaderboard & Category Box Plot
col_bar, col_box = st.columns([1, 1])

with col_bar:
    st.subheader("🏁 OTIF Leaderboard by Supplier")
    sup_otif = df_clean.groupby(['supplier_name', 'category_name']).agg(
        total=('delivery_id', 'count'),
        otif=('is_otif', 'sum')
    ).reset_index()
    sup_otif['otif_pct'] = (sup_otif['otif'] / sup_otif['total'] * 100.0).round(1)
    sup_otif = sup_otif.sort_values(by='otif_pct', ascending=True)

    fig_bar = px.bar(
        sup_otif,
        x='otif_pct',
        y='supplier_name',
        orientation='h',
        color='otif_pct',
        color_continuous_scale=['#EF4444', '#F59E0B', '#10B981'],
        labels={'otif_pct': 'OTIF Percentage (%)', 'supplier_name': 'Supplier'},
        text='otif_pct'
    )
    fig_bar.add_vline(x=TARGET_OTIF_PCT, line_dash="dash", line_color="#1E293B", annotation_text="95% SLA Target")
    fig_bar.update_layout(margin=dict(l=20, r=20, t=20, b=20), height=420, coloraxis_showscale=False)
    st.plotly_chart(fig_bar, use_container_width=True)

with col_box:
    st.subheader("📦 Lead Time Distribution by Category")
    fig_box = px.box(
        df_clean.dropna(subset=['delivery_lead_time_days']),
        x='category_name',
        y='delivery_lead_time_days',
        color='category_name',
        points='outliers',
        labels={'delivery_lead_time_days': 'Lead Time (Days)', 'category_name': 'Category'}
    )
    fig_box.update_layout(
        margin=dict(l=20, r=20, t=20, b=20),
        height=420,
        showlegend=False,
        template="plotly_white"
    )
    st.plotly_chart(fig_box, use_container_width=True)

# 3. Delay Tail Outliers (Tukey's Fences)
st.markdown("---")
st.subheader("⚠️ Severe Delivery Delay Outliers (Tukey's IQR Method)")

lead_summary, outliers_df = compute_lead_time_distribution(df_clean)

col_sum, col_out = st.columns([1, 2])

with col_sum:
    st.write("**Category Statistical Dispersion:**")
    st.dataframe(
        lead_summary[['category', 'mean_days', 'median_days', 'iqr_days', 'upper_fence', 'outliers_count']],
        use_container_width=True,
        hide_index=True
    )

with col_out:
    st.write(f"**Identified Outliers Exceeding Category Fences ({len(outliers_df)} Deliveries):**")
    if len(outliers_df) > 0:
        st.dataframe(
            outliers_df[['delivery_id', 'supplier_name', 'category_name', 'promised_delivery_date', 'actual_delivery_date', 'delivery_lead_time_days', 'delay_days']].sort_values(by='delay_days', ascending=False),
            use_container_width=True,
            hide_index=True
        )
    else:
        st.success("No statistical outliers detected in delivery lead times.")
