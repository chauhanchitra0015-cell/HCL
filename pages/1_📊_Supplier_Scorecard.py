"""
SUPPLIER PERFORMANCE ANALYSIS — SCORECARD & CPI TIERING (UI-02)
Author: Chitra Chauhan (Roll No: 2500320140043, MCA 2nd Year DA, ABES EC)
Mentor: Mr. Dinesh
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

from src.data_loader import load_analytical_data
from src.analytics import compute_supplier_cpi_table, compute_spend_at_risk
from src.ui import render_tailwind_header, render_kpi_cards, render_insight_box
from src.utils import format_inr, to_csv_bytes, to_excel_bytes

st.set_page_config(page_title="Supplier Scorecard & CPI", page_icon="📊", layout="wide")

# Load Data
df_clean, _ = load_analytical_data()

render_tailwind_header(
    title="Multi-Metric Supplier Scorecard & CPI Tiering",
    subtitle="Composite Performance Index (CPI) evaluation across Delivery, Quality, and Cost dimensions",
    badge_text="UI-02 • SCORECARD ENGINE"
)

st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)

# Weight Customization Expander
with st.expander("⚙️ Customize CPI Weight Sensitivities (Contract Governance)", expanded=False):
    st.caption("Adjust relative importance of fulfillment punctuality, receiving quality, and price variance.")
    col_w1, col_w2, col_w3 = st.columns(3)
    w_del = col_w1.slider("Delivery Weight (OTIF %)", min_value=0.10, max_value=0.80, value=0.40, step=0.05)
    w_qual = col_w2.slider("Quality Weight (PPM)", min_value=0.10, max_value=0.80, value=0.35, step=0.05)
    w_cost = col_w3.slider("Cost Weight (Price Variance)", min_value=0.10, max_value=0.80, value=0.25, step=0.05)
    
    # Normalize weights
    total_w = w_del + w_qual + w_cost
    w_del_norm = round(w_del / total_w, 2)
    w_qual_norm = round(w_qual / total_w, 2)
    w_cost_norm = round(1.0 - (w_del_norm + w_qual_norm), 2)
    st.info(f"Normalized Weights Applied: Delivery = **{w_del_norm*100:.0f}%** | Quality = **{w_qual_norm*100:.0f}%** | Cost = **{w_cost_norm*100:.0f}%**")

# Calculate CPI Table
cpi_df = compute_supplier_cpi_table(df_clean, w_delivery=w_del_norm, w_quality=w_qual_norm, w_cost=w_cost_norm)

tier_a_count = len(cpi_df[cpi_df['performance_tier'].str.contains("Tier A")])
tier_b_count = len(cpi_df[cpi_df['performance_tier'].str.contains("Tier B")])
tier_c_count = len(cpi_df[cpi_df['performance_tier'].str.contains("Tier C")])
spend_at_risk = compute_spend_at_risk(cpi_df)

# KPI Cards
scorecard_kpis = [
    {
        "label": "Tier A (Preferred)",
        "value": str(tier_a_count),
        "benchmark": "CPI ≥ 85.0",
        "status": "success",
        "subtext": "Automatic contract renewal eligible"
    },
    {
        "label": "Tier B (Conditional)",
        "value": str(tier_b_count),
        "benchmark": "70.0 ≤ CPI < 85.0",
        "status": "warning",
        "subtext": "Active monitoring required"
    },
    {
        "label": "Tier C (At-Risk)",
        "value": str(tier_c_count),
        "benchmark": "CPI < 70.0",
        "status": "danger" if tier_c_count > 0 else "success",
        "subtext": "Mandatory CAPA audit triggered"
    },
    {
        "label": "Spend-at-Risk %",
        "value": f"{spend_at_risk:.1f}%",
        "benchmark": "Target < 10.0%",
        "status": "success" if spend_at_risk <= 10.0 else "danger",
        "subtext": "Procurement capital exposed"
    }
]

render_kpi_cards(scorecard_kpis)

st.markdown("<div style='height: 15px;'></div>", unsafe_allow_html=True)

# Charts Section: Tier Breakdown & Radar Chart
col_pie, col_radar = st.columns([1, 2])

with col_pie:
    st.subheader("📊 Supplier Tier Distribution")
    tier_counts = cpi_df['performance_tier'].value_counts().reset_index()
    tier_counts.columns = ['Tier', 'Suppliers']
    
    fig_pie = px.pie(
        tier_counts,
        names='Tier',
        values='Suppliers',
        color='Tier',
        color_discrete_map={
            "Tier A (Preferred)": "#10B981",
            "Tier B (Conditional)": "#F59E0B",
            "Tier C (At-Risk)": "#EF4444"
        },
        hole=0.4
    )
    fig_pie.update_layout(margin=dict(l=10, r=10, t=10, b=10), height=320)
    st.plotly_chart(fig_pie, use_container_width=True)

with col_radar:
    st.subheader("🎯 Vendor Multi-Dimensional Radar Comparison")
    selected_for_radar = st.multiselect(
        "Select up to 3 suppliers to benchmark against average:",
        options=cpi_df['supplier_name'].unique().tolist(),
        default=cpi_df['supplier_name'].head(2).tolist()
    )
    
    categories_radar = ['Delivery Score', 'Quality Score', 'Cost Score']
    fig_radar = go.Figure()
    
    # Add Average baseline
    avg_del = cpi_df['delivery_score'].mean()
    avg_qual = cpi_df['quality_score'].mean()
    avg_cost = cpi_df['cost_score'].mean()
    fig_radar.add_trace(go.Scatterpolar(
        r=[avg_del, avg_qual, avg_cost],
        theta=categories_radar,
        fill='toself',
        name='Portfolio Average',
        line=dict(color='#94A3B8', dash='dash')
    ))
    
    colors = ['#1A365D', '#D69E2E', '#E53E3E']
    for idx, s_name in enumerate(selected_for_radar[:3]):
        s_row = cpi_df[cpi_df['supplier_name'] == s_name].iloc[0]
        fig_radar.add_trace(go.Scatterpolar(
            r=[s_row['delivery_score'], s_row['quality_score'], s_row['cost_score']],
            theta=categories_radar,
            fill='toself',
            name=s_name,
            line=dict(color=colors[idx % len(colors)], width=2)
        ))
        
    fig_radar.update_layout(
        polar=dict(radialaxis=dict(visible=True, range=[0, 100])),
        margin=dict(l=40, r=40, t=20, b=20),
        height=320,
        showlegend=True
    )
    st.plotly_chart(fig_radar, use_container_width=True)

# Master Scorecard Table
st.subheader("📋 Comprehensive Supplier Scorecard")

display_cols = [
    'supplier_name', 'category_name', 'total_deliveries', 'otif_pct',
    'defect_ppm', 'price_variance_pct', 'delivery_score', 'quality_score',
    'cost_score', 'cpi_score', 'performance_tier', 'capa_status'
]

formatted_df = cpi_df[display_cols].rename(columns={
    'supplier_name': 'Supplier Name',
    'category_name': 'Category',
    'total_deliveries': 'Deliveries',
    'otif_pct': 'OTIF %',
    'defect_ppm': 'Defect PPM',
    'price_variance_pct': 'Price Var %',
    'delivery_score': 'Del Score',
    'quality_score': 'Qual Score',
    'cost_score': 'Cost Score',
    'cpi_score': 'CPI Score',
    'performance_tier': 'Tier',
    'capa_status': 'Governance Status'
})

st.dataframe(
    formatted_df.style.format({
        'OTIF %': '{:.1f}%',
        'Defect PPM': '{:,}',
        'Price Var %': '{:+.1f}%',
        'Del Score': '{:.1f}',
        'Qual Score': '{:.1f}',
        'Cost Score': '{:.1f}',
        'CPI Score': '{:.1f}'
    }),
    use_container_width=True,
    hide_index=True
)

# Export Buttons
col_exp1, col_exp2, _ = st.columns([1, 1, 3])
with col_exp1:
    st.download_button(
        "📥 Download Scorecard (CSV)",
        data=to_csv_bytes(formatted_df),
        file_name="supplier_scorecard.csv",
        mime="text/csv"
    )
with col_exp2:
    st.download_button(
        "📥 Download Scorecard (Excel)",
        data=to_excel_bytes(formatted_df, sheet_name="Scorecard"),
        file_name="supplier_scorecard.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# CAPA Triggers Section
st.markdown("---")
st.subheader("🚨 Corrective Action Plan (CAPA) Triggers")
tier_c_vendors = cpi_df[cpi_df['performance_tier'].str.contains("Tier C")]

if len(tier_c_vendors) > 0:
    for _, vendor in tier_c_vendors.iterrows():
        st.warning(f"""
        **Supplier:** {vendor['supplier_name']} ({vendor['category_name']}) | **CPI Score:** {vendor['cpi_score']:.1f}  
        * **Identified Deficiencies:** OTIF = {vendor['otif_pct']:.1f}% (Benchmark ≥ 95%), Defect PPM = {vendor['defect_ppm']:,} (Benchmark < 1,000), Price Variance = {vendor['price_variance_pct']:+.1f}%  
        * **Mandatory Action:** Formal CAPA issued. Vendor must submit 8D Root Cause Analysis within 14 business days. Contract renewal frozen.
        """)
else:
    st.success("✅ All suppliers are currently operating within acceptable performance bounds (Tier A or Tier B). No active CAPA triggers.")
