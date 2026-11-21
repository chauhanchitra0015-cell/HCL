"""
SUPPLIER PERFORMANCE ANALYSIS — RECEIVING QUALITY & PARETO RISK (UI-04)
Author: Chitra Chauhan (Roll No: 2500320140043, MCA 2nd Year DA, ABES EC)
Mentor: Mr. Dinesh
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

from src.data_loader import load_analytical_data
from src.analytics import compute_defect_ppm, compute_defect_pareto
from src.ui import render_tailwind_header, render_kpi_cards, render_insight_box
from src.config import MAX_DEFECT_PPM

st.set_page_config(page_title="Quality & Defect Risk", page_icon="🔬", layout="wide")

df_clean, _ = load_analytical_data()

render_tailwind_header(
    title="Receiving Quality & Defect Pareto Risk",
    subtitle="Statistical quality control, defect PPM tracking, and Pareto 80/20 root cause prioritization",
    badge_text="UI-04 • QUALITY ASSURANCE"
)

st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)

# 1. Quality KPIs
total_inspected = int(df_clean['inspected_qty'].sum())
total_defective = int(df_clean['defective_qty'].sum())
macro_ppm = compute_defect_ppm(total_inspected, total_defective)
quarantine_count = len(df_clean[df_clean['defective_qty'] > 0])

quality_kpis = [
    {
        "label": "Receiving Defect PPM",
        "value": f"{macro_ppm:,}",
        "benchmark": f"Target < {MAX_DEFECT_PPM:,} PPM",
        "status": "success" if macro_ppm <= MAX_DEFECT_PPM else "danger",
        "subtext": "Parts Per Million non-conforming"
    },
    {
        "label": "Total Inspected Units",
        "value": f"{total_inspected:,}",
        "benchmark": "100% Receiving QA",
        "status": "neutral",
        "subtext": "Components checked at receipt"
    },
    {
        "label": "Scrapped / Defective Units",
        "value": f"{total_defective:,}",
        "benchmark": "Physical rejects",
        "status": "warning" if total_defective > 0 else "success",
        "subtext": "Failed dimensional/electrical QA"
    },
    {
        "label": "Quarantined Deliveries",
        "value": str(quarantine_count),
        "benchmark": "Containment lots",
        "status": "warning" if quarantine_count > 0 else "success",
        "subtext": "Lots requiring QA review"
    }
]

render_kpi_cards(quality_kpis)

st.markdown("<div style='height: 15px;'></div>", unsafe_allow_html=True)

# 2. Charts: Dual-Axis Pareto 80/20 & Defect PPM by Vendor
col_pareto, col_ppm = st.columns([1, 1])

with col_pareto:
    st.subheader("📊 Defect Root Cause Pareto (80/20 Rule)")
    pareto_df = compute_defect_pareto(df_clean)
    
    if len(pareto_df) > 0:
        fig_pareto = go.Figure()
        # Bar trace
        fig_pareto.add_trace(go.Bar(
            x=pareto_df['defect_reason'],
            y=pareto_df['defect_count'],
            name='Defect Units',
            marker_color='#1A365D',
            yaxis='y'
        ))
        # Cumulative line trace
        fig_pareto.add_trace(go.Scatter(
            x=pareto_df['defect_reason'],
            y=pareto_df['cumulative_share_pct'],
            name='Cumulative %',
            marker_color='#EF4444',
            line=dict(width=2.5),
            mode='lines+markers',
            yaxis='y2'
        ))
        # 80% reference line
        fig_pareto.add_hline(
            y=80.0,
            yref='y2',
            line_dash="dot",
            line_color="#F59E0B",
            annotation_text="80% Pareto Cutoff",
            annotation_position="bottom left"
        )
        fig_pareto.update_layout(
            yaxis=dict(title="Defective Units (Count)", showgrid=False),
            yaxis2=dict(title="Cumulative Share (%)", overlaying='y', side='right', range=[0, 105], showgrid=True),
            xaxis=dict(title="Defect Classification"),
            margin=dict(l=30, r=40, t=20, b=40),
            height=380,
            template="plotly_white",
            legend=dict(x=0.02, y=0.95)
        )
        st.plotly_chart(fig_pareto, use_container_width=True)
    else:
        st.info("No recorded defects in dataset.")

with col_ppm:
    st.subheader("🎯 Vendor Defect PPM Distribution")
    vendor_ppm = df_clean.groupby(['supplier_name', 'category_name']).agg(
        inspected=('inspected_qty', 'sum'),
        defective=('defective_qty', 'sum')
    ).reset_index()
    vendor_ppm['ppm'] = vendor_ppm.apply(lambda r: compute_defect_ppm(r['inspected'], r['defective']), axis=1)
    vendor_ppm = vendor_ppm.sort_values(by='ppm', ascending=False)

    fig_ppm = px.bar(
        vendor_ppm.head(10),
        x='ppm',
        y='supplier_name',
        orientation='h',
        color='ppm',
        color_continuous_scale=['#10B981', '#F59E0B', '#EF4444'],
        labels={'ppm': 'Defect PPM', 'supplier_name': 'Supplier'},
        text='ppm'
    )
    fig_ppm.add_vline(x=MAX_DEFECT_PPM, line_dash="dash", line_color="#1E293B", annotation_text="1,000 PPM Limit")
    fig_ppm.update_layout(margin=dict(l=20, r=20, t=20, b=20), height=380, coloraxis_showscale=False)
    st.plotly_chart(fig_ppm, use_container_width=True)

# 3. Defect Quarantine Audit Log
st.markdown("---")
st.subheader("📋 Component Quarantine & Inspection Audit Log")

quarantined_orders = df_clean[df_clean['defective_qty'] > 0][
    ['delivery_id', 'order_id', 'supplier_name', 'product_name', 'delivered_qty', 'inspected_qty', 'defective_qty', 'defect_ppm', 'defect_reason', 'qc_inspection_status']
].sort_values(by='defective_qty', ascending=False)

if len(quarantined_orders) > 0:
    st.dataframe(
        quarantined_orders.style.format({
            'delivered_qty': '{:,}',
            'inspected_qty': '{:,}',
            'defective_qty': '{:,}',
            'defect_ppm': '{:,}'
        }),
        use_container_width=True,
        hide_index=True
    )
else:
    st.success("No defective deliveries detected.")
