"""
SUPPLIER PERFORMANCE ANALYSIS — COST VARIANCE & SPEND-AT-RISK (UI-05)
Author: Chitra Chauhan (Roll No: 2500320140043, MCA 2nd Year DA, ABES EC)
Mentor: Prof. Kalpana Dwivedi
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

from src.data_loader import load_analytical_data
from src.analytics import compute_supplier_cpi_table, compute_spend_at_risk, compute_price_variance_pct
from src.ui import render_tailwind_header, render_kpi_cards, render_insight_box
from src.utils import format_inr, format_percent
from src.config import MAX_SPEND_AT_RISK_PCT

st.set_page_config(page_title="Cost Variance & Spend Risk", page_icon="💰", layout="wide")

df_clean, _ = load_analytical_data()

render_tailwind_header(
    title="Cost Variance & Spend-at-Risk Governance",
    subtitle="Financial audit of Purchase Price Variance (PPV), unapproved price creep, and Tier C spend exposure",
    badge_text="UI-05 • FINANCIAL AUDIT"
)

st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)

# 1. Cost KPIs
total_contracted = df_clean['contracted_line_total'].sum()
total_invoiced = df_clean['invoiced_line_total'].sum()
net_variance_amt = total_invoiced - total_contracted
net_variance_pct = compute_price_variance_pct(total_contracted, total_invoiced)

cpi_table = compute_supplier_cpi_table(df_clean)
spend_at_risk_pct = compute_spend_at_risk(cpi_table)

cost_kpis = [
    {
        "label": "Total Invoiced Spend",
        "value": format_inr(total_invoiced),
        "benchmark": "Procurement Outflow",
        "status": "neutral",
        "subtext": "Actual disbursed spend"
    },
    {
        "label": "Net Price Variance",
        "value": f"{net_variance_pct:+.2f}%",
        "benchmark": f"{format_inr(net_variance_amt)} delta",
        "status": "danger" if net_variance_pct > 1.0 else ("warning" if net_variance_pct > 0 else "success"),
        "subtext": "Invoiced vs Contracted delta"
    },
    {
        "label": "Spend-at-Risk %",
        "value": f"{spend_at_risk_pct:.1f}%",
        "benchmark": f"Target < {MAX_SPEND_AT_RISK_PCT:.0f}%",
        "status": "success" if spend_at_risk_pct <= MAX_SPEND_AT_RISK_PCT else "danger",
        "subtext": "Spend committed to Tier C vendors"
    },
    {
        "label": "Price Creep Lines",
        "value": str(len(df_clean[df_clean['price_variance_pct'] > 0])),
        "benchmark": "Invoices with premium",
        "status": "warning",
        "subtext": "Deliveries above contracted rate"
    }
]

render_kpi_cards(cost_kpis)

st.markdown("<div style='height: 15px;'></div>", unsafe_allow_html=True)

# 2. Charts: Spend-at-Risk Gauge & Price Variance Distribution
col_gauge, col_dist = st.columns([1, 2])

with col_gauge:
    st.subheader("🎯 Spend-at-Risk Exposure Gauge")
    fig_gauge = go.Figure(go.Indicator(
        mode="gauge+number+delta",
        value=spend_at_risk_pct,
        number={'suffix': "%"},
        delta={'reference': MAX_SPEND_AT_RISK_PCT, 'increasing': {'color': '#EF4444'}, 'decreasing': {'color': '#10B981'}},
        gauge={
            'axis': {'range': [0, 25], 'ticksuffix': "%"},
            'bar': {'color': "#1A365D"},
            'steps': [
                {'range': [0, 10], 'color': "#D1FAE5"},
                {'range': [10, 15], 'color': "#FEF3C7"},
                {'range': [15, 25], 'color': "#FEE2E2"}
            ],
            'threshold': {
                'line': {'color': "#EF4444", 'width': 4},
                'thickness': 0.75,
                'value': MAX_SPEND_AT_RISK_PCT
            }
        }
    ))
    fig_gauge.update_layout(margin=dict(l=20, r=20, t=30, b=20), height=320)
    st.plotly_chart(fig_gauge, use_container_width=True)

with col_dist:
    st.subheader("📊 Purchase Price Variance (PPV %) Distribution")
    fig_hist = px.histogram(
        df_clean,
        x='price_variance_pct',
        nbins=25,
        color_discrete_sequence=['#1A365D'],
        labels={'price_variance_pct': 'Price Variance (%)'}
    )
    fig_hist.add_vline(x=0.0, line_dash="dash", line_color="#EF4444", annotation_text="0% Benchmark (Contract Rate)")
    fig_hist.update_layout(
        margin=dict(l=20, r=20, t=20, b=20),
        height=320,
        yaxis_title="Delivery Count",
        template="plotly_white"
    )
    st.plotly_chart(fig_hist, use_container_width=True)

# 3. Price Creep Audit Table
st.markdown("---")
st.subheader("📋 Price Creep Audit Log (Invoices Exceeding Contract Price)")

price_creep_df = df_clean[df_clean['price_variance_pct'] > 0][
    ['delivery_id', 'order_id', 'supplier_name', 'category_name', 'product_name',
     'contracted_unit_price', 'invoiced_unit_price', 'price_variance_pct', 'cost_variance_amount']
].sort_values(by='cost_variance_amount', ascending=False)

if len(price_creep_df) > 0:
    st.dataframe(
        price_creep_df.style.format({
            'contracted_unit_price': '₹ {:.2f}',
            'invoiced_unit_price': '₹ {:.2f}',
            'price_variance_pct': '+{:.2f}%',
            'cost_variance_amount': '₹ {:,.2f}'
        }),
        use_container_width=True,
        hide_index=True
    )
else:
    st.success("Zero unapproved price deviations detected. 100% invoice price compliance.")
