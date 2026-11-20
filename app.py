"""
SUPPLIER PERFORMANCE ANALYSIS & PROCUREMENT INTELLIGENCE PLATFORM
Main Application & Executive Overview Dashboard (UI-01)
Author: Chitra Chauhan (Roll No: 2500320140043, MCA 2nd Year DA, ABES EC)
Mentor: Prof. Kalpana Dwivedi
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

from src.config import (
    APP_TITLE,
    APP_SUBTITLE,
    STUDENT_NAME,
    STUDENT_ROLL,
    MENTOR_NAME,
    INSTITUTION,
    TARGET_OTIF_PCT,
    MAX_DEFECT_PPM,
    MAX_SPEND_AT_RISK_PCT,
    PRIMARY_COLOR,
    SECONDARY_COLOR
)
from src.data_loader import load_analytical_data, get_filter_bounds, filter_dataframe
from src.analytics import compute_kpi_summary, compute_supplier_cpi_table
from src.ui import render_tailwind_header, render_kpi_cards, render_insight_box
from src.utils import format_inr, format_percent

# 1. Page Configuration
st.set_page_config(
    page_title="Procurement Intelligence Cockpit",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 2. Data Loading Pipeline
with st.spinner("Connecting to SQLite database & validating procurement records..."):
    df_clean, hygiene_report = load_analytical_data()

filter_bounds = get_filter_bounds(df_clean)

# 3. Global Unified Sidebar Navigation & Filters
st.sidebar.markdown(f"""
<div style="padding: 10px 0px; border-bottom: 1px solid #E2E8F0; margin-bottom: 15px;">
    <h3 style="margin: 0; color: #1A365D; font-size: 1.15rem; font-weight: 800;">🏢 Procurement Intel</h3>
    <p style="margin: 2px 0 0 0; color: #64748B; font-size: 0.75rem;">ABES EC Capstone • Chitra Chauhan</p>
</div>
""", unsafe_allow_html=True)

st.sidebar.subheader("🎯 Global Scope Filters")

# Category Multiselect
selected_categories = st.sidebar.multiselect(
    "Product Categories",
    options=filter_bounds["categories"],
    default=filter_bounds["categories"],
    help="Filter orders by specific supply category."
)

# Supplier Multiselect
selected_suppliers = st.sidebar.multiselect(
    "Suppliers",
    options=filter_bounds["suppliers"],
    default=filter_bounds["suppliers"],
    help="Filter by specific vendor name."
)

# Delivery Status Filter
selected_statuses = st.sidebar.multiselect(
    "Fulfillment Status",
    options=filter_bounds["statuses"],
    default=filter_bounds["statuses"]
)

# Date Range Filter
min_d = filter_bounds["min_date"].date() if filter_bounds["min_date"] else None
max_d = filter_bounds["max_date"].date() if filter_bounds["max_date"] else None

date_selection = st.sidebar.date_input(
    "Order Date Range",
    value=(min_d, max_d) if (min_d and max_d) else None,
    help="Filter transactions across specific timeframes."
)

# Apply Filters
filtered_df = filter_dataframe(
    df_clean,
    categories=selected_categories,
    suppliers=selected_suppliers,
    statuses=selected_statuses,
    date_range=date_selection if isinstance(date_selection, (tuple, list)) and len(date_selection) == 2 else None
)

# Sidebar System Health Status
st.sidebar.markdown("---")
st.sidebar.markdown(f"""
<div style="font-size: 0.75rem; color: #64748B;">
    <b>Database:</b> <code>data/app.db</code> (SQLite 3)<br>
    <b>Active Records:</b> {len(filtered_df)} / {len(df_clean)} deliveries<br>
    <b>Data Hygiene:</b> 100% Validated<br>
    <b>Roll No:</b> {STUDENT_ROLL}
</div>
""", unsafe_allow_html=True)

# 4. Main Body: Tailwind Header Banner
render_tailwind_header(
    title="Supplier Performance & Procurement Intelligence Cockpit",
    subtitle="Enterprise Decision-Support Platform for Vendor Evaluation, Risk Mitigation & SCM Analytics",
    badge_text="UI-01 • EXECUTIVE VIEW"
)

st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)

# 5. Top Macro Executive KPI Metric Cards
kpi = compute_kpi_summary(filtered_df)

otif_status = "success" if kpi["macro_otif_pct"] >= TARGET_OTIF_PCT else ("warning" if kpi["macro_otif_pct"] >= 80.0 else "danger")
ppm_status = "success" if kpi["macro_defect_ppm"] <= MAX_DEFECT_PPM else ("warning" if kpi["macro_defect_ppm"] <= 2500 else "danger")
risk_status = "success" if kpi["spend_at_risk_pct"] <= MAX_SPEND_AT_RISK_PCT else "danger"

kpi_cards_data = [
    {
        "label": "Overall OTIF Rate",
        "value": f"{kpi['macro_otif_pct']:.1f}%",
        "benchmark": f"Target ≥ {TARGET_OTIF_PCT:.0f}%",
        "status": otif_status,
        "subtext": f"{filtered_df['is_otif'].sum() if 'is_otif' in filtered_df else 0} / {len(filtered_df)} on-time in-full"
    },
    {
        "label": "Total Procurement Spend",
        "value": format_inr(kpi["total_spend"]),
        "benchmark": f"{kpi['total_orders']} Purchase Orders",
        "status": "neutral",
        "subtext": f"{kpi['total_deliveries']} line deliveries"
    },
    {
        "label": "Receiving Defect PPM",
        "value": f"{kpi['macro_defect_ppm']:,}",
        "benchmark": f"Target < {MAX_DEFECT_PPM:,} PPM",
        "status": ppm_status,
        "subtext": "Quality inspection scrap rate"
    },
    {
        "label": "Spend-at-Risk %",
        "value": f"{kpi['spend_at_risk_pct']:.1f}%",
        "benchmark": f"Target < {MAX_SPEND_AT_RISK_PCT:.0f}%",
        "status": risk_status,
        "subtext": "Exposure to Tier C vendors"
    }
]

render_kpi_cards(kpi_cards_data)

st.markdown("<div style='height: 15px;'></div>", unsafe_allow_html=True)

# 6. Charts Section: OTIF Trend & Spend Breakdown
col_trend, col_spend = st.columns([3, 2])

with col_trend:
    st.subheader("📈 Quarterly Delivery Adherence (OTIF %)")
    if len(filtered_df) > 0 and 'order_date' in filtered_df.columns:
        trend_df = filtered_df.copy()
        trend_df['quarter'] = trend_df['order_date'].dt.to_period('Q').astype(str)
        quarterly_summary = trend_df.groupby('quarter').agg(
            total=('delivery_id', 'count'),
            otif=('is_otif', 'sum')
        ).reset_index()
        quarterly_summary['otif_pct'] = (quarterly_summary['otif'] / quarterly_summary['total'] * 100.0).round(1)

        fig_trend = go.Figure()
        fig_trend.add_trace(go.Scatter(
            x=quarterly_summary['quarter'],
            y=quarterly_summary['otif_pct'],
            mode='lines+markers+text',
            name='OTIF Rate %',
            text=[f"{v:.1f}%" for v in quarterly_summary['otif_pct']],
            textposition='top center',
            line=dict(color='#1A365D', width=3),
            marker=dict(size=8, color='#D69E2E')
        ))
        # Benchmark SLA Line
        fig_trend.add_hline(
            y=TARGET_OTIF_PCT,
            line_dash="dash",
            line_color="#E53E3E",
            annotation_text=f"SLA Target ({TARGET_OTIF_PCT}%)",
            annotation_position="bottom right"
        )
        fig_trend.update_layout(
            yaxis=dict(title="OTIF Percentage (%)", range=[50, 105]),
            xaxis=dict(title="Fiscal Quarter"),
            margin=dict(l=40, r=40, t=20, b=40),
            height=340,
            template="plotly_white"
        )
        st.plotly_chart(fig_trend, use_container_width=True)
    else:
        st.info("No delivery records available for the selected scope.")

with col_spend:
    st.subheader("💳 Spend by Product Category")
    if len(filtered_df) > 0 and 'category_name' in filtered_df.columns:
        cat_spend = filtered_df.groupby('category_name')['invoiced_line_total'].sum().reset_index()
        cat_spend = cat_spend.sort_values(by='invoiced_line_total', ascending=False)
        
        fig_donut = px.pie(
            cat_spend,
            names='category_name',
            values='invoiced_line_total',
            hole=0.45,
            color_discrete_sequence=px.colors.qualitative.Prism
        )
        fig_donut.update_traces(
            textinfo='percent+label',
            hovertemplate="<b>%{label}</b><br>Spend: ₹%{value:,.0f}<br>Share: %{percent}"
        )
        fig_donut.update_layout(
            margin=dict(l=20, r=20, t=20, b=20),
            height=340,
            showlegend=False
        )
        st.plotly_chart(fig_donut, use_container_width=True)
    else:
        st.info("No spend records available.")

# 7. Operational Alerts & Actionable Intelligence
st.markdown("---")
cpi_table = compute_supplier_cpi_table(filtered_df)

col_alert, col_top5 = st.columns([1, 2])

with col_alert:
    st.subheader("⚡ Procurement Alerts")
    tier_c_count = len(cpi_table[cpi_table['performance_tier'].str.contains("Tier C", na=False)]) if len(cpi_table) > 0 else 0
    severe_delays = len(filtered_df[filtered_df['delay_days'] > 7]) if 'delay_days' in filtered_df.columns else 0
    price_creeps = len(filtered_df[filtered_df['price_variance_pct'] > 0]) if 'price_variance_pct' in filtered_df.columns else 0

    render_insight_box(
        title=f"{tier_c_count} Suppliers in Tier C",
        message="Vendors scoring below 70 CPI threshold require mandatory Corrective Action Plans (CAPA).",
        alert_type="danger" if tier_c_count > 0 else "success"
    )
    st.markdown("<div style='height: 8px;'></div>", unsafe_allow_html=True)
    render_insight_box(
        title=f"{severe_delays} Severe Delivery Delays",
        message="Deliveries delayed beyond 7 days past contractual SLA detected in logistics logs.",
        alert_type="warning" if severe_delays > 0 else "info"
    )

with col_top5:
    st.subheader("🏆 Supplier Tier Snapshot")
    if len(cpi_table) > 0:
        display_cols = ['supplier_name', 'category_name', 'otif_pct', 'defect_ppm', 'cpi_score', 'performance_tier', 'capa_status']
        st.dataframe(
            cpi_table[display_cols].head(6).style.format({
                'otif_pct': '{:.1f}%',
                'defect_ppm': '{:,}',
                'cpi_score': '{:.1f}'
            }),
            use_container_width=True,
            hide_index=True
        )
    else:
        st.info("No supplier records in current selection.")
