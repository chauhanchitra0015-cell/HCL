"""
SUPPLIER PERFORMANCE ANALYSIS — ABOUT & ACADEMIC ARCHITECTURE (UI-08)
Author: Chitra Chauhan (Roll No: 2500320140043, MCA 2nd Year DA, ABES EC)
Mentor: Prof. Kalpana Dwivedi
"""

import streamlit as st
from src.ui import render_tailwind_header
from src.config import (
    STUDENT_NAME,
    STUDENT_ROLL,
    STUDENT_PROGRAM,
    MENTOR_NAME,
    INSTITUTION,
    APP_VERSION
)

st.set_page_config(page_title="About & Architecture", page_icon="ℹ️", layout="wide")

render_tailwind_header(
    title="Project Architecture & Academic Reference",
    subtitle="System specifications, mathematical formulation, and academic capstone attribution",
    badge_text=f"UI-08 • {APP_VERSION}"
)

st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)

col_stu, col_arc = st.columns([1, 1])

with col_stu:
    st.subheader("🎓 Student & Capstone Attribution")
    st.markdown(f"""
    * **Student Name:** {STUDENT_NAME}
    * **University Roll No:** `{STUDENT_ROLL}`
    * **Program:** {STUDENT_PROGRAM}
    * **Project Mentor:** {MENTOR_NAME}
    * **Department:** Department of Computer Applications
    * **Institution:** {INSTITUTION}
    * **Course:** MCA Major Capstone Project (2024–2026 Batch)
    * **Specialization:** Enterprise Data Analytics & Decision Support Systems
    """)

    st.markdown("---")
    st.subheader("📐 Mathematical Modeling & KPI Equations")
    st.markdown("""
    1. **On-Time In-Full Rate (OTIF %):**
       $$\\text{OTIF} = \\frac{\\sum \\text{is\\_otif}}{\\text{Total Deliveries}} \\times 100$$
       *Enforces that delivery arrived on/before SLA AND delivered 100% ordered quantity.*

    2. **Receiving Defect Parts Per Million (PPM):**
       $$\\text{PPM} = \\left( \\frac{\\text{Defective Qty}}{\\text{Inspected Qty}} \\right) \\times 1,000,000$$

    3. **Purchase Price Variance % (PPV):**
       $$\\text{PPV \\%} = \\left( \\frac{\\text{Invoiced Price} - \\text{Contracted Price}}{\\text{Contracted Price}} \\right) \\times 100$$

    4. **Composite Performance Index (CPI):**
       $$\\text{CPI} = 0.40 \\times \\text{OTIF\\%} + 0.35 \\times \\text{Quality Score} + 0.25 \\times \\text{Cost Score}$$

    5. **Governance Tiers:**
       * **Tier A (Preferred):** $\\text{CPI} \\ge 85.0$
       * **Tier B (Conditional):** $70.0 \\le \\text{CPI} < 85.0$
       * **Tier C (At-Risk):** $\\text{CPI} < 70.0$ *(Mandatory CAPA Audit)*
    """)

with col_arc:
    st.subheader("🏛️ Enterprise Multi-Tier Architecture")
    st.markdown("""
    ```mermaid
    flowchart TD
        subgraph Data_Layer ["Data Tier (SQLite 3)"]
            DB[(data/app.db)]
            V[View: v_supplier_performance_analysis]
            DB --> V
        end

        subgraph Ingestion_Layer ["Ingestion & Validation Engine"]
            LOAD[src/data_loader.py]
            VAL[src/validator.py]
            V --> LOAD
            LOAD --> VAL
        end

        subgraph Analytical_Engine ["Analytics & Statistical Models"]
            KPI[src/analytics.py]
            STAT[Lead Time IQR & Tukey Fences]
            PAR[Pareto 80/20 Defect Analysis]
            CPI_MOD[Weighted CPI & Tier Assignment]
            VAL --> KPI
            KPI --> STAT
            KPI --> PAR
            KPI --> CPI_MOD
        end

        subgraph Presentation_Layer ["Streamlit UI & Tailwind CDN"]
            APP[app.py - Executive Cockpit]
            P1[pages/1 - Scorecard & CPI]
            P2[pages/2 - Delivery Performance]
            P3[pages/3 - Quality Risk]
            P4[pages/4 - Cost Variance]
            P5[pages/5 - Data Explorer]
            P6[pages/6 - Data Quality]
            P7[pages/7 - About & Docs]
            CPI_MOD --> APP
            CPI_MOD --> P1
            CPI_MOD --> P2
            CPI_MOD --> P3
            CPI_MOD --> P4
            CPI_MOD --> P5
            CPI_MOD --> P6
            CPI_MOD --> P7
        end
    ```
    """)
    st.info("Architecture engineered for sub-second query latency and zero-cost persistent deployment on Streamlit Community Cloud.")
