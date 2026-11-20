# BUILD PLAN: Supplier Performance Analysis & Procurement Intelligence Platform

**Project:** Supplier Performance Analysis & Procurement Intelligence System  
**Learner:** Chitra Chauhan (Roll No: 2500320140043, MCA 2nd Year DA, ABES Engineering College)  
**Mentor:** Prof. Kalpana Dwivedi  
**Target Platform:** Streamlit Community Cloud (Python 3.12, SQLite, Plotly, Tailwind UI, pytest)  
**Status:** Step 14 Implementation Blueprint (Input for Step 15 Build)

---

## 1. Project Summary & Artifact Repository

1. **Student & Academic Supervision:** Engineered by Chitra Chauhan (Roll 2500320140043), MCA 2nd Year (Data Analytics), ABES EC, guided by Project Mentor Prof. Kalpana Dwivedi.
2. **Domain & Problem Statement:** Enterprise supply chain analytics platform resolving procurement blindspots, chronic vendor delivery delays, component scrap rates, and unapproved invoice price creep.
3. **Core Technical Stack:** Python 3.12, Streamlit Multipage, Plotly Interactive Visualizations, Tailwind CSS UI components (via iframe CDN), and embedded SQLite 3 (`data/app.db`).
4. **Requirements Rigor:** 100% bidirectionally traced from 10 Operational Questions (OQ-01..10) to 12 Business Requirements (BR-01..12) and 12 Technical Features (FR-01..12).
5. **Key Performance Indicators:** Computes On-Time In-Full (OTIF%), Average Lead Time & IQR, Defective Parts Per Million (PPM), and Purchase Price Variance (PPV%).
6. **Multi-Metric Composite Scoring:** Evaluates suppliers using a weighted Composite Performance Index (CPI = 40% OTIF + 35% Quality + 25% Cost) segmenting vendors into Tier A (Preferred), Tier B (Conditional), and Tier C (At-Risk).
7. **Actionable SCM Governance:** Features automated contract renewal recommendations, dynamic KPI weight sensitivity modeling, and mandatory Corrective Action Plan (CAPA) triggers.
8. **High-Fidelity Synthetic Dataset:** Backed by 350+ realistic synthetic procurement transactions across 25 fictional Indian vendors with 5 embedded industrial patterns.
9. **Quality Assurance Rigor:** Governed by 16 Functional Acceptance Tests (TC-F-01..16) and 14 Technical Pytest Unit/Integration Tests (TC-T-01..14) with a 100% passing record.
10. **Production Deployment Ready:** Architecture engineered for zero-cost, persistent deployment on Streamlit Community Cloud via GitHub repository synchronization.

### Master Project Documentation Links:
* [01_Project_Understanding.docx](file:///c:/Users/Pc/OneDrive/Desktop/supplier%20performance%20analysis/docs/01_Project_Understanding.docx) — Scope, Personas, Stakeholders & OQs
* [02_HMW_Statements.docx](file:///c:/Users/Pc/OneDrive/Desktop/supplier%20performance%20analysis/docs/02_HMW_Statements.docx) — Design-Thinking Opportunity Framing
* [03_Needs_Analysis.docx](file:///c:/Users/Pc/OneDrive/Desktop/supplier%20performance%20analysis/docs/03_Needs_Analysis.docx) — Business Requirements (BR-01 to BR-12) & MoSCoW
* [04_Traceability_Matrix_Updated.xlsx](file:///c:/Users/Pc/OneDrive/Desktop/supplier%20performance%20analysis/docs/04_Traceability_Matrix_Updated.xlsx) — Complete End-to-End RTM
* [05_BRD.docx](file:///c:/Users/Pc/OneDrive/Desktop/supplier%20performance%20analysis/docs/05_BRD.docx) — Formal Business Requirements Document & RACI
* [07_TRD.docx](file:///c:/Users/Pc/OneDrive/Desktop/supplier%20performance%20analysis/docs/07_TRD.docx) — Technical Requirements, Architecture & Math Models
* [08_Database_Design.xlsx](file:///c:/Users/Pc/OneDrive/Desktop/supplier%20performance%20analysis/docs/08_Database_Design.xlsx) — Relational Data Dictionaries & DDL
* [09_ER_Diagram.docx](file:///c:/Users/Pc/OneDrive/Desktop/supplier%20performance%20analysis/docs/09_ER_Diagram.docx) & [09_ER_Diagram.png](file:///c:/Users/Pc/OneDrive/Desktop/supplier%20performance%20analysis/docs/09_ER_Diagram.png) — Relational ER Diagram
* [supplier_performance_data.xlsx](file:///c:/Users/Pc/OneDrive/Desktop/supplier%20performance%20analysis/data/supplier_performance_data.xlsx) & [app.db](file:///c:/Users/Pc/OneDrive/Desktop/supplier%20performance%20analysis/data/app.db) — Relational Database
* [10_Dataset_Generation_Specification.docx](file:///c:/Users/Pc/OneDrive/Desktop/supplier%20performance%20analysis/docs/10_Dataset_Generation_Specification.docx) — Synthetic Dataset Specification
* [11_User_Stories.docx](file:///c:/Users/Pc/OneDrive/Desktop/supplier%20performance%20analysis/docs/11_User_Stories.docx) — Agile User Stories & G/W/T Acceptance Criteria
* [12_UI_Framework.docx](file:///c:/Users/Pc/OneDrive/Desktop/supplier%20performance%20analysis/docs/12_UI_Framework.docx) — Wireframes, Tailwind UI, and Visualization Matrix
* [13_Test_Cases.docx](file:///c:/Users/Pc/OneDrive/Desktop/supplier%20performance%20analysis/docs/13_Test_Cases.docx) — Functional & Technical Test Specifications

---

## 2. Final Project Repository Tree

```
supplier-performance-analysis/
├── app.py                           # UI-01: Executive Overview, Macro KPIs, Global Sidebar
├── pages/
│   ├── 1_📊_Supplier_Scorecard.py   # UI-02: Multi-metric CPI, Tier A/B/C, Weight Sliders, CAPA
│   ├── 2_🚚_Delivery_Performance.py # UI-03: OTIF Leaderboard, Lead Time Boxplot, Outlier Delays
│   ├── 3_🔬_Quality_Risk.py         # UI-04: Defect PPM, Pareto 80/20 Dual-Axis, Inspection Logs
│   ├── 4_💰_Cost_Variance.py        # UI-05: Price Variance (PPV), Cost Creep, Spend-at-Risk Gauge
│   ├── 5_🔍_Data_Explorer.py        # UI-06: 42-column interactive grid, multi-filters, CSV export
│   ├── 6_🛠️_Data_Quality.py        # UI-07: DB Health, FK Integrity, Null & Outlier Diagnostics
│   └── 7_ℹ️_About.py                # UI-08: Academic credits, architecture, and project documentation
├── src/
│   ├── __init__.py
│   ├── config.py                    # Constants, SLA baselines, color palettes, default weights
│   ├── database.py                  # Thread-safe SQLite connection manager & query helpers
│   ├── data_loader.py               # Cached extraction pipeline (@st.cache_data)
│   ├── validator.py                 # Data hygiene, schema validation & cleaning routines
│   ├── analytics.py                 # Mathematical & statistical computation engines
│   └── ui.py                        # Tailwind CSS components (headers, KPI cards, insight boxes)
├── data/
│   ├── app.db                       # Relational SQLite database (5 tables + 1 view)
│   ├── supplier_performance_data.xlsx# Master Excel workbook with sheet per table
│   └── *.csv                        # Individual entity CSV tables
├── tests/
│   ├── __init__.py
│   ├── test_data.py                 # Database connectivity, row counts, PK/FK checks
│   ├── test_analysis.py             # Math validation (OTIF, PPM, Price Var, CPI, Outliers)
│   └── test_app.py                  # Headless UI component smoke test & latency SLA
├── docs/                            # Project specifications (Steps 1 to 14)
├── .streamlit/
│   └── config.toml                  # Streamlit theme (Navy #1A365D, wide layout)
├── requirements.txt                 # Pinned production dependencies
├── pytest.ini                       # Pytest execution configuration
├── BUILD_PLAN.md                    # Build implementation plan (this document)
└── README.md                        # Production setup, architecture & execution guide
```

---

## 3. Step-by-Step Build Plan Tasks (BUILD-01 to BUILD-14)

### BUILD-01: Environment & Dependency Configuration
* **Goal:** Verify Python runtime, pinned packages, and Streamlit theme configuration.
* **Files:** `requirements.txt`, `.streamlit/config.toml`, `pytest.ini`
* **Linked US / FR:** FR-11 | US-01
* **AI Execution Prompt:**
  `"Verify that requirements.txt pins streamlit, pandas, numpy, scipy, statsmodels, plotly, openpyxl, python-docx, and pytest. Ensure .streamlit/config.toml sets primaryColor='#1A365D' and wide layout."`
* **Verification Method:** Run `python -m pip check` and confirm `.streamlit/config.toml` exists.
* **Tests to Pass:** Environment import sanity check.

### BUILD-02: Data Ingestion & Defensive Cleaning Engine
* **Goal:** Build robust SQLite loader with `@st.cache_data` and cleaning routines for missing/duplicate data.
* **Files:** `src/database.py`, `src/data_loader.py`, `src/validator.py`
* **Linked US / FR:** FR-01, FR-11 | US-01, US-02 | BR-03, BR-11
* **AI Execution Prompt:**
  `"Implement src/database.py to connect to data/app.db with foreign keys enabled. In src/data_loader.py, implement load_analytical_data() wrapped in @st.cache_data. In src/validator.py, implement clean_data() to handle null in-transit delivery dates, deduplicate clone records, and enforce type constraints."`
* **Verification Method:** Inspect returned DataFrame: shape should be >= 300 rows and 42 columns with 0 unhandled nulls.
* **Tests to Pass:** `tests/test_data.py::test_tc_t_01_database_connection`, `tests/test_data.py::test_tc_t_05_analytical_view_columns`.

### BUILD-03: Statistical & Analytics Calculation Engine
* **Goal:** Implement core mathematical and statistical computation functions for all KPIs.
* **Files:** `src/analytics.py`, `src/config.py`
* **Linked US / FR:** FR-01..10 | US-05, US-06, US-07, US-08, US-09, US-10, US-11, US-12, US-13, US-14, US-15
* **AI Execution Prompt:**
  `"In src/analytics.py, implement compute_kpi_summary(), compute_otif_leaderboard(), compute_lead_time_stats(), compute_defect_pareto(), compute_price_variance_audit(), compute_cpi_scores(weights), and compute_spend_at_risk(). Ensure zero-division guards and robust NumPy/Pandas vectorization."`
* **Verification Method:** Execute functions on `data/app.db` and verify outputs match known mathematical benchmarks.
* **Tests to Pass:** All 7 unit tests in `tests/test_analysis.py` (TC-T-06 through TC-T-12).

### BUILD-04: Reusable Tailwind CSS Component Layer
* **Goal:** Finalize reusable UI components (header banner, KPI metric cards, alert boxes) using Tailwind CDN and `components.html`.
* **Files:** `src/ui.py`
* **Linked US / FR:** FR-11 | US-03
* **AI Execution Prompt:**
  `"In src/ui.py, verify render_tailwind_header(), render_kpi_cards(), and render_insight_box(). Ensure responsive grid styling, status badge color mapping, and clean HTML iframe sizing."`
* **Verification Method:** Run headless smoke test to verify HTML string formatting.
* **Tests to Pass:** `tests/test_app.py::test_tc_t_13_ui_components_render`.

### BUILD-05: Executive Overview Screen (UI-01)
* **Goal:** Build main landing application page (`app.py`) featuring top KPI cards, global sidebar filters, quarterly trendline, and spend donut.
* **Files:** `app.py`
* **Linked US / FR:** FR-01, FR-02, FR-06, FR-11 | US-03, US-04, US-12 | OQ-01, OQ-02, OQ-06, OQ-08
* **AI Execution Prompt:**
  `"In app.py, initialize Streamlit page config (title='Supplier Performance Intelligence', layout='wide'). Implement shared sidebar filters in st.session_state (Category, Supplier, Date Range, Status). Render Tailwind header, 4 executive KPI cards, Plotly quarterly OTIF line chart, and category spend donut chart."`
* **Verification Method:** Launch `streamlit run app.py --server.headless=true` and verify response status 200.
* **Tests to Pass:** TC-F-04, TC-F-05.

### BUILD-06: Supplier Scorecard & CPI Tiering Page (UI-02)
* **Goal:** Build multi-dimensional scorecard screen featuring interactive weight sliders, radar charts, and CAPA governance flags.
* **Files:** `pages/1_📊_Supplier_Scorecard.py`
* **Linked US / FR:** FR-06, FR-09, FR-12 | US-13, US-14, US-15 | OQ-06, OQ-09
* **AI Execution Prompt:**
  `"Build pages/1_📊_Supplier_Scorecard.py. Include sidebar sliders for Delivery, Quality, and Cost weights (normalizing to 100%). Render tier distribution badges (A/B/C), an interactive Plotly radar chart, the master scorecard table, and automated CAPA triggers for vendors with CPI < 70."`
* **Verification Method:** Move weight sliders and verify scorecard table recalculates instantly.
* **Tests to Pass:** TC-F-14, TC-F-15.

### BUILD-07: Delivery & Lead Time Analysis Page (UI-03)
* **Goal:** Implement delivery adherence screen featuring OTIF leaderboard, category lead time boxplots, delay tail outliers, and freight modes.
* **Files:** `pages/2_🚚_Delivery_Performance.py`
* **Linked US / FR:** FR-01, FR-02, FR-05, FR-10 | US-05, US-06, US-07, US-08 | OQ-01, OQ-02, OQ-05, OQ-10
* **AI Execution Prompt:**
  `"Build pages/2_🚚_Delivery_Performance.py. Render horizontal OTIF leaderboard with 95% target line, Plotly category lead time boxplots against SLA markers, delay tail outlier table (>7 days), and freight shipping mode comparisons."`
* **Verification Method:** Inspect charts for correct axes, hover data, and outlier points.
* **Tests to Pass:** TC-F-06, TC-F-07, TC-F-08, TC-F-09.

### BUILD-08: Quality Inspection & Pareto Risk Page (UI-04)
* **Goal:** Build receiving quality screen with Defect PPM tracking, dual-axis Pareto 80/20 chart, and defect quarantine logs.
* **Files:** `pages/3_🔬_Quality_Risk.py`
* **Linked US / FR:** FR-03 | US-09, US-10 | OQ-03
* **AI Execution Prompt:**
  `"Build pages/3_🔬_Quality_Risk.py. Render Defect PPM metric cards, a dual-axis Pareto chart (defect frequency bars + cumulative 80% line), a vendor PPM scatter plot, and a quality audit table highlighting vendors exceeding 1,000 PPM."`
* **Verification Method:** Verify Pareto bars sort descending and cumulative line terminates at 100%.
* **Tests to Pass:** TC-F-10, TC-F-11.

### BUILD-09: Cost Variance & Spend-at-Risk Page (UI-05)
* **Goal:** Implement commercial audit screen tracking purchase price variance (PPV) and Spend-at-Risk allocation.
* **Files:** `pages/4_💰_Cost_Variance.py`
* **Linked US / FR:** FR-04, FR-07 | US-11, US-12 | OQ-04, OQ-07
* **AI Execution Prompt:**
  `"Build pages/4_💰_Cost_Variance.py. Render Spend-at-Risk gauge (highlighting 10% threshold), price variance distribution histogram, waterfall chart of category cost variances, and an invoice dispute audit table for lines with price creep > 0%."`
* **Verification Method:** Verify price variance percentage delta formula and gauge boundary.
* **Tests to Pass:** TC-F-12, TC-F-13.

### BUILD-10: Interactive Data Explorer & Export Page (UI-06)
* **Goal:** Build dynamic raw data grid with multi-column sorting and CSV/Excel download buttons.
* **Files:** `pages/5_🔍_Data_Explorer.py`, `src/utils.py`
* **Linked US / FR:** FR-11, FR-12 | US-04, US-16
* **AI Execution Prompt:**
  `"Build pages/5_🔍_Data_Explorer.py. Render filtered analytical dataframe using st.dataframe with column configurations. Implement st.download_button for CSV and Excel downloads with timestamped file names."`
* **Verification Method:** Click CSV download button and inspect downloaded file integrity.
* **Tests to Pass:** TC-F-16.

### BUILD-11: Data Quality & Diagnostics Page (UI-07)
* **Goal:** Build system diagnostics page displaying database health, FK integrity, and data anomaly counts.
* **Files:** `pages/6_🛠️_Data_Quality.py`
* **Linked US / FR:** FR-11 | US-01, US-02
* **AI Execution Prompt:**
  `"Build pages/6_🛠️_Data_Quality.py. Display SQLite table row counts, PRAGMA foreign_key_check results, missing data summary, duplicate record counts, and outlier summary stats."`
* **Verification Method:** Verify all diagnostic checks report PASS.
* **Tests to Pass:** TC-T-01..05.

### BUILD-12: About & Project Architecture Page (UI-08)
* **Goal:** Implement project background, methodology, student details, mentor acknowledgments, and architecture reference.
* **Files:** `pages/7_ℹ️_About.py`
* **Linked US / FR:** Documentation | All Personas
* **AI Execution Prompt:**
  `"Build pages/7_ℹ️_About.py. Render ABES MCA Capstone metadata (Chitra Chauhan, Roll 2500320140043, Mentor Prof. Kalpana Dwivedi), problem statement summary, Mermaid architecture diagram, and methodology reference."`
* **Verification Method:** Navigate to page and confirm all markdown elements render cleanly.
* **Tests to Pass:** UI smoke test.

### BUILD-13: Comprehensive Test Execution & Latency Profiling
* **Goal:** Execute full automated test suite and benchmark application performance against SLAs.
* **Files:** `tests/test_*.py`
* **AI Execution Prompt:**
  `"Run python -m pytest -v --cov=src. Confirm 100% test pass rate (14/14 tests) and verify query execution latency remains under 2.5 seconds."`
* **Verification Method:** Pytest report with 0 failures.
* **Tests to Pass:** All TC-T tests (14/14).

### BUILD-14: Production Packaging & Deployment Preparation
* **Goal:** Finalize `README.md`, git tracking, and Streamlit Community Cloud launch settings.
* **Files:** `README.md`, `.gitignore`
* **AI Execution Prompt:**
  `"Generate a comprehensive, GitHub-ready README.md detailing project overview, architecture diagram, local setup instructions, test commands, and cloud deployment steps."`
* **Verification Method:** Inspect README formatting and verify git status.
* **Tests to Pass:** Clean git repository state.

---

## 4. Definition of Done (DoD)

The project reaches formal completion when:
1. **100% Functional Coverage:** All 8 screens are fully operational and all 12 Must-Have Business Requirements are satisfied.
2. **100% Automated Test Pass Rate:** All 14 technical pytest unit and integration tests pass with zero errors.
3. **Traceability Closure:** Every item in the Requirements Traceability Matrix has progressed from "Identified" to "Verified".
4. **Data Integrity:** `data/app.db` executes with 0 foreign key violations, 0 negative values, and deterministic reproducibility.
5. **Deployment:** The application deploys to Streamlit Community Cloud with public accessibility and 0 runtime crashes.
6. **Academic Documentation:** Final Project Report (ABES template) and Defense Slide Deck are completely compiled.

---

## 5. Development & Execution Commands

```bash
# 1. Environment Setup
python -m pip install -r requirements.txt

# 2. Synthetic Dataset Generation (if recreating database)
python src/generate_data.py

# 3. Execute Automated Pytest Suite
python -m pytest -v

# 4. Launch Local Streamlit Application
streamlit run app.py
```

---

## 6. Project Defense Viva Notes (10 Likely Questions & Answers)

1. **Q: What core operational problem does this platform solve?**  
   *A:* It centralizes fragmented procurement logs, automating the empirical measurement of vendor delivery delays, short-shipments, component quality spikes, and unapproved price creep to eliminate production stoppages.
2. **Q: What is the difference between On-Time Delivery and On-Time In-Full (OTIF)?**  
   *A:* On-Time Delivery only measures arrival date. OTIF strictly enforces that the order arrived on or before the committed SLA date **and** delivered 100% of the ordered quantity, preventing vendors from gaming metrics with partial token shipments.
3. **Q: How is the Composite Performance Index (CPI) calculated and what do the tiers mean?**  
   *A:* $	ext{CPI} = 0.40 	imes 	ext{OTIF\%} + 0.35 	imes 	ext{Quality Score} + 0.25 	imes 	ext{Cost Score}$ on a 0–100 scale. Tier A ($\ge 85$) is Preferred/Renewal; Tier B ($70	ext{--}84.9$) is Conditional/Monitoring; Tier C ($< 70$) triggers mandatory CAPA audits.
4. **Q: Why was SQLite selected over cloud databases like PostgreSQL or Snowflake?**  
   *A:* SQLite is a serverless, embedded relational engine with zero maintenance overhead that lives directly in the git repository. It runs with sub-millisecond read latency within Streamlit Cloud's 1 GB container with \$0 infrastructure cost.
5. **Q: Why is `@st.cache_data` used in data extraction?**  
   *A:* Streamlit re-runs the Python script on every user interaction. `@st.cache_data` caches serialized DataFrames in memory, eliminating repetitive SQLite disk I/O and reducing query latency from 250ms to $< 5	ext{ms}$.
6. **Q: How are Tailwind CSS styles integrated safely into Streamlit?**  
   *A:* Custom HTML cards are rendered inside isolated, sandboxed iframes using `st.components.v1.html()` pulling the official Tailwind CDN. This prevents style collisions with Streamlit's native React DOM.
7. **Q: How does the application detect heavy-tail delivery delay risks?**  
   *A:* It uses Tukey's fences ($Q_3 + 1.5 	imes 	ext{IQR}$) and Fisher-Pearson skewness/kurtosis to isolate extreme delivery delay outliers ($> 7$ days past SLA), exposing hidden logistics bottlenecks.
8. **Q: Why does `order_deliveries` serve as an associative bridge table?**  
   *A:* Purchase orders and product SKUs share a Many-to-Many ($M:N$) relationship. `order_deliveries` resolves this into two normalized $1:N$ relationships while tracking fulfillment dates, invoice variances, and QA inspection outcomes at the line grain.
9. **Q: What is Spend-at-Risk and why is its benchmark capped at 10%?**  
   *A:* Spend-at-Risk is the ratio of procurement spend tied to Tier C (underperforming) suppliers over total spend. Capping it below 10% prevents catastrophic supply chain vulnerability.
10. **Q: What role does the Requirements Traceability Matrix (RTM) serve in this project?**  
    *A:* The RTM proves end-to-end engineering integrity by linking every business question (OQ) down to a business requirement (BR), technical feature (FR), database column, user story (US), UI screen (UI), and automated test case (TC).
