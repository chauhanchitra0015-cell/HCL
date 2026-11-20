import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

os.makedirs('docs', exist_ok=True)
doc = docx.Document()

# Page Margins
for s in doc.sections:
    s.top_margin = Inches(1)
    s.bottom_margin = Inches(1)
    s.left_margin = Inches(1)
    s.right_margin = Inches(1)

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>'))

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

# Title & Subtitle
p_title = doc.add_paragraph()
r_title = p_title.add_run('TECHNICAL REQUIREMENTS DOCUMENT (TRD)')
r_title.bold = True
r_title.font.size = Pt(22)
r_title.font.color.rgb = RGBColor(24, 43, 73)
p_title.paragraph_format.space_after = Pt(2)

p_sub = doc.add_paragraph()
r_sub = p_sub.add_run('System Architecture, Engineering Specifications & Statistical Methodology')
r_sub.font.size = Pt(13)
r_sub.font.color.rgb = RGBColor(90, 105, 120)
p_sub.paragraph_format.space_after = Pt(14)

# Document Control Table
meta_table = doc.add_table(rows=6, cols=2)
meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
meta_data = [
    ('Project Title', 'Supplier Performance Analysis & Procurement Intelligence System'),
    ('Document Version', 'Version 1.0 (Formal Technical Baseline)'),
    ('Lead Author & Role', 'Chitra Chauhan | Roll No: 2500320140043 | Solution Architect & Developer'),
    ('Academic Mentor', 'Prof. Kalpana Dwivedi (Dept. of Computer Applications, ABES EC)'),
    ('Target Architecture', 'Python 3.12, Streamlit Multipage, Plotly, Tailwind UI, SQLite, pytest'),
    ('Deployment Target', 'Streamlit Community Cloud (Linux Container, Free Tier, Zero-Cost)')
]
for idx, (label, val) in enumerate(meta_data):
    r = meta_table.rows[idx]
    r.cells[0].text = label
    r.cells[1].text = val
    set_cell_background(r.cells[0], 'F0F4F8')
    r.cells[0].paragraphs[0].runs[0].bold = True
    r.cells[0].paragraphs[0].runs[0].font.size = Pt(9.5)
    r.cells[1].paragraphs[0].runs[0].font.size = Pt(9.5)
    set_cell_margins(r.cells[0], 70, 70, 100, 100)
    set_cell_margins(r.cells[1], 70, 70, 100, 100)

doc.add_paragraph().paragraph_format.space_after = Pt(8)

# Section 1: System Architecture
h1 = doc.add_heading('1. System Architecture & End-to-End Data Pipeline', level=1)
h1.paragraph_format.space_before = Pt(12)
h1.paragraph_format.space_after = Pt(6)

p_arch = doc.add_paragraph(
    'The system implements a decoupled, modular data analytics pipeline designed for rapid interactive querying, robust caching, and rich multi-page visual presentation. The architecture follows a linear, non-cyclic data flow: Data Persistence → Ingestion & Caching → Schema Validation → Analytical & Statistical Processing → Presentation Layer.'
)
p_arch.paragraph_format.line_spacing = 1.15
p_arch.paragraph_format.space_after = Pt(6)

p_mermaid = doc.add_paragraph()
r_m_title = p_mermaid.add_run('Mermaid Architecture Flowchart Specification:\n')
r_m_title.bold = True
r_m_code = p_mermaid.add_run(
    'graph TD\n'
    '    A[SQLite Database: data/app.db] -->|Raw SQL Query via sqlite3| B[Data Loader: src/data_loader.py]\n'
    '    B -->|st.cache_data In-Memory Cache| C[Data Validator: src/validator.py]\n'
    '    C -->|Validated Pandas DataFrames| D[Analytics Engine: src/analytics.py]\n'
    '    D -->|OTIF, Lead Time, Defect PPM, Cost Var, CPI| E[Streamlit Multi-Page UI]\n'
    '    E --> F[app.py: Executive Overview & Global Filters]\n'
    '    E --> G[pages/1_Scorecard.py: Vendor CPI & Tiering]\n'
    '    E --> H[pages/2_Delivery.py: Lead Time & OTIF]\n'
    '    E --> I[pages/3_Quality.py: Defect PPM & Pareto]\n'
    '    E --> J[pages/4_Cost.py: Price Variance & Spend-at-Risk]\n'
    '    E --> K[pages/5_Explorer.py: Data Grid & Export]\n'
    '    L[Tailwind CDN Components: src/components.py] -->|HTML iframe Ingestion| E'
)
r_m_code.font.name = 'Courier New'
r_m_code.font.size = Pt(8.5)
p_mermaid.paragraph_format.space_after = Pt(8)

# Section 2: Folder Structure
h2 = doc.add_heading('2. Modular Project Repository Structure', level=1)
h2.paragraph_format.space_before = Pt(12)
h2.paragraph_format.space_after = Pt(6)

p_tree = doc.add_paragraph()
r_tree = p_tree.add_run(
    'supplier-performance-analysis/\n'
    '├── app.py                      # Main entrypoint, executive Tailwind KPI cards, global filters\n'
    '├── pages/                      # Multi-page Streamlit application modules\n'
    '│   ├── 1_📊_Supplier_Scorecard.py   # Multi-metric CPI scoring and Tier A/B/C classification\n'
    '│   ├── 2_🚚_Delivery_Performance.py # OTIF compliance, lead time distributions, delay outliers\n'
    '│   ├── 3_🔬_Quality_Risk_Analysis.py# Defect PPM, Pareto defect reasons, QA inspection audits\n'
    '│   ├── 4_💰_Cost_Variance_Audit.py  # Purchase Price Variance (PPV) and Spend-at-Risk analytics\n'
    '│   └── 5_🔍_Data_Explorer.py        # Raw data slicing, dynamic filtering, and CSV/Excel exports\n'
    '├── src/                        # Core application business logic and engines\n'
    '│   ├── __init__.py\n'
    '│   ├── config.py               # Constants, theme colors, default weights, and SLA thresholds\n'
    '│   ├── database.py             # SQLite connection pooling, query execution, and transactions\n'
    '│   ├── data_loader.py          # Cached data extraction pipelines using @st.cache_data\n'
    '│   ├── validator.py            # Data type enforcement, range boundary and schema validation\n'
    '│   ├── analytics.py            # Mathematical calculation engines (OTIF, CPI, statistics)\n'
    '│   ├── components.py           # Tailwind CSS styled KPI cards and HTML iframe wrappers\n'
    '│   └── utils.py                # Formatting helpers, logging, and tabular export utilities\n'
    '├── data/                       # Local relational storage and DDL\n'
    '│   ├── app.db                  # Persistent SQLite relational database\n'
    '│   └── schema.sql              # Relational DDL schema script\n'
    '├── tests/                      # Automated test suite (pytest)\n'
    '│   ├── __init__.py\n'
    '│   ├── test_database.py        # Schema integrity, foreign key, and connection tests\n'
    '│   ├── test_analytics.py       # Math validation (OTIF, Lead Time, Defect PPM, Cost, CPI)\n'
    '│   ├── test_validator.py       # Data validation boundary and edge-case tests\n'
    '│   └── test_app_ui.py          # Streamlit AppTest headless UI integration smoke tests\n'
    '├── docs/                       # Project artifacts (BRD, TRD, RTM, Presentation Deck)\n'
    '├── .streamlit/\n'
    '│   └── config.toml             # Custom theme palette, layout width, and server settings\n'
    '├── requirements.txt            # Strict, pinned production dependencies\n'
    '├── pytest.ini                  # Pytest test discovery and execution configuration\n'
    '└── README.md                   # Installation, execution, and deployment manual'
)
r_tree.font.name = 'Courier New'
r_tree.font.size = Pt(8.5)
p_tree.paragraph_format.space_after = Pt(8)

# Section 3: Functional Requirements (FR-01 to FR-12)
h3 = doc.add_heading('3. Functional Requirements Catalog (FR-01 to FR-12)', level=1)
h3.paragraph_format.space_before = Pt(12)
h3.paragraph_format.space_after = Pt(6)

fr_table = doc.add_table(rows=13, cols=5)
fr_table.alignment = WD_TABLE_ALIGNMENT.CENTER
fr_headers = ['FR ID', 'Technical Functional Requirement', 'Linked BR', 'Target Python Module', 'Verification Criteria']
for j, h in enumerate(fr_headers):
    c = fr_table.cell(0, j)
    c.text = h
    set_cell_background(c, '1A365D')
    p = c.paragraphs[0]
    p.runs[0].bold = True
    p.runs[0].font.color.rgb = RGBColor(255, 255, 255)
    p.runs[0].font.size = Pt(8.5)
    set_cell_margins(c, 80, 80, 80, 80)

fr_data = [
    ('FR-01', 'Calculate On-Time In-Full (OTIF) fulfillment rate per vendor and order line item using dual conditional boolean evaluation.', 'BR-01', 'src/analytics.py', 'OTIF = 1 iff actual_date <= promised_date AND delivered_qty >= ordered_qty; returns float percentage.'),
    ('FR-02', 'Compute parametric (Mean, Std Dev) and non-parametric (Median, IQR) statistics for delivery lead times against SLA benchmarks.', 'BR-02', 'src/analytics.py', 'Elapsed calendar days calculation; returns lead time distribution metrics grouped by supplier/category.'),
    ('FR-03', 'Compute Defective Parts Per Million (PPM) and percentage defect rates across suppliers, products, and defect reason codes.', 'BR-03', 'src/analytics.py', 'Defect PPM = (rejected / inspected) * 1,000,000; returns aggregated quality metrics and flags PPM > 1000.'),
    ('FR-04', 'Audit Purchase Price Variance (PPV) and currency cost delta between contracted baseline prices and actual invoiced prices.', 'BR-04', 'src/analytics.py', 'Variance (%) = ((invoiced - contracted) / contracted) * 100; flags positive cost overruns.'),
    ('FR-05', 'Detect and isolate delivery delay tail outliers (> 7 calendar days) and compute distribution skewness and kurtosis.', 'BR-05', 'src/analytics.py', 'Filters delays > 7 days; computes Fisher-Pearson skewness and excess kurtosis coefficients.'),
    ('FR-06', 'Synthesize delivery, quality, and cost metrics into a weighted Composite Performance Index (CPI) on a 0-100 scale.', 'BR-06', 'src/analytics.py', 'CPI = (0.40 * OTIF) + (0.35 * Quality) + (0.25 * Cost); validates weights sum to 1.0; outputs Tier A/B/C.'),
    ('FR-07', 'Quantify Spend-at-Risk exposure and total capital allocation concentrated in Tier C underperforming suppliers.', 'BR-07', 'src/analytics.py', 'Spend-at-Risk = (Spend Tier C / Total Spend) * 100; triggers warning if ratio exceeds 10%.'),
    ('FR-08', 'Generate monthly and quarterly time-series trend aggregations for OTIF% and Defect PPM across rolling quarters.', 'BR-08', 'src/analytics.py', 'Quarterly date grouping; computes rolling mean and trend slope for longitudinal visualization.'),
    ('FR-09', 'Evaluate governance rules to classify suppliers as "Renewal Eligible", "Conditional Monitoring", or "CAPA Mandatory".', 'BR-09', 'src/analytics.py', 'Deterministic rule evaluation based on CPI thresholds and consecutive late delivery counts.'),
    ('FR-10', 'Aggregate and benchmark lead times and delay days across shipping logistics modes (Air, Sea, Road, Rail) and geographic origins.', 'BR-10', 'src/analytics.py', 'Groupby shipping_mode and supplier_country; computes mean lead time and delay variance.'),
    ('FR-11', 'Provide reactive global sidebar filter widgets (Date Range, Supplier, Category, Status) with instant downstream propagation.', 'BR-11', 'src/data_loader.py, app.py', 'Filter changes dynamically trigger recalculation of all charts and metrics in < 1.0s.'),
    ('FR-12', 'Generate formatted, downloadable tabular snapshots of filtered scorecards and raw analytical records into CSV / Excel.', 'BR-12', 'src/utils.py, pages/1_Scorecard.py', 'Renders st.download_button delivering sanitized CSV/Excel payloads with custom timestamps.')
]

for i, row in enumerate(fr_data):
    for j, val in enumerate(row):
        c = fr_table.cell(i+1, j)
        c.text = val
        if (i % 2) == 1:
            set_cell_background(c, 'F7FAFC')
        c.paragraphs[0].runs[0].font.size = Pt(8)
        set_cell_margins(c, 60, 60, 60, 60)

doc.add_paragraph().paragraph_format.space_after = Pt(8)

# Section 4: Data Layer
h4 = doc.add_heading('4. Data Layer Architecture: Loading, Caching & Validation', level=1)
h4.paragraph_format.space_before = Pt(12)
h4.paragraph_format.space_after = Pt(6)

doc.add_paragraph(
    '4.1 Database Engine & Connection Management:\n'
    '• SQLite 3 (version 3.35+) embedded database located at data/app.db.\n'
    '• Thread-safe connection handling using context managers (with sqlite3.connect(...) as conn) with check_same_thread=False.\n'
    '• Write-Ahead Logging (PRAGMA journal_mode=WAL) enabled for high-concurrency read operations.\n\n'
    '4.2 Streamlit In-Memory Caching Strategy:\n'
    '• All database reads are wrapped with the @st.cache_data(ttl=3600, show_spinner=False) decorator.\n'
    '• Cached dataframes are treated as immutable singletons across user browser sessions, preventing duplicate disk I/O and reducing query execution time to < 5ms.\n\n'
    '4.3 Data Validation Engine (src/validator.py):\n'
    'Before analytics processing, all extracted tables pass through an automated validation barrier:\n'
    '• Schema & Type Integrity: Dates cast to datetime64[ns], quantities to int64, financial metrics to float64.\n'
    '• Boundary & Range Checks: delivered_qty >= 0, inspected_qty >= defective_qty, contract_price > 0.\n'
    '• Foreign Key Referential Integrity: Deliveries.order_id must resolve to Orders.order_id; Quality_Inspections.delivery_id must resolve to Deliveries.delivery_id.\n'
    '• Null Handling: Null actual_delivery_dates are imputed as \"In-Transit\" and isolated from completed OTIF calculations.'
)

# Section 5: Statistical Methods
h5 = doc.add_heading('5. Statistical Methodology for Operational Questions (OQ-01 to OQ-10)', level=1)
h5.paragraph_format.space_before = Pt(12)
h5.paragraph_format.space_after = Pt(6)

stat_table = doc.add_table(rows=11, cols=4)
stat_table.alignment = WD_TABLE_ALIGNMENT.CENTER
stat_headers = ['OQ ID', 'Analytical Goal', 'Statistical & Mathematical Method', 'Scientific Libraries Used']
for j, h in enumerate(stat_headers):
    c = stat_table.cell(0, j)
    c.text = h
    set_cell_background(c, '1A365D')
    p = c.paragraphs[0]
    p.runs[0].bold = True
    p.runs[0].font.color.rgb = RGBColor(255, 255, 255)
    p.runs[0].font.size = Pt(8.5)
    set_cell_margins(c, 80, 80, 80, 80)

stat_data = [
    ('OQ-01', 'OTIF Compliance', 'Bernoulli Trial Aggregation: p_hat = sum(OTIF_i) / n. Binomial proportion 95% Wilson score confidence interval.', 'NumPy, SciPy.stats'),
    ('OQ-02', 'Lead Time & SLA', 'Parametric Mean (mu) and Std Dev (sigma); Non-parametric Median and Interquartile Range (IQR = Q3 - Q1). Mann-Whitney U test for SLA delta.', 'NumPy, SciPy.stats'),
    ('OQ-03', 'Defect Rate & PPM', 'Defect PPM = (Rejected / Inspected) * 10^6. Poisson process modeling for defect counts. Pareto 80/20 cumulative distribution.', 'Pandas, NumPy'),
    ('OQ-04', 'Price Variance', 'Percentage Relative Delta: ((P_inv - P_base) / P_base) * 100. Wilcoxon signed-rank test against zero price deviation hypothesis.', 'NumPy, SciPy.stats'),
    ('OQ-05', 'Tail Delay Outliers', 'Tukey Fences Outlier Isolation: Delay > Q3 + 1.5 * IQR. Fisher-Pearson standardized skewness and excess kurtosis.', 'SciPy.stats, Pandas'),
    ('OQ-06', 'Composite Scoring', 'Min-Max Feature Scaling: X_norm = (X - X_min) / (X_max - X_min). Weighted Linear Combination: CPI = sum(w_i * X_norm_i).', 'NumPy, Pandas'),
    ('OQ-07', 'Spend-at-Risk', 'Bivariate Spend vs Risk segmentation. Pareto Spend Concentration (Lorenz Curve and Gini Coefficient calculation on supplier spend).', 'NumPy, SciPy.stats'),
    ('OQ-08', 'Seasonal Trends', 'Rolling window moving average (3-month window). Ordinary Least Squares (OLS) linear trend regression: y = beta_0 + beta_1 * t + e.', 'statsmodels, NumPy'),
    ('OQ-09', 'Governance Triggers', 'Deterministic Decision-Tree Rule Induction combining continuous CPI score thresholds and discrete consecutive late order run-lengths.', 'Python standard logic'),
    ('OQ-10', 'Logistics Modes', 'One-Way Analysis of Variance (ANOVA) and Kruskal-Wallis H-test comparing lead-time variance across freight modes (Air, Sea, Road, Rail).', 'SciPy.stats')
]

for i, row in enumerate(stat_data):
    for j, val in enumerate(row):
        c = stat_table.cell(i+1, j)
        c.text = val
        if (i % 2) == 1:
            set_cell_background(c, 'F7FAFC')
        c.paragraphs[0].runs[0].font.size = Pt(8)
        set_cell_margins(c, 60, 60, 60, 60)

doc.add_paragraph().paragraph_format.space_after = Pt(8)

# Section 6: UI & Styling Architecture
h6 = doc.add_heading('6. User Interface, Tailwind CDN & Responsive Styling', level=1)
h6.paragraph_format.space_before = Pt(12)
h6.paragraph_format.space_after = Pt(6)

doc.add_paragraph(
    '6.1 Tailwind CSS Integration via iframe Containers:\n'
    'To achieve enterprise-grade visual polish beyond standard Streamlit widgets, custom HTML/CSS components are injected using streamlit.components.v1.html(..., height=...).\n'
    '• Tailwind CSS CDN script (<script src=\"https://cdn.tailwindcss.com\"></script>) is injected inside iframe headers.\n'
    '• Executive KPI Cards feature shadow-md, rounded-xl, gradient accent bars, and micro-typography.\n'
    '• Vendor Governance Tier Badges (Tier A Emerald, Tier B Amber, Tier C Red) provide instant visual status indicators.\n\n'
    '6.2 Theming & Layout (.streamlit/config.toml):\n'
    '• Base Theme: Light mode default with navy blue accents (#1A365D) and slate gray neutrals (#4B5563).\n'
    '• Layout Configuration: layout=\"wide\" enabled across all pages, optimizing desktop and laptop screen real estate (1366x768 to 1920x1080).\n'
    '• Plotly Visualization Palette: High-contrast accessible color scales (Emerald #10B981, Crimson #EF4444, Indigo #4F46E5) with template=\"plotly_white\".'
)

# Section 7: Performance, Error Handling & Security
h7 = doc.add_heading('7. Performance SLAs, Error Handling & Security Architecture', level=1)
h7.paragraph_format.space_before = Pt(12)
h7.paragraph_format.space_after = Pt(6)

doc.add_paragraph(
    '7.1 Performance Benchmarks & SLAs:\n'
    '• Initial Dashboard Cold Start: < 2.5 seconds (including SQLite loading and cached memory indexing).\n'
    '• Interactive Filter Recalculation Latency: < 1.0 second across 5,000 purchase order line items.\n'
    '• Memory Consumption: Maintained under 350 MB RAM, well within Streamlit Community Cloud\'s 1.0 GB ceiling.\n\n'
    '7.2 Defensive Error Handling Architecture:\n'
    '• Graceful Degradation: All analytical routines wrap computations in try...except blocks, logging failures and rendering st.info or st.warning notices rather than uncaught stack traces.\n'
    '• Division by Zero Protection: Automatic safety guards (e.g., if inspected_units == 0: return 0.0) applied across all ratio metrics.\n'
    '• Missing Filter Handlers: Empty dataset filters display a clean \"No records matching current filter criteria\" callout.\n\n'
    '7.3 Security & Sandbox Integrity:\n'
    '• SQL Injection Prevention: Exclusively uses parameterized queries (cursor.execute(\"SELECT ... WHERE id = ?\", (param,))).\n'
    '• Zero PII Exposure: All data is 100% synthetic supply chain records with zero personal identifiable information.\n'
    '• Read-Only Analytical Execution: Interactive user exploration is restricted to read-only queries with zero unauthorized database modifications.'
)

# Section 8: Pinned requirements.txt
h8 = doc.add_heading('8. Production Dependency Specifications (requirements.txt)', level=1)
h8.paragraph_format.space_before = Pt(12)
h8.paragraph_format.space_after = Pt(6)

p_req = doc.add_paragraph()
r_req = p_req.add_run(
    '# Core Analytical Application Stack (Python 3.12 Compatible)\n'
    'streamlit==1.38.0\n'
    'pandas==2.2.2\n'
    'numpy==1.26.4\n'
    'scipy==1.14.0\n'
    'statsmodels==0.14.2\n'
    'plotly==5.24.0\n'
    'openpyxl==3.1.5\n'
    'python-docx==1.2.0\n\n'
    '# Automated Testing Suite\n'
    'pytest==8.3.2\n'
    'pytest-cov==5.0.0'
)
r_req.font.name = 'Courier New'
r_req.font.size = Pt(9)
p_req.paragraph_format.space_after = Pt(8)

# Section 9: Deployment Requirements
h9 = doc.add_heading('9. Cloud Deployment & CI/CD Architecture', level=1)
h9.paragraph_format.space_before = Pt(12)
h9.paragraph_format.space_after = Pt(6)

doc.add_paragraph(
    '• Platform: Streamlit Community Cloud (hosted containerized Linux environment).\n'
    '• Repository Source: GitHub public/private repository linked via OAuth.\n'
    '• Trigger & Build Pipeline: Automatic deployment webhooks trigger rebuilds on each git push to main branch.\n'
    '• Environment Configuration: Python 3.12 specified in runtime environment configuration; dependencies resolved via root requirements.txt.\n'
    '• Persistence Architecture: data/app.db tracked directly in git repository for zero-setup database deployment.'
)

# Section 10: Technical Risks
h10 = doc.add_heading('10. Technical Risk Management Matrix', level=1)
h10.paragraph_format.space_before = Pt(12)
h10.paragraph_format.space_after = Pt(6)

tech_risk_table = doc.add_table(rows=5, cols=4)
tech_risk_table.alignment = WD_TABLE_ALIGNMENT.CENTER
tech_risk_headers = ['Risk ID', 'Technical Risk Description', 'Severity', 'Technical Mitigation Strategy']
for j, h in enumerate(tech_risk_headers):
    c = tech_risk_table.cell(0, j)
    c.text = h
    set_cell_background(c, '1A365D')
    p = c.paragraphs[0]
    p.runs[0].bold = True
    p.runs[0].font.color.rgb = RGBColor(255, 255, 255)
    p.runs[0].font.size = Pt(8.5)
    set_cell_margins(c, 80, 80, 80, 80)

tech_risks = [
    ('RISK-T01', 'Streamlit session memory bloat under heavy filter recalculations.', 'Medium', 'Enforce immutable cached dataframes using @st.cache_data(max_entries=50); perform transformations on lightweight views.'),
    ('RISK-T02', 'Plotly interactive charts rendering slowly on client mobile/laptop browsers.', 'Low', 'Pre-aggregate time series and category bins in Pandas before passing payload to Plotly; enable WebGL rendering for scatter plots.'),
    ('RISK-T03', 'Tailwind CSS CDN loading delay in iframe causing unstyled metric cards.', 'Low', 'Inject minimal fallback inline CSS styles inside HTML template to ensure readable typography prior to CDN script execution.'),
    ('RISK-T04', 'SQLite file locking conflicts if simultaneous background processes execute.', 'Low', 'Configure SQLite with PRAGMA busy_timeout=5000 and enable Write-Ahead Logging (WAL mode); enforce strictly read-only connections in Streamlit.')
]

for i, row in enumerate(tech_risks):
    for j, val in enumerate(row):
        c = tech_risk_table.cell(i+1, j)
        c.text = val
        if (i % 2) == 1:
            set_cell_background(c, 'F7FAFC')
        c.paragraphs[0].runs[0].font.size = Pt(8.5)
        set_cell_margins(c, 60, 60, 70, 70)

# Save Document
output_path = os.path.join('docs', '07_TRD.docx')
doc.save(output_path)
print(f'Successfully generated: {output_path}')
