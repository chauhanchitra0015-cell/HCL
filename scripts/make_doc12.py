import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

os.makedirs('docs', exist_ok=True)
doc = docx.Document()

# Margins
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
r_title = p_title.add_run('STEP 12: USER INTERFACE (UI) FRAMEWORK SPECIFICATION')
r_title.bold = True
r_title.font.size = Pt(22)
r_title.font.color.rgb = RGBColor(24, 43, 73)
p_title.paragraph_format.space_after = Pt(2)

p_sub = doc.add_paragraph()
r_sub = p_sub.add_run('Streamlit Multi-Page Architecture, Tailwind UI Components & Data Visualization Standards')
r_sub.font.size = Pt(13)
r_sub.font.color.rgb = RGBColor(90, 105, 120)
p_sub.paragraph_format.space_after = Pt(14)

# Document Control Table
meta_table = doc.add_table(rows=6, cols=2)
meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
meta_data = [
    ('Project Title', 'Supplier Performance Analysis & Procurement Intelligence System'),
    ('Student Details', 'Chitra Chauhan | Roll No: 2500320140043 | MCA 2nd Year (Data Analytics), ABES'),
    ('Project Mentor', 'Prof. Kalpana Dwivedi (Department of Computer Applications, ABES EC)'),
    ('UI Tech Stack', 'Streamlit Multipage, Tailwind CSS (via components.html CDN), Plotly White'),
    ('Screen Specifications', '8 Distinct Multi-Page Screens (UI-01 through UI-08)'),
    ('Component Module', 'src/ui.py (Reusable Tailwind Cards, Headers, and Insight Callouts)')
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

doc.add_paragraph().paragraph_format.space_after = Pt(10)

# Section 1: Screens Catalog
h1 = doc.add_heading('1. Master Screens Architecture & User Mapping (UI-01 to UI-08)', level=1)
h1.paragraph_format.space_before = Pt(12)
h1.paragraph_format.space_after = Pt(6)

screen_table = doc.add_table(rows=9, cols=6)
screen_table.alignment = WD_TABLE_ALIGNMENT.CENTER
headers = ['UI ID', 'Screen Title & Route', 'Core Purpose', 'Key UI Components', 'Active Filters', 'User Stories']
for j, h in enumerate(headers):
    c = screen_table.cell(0, j)
    c.text = h
    set_cell_background(c, '1A365D')
    p = c.paragraphs[0]
    p.runs[0].bold = True
    p.runs[0].font.color.rgb = RGBColor(255, 255, 255)
    p.runs[0].font.size = Pt(8.5)
    set_cell_margins(c, 80, 80, 80, 80)

screens_data = [
    ('UI-01', 'Executive Overview (app.py)', 'High-level procurement health summary and macro KPI monitoring.', 'Tailwind Header, 4 KPI Metric Cards, Spend Distribution Donut, Quarterly OTIF Trendline.', 'Category, Supplier, Date Range, Status', 'US-03, US-04, US-12'),
    ('UI-02', 'Supplier Scorecard (pages/1_Scorecard.py)', 'Multi-metric CPI scoring, Tier A/B/C classification, and CAPA triggers.', 'Weight Sliders, CPI Scorecard Table, Radar Chart, Tier Badges, CAPA Notice Box.', 'Tier Filter (A/B/C), Minimum Rating Slider', 'US-13, US-14, US-15, US-16'),
    ('UI-03', 'Delivery Performance (pages/2_Delivery.py)', 'Fulfillment punctuality, lead time dispersion, and logistics benchmarking.', 'OTIF Leaderboard Bar Chart, Lead Time Boxplots, Outlier Delay Table, Freight Mode Comparison.', 'Category, Date Range, Shipping Mode', 'US-05, US-06, US-07, US-08'),
    ('UI-04', 'Quality Risk Analysis (pages/3_Quality.py)', 'Component defect rate auditing, Defect PPM, and root cause Pareto.', 'Defect PPM Ranking, Dual-Axis Pareto Chart, Quality Audit Table, Defect Scatter.', 'Supplier, Product, Defect Code', 'US-09, US-10'),
    ('UI-05', 'Cost Variance & Spend (pages/4_Cost.py)', 'Purchase Price Variance (PPV) audits and Spend-at-Risk tracking.', 'Spend-at-Risk Gauge, Price Variance Histogram, Disputed Line Item Table, Waterfall Plot.', 'Category, Invoiced vs Base Variance', 'US-11, US-12'),
    ('UI-06', 'Data Explorer & Export (pages/5_Explorer.py)', 'Granular slice-and-dice data grid with CSV/Excel download capabilities.', '42-column interactive dataframe, multi-column search, CSV/Excel download buttons.', 'Global multi-select cross-filters', 'US-04, US-16'),
    ('UI-07', 'Data Quality Diagnostics (pages/6_Quality_Audit.py)', 'Database integrity monitoring, foreign key validation, and anomaly checks.', 'FK Integrity Badge, Missing Value Heatmap, Outlier Summary Box, Row Counter.', 'None (System Administrative View)', 'US-01, US-02'),
    ('UI-08', 'About & Architecture (pages/7_About.py)', 'Academic background, solution architecture, methodology, and citations.', 'Architecture Flowchart, Technology Stack Pill Badges, Academic Credits (ABES MCA).', 'None (Informational View)', 'Documentation')
]

for i, row in enumerate(screens_data):
    for j, val in enumerate(row):
        c = screen_table.cell(i+1, j)
        c.text = val
        if (i % 2) == 1:
            set_cell_background(c, 'F7FAFC')
        c.paragraphs[0].runs[0].font.size = Pt(8)
        if j in [0, 4, 5]:
            c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_cell_margins(c, 60, 60, 60, 60)

doc.add_paragraph().paragraph_format.space_after = Pt(8)

# Section 2: ASCII Wireframes
h2 = doc.add_heading('2. ASCII Wireframe Specifications for Core Screens', level=1)
h2.paragraph_format.space_before = Pt(12)
h2.paragraph_format.space_after = Pt(6)

wireframes = [
    ('UI-01: Executive Overview Wireframe',
"""+---------------------------------------------------------------------------------------------------+
|  [LOGO] SUPPLIER PERFORMANCE INTELLIGENCE PLATFORM                [STATUS: SQLite Connected (Live)]|
|  Comprehensive Procurement KPI Dashboard • ABES MCA Capstone 2024-2026                            |
+---------------------------------------------------------------------------------------------------+
| SIDEBAR FILTERS          | MAIN DASHBOARD CONTENT AREA                                            |
| -------------------------+------------------------------------------------------------------------|
| Category: [All Categories] | [CARD 1: Overall OTIF%] [CARD 2: Mean Lead Time] [CARD 3: Defect PPM] [CARD 4: Spend]  |
| Supplier: [All Suppliers]  |  Value: 88.6%           Value: 18.4 Days         Value: 740 PPM      Value: INR 14.8M  |
| Date: [2024-01 to 2024-12] |  Status: Warning        Status: In-Range         Status: Compliant   Status: Active    |
| Order Status: [Completed]  +------------------------------------------------------------------------+
|                            | ROW 1: MACRO TRENDS & SPEND DISTRIBUTION                               |
| [Reset Filters Button]     | +------------------------------------+-------------------------------+ |
|                            | | Quarterly OTIF% Trend (Line Plotly) | Spend by Category (Donut Chart)| |
|                            | | X: Q1, Q2, Q3, Q4 | Y: OTIF %       | Outer: Category | Val: Spend  | |
|                            | +------------------------------------+-------------------------------+ |
|                            | ROW 2: CRITICAL ALERTS & SPEND-AT-RISK GAUGE                           |
|                            | +--------------------------------------------------------------------+ |
|                            | | [!] ALERT: Spend allocated to Tier C suppliers is 9.8% (Target <10%)| |
|                            | | Top 3 Suppliers requiring immediate review: SUP-021, SUP-023, SUP-025|
|                            | +--------------------------------------------------------------------+ |
+----------------------------+------------------------------------------------------------------------+"""),

    ('UI-02: Supplier Scorecard & CPI Tiering Wireframe',
"""+---------------------------------------------------------------------------------------------------+
|  SUPPLIER COMPOSITE PERFORMANCE SCORECARD (CPI) & VENDOR GOVERNANCE                                |
+---------------------------------------------------------------------------------------------------+
| SIDEBAR CONTROLS         | MAIN SCORECARD CONTENT                                                 |
| -------------------------+------------------------------------------------------------------------|
| WEIGHT CONFIGURATION     | ROW 1: TIER DISTRIBUTION SUMMARY                                       |
| Delivery Wt: [====o] 40% | [ Tier A: 8 Vendors (32%) ]  [ Tier B: 12 Vendors (48%) ]  [ Tier C: 5 Vendors (20%) ]|
| Quality Wt : [===o-] 35% +------------------------------------------------------------------------+
| Cost Wt    : [==o--] 25% | ROW 2: MULTI-AXIS VENDOR COMPARISON                                    |
| [Recalculate Scores]     | +------------------------------------+-------------------------------+ |
|                          | | Composite Score Ranking (Bar Plot) | Radar Chart (Delivery/Quality/Cost)
| TIER FILTER              | | X: Vendor Name | Y: CPI Score      | Polar: 3 Axes for Selected Vendor|
| [X] Tier A  [X] Tier B   | | Bars: Green (>=85), Amber, Red (<70)| Benchmark against Cohort Avg  | |
| [X] Tier C               | +------------------------------------+-------------------------------+ |
|                          | ROW 3: DETAILED SCORECARD TABLE & ACTIONABLE GOVERNANCE                |
|                          | +--------------------------------------------------------------------+ |
|                          | | ID     | Vendor Name       | OTIF% | PPM   | Price Var% | CPI | Action   | |
|                          | | SUP-001| Bharat Precision  | 96.2% | 420   | +0.2%      | 94.8| Renew [V]| |
|                          | | SUP-021| Ananya Supplies   | 61.4% | 4850  | +5.8%      | 62.1| CAPA [!]| |
|                          | +--------------------------------------------------------------------+ |
+--------------------------+------------------------------------------------------------------------+"""),

    ('UI-03: Delivery Performance & Lead Time Analytics Wireframe',
"""+---------------------------------------------------------------------------------------------------+
|  DELIVERY ADHERENCE, LEAD TIME DISPERSION & OUTLIER DETECTION                                     |
+---------------------------------------------------------------------------------------------------+
| FILTERS                  | MAIN DELIVERY ANALYTICS                                                |
| -------------------------+------------------------------------------------------------------------|
| Category: [All]          | [ Overall OTIF: 88.6% ]  [ Delayed Orders: 42 ]  [ Avg Delay: 3.4 Days ] |
| Shipping: [All Modes]    +------------------------------------------------------------------------+
| Threshold: [Delay > 7d]  | ROW 1: OTIF LEADERBOARD & CATEGORY LEAD TIME DISPERSION                |
|                          | +------------------------------------+-------------------------------+ |
|                          | | Vendor OTIF Leaderboard (Horiz Bar)| Lead Time Boxplot vs SLA Line  | |
|                          | | Y: Vendor Names | X: OTIF %        | Category Boxplots (Q1, Med, Q3)| |
|                          | | Dotted vertical target at 95.0%    | Dotted reference line at SLA   | |
|                          | +------------------------------------+-------------------------------+ |
|                          | ROW 2: TRANSIT LOGISTICS & SEVERE DELAY OUTLIER AUDIT                  |
|                          | +------------------------------------+-------------------------------+ |
|                          | | Freight Shipping Mode Breakdown    | Delay Tail Outliers (> 7 Days)| |
|                          | | Air Freight vs Road vs Sea vs Rail | DEL-0023 | Delay: 34d | Castings | |
|                          | | Mean delay days and volume bar     | DEL-0069 | Delay: 32d | Valves   | |
|                          | +------------------------------------+-------------------------------+ |
+--------------------------+------------------------------------------------------------------------+"""),

    ('UI-04: Quality Risk Analysis & Pareto Wireframe',
"""+---------------------------------------------------------------------------------------------------+
|  RECEIVING QUALITY INSPECTION, DEFECT PPM & PARETO ROOT CAUSE AUDIT                               |
+---------------------------------------------------------------------------------------------------+
| FILTERS                  | MAIN QUALITY ANALYTICS                                                 |
| -------------------------+------------------------------------------------------------------------|
| Product: [All SKUs]      | [ Avg Defect PPM: 740 ]  [ Rejection Rate: 1.1% ]  [ Inspected: 18,450 ]  |
| Defect Code: [All Codes] +------------------------------------------------------------------------+
|                          | ROW 1: PARETO 80/20 ROOT CAUSE DECOMPOSITION                           |
|                          | +--------------------------------------------------------------------+ |
|                          | | Dual-Axis Pareto Chart (Bar Frequency + Cumulative Line)           | |
|                          | | Bars: Dimensional (45%), Surface (30%), Fatigue (12%), Packaging   | |
|                          | | Overlaid red line showing 80% cumulative threshold cutoff          | |
|                          | +--------------------------------------------------------------------+ |
|                          | ROW 2: VENDOR QUALITY BENCHMARKING & HIGH-RISK QUARANTINE LIST         |
|                          | +------------------------------------+-------------------------------+ |
|                          | | Defect PPM by Vendor (Scatter Plot)| Critical Defect Quarantine Log| |
|                          | | X: Defect PPM | Y: Inspected Qty   | Vendor   | SKU     | Defect   | |
|                          | | Dotted threshold at 1,000 PPM      | SUP-021  | GGG-40  | Dimension| |
|                          | +------------------------------------+-------------------------------+ |
+--------------------------+------------------------------------------------------------------------+""")
]

for title, w_code in wireframes:
    p = doc.add_paragraph()
    r = p.add_run(title)
    r.bold = True
    r.font.size = Pt(11)
    r.font.color.rgb = RGBColor(26, 54, 93)
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(2)
    
    p_box = doc.add_paragraph()
    r_c = p_box.add_run(w_code)
    r_c.font.name = 'Courier New'
    r_c.font.size = Pt(7.5)
    p_box.paragraph_format.space_after = Pt(6)

# Section 3: Global Layout & Theming
h3 = doc.add_heading('3. Global Layout Architecture & Color Palette Standards', level=1)
h3.paragraph_format.space_before = Pt(12)
h3.paragraph_format.space_after = Pt(6)

doc.add_paragraph(
    '3.1 Global State Management (st.session_state):\n'
    'To guarantee seamless multi-page transitions, all sidebar filter selections (Category, Supplier, Date Boundaries, Order Status, and CPI Weights) are stored in st.session_state. Navigating between Overview, Scorecard, Delivery, and Quality preserves active filter boundaries without resetting.\n\n'
    '3.2 Brand Color Palette:\n'
    '• Primary Navy Accent: #1A365D (Headers, primary brand framing, active tabs)\n'
    '• Brand Indigo: #4F46E5 (Key interactive controls, Plotly accent series)\n'
    '• Success Emerald: #10B981 (Tier A Preferred, OTIF Compliant >= 95%, Low PPM)\n'
    '• Warning Amber: #F59E0B (Tier B Conditional, Moderate Delays, PPM 800-1000)\n'
    '• Danger Rose: #EF4444 (Tier C At-Risk, Severe Delays > 7d, Price Creep > 3%)\n'
    '• Neutral Light Background: #F8FAFC (Secondary backgrounds, table zebra striping)\n\n'
    '3.3 Responsive & Mobile Adaptive Layout:\n'
    'Using Tailwind\'s grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 classes, KPI card rows automatically collapse from a 4-column horizontal desktop strip into a 2-column tablet layout or single-column mobile view without horizontal scroll clipping.'
)

# Section 4: Best Chart per OQ Table
h4 = doc.add_heading('4. Visualization Matrix: Best Chart per Operational Question (OQ-01 to OQ-10)', level=1)
h4.paragraph_format.space_before = Pt(12)
h4.paragraph_format.space_after = Pt(6)

chart_table = doc.add_table(rows=11, cols=4)
chart_table.alignment = WD_TABLE_ALIGNMENT.CENTER
chart_headers = ['OQ ID', 'Operational Goal', 'Optimal Visualization Type', 'Analytical Rationale']
for j, h in enumerate(chart_headers):
    c = chart_table.cell(0, j)
    c.text = h
    set_cell_background(c, '1A365D')
    p = c.paragraphs[0]
    p.runs[0].bold = True
    p.runs[0].font.color.rgb = RGBColor(255, 255, 255)
    p.runs[0].font.size = Pt(8.5)
    set_cell_margins(c, 80, 80, 80, 80)

chart_data = [
    ('OQ-01', 'OTIF Compliance Leaderboard', 'Horizontal Ranked Bar Chart with Target SLA Line', 'Horizontal bars provide legible space for vendor names; color coding instantly flags compliant (>=95%) vs failing suppliers against a vertical threshold line.'),
    ('OQ-02', 'Lead Time & SLA Variability', 'Grouped Box & Whisker Plot by Commodity Category', 'Unlike simple averages, boxplots communicate median, dispersion (IQR), and extreme tail outliers across categories alongside contracted SLA benchmark markers.'),
    ('OQ-03', 'Defect PPM & Root Causes', 'Dual-Axis Pareto Chart (Bar Frequency + Cumulative % Line)', 'Standard industrial quality control tool that highlights the \"vital few\" failure causes accounting for 80% of part rejections.'),
    ('OQ-04', 'Price Variance / Cost Creep', 'Diverging Bar / Scatter Waterfall (Price Delta % vs Spend)', 'Diverging red/green colors highlight price increases vs cost savings, while bubble sizes reflect total procurement volume.'),
    ('OQ-05', 'Severe Delay Tail Outliers (>7d)', 'Histogram with KDE Fit & Outlier Jitter Strip', 'Visualizes heavy-tailed right-skewness and excess kurtosis, pinpointing catastrophic delay probabilities.'),
    ('OQ-06', 'Composite Scorecard (CPI)', 'Multi-Metric Polar Radar Chart & Tiered Leaderboard', 'Visualizes simultaneous trade-offs across the 3 competing dimensions (Delivery, Quality, Cost) for any selected vendor.'),
    ('OQ-07', 'Spend-at-Risk Allocation', 'Interactive Donut Gauge & 2x2 Risk-Spend Scatter Grid', 'Plots Vendor Spend (Y) against Performance Tier (X), isolating high-spend, low-performance vendors in the danger zone.'),
    ('OQ-08', 'Seasonal & Quarterly Trends', 'Multi-Series Time-Series Line with Rolling Moving Average', 'Rolling trendlines filter out transactional weekly noise to expose cyclical Q3/Q4 performance dips.'),
    ('OQ-09', 'Contract Renewal & CAPA Triggers', 'Categorical Governance Status Kanban Matrix', 'Classifies vendors into clear action bins: Renewal Expansion (Green), Conditional Monitor (Amber), and CAPA Mandatory (Red).'),
    ('OQ-10', 'Shipping Modes & Geography', 'Grouped Violin Plot / Comparative Bar Breakdown', 'Exposes transit delay probability distributions across logistics modes (Air, Sea, Road, Rail) and supplier regions.')
]

for i, row in enumerate(chart_data):
    for j, val in enumerate(row):
        c = chart_table.cell(i+1, j)
        c.text = val
        if (i % 2) == 1:
            set_cell_background(c, 'F7FAFC')
        c.paragraphs[0].runs[0].font.size = Pt(8)
        if j == 0:
            c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_cell_margins(c, 60, 60, 60, 60)

doc.add_paragraph().paragraph_format.space_after = Pt(8)

# Section 5: Reusable Tailwind Components & Configuration
h5 = doc.add_heading('5. Reusable Tailwind Components & Configuration Implementation', level=1)
h5.paragraph_format.space_before = Pt(12)
h5.paragraph_format.space_after = Pt(6)

doc.add_paragraph(
    '5.1 Streamlit Theme Configuration (.streamlit/config.toml):\n'
    'Configured with primaryColor=\"#1A365D\", backgroundColor=\"#FFFFFF\", secondaryBackgroundColor=\"#F8FAFC\", textColor=\"#0F172A\", and font=\"sans serif\".\n\n'
    '5.2 Starter Component Module (src/ui.py):\n'
    'Implemented reusable Python functions: render_tailwind_header(title, subtitle, badge), render_kpi_cards(kpi_data), and render_insight_box(title, message, alert_type) utilizing streamlit.components.v1.html() and Tailwind CDN.'
)

# Save Document
output_docx = os.path.join('docs', '12_UI_Framework.docx')
doc.save(output_docx)
print(f'Successfully generated: {output_docx}')

# -------------------------------------------------------------
# Update RTM Matrix with UI Screen IDs
# -------------------------------------------------------------
rtm_path = os.path.join('docs', '04_Traceability_Matrix_Updated.xlsx')
wb = openpyxl.load_workbook(rtm_path)
ws_rtm = wb["RTM_Matrix"]

ui_mappings = [
    # Row 2 (BR-01, US-05) -> UI-01, UI-03
    (2, "UI-01, UI-03", "UI Designed"),
    # Row 3 (BR-11, US-04) -> UI-01..06 (Sidebar)
    (3, "UI-01..06 (Sidebar)", "UI Designed"),
    # Row 4 (BR-02, US-06) -> UI-01, UI-03
    (4, "UI-01, UI-03", "UI Designed"),
    # Row 5 (BR-11, US-04) -> UI-01..06 (Sidebar)
    (5, "UI-01..06 (Sidebar)", "UI Designed"),
    # Row 6 (BR-03, US-09, 10) -> UI-01, UI-04
    (6, "UI-01, UI-04", "UI Designed"),
    # Row 7 (BR-11, US-04) -> UI-01..06 (Sidebar)
    (7, "UI-01..06 (Sidebar)", "UI Designed"),
    # Row 8 (BR-04, US-11) -> UI-05
    (8, "UI-05", "UI Designed"),
    # Row 9 (BR-05, US-07) -> UI-03
    (9, "UI-03", "UI Designed"),
    # Row 10 (BR-06, US-13, 14) -> UI-01, UI-02
    (10, "UI-01, UI-02", "UI Designed"),
    # Row 11 (BR-12, US-16) -> UI-02, UI-06
    (11, "UI-02, UI-06", "UI Designed"),
    # Row 12 (BR-07, US-12) -> UI-01, UI-05
    (12, "UI-01, UI-05", "UI Designed"),
    # Row 13 (BR-08, US-03, 05) -> UI-01, UI-03
    (13, "UI-01, UI-03", "UI Designed"),
    # Row 14 (BR-09, US-15) -> UI-02
    (14, "UI-02", "UI Designed"),
    # Row 15 (BR-12, US-16) -> UI-02, UI-06
    (15, "UI-02, UI-06", "UI Designed"),
    # Row 16 (BR-10, US-08) -> UI-03
    (16, "UI-03", "UI Designed")
]

BORDER_GRAY = "D1D5DB"
thin_border = Border(
    left=Side(style='thin', color=BORDER_GRAY),
    right=Side(style='thin', color=BORDER_GRAY),
    top=Side(style='thin', color=BORDER_GRAY),
    bottom=Side(style='thin', color=BORDER_GRAY)
)

for row_num, ui_id, status in ui_mappings:
    c_status = ws_rtm.cell(row=row_num, column=7)
    c_status.value = status
    c_status.font = Font(name="Calibri", size=10, bold=True, color="059669")
    
    c_ui = ws_rtm.cell(row=row_num, column=11)
    c_ui.value = ui_id
    c_ui.alignment = Alignment(horizontal="center", vertical="center")
    c_ui.font = Font(name="Calibri", size=10, bold=True, color="D97706")
    c_ui.border = thin_border

ws_rtm.column_dimensions["K"].width = 22

wb.save(rtm_path)
print(f'Successfully updated: {rtm_path}')
