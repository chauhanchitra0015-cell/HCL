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
r_title = p_title.add_run('STEP 11: AGILE USER STORIES CATALOG & SPRINT PLAN')
r_title.bold = True
r_title.font.size = Pt(22)
r_title.font.color.rgb = RGBColor(24, 43, 73)
p_title.paragraph_format.space_after = Pt(2)

p_sub = doc.add_paragraph()
r_sub = p_sub.add_run('User-Centric Agile Requirements, Given/When/Then Acceptance Criteria & Sprint Schedule')
r_sub.font.size = Pt(13)
r_sub.font.color.rgb = RGBColor(90, 105, 120)
p_sub.paragraph_format.space_after = Pt(14)

# Document Control Table
meta_table = doc.add_table(rows=6, cols=2)
meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
meta_data = [
    ('Project Title', 'Supplier Performance Analysis & Procurement Intelligence System'),
    ('Student Details', 'Chitra Chauhan | Roll No: 2500320140043 | MCA 2nd Year (Data Analytics), ABES'),
    ('Academic Mentor', 'Mr. Dinesh (Department of Computer Applications, ABES EC)'),
    ('Total User Stories', '16 User Stories across 5 Agile Epics (40 Total Story Points)'),
    ('Velocity Benchmark', '4 Story Points / Day (Buildable in under half a day per story)'),
    ('Traceability Mapping', '100% Mapped to OQ-01..10, BR-01..12, and FR-01..12 in RTM')
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

# Section 1: Agile Methodology
h1 = doc.add_heading('1. Agile User Story Framework & Estimation Standard', level=1)
h1.paragraph_format.space_before = Pt(12)
h1.paragraph_format.space_after = Pt(6)

doc.add_paragraph(
    'In Step 11, the formal technical requirements (FR) and business requirements (BR) are decomposed into granular, buildable Agile User Stories following the standard INVEST criteria (Independent, Negotiable, Valuable, Estimable, Small, and Testable). Every user story is estimated on the Fibonacci scale (1 to 8 story points), where each story represents less than half a day of development effort. Stories are grouped into 5 thematic Epics covering the entire product lifecycle from data persistence to executive governance.'
)

# Section 2: User Story Master Table
h2 = doc.add_heading('2. Master User Stories Catalog (US-01 to US-16)', level=1)
h2.paragraph_format.space_before = Pt(12)
h2.paragraph_format.space_after = Pt(6)

stories_table = doc.add_table(rows=17, cols=7)
stories_table.alignment = WD_TABLE_ALIGNMENT.CENTER
headers = ['US ID', 'Persona', 'User Story Statement (As a / I want / So that)', 'FR ID', 'BR ID', 'Priority', 'Points']
for j, h in enumerate(headers):
    c = stories_table.cell(0, j)
    c.text = h
    set_cell_background(c, '1A365D')
    p = c.paragraphs[0]
    p.runs[0].bold = True
    p.runs[0].font.color.rgb = RGBColor(255, 255, 255)
    p.runs[0].font.size = Pt(8.5)
    set_cell_margins(c, 80, 80, 80, 80)

user_stories_data = [
    ('US-01', 'Data Analyst', 'As a data analyst, I want to load and cache relational tables from SQLite (data/app.db) using @st.cache_data so that the application has instant, thread-safe data access.', 'FR-11', 'BR-11', 'Must', '3 pts'),
    ('US-02', 'QA Lead', 'As a QA lead, I want raw delivery data to be validated against physical range rules so that corrupted records are cleaned and prevented from polluting analytics.', 'FR-01', 'BR-03', 'Must', '2 pts'),
    ('US-03', 'Head of SCM', 'As Head of SCM, I want to view top-level KPI cards (OTIF%, Lead Time, Defect PPM, Spend) styled with Tailwind CSS so that I can gauge procurement health in under 5 seconds.', 'FR-01, 02', 'BR-01, 06', 'Must', '3 pts'),
    ('US-04', 'Sourcing Specialist', 'As a sourcing specialist, I want to filter data by Category, Supplier, Date Range, and Status via a sidebar so that all visual charts update synchronously.', 'FR-11', 'BR-11', 'Must', '3 pts'),
    ('US-05', 'Sourcing Specialist', 'As a sourcing specialist, I want to view an OTIF compliance leaderboard ranked by vendor with a 95% target reference line so that I can award orders to dependable suppliers.', 'FR-01', 'BR-01', 'Must', '2 pts'),
    ('US-06', 'Warehouse Supervisor', 'As a warehouse supervisor, I want to inspect lead time distribution boxplots grouped by category against standard SLA days so that I can calibrate safety stock buffers.', 'FR-02', 'BR-02', 'Must', '3 pts'),
    ('US-07', 'Sourcing Specialist', 'As a sourcing specialist, I want to isolate delivery tail outliers (> 7 calendar days delay) so that highly volatile suppliers can be removed from critical schedules.', 'FR-05', 'BR-05', 'Should', '2 pts'),
    ('US-08', 'Warehouse Supervisor', 'As a logistics planner, I want to benchmark lead times across shipping freight modes (Air, Sea, Road, Rail) so that transport routes can be optimized.', 'FR-10', 'BR-10', 'Should', '2 pts'),
    ('US-09', 'QA Manager', 'As a QA manager, I want to track Defective Parts Per Million (PPM) by vendor with an automated alert for suppliers exceeding 1,000 PPM so that sub-standard lots are quarantined.', 'FR-03', 'BR-03', 'Must', '3 pts'),
    ('US-10', 'QA Manager', 'As a QA manager, I want to view a defect root-cause Pareto chart combining defect count bars with an 80% cumulative threshold so that quality audits target primary failure causes.', 'FR-03', 'BR-03', 'Must', '3 pts'),
    ('US-11', 'Finance Controller', 'As a finance controller, I want to audit unit cost variances between contracted baselines and invoiced amounts so that unapproved price creep is recovered.', 'FR-04', 'BR-04', 'Must', '3 pts'),
    ('US-12', 'Head of SCM', 'As Head of SCM, I want to view the proportion of spend allocated to Tier C underperforming suppliers so that enterprise Spend-at-Risk remains below 10%.', 'FR-07', 'BR-07', 'Must', '2 pts'),
    ('US-13', 'Head of SCM', 'As Head of SCM, I want an objective Composite Performance Index (CPI) score (0-100 scale) synthesizing delivery, quality, and cost so that suppliers are classified into Tier A, B, and C.', 'FR-06', 'BR-06', 'Must', '3 pts'),
    ('US-14', 'Head of SCM', 'As Head of SCM, I want to interactively adjust the relative weights for Delivery, Quality, and Cost so that I can simulate what-if contract renegotiation models.', 'FR-06', 'BR-06', 'Should', '2 pts'),
    ('US-15', 'Head of SCM', 'As Head of SCM, I want automated governance classification badges (Renewal Eligible, Conditional, CAPA Mandatory) so that vendor accountability is enforced systematically.', 'FR-09', 'BR-09', 'Must', '2 pts'),
    ('US-16', 'Sourcing Specialist', 'As a sourcing specialist, I want to view an interactive raw data table with sorting and export filtered scorecards to CSV/Excel so that empirical data can be shared in vendor reviews.', 'FR-12', 'BR-12', 'Could', '2 pts')
]

for i, row in enumerate(user_stories_data):
    for j, val in enumerate(row):
        c = stories_table.cell(i+1, j)
        c.text = val
        if (i % 2) == 1:
            set_cell_background(c, 'F7FAFC')
        c.paragraphs[0].runs[0].font.size = Pt(8)
        if j in [0, 3, 4, 5, 6]:
            c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_cell_margins(c, 50, 50, 50, 50)

doc.add_paragraph().paragraph_format.space_after = Pt(8)

# Section 3: Acceptance Criteria per User Story (Given / When / Then)
h3 = doc.add_heading('3. Detailed Given / When / Then Acceptance Criteria', level=1)
h3.paragraph_format.space_before = Pt(12)
h3.paragraph_format.space_after = Pt(6)

criteria_catalog = [
    ('US-01: SQLite Ingestion & Connection Pooling',
     '• Given the relational SQLite database data/app.db exists,\n'
     '• When the application initialization script calls get_database_connection(),\n'
     '• Then connection is established using check_same_thread=False, PRAGMA foreign_keys=ON, and tables load into cached memory in < 500ms.'),

    ('US-02: Data Hygiene & Boundary Validation',
     '• Given raw delivery data contains intentional missing delivery dates and potential duplicates,\n'
     '• When src/validator.py executes clean_delivery_records(df),\n'
     '• Then in-transit null dates are safely isolated without imputation distortion, duplicate rows are deduplicated, and zero unhandled exceptions occur.'),

    ('US-03: Executive KPI Summary Banners (Tailwind Cards)',
     '• Given the application Home page is rendered in a web browser,\n'
     '• When the executive summary component executes,\n'
     '• Then four distinct Tailwind CSS styled cards render inside an iframe displaying: Overall OTIF%, Average Lead Time, Defect PPM, and Total Spend, with dynamic colored accent badges matching status thresholds.'),

    ('US-04: Global Reactive Multi-Dimensional Sidebar Filters',
     '• Given the user navigates across any multi-page view,\n'
     '• When the user selects \"Mechanical Castings\" in Category and \"2024-01-01 to 2024-06-30\" in Date Range,\n'
     '• Then 100% of downstream charts and metric calculations dynamically filter to the matching records in < 1.0 second.'),

    ('US-05: On-Time In-Full (OTIF) Compliance Ranking',
     '• Given the Delivery Performance page is open,\n'
     '• When viewing the OTIF leaderboard horizontal bar chart,\n'
     '• Then vendors are ordered descending by OTIF%, a vertical dotted reference line is placed at 95.0%, and bars are color-coded (Tier A Emerald, Tier B Amber, Tier C Red).'),

    ('US-06: Lead Time Distribution & SLA Delta Boxplots',
     '• Given the Delivery Performance view is selected,\n'
     '• When the category lead-time boxplot renders,\n'
     '• Then each category box displays minimum, Q1, median, Q3, and maximum whiskers, alongside marker points for SLA baselines and delay outliers.'),

    ('US-07: Delivery Delay Outlier Isolation (> 7 Days)',
     '• Given the Delivery Performance page is open,\n'
     '• When the user toggles the \"Severe Delay Outliers (>7 Days)\" filter,\n'
     '• Then an isolated table displays only orders where delay_days > 7 with highlighted delay day counts and affected item descriptions.'),

    ('US-08: Shipping Logistics Mode & Transit Benchmarks',
     '• Given the Delivery Performance page is rendered,\n'
     '• When viewing the transport logistics section,\n'
     '• Then a grouped bar chart displays mean lead times and transit delays across freight modes (Air Freight, Sea Ocean, Road Express, Rail Cargo).'),

    ('US-09: Incoming Component Defect PPM Tracking',
     '• Given the Quality Risk Analysis page is loaded,\n'
     '• When the defect metric calculation executes,\n'
     '• Then each supplier\'s Defect PPM is computed as (rejected / inspected) * 1,000,000, and any vendor exceeding 1,000 PPM displays an automated warning badge.'),

    ('US-10: Defect Root-Cause Pareto 80/20 Analysis',
     '• Given the Quality Risk Analysis page is open,\n'
     '• When viewing the Pareto defect chart,\n'
     '• Then defect reason bars are arranged in descending frequency with an overlaid cumulative percentage line showing the 80% cut-off point.'),

    ('US-11: Purchase Price Variance (PPV) & Cost Creep Audit',
     '• Given the Cost Variance page is open,\n'
     '• When viewing invoice cost discrepancies,\n'
     '• Then an interactive audit table lists line items with positive price variance (> 0%), displaying contracted price, invoiced price, variance %, and net monetary impact.'),

    ('US-12: Spend-at-Risk Exposure Analytics',
     '• Given the Cost Variance page is open,\n'
     '• When the spend-at-risk module executes,\n'
     '• Then a visual gauge and donut chart display the percentage of total spend allocated to Tier C suppliers, flagging an alert if Spend-at-Risk > 10%.'),

    ('US-13: Composite Performance Index (CPI) Scorecard',
     '• Given the Supplier Scorecard page is open,\n'
     '• When the scorecard engine computes composite scores,\n'
     '• Then CPI is computed as 0.40 * OTIF + 0.35 * Quality + 0.25 * Cost, classifying vendors into Tier A (>= 85), Tier B (70-84.9), and Tier C (< 70).'),

    ('US-14: Interactive CPI Weight Sensitivity Sliders',
     '• Given the user adjusts the Delivery weight slider from 40% to 50% in the sidebar,\n'
     '• When the slider value is released,\n'
     '• Then the system normalizes remaining weights to sum to 100% and dynamically updates all vendor composite rankings without page reload.'),

    ('US-15: Automated CAPA and Contract Renewal Triggers',
     '• Given a supplier record has CPI < 70 or consecutive late deliveries >= 3,\n'
     '• When viewing the vendor detail card,\n'
     '• Then the governance status is set to \"CAPA Audit Mandatory\" with an interactive action checklist for procurement managers.'),

    ('US-16: Tabular Scorecard Data Slicing & CSV Export',
     '• Given the Data Explorer page is open,\n'
     '• When the user clicks the \"Download Filtered Dataset as CSV\" button,\n'
     '• Then the browser downloads a sanitized, UTF-8 encoded CSV file containing 100% of matching filtered rows within 1.5 seconds.')
]

for title, criteria in criteria_catalog:
    p = doc.add_paragraph()
    r = p.add_run(title)
    r.bold = True
    r.font.size = Pt(10.5)
    r.font.color.rgb = RGBColor(26, 54, 93)
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(2)
    p_b = doc.add_paragraph(criteria)
    p_b.paragraph_format.line_spacing = 1.15
    p_b.paragraph_format.space_after = Pt(4)

# Section 4: Sprint Plan
h4 = doc.add_heading('4. 10-Day Agile Sprint Implementation Schedule', level=1)
h4.paragraph_format.space_before = Pt(12)
h4.paragraph_format.space_after = Pt(6)

sprint_table = doc.add_table(rows=11, cols=4)
sprint_table.alignment = WD_TABLE_ALIGNMENT.CENTER
sprint_headers = ['Sprint Day', 'Target Epics', 'Scheduled User Stories', 'Daily Points']
for j, h in enumerate(sprint_headers):
    c = sprint_table.cell(0, j)
    c.text = h
    set_cell_background(c, '1A365D')
    p = c.paragraphs[0]
    p.runs[0].bold = True
    p.runs[0].font.color.rgb = RGBColor(255, 255, 255)
    p.runs[0].font.size = Pt(8.5)
    set_cell_margins(c, 80, 80, 80, 80)

sprint_schedule = [
    ('Day 1 (Mon)', 'Epic 1: Data Architecture', 'US-01 (DB Ingestion & Cache), US-02 (Data Validation)', '5 pts'),
    ('Day 2 (Tue)', 'Epic 2: Core Dashboard', 'US-03 (Tailwind KPI Banners), US-04 (Global Sidebar Filters)', '6 pts'),
    ('Day 3 (Wed)', 'Epic 3: Delivery Analysis', 'US-05 (OTIF Compliance Leaderboard), US-06 (Lead Time Boxplots)', '5 pts'),
    ('Day 4 (Thu)', 'Epic 3: Delivery Analysis', 'US-07 (Tail Delay Outliers), US-08 (Logistics & Geographic Modes)', '4 pts'),
    ('Day 5 (Fri)', 'Mid-Sprint Milestone', 'Code Review, Integration Testing, and Architecture Validation', 'Milestone Review'),
    ('Day 6 (Mon)', 'Epic 4: Quality & Cost Risk', 'US-09 (Defect PPM Tracking), US-10 (Defect Pareto 80/20 Chart)', '6 pts'),
    ('Day 7 (Tue)', 'Epic 4: Quality & Cost Risk', 'US-11 (Purchase Price Variance), US-12 (Spend-at-Risk Gauge)', '5 pts'),
    ('Day 8 (Wed)', 'Epic 5: Strategic Insights', 'US-13 (CPI Scorecard Engine), US-14 (Dynamic Weight Sliders)', '5 pts'),
    ('Day 9 (Thu)', 'Epic 5: Strategic Insights', 'US-15 (CAPA Governance Triggers), US-16 (Data Grid & CSV Export)', '4 pts'),
    ('Day 10 (Fri)', 'Sprint Finalization', 'End-to-End Regression Testing, Documentation, Mentor Demo', 'Final Acceptance')
]

for i, row in enumerate(sprint_schedule):
    for j, val in enumerate(row):
        c = sprint_table.cell(i+1, j)
        c.text = val
        if (i % 2) == 1:
            set_cell_background(c, 'F7FAFC')
        c.paragraphs[0].runs[0].font.size = Pt(8.5)
        set_cell_margins(c, 60, 60, 70, 70)

# Save Document
output_docx = os.path.join('docs', '11_User_Stories.docx')
doc.save(output_docx)
print(f'Successfully generated: {output_docx}')

# -------------------------------------------------------------
# Update RTM Matrix with User Story IDs
# -------------------------------------------------------------
rtm_path = os.path.join('docs', '04_Traceability_Matrix_Updated.xlsx')
wb = openpyxl.load_workbook(rtm_path)
ws_rtm = wb["RTM_Matrix"]

us_mappings = [
    # Row 2 (BR-01) -> US-05
    (2, "US-05", "Story Specified"),
    # Row 3 (BR-11) -> US-04
    (3, "US-04", "Story Specified"),
    # Row 4 (BR-02) -> US-06
    (4, "US-06", "Story Specified"),
    # Row 5 (BR-11) -> US-04
    (5, "US-04", "Story Specified"),
    # Row 6 (BR-03) -> US-09, US-10
    (6, "US-09, US-10", "Story Specified"),
    # Row 7 (BR-11) -> US-04
    (7, "US-04", "Story Specified"),
    # Row 8 (BR-04) -> US-11
    (8, "US-11", "Story Specified"),
    # Row 9 (BR-05) -> US-07
    (9, "US-07", "Story Specified"),
    # Row 10 (BR-06) -> US-13, US-14
    (10, "US-13, US-14", "Story Specified"),
    # Row 11 (BR-12) -> US-16
    (11, "US-16", "Story Specified"),
    # Row 12 (BR-07) -> US-12
    (12, "US-12", "Story Specified"),
    # Row 13 (BR-08) -> US-03, US-05
    (13, "US-03, US-05", "Story Specified"),
    # Row 14 (BR-09) -> US-15
    (14, "US-15", "Story Specified"),
    # Row 15 (BR-12) -> US-16
    (15, "US-16", "Story Specified"),
    # Row 16 (BR-10) -> US-08
    (16, "US-08", "Story Specified")
]

BORDER_GRAY = "D1D5DB"
thin_border = Border(
    left=Side(style='thin', color=BORDER_GRAY),
    right=Side(style='thin', color=BORDER_GRAY),
    top=Side(style='thin', color=BORDER_GRAY),
    bottom=Side(style='thin', color=BORDER_GRAY)
)

for row_num, us_id, status in us_mappings:
    c_status = ws_rtm.cell(row=row_num, column=7)
    c_status.value = status
    c_status.font = Font(name="Calibri", size=10, bold=True, color="059669")
    
    c_us = ws_rtm.cell(row=row_num, column=10)
    c_us.value = us_id
    c_us.alignment = Alignment(horizontal="center", vertical="center")
    c_us.font = Font(name="Calibri", size=10, bold=True, color="7C3AED")
    c_us.border = thin_border

ws_rtm.column_dimensions["J"].width = 18

wb.save(rtm_path)
print(f'Successfully updated: {rtm_path}')
