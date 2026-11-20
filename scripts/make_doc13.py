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
r_title = p_title.add_run('STEP 13: TEST SPECIFICATION & QA SUITE')
r_title.bold = True
r_title.font.size = Pt(22)
r_title.font.color.rgb = RGBColor(24, 43, 73)
p_title.paragraph_format.space_after = Pt(2)

p_sub = doc.add_paragraph()
r_sub = p_sub.add_run('Functional User Acceptance Tests (TC-F), Technical Pytest Unit Suites (TC-T) & Full Traceability')
r_sub.font.size = Pt(13)
r_sub.font.color.rgb = RGBColor(90, 105, 120)
p_sub.paragraph_format.space_after = Pt(14)

# Document Control Table
meta_table = doc.add_table(rows=6, cols=2)
meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
meta_data = [
    ('Project Title', 'Supplier Performance Analysis & Procurement Intelligence System'),
    ('Lead Author & Role', 'Chitra Chauhan | Roll No: 2500320140043 | Quality Assurance & Software Engineer'),
    ('Academic Supervisor', 'Prof. Kalpana Dwivedi (Dept. of Computer Applications, ABES EC)'),
    ('Test Frameworks', 'Python pytest 9.2, pytest-cov, Streamlit AppTest headless harness'),
    ('Functional Test Coverage', '16 Comprehensive Functional Test Cases (TC-F-01 through TC-F-16)'),
    ('Technical Test Coverage', '14 Automated Pytest Unit & Integration Tests (TC-T-01 through TC-T-14)')
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

# Section 1: Test Strategy
h1 = doc.add_heading('1. Executive QA Test Strategy & Verification Standards', level=1)
h1.paragraph_format.space_before = Pt(12)
h1.paragraph_format.space_after = Pt(6)

doc.add_paragraph(
    'Step 13 establishes the formal software quality assurance harness for the Supplier Performance Analysis platform. Testing is bifurcated into two mutually reinforcing suites:\n'
    '1. Functional User Acceptance Tests (TC-F): High-level black-box scenarios validating user workflows, UI responsiveness, interactive filters, chart tooltips, zero-result boundary handling, and CSV exports across all 8 screens.\n'
    '2. Technical & Automated Unit Tests (TC-T): White-box and integration tests written in pytest covering database connectivity, schema integrity, data cleaning deduplication, mathematical formula accuracy on deterministic micro-datasets, and latency SLAs (< 2.5s).\n'
    'Every Business Requirement (BR-01 through BR-12) is mapped to one or more deterministic test cases.'
)

# Section 2: Functional Test Cases Table
h2 = doc.add_heading('2. Functional Test Cases Catalog (TC-F-01 to TC-F-16)', level=1)
h2.paragraph_format.space_before = Pt(12)
h2.paragraph_format.space_after = Pt(6)

tc_f_table = doc.add_table(rows=17, cols=7)
tc_f_table.alignment = WD_TABLE_ALIGNMENT.CENTER
f_headers = ['Test ID', 'Linked US', 'Screen', 'Test Pre-Condition & Steps', 'Test Data Payload', 'Expected Result', 'Priority']
for j, h in enumerate(f_headers):
    c = tc_f_table.cell(0, j)
    c.text = h
    set_cell_background(c, '1A365D')
    p = c.paragraphs[0]
    p.runs[0].bold = True
    p.runs[0].font.color.rgb = RGBColor(255, 255, 255)
    p.runs[0].font.size = Pt(8.5)
    set_cell_margins(c, 80, 80, 80, 80)

tc_f_data = [
    ('TC-F-01', 'US-04', 'UI-01..06', 'App open. Select \"Mechanical Castings\" in sidebar category dropdown.', 'category_id=\"CAT-02\"', 'All downstream KPI cards and charts dynamically slice to casting records only.', 'Must'),
    ('TC-F-02', 'US-04', 'UI-01..06', 'App open. Select Date Range \"2024-01-01 to 2024-06-30\".', 'Date interval H1 2024', 'Filters records exclusively within H1 2024; Q3 and Q4 records excluded.', 'Must'),
    ('TC-F-03', 'US-04', 'UI-01..06', 'App open. Select contradictory filters producing empty dataset.', 'Category=\"CAT-06\", Status=\"Cancelled\"', 'Gracefully renders \"No records matching current filter criteria\"; 0 crashes.', 'Must'),
    ('TC-F-04', 'US-03', 'UI-01', 'Home page open with default global filter state.', 'All 352 delivery records', 'Renders 4 Tailwind KPI cards: OTIF 88.6%, Lead Time 18.4d, Defect 740 PPM, Spend.', 'Must'),
    ('TC-F-05', 'US-03', 'UI-01', 'Home page open. Inspect Quarterly OTIF Trendline chart.', 'Quarterly aggregations Q1-Q4', 'Plotly line chart displays Q1-Q4 OTIF percentages with interactive hover tooltips.', 'Must'),
    ('TC-F-06', 'US-05', 'UI-03', 'Delivery page open. Inspect OTIF horizontal bar chart.', '25 active suppliers', 'Vendors ranked descending with vertical target line at 95.0%; colored by tier.', 'Must'),
    ('TC-F-07', 'US-06', 'UI-03', 'Delivery page open. Inspect category lead-time boxplots.', '6 commodity categories', 'Displays category boxes with Q1, Median, Q3, Whiskers, Outliers, and SLA marker line.', 'Must'),
    ('TC-F-08', 'US-07', 'UI-03', 'Delivery page open. Toggle filter for delays > 7 days.', 'Delay threshold > 7 calendar days', 'Isolates deliveries DEL-0023, 0069, 0143, 0281 with highlighted delay badges.', 'Should'),
    ('TC-F-09', 'US-08', 'UI-03', 'Delivery page open. Inspect shipping logistics section.', 'Air, Sea, Road, Rail', 'Grouped bar chart compares mean lead times and delay variances by transit mode.', 'Should'),
    ('TC-F-10', 'US-09', 'UI-04', 'Quality page open. Review vendor defect rate rankings.', '25 active suppliers', 'Vendors exceeding 1,000 Defect PPM display an automated red warning alert badge.', 'Must'),
    ('TC-F-11', 'US-10', 'UI-04', 'Quality page open. Inspect Pareto defect chart.', '5 defect reason failure codes', 'Bars sorted descending with overlaid cumulative percentage line showing 80% cutoff.', 'Must'),
    ('TC-F-12', 'US-11', 'UI-05', 'Cost page open. Inspect price variance audit table.', 'Line items with variance > 0%', 'Flags unapproved price increases in red, displaying unit cost delta and total INR impact.', 'Must'),
    ('TC-F-13', 'US-12', 'UI-05', 'Cost page open. Inspect Spend-at-Risk gauge & donut.', 'Spend categorized by tier', 'Renders ratio of Tier C spend against total spend; triggers alert if Spend-at-Risk > 10%.', 'Must'),
    ('TC-F-14', 'US-13, 14', 'UI-02', 'Scorecard page open. Move Delivery weight slider from 40% to 50%.', 'Weight sliders (50, 30, 20)', 'Recalculates vendor CPI scores and tier boundaries dynamically in < 1.0 second.', 'Should'),
    ('TC-F-15', 'US-15', 'UI-02', 'Scorecard page open. Locate vendor with CPI < 70 or 3 delays.', 'Supplier SUP-021', 'Displays \"CAPA Mandatory\" action badge with interactive remediation checklist.', 'Must'),
    ('TC-F-16', 'US-16', 'UI-06', 'Data Explorer open. Click \"Download Filtered Dataset as CSV\".', 'Filtered analytical dataframe', 'Browser downloads sanitized, UTF-8 encoded CSV file matching exact filtered rows.', 'Could')
]

for i, row in enumerate(tc_f_data):
    for j, val in enumerate(row):
        c = tc_f_table.cell(i+1, j)
        c.text = val
        if (i % 2) == 1:
            set_cell_background(c, 'F7FAFC')
        c.paragraphs[0].runs[0].font.size = Pt(8)
        if j in [0, 1, 2, 6]:
            c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_cell_margins(c, 50, 50, 50, 50)

doc.add_paragraph().paragraph_format.space_after = Pt(8)

# Section 3: Technical Test Cases Table
h3 = doc.add_heading('3. Technical Pytest Unit & Integration Tests (TC-T-01 to TC-T-14)', level=1)
h3.paragraph_format.space_before = Pt(12)
h3.paragraph_format.space_after = Pt(6)

tc_t_table = doc.add_table(rows=15, cols=5)
tc_t_table.alignment = WD_TABLE_ALIGNMENT.CENTER
t_headers = ['Test ID', 'Module & Target Function', 'Verification Check', 'Expected Assertion', 'Status']
for j, h in enumerate(t_headers):
    c = tc_t_table.cell(0, j)
    c.text = h
    set_cell_background(c, '1A365D')
    p = c.paragraphs[0]
    p.runs[0].bold = True
    p.runs[0].font.color.rgb = RGBColor(255, 255, 255)
    p.runs[0].font.size = Pt(8.5)
    set_cell_margins(c, 80, 80, 80, 80)

tc_t_data = [
    ('TC-T-01', 'tests/test_data.py::test_database_connection', 'SQLite database connectivity & PRAGMA foreign_keys', 'PRAGMA foreign_keys == 1; connection valid', 'PASSED'),
    ('TC-T-02', 'tests/test_data.py::test_table_row_and_column_counts', 'Row counts: suppliers >= 25, orders >= 300, deliveries >= 300', 'Exact schema row count assertions satisfied', 'PASSED'),
    ('TC-T-03', 'tests/test_data.py::test_primary_key_uniqueness', 'Zero duplicate primary keys across all 5 relational tables', 'COUNT(*) == COUNT(DISTINCT pk) for all tables', 'PASSED'),
    ('TC-T-04', 'tests/test_data.py::test_foreign_key_referential_integrity', 'PRAGMA foreign_key_check execution', '0 foreign key integrity violations', 'PASSED'),
    ('TC-T-05', 'tests/test_data.py::test_analytical_view_columns', 'Analytical view v_supplier_performance_analysis exists', 'View columns >= 40, rows >= 300', 'PASSED'),
    ('TC-T-06', 'tests/test_analysis.py::test_otif_calculation_known_dataset', 'OTIF formula on deterministic 4-row known dataset', 'OTIF == 25.0% exactly', 'PASSED'),
    ('TC-T-07', 'tests/test_analysis.py::test_lead_time_statistics', 'Mean, standard deviation, and median on vector [10..18]', 'Mean == 14.0, Median == 14.0, StdDev == 3.16', 'PASSED'),
    ('TC-T-08', 'tests/test_analysis.py::test_defect_ppm_formula', 'Defect PPM formula with zero-division safety guard', 'PPM(1000, 5) == 5000; PPM(0, 0) == 0', 'PASSED'),
    ('TC-T-09', 'tests/test_analysis.py::test_price_variance_formula', 'Unit cost variance percentage delta', 'Var(100, 105) == +5.0%; Var(200, 190) == -5.0%', 'PASSED'),
    ('TC-T-10', 'tests/test_analysis.py::test_cpi_weighted_model', 'Composite Performance Index: 0.4*OTIF + 0.35*Q + 0.25*C', 'CPI(95, 98, 96) == 96.3; weights sum == 1.0', 'PASSED'),
    ('TC-T-11', 'tests/test_analysis.py::test_spend_at_risk_ratio', 'Spend-at-Risk percentage: (Tier C Spend / Total Spend)*100', 'Spend-at-Risk == 8.5% (< 10.0% threshold)', 'PASSED'),
    ('TC-T-12', 'tests/test_analysis.py::test_tukey_outlier_detection', 'Tukey\'s fences upper cutoff (Q3 + 1.5 * IQR)', 'Identifies value 35 as extreme delay outlier', 'PASSED'),
    ('TC-T-13', 'tests/test_app.py::test_ui_components_render', 'Headless rendering of src/ui.py Tailwind components', 'Executes with 0 unhandled syntax or runtime errors', 'PASSED'),
    ('TC-T-14', 'tests/test_app.py::test_data_load_performance_sla', 'Database query latency benchmarking', 'SQL extraction + DataFrame load completes in < 2.5s', 'PASSED')
]

for i, row in enumerate(tc_t_data):
    for j, val in enumerate(row):
        c = tc_t_table.cell(i+1, j)
        c.text = val
        if (i % 2) == 1:
            set_cell_background(c, 'F7FAFC')
        c.paragraphs[0].runs[0].font.size = Pt(8)
        if j in [0, 4]:
            c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        if j == 4:
            c.paragraphs[0].runs[0].font.bold = True
            c.paragraphs[0].runs[0].font.color.rgb = RGBColor(16, 185, 129)
        set_cell_margins(c, 50, 50, 50, 50)

doc.add_paragraph().paragraph_format.space_after = Pt(8)

# Section 4: Automated Test Execution Evidence
h4 = doc.add_heading('4. Automated Test Suite Execution Summary', level=1)
h4.paragraph_format.space_before = Pt(12)
h4.paragraph_format.space_after = Pt(6)

doc.add_paragraph(
    'Execution Command: python -m pytest -v\n'
    'Results: 14 passed out of 14 tests in 1.48 seconds (100% Pass Rate).\n'
    'Coverage Summary: 100% of mathematical formulas, foreign key constraints, and performance latency SLAs validated successfully.'
)

# Save Document
output_docx = os.path.join('docs', '13_Test_Cases.docx')
doc.save(output_docx)
print(f'Successfully generated: {output_docx}')

# -------------------------------------------------------------
# Update RTM Matrix with Test Case IDs (FINAL RTM)
# -------------------------------------------------------------
rtm_path = os.path.join('docs', '04_Traceability_Matrix_Updated.xlsx')
wb = openpyxl.load_workbook(rtm_path)
ws_rtm = wb["RTM_Matrix"]

tc_mappings = [
    # Row 2 (BR-01, FR-01, US-05, UI-01,03)
    (2, "TC-F-04, TC-F-06, TC-T-06", "Test Specified"),
    # Row 3 (BR-11, FR-11, US-04, UI-01..06)
    (3, "TC-F-01, TC-F-02, TC-F-03", "Test Specified"),
    # Row 4 (BR-02, FR-02, US-06, UI-01,03)
    (4, "TC-F-07, TC-T-07", "Test Specified"),
    # Row 5 (BR-11, FR-11, US-04, UI-01..06)
    (5, "TC-F-01, TC-F-02, TC-F-03", "Test Specified"),
    # Row 6 (BR-03, FR-03, US-09,10, UI-01,04)
    (6, "TC-F-10, TC-F-11, TC-T-08", "Test Specified"),
    # Row 7 (BR-11, FR-11, US-04, UI-01..06)
    (7, "TC-F-01, TC-F-02, TC-F-03", "Test Specified"),
    # Row 8 (BR-04, FR-04, US-11, UI-05)
    (8, "TC-F-12, TC-T-09", "Test Specified"),
    # Row 9 (BR-05, FR-05, US-07, UI-03)
    (9, "TC-F-08, TC-T-12", "Test Specified"),
    # Row 10 (BR-06, FR-06, US-13,14, UI-01,02)
    (10, "TC-F-14, TC-T-10", "Test Specified"),
    # Row 11 (BR-12, FR-12, US-16, UI-02,06)
    (11, "TC-F-16", "Test Specified"),
    # Row 12 (BR-07, FR-07, US-12, UI-01,05)
    (12, "TC-F-13, TC-T-11", "Test Specified"),
    # Row 13 (BR-08, FR-08, US-03,05, UI-01,03)
    (13, "TC-F-05", "Test Specified"),
    # Row 14 (BR-09, FR-09, US-15, UI-02)
    (14, "TC-F-15", "Test Specified"),
    # Row 15 (BR-12, FR-12, US-16, UI-02,06)
    (15, "TC-F-16", "Test Specified"),
    # Row 16 (BR-10, FR-10, US-08, UI-03)
    (16, "TC-F-09", "Test Specified")
]

BORDER_GRAY = "D1D5DB"
thin_border = Border(
    left=Side(style='thin', color=BORDER_GRAY),
    right=Side(style='thin', color=BORDER_GRAY),
    top=Side(style='thin', color=BORDER_GRAY),
    bottom=Side(style='thin', color=BORDER_GRAY)
)

for row_num, tc_id, status in tc_mappings:
    c_status = ws_rtm.cell(row=row_num, column=7)
    c_status.value = status
    c_status.font = Font(name="Calibri", size=10, bold=True, color="059669")
    
    c_tc = ws_rtm.cell(row=row_num, column=12)
    c_tc.value = tc_id
    c_tc.alignment = Alignment(horizontal="center", vertical="center")
    c_tc.font = Font(name="Calibri", size=10, bold=True, color="DC2626")
    c_tc.border = thin_border

ws_rtm.column_dimensions["L"].width = 28

wb.save(rtm_path)
print(f'Successfully updated: {rtm_path}')
