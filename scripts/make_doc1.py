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

# Helper function for cell background color
def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>'))

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

# Title & Header
p_title = doc.add_paragraph()
r_title = p_title.add_run('STEP 1: PROJECT UNDERSTANDING DOCUMENT')
r_title.bold = True
r_title.font.size = Pt(20)
r_title.font.color.rgb = RGBColor(24, 43, 73)
p_title.paragraph_format.space_after = Pt(2)

p_sub = doc.add_paragraph()
r_sub = p_sub.add_run('Comprehensive Project Scope, Stakeholder Matrix & Operational Requirements')
r_sub.font.size = Pt(12)
r_sub.font.color.rgb = RGBColor(90, 105, 120)
p_sub.paragraph_format.space_after = Pt(14)

# Project Metadata Table
meta_table = doc.add_table(rows=5, cols=2)
meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
meta_data = [
    ('Project Title', 'Supplier Performance Analysis & Procurement Intelligence System'),
    ('Student Details', 'Chitra Chauhan | Roll No: 2500320140043 | MCA 2nd Year (Data Analytics), ABES'),
    ('Project Mentor', 'Mr. Dinesh'),
    ('Technology Stack', 'Python 3.11, Streamlit, Plotly, SQLite, pytest, Streamlit Community Cloud'),
    ('Document Version & Date', 'Version 1.0 (Final) | Academic Year 2024-2026')
]
for idx, (label, val) in enumerate(meta_data):
    r = meta_table.rows[idx]
    r.cells[0].text = label
    r.cells[1].text = val
    set_cell_background(r.cells[0], 'F0F4F8')
    r.cells[0].paragraphs[0].runs[0].bold = True
    r.cells[0].paragraphs[0].runs[0].font.size = Pt(9.5)
    r.cells[1].paragraphs[0].runs[0].font.size = Pt(9.5)
    set_cell_margins(r.cells[0], 80, 80, 120, 120)
    set_cell_margins(r.cells[1], 80, 80, 120, 120)

doc.add_paragraph().paragraph_format.space_after = Pt(10)

# Section 1: Executive Summary & Context
h1 = doc.add_heading('1. Executive Summary & Context', level=1)
h1.paragraph_format.space_before = Pt(12)
h1.paragraph_format.space_after = Pt(6)
p = doc.add_paragraph(
    'In modern manufacturing and retail supply chains, procurement decisions heavily influence production continuity, product quality, and cost efficiency. Organizations engage hundreds of suppliers across multiple commodity categories. However, many procurement divisions continue to rely on siloed spreadsheets and fragmented operational logs, obscuring systemic supplier delays, quality defects, and pricing irregularities. This project establishes an interactive, data-driven analytical application that evaluates vendor reliability, quantifies supply chain risk, and provides actionable intelligence for supplier contract management.'
)
p.paragraph_format.line_spacing = 1.15
p.paragraph_format.space_after = Pt(8)

# Section 2: Problem Statement
h2 = doc.add_heading('2. Problem Statement', level=1)
h2.paragraph_format.space_before = Pt(12)
h2.paragraph_format.space_after = Pt(6)

p_prob = doc.add_paragraph()
r_prob_quote = p_prob.add_run('"Procurement managers and supply chain teams lack a unified, data-driven analytical platform to systematically track, evaluate, and benchmark supplier performance across delivery timelines, quality compliance, and cost variance, leading to delayed production cycles, undetected defect spikes, and suboptimal vendor negotiations."')
r_prob_quote.bold = True
r_prob_quote.italic = True
r_prob_quote.font.size = Pt(11)
r_prob_quote.font.color.rgb = RGBColor(160, 32, 32)
p_prob.paragraph_format.space_after = Pt(10)

# Section 3: Intended Audience & User Personas
h3 = doc.add_heading('3. Intended Audience & User Personas', level=1)
h3.paragraph_format.space_before = Pt(12)
h3.paragraph_format.space_after = Pt(6)

p_aud_intro = doc.add_paragraph('The application is designed for operational and strategic personas across supply chain management:')
aud_table = doc.add_table(rows=5, cols=4)
aud_table.alignment = WD_TABLE_ALIGNMENT.CENTER
aud_headers = ['Target Persona', 'Access Frequency', 'Primary Role', 'Core Workflow in App']
for j, h in enumerate(aud_headers):
    c = aud_table.cell(0, j)
    c.text = h
    set_cell_background(c, '1A365D')
    p = c.paragraphs[0]
    p.runs[0].bold = True
    p.runs[0].font.color.rgb = RGBColor(255, 255, 255)
    p.runs[0].font.size = Pt(9.5)
    set_cell_margins(c, 100, 100, 120, 120)

aud_rows = [
    ('Procurement Specialist', 'Daily', 'Operational PO tracking & order allocation', 'Monitor open POs, check real-time vendor lead times, and identify pending deliveries at risk of delay.'),
    ('Supply Chain Category Manager', 'Weekly', 'Vendor performance reviews & quarterly audits', 'Evaluate vendor scorecards, analyze defect rate distributions, and benchmark suppliers within specific categories.'),
    ('Quality Assurance Lead', 'Bi-Weekly', 'Quality control inspection & rejection analysis', 'Trace defective parts, inspect vendor rejection trends, and prepare corrective action requests (CAR).'),
    ('Head of Procurement / VP SCM', 'Monthly / Quarterly', 'Strategic sourcing & contract renewals', 'Review executive KPI summaries, vendor tier rankings (A/B/C), and total spend concentration versus supplier risk.')
]
for i, row in enumerate(aud_rows):
    for j, val in enumerate(row):
        c = aud_table.cell(i+1, j)
        c.text = val
        if (i % 2) == 1:
            set_cell_background(c, 'F7FAFC')
        c.paragraphs[0].runs[0].font.size = Pt(9)
        set_cell_margins(c, 80, 80, 100, 100)

doc.add_paragraph().paragraph_format.space_after = Pt(8)

# Section 4: Stakeholder Matrix
h4 = doc.add_heading('4. Stakeholder Matrix', level=1)
h4.paragraph_format.space_before = Pt(12)
h4.paragraph_format.space_after = Pt(6)

sh_table = doc.add_table(rows=6, cols=6)
sh_table.alignment = WD_TABLE_ALIGNMENT.CENTER
sh_headers = ['SH ID', 'Stakeholder', 'Role in Organization', 'Interest', 'Influence', 'Decisions Made Using Application']
for j, h in enumerate(sh_headers):
    c = sh_table.cell(0, j)
    c.text = h
    set_cell_background(c, '1A365D')
    p = c.paragraphs[0]
    p.runs[0].bold = True
    p.runs[0].font.color.rgb = RGBColor(255, 255, 255)
    p.runs[0].font.size = Pt(9)
    set_cell_margins(c, 100, 100, 100, 100)

sh_data = [
    ('SH-01', 'Head of Procurement / Director', 'Project Business Owner', 'High', 'High', 'Approves strategic supplier contracts, annual vendor consolidation, tier re-allocations, and supplier delisting.'),
    ('SH-02', 'Procurement & Sourcing Specialist', 'Primary Operational End-User', 'High', 'Medium', 'Allocates new purchase orders to reliable suppliers based on historical lead times and OTIF compliance.'),
    ('SH-03', 'Quality Assurance (QA) Manager', 'Quality & Compliance Lead', 'High', 'Medium', 'Issues vendor warning notices, triggers mandatory on-site audits, and enforces return-to-vendor (RTV) penalties.'),
    ('SH-04', 'Warehouse & Inventory Supervisor', 'Logistics Contributor', 'Medium', 'Low', 'Adjusts buffer stock levels and safety stock configurations in response to identified supplier lead-time variances.'),
    ('SH-05', 'Finance & Procurement Controller', 'Commercial & Spend Auditor', 'Medium', 'High', 'Enforces contracted pricing baselines, audits unit cost variances, and reconciles disputed invoice figures.')
]

for i, row in enumerate(sh_data):
    for j, val in enumerate(row):
        c = sh_table.cell(i+1, j)
        c.text = val
        if (i % 2) == 1:
            set_cell_background(c, 'F7FAFC')
        c.paragraphs[0].runs[0].font.size = Pt(8.5)
        set_cell_margins(c, 70, 70, 80, 80)

doc.add_paragraph().paragraph_format.space_after = Pt(8)

# Section 5: Operational Questions (OQ-01 to OQ-10)
h5 = doc.add_heading('5. Operational Questions & Data Validation Matrix', level=1)
h5.paragraph_format.space_before = Pt(12)
h5.paragraph_format.space_after = Pt(6)

p_oq = doc.add_paragraph('Each operational question (OQ) is directly answerable from the relational dataset schema:')
oq_table = doc.add_table(rows=11, cols=4)
oq_table.alignment = WD_TABLE_ALIGNMENT.CENTER
oq_headers = ['OQ ID', 'Operational Question', 'Required Analytical Metrics', 'Underlying Relational Data Attributes']
for j, h in enumerate(oq_headers):
    c = oq_table.cell(0, j)
    c.text = h
    set_cell_background(c, '1A365D')
    p = c.paragraphs[0]
    p.runs[0].bold = True
    p.runs[0].font.color.rgb = RGBColor(255, 255, 255)
    p.runs[0].font.size = Pt(9)
    set_cell_margins(c, 100, 100, 100, 100)

oq_data = [
    ('OQ-01', 'Which suppliers have the highest and lowest On-Time In-Full (OTIF) delivery compliance rates?', 'OTIF Percentage = (Total On-Time & In-Full Deliveries / Total Orders) * 100', 'promised_date, actual_delivery_date, ordered_quantity, delivered_quantity'),
    ('OQ-02', 'What is the average lead time and lead time variability (Std Dev / IQR) per vendor and category?', 'Mean Lead Time (Days), Standard Deviation, Delivery Delay Variance', 'po_creation_date, actual_delivery_date, category_id, supplier_id'),
    ('OQ-03', 'What is the defect rate percentage and Defective PPM across suppliers and component types?', 'Defect Rate (%) = (Rejected Units / Inspected Units) * 100, Defect PPM', 'inspected_quantity, rejected_quantity, defect_reason_code, part_id'),
    ('OQ-04', 'What is the cost variance between contracted baseline unit prices and actual invoiced prices?', 'Cost Variance (%) = ((Invoiced Unit Price - Contract Unit Price) / Contract Price) * 100', 'contract_unit_price, actual_invoiced_price, total_po_value'),
    ('OQ-05', 'Which suppliers exhibit severe delivery volatility posing critical supply chain bottleneck risks?', 'Lead Time Kurtosis/Skewness, Delivery Delay Outlier Counts (Delay > 7 days)', 'delay_days, delivery_status, historical_po_records'),
    ('OQ-06', 'How are suppliers ranked on a multi-criteria Composite Performance Index (CPI)?', 'Weighted Score = 40% OTIF + 35% Quality Score + 25% Cost Adherence Score', 'computed_otif_pct, computed_defect_pct, computed_cost_variance'),
    ('OQ-07', 'What proportion of procurement capital expenditure is concentrated among high-risk or low-tier vendors?', 'Spend-at-Risk (%) = (Total Spend with Tier C Suppliers / Total Overall Spend) * 100', 'po_line_amount, vendor_composite_tier (A/B/C)'),
    ('OQ-08', 'How do vendor performance metrics trend across quarterly and monthly seasonal cycles?', 'Monthly / Quarterly Time-series aggregations of OTIF%, Defect%, and Delay Days', 'po_date, delivery_month, delivery_quarter'),
    ('OQ-09', 'Which suppliers meet criteria for contract renewal or require formal Corrective Action Plans (CAPA)?', 'Action Classification (Tier A: Retain/Promote, Tier B: Monitor, Tier C: CAPA/Audit)', 'composite_score, consecutive_low_otif_flag, defect_spike_flag'),
    ('OQ-10', 'How do geographic supplier locations and logistics modes impact lead times and delivery delays?', 'Average Lead Time and Delay Days grouped by Logistics Mode and Supplier Country', 'supplier_country, shipping_mode (Air/Sea/Road/Rail), delay_days')
]

for i, row in enumerate(oq_data):
    for j, val in enumerate(row):
        c = oq_table.cell(i+1, j)
        c.text = val
        if (i % 2) == 1:
            set_cell_background(c, 'F7FAFC')
        c.paragraphs[0].runs[0].font.size = Pt(8.5)
        set_cell_margins(c, 70, 70, 80, 80)

doc.add_paragraph().paragraph_format.space_after = Pt(8)

# Section 6: SMART Objectives
h6 = doc.add_heading('6. SMART Objectives', level=1)
h6.paragraph_format.space_before = Pt(12)
h6.paragraph_format.space_after = Pt(6)

obj_bullets = [
    ('OBJ-01 (Specific & Measurable):', ' Develop an automated analytical processing pipeline that calculates OTIF, Defect Rate, Cost Variance, and Composite Performance Scores for 100% of historical purchase orders in < 2.0 seconds.'),
    ('OBJ-02 (Actionable & Relevant):', ' Build an interactive, multi-dimensional supplier scorecard ranking 100% of active suppliers into clear governance tiers (Tier A: Preferred, Tier B: Conditional, Tier C: At-Risk) with automated CAPA triggers.'),
    ('OBJ-03 (Attainable & Realistic):', ' Design and populate a 6-table normalized relational SQLite database containing 3,000+ realistic synthetic procurement transactions exhibiting authentic business distributions, seasonal variances, and realistic anomaly patterns.'),
    ('OBJ-04 (Quality & Reliability):', ' Achieve a minimum of 90% code coverage across analytical KPI computation modules validated through an automated pytest test suite covering both functional and edge-case scenarios.'),
    ('OBJ-05 (Time-bound & Deliverable):', ' Successfully deploy the containerized/cloud-ready Streamlit application to Streamlit Community Cloud and complete all ABES MCA academic report requirements within the semester schedule.')
]
for title, text in obj_bullets:
    p = doc.add_paragraph(style='List Bullet')
    r1 = p.add_run(title)
    r1.bold = True
    r2 = p.add_run(text)
    p.paragraph_format.space_after = Pt(4)

# Section 7: Scope Definition
h7 = doc.add_heading('7. Project Scope Definition', level=1)
h7.paragraph_format.space_before = Pt(12)
h7.paragraph_format.space_after = Pt(6)

scope_table = doc.add_table(rows=7, cols=2)
scope_table.alignment = WD_TABLE_ALIGNMENT.CENTER
scope_headers = ['In-Scope (System Capabilities)', 'Out-of-Scope (Explicit Boundaries)']
for j, h in enumerate(scope_headers):
    c = scope_table.cell(0, j)
    c.text = h
    set_cell_background(c, '1A365D')
    p = c.paragraphs[0]
    p.runs[0].bold = True
    p.runs[0].font.color.rgb = RGBColor(255, 255, 255)
    p.runs[0].font.size = Pt(9.5)
    set_cell_margins(c, 100, 100, 120, 120)

scope_data = [
    ('Synthetic enterprise procurement dataset generation (3,000+ records) with realistic distributions.', 'Direct real-time API integrations with commercial ERP systems (SAP S/4HANA, Oracle SCM).'),
    ('Relational SQLite database design with full referential integrity and optimized indexes.', 'Live multi-tenant enterprise cloud database hosting requiring paid enterprise infrastructure.'),
    ('Automated KPI calculation: OTIF%, Lead Time variability, Defect PPM, Cost Variance, and Composite Score.', 'Automated banking and payment gateway execution for vendor invoice disbursements.'),
    ('Interactive Streamlit analytics dashboard featuring dynamic Plotly visualizations and Tailwind styling.', 'Real-time GPS/IoT physical shipment tracking of containers in maritime or road transit.'),
    ('Multi-dimensional filtering (Date Range, Supplier, Item Category, Delivery Status, Performance Tier).', 'Machine learning deep neural network predictive forecasting (project scope centers on statistical and diagnostic analytics).'),
    ('Automated regression and unit test suite executed via pytest ensuring computational precision.', 'Mobile native application development for iOS and Android platforms.')
]
for i, row in enumerate(scope_data):
    for j, val in enumerate(row):
        c = scope_table.cell(i+1, j)
        c.text = val
        if (i % 2) == 1:
            set_cell_background(c, 'F7FAFC')
        c.paragraphs[0].runs[0].font.size = Pt(8.5)
        set_cell_margins(c, 70, 70, 80, 80)

doc.add_paragraph().paragraph_format.space_after = Pt(8)

# Section 8: Assumptions & Constraints
h8 = doc.add_heading('8. Assumptions & Constraints', level=1)
h8.paragraph_format.space_before = Pt(12)
h8.paragraph_format.space_after = Pt(6)

p_assump = doc.add_paragraph()
p_assump.add_run('Project Assumptions:').bold = True
assumptions = [
    ('ASSUMP-01: Synthetic Data Fidelity - ', 'The synthetic dataset generated in Step 10 accurately simulates real-world supply chain dynamics, including lead-time seasonality, supplier tier variations, and realistic defect distributions without containing proprietary corporate data.'),
    ('ASSUMP-02: Deterministic Business Rules - ', 'Lead times and defect calculations are computed based on confirmed delivery and QA inspection records; pending or in-transit orders are evaluated strictly against active status criteria.'),
    ('ASSUMP-03: Local Persistent Storage - ', 'SQLite is sufficient for single-node concurrent analytical querying without requiring network-attached distributed databases.'),
    ('ASSUMP-04: User Authorization Scope - ', 'The application operates in a unified analytical dashboard context tailored for procurement analysts and academic review without multi-tenant RBAC token servers.')
]
for a_id, a_text in assumptions:
    p = doc.add_paragraph(style='List Bullet')
    r1 = p.add_run(a_id)
    r1.bold = True
    r2 = p.add_run(a_text)
    p.paragraph_format.space_after = Pt(3)

p_const = doc.add_paragraph()
p_const.add_run('Project Constraints:').bold = True
constraints = [
    ('CONST-01: Zero Budget Constraint - ', 'The entire solution is built utilizing 100% free and open-source tools (Python, SQLite, Streamlit Community Cloud) with zero commercial licensing expenses.'),
    ('CONST-02: Computational & Memory Limit - ', 'Application execution is optimized to operate smoothly within Streamlit Community Cloud resource limits (1 GB RAM allocation).'),
    ('CONST-03: Academic Timeline & Milestones - ', 'Design and implementation must strictly follow the 18-step project lifecycle and adhere to ABES MCA academic project standards.'),
    ('CONST-04: Cross-Platform Compatibility - ', 'The Streamlit web interface must render seamlessly across modern web browsers (Google Chrome, Mozilla Firefox, Microsoft Edge) at 1080p and laptop resolutions.')
]
for c_id, c_text in constraints:
    p = doc.add_paragraph(style='List Bullet')
    r1 = p.add_run(c_id)
    r1.bold = True
    r2 = p.add_run(c_text)
    p.paragraph_format.space_after = Pt(3)

# Section 9: Measurable Success Criteria
h9 = doc.add_heading('9. Measurable Success Criteria', level=1)
h9.paragraph_format.space_before = Pt(12)
h9.paragraph_format.space_after = Pt(6)

sc_table = doc.add_table(rows=6, cols=4)
sc_table.alignment = WD_TABLE_ALIGNMENT.CENTER
sc_headers = ['Criteria ID', 'Success Dimension', 'Target Metric / Standard', 'Verification Method']
for j, h in enumerate(sc_headers):
    c = sc_table.cell(0, j)
    c.text = h
    set_cell_background(c, '1A365D')
    p = c.paragraphs[0]
    p.runs[0].bold = True
    p.runs[0].font.color.rgb = RGBColor(255, 255, 255)
    p.runs[0].font.size = Pt(9)
    set_cell_margins(c, 100, 100, 100, 100)

sc_data = [
    ('SC-01', 'OQ Coverage', '100% (10 out of 10 Operational Questions fully answered with interactive charts and KPI metrics).', 'System walkthrough against OQ-01 through OQ-10 verification checklist.'),
    ('SC-02', 'Query & Render Performance', 'Dashboard cold start < 3.0s; interactive filter recalculation latency < 1.5s for 3,000+ records.', 'Execution profiling using Python time benchmarking.'),
    ('SC-03', 'Automated Test Pass Rate', '100% test pass rate across all unit and integration tests with >= 90% statement coverage.', 'Automated execution of pytest with coverage reporting.'),
    ('SC-04', 'Cloud Deployment Reliability', '100% availability on Streamlit Community Cloud with zero unhandled exceptions.', 'Live URL access verification and smoke testing.'),
    ('SC-05', 'Academic Evaluation Standards', 'Formal acceptance of the project document, source repository, and slide deck by academic mentor.', 'Evaluation review by Mr. Dinesh and ABES MCA examination panel.')
]

for i, row in enumerate(sc_data):
    for j, val in enumerate(row):
        c = sc_table.cell(i+1, j)
        c.text = val
        if (i % 2) == 1:
            set_cell_background(c, 'F7FAFC')
        c.paragraphs[0].runs[0].font.size = Pt(8.5)
        set_cell_margins(c, 70, 70, 80, 80)

doc.add_paragraph().paragraph_format.space_after = Pt(14)

# Section 10: OQ Data Feasibility Verification
h10 = doc.add_heading('10. Data Feasibility Verification Audit', level=1)
h10.paragraph_format.space_before = Pt(12)
h10.paragraph_format.space_after = Pt(6)

p_audit = doc.add_paragraph(
    'An audit of all 10 operational questions was performed against the planned 6-table database schema (Suppliers, Categories, Products, Purchase_Orders, Deliveries, Quality_Inspections). Result: 100% of the required metrics (OTIF%, lead time days, defect rate, cost variance, composite index, spend concentration, risk tiering) can be mathematically computed and joined via primary/foreign key relationships without missing data dependencies.'
)
p_audit.paragraph_format.space_after = Pt(12)

# Save document
output_path1 = os.path.join('docs', '01_Project_Understanding.docx')
output_path2 = os.path.join('docs', '01_name_project_understa.docx')
doc.save(output_path1)
doc.save(output_path2)
print(f'Successfully generated: {output_path1} and {output_path2}')
