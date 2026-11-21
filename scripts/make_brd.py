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

# Document Header & Title
p_title = doc.add_paragraph()
r_title = p_title.add_run('BUSINESS REQUIREMENTS DOCUMENT (BRD)')
r_title.bold = True
r_title.font.size = Pt(22)
r_title.font.color.rgb = RGBColor(24, 43, 73)
p_title.paragraph_format.space_after = Pt(2)

p_sub = doc.add_paragraph()
r_sub = p_sub.add_run('Enterprise Supplier Performance Analysis & Procurement Intelligence System')
r_sub.font.size = Pt(13)
r_sub.font.color.rgb = RGBColor(90, 105, 120)
p_sub.paragraph_format.space_after = Pt(14)

# Section 1: Document Control
h1 = doc.add_heading('1. Document Control & Version History', level=1)
h1.paragraph_format.space_before = Pt(12)
h1.paragraph_format.space_after = Pt(6)

ctrl_table = doc.add_table(rows=6, cols=2)
ctrl_table.alignment = WD_TABLE_ALIGNMENT.CENTER
ctrl_data = [
    ('Document Title', 'Business Requirements Document (BRD) - Supplier Performance Analysis'),
    ('Document Version', 'Version 1.0 (Formal Baseline)'),
    ('Author / Lead Analyst', 'Chitra Chauhan | Roll No: 2500320140043 | MCA 2nd Year (Data Analytics), ABES'),
    ('Academic Supervisor / Mentor', 'Mr. Dinesh (Department of Computer Applications, ABES EC)'),
    ('Document Classification', 'Academic & Industry Applied Capstone Engineering Deliverable'),
    ('Creation & Baseline Date', 'Academic Session 2024-2026 | Document Date: Current Baseline')
]
for idx, (label, val) in enumerate(ctrl_data):
    r = ctrl_table.rows[idx]
    r.cells[0].text = label
    r.cells[1].text = val
    set_cell_background(r.cells[0], 'F0F4F8')
    r.cells[0].paragraphs[0].runs[0].bold = True
    r.cells[0].paragraphs[0].runs[0].font.size = Pt(9.5)
    r.cells[1].paragraphs[0].runs[0].font.size = Pt(9.5)
    set_cell_margins(r.cells[0], 70, 70, 100, 100)
    set_cell_margins(r.cells[1], 70, 70, 100, 100)

doc.add_paragraph().paragraph_format.space_after = Pt(8)

# Section 2: Executive Summary
h2 = doc.add_heading('2. Executive Summary', level=1)
h2.paragraph_format.space_before = Pt(12)
h2.paragraph_format.space_after = Pt(6)

p_exec = doc.add_paragraph(
    'In modern enterprise manufacturing and distribution operations, procurement performance directly dictates operational stability, manufacturing line throughput, and operating cash margins. Organizations spend millions of dollars annually procuring thousands of parts across hundreds of external vendors. However, management frequently lacks consolidated operational intelligence to systematically track vendor delivery punctuality, order completeness, receiving quality, and unit pricing adherence.\n\n'
    'This Business Requirements Document (BRD) defines the mandatory capabilities for an end-to-end Supplier Performance Analysis Platform. By integrating operational purchase orders, receiving logs, quality inspection records, and financial baseline contracts, the system automates the computation of On-Time In-Full (OTIF) rates, lead-time variance distributions, Defective Parts Per Million (PPM), and Purchase Price Variances (PPV). These disparate metrics are synthesized into a weighted, objective Composite Performance Index (CPI) that segments suppliers into Tier A (Preferred), Tier B (Conditional), and Tier C (At-Risk) governance tiers, providing leadership with immediate transparency into Spend-at-Risk exposure and automating contractual corrective action triggers.'
)
p_exec.paragraph_format.line_spacing = 1.15
p_exec.paragraph_format.space_after = Pt(8)

# Section 3: Business Problem & Opportunity
h3 = doc.add_heading('3. Business Problem & Opportunity Statement', level=1)
h3.paragraph_format.space_before = Pt(12)
h3.paragraph_format.space_after = Pt(6)

doc.add_paragraph(
    '3.1 The Business Problem:\n'
    'Procurement departments face significant operational risks due to fragmented, siloed reporting systems. Critical vendor performance data is scattered across legacy spreadsheets, warehouse receiving logs, and manual quality inspection clipboards. Consequently:\n'
    '• Chronic vendor delivery delays force unexpected assembly line stoppages.\n'
    '• Undetected short-shipments generate stockouts and administrative back-order chaos.\n'
    '• Defective component batches slip past receiving into active production, triggering expensive manufacturing scrap.\n'
    '• Invoiced prices silently drift above negotiated master contract rates without financial controller awareness.\n'
    '• Annual vendor contract renewals rely on subjective personal impressions rather than empirical performance scorecards.\n\n'
    '3.2 The Market Opportunity:\n'
    'Implementing a centralized, data-driven analytical platform creates a unified single source of truth across Procurement, Quality Assurance, Warehousing, and Finance. Automated scorecards empower buyers to allocate volume strategically, enable QA teams to enforce contractual quality standards through formal CAPA mechanisms, allow warehouse managers to optimize safety stock buffers, and assist controllers in recovering unauthorized pricing creep.'
)

# Section 4: Objectives & Success Metrics (KPIs)
h4 = doc.add_heading('4. Business Objectives & Key Performance Indicators (KPIs)', level=1)
h4.paragraph_format.space_before = Pt(12)
h4.paragraph_format.space_after = Pt(6)

kpi_table = doc.add_table(rows=7, cols=4)
kpi_table.alignment = WD_TABLE_ALIGNMENT.CENTER
kpi_headers = ['Objective ID', 'Strategic Business Objective', 'Primary KPI & Target Benchmark', 'Organizational Impact']
for j, h in enumerate(kpi_headers):
    c = kpi_table.cell(0, j)
    c.text = h
    set_cell_background(c, '1A365D')
    p = c.paragraphs[0]
    p.runs[0].bold = True
    p.runs[0].font.color.rgb = RGBColor(255, 255, 255)
    p.runs[0].font.size = Pt(9)
    set_cell_margins(c, 100, 100, 100, 100)

kpi_rows = [
    ('OBJ-01', 'Fulfillment Reliability & OTIF Optimization', 'Average OTIF >= 95% for Tier A vendors; Zero hidden short-shipments.', 'Eliminates unexpected manufacturing assembly halts and stockout delays.'),
    ('OBJ-02', 'Lead Time Predictability & Safety Buffer Calibration', 'Lead time variance reduction; Mean lead time tracked within +/- 2 days of SLA.', 'Prevents working capital bloat caused by excessive warehouse safety stock.'),
    ('OBJ-03', 'Component Quality Compliance & Scrap Reduction', 'Incoming Defect PPM < 1,000 across critical categories; Scrap rate < 1.5%.', 'Drastically reduces factory floor scrap, rework costs, and warranty liabilities.'),
    ('OBJ-04', 'Purchase Price Variance (PPV) & Cost Leakage Audit', '100% automated detection of price creep; Invoiced unit cost variance <= 0.0%.', 'Protects procurement gross profit margins and enables invoice recovery.'),
    ('OBJ-05', 'Spend-at-Risk Mitigation & Vendor Governance', 'Spend allocated to Tier C (At-Risk) suppliers kept below 10% of total spend.', 'Protects enterprise operational continuity from single-point vendor failure.'),
    ('OBJ-06', 'Computational & Analytical Efficiency', 'Analytical scorecard query and KPI recalculation latency < 2.0 seconds.', 'Enables real-time operational decision-making during buyer negotiations.')
]

for i, row in enumerate(kpi_rows):
    for j, val in enumerate(row):
        c = kpi_table.cell(i+1, j)
        c.text = val
        if (i % 2) == 1:
            set_cell_background(c, 'F7FAFC')
        c.paragraphs[0].runs[0].font.size = Pt(8.5)
        set_cell_margins(c, 70, 70, 80, 80)

doc.add_paragraph().paragraph_format.space_after = Pt(8)

# Section 5: Stakeholder RACI Matrix
h5 = doc.add_heading('5. Stakeholder RACI Governance Matrix', level=1)
h5.paragraph_format.space_before = Pt(12)
h5.paragraph_format.space_after = Pt(6)

p_raci = doc.add_paragraph('Legend: R = Responsible (does work), A = Accountable (approver), C = Consulted (provides input), I = Informed (updated).')
p_raci.paragraph_format.space_after = Pt(4)

raci_table = doc.add_table(rows=8, cols=6)
raci_table.alignment = WD_TABLE_ALIGNMENT.CENTER
raci_headers = ['Operational Activity / Deliverable', 'SH-01 Head SCM', 'SH-02 Buyer', 'SH-03 QA Lead', 'SH-04 Warehouse', 'SH-05 Finance']
for j, h in enumerate(raci_headers):
    c = raci_table.cell(0, j)
    c.text = h
    set_cell_background(c, '1A365D')
    p = c.paragraphs[0]
    p.runs[0].bold = True
    p.runs[0].font.color.rgb = RGBColor(255, 255, 255)
    p.runs[0].font.size = Pt(8.5)
    set_cell_margins(c, 80, 80, 80, 80)

raci_data = [
    ('Vendor Lead Time & OTIF Benchmarking', 'Accountable', 'Responsible', 'Informed', 'Consulted', 'Informed'),
    ('Quality Defect Tracking & PPM Auditing', 'Informed', 'Informed', 'Accountable / Resp.', 'Consulted', 'Informed'),
    ('Purchase Price Variance (PPV) Audits', 'Informed', 'Consulted', 'Informed', 'Informed', 'Accountable / Resp.'),
    ('Safety Stock & Warehouse Calibration', 'Informed', 'Consulted', 'Informed', 'Accountable / Resp.', 'Informed'),
    ('Vendor Scorecard Review & Tiering (A/B/C)', 'Accountable', 'Responsible', 'Consulted', 'Consulted', 'Consulted'),
    ('Spend-at-Risk & Contract Renewal Decisions', 'Accountable / Resp.', 'Consulted', 'Consulted', 'Informed', 'Consulted'),
    ('Corrective Action Plan (CAPA) Enforcement', 'Accountable', 'Informed', 'Responsible', 'Informed', 'Informed')
]

for i, row in enumerate(raci_data):
    for j, val in enumerate(row):
        c = raci_table.cell(i+1, j)
        c.text = val
        if (i % 2) == 1:
            set_cell_background(c, 'F7FAFC')
        c.paragraphs[0].runs[0].font.size = Pt(8.5)
        if j > 0:
            c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_cell_margins(c, 70, 70, 70, 70)

doc.add_paragraph().paragraph_format.space_after = Pt(8)

# Section 6: Scope
h6 = doc.add_heading('6. Scope: In-Scope vs Out-of-Scope', level=1)
h6.paragraph_format.space_before = Pt(12)
h6.paragraph_format.space_after = Pt(6)

doc.add_paragraph(
    '6.1 In-Scope Capabilities:\n'
    '• Generation of high-fidelity synthetic procurement relational database (3,000+ orders, 6 normalized entities).\n'
    '• Automated calculation of core KPIs: OTIF%, Lead Time (Mean, Std Dev, IQR), Defect PPM, Cost Variance%, and Composite Performance Index (CPI).\n'
    '• Interactive analytical dashboard featuring multi-dimensional filtering (Date Range, Supplier, Category, Status, Risk Tier).\n'
    '• Visual exploratory data analysis: OTIF compliance distributions, defect Pareto charts, lead-time boxplots, and Spend vs Risk scatter matrices.\n'
    '• Automated governance flags identifying vendors eligible for renewal or requiring Corrective Action Plans (CAPA).\n'
    '• Automated unit and integration testing via pytest achieving >= 90% statement coverage.\n'
    '• Public deployment on Streamlit Community Cloud with persistent local SQLite storage.\n\n'
    '6.2 Out-of-Scope Capabilities:\n'
    '• Live bidirectional API integration with proprietary ERP platforms (SAP S/4HANA, Oracle Cloud).\n'
    '• Automated electronic banking / wire transfer execution for vendor invoices.\n'
    '• Real-time IoT telematics and satellite GPS container tracking in maritime/road transit.\n'
    '• Deep neural network machine learning models for demand forecasting (focus is on descriptive and diagnostic analytics).\n'
    '• Native mobile smartphone application development (iOS / Android).'
)

# Section 7: Detailed Business Requirements (BR-01 to BR-12)
h7 = doc.add_heading('7. Detailed Business Requirements Catalog', level=1)
h7.paragraph_format.space_before = Pt(12)
h7.paragraph_format.space_after = Pt(6)

br_details = [
    ('BR-01: On-Time In-Full (OTIF) Delivery Tracking',
     'The business needs to track and benchmark vendor On-Time In-Full (OTIF) fulfillment compliance rates across all purchase orders so that procurement teams can allocate orders exclusively to dependable suppliers and eliminate production halts.',
     'Late and partial deliveries disrupt master manufacturing production schedules and force expensive expediting costs.',
     'Must-Have', 'HMW-01',
     'Given a closed purchase order, when evaluating fulfillment, then OTIF = 1 if and only if actual delivery date <= promised date AND delivered quantity >= ordered quantity; else OTIF = 0. Dashboard must report overall and supplier-specific OTIF percentages.'),

    ('BR-02: Lead Time Variability & SLA Adherence',
     'The business needs to compute average lead times, standard deviations, and SLA variance per supplier and product category so that warehouse operations can calibrate safety stock levels without inflating inventory holding costs.',
     'High lead-time variance causes inventory stockouts or forces excessive safety stock holding bloat.',
     'Must-Have', 'HMW-02',
     'Given delivery records, system calculates elapsed lead time = actual delivery date - PO creation date. Computes mean, standard deviation, and delta against contracted SLA days for 100% of orders.'),

    ('BR-03: Component Defect Rate & Defect PPM Monitoring',
     'The business needs to monitor incoming component defect rates and Defective Parts Per Million (PPM) at the supplier and part level so that quality managers can prevent sub-standard materials from entering assembly processes.',
     'Defective components entering assembly cause severe production line scrap, costly rework, and potential safety recalls.',
     'Must-Have', 'HMW-03',
     'Given QA inspection records, system computes Defect Rate (%) = (Rejected Units / Inspected Units) * 100 and Defect PPM = Defect Rate * 10,000. Flags any supplier exceeding 1,000 PPM threshold.'),

    ('BR-04: Purchase Price Variance (PPV) & Cost Creep Audit',
     'The business needs to audit unit purchase price variances between contracted baselines and actual invoiced costs across all order line items so that finance controllers can prevent cost leakage and recover unauthorized price increments.',
     'Unapproved unit price increases erode procurement margins across thousands of purchase order line items.',
     'Must-Have', 'HMW-04',
     'Given order line items, system computes Cost Variance (%) = ((Invoiced Unit Price - Contract Unit Price) / Contract Unit Price) * 100. Any positive cost variance (> 0%) is flagged as unauthorized price creep.'),

    ('BR-05: Delivery Delay Outlier & Tail Risk Isolation',
     'The business needs to isolate supplier delivery outliers and tail delays exceeding seven business days so that sourcing teams can remove highly volatile vendors from critical path manufacturing schedules.',
     'Severe tail delays (> 7 days) inflict disproportionate damage on downstream customer commitments compared to minor 1-day delays.',
     'Should-Have', 'HMW-05',
     'System flags and filters all delivery events where actual delivery date - promised date > 7 calendar days. Provides distribution histograms and statistical outlier counts.'),

    ('BR-06: Composite Performance Index (CPI) Scorecard',
     'The business needs to synthesize delivery timeliness, quality compliance, and price adherence into a weighted Composite Performance Index (CPI) ranking so that management can objectively categorize suppliers into standardized governance tiers.',
     'Single-metric vendor evaluations create dangerous operational blindspots; an objective multi-metric benchmark is required.',
     'Must-Have', 'HMW-06',
     'CPI Score = (0.40 * OTIF%) + (0.35 * (100 - Defect Rate%)) + (0.25 * (100 - Cost Variance%)). Assigns Tier A (>= 85), Tier B (70-84.9), and Tier C (< 70).'),

    ('BR-07: Spend-at-Risk Exposure Analytics',
     'The business needs to quantify the total procurement expenditure allocated to high-risk and underperforming suppliers so that leadership can mitigate supply chain vulnerabilities and reallocate capital toward reliable vendors.',
     'High spend concentrated with low-performing vendors creates existential supply chain vulnerability.',
     'Must-Have', 'HMW-07',
     'Spend-at-Risk (%) = (Total Annual Spend with Tier C Suppliers / Total Overall Spend) * 100. Alerts when Spend-at-Risk exceeds 10% organizational risk limit.'),

    ('BR-08: Chronological & Quarterly Seasonal Trend Analysis',
     'The business needs to analyze chronological and quarterly trends in vendor fulfillment and quality metrics so that category managers can anticipate seasonal slumps and secure capacity ahead of peak operational cycles.',
     'Supplier capacity bottlenecks and quality degradation frequently follow seasonal and quarter-end cycles.',
     'Should-Have', 'HMW-08',
     'Aggregates monthly and quarterly time-series tracking of OTIF%, defect rates, and delivery delays across 4 rolling quarters with interactive trendlines.'),

    ('BR-09: Automated CAPA & Contract Renewal Governance',
     'The business needs to identify suppliers eligible for strategic contract renewal or requiring mandatory Corrective Action Plans (CAPA) based on objective threshold criteria so that vendor governance is enforced systematically and legally.',
     'Enforcing vendor accountability requires standardized contractual audit triggers rather than subjective discretion.',
     'Must-Have', 'HMW-09',
     'Suppliers with CPI >= 85 flagged for "Contract Renewal / Expansion"; suppliers with CPI < 70 or consecutive late deliveries >= 3 flagged for "CAPA Audit Mandatory".'),

    ('BR-10: Shipping Logistics Mode & Geographic Impact',
     'The business needs to evaluate vendor delivery lead times and transit delays across shipping modes and geographic vendor origins so that logistics planners can select the most predictable transport routes and origin locations.',
     'Inability to distinguish vendor manufacturing delays from freight transit mode delays obscures logistics optimization.',
     'Should-Have', 'HMW-10',
     'Calculates average lead time and delay days grouped across freight shipping modes (Air, Sea, Road, Rail) and supplier country of origin.'),

    ('BR-11: Multi-Dimensional Slicing & Filtering',
     'The business needs to filter and segment procurement performance by product category, date range, supplier, and order status so that cross-functional teams can conduct targeted root-cause investigations.',
     'Users require targeted root-cause analysis across specific timeframes, high-spend categories, and individual vendors.',
     'Must-Have', 'HMW-01, 02, 03',
     'Interactive multi-select filters for Category, Supplier, Date Range, and Status dynamically update 100% of downstream charts and KPI cards within 1.5 seconds.'),

    ('BR-12: Structured Scorecard & Report Export',
     'The business needs to export standardized supplier performance scorecards and audit summaries into structured reports so that procurement teams can present empirical data during formal vendor contract reviews.',
     'Procurement specialists require portable, structured evidence to share with suppliers during formal negotiation sessions.',
     'Could-Have', 'HMW-06, 09',
     'Provides an on-demand export capability generating formatted tabular data snapshots of filtered supplier performance metrics.')
]

for title, desc, rat, prio, hmw, crit in br_details:
    p_b = doc.add_paragraph()
    r = p_b.add_run(title)
    r.bold = True
    r.font.size = Pt(11)
    r.font.color.rgb = RGBColor(26, 54, 93)
    p_b.paragraph_format.space_before = Pt(6)
    p_b.paragraph_format.space_after = Pt(2)
    
    table_b = doc.add_table(rows=5, cols=2)
    table_b.alignment = WD_TABLE_ALIGNMENT.CENTER
    fields = [
        ('Requirement Description', desc),
        ('Business Rationale', rat),
        ('MoSCoW Priority & Source', f'{prio} | Derived from {hmw}'),
        ('Acceptance Criteria', crit),
        ('Traceability Status', 'Baseline Requirement | Mapped to Step 4 RTM')
    ]
    for idx, (lbl, val) in enumerate(fields):
        row = table_b.rows[idx]
        row.cells[0].text = lbl
        row.cells[1].text = val
        set_cell_background(row.cells[0], 'F0F4F8')
        row.cells[0].paragraphs[0].runs[0].bold = True
        row.cells[0].paragraphs[0].runs[0].font.size = Pt(8.5)
        row.cells[1].paragraphs[0].runs[0].font.size = Pt(8.5)
        set_cell_margins(row.cells[0], 60, 60, 80, 80)
        set_cell_margins(row.cells[1], 60, 60, 80, 80)
        
    doc.add_paragraph().paragraph_format.space_after = Pt(4)

# Section 8: Business Rules
h8 = doc.add_heading('8. Business Rules Catalog', level=1)
h8.paragraph_format.space_before = Pt(12)
h8.paragraph_format.space_after = Pt(6)

rules = [
    ('B-RULE-01 (OTIF Delivery Condition):', ' An order line is strictly classified as On-Time In-Full (OTIF = 1) if and only if actual delivery date <= promised delivery date AND delivered quantity >= ordered quantity. Partial deliveries arriving on time receive OTIF = 0.'),
    ('B-RULE-02 (Lead Time Computation):', ' Delivery lead time is computed as elapsed calendar days: Lead_Time = actual_delivery_date - po_creation_date. If delivery is pending/in-transit, Lead_Time = current_date - po_creation_date.'),
    ('B-RULE-03 (Defect PPM Calculation):', ' Defective Parts Per Million (PPM) is calculated as: Defect_PPM = (rejected_units / inspected_units) * 1,000,000. If inspected units = 0, defect metrics are excluded from calculation.'),
    ('B-RULE-04 (Price Variance & PPV):', ' Cost Variance (%) = ((invoiced_unit_price - contracted_unit_price) / contracted_unit_price) * 100. Any positive variance (> 0%) triggers a price discrepancy warning.'),
    ('B-RULE-05 (Composite Performance Weighting):', ' Composite Performance Index (CPI) is computed on a 0-100 scale: CPI = 0.40 * OTIF% + 0.35 * (100 - Defect Rate%) + 0.25 * (100 - Cost Variance%). Total weights must strictly sum to 1.0 (100%).'),
    ('B-RULE-06 (Vendor Governance Tiering):', ' Suppliers are classified based on composite scores: Tier A (Preferred / Renewal) for CPI >= 85; Tier B (Conditional / Monitoring) for 70 <= CPI < 85; Tier C (At-Risk / CAPA Mandatory) for CPI < 70 OR consecutive late orders >= 3 OR Defect PPM > 5,000.')
]
for r_id, r_txt in rules:
    p = doc.add_paragraph(style='List Bullet')
    r1 = p.add_run(r_id)
    r1.bold = True
    r2 = p.add_run(r_txt)
    p.paragraph_format.space_after = Pt(3)

# Section 9: Data Requirements & Ingestion
h9 = doc.add_heading('9. Data Requirements & Ingestion Lifecycle', level=1)
h9.paragraph_format.space_before = Pt(12)
h9.paragraph_format.space_after = Pt(6)

doc.add_paragraph(
    'The platform operates on a normalized relational SQLite database refreshed via automated batch ingestion script. The schema comprises six primary entities:\n'
    '1. Suppliers: Master vendor registry (Supplier_ID, Name, Country, Rating, Tier, Contact).\n'
    '2. Categories: Procurement commodity taxonomy (Category_ID, Category_Name, Risk_Level).\n'
    '3. Products: Master part catalog (Product_ID, Product_Name, Category_ID, Baseline_Contract_Price, SLA_Lead_Time_Days).\n'
    '4. Purchase_Orders: Order transactional headers (PO_ID, Supplier_ID, Order_Date, Promised_Delivery_Date, Status, Total_Amount).\n'
    '5. Deliveries: Fulfillment records (Delivery_ID, PO_ID, Actual_Delivery_Date, Delivered_Qty, Delay_Days, Shipping_Mode).\n'
    '6. Quality_Inspections: Receiving quality audits (Inspection_ID, Delivery_ID, Product_ID, Inspected_Qty, Defective_Qty, Defect_Reason_Code).\n\n'
    'Refresh Cadence: Ingestion executes during application startup; analytical cache persists in session memory for zero-latency filter recalculation.'
)

# Section 10: Assumptions, Constraints & Dependencies
h10 = doc.add_heading('10. Assumptions, Constraints & Dependencies', level=1)
h10.paragraph_format.space_before = Pt(12)
h10.paragraph_format.space_after = Pt(6)

doc.add_paragraph(
    '10.1 Project Assumptions:\n'
    '• ASSUMP-01: Synthetic data generated in Step 10 accurately simulates enterprise procurement distributions.\n'
    '• ASSUMP-02: Purchase order delivery and QA records are immutable once marked completed.\n'
    '• ASSUMP-03: Single-node persistent SQLite database is sufficient for the target dataset size (3,000 to 5,000 records).\n\n'
    '10.2 Technical Constraints:\n'
    '• CONST-01: Total project software and licensing budget is $0 (100% open-source stack).\n'
    '• CONST-02: Application memory consumption must stay within 1 GB RAM for Streamlit Community Cloud hosting.\n'
    '• CONST-03: System must strictly follow ABES MCA academic project standards and viva evaluation rubrics.\n\n'
    '10.3 System Dependencies:\n'
    '• DEP-01: Python 3.11 runtime environment.\n'
    '• DEP-02: Streamlit web application framework and Plotly visualization libraries.\n'
    '• DEP-03: GitHub version control repository linked to Streamlit Community Cloud deployment webhooks.'
)

# Section 11: Risks & Mitigation Strategy
h11 = doc.add_heading('11. Risk Management & Mitigation Matrix', level=1)
h11.paragraph_format.space_before = Pt(12)
h11.paragraph_format.space_after = Pt(6)

risk_table = doc.add_table(rows=6, cols=5)
risk_table.alignment = WD_TABLE_ALIGNMENT.CENTER
risk_headers = ['Risk ID', 'Risk Description', 'Likelihood', 'Impact', 'Mitigation & Contingency Strategy']
for j, h in enumerate(risk_headers):
    c = risk_table.cell(0, j)
    c.text = h
    set_cell_background(c, '1A365D')
    p = c.paragraphs[0]
    p.runs[0].bold = True
    p.runs[0].font.color.rgb = RGBColor(255, 255, 255)
    p.runs[0].font.size = Pt(8.5)
    set_cell_margins(c, 80, 80, 80, 80)

risk_data = [
    ('RISK-01', 'Synthetic data lacks realistic statistical anomalies or business edge cases.', 'Medium', 'High', 'Utilize calibrated parametric distributions (log-normal spend, Weibull lead times, Pareto defects) with dedicated unit test validation.'),
    ('RISK-02', 'Complex multi-filter analytical recalculations exceed latency SLA (> 2.0s).', 'Low', 'Medium', 'Implement SQLite indexing on foreign keys and apply Streamlit caching (@st.cache_data) for analytical aggregations.'),
    ('RISK-03', 'Streamlit Cloud container restarts clear active session parameters.', 'Medium', 'Low', 'Store all master transactions and generated records persistently in SQLite; use deterministic seed initialization.'),
    ('RISK-04', 'Disagreement on composite score KPI weightings among stakeholder personas.', 'High', 'Medium', 'Implement dynamic sidebar configuration sliders allowing users to customize KPI weights interactively during evaluation.'),
    ('RISK-05', 'Deployment pipeline breaks prior to academic viva examination.', 'Low', 'High', 'Maintain automated pytest test suite and secondary offline localhost execution package as immediate contingency.')
]

for i, row in enumerate(risk_data):
    for j, val in enumerate(row):
        c = risk_table.cell(i+1, j)
        c.text = val
        if (i % 2) == 1:
            set_cell_background(c, 'F7FAFC')
        c.paragraphs[0].runs[0].font.size = Pt(8.5)
        set_cell_margins(c, 60, 60, 70, 70)

doc.add_paragraph().paragraph_format.space_after = Pt(8)

# Section 12: Glossary
h12 = doc.add_heading('12. Project Glossary & Acronyms', level=1)
h12.paragraph_format.space_before = Pt(12)
h12.paragraph_format.space_after = Pt(6)

glossary_terms = [
    ('OTIF (On-Time In-Full):', ' A stringent supply chain delivery metric tracking the percentage of orders delivered on or before the committed date with 100% of the requested quantity.'),
    ('PPM (Parts Per Million):', ' A high-precision defect measurement standard calculated as (Defective Units / Total Inspected Units) * 1,000,000.'),
    ('CAPA (Corrective and Preventive Action):', ' A formal quality management process requiring suppliers to investigate root causes of systemic failures and implement corrective processes.'),
    ('PPV (Purchase Price Variance):', ' The percentage and currency difference between the contracted standard price and the actual invoiced purchase price.'),
    ('CPI (Composite Performance Index):', ' A weighted 0-100 score synthesizing delivery timeliness, quality adherence, and price variance into a single benchmark.'),
    ('SLA (Service Level Agreement):', ' Contractually committed delivery lead times and quality defect tolerances agreed upon between buyer and vendor.'),
    ('RTV (Return-to-Vendor):', ' Defective materials rejected at receiving inspection and returned to the vendor for credit or replacement.'),
    ('RTM (Requirements Traceability Matrix):', ' A grid tracking requirements across their lifecycle from business need down to automated test cases.')
]
for term, definition in glossary_terms:
    p = doc.add_paragraph(style='List Bullet')
    r1 = p.add_run(term)
    r1.bold = True
    r2 = p.add_run(definition)
    p.paragraph_format.space_after = Pt(3)

# Section 13: Formal Sign-off
h13 = doc.add_heading('13. Formal Project Sign-off & Approvals', level=1)
h13.paragraph_format.space_before = Pt(12)
h13.paragraph_format.space_after = Pt(6)

sign_table = doc.add_table(rows=4, cols=5)
sign_table.alignment = WD_TABLE_ALIGNMENT.CENTER
sign_headers = ['Role', 'Name', 'Designation / Department', 'Signature', 'Approval Date']
for j, h in enumerate(sign_headers):
    c = sign_table.cell(0, j)
    c.text = h
    set_cell_background(c, '1A365D')
    p = c.paragraphs[0]
    p.runs[0].bold = True
    p.runs[0].font.color.rgb = RGBColor(255, 255, 255)
    p.runs[0].font.size = Pt(8.5)
    set_cell_margins(c, 80, 80, 80, 80)

sign_data = [
    ('Lead Analyst / Author', 'Chitra Chauhan', 'MCA Student (Data Analytics), ABES EC', '[Signed Electronically]', 'Current Session'),
    ('Academic Mentor', 'Mr. Dinesh', 'Assistant Professor, Dept of CA, ABES EC', '[Approved for Baseline]', 'Current Session'),
    ('Department Reviewer', 'Head of Department / Examiner', 'Dept of Computer Applications, ABES EC', '[Pending Final Viva]', 'Scheduled Viva Date')
]
for i, row in enumerate(sign_data):
    for j, val in enumerate(row):
        c = sign_table.cell(i+1, j)
        c.text = val
        if (i % 2) == 1:
            set_cell_background(c, 'F7FAFC')
        c.paragraphs[0].runs[0].font.size = Pt(8.5)
        set_cell_margins(c, 70, 70, 70, 70)

# Save Document
output_path = os.path.join('docs', '05_BRD.docx')
doc.save(output_path)
print(f'Successfully generated: {output_path}')
