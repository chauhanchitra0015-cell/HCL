import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

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
r_title = p_title.add_run('STEP 3: BUSINESS NEEDS ANALYSIS DOCUMENT')
r_title.bold = True
r_title.font.size = Pt(20)
r_title.font.color.rgb = RGBColor(24, 43, 73)
p_title.paragraph_format.space_after = Pt(2)

p_sub = doc.add_paragraph()
r_sub = p_sub.add_run('Business Requirements Formulation, MoSCoW Prioritization & Gap Analysis')
r_sub.font.size = Pt(12)
r_sub.font.color.rgb = RGBColor(90, 105, 120)
p_sub.paragraph_format.space_after = Pt(14)

# Project Metadata Table
meta_table = doc.add_table(rows=5, cols=2)
meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
meta_data = [
    ('Project Title', 'Supplier Performance Analysis & Procurement Intelligence System'),
    ('Student Details', 'Chitra Chauhan | Roll No: 2500320140043 | MCA 2nd Year (Data Analytics), ABES'),
    ('Project Mentor', 'Prof. Kalpana Dwivedi'),
    ('Engineering Phase', 'Step 3: Needs Analysis (BR Transformation)'),
    ('Document Reference', 'docs/03_Needs_Analysis.docx | Derived from Step 2 HMW Statements')
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

# Section 1: Introduction
h1 = doc.add_heading('1. Executive Context & Methodology', level=1)
h1.paragraph_format.space_before = Pt(12)
h1.paragraph_format.space_after = Pt(6)

p_intro = doc.add_paragraph(
    'Step 3 bridges user-centered design thinking and formal technical requirements engineering. In this phase, each \"How Might We\" (HMW) statement developed in Step 2 is translated into a rigorous Business Requirement (BR). Following standard enterprise Business Analysis principles (IIBA BABOK), each requirement is articulated strictly from a business perspective: WHAT capabilities the organization must possess and WHAT tangible value is realized, purposefully avoiding technology-specific implementation choices (such as SQL syntax, Streamlit components, or Python libraries). Every requirement is prioritized using the MoSCoW framework and assigned measurable verification criteria.'
)
p_intro.paragraph_format.line_spacing = 1.15
p_intro.paragraph_format.space_after = Pt(8)

# Section 2: Business Requirements Table
h2 = doc.add_heading('2. Master Business Requirements Matrix (BR-01 to BR-12)', level=1)
h2.paragraph_format.space_before = Pt(12)
h2.paragraph_format.space_after = Pt(6)

br_table = doc.add_table(rows=13, cols=6)
br_table.alignment = WD_TABLE_ALIGNMENT.CENTER
br_headers = ['BR ID', 'Business Requirement Statement', 'Primary Stakeholder', 'Source HMW', 'Priority', 'Measurement / Acceptance Criteria']
for j, h in enumerate(br_headers):
    c = br_table.cell(0, j)
    c.text = h
    set_cell_background(c, '1A365D')
    p = c.paragraphs[0]
    p.runs[0].bold = True
    p.runs[0].font.color.rgb = RGBColor(255, 255, 255)
    p.runs[0].font.size = Pt(9)
    set_cell_margins(c, 100, 100, 100, 100)

br_data = [
    ('BR-01', 'The business needs to track and benchmark vendor On-Time In-Full (OTIF) fulfillment compliance rates across all purchase orders so that procurement teams can allocate orders exclusively to dependable suppliers and eliminate production halts.', 'SH-02 (Buyer), SH-01 (Head SCM)', 'HMW-01', 'Must', 'Percentage calculation: (On-Time & Full Deliveries / Total Completed Orders) * 100. Target benchmark >= 95% for Tier A vendors.'),
    ('BR-02', 'The business needs to compute average lead times, standard deviations, and SLA variance per supplier and product category so that warehouse operations can calibrate safety stock levels without inflating inventory holding costs.', 'SH-04 (Warehouse), SH-02 (Buyer)', 'HMW-02', 'Must', 'Mean lead time in elapsed calendar days and standard deviation per vendor; variance against agreed contracted SLA days.'),
    ('BR-03', 'The business needs to monitor incoming component defect rates and Defective Parts Per Million (PPM) at the supplier and part level so that quality managers can prevent sub-standard materials from entering assembly processes.', 'SH-03 (QA Manager)', 'HMW-03', 'Must', 'Ratio of rejected units to total inspected units converted to percentage and Defect PPM; vendor target benchmark < 1,000 PPM.'),
    ('BR-04', 'The business needs to audit unit purchase price variances between contracted baselines and actual invoiced costs across all order line items so that finance controllers can prevent cost leakage and recover unauthorized price increments.', 'SH-05 (Finance Controller)', 'HMW-04', 'Must', 'Percentage and currency cost variance: ((Invoiced Price - Contract Price) / Contract Price) * 100. Automated flagging of positive variance.'),
    ('BR-05', 'The business needs to isolate supplier delivery outliers and tail delays exceeding seven business days so that sourcing teams can remove highly volatile vendors from critical path manufacturing schedules.', 'SH-02 (Buyer), SH-01 (Head SCM)', 'HMW-05', 'Should', 'Count and statistical distribution frequency of delivery delays exceeding 7 calendar days beyond promised date.'),
    ('BR-06', 'The business needs to synthesize delivery timeliness, quality compliance, and price adherence into a weighted Composite Performance Index (CPI) ranking so that management can objectively categorize suppliers into standardized governance tiers.', 'SH-01 (Head SCM)', 'HMW-06', 'Must', 'Composite index score calculated on a 0-100 scale: (40% OTIF + 35% Quality Score + 25% Cost Score); validation against Tier A/B/C thresholds.'),
    ('BR-07', 'The business needs to quantify the total procurement expenditure allocated to high-risk and underperforming suppliers so that leadership can mitigate supply chain vulnerabilities and reallocate capital toward reliable vendors.', 'SH-01 (Head SCM), SH-05 (Finance)', 'HMW-07', 'Must', 'Spend-at-Risk ratio: (Total spend with Tier C vendors / Total procurement spend) * 100. Organizational target < 10% Spend-at-Risk.'),
    ('BR-08', 'The business needs to analyze chronological and quarterly trends in vendor fulfillment and quality metrics so that category managers can anticipate seasonal slumps and secure capacity ahead of peak operational cycles.', 'SH-01 (Head SCM), SH-02 (Buyer)', 'HMW-08', 'Should', 'Quarterly trend tracking of OTIF% and defect rates across 4 rolling quarters to identify cyclical dips exceeding 5 percentage points.'),
    ('BR-09', 'The business needs to identify suppliers eligible for strategic contract renewal or requiring mandatory Corrective Action Plans (CAPA) based on objective threshold criteria so that vendor governance is enforced systematically and legally.', 'SH-01 (Head SCM), SH-03 (QA Lead)', 'HMW-09', 'Must', 'Automated governance status classification (Renew / Retain / CAPA Audit) triggered when scores breach agreed operational thresholds.'),
    ('BR-10', 'The business needs to evaluate vendor delivery lead times and transit delays across shipping modes and geographic vendor origins so that logistics planners can select the most predictable transport routes and origin locations.', 'SH-04 (Warehouse), SH-02 (Buyer)', 'HMW-10', 'Should', 'Comparative average transit days and delay metrics broken down across freight shipping modes (Air, Sea, Road, Rail) and supplier regions.'),
    ('BR-11', 'The business needs to filter and segment procurement performance by product category, date range, supplier, and order status so that cross-functional teams can conduct targeted root-cause investigations.', 'All Stakeholders (SH-01 - SH-05)', 'HMW-01, HMW-02, HMW-03', 'Must', 'Dynamic segmentation across all analytical views enabling instantaneous slicing across categories, vendors, and date intervals.'),
    ('BR-12', 'The business needs to export standardized supplier performance scorecards and audit summaries into structured reports so that procurement teams can present empirical data during formal vendor contract reviews.', 'SH-01 (Head SCM), SH-02 (Buyer)', 'HMW-06, HMW-09', 'Could', 'On-demand export of tabular vendor scorecards and KPI performance records into portable, standardized file formats.')
]

for i, row in enumerate(br_data):
    for j, val in enumerate(row):
        c = br_table.cell(i+1, j)
        c.text = val
        if (i % 2) == 1:
            set_cell_background(c, 'F7FAFC')
        c.paragraphs[0].runs[0].font.size = Pt(8)
        set_cell_margins(c, 70, 70, 70, 70)

doc.add_paragraph().paragraph_format.space_after = Pt(8)

# Section 3: MoSCoW Prioritization Distribution
h3 = doc.add_heading('3. MoSCoW Prioritization Breakdown', level=1)
h3.paragraph_format.space_before = Pt(12)
h3.paragraph_format.space_after = Pt(6)

p_moscow = doc.add_paragraph(
    'The requirements portfolio adheres strictly to balanced agile scoping standards, ensuring that core operational capabilities are non-negotiable while maintaining distinct tiers for secondary enhancements:'
)
moscow_summary = [
    ('Must-Have (8 Requirements - 66.7%):', ' BR-01 (OTIF Compliance), BR-02 (Lead Time & SLA), BR-03 (Defect Rate & PPM), BR-04 (Price Variance Audit), BR-06 (Composite Scorecard), BR-07 (Spend-at-Risk), BR-09 (CAPA & Renewal Governance), BR-11 (Multi-dimensional Segmentation). These form the core analytical foundation.'),
    ('Should-Have (3 Requirements - 25.0%):', ' BR-05 (Tail Delay Outliers), BR-08 (Seasonal & Quarterly Trends), BR-10 (Logistics Mode & Geography). High-value analytical capabilities that enrich category management decision-making.'),
    ('Could-Have (1 Requirement - 8.3%):', ' BR-12 (Scorecard Report Exporting). Useful capability for offline supplier meetings.'),
    ('Won\'t-Have (Current Release Boundary):', ' Bidirectional enterprise ERP transactional writes and automated electronic bank invoice clearing (explicitly excluded to preserve academic sandbox safety and eliminate cloud infrastructure costs).')
]
for title, desc in moscow_summary:
    p = doc.add_paragraph(style='List Bullet')
    r1 = p.add_run(title)
    r1.bold = True
    r2 = p.add_run(desc)
    p.paragraph_format.space_after = Pt(3)

# Section 4: Required Underlying Data Inventory
h4 = doc.add_heading('4. Underlying Data Requirements Mapping', level=1)
h4.paragraph_format.space_before = Pt(12)
h4.paragraph_format.space_after = Pt(6)

data_table = doc.add_table(rows=13, cols=3)
data_table.alignment = WD_TABLE_ALIGNMENT.CENTER
data_headers = ['BR ID', 'Business Capability', 'Required Relational Data Attributes & Entities']
for j, h in enumerate(data_headers):
    c = data_table.cell(0, j)
    c.text = h
    set_cell_background(c, '1A365D')
    p = c.paragraphs[0]
    p.runs[0].bold = True
    p.runs[0].font.color.rgb = RGBColor(255, 255, 255)
    p.runs[0].font.size = Pt(9)
    set_cell_margins(c, 100, 100, 100, 100)

data_mapping = [
    ('BR-01', 'OTIF Compliance Tracking', 'Orders.promised_delivery_date, Deliveries.actual_delivery_date, Orders.ordered_qty, Deliveries.delivered_qty'),
    ('BR-02', 'Lead Time & SLA Adherence', 'Orders.po_creation_date, Deliveries.actual_delivery_date, Products.contracted_lead_time_days, Categories.category_name'),
    ('BR-03', 'Defect Rate & PPM Monitoring', 'Inspections.inspected_units, Inspections.rejected_units, Inspections.defect_code, Products.product_id, Suppliers.supplier_id'),
    ('BR-04', 'Unit Price Variance Auditing', 'Order_Items.contracted_unit_price, Order_Items.invoiced_unit_price, Order_Items.quantity, Orders.po_total_value'),
    ('BR-05', 'Delivery Volatility & Outlier Isolation', 'Deliveries.delay_days, Deliveries.delivery_status_code, Orders.order_id, Suppliers.supplier_name'),
    ('BR-06', 'Composite Performance Scoring (CPI)', 'Computed OTIF score, computed Quality compliance score, computed Price variance compliance score, KPI configuration weights'),
    ('BR-07', 'Spend-at-Risk Exposure Analytics', 'Orders.po_total_value, Suppliers.composite_performance_tier (A/B/C), Orders.fiscal_quarter'),
    ('BR-08', 'Seasonal & Quarterly Trend Analysis', 'Orders.po_creation_date, Orders.fiscal_year, Orders.fiscal_quarter, Deliveries.delivery_month'),
    ('BR-09', 'Automated CAPA & Renewal Governance', 'Suppliers.composite_score, Suppliers.consecutive_delayed_orders_count, Inspections.critical_defect_count, Contracts.expiration_date'),
    ('BR-10', 'Logistics Mode & Geography Evaluation', 'Suppliers.supplier_country, Shipments.shipping_mode (Air/Sea/Road/Rail), Deliveries.transit_lead_time_days'),
    ('BR-11', 'Multi-dimensional Filter & Slicing', 'Categories.category_id, Suppliers.supplier_id, Orders.order_date_range, Orders.order_status'),
    ('BR-12', 'Scorecard Audit Summary Export', 'Aggregated supplier scorecard summary view, KPI metrics snapshot, export generation timestamp')
]

for i, row in enumerate(data_mapping):
    for j, val in enumerate(row):
        c = data_table.cell(i+1, j)
        c.text = val
        if (i % 2) == 1:
            set_cell_background(c, 'F7FAFC')
        c.paragraphs[0].runs[0].font.size = Pt(8.5)
        set_cell_margins(c, 70, 70, 80, 80)

doc.add_paragraph().paragraph_format.space_after = Pt(8)

# Section 5: Bidirectional Gap Analysis
h5 = doc.add_heading('5. Bidirectional Traceability Gap Analysis', level=1)
h5.paragraph_format.space_before = Pt(12)
h5.paragraph_format.space_after = Pt(6)

p_gap1 = doc.add_paragraph()
r_g1 = p_gap1.add_run('A. Forward Traceability Audit (HMW → BR):')
r_g1.bold = True
doc.add_paragraph(
    'Every single opportunity statement from Step 2 (HMW-01 through HMW-10) is mapped to one or more Business Requirements. Specifically:\n'
    '• HMW-01 → BR-01, BR-11\n'
    '• HMW-02 → BR-02, BR-11\n'
    '• HMW-03 → BR-03, BR-11\n'
    '• HMW-04 → BR-04\n'
    '• HMW-05 → BR-05\n'
    '• HMW-06 → BR-06, BR-12\n'
    '• HMW-07 → BR-07\n'
    '• HMW-08 → BR-08\n'
    '• HMW-09 → BR-09, BR-12\n'
    '• HMW-10 → BR-10\n'
    'Forward Gap Count: 0 unmapped HMW statements.'
)

p_gap2 = doc.add_paragraph()
r_g2 = p_gap2.add_run('B. Backward Traceability Audit (BR → HMW):')
r_g2.bold = True
doc.add_paragraph(
    'Every defined Business Requirement (BR-01 through BR-12) originates from a validated HMW opportunity statement. No orphan, superfluous, or scope-creep requirements were introduced.\n'
    'Backward Gap Count: 0 ungrounded Business Requirements.'
)

p_gap3 = doc.add_paragraph()
r_g3 = p_gap3.add_run('C. Testability & Measurability Audit:')
r_g3.bold = True
doc.add_paragraph(
    '100% of the 12 business requirements specify clear, unambiguous quantitative metrics and mathematical criteria (e.g., percentages, PPM, calendar day deltas, currency variances). None are ambiguous or subjective. All requirements are fully testable in the QA phase (Step 13).'
)

# Save Document
output_path = os.path.join('docs', '03_Needs_Analysis.docx')
doc.save(output_path)
print(f'Successfully generated: {output_path}')
