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

# Title & Header
p_title = doc.add_paragraph()
r_title = p_title.add_run('STEP 2: HOW MIGHT WE (HMW) STATEMENTS')
r_title.bold = True
r_title.font.size = Pt(20)
r_title.font.color.rgb = RGBColor(24, 43, 73)
p_title.paragraph_format.space_after = Pt(2)

p_sub = doc.add_paragraph()
r_sub = p_sub.add_run('Design-Thinking Problem Reframing & Stakeholder Opportunity Mapping')
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
    ('Document Reference', 'docs/02_HMW_Statements.docx | Derived from Step 1 Project Understanding')
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

# Section 1: Design Thinking Framework
h1 = doc.add_heading('1. Design-Thinking Problem Reframing Methodology', level=1)
h1.paragraph_format.space_before = Pt(12)
h1.paragraph_format.space_after = Pt(6)

p_meth = doc.add_paragraph(
    'In Phase 2 of our engineering lifecycle, we employ the Stanford d.school Design Thinking framework to translate organizational friction into actionable innovation vectors. Rather than prematurely jumping to functional UI widgets or database queries, we dissect the specific operational pain points of our five core stakeholder groups (SH-01 through SH-05) identified in Step 1. By reframing each unresolved friction point into an open, generative \"How Might We\" (HMW) statement, we ensure the analytical architecture addresses underlying human and organizational needs rather than merely showcasing technical features.'
)
p_meth.paragraph_format.line_spacing = 1.15
p_meth.paragraph_format.space_after = Pt(8)

# Section 2: Stakeholder Pain Points
h2 = doc.add_heading('2. Stakeholder Pain Point Inventory', level=1)
h2.paragraph_format.space_before = Pt(12)
h2.paragraph_format.space_after = Pt(6)

sh_pain_points = [
    ('SH-01: Head of Procurement / SCM Director', [
        'Lack of macro-level visibility into total procurement spend exposed to poor-performing or high-risk suppliers.',
        'Annual contract renegotiations and vendor tier re-allocations are hindered by subjective, anecdotal arguments rather than defensible multi-dimensional scorecard data.',
        'Inability to anticipate seasonal or quarterly fulfillment degradations before they compromise customer order deliveries.'
    ]),
    ('SH-02: Procurement & Sourcing Specialist', [
        'Allocating purchase orders to suppliers who show high lead-time volatility, resulting in severe scheduling bottlenecks on assembly lines.',
        'Difficulty verifying whether vendors consistently fulfill orders in full versus routinely making partial shipments that force complex back-order tracking.',
        'Inability to benchmark suppliers within the same category to identify alternative high-performing vendors.'
    ]),
    ('SH-03: Quality Assurance (QA) Manager', [
        'Inability to quickly isolate component-level defect spikes and detect whether failures stem from specific supplier batches or chronic process flaws.',
        'Lacking centralized empirical defect metrics (e.g., Defect PPM) to formally trigger Corrective Action Plans (CAPA) or demand vendor credit notes.',
        'Time-consuming manual cross-referencing between warehouse receipt records and QA inspection logs.'
    ]),
    ('SH-04: Warehouse & Inventory Supervisor', [
        'Erratic, unpredictable delivery lead times force the warehouse to maintain unnecessarily large safety stock buffers, tying up working capital.',
        'Lack of visibility into how logistics shipping modes and geographic vendor origins impact receiving dock schedules and dock congestion.'
    ]),
    ('SH-05: Finance & Procurement Controller', [
        'Unapproved purchase price creeping where invoiced unit prices quietly exceed agreed contracted baselines across thousands of PO line items.',
        'Difficulty quantifying the aggregate financial cost variance across spend categories caused by vendor non-compliance and inflation.'
    ])
]

for sh_title, points in sh_pain_points:
    p_sh = doc.add_paragraph()
    r = p_sh.add_run(sh_title)
    r.bold = True
    r.font.color.rgb = RGBColor(26, 54, 93)
    p_sh.paragraph_format.space_after = Pt(2)
    for pt in points:
        p_pt = doc.add_paragraph(style='List Bullet')
        p_pt.add_run(pt)
        p_pt.paragraph_format.space_after = Pt(2)

doc.add_paragraph().paragraph_format.space_after = Pt(8)

# Section 3: HMW Statements Table
h3 = doc.add_heading('3. \"How Might We\" (HMW) Opportunity Mapping Matrix', level=1)
h3.paragraph_format.space_before = Pt(12)
h3.paragraph_format.space_after = Pt(6)

hmw_table = doc.add_table(rows=11, cols=5)
hmw_table.alignment = WD_TABLE_ALIGNMENT.CENTER
hmw_headers = ['HMW ID', 'How Might We ...?', 'Stakeholder', 'Root Pain Point Addressed', 'Linked OQ']
for j, h in enumerate(hmw_headers):
    c = hmw_table.cell(0, j)
    c.text = h
    set_cell_background(c, '1A365D')
    p = c.paragraphs[0]
    p.runs[0].bold = True
    p.runs[0].font.color.rgb = RGBColor(255, 255, 255)
    p.runs[0].font.size = Pt(9)
    set_cell_margins(c, 100, 100, 100, 100)

hmw_data = [
    ('HMW-01', 'How might we help procurement specialists instantly identify suppliers who reliably deliver On-Time In-Full (OTIF) versus those who routinely short-ship or delay orders?', 'SH-02 (Buyer)', 'Inability to detect chronic short-shipping and late arrivals prior to PO award.', 'OQ-01'),
    ('HMW-02', 'How might we enable warehouse and procurement teams to uncover lead time variability and SLA deviations so safety stock buffers can be accurately calibrated?', 'SH-04 (Warehouse), SH-02 (Buyer)', 'Unpredictable vendor lead times driving safety stock inflation and inventory bloat.', 'OQ-02'),
    ('HMW-03', 'How might we empower QA managers to rapidly detect component defect spikes and identify vendors exceeding acceptable Defective PPM thresholds?', 'SH-03 (QA Lead)', 'Delayed detection of sub-standard supplier components entering production lines.', 'OQ-03'),
    ('HMW-04', 'How might we enable procurement controllers to continuously detect unit price creep between contracted baselines and actual invoiced amounts across purchase orders?', 'SH-05 (Finance Controller)', 'Undetected invoice price discrepancies eroding procurement profit margins.', 'OQ-04'),
    ('HMW-05', 'How might we help sourcing teams isolate vendors exhibiting severe delivery volatility and chronic tail delays exceeding one business week?', 'SH-02 (Buyer), SH-01 (Head of SCM)', 'High-variance supplier fulfillment causing unpredictable line stoppages.', 'OQ-05'),
    ('HMW-06', 'How might we provide procurement leadership with a balanced, multi-dimensional Composite Performance Index (CPI) that synthesizes delivery, quality, and cost metrics into an objective vendor tier?', 'SH-01 (Head of SCM)', 'Vendor evaluations relying on fragmented metrics or gut feel during contract renewals.', 'OQ-06'),
    ('HMW-07', 'How might we give procurement executives instant visibility into what portion of enterprise spend is exposed to high-risk, underperforming suppliers?', 'SH-01 (Head of SCM), SH-05 (Finance)', 'Blindness toward total capital expenditure concentrated with unreliable Tier C vendors.', 'OQ-07'),
    ('HMW-08', 'How might we assist category managers in identifying cyclical, seasonal, or quarterly performance degradation patterns before peak operational periods?', 'SH-01 (Head of SCM), SH-02 (Buyer)', 'Lack of foresight into seasonal fulfillment slumps and quarterly quality deterioration.', 'OQ-08'),
    ('HMW-09', 'How might we provide automated governance triggers that recommend suppliers for strategic contract renewal or mandate formal Corrective Action Plans (CAPA)?', 'SH-01 (Head of SCM), SH-03 (QA Lead)', 'Subjective or delayed enforcement of corrective action against chronic underperforming suppliers.', 'OQ-09'),
    ('HMW-10', 'How might we help logistics coordinators and sourcing planners evaluate the impact of supplier geography and transit modes on delivery predictability?', 'SH-04 (Warehouse), SH-02 (Buyer)', 'Lack of analytical insight into whether delays are caused by vendor logistics modes or geography.', 'OQ-10')
]

for i, row in enumerate(hmw_data):
    for j, val in enumerate(row):
        c = hmw_table.cell(i+1, j)
        c.text = val
        if (i % 2) == 1:
            set_cell_background(c, 'F7FAFC')
        c.paragraphs[0].runs[0].font.size = Pt(8.5)
        set_cell_margins(c, 70, 70, 80, 80)

doc.add_paragraph().paragraph_format.space_after = Pt(8)

# Section 4: Top 5 Priority Opportunities
h4 = doc.add_heading('4. Top 5 Priority Opportunities & Strategic Justifications', level=1)
h4.paragraph_format.space_before = Pt(12)
h4.paragraph_format.space_after = Pt(6)

priorities = [
    ('PRIORITY 1: HMW-01 (OTIF Delivery Adherence):', ' OTIF fulfillment reliability is the fundamental operational prerequisite to preventing manufacturing assembly halts and downstream customer delivery failures.'),
    ('PRIORITY 2: HMW-02 (Lead Time Variability & SLA Adherence):', ' Lead time unpredictability directly drives excess safety stock holding costs and causes master production schedule instability.'),
    ('PRIORITY 3: HMW-03 (Component Defect Spikes & Quality Audits):', ' Defective raw materials lead directly to catastrophic production line scrap, costly warranty liabilities, and irreversible brand damage.'),
    ('PRIORITY 4: HMW-06 (Composite Performance Index & Tiering):', ' Strategic vendor consolidation and annual contract negotiations require a single, mathematically sound composite benchmark that eliminates subjective bias.'),
    ('PRIORITY 5: HMW-07 (Spend-at-Risk Exposure Analytics):', ' Pinpoints critical enterprise-wide financial and supply continuity exposure by quantifying capital tied to vulnerable or underperforming suppliers.')
]

for title, reason in priorities:
    p = doc.add_paragraph(style='List Bullet')
    r1 = p.add_run(title)
    r1.bold = True
    r2 = p.add_run(reason)
    p.paragraph_format.space_after = Pt(3)

# Section 5: Coverage Audit
h5 = doc.add_heading('5. Operational Question (OQ) Coverage Audit', level=1)
h5.paragraph_format.space_before = Pt(12)
h5.paragraph_format.space_after = Pt(6)

p_cov = doc.add_paragraph(
    'A cross-verification audit confirms that 100% of the operational questions defined in Step 1 (OQ-01 through OQ-10) are mapped to at least one generative HMW statement (HMW-01 through HMW-10). This establishes bidirectional alignment between user needs, business questions, and technical requirements.'
)
p_cov.paragraph_format.space_after = Pt(12)

# Save Document
output_path = os.path.join('docs', '02_HMW_Statements.docx')
doc.save(output_path)
print(f'Successfully generated: {output_path}')
