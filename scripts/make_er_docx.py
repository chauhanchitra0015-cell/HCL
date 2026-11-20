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
r_title = p_title.add_run('STEP 9: ENTITY RELATIONSHIP (ER) SPECIFICATION')
r_title.bold = True
r_title.font.size = Pt(22)
r_title.font.color.rgb = RGBColor(24, 43, 73)
p_title.paragraph_format.space_after = Pt(2)

p_sub = doc.add_paragraph()
r_sub = p_sub.add_run('Conceptual & Logical Relational Schema Modeling, Cardinality & Normalization Review')
r_sub.font.size = Pt(13)
r_sub.font.color.rgb = RGBColor(90, 105, 120)
p_sub.paragraph_format.space_after = Pt(14)

# Document Control Table
meta_table = doc.add_table(rows=6, cols=2)
meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
meta_data = [
    ('Project Title', 'Supplier Performance Analysis & Procurement Intelligence System'),
    ('Document Version', 'Version 1.0 (Formal Engineering Baseline)'),
    ('Lead Author & Role', 'Chitra Chauhan | Roll No: 2500320140043 | Solution Architect & Developer'),
    ('Academic Mentor', 'Prof. Kalpana Dwivedi (Dept. of Computer Applications, ABES EC)'),
    ('Relational Architecture', '3NF Star/Snowflake Schema (3 Dimensions, 2 Fact Tables, 1 Analytical View)'),
    ('Associated Graphic', 'docs/09_ER_Diagram.png (High-Resolution Graphic Embedded Below)')
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

# Section 1: Embedded ER Diagram Graphic
h1 = doc.add_heading('1. Visual Entity Relationship Diagram', level=1)
h1.paragraph_format.space_before = Pt(12)
h1.paragraph_format.space_after = Pt(6)

png_path = os.path.join('docs', '09_ER_Diagram.png')
if os.path.exists(png_path):
    p_img = doc.add_paragraph()
    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_picture(png_path, width=Inches(6.5))
    p_cap = doc.add_paragraph('Figure 9.1: Formal Relational Entity Relationship Diagram (ERD) with Crow\'s Foot Cardinality Notation.')
    p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cap.runs[0].font.size = Pt(9)
    p_cap.runs[0].font.italic = True
    p_cap.paragraph_format.space_after = Pt(12)

# Section 2: Mermaid erDiagram Specification
h2 = doc.add_heading('2. Mermaid erDiagram Code Specification', level=1)
h2.paragraph_format.space_before = Pt(12)
h2.paragraph_format.space_after = Pt(6)

p_m = doc.add_paragraph(
    'The formal text-based structural definition of the ERD below is compatible with live Mermaid rendering engines:'
)
p_m_code = doc.add_paragraph()
mermaid_code = """erDiagram
    suppliers ||--o{ purchase_orders : "fulfills"
    categories ||--o{ products : "classifies"
    purchase_orders ||--|{ order_deliveries : "contains"
    products ||--o{ order_deliveries : "itemizes"

    suppliers {
        string supplier_id PK
        string supplier_name
        string supplier_country
        string supplier_city
        string contact_email
        string supplier_tier
        decimal rating_score
        date contract_start_date
        date contract_end_date
        string status
    }

    categories {
        string category_id PK
        string category_name
        string spend_tier
        decimal risk_weight
        string description
    }

    products {
        string product_id PK
        string product_name
        string category_id FK
        string unit_of_measure
        decimal baseline_contract_price
        int standard_lead_time_days
        int target_defect_ppm
    }

    purchase_orders {
        string order_id PK
        string po_number
        string supplier_id FK
        date order_date
        date promised_delivery_date
        string order_status
        string payment_terms
        decimal total_po_value
        int fiscal_year
        string fiscal_quarter
    }

    order_deliveries {
        string delivery_id PK
        string order_id FK
        string product_id FK
        int ordered_quantity
        int delivered_quantity
        date actual_delivery_date
        int delivery_lead_time_days
        int delay_days
        string shipping_mode
        decimal contracted_unit_price
        decimal invoiced_unit_price
        decimal price_variance_amount
        decimal price_variance_pct
        int inspected_quantity
        int defective_quantity
        decimal defect_rate_pct
        int defect_ppm
        string defect_reason_code
        int is_on_time
        int is_in_full
        int is_otif
    }"""
r_mc = p_m_code.add_run(mermaid_code)
r_mc.font.name = 'Courier New'
r_mc.font.size = Pt(8.5)
p_m_code.paragraph_format.space_after = Pt(10)

# Section 3: Plain-English Relationship Statements
h3 = doc.add_heading('3. Plain-English Relationship & Cardinality Declarations', level=1)
h3.paragraph_format.space_before = Pt(12)
h3.paragraph_format.space_after = Pt(6)

rel_statements = [
    ('Relationship 1: suppliers to purchase_orders (1 : N)',
     'One registered supplier receives and fulfills zero, one, or many purchase orders over time, while each purchase order is contractually placed with exactly one supplier.',
     'suppliers.supplier_id (PK) -> purchase_orders.supplier_id (FK)',
     'Active Verb: "fulfills" | Cardinality: 1 : 0..N | Delete Rule: RESTRICT'),

    ('Relationship 2: categories to products (1 : N)',
     'One commodity category classifies zero, one, or many component products, while each individual product catalog item belongs to exactly one procurement category.',
     'categories.category_id (PK) -> products.category_id (FK)',
     'Active Verb: "classifies" | Cardinality: 1 : 0..N | Delete Rule: RESTRICT'),

    ('Relationship 3: purchase_orders to order_deliveries (1 : N)',
     'One purchase order contains one or many delivery line items, while each individual delivery line item is physically tied to exactly one parent purchase order.',
     'purchase_orders.order_id (PK) -> order_deliveries.order_id (FK)',
     'Active Verb: "contains" | Cardinality: 1 : 1..N | Delete Rule: CASCADE'),

    ('Relationship 4: products to order_deliveries (1 : N)',
     'One component product is itemized across zero, one, or many delivery line items, while each delivery line item specifies exactly one product SKU.',
     'products.product_id (PK) -> order_deliveries.product_id (FK)',
     'Active Verb: "itemizes" | Cardinality: 1 : 0..N | Delete Rule: RESTRICT')
]

for title, desc, path, tech in rel_statements:
    p = doc.add_paragraph()
    r1 = p.add_run(title + '\n')
    r1.bold = True
    r1.font.color.rgb = RGBColor(26, 54, 93)
    p.add_run(f'• Plain-English Description: {desc}\n')
    p.add_run(f'• Foreign Key Mapping: {path}\n')
    p.add_run(f'• Technical Specification: {tech}\n')
    p.paragraph_format.space_after = Pt(4)

# Section 4: Business Rules Enforced by the Schema
h4 = doc.add_heading('4. Business Rules Enforced by the Relational Architecture', level=1)
h4.paragraph_format.space_before = Pt(12)
h4.paragraph_format.space_after = Pt(6)

rules = [
    ('B-RULE-01 (Referential Integrity & Master Retention):', ' Deletion of a vendor in suppliers or a SKU in products is strictly restricted (ON DELETE RESTRICT) if historical purchase orders or delivery records reference their keys. This preserves historical scorecard audits.'),
    ('B-RULE-02 (Cascading Header-Detail Deletion):', ' When a purchase order header is cancelled or purged prior to fulfillment, all associated delivery detail lines are automatically cleaned up synchronously (ON DELETE CASCADE).'),
    ('B-RULE-03 (Chronological Date Ordering):', ' A purchase order cannot have a promised delivery date that precedes its issuance date: CHECK(promised_delivery_date >= order_date). Similarly, vendor contracts enforce: CHECK(contract_end_date >= contract_start_date).'),
    ('B-RULE-04 (Inspection Quantity Conservation):', ' The number of defective components identified cannot mathematically exceed the total quantity of inspected units: CHECK(defective_quantity <= inspected_quantity). Furthermore, all physical quantities must be non-negative.'),
    ('B-RULE-05 (Deterministic Boolean KPI Domains):', ' To ensure zero-ambiguity vectorized operations in Python, KPI indicators (is_on_time, is_in_full, is_otif) are constrained strictly to integer binary states: CHECK(is_otif IN (0, 1)).')
]
for title, text in rules:
    p = doc.add_paragraph(style='List Bullet')
    r1 = p.add_run(title)
    r1.bold = True
    r2 = p.add_run(text)
    p.paragraph_format.space_after = Pt(3)

# Section 5: Design Problems, Many-to-Many Resolution & Architectural Fixes
h5 = doc.add_heading('5. Architectural Design Review, Dilemmas & Engineering Fixes', level=1)
h5.paragraph_format.space_before = Pt(12)
h5.paragraph_format.space_after = Pt(6)

design_reviews = [
    ('5.1 Resolution of Many-to-Many (M:N) Relationship via Associative Bridge Entity:',
     '• The Problem: In enterprise supply chains, a single purchase order contract contains multiple distinct product SKUs, and conversely, a given product SKU is ordered repeatedly across multiple purchase orders. Direct association between purchase_orders and products creates an M:N cardinality trap that relational engines cannot represent without massive data redundancy.\n'
     '• The Engineering Fix: We introduced order_deliveries as an associative bridge entity. By incorporating both order_id (FK) and product_id (FK), order_deliveries decomposes the M:N relationship into two clean, normalized 1:N relationships. Furthermore, by attaching fulfillment dates, invoicing variances, and QA inspection outcomes directly to this line-item grain, the entity models the precise operational reality of procurement.'),

    ('5.2 Performance Optimization vs. Pure Normalization (Derived Attribute Strategy):',
     '• The Problem: In strict Third Normal Form (3NF), calculated attributes—such as delivery_lead_time_days, delay_days, price_variance_pct, and is_otif—are considered transitive dependencies that should never be persisted physically on disk, but calculated dynamically on every query.\n'
     '• The Dilemma: Streamlit operates in a reactive execution model where every UI filter click re-runs the Python script. Recalculating calendar deltas and multi-condition boolean checks across thousands of historical records in Python on every click introduces perceptible UI latency.\n'
     '• The Engineering Fix: We adopted a hybrid analytical storage pattern: derived attributes are physically stored in order_deliveries and indexed. To prevent data corruption, rigorous database CHECK constraints enforce data consistency at the persistence layer, while eliminating computational bottlenecks during dashboard interaction.'),

    ('5.3 Entity Consolidation vs. Over-Normalization (Fulfillment & QA Inspection):',
     '• The Problem: Traditional academic database models frequently split receiving docks from inspection labs, creating a 6th table (quality_inspections).\n'
     '• The Dilemma: In our procurement domain, incoming shipments undergo immediate dockside sampling before inventory receipt. Creating an additional table introduces a redundant 1:1 or 1:N join on every single query without providing additional analytical granularity.\n'
     '• The Engineering Fix: Receiving fulfillment and QA inspection metrics were consolidated into order_deliveries. This maintains a clean 5-table architecture, optimizes cache density, and ensures instantaneous analytical joins in SQLite.')
]

for section_title, content in design_reviews:
    p = doc.add_paragraph()
    r = p.add_run(section_title)
    r.bold = True
    r.font.size = Pt(11)
    r.font.color.rgb = RGBColor(26, 54, 93)
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(2)
    
    p_body = doc.add_paragraph(content)
    p_body.paragraph_format.line_spacing = 1.15
    p_body.paragraph_format.space_after = Pt(6)

# Save Document
output_path = os.path.join('docs', '09_ER_Diagram.docx')
doc.save(output_path)
print(f'Successfully generated: {output_path}')
