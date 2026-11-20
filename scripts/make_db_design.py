import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

os.makedirs('docs', exist_ok=True)
wb = openpyxl.Workbook()

# Style Palette
NAVY_HEADER = "1A365D"
WHITE = "FFFFFF"
LIGHT_BLUE = "F0F4F8"
ZEBRA_FILL = "F9FAFB"
BORDER_GRAY = "D1D5DB"
SUCCESS_GREEN = "10B981"
PK_COLOR = "7C3AED"
FK_COLOR = "2563EB"

header_fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
header_font = Font(name="Calibri", size=11, bold=True, color=WHITE)
header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

thin_border = Border(
    left=Side(style='thin', color=BORDER_GRAY),
    right=Side(style='thin', color=BORDER_GRAY),
    top=Side(style='thin', color=BORDER_GRAY),
    bottom=Side(style='thin', color=BORDER_GRAY)
)

def style_table(ws, start_row, headers, data, col_widths, center_cols=[]):
    ws.row_dimensions[start_row].height = 26
    for col_idx, h in enumerate(headers, start=1):
        c = ws.cell(row=start_row, column=col_idx, value=h)
        c.fill = header_fill
        c.font = header_font
        c.alignment = header_align
        c.border = thin_border
        
    for r_offset, row_vals in enumerate(data):
        curr_row = start_row + 1 + r_offset
        ws.row_dimensions[curr_row].height = 22
        zebra = PatternFill(start_color=ZEBRA_FILL if r_offset % 2 == 1 else "FFFFFF", fill_type="solid")
        for col_idx, val in enumerate(row_vals, start=1):
            c = ws.cell(row=curr_row, column=col_idx, value=val)
            c.font = Font(name="Calibri", size=10)
            c.fill = zebra
            c.border = thin_border
            if col_idx in center_cols:
                c.alignment = Alignment(horizontal="center", vertical="center")
            else:
                c.alignment = Alignment(horizontal="left", vertical="center")
                
            # Key badges
            if str(val) == "PK":
                c.font = Font(name="Calibri", size=10, bold=True, color=PK_COLOR)
            elif str(val) == "FK":
                c.font = Font(name="Calibri", size=10, bold=True, color=FK_COLOR)
                
    for col_letter, w in col_widths.items():
        ws.column_dimensions[col_letter].width = w

# -------------------------------------------------------------
# Sheet 1: Overview
# -------------------------------------------------------------
ws_ov = wb.active
ws_ov.title = "Overview"
ws_ov.views.sheetView[0].showGridLines = True

ws_ov.merge_cells("A1:G1")
top_c = ws_ov["A1"]
top_c.value = "SUPPLIER PERFORMANCE ANALYSIS — RELATIONAL DATABASE ARCHITECTURE"
top_c.font = Font(name="Calibri", size=14, bold=True, color=WHITE)
top_c.fill = header_fill
top_c.alignment = Alignment(horizontal="center", vertical="center")
ws_ov.row_dimensions[1].height = 32

meta_rows = [
    ("Project Title", "Supplier Performance Analysis & Procurement Intelligence System"),
    ("Learner & Roll", "Chitra Chauhan | Roll No: 2500320140043 | MCA 2nd Year (Data Analytics), ABES"),
    ("Academic Mentor", "Prof. Kalpana Dwivedi (Department of Computer Applications, ABES EC)"),
    ("Database Engine", "SQLite 3 (Embedded Relational Database in data/app.db)"),
    ("Schema Pattern", "Normalized 3NF Relational Star/Snowflake Schema (3 Master Dimensions, 2 Fact Tables)"),
    ("Analytical View", "v_supplier_performance_analysis (41 fully joined analytical attributes)")
]
for idx, (lbl, val) in enumerate(meta_rows, start=3):
    ws_ov.cell(row=idx, column=1, value=lbl).font = Font(name="Calibri", size=10, bold=True, color="1F2937")
    ws_ov.cell(row=idx, column=2, value=val).font = Font(name="Calibri", size=10, color="374151")
    ws_ov.row_dimensions[idx].height = 20

ws_ov.column_dimensions["A"].width = 20
ws_ov.column_dimensions["B"].width = 75

ws_ov.cell(row=10, column=1, value="Core Design Principles & Analytical Capabilities").font = Font(name="Calibri", size=12, bold=True, color=NAVY_HEADER)
principles = [
    "1. Strict Referential Integrity: Enforces foreign key constraints between suppliers, products, categories, orders, and deliveries.",
    "2. Granular Traceability: Every single Operational Question (OQ-01 through OQ-10) is directly backed by structured table columns.",
    "3. High-Performance Querying: Indexes established on all foreign key paths and date dimensions for sub-second aggregations.",
    "4. Deterministic Business Rules: CHECK constraints guarantee data hygiene (OTIF in 0/1, quantities >= 0, valid shipping modes)."
]
for idx, p in enumerate(principles, start=11):
    ws_ov.cell(row=idx, column=1, value=p).font = Font(name="Calibri", size=10, color="4B5563")
    ws_ov.row_dimensions[idx].height = 20

# -------------------------------------------------------------
# Sheet 2: Tables_Summary
# -------------------------------------------------------------
ws_sum = wb.create_sheet(title="Tables_Summary")
ws_sum.views.sheetView[0].showGridLines = True

sum_headers = ["Table Name", "Entity Type", "Grain / Level of Detail", "Primary Key", "Foreign Keys", "Column Count", "Target Rows", "Core Purpose"]
sum_data = [
    ["suppliers", "Master Dimension", "One row per vendor organization", "supplier_id", "None", 10, "25", "Stores vendor demographics, countries, ratings, and master tier classification."],
    ["categories", "Master Dimension", "One row per commodity spend category", "category_id", "None", 5, "6", "Procurement category taxonomy, spend classifications, and category risk weights."],
    ["products", "Master Dimension", "One row per purchased SKU / item", "product_id", "category_id", 7, "40", "Product catalog, baseline contracted prices, SLA lead time days, and target PPM."],
    ["purchase_orders", "Fact Table (Header)", "One row per purchase order contract", "order_id", "supplier_id", 10, "300", "Tracks order placement date, promised delivery dates, order totals, and status."],
    ["order_deliveries", "Fact Table (Line Detail)", "One row per order fulfillment & QA event", "delivery_id", "order_id, product_id", 21, "300 - 400", "Captures actual deliveries, lead times, delay days, cost variances, and QA inspection outcomes."],
    ["v_supplier_performance_analysis", "Analytical View", "One row per analytical delivery transaction", "delivery_id", "Joined Keys", 41, "300 - 400", "Unified analytical view joining all dimensions and facts for Streamlit and Plotly."]
]
sum_widths = {"A": 28, "B": 20, "C": 35, "D": 16, "E": 20, "F": 14, "G": 14, "H": 48}
style_table(ws_sum, 1, sum_headers, sum_data, sum_widths, center_cols=[2, 4, 5, 6, 7])

# -------------------------------------------------------------
# Sheet 3-7: Data Dictionaries
# -------------------------------------------------------------
dd_specs = {
    "DD_Suppliers": (
        ["Column Name", "Data Type", "Key", "Nullable", "Allowed Values / Constraints", "Description", "OQ / BR Supported"],
        [
            ["supplier_id", "VARCHAR(10)", "PK", "N", "Format: SUP-001 to SUP-025", "Unique primary identifier for supplier", "All OQs, BR-01..12"],
            ["supplier_name", "VARCHAR(100)", "None", "N", "Unique non-empty string", "Legal operating entity name of the vendor", "OQ-01, BR-01"],
            ["supplier_country", "VARCHAR(50)", "None", "N", "India, Germany, USA, Japan, South Korea, Vietnam, Taiwan", "Country of vendor manufacturing headquarters", "OQ-10, BR-10"],
            ["supplier_city", "VARCHAR(50)", "None", "N", "Valid municipality name", "City of primary supplier dispatch facility", "OQ-10, BR-10"],
            ["contact_email", "VARCHAR(100)", "None", "N", "Standard RFC email pattern", "Official representative procurement email", "BR-09"],
            ["supplier_tier", "VARCHAR(10)", "None", "N", "Tier A, Tier B, Tier C", "Strategic governance classification tier", "OQ-06, OQ-07, BR-06, BR-07"],
            ["rating_score", "DECIMAL(4,2)", "None", "N", "0.00 to 100.00", "Baseline historical supplier audit rating", "OQ-06, BR-06"],
            ["contract_start_date", "DATE", "None", "N", "YYYY-MM-DD", "Effective inception date of master service contract", "BR-09"],
            ["contract_end_date", "DATE", "None", "N", "YYYY-MM-DD >= contract_start_date", "Contractual expiration date for renewal audit", "OQ-09, BR-09"],
            ["status", "VARCHAR(20)", "None", "N", "Active, Conditional, Suspended", "Commercial operational status of vendor", "OQ-09, BR-09"]
        ],
        {"A": 22, "B": 16, "C": 10, "D": 10, "E": 30, "F": 45, "G": 22}
    ),
    "DD_Categories": (
        ["Column Name", "Data Type", "Key", "Nullable", "Allowed Values / Constraints", "Description", "OQ / BR Supported"],
        [
            ["category_id", "VARCHAR(10)", "PK", "N", "Format: CAT-01 to CAT-06", "Unique primary identifier for commodity category", "All OQs, BR-11"],
            ["category_name", "VARCHAR(50)", "None", "N", "Unique spend category title", "Procurement commodity taxonomy title", "OQ-02, BR-02, BR-11"],
            ["spend_tier", "VARCHAR(10)", "None", "N", "Strategic, Leverage, Bottleneck, Routine", "Kraljic portfolio purchasing classification", "OQ-07, BR-07"],
            ["risk_weight", "DECIMAL(3,2)", "None", "N", "0.10 to 1.00", "Category criticality multiplier for risk scoring", "OQ-06, BR-06"],
            ["description", "TEXT", "None", "Y", "Freeform string", "Operational scope and commodity definitions", "Documentation"]
        ],
        {"A": 20, "B": 16, "C": 10, "D": 10, "E": 28, "F": 45, "G": 22}
    ),
    "DD_Products": (
        ["Column Name", "Data Type", "Key", "Nullable", "Allowed Values / Constraints", "Description", "OQ / BR Supported"],
        [
            ["product_id", "VARCHAR(10)", "PK", "N", "Format: PRD-001 to PRD-040", "Unique primary identifier for component/SKU", "OQ-03, BR-03"],
            ["product_name", "VARCHAR(100)", "None", "N", "Non-empty string", "Item engineering description", "OQ-03, BR-03"],
            ["category_id", "VARCHAR(10)", "FK", "N", "References categories.category_id", "Parent commodity category identifier", "OQ-02, BR-02, BR-11"],
            ["unit_of_measure", "VARCHAR(15)", "None", "N", "Pieces, Kg, Meters, Sets, Boxes", "Standard inventory accounting unit", "Operations"],
            ["baseline_contract_price", "DECIMAL(10,2)", "None", "N", "> 0.00", "Agreed contracted baseline unit cost", "OQ-04, BR-04"],
            ["standard_lead_time_days", "INTEGER", "None", "N", "3 to 90 days", "Contractual SLA baseline manufacturing lead time", "OQ-02, BR-02"],
            ["target_defect_ppm", "INTEGER", "None", "N", ">= 0", "Maximum tolerable defective parts per million SLA", "OQ-03, BR-03"]
        ],
        {"A": 24, "B": 16, "C": 10, "D": 10, "E": 26, "F": 45, "G": 22}
    ),
    "DD_Purchase_Orders": (
        ["Column Name", "Data Type", "Key", "Nullable", "Allowed Values / Constraints", "Description", "OQ / BR Supported"],
        [
            ["order_id", "VARCHAR(15)", "PK", "N", "Format: PO-2024-001 to 300", "Unique primary identifier for purchase order", "All OQs, BR-01..12"],
            ["po_number", "VARCHAR(20)", "None", "N", "Unique PO reference string", "Official corporate ERP purchase order number", "BR-01"],
            ["supplier_id", "VARCHAR(10)", "FK", "N", "References suppliers.supplier_id", "Associated vendor fulfilling the order", "All OQs, BR-01..12"],
            ["order_date", "DATE", "None", "N", "YYYY-MM-DD", "Timestamp when purchase order was officially issued", "OQ-02, OQ-08, BR-02, BR-08"],
            ["promised_delivery_date", "DATE", "None", "N", "YYYY-MM-DD >= order_date", "Committed SLA delivery arrival deadline", "OQ-01, OQ-02, BR-01, BR-02"],
            ["order_status", "VARCHAR(20)", "None", "N", "Completed, In-Transit, Cancelled", "Current lifecycle status of purchase order", "BR-11"],
            ["payment_terms", "VARCHAR(20)", "None", "N", "Net 30, Net 45, Net 60, Immediate", "Commercial invoicing payment conditions", "Finance"],
            ["total_po_value", "DECIMAL(12,2)", "None", "N", ">= 0.00", "Aggregate monetary commitment of order in USD", "OQ-07, BR-07"],
            ["fiscal_year", "INTEGER", "None", "N", "2023 to 2026", "Accounting fiscal reporting year", "OQ-08, BR-08"],
            ["fiscal_quarter", "VARCHAR(5)", "None", "N", "Q1, Q2, Q3, Q4", "Accounting calendar fiscal quarter", "OQ-08, BR-08"]
        ],
        {"A": 22, "B": 16, "C": 10, "D": 10, "E": 28, "F": 45, "G": 22}
    ),
    "DD_Order_Deliveries": (
        ["Column Name", "Data Type", "Key", "Nullable", "Allowed Values / Constraints", "Description", "OQ / BR Supported"],
        [
            ["delivery_id", "VARCHAR(15)", "PK", "N", "Format: DEL-001 to DEL-300", "Unique primary identifier for delivery line", "All OQs, BR-01..12"],
            ["order_id", "VARCHAR(15)", "FK", "N", "References purchase_orders.order_id", "Parent purchase order identifier", "All OQs, BR-01..12"],
            ["product_id", "VARCHAR(10)", "FK", "N", "References products.product_id", "SKU component delivered in line item", "OQ-03, OQ-04, BR-03, BR-04"],
            ["ordered_quantity", "INTEGER", "None", "N", "> 0", "Total quantity of units ordered", "OQ-01, BR-01"],
            ["delivered_quantity", "INTEGER", "None", "N", ">= 0", "Total physical units received at warehouse dock", "OQ-01, BR-01"],
            ["actual_delivery_date", "DATE", "None", "Y", "YYYY-MM-DD", "Official physical dock receiving timestamp", "OQ-01, OQ-02, BR-01, BR-02"],
            ["delivery_lead_time_days", "INTEGER", "None", "Y", "actual_delivery_date - order_date", "Elapsed calendar days to fulfill delivery", "OQ-02, BR-02"],
            ["delay_days", "INTEGER", "None", "N", "Default 0; >= 0 if late", "Elapsed delay past promised delivery date", "OQ-05, BR-05"],
            ["shipping_mode", "VARCHAR(15)", "None", "N", "Air Freight, Sea Ocean, Road Express, Rail Cargo", "Freight transport transit mode", "OQ-10, BR-10"],
            ["contracted_unit_price", "DECIMAL(10,2)", "None", "N", "> 0.00", "Baseline contracted price at order time", "OQ-04, BR-04"],
            ["invoiced_unit_price", "DECIMAL(10,2)", "None", "N", "> 0.00", "Actual invoiced billing unit price", "OQ-04, BR-04"],
            ["price_variance_amount", "DECIMAL(10,2)", "None", "N", "invoiced_unit_price - contracted_unit_price", "Unit monetary cost delta", "OQ-04, BR-04"],
            ["price_variance_pct", "DECIMAL(6,2)", "None", "N", "(variance_amount / contracted_price) * 100", "Percentage purchase price variance", "OQ-04, BR-04"],
            ["inspected_quantity", "INTEGER", "None", "N", ">= 0; Default 0", "Number of units sampled for QA inspection", "OQ-03, BR-03"],
            ["defective_quantity", "INTEGER", "None", "N", ">= 0 and <= inspected_quantity", "Number of units failing QA inspection", "OQ-03, BR-03"],
            ["defect_rate_pct", "DECIMAL(6,2)", "None", "N", "(defective / inspected) * 100", "Percentage of defective units in batch", "OQ-03, BR-03"],
            ["defect_ppm", "INTEGER", "None", "N", "defect_rate_pct * 10,000", "Defective parts per million benchmark", "OQ-03, BR-03"],
            ["defect_reason_code", "VARCHAR(30)", "None", "N", "Dimensional Non-Conformance, Surface, Fatigue, Packaging, Electrical, None", "Root cause defect failure categorization", "OQ-03, BR-03"],
            ["is_on_time", "INTEGER", "None", "N", "0 or 1", "Boolean: actual_date <= promised_date", "OQ-01, BR-01"],
            ["is_in_full", "INTEGER", "None", "N", "0 or 1", "Boolean: delivered_qty >= ordered_qty", "OQ-01, BR-01"],
            ["is_otif", "INTEGER", "None", "N", "0 or 1", "Boolean: is_on_time == 1 AND is_in_full == 1", "OQ-01, BR-01"]
        ],
        {"A": 24, "B": 16, "C": 10, "D": 10, "E": 28, "F": 45, "G": 22}
    )
}

for sheet_title, (headers, rows, widths) in dd_specs.items():
    ws_d = wb.create_sheet(title=sheet_title)
    ws_d.views.sheetView[0].showGridLines = True
    style_table(ws_d, 1, headers, rows, widths, center_cols=[3, 4])

# -------------------------------------------------------------
# Sheet 8: Relationships
# -------------------------------------------------------------
ws_rel = wb.create_sheet(title="Relationships")
ws_rel.views.sheetView[0].showGridLines = True

rel_headers = ["Relationship ID", "Parent Table", "Parent Column (PK)", "Child Table", "Child Column (FK)", "Cardinality", "On Delete", "On Update", "Business Rule Enforced"]
rel_data = [
    ["REL-01", "suppliers", "supplier_id", "purchase_orders", "supplier_id", "1 : Many (1:N)", "RESTRICT", "CASCADE", "A purchase order must belong to an existing registered master supplier."],
    ["REL-02", "categories", "category_id", "products", "category_id", "1 : Many (1:N)", "RESTRICT", "CASCADE", "Every product item must be cataloged under an authorized procurement category."],
    ["REL-03", "purchase_orders", "order_id", "order_deliveries", "order_id", "1 : Many (1:N)", "CASCADE", "CASCADE", "Order delivery lines cannot exist without a valid parent purchase order."],
    ["REL-04", "products", "product_id", "order_deliveries", "product_id", "1 : Many (1:N)", "RESTRICT", "CASCADE", "Every delivered line item must map to a certified catalog SKU."]
]
rel_widths = {"A": 16, "B": 18, "C": 18, "D": 20, "E": 18, "F": 16, "G": 14, "H": 14, "I": 45}
style_table(ws_rel, 1, rel_headers, rel_data, rel_widths, center_cols=[1, 6, 7, 8])

# -------------------------------------------------------------
# Sheet 9: KPIs_Measures
# -------------------------------------------------------------
ws_kpi = wb.create_sheet(title="KPIs_Measures")
ws_kpi.views.sheetView[0].showGridLines = True

kpi_headers = ["Measure Name", "Business Formula", "SQL / Algorithmic Expression", "Target Benchmark", "Linked OQ / BR"]
kpi_data = [
    ["On-Time In-Full (OTIF) Rate", "(Count of OTIF Orders / Total Completed Orders) * 100", "SUM(is_otif) * 100.0 / COUNT(*)", ">= 95.0% for Tier A", "OQ-01, BR-01"],
    ["Average Lead Time", "Sum of Lead Time Days / Total Completed Orders", "AVG(delivery_lead_time_days)", "+/- 2 days of SLA", "OQ-02, BR-02"],
    ["Lead Time Standard Deviation", "Square root of variance in delivery lead time", "STDEV(delivery_lead_time_days)", "< 3.0 Days", "OQ-02, BR-02"],
    ["Defective Parts Per Million (PPM)", "(Total Defective Units / Total Inspected Units) * 1,000,000", "(SUM(defective_qty) * 1000000.0) / SUM(inspected_qty)", "< 1,000 PPM", "OQ-03, BR-03"],
    ["Defect Rate Percentage", "(Total Defective Units / Total Inspected Units) * 100", "(SUM(defective_qty) * 100.0) / SUM(inspected_qty)", "< 1.5%", "OQ-03, BR-03"],
    ["Purchase Price Variance (PPV) %", "((Invoiced Price - Contract Price) / Contract Price) * 100", "AVG((invoiced_unit_price - contracted_unit_price) * 100.0 / contracted_unit_price)", "<= 0.0%", "OQ-04, BR-04"],
    ["Tail Delay Outlier Count", "Count of delivery lines where delay_days > 7 days", "COUNT(CASE WHEN delay_days > 7 THEN 1 END)", "0 Severe Delays", "OQ-05, BR-05"],
    ["Composite Performance Index (CPI)", "0.40 * OTIF% + 0.35 * (100 - Defect%) + 0.25 * (100 - CostVar%)", "(0.40 * otif_pct) + (0.35 * quality_score) + (0.25 * cost_score)", ">= 85.0 (Tier A)", "OQ-06, BR-06"],
    ["Spend-at-Risk Ratio", "(Total Spend with Tier C Vendors / Total Spend) * 100", "SUM(CASE WHEN supplier_tier = 'Tier C' THEN total_po_value END) * 100.0 / SUM(total_po_value)", "< 10.0%", "OQ-07, BR-07"]
]
kpi_widths = {"A": 28, "B": 42, "C": 48, "D": 22, "E": 20}
style_table(ws_kpi, 1, kpi_headers, kpi_data, kpi_widths, center_cols=[4, 5])

# -------------------------------------------------------------
# Sheet 10: SQLite_DDL
# -------------------------------------------------------------
ws_ddl = wb.create_sheet(title="SQLite_DDL")
ws_ddl.views.sheetView[0].showGridLines = True

ddl_script = """-- ====================================================================
-- SUPPLIER PERFORMANCE ANALYSIS PLATFORM - RELATIONAL SQLITE SCHEMA
-- Database Version: 1.0 (Production Release)
-- Target Platform: SQLite 3.35+
-- ====================================================================

PRAGMA foreign_keys = ON;

-- 1. Master Suppliers Table
CREATE TABLE suppliers (
    supplier_id VARCHAR(10) PRIMARY KEY,
    supplier_name VARCHAR(100) NOT NULL UNIQUE,
    supplier_country VARCHAR(50) NOT NULL CHECK(supplier_country IN ('India', 'Germany', 'USA', 'Japan', 'South Korea', 'Vietnam', 'Taiwan')),
    supplier_city VARCHAR(50) NOT NULL,
    contact_email VARCHAR(100) NOT NULL,
    supplier_tier VARCHAR(10) NOT NULL CHECK(supplier_tier IN ('Tier A', 'Tier B', 'Tier C')),
    rating_score DECIMAL(4,2) NOT NULL CHECK(rating_score BETWEEN 0.00 AND 100.00),
    contract_start_date DATE NOT NULL,
    contract_end_date DATE NOT NULL,
    status VARCHAR(20) NOT NULL CHECK(status IN ('Active', 'Conditional', 'Suspended')),
    CHECK(contract_end_date >= contract_start_date)
);

-- 2. Master Categories Table
CREATE TABLE categories (
    category_id VARCHAR(10) PRIMARY KEY,
    category_name VARCHAR(50) NOT NULL UNIQUE,
    spend_tier VARCHAR(10) NOT NULL CHECK(spend_tier IN ('Strategic', 'Leverage', 'Bottleneck', 'Routine')),
    risk_weight DECIMAL(3,2) NOT NULL CHECK(risk_weight BETWEEN 0.10 AND 1.00),
    description TEXT
);

-- 3. Master Products Catalog Table
CREATE TABLE products (
    product_id VARCHAR(10) PRIMARY KEY,
    product_name VARCHAR(100) NOT NULL,
    category_id VARCHAR(10) NOT NULL,
    unit_of_measure VARCHAR(15) NOT NULL CHECK(unit_of_measure IN ('Pieces', 'Kg', 'Meters', 'Sets', 'Boxes')),
    baseline_contract_price DECIMAL(10,2) NOT NULL CHECK(baseline_contract_price > 0),
    standard_lead_time_days INTEGER NOT NULL CHECK(standard_lead_time_days BETWEEN 3 AND 90),
    target_defect_ppm INTEGER NOT NULL CHECK(target_defect_ppm >= 0),
    FOREIGN KEY (category_id) REFERENCES categories(category_id) ON UPDATE CASCADE ON DELETE RESTRICT
);

-- 4. Fact Purchase Orders Table
CREATE TABLE purchase_orders (
    order_id VARCHAR(15) PRIMARY KEY,
    po_number VARCHAR(20) NOT NULL UNIQUE,
    supplier_id VARCHAR(10) NOT NULL,
    order_date DATE NOT NULL,
    promised_delivery_date DATE NOT NULL,
    order_status VARCHAR(20) NOT NULL CHECK(order_status IN ('Completed', 'In-Transit', 'Cancelled')),
    payment_terms VARCHAR(20) NOT NULL CHECK(payment_terms IN ('Net 30', 'Net 45', 'Net 60', 'Immediate')),
    total_po_value DECIMAL(12,2) NOT NULL CHECK(total_po_value >= 0),
    fiscal_year INTEGER NOT NULL CHECK(fiscal_year BETWEEN 2023 AND 2026),
    fiscal_quarter VARCHAR(5) NOT NULL CHECK(fiscal_quarter IN ('Q1', 'Q2', 'Q3', 'Q4')),
    FOREIGN KEY (supplier_id) REFERENCES suppliers(supplier_id) ON UPDATE CASCADE ON DELETE RESTRICT,
    CHECK(promised_delivery_date >= order_date)
);

-- 5. Fact Order Deliveries & QA Inspection Table
CREATE TABLE order_deliveries (
    delivery_id VARCHAR(15) PRIMARY KEY,
    order_id VARCHAR(15) NOT NULL,
    product_id VARCHAR(10) NOT NULL,
    ordered_quantity INTEGER NOT NULL CHECK(ordered_quantity > 0),
    delivered_quantity INTEGER NOT NULL CHECK(delivered_quantity >= 0),
    actual_delivery_date DATE,
    delivery_lead_time_days INTEGER,
    delay_days INTEGER NOT NULL DEFAULT 0,
    shipping_mode VARCHAR(15) NOT NULL CHECK(shipping_mode IN ('Air Freight', 'Sea Ocean', 'Road Express', 'Rail Cargo')),
    contracted_unit_price DECIMAL(10,2) NOT NULL CHECK(contracted_unit_price > 0),
    invoiced_unit_price DECIMAL(10,2) NOT NULL CHECK(invoiced_unit_price > 0),
    price_variance_amount DECIMAL(10,2) NOT NULL,
    price_variance_pct DECIMAL(6,2) NOT NULL,
    inspected_quantity INTEGER NOT NULL DEFAULT 0 CHECK(inspected_quantity >= 0),
    defective_quantity INTEGER NOT NULL DEFAULT 0 CHECK(defective_quantity >= 0),
    defect_rate_pct DECIMAL(6,2) NOT NULL DEFAULT 0.0,
    defect_ppm INTEGER NOT NULL DEFAULT 0,
    defect_reason_code VARCHAR(30) NOT NULL CHECK(defect_reason_code IN ('Dimensional Non-Conformance', 'Surface Imperfection', 'Material Fatigue', 'Packaging Damage', 'Electrical Fault', 'None')),
    is_on_time INTEGER NOT NULL CHECK(is_on_time IN (0, 1)),
    is_in_full INTEGER NOT NULL CHECK(is_in_full IN (0, 1)),
    is_otif INTEGER NOT NULL CHECK(is_otif IN (0, 1)),
    FOREIGN KEY (order_id) REFERENCES purchase_orders(order_id) ON UPDATE CASCADE ON DELETE CASCADE,
    FOREIGN KEY (product_id) REFERENCES products(product_id) ON UPDATE CASCADE ON DELETE RESTRICT,
    CHECK(defective_quantity <= inspected_quantity)
);

-- High-Performance Indexes
CREATE INDEX idx_po_supplier ON purchase_orders(supplier_id);
CREATE INDEX idx_po_order_date ON purchase_orders(order_date);
CREATE INDEX idx_del_order ON order_deliveries(order_id);
CREATE INDEX idx_del_product ON order_deliveries(product_id);
CREATE INDEX idx_del_actual_date ON order_deliveries(actual_delivery_date);

-- ====================================================================
-- 6. Unified Analytical Joined View (41 Columns)
-- ====================================================================
CREATE VIEW v_supplier_performance_analysis AS
SELECT 
    d.delivery_id,
    po.order_id,
    po.po_number,
    po.order_date,
    po.promised_delivery_date,
    d.actual_delivery_date,
    po.fiscal_year,
    po.fiscal_quarter,
    po.order_status,
    s.supplier_id,
    s.supplier_name,
    s.supplier_country,
    s.supplier_city,
    s.supplier_tier,
    s.status AS supplier_status,
    c.category_id,
    c.category_name,
    c.spend_tier,
    p.product_id,
    p.product_name,
    p.unit_of_measure,
    d.ordered_quantity,
    d.delivered_quantity,
    d.delivery_lead_time_days,
    p.standard_lead_time_days,
    (d.delivery_lead_time_days - p.standard_lead_time_days) AS lead_time_variance_days,
    d.delay_days,
    d.shipping_mode,
    d.contracted_unit_price,
    d.invoiced_unit_price,
    (d.ordered_quantity * d.contracted_unit_price) AS line_total_contracted_spend,
    (d.delivered_quantity * d.invoiced_unit_price) AS line_total_invoiced_spend,
    d.price_variance_amount,
    d.price_variance_pct,
    d.inspected_quantity,
    d.defective_quantity,
    d.defect_rate_pct,
    d.defect_ppm,
    d.defect_reason_code,
    d.is_on_time,
    d.is_in_full,
    d.is_otif
FROM order_deliveries d
JOIN purchase_orders po ON d.order_id = po.order_id
JOIN suppliers s ON po.supplier_id = s.supplier_id
JOIN products p ON d.product_id = p.product_id
JOIN categories c ON p.category_id = c.category_id;
"""

ws_ddl.cell(row=1, column=1, value="Executable SQLite DDL Specification").font = Font(name="Calibri", size=12, bold=True, color=NAVY_HEADER)
ws_ddl.row_dimensions[1].height = 24

for line_idx, line in enumerate(ddl_script.strip().split("\n"), start=3):
    c = ws_ddl.cell(row=line_idx, column=1, value=line)
    c.font = Font(name="Courier New", size=9)
    ws_ddl.row_dimensions[line_idx].height = 16

ws_ddl.column_dimensions["A"].width = 120

output_db_design = os.path.join('docs', '08_Database_Design.xlsx')
wb.save(output_db_design)
print(f'Successfully generated: {output_db_design}')
