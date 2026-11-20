import os
from PIL import Image, ImageDraw, ImageFont

os.makedirs('docs', exist_ok=True)

# Image dimensions
width = 1600
height = 1150
img = Image.new('RGB', (width, height), color='#F8FAFC')
draw = ImageDraw.Draw(img)

# Colors
NAVY = '#1A365D'
WHITE = '#FFFFFF'
CARD_BG = '#FFFFFF'
BORDER = '#CBD5E1'
HEADER_BG = '#1E293B'
PK_COLOR = '#7C3AED'
FK_COLOR = '#2563EB'
TEXT_DARK = '#0F172A'
TEXT_MUTED = '#64748B'
LINE_COLOR = '#334155'
LABEL_BG = '#E2E8F0'

# Fonts - using basic default or load PIL default
try:
    font_title = ImageFont.truetype("arial.ttf", 26)
    font_subtitle = ImageFont.truetype("arial.ttf", 15)
    font_tbl_hdr = ImageFont.truetype("arialbd.ttf", 15)
    font_col = ImageFont.truetype("arial.ttf", 12)
    font_badge = ImageFont.truetype("arialbd.ttf", 11)
    font_rel = ImageFont.truetype("arialbd.ttf", 12)
except:
    font_title = ImageFont.load_default()
    font_subtitle = ImageFont.load_default()
    font_tbl_hdr = ImageFont.load_default()
    font_col = ImageFont.load_default()
    font_badge = ImageFont.load_default()
    font_rel = ImageFont.load_default()

# Draw Canvas Title
draw.rectangle([(0, 0), (width, 80)], fill=NAVY)
draw.text((40, 16), "SUPPLIER PERFORMANCE ANALYSIS — ENTITY RELATIONSHIP DIAGRAM (ERD)", fill=WHITE, font=font_title)
draw.text((40, 50), "Normalized 3NF Relational Architecture | Primary Keys (PK) & Foreign Keys (FK) Mappings", fill='#94A3B8', font=font_subtitle)

# Helper function to draw an entity card
def draw_table(x, y, w, title, subtitle, columns):
    header_h = 42
    row_h = 22
    total_h = header_h + len(columns) * row_h + 12
    
    # Outer box shadow / border
    draw.rounded_rectangle([(x+3, y+3), (x+w+3, y+total_h+3)], radius=8, fill='#E2E8F0')
    draw.rounded_rectangle([(x, y), (x+w, y+total_h)], radius=8, fill=CARD_BG, outline=BORDER, width=2)
    
    # Header
    draw.rounded_rectangle([(x, y), (x+w, y+header_h)], radius=8, fill=HEADER_BG)
    # Fill bottom corners of header so it looks square at the bottom
    draw.rectangle([(x, y+header_h-8), (x+w, y+header_h)], fill=HEADER_BG)
    draw.text((x+14, y+8), title, fill=WHITE, font=font_tbl_hdr)
    draw.text((x+14, y+24), subtitle, fill='#94A3B8', font=font_col)
    
    # Columns
    cur_y = y + header_h + 6
    for col_name, col_type, key_type in columns:
        # Background zebra
        # draw text
        if key_type == 'PK':
            draw.rounded_rectangle([(x+10, cur_y+2), (x+34, cur_y+18)], radius=3, fill='#EDE9FE')
            draw.text((x+14, cur_y+3), "PK", fill=PK_COLOR, font=font_badge)
        elif key_type == 'FK':
            draw.rounded_rectangle([(x+10, cur_y+2), (x+34, cur_y+18)], radius=3, fill='#DBEAFE')
            draw.text((x+14, cur_y+3), "FK", fill=FK_COLOR, font=font_badge)
        else:
            draw.ellipse([(x+18, cur_y+8), (x+24, cur_y+14)], fill='#CBD5E1')
            
        draw.text((x+42, cur_y+3), col_name, fill=TEXT_DARK, font=font_col)
        # Type right aligned
        draw.text((x+w-120, cur_y+3), col_type, fill=TEXT_MUTED, font=font_col)
        cur_y += row_h
        
    return x, y, w, total_h

# 1. suppliers (Top Left)
suppliers_cols = [
    ('supplier_id', 'VARCHAR(10)', 'PK'),
    ('supplier_name', 'VARCHAR(100)', ''),
    ('supplier_country', 'VARCHAR(50)', ''),
    ('supplier_city', 'VARCHAR(50)', ''),
    ('contact_email', 'VARCHAR(100)', ''),
    ('supplier_tier', 'VARCHAR(10)', ''),
    ('rating_score', 'DECIMAL(4,2)', ''),
    ('contract_start_date', 'DATE', ''),
    ('contract_end_date', 'DATE', ''),
    ('status', 'VARCHAR(20)', '')
]
t1 = draw_table(50, 110, 320, "suppliers", "Master Vendor Dimension (25 rows)", suppliers_cols)

# 2. categories (Bottom Left)
categories_cols = [
    ('category_id', 'VARCHAR(10)', 'PK'),
    ('category_name', 'VARCHAR(50)', ''),
    ('spend_tier', 'VARCHAR(10)', ''),
    ('risk_weight', 'DECIMAL(3,2)', ''),
    ('description', 'TEXT', '')
]
t2 = draw_table(50, 480, 320, "categories", "Commodity Taxonomy (6 rows)", categories_cols)

# 3. products (Middle Bottom)
products_cols = [
    ('product_id', 'VARCHAR(10)', 'PK'),
    ('product_name', 'VARCHAR(100)', ''),
    ('category_id', 'VARCHAR(10)', 'FK'),
    ('unit_of_measure', 'VARCHAR(15)', ''),
    ('baseline_contract_price', 'DECIMAL(10,2)', ''),
    ('standard_lead_time_days', 'INTEGER', ''),
    ('target_defect_ppm', 'INTEGER', '')
]
t3 = draw_table(50, 750, 340, "products", "Item & SKU Catalog (40 rows)", products_cols)

# 4. purchase_orders (Top Center)
po_cols = [
    ('order_id', 'VARCHAR(15)', 'PK'),
    ('po_number', 'VARCHAR(20)', ''),
    ('supplier_id', 'VARCHAR(10)', 'FK'),
    ('order_date', 'DATE', ''),
    ('promised_delivery_date', 'DATE', ''),
    ('order_status', 'VARCHAR(20)', ''),
    ('payment_terms', 'VARCHAR(20)', ''),
    ('total_po_value', 'DECIMAL(12,2)', ''),
    ('fiscal_year', 'INTEGER', ''),
    ('fiscal_quarter', 'VARCHAR(5)', '')
]
t4 = draw_table(550, 110, 340, "purchase_orders", "PO Transactions Fact (300 rows)", po_cols)

# 5. order_deliveries (Center-Right, Large Bridge Fact)
deliv_cols = [
    ('delivery_id', 'VARCHAR(15)', 'PK'),
    ('order_id', 'VARCHAR(15)', 'FK'),
    ('product_id', 'VARCHAR(10)', 'FK'),
    ('ordered_quantity', 'INTEGER', ''),
    ('delivered_quantity', 'INTEGER', ''),
    ('actual_delivery_date', 'DATE', ''),
    ('delivery_lead_time_days', 'INTEGER', ''),
    ('delay_days', 'INTEGER', ''),
    ('shipping_mode', 'VARCHAR(15)', ''),
    ('contracted_unit_price', 'DECIMAL(10,2)', ''),
    ('invoiced_unit_price', 'DECIMAL(10,2)', ''),
    ('price_variance_amount', 'DECIMAL(10,2)', ''),
    ('price_variance_pct', 'DECIMAL(6,2)', ''),
    ('inspected_quantity', 'INTEGER', ''),
    ('defective_quantity', 'INTEGER', ''),
    ('defect_rate_pct', 'DECIMAL(6,2)', ''),
    ('defect_ppm', 'INTEGER', ''),
    ('defect_reason_code', 'VARCHAR(30)', ''),
    ('is_on_time', 'INTEGER', ''),
    ('is_in_full', 'INTEGER', ''),
    ('is_otif', 'INTEGER', '')
]
t5 = draw_table(1080, 110, 440, "order_deliveries", "Fulfillment & QA Fact Bridge (300-400 rows)", deliv_cols)

# Helper function to draw relationship lines with crow's foot
def draw_relationship(p1, p2, label, offset_label=(0, 0)):
    # Draw line segments
    draw.line([p1, p2], fill=LINE_COLOR, width=2)
    # Relationship verb badge
    mid_x = (p1[0] + p2[0]) // 2 + offset_label[0]
    mid_y = (p1[1] + p2[1]) // 2 + offset_label[1]
    w_box = len(label) * 8 + 16
    draw.rounded_rectangle([(mid_x - w_box//2, mid_y - 12), (mid_x + w_box//2, mid_y + 12)], radius=5, fill=LABEL_BG, outline=BORDER)
    draw.text((mid_x - w_box//2 + 8, mid_y - 7), label, fill=LINE_COLOR, font=font_rel)

# 1. suppliers -> purchase_orders (REL-01: 1 to N)
draw.line([(370, 200), (550, 200)], fill=LINE_COLOR, width=3)
# Crow's foot at (550, 200)
draw.line([(535, 190), (550, 200)], fill=LINE_COLOR, width=2)
draw.line([(535, 210), (550, 200)], fill=LINE_COLOR, width=2)
# One bar at (385, 200)
draw.line([(385, 190), (385, 210)], fill=LINE_COLOR, width=2)
draw_relationship((370, 200), (550, 200), "fulfills (1:N)", (0, -18))

# 2. categories -> products (REL-02: 1 to N)
draw.line([(210, 650), (210, 750)], fill=LINE_COLOR, width=3)
draw.line([(200, 735), (210, 750)], fill=LINE_COLOR, width=2)
draw.line([(220, 735), (210, 750)], fill=LINE_COLOR, width=2)
draw.line([(200, 665), (220, 665)], fill=LINE_COLOR, width=2)
draw_relationship((210, 650), (210, 750), "classifies (1:N)", (50, 0))

# 3. purchase_orders -> order_deliveries (REL-03: 1 to N)
draw.line([(890, 250), (1080, 250)], fill=LINE_COLOR, width=3)
draw.line([(1065, 240), (1080, 250)], fill=LINE_COLOR, width=2)
draw.line([(1065, 260), (1080, 250)], fill=LINE_COLOR, width=2)
draw.line([(905, 240), (905, 260)], fill=LINE_COLOR, width=2)
draw_relationship((890, 250), (1080, 250), "contains (1:N)", (0, -18))

# 4. products -> order_deliveries (REL-04: 1 to N)
# Draw stepped line: from (390, 850) -> right to (980, 850) -> up to (980, 480) -> right to (1080, 480)
draw.line([(390, 850), (980, 850)], fill=LINE_COLOR, width=3)
draw.line([(980, 850), (980, 480)], fill=LINE_COLOR, width=3)
draw.line([(980, 480), (1080, 480)], fill=LINE_COLOR, width=3)
draw.line([(1065, 470), (1080, 480)], fill=LINE_COLOR, width=2)
draw.line([(1065, 490), (1080, 480)], fill=LINE_COLOR, width=2)
draw.line([(405, 840), (405, 860)], fill=LINE_COLOR, width=2)
draw_relationship((700, 850), (700, 850), "itemizes (1:N)", (0, 0))

# Save image
output_png = os.path.join('docs', '09_ER_Diagram.png')
img.save(output_png, quality=95)
print(f'Successfully generated: {output_png}')
