import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

output_path = os.path.join('docs', '04_Traceability_Matrix.xlsx')
wb = openpyxl.load_workbook(output_path)

ws_rtm = wb["RTM_Matrix"]

# Update mappings
# Col 8 is Feature (Step 6/7)
# Col 9 is DB table.column (Step 8)
# Col 7 is Status -> Update to "Schema Designed"

updates = [
    # Row 2 (index 2)
    (2, "FR-01", "order_deliveries.is_otif, purchase_orders.promised_delivery_date, order_deliveries.actual_delivery_date, order_deliveries.ordered_quantity, order_deliveries.delivered_quantity", "Schema Designed"),
    # Row 3
    (3, "FR-11", "categories.category_id, suppliers.supplier_id, purchase_orders.order_date, purchase_orders.order_status", "Schema Designed"),
    # Row 4
    (4, "FR-02", "order_deliveries.delivery_lead_time_days, products.standard_lead_time_days, suppliers.supplier_id, categories.category_id", "Schema Designed"),
    # Row 5
    (5, "FR-11", "categories.category_id, suppliers.supplier_id, purchase_orders.order_date, purchase_orders.order_status", "Schema Designed"),
    # Row 6
    (6, "FR-03", "order_deliveries.inspected_quantity, order_deliveries.defective_quantity, order_deliveries.defect_rate_pct, order_deliveries.defect_ppm, order_deliveries.defect_reason_code", "Schema Designed"),
    # Row 7
    (7, "FR-11", "categories.category_id, suppliers.supplier_id, purchase_orders.order_date, purchase_orders.order_status", "Schema Designed"),
    # Row 8
    (8, "FR-04", "order_deliveries.contracted_unit_price, order_deliveries.invoiced_unit_price, order_deliveries.price_variance_amount, order_deliveries.price_variance_pct", "Schema Designed"),
    # Row 9
    (9, "FR-05", "order_deliveries.delay_days, order_deliveries.is_on_time, purchase_orders.order_status", "Schema Designed"),
    # Row 10
    (10, "FR-06", "suppliers.rating_score, suppliers.supplier_tier, order_deliveries.is_otif, order_deliveries.defect_rate_pct, order_deliveries.price_variance_pct", "Schema Designed"),
    # Row 11
    (11, "FR-12", "v_supplier_performance_analysis.*, suppliers.supplier_id, suppliers.supplier_name", "Schema Designed"),
    # Row 12
    (12, "FR-07", "purchase_orders.total_po_value, suppliers.supplier_tier, order_deliveries.invoiced_unit_price", "Schema Designed"),
    # Row 13
    (13, "FR-08", "purchase_orders.order_date, purchase_orders.fiscal_year, purchase_orders.fiscal_quarter, order_deliveries.actual_delivery_date", "Schema Designed"),
    # Row 14
    (14, "FR-09", "suppliers.supplier_tier, suppliers.contract_end_date, suppliers.status, order_deliveries.delay_days", "Schema Designed"),
    # Row 15
    (15, "FR-12", "v_supplier_performance_analysis.*, suppliers.supplier_id, suppliers.supplier_name", "Schema Designed"),
    # Row 16
    (16, "FR-10", "order_deliveries.shipping_mode, suppliers.supplier_country, suppliers.supplier_city, order_deliveries.delivery_lead_time_days, order_deliveries.delay_days", "Schema Designed"),
]

BORDER_GRAY = "D1D5DB"
thin_border = Border(
    left=Side(style='thin', color=BORDER_GRAY),
    right=Side(style='thin', color=BORDER_GRAY),
    top=Side(style='thin', color=BORDER_GRAY),
    bottom=Side(style='thin', color=BORDER_GRAY)
)

for row_num, feature, db_col, status in updates:
    cell_status = ws_rtm.cell(row=row_num, column=7)
    cell_status.value = status
    cell_status.font = Font(name="Calibri", size=10, bold=True, color="059669")
    
    cell_feat = ws_rtm.cell(row=row_num, column=8)
    cell_feat.value = feature
    cell_feat.alignment = Alignment(horizontal="center", vertical="center")
    cell_feat.font = Font(name="Calibri", size=10, bold=True, color="4F46E5")
    cell_feat.border = thin_border
    
    cell_db = ws_rtm.cell(row=row_num, column=9)
    cell_db.value = db_col
    cell_db.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    cell_db.font = Font(name="Calibri", size=9)
    cell_db.border = thin_border

ws_rtm.column_dimensions["I"].width = 52
ws_rtm.column_dimensions["H"].width = 16

wb.save(output_path)
print(f'Successfully updated: {output_path}')
