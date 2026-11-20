import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

os.makedirs('docs', exist_ok=True)
wb = openpyxl.Workbook()

# Sheet 1: Traceability Matrix
ws_rtm = wb.active
ws_rtm.title = "RTM_Matrix"
ws_rtm.views.sheetView[0].showGridLines = True

# Palettes
NAVY_HEADER = "1A365D"
WHITE = "FFFFFF"
LIGHT_BLUE = "F0F4F8"
ZEBRA_FILL = "F9FAFB"
BORDER_GRAY = "D1D5DB"
SUCCESS_GREEN = "10B981"
MUST_COLOR = "DC2626"
SHOULD_COLOR = "D97706"
COULD_COLOR = "2563EB"

header_fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
header_font = Font(name="Calibri", size=11, bold=True, color=WHITE)
header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

thin_border = Border(
    left=Side(style='thin', color=BORDER_GRAY),
    right=Side(style='thin', color=BORDER_GRAY),
    top=Side(style='thin', color=BORDER_GRAY),
    bottom=Side(style='thin', color=BORDER_GRAY)
)

headers = [
    "OQ ID",
    "HMW ID",
    "Req ID (BR)",
    "Business Requirement Description",
    "Stakeholder (SH)",
    "Priority",
    "Status",
    "Feature (Step 6)",
    "DB table.column (Step 8)",
    "User Story (Step 11)",
    "UI Screen (Step 12)",
    "Test Case (Step 13)"
]

ws_rtm.append(headers)
ws_rtm.row_dimensions[1].height = 28

for col_num in range(1, len(headers) + 1):
    cell = ws_rtm.cell(row=1, column=col_num)
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = header_align
    cell.border = thin_border

matrix_data = [
    ("OQ-01", "HMW-01", "BR-01", "The business needs to track and benchmark vendor On-Time In-Full (OTIF) fulfillment compliance rates across all purchase orders so that procurement teams can allocate orders exclusively to dependable suppliers and eliminate production halts.", "SH-02, SH-01", "Must", "Identified", "", "", "", "", ""),
    ("OQ-01", "HMW-01", "BR-11", "The business needs to filter and segment procurement performance by product category, date range, supplier, and order status so that cross-functional teams can conduct targeted root-cause investigations.", "SH-01, SH-02, SH-03, SH-04, SH-05", "Must", "Identified", "", "", "", "", ""),
    ("OQ-02", "HMW-02", "BR-02", "The business needs to compute average lead times, standard deviations, and SLA variance per supplier and product category so that warehouse operations can calibrate safety stock levels without inflating inventory holding costs.", "SH-04, SH-02", "Must", "Identified", "", "", "", "", ""),
    ("OQ-02", "HMW-02", "BR-11", "The business needs to filter and segment procurement performance by product category, date range, supplier, and order status so that cross-functional teams can conduct targeted root-cause investigations.", "SH-01, SH-02, SH-03, SH-04, SH-05", "Must", "Identified", "", "", "", "", ""),
    ("OQ-03", "HMW-03", "BR-03", "The business needs to monitor incoming component defect rates and Defective Parts Per Million (PPM) at the supplier and part level so that quality managers can prevent sub-standard materials from entering assembly processes.", "SH-03", "Must", "Identified", "", "", "", "", ""),
    ("OQ-03", "HMW-03", "BR-11", "The business needs to filter and segment procurement performance by product category, date range, supplier, and order status so that cross-functional teams can conduct targeted root-cause investigations.", "SH-01, SH-02, SH-03, SH-04, SH-05", "Must", "Identified", "", "", "", "", ""),
    ("OQ-04", "HMW-04", "BR-04", "The business needs to audit unit purchase price variances between contracted baselines and actual invoiced costs across all order line items so that finance controllers can prevent cost leakage and recover unauthorized price increments.", "SH-05", "Must", "Identified", "", "", "", "", ""),
    ("OQ-05", "HMW-05", "BR-05", "The business needs to isolate supplier delivery outliers and tail delays exceeding seven business days so that sourcing teams can remove highly volatile vendors from critical path manufacturing schedules.", "SH-02, SH-01", "Should", "Identified", "", "", "", "", ""),
    ("OQ-06", "HMW-06", "BR-06", "The business needs to synthesize delivery timeliness, quality compliance, and price adherence into a weighted Composite Performance Index (CPI) ranking so that management can objectively categorize suppliers into standardized governance tiers.", "SH-01", "Must", "Identified", "", "", "", "", ""),
    ("OQ-06", "HMW-06", "BR-12", "The business needs to export standardized supplier performance scorecards and audit summaries into structured reports so that procurement teams can present empirical data during formal vendor contract reviews.", "SH-01, SH-02", "Could", "Identified", "", "", "", "", ""),
    ("OQ-07", "HMW-07", "BR-07", "The business needs to quantify the total procurement expenditure allocated to high-risk and underperforming suppliers so that leadership can mitigate supply chain vulnerabilities and reallocate capital toward reliable vendors.", "SH-01, SH-05", "Must", "Identified", "", "", "", "", ""),
    ("OQ-08", "HMW-08", "BR-08", "The business needs to analyze chronological and quarterly trends in vendor fulfillment and quality metrics so that category managers can anticipate seasonal slumps and secure capacity ahead of peak operational cycles.", "SH-01, SH-02", "Should", "Identified", "", "", "", "", ""),
    ("OQ-09", "HMW-09", "BR-09", "The business needs to identify suppliers eligible for strategic contract renewal or requiring mandatory Corrective Action Plans (CAPA) based on objective threshold criteria so that vendor governance is enforced systematically and legally.", "SH-01, SH-03", "Must", "Identified", "", "", "", "", ""),
    ("OQ-09", "HMW-09", "BR-12", "The business needs to export standardized supplier performance scorecards and audit summaries into structured reports so that procurement teams can present empirical data during formal vendor contract reviews.", "SH-01, SH-02", "Could", "Identified", "", "", "", "", ""),
    ("OQ-10", "HMW-10", "BR-10", "The business needs to evaluate vendor delivery lead times and transit delays across shipping modes and geographic vendor origins so that logistics planners can select the most predictable transport routes and origin locations.", "SH-04, SH-02", "Should", "Identified", "", "", "", "", "")
]

row_idx = 2
for row_data in matrix_data:
    ws_rtm.append(row_data)
    ws_rtm.row_dimensions[row_idx].height = 36
    zebra = PatternFill(start_color=ZEBRA_FILL if row_idx % 2 == 0 else "FFFFFF", fill_type="solid")
    
    for c_idx in range(1, len(row_data) + 1):
        cell = ws_rtm.cell(row=row_idx, column=c_idx)
        cell.font = Font(name="Calibri", size=10)
        cell.fill = zebra
        cell.border = thin_border
        
        # Center align IDs, priority, status
        if c_idx in [1, 2, 3, 6, 7]:
            cell.alignment = Alignment(horizontal="center", vertical="center")
        elif c_idx == 4:
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        else:
            cell.alignment = Alignment(horizontal="left", vertical="center")
            
        # Priority font styling
        if c_idx == 6:
            if cell.value == "Must":
                cell.font = Font(name="Calibri", size=10, bold=True, color=MUST_COLOR)
            elif cell.value == "Should":
                cell.font = Font(name="Calibri", size=10, bold=True, color=SHOULD_COLOR)
            elif cell.value == "Could":
                cell.font = Font(name="Calibri", size=10, bold=True, color=COULD_COLOR)
                
        # Status font styling
        if c_idx == 7:
            cell.font = Font(name="Calibri", size=10, bold=True, color="2563EB")
            
    row_idx += 1

# Auto-filter
ws_rtm.auto_filter.ref = f"A1:L{row_idx-1}"

# Column widths
col_widths = {
    "A": 12,  # OQ ID
    "B": 12,  # HMW ID
    "C": 14,  # Req ID (BR)
    "D": 55,  # Business Requirement Description
    "E": 22,  # Stakeholder
    "F": 12,  # Priority
    "G": 14,  # Status
    "H": 18,  # Feature
    "I": 24,  # DB table.column
    "J": 18,  # User Story
    "K": 18,  # UI Screen
    "L": 22   # Test Case
}
for col_letter, width in col_widths.items():
    ws_rtm.column_dimensions[col_letter].width = width

# -------------------------------------------------------------
# Sheet 2: Coverage Analysis
# -------------------------------------------------------------
ws_cov = wb.create_sheet(title="Coverage_Analysis")
ws_cov.views.sheetView[0].showGridLines = True

# Title block
ws_cov.merge_cells("A1:F1")
top_cell = ws_cov["A1"]
top_cell.value = "REQUIREMENTS TRACEABILITY & COVERAGE AUDIT DASHBOARD"
top_cell.font = Font(name="Calibri", size=14, bold=True, color=WHITE)
top_cell.fill = header_fill
top_cell.alignment = Alignment(horizontal="center", vertical="center")
ws_cov.row_dimensions[1].height = 32

# KPI Summary Cards
kpi_labels = [
    ("Total OQs Defined", 10, "A3:B3", "A4:B4"),
    ("Total HMW Statements", 10, "C3:D3", "C4:D4"),
    ("Total Business Requirements", 12, "E3:F3", "E4:F4"),
]
for lbl, val, r1, r2 in kpi_labels:
    ws_cov.merge_cells(r1)
    ws_cov.merge_cells(r2)
    c1 = ws_cov[r1.split(":")[0]]
    c2 = ws_cov[r2.split(":")[0]]
    c1.value = lbl
    c1.font = Font(name="Calibri", size=10, bold=True, color="4B5563")
    c1.alignment = Alignment(horizontal="center", vertical="center")
    c1.fill = PatternFill(start_color="F3F4F6", fill_type="solid")
    c2.value = val
    c2.font = Font(name="Calibri", size=16, bold=True, color=NAVY_HEADER)
    c2.alignment = Alignment(horizontal="center", vertical="center")
    c2.fill = PatternFill(start_color="FFFFFF", fill_type="solid")

kpi_gaps = [
    ("Uncovered OQs", 0, "A6:B6", "A7:B7", SUCCESS_GREEN),
    ("HMWs without BR", 0, "C6:D6", "C7:D7", SUCCESS_GREEN),
    ("BRs without HMW", 0, "E6:F6", "E7:F7", SUCCESS_GREEN),
]
for lbl, val, r1, r2, color in kpi_gaps:
    ws_cov.merge_cells(r1)
    ws_cov.merge_cells(r2)
    c1 = ws_cov[r1.split(":")[0]]
    c2 = ws_cov[r2.split(":")[0]]
    c1.value = lbl
    c1.font = Font(name="Calibri", size=10, bold=True, color="4B5563")
    c1.alignment = Alignment(horizontal="center", vertical="center")
    c1.fill = PatternFill(start_color="F3F4F6", fill_type="solid")
    c2.value = f"{val} Gaps (100% Covered)"
    c2.font = Font(name="Calibri", size=12, bold=True, color=color)
    c2.alignment = Alignment(horizontal="center", vertical="center")
    c2.fill = PatternFill(start_color="FFFFFF", fill_type="solid")

# Section: OQ to HMW to BR Detailed Mapping Table
ws_cov.cell(row=9, column=1, value="Operational Question (OQ) Coverage Verification").font = Font(name="Calibri", size=12, bold=True, color=NAVY_HEADER)
ws_cov.row_dimensions[9].height = 24

cov_headers = ["OQ ID", "Operational Question Title", "Mapped HMW ID", "Mapped Business Requirements", "Coverage Status", "Verification Method"]
ws_cov.row_dimensions[10].height = 24
for j, h in enumerate(cov_headers, start=1):
    c = ws_cov.cell(row=10, column=j, value=h)
    c.fill = header_fill
    c.font = header_font
    c.alignment = header_align
    c.border = thin_border

oq_cov_data = [
    ("OQ-01", "OTIF Delivery Compliance", "HMW-01", "BR-01, BR-11", "Fully Covered", "Mathematical calculation of on-time & full order ratio"),
    ("OQ-02", "Lead Time Variability & SLA Adherence", "HMW-02", "BR-02, BR-11", "Fully Covered", "Statistical mean, std dev, and SLA delta days"),
    ("OQ-03", "Component Defect Rate & Defect PPM", "HMW-03", "BR-03, BR-11", "Fully Covered", "Defect ratio and Defective PPM calculation per part"),
    ("OQ-04", "Contract vs Invoiced Cost Variance", "HMW-04", "BR-04", "Fully Covered", "Unit purchase price delta and aggregate currency variance"),
    ("OQ-05", "Severe Delivery Delay Volatility (>7 Days)", "HMW-05", "BR-05", "Fully Covered", "Outlier tail detection and kurtosis distribution analysis"),
    ("OQ-06", "Composite Performance Index (CPI) Scorecard", "HMW-06", "BR-06, BR-12", "Fully Covered", "Weighted multi-metric scoring model (0-100 scale)"),
    ("OQ-07", "Spend-at-Risk Allocation in Poor Vendors", "HMW-07", "BR-07", "Fully Covered", "Tier C spend exposure ratio against total spend"),
    ("OQ-08", "Seasonal & Quarterly Trend Variations", "HMW-08", "BR-08", "Fully Covered", "Quarterly rolling time-series aggregation"),
    ("OQ-09", "Contract Renewal & CAPA Audit Governance", "HMW-09", "BR-09, BR-12", "Fully Covered", "Automated threshold logic for contract governance status"),
    ("OQ-10", "Transit Mode & Geographic Origin Impacts", "HMW-10", "BR-10", "Fully Covered", "Mean delay comparisons across logistics modes & regions")
]

for idx, row in enumerate(oq_cov_data, start=11):
    ws_cov.row_dimensions[idx].height = 22
    for j, val in enumerate(row, start=1):
        c = ws_cov.cell(row=idx, column=j, value=val)
        c.font = Font(name="Calibri", size=10)
        c.border = thin_border
        if j in [1, 3, 5]:
            c.alignment = Alignment(horizontal="center", vertical="center")
        else:
            c.alignment = Alignment(horizontal="left", vertical="center")
        if j == 5:
            c.font = Font(name="Calibri", size=10, bold=True, color=SUCCESS_GREEN)

# Set widths for coverage sheet
cov_col_widths = {
    "A": 12,
    "B": 38,
    "C": 14,
    "D": 22,
    "E": 18,
    "F": 45
}
for col_letter, width in cov_col_widths.items():
    ws_cov.column_dimensions[col_letter].width = width

output_path = os.path.join('docs', '04_Traceability_Matrix.xlsx')
wb.save(output_path)
print(f'Successfully generated: {output_path}')
