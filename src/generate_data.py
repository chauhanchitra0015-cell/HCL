"""
SUPPLIER PERFORMANCE ANALYSIS — SYNTHETIC DATASET GENERATOR
Module: src/generate_data.py
Seed: 42 (Strict Reproducibility)
Architecture: Relational 3NF (5 Tables + 1 Analytical View)
Fictional Indian Supplier Names & Cities Only
"""

import os
import random
import sqlite3
import datetime
import numpy as np
import pandas as pd
import openpyxl

# Set random seeds for absolute reproducibility
RANDOM_SEED = 42
random.seed(RANDOM_SEED)
np.random.seed(RANDOM_SEED)

os.makedirs('data', exist_ok=True)
os.makedirs('src', exist_ok=True)

# -------------------------------------------------------------
# 1. Master Dimension: Categories (6 Records)
# -------------------------------------------------------------
categories_data = [
    {
        "category_id": "CAT-01",
        "category_name": "Electronics & Sensors",
        "spend_tier": "Strategic",
        "risk_weight": 0.85,
        "description": "High-precision microcontrollers, IoT sensors, wiring harnesses, and PCB assemblies."
    },
    {
        "category_id": "CAT-02",
        "category_name": "Mechanical Castings",
        "spend_tier": "Bottleneck",
        "risk_weight": 0.90,
        "description": "Heavy ductile iron castings, engine blocks, transmission housings, and die-cast alloys."
    },
    {
        "category_id": "CAT-03",
        "category_name": "Hydraulics & Pneumatics",
        "spend_tier": "Strategic",
        "risk_weight": 0.75,
        "description": "Hydraulic cylinders, directional control valves, fluid hoses, and pressure accumulators."
    },
    {
        "category_id": "CAT-04",
        "category_name": "Fasteners & Hardware",
        "spend_tier": "Routine",
        "risk_weight": 0.30,
        "description": "High-tensile Grade 8.8 bolts, locknuts, hex screws, washers, and industrial rivets."
    },
    {
        "category_id": "CAT-05",
        "category_name": "Polymers & Packaging",
        "spend_tier": "Leverage",
        "risk_weight": 0.45,
        "description": "Molded polymer gaskets, protective corrugated containers, pallets, and shrink wraps."
    },
    {
        "category_id": "CAT-06",
        "category_name": "Raw Metals & Alloys",
        "spend_tier": "Leverage",
        "risk_weight": 0.80,
        "description": "Cold-rolled steel coils, extruded aluminum profiles, brass rods, and structural bars."
    }
]
df_categories = pd.DataFrame(categories_data)

# -------------------------------------------------------------
# 2. Master Dimension: Suppliers (25 Fictional Indian Vendors)
# -------------------------------------------------------------
indian_cities = [
    "Pune", "Bengaluru", "Chennai", "Coimbatore", "Faridabad",
    "Gurugram", "Ahmedabad", "Jamshedpur", "Vadodara", "Nashik",
    "Hyderabad", "Rajkot", "Surat", "Ghaziabad", "Indore"
]

fictional_company_names = [
    # Tier A (Preferred Vendors: 8 suppliers)
    ("Bharat Precision Forgings Ltd", "Pune", "Tier A", 96.50, "Active"),
    ("Kaveri Industrial Sensors Pvt Ltd", "Bengaluru", "Tier A", 95.20, "Active"),
    ("Sahyadri Heavy Castings Ltd", "Coimbatore", "Tier A", 94.80, "Active"),
    ("Godavari Hydraulic Works", "Hyderabad", "Tier A", 93.60, "Active"),
    ("Narmada Precision Fasteners Ltd", "Faridabad", "Tier A", 95.90, "Active"),
    ("Vindhya Engineered Polymers", "Vadodara", "Tier A", 92.80, "Active"),
    ("Deccan Advanced Tooling Corp", "Chennai", "Tier A", 94.10, "Active"),
    ("Himalaya Alloy Steels Ltd", "Jamshedpur", "Tier A", 93.40, "Active"),

    # Tier B (Conditional Vendors: 12 suppliers)
    ("Chola Machining Systems", "Coimbatore", "Tier B", 84.50, "Active"),
    ("Maurya Fluid Dynamics Pvt Ltd", "Ahmedabad", "Tier B", 82.30, "Active"),
    ("Konkan Micro-Components Ltd", "Nashik", "Tier B", 81.90, "Active"),
    ("Satpura Industrial Hardware", "Indore", "Tier B", 79.40, "Conditional"),
    ("Aravalli Rubber & Gaskets", "Gurugram", "Tier B", 83.10, "Active"),
    ("Kalinga Metal Extrusions", "Jamshedpur", "Tier B", 78.50, "Conditional"),
    ("Tapti Engineering Solutions", "Surat", "Tier B", 80.20, "Active"),
    ("Marwar Fasteners & Springs", "Rajkot", "Tier B", 77.80, "Conditional"),
    ("Malabar Valves & Actuators", "Bengaluru", "Tier B", 82.00, "Active"),
    ("Bundelkhand Forging Hub", "Ghaziabad", "Tier B", 76.50, "Conditional"),
    ("Coromandel Foundry Craft", "Chennai", "Tier B", 81.20, "Active"),
    ("Vidarbha Electro-Pneumatics", "Pune", "Tier B", 79.80, "Active"),

    # Tier C (At-Risk / Underperforming Vendors: 5 suppliers)
    ("Ananya Industrial Supplies", "Faridabad", "Tier C", 64.20, "Conditional"),
    ("Durgapur Basic Castings", "Jamshedpur", "Tier C", 61.50, "Conditional"),
    ("Gomti Polymer Packaging", "Ghaziabad", "Tier C", 58.70, "Suspended"),
    ("Rohini Sheet Metal Works", "Gurugram", "Tier C", 63.80, "Conditional"),
    ("Shivalik Hardware Fabricators", "Indore", "Tier C", 56.40, "Suspended")
]

suppliers_data = []
for idx, (name, city, tier, rating, status) in enumerate(fictional_company_names, start=1):
    supp_id = f"SUP-{idx:03d}"
    email_slug = name.lower().replace(" ", "").replace("&", "").replace("-", "")[:12]
    suppliers_data.append({
        "supplier_id": supp_id,
        "supplier_name": name,
        "supplier_country": "India",
        "supplier_city": city,
        "contact_email": f"procurement@{email_slug}.co.in",
        "supplier_tier": tier,
        "rating_score": rating,
        "contract_start_date": "2023-01-01",
        "contract_end_date": "2025-12-31" if tier != "Tier C" else "2024-12-31",
        "status": status
    })
df_suppliers = pd.DataFrame(suppliers_data)

# -------------------------------------------------------------
# 3. Master Dimension: Products (40 SKU Records)
# -------------------------------------------------------------
products_catalog = [
    # Electronics (CAT-01): 7 items
    ("Opto-Electronic Proximity Sensor 24V", "CAT-01", "Pieces", 450.00, 14, 500),
    ("Piezoelectric Pressure Transducer 100bar", "CAT-01", "Pieces", 1250.00, 21, 400),
    ("Microcontroller Control Board PCB-X4", "CAT-01", "Pieces", 2800.00, 28, 600),
    ("Industrial Thermocouple Sensor Type-K", "CAT-01", "Pieces", 320.00, 10, 300),
    ("CAN-Bus Heavy Harness Assembly 5m", "CAT-01", "Sets", 1100.00, 18, 450),
    ("Rotary Optical Shaft Encoder 1024PPR", "CAT-01", "Pieces", 1850.00, 25, 350),
    ("High-Voltage Switching Contactor 60A", "CAT-01", "Pieces", 780.00, 12, 550),

    # Mechanical Castings (CAT-02): 7 items
    ("Heavy Ductile Iron Engine Housing GGG-40", "CAT-02", "Pieces", 4800.00, 35, 1200),
    ("Hydraulic Transmission Gearbox Casing", "CAT-02", "Pieces", 6200.00, 42, 1400),
    ("Cast Steel Flywheel Blank 350mm", "CAT-02", "Pieces", 2400.00, 28, 900),
    ("Cast Iron Exhaust Manifold Adapter", "CAT-02", "Pieces", 1450.00, 21, 1100),
    ("Precision Aluminum Die-Cast Arm Bracket", "CAT-02", "Pieces", 890.00, 16, 750),
    ("Centrifugal Pump Impeller Bronze Casting", "CAT-02", "Pieces", 3100.00, 30, 850),
    ("High-Strength Cast Iron Brake Drum 400mm", "CAT-02", "Pieces", 1950.00, 24, 950),

    # Hydraulics & Pneumatics (CAT-03): 7 items
    ("Double-Acting Hydraulic Cylinder 80x400", "CAT-03", "Pieces", 5400.00, 28, 800),
    ("Electro-Hydraulic Directional Valve 4/3", "CAT-03", "Pieces", 3600.00, 20, 650),
    ("Reinforced High-Pressure Hydraulic Hose 2m", "CAT-03", "Pieces", 420.00, 10, 500),
    ("Pneumatic Linear Actuator 50x200mm", "CAT-03", "Pieces", 1750.00, 14, 600),
    ("Hydraulic Filter Assembly 10-Micron", "CAT-03", "Pieces", 680.00, 12, 400),
    ("Proportional Pressure Relief Valve 250bar", "CAT-03", "Pieces", 4200.00, 24, 700),
    ("Rotary Hydraulic Motor 160cc/rev", "CAT-03", "Pieces", 7800.00, 32, 750),

    # Fasteners & Hardware (CAT-04): 7 items
    ("High-Tensile Hex Head Bolt M12x60 Gr8.8", "CAT-04", "Boxes", 180.00, 7, 250),
    ("Nyloc Self-Locking Nut M12 Pack-100", "CAT-04", "Boxes", 95.00, 5, 200),
    ("Socket Head Cap Screw M8x40 Gr10.9", "CAT-04", "Boxes", 140.00, 7, 220),
    ("Heavy Spring Lock Washer M12 Pack-500", "CAT-04", "Boxes", 75.00, 5, 150),
    ("Structural Blind Rivet 4.8x12mm Pack-500", "CAT-04", "Boxes", 110.00, 6, 180),
    ("Hardened Flat Steel Washer M16 Pack-200", "CAT-04", "Boxes", 85.00, 5, 120),
    ("Precision Dowel Pin 10x50mm Pack-50", "CAT-04", "Boxes", 160.00, 8, 300),

    # Polymers & Packaging (CAT-05): 6 items
    ("Molded Nitrile O-Ring Seal Kit 200pc", "CAT-05", "Sets", 320.00, 8, 450),
    ("High-Density Polyethylene Fluid Reservoir", "CAT-05", "Pieces", 480.00, 14, 600),
    ("Heavy-Duty Corrugated Shipping Box 50x50", "CAT-05", "Boxes", 65.00, 4, 300),
    ("Vibration Isolation Rubber Damper Mount", "CAT-05", "Pieces", 185.00, 9, 350),
    ("Stretch Film Industrial Pallet Wrap 500m", "CAT-05", "Boxes", 240.00, 6, 200),
    ("Polyurethane Timing Belt 25mm x 1200mm", "CAT-05", "Pieces", 520.00, 12, 400),

    # Raw Metals (CAT-06): 6 items
    ("Cold-Rolled Carbon Steel Sheet 2.5mm", "CAT-06", "Kg", 85.00, 21, 600),
    ("Extruded Aluminum T-Slot Profile 40x40", "CAT-06", "Meters", 310.00, 15, 500),
    ("Alloy Steel Round Bar 50mm EN-19", "CAT-06", "Kg", 120.00, 24, 700),
    ("Phosphor Bronze Bushing Hollow Bar 40mm", "CAT-06", "Kg", 480.00, 28, 650),
    ("Structural Steel Channel Section 100x50", "CAT-06", "Meters", 550.00, 20, 550),
    ("Stainless Steel Seamless Tube 25mm 304L", "CAT-06", "Meters", 720.00, 26, 600)
]

products_data = []
for idx, (p_name, c_id, uom, price, lead_days, ppm) in enumerate(products_catalog, start=1):
    products_data.append({
        "product_id": f"PRD-{idx:03d}",
        "product_name": p_name,
        "category_id": c_id,
        "unit_of_measure": uom,
        "baseline_contract_price": float(price),
        "standard_lead_time_days": int(lead_days),
        "target_defect_ppm": int(ppm)
    })
df_products = pd.DataFrame(products_catalog, columns=[
    "product_name", "category_id", "unit_of_measure", "baseline_contract_price",
    "standard_lead_time_days", "target_defect_ppm"
])
df_products["product_id"] = [f"PRD-{i:03d}" for i in range(1, len(df_products)+1)]
df_products = df_products[["product_id", "product_name", "category_id", "unit_of_measure",
                           "baseline_contract_price", "standard_lead_time_days", "target_defect_ppm"]]

# -------------------------------------------------------------
# 4. Fact Purchase Orders (300 Orders) & Order Deliveries (350 Lines)
# -------------------------------------------------------------
# Supplier tier mapping lookup
supplier_tier_map = dict(zip(df_suppliers['supplier_id'], df_suppliers['supplier_tier']))
products_dict = df_products.set_index('product_id').to_dict('index')

# Start date for timeline: 2024-01-01 to 2024-12-31 (Full year 2024)
start_date = datetime.date(2024, 1, 1)

orders_records = []
deliveries_records = []

delivery_counter = 1

for order_num in range(1, 301):
    order_id = f"PO-2024-{order_num:03d}"
    po_number = f"PO/2024/{1000 + order_num}"
    
    # Assign supplier: 55% to Tier A, 35% to Tier B, 10% to Tier C
    tier_pick = np.random.choice(["Tier A", "Tier B", "Tier C"], p=[0.55, 0.35, 0.10])
    eligible_suppliers = df_suppliers[df_suppliers['supplier_tier'] == tier_pick]['supplier_id'].tolist()
    supplier_id = random.choice(eligible_suppliers)
    
    # Date of order spread across 2024
    day_offset = int(np.random.beta(1.2, 1.2) * 350)
    order_date = start_date + datetime.timedelta(days=day_offset)
    
    fiscal_year = 2024
    month = order_date.month
    if month in [1, 2, 3]:
        fiscal_quarter = "Q1"
    elif month in [4, 5, 6]:
        fiscal_quarter = "Q2"
    elif month in [7, 8, 9]:
        fiscal_quarter = "Q3"
    else:
        fiscal_quarter = "Q4"
        
    payment_terms = random.choice(["Net 30", "Net 45", "Net 60", "Immediate"])
    
    # Number of lines in this PO (1 line: 85%, 2 lines: 15%)
    num_lines = 1 if order_num <= 250 else 2
    
    po_line_totals = 0.0
    
    for l_idx in range(num_lines):
        delivery_id = f"DEL-{delivery_counter:04d}"
        delivery_counter += 1
        
        # Pick a product
        product_id = random.choice(df_products['product_id'].tolist())
        p_info = products_dict[product_id]
        std_lead = p_info['standard_lead_time_days']
        base_price = p_info['baseline_contract_price']
        
        # Promised delivery date based on standard SLA lead time
        promised_delivery_date = order_date + datetime.timedelta(days=std_lead)
        
        # Order quantity (log-normal distribution)
        ordered_qty = int(np.random.choice([10, 25, 50, 100, 250, 500], p=[0.25, 0.30, 0.20, 0.15, 0.07, 0.03]))
        
        # Order Status: 95% Completed, 5% In-Transit (towards end of year)
        if order_date > datetime.date(2024, 12, 10) and random.random() < 0.60:
            order_status = "In-Transit"
            delivered_qty = 0
            actual_delivery_date = None
            lead_time_days = None
            delay_days = 0
            is_on_time = 0
            is_in_full = 0
            is_otif = 0
            inspected_qty = 0
            defective_qty = 0
            defect_rate_pct = 0.0
            defect_ppm = 0
            defect_reason_code = "None"
            invoiced_price = base_price
            price_var_amt = 0.0
            price_var_pct = 0.0
            shipping_mode = random.choice(["Road Express", "Air Freight"])
        else:
            order_status = "Completed"
            
            # --- PATTERN 1: Supplier Tier Performance ---
            if tier_pick == "Tier A":
                # High reliability: OTIF 95%, lead time near SLA, low defect rate
                lead_delta = int(np.random.normal(loc=-1.0, scale=1.5))
                delivered_qty = ordered_qty  # In-Full 98% of the time
                if random.random() < 0.02:
                    delivered_qty = int(ordered_qty * 0.90)
                
                # Defect rate very low (< 600 PPM)
                inspected_qty = delivered_qty
                if random.random() < 0.15:
                    defective_qty = 1 if inspected_qty < 100 else int(inspected_qty * 0.005)
                else:
                    defective_qty = 0
                    
                # Price variance near zero (-1% to +1%)
                price_var_pct = round(float(np.random.normal(loc=0.0, scale=0.5)), 2)
                
            elif tier_pick == "Tier B":
                # Moderate reliability: OTIF ~82%, lead delta +2 to +4 days, PPM ~1,800
                lead_delta = int(np.random.normal(loc=2.0, scale=3.0))
                if random.random() < 0.10:
                    delivered_qty = int(ordered_qty * random.uniform(0.85, 0.95))
                else:
                    delivered_qty = ordered_qty
                    
                inspected_qty = delivered_qty
                if random.random() < 0.35:
                    defective_qty = max(1, int(inspected_qty * np.random.uniform(0.01, 0.03)))
                else:
                    defective_qty = 0
                    
                price_var_pct = round(float(np.random.normal(loc=1.2, scale=1.0)), 2)
                
            else:  # Tier C
                # Low reliability: OTIF ~60%, heavy lead time delays, defect PPM > 4,500
                lead_delta = int(np.random.exponential(scale=6.0))
                if random.random() < 0.30:
                    delivered_qty = int(ordered_qty * random.uniform(0.70, 0.90))
                else:
                    delivered_qty = ordered_qty
                    
                inspected_qty = delivered_qty
                if random.random() < 0.65:
                    defective_qty = max(2, int(inspected_qty * np.random.uniform(0.03, 0.08)))
                else:
                    defective_qty = 0
                    
                # Chronic price creep (+3% to +8%)
                price_var_pct = round(float(np.random.uniform(2.5, 7.5)), 2)
            
            # --- PATTERN 4: Seasonal Q3/Q4 Performance Dip ---
            if fiscal_quarter in ["Q3", "Q4"]:
                # Extra seasonal delivery delay
                lead_delta += int(np.random.choice([0, 1, 2, 3], p=[0.4, 0.3, 0.2, 0.1]))
            
            # Ensure lead time is at least 2 days
            actual_lead = max(2, std_lead + lead_delta)
            actual_delivery_date = order_date + datetime.timedelta(days=actual_lead)
            lead_time_days = actual_lead
            
            delay_days = max(0, (actual_delivery_date - promised_delivery_date).days)
            
            # Boolean Flags
            is_on_time = 1 if actual_delivery_date <= promised_delivery_date else 0
            is_in_full = 1 if delivered_qty >= ordered_qty else 0
            is_otif = 1 if (is_on_time == 1 and is_in_full == 1) else 0
            
            # Invoiced price and monetary variance
            invoiced_price = round(base_price * (1.0 + price_var_pct / 100.0), 2)
            price_var_amt = round(invoiced_price - base_price, 2)
            
            # Defect calculations
            if inspected_qty > 0 and defective_qty > 0:
                defect_rate_pct = round((defective_qty / inspected_qty) * 100.0, 2)
                defect_ppm = int(defect_rate_pct * 10000)
                # --- PATTERN 5: Defect Pareto Distribution ---
                defect_reason_code = np.random.choice(
                    ["Dimensional Non-Conformance", "Surface Imperfection", "Material Fatigue", "Packaging Damage", "Electrical Fault"],
                    p=[0.45, 0.30, 0.12, 0.08, 0.05]
                )
            else:
                defect_rate_pct = 0.0
                defect_ppm = 0
                defect_reason_code = "None"
                
            # Shipping Mode
            # --- PATTERN 3: Logistics Mode Characteristics ---
            shipping_mode = np.random.choice(
                ["Road Express", "Air Freight", "Sea Ocean", "Rail Cargo"],
                p=[0.55, 0.20, 0.15, 0.10]
            )
            
        po_line_totals += round(ordered_qty * base_price, 2)
        
        deliveries_records.append({
            "delivery_id": delivery_id,
            "order_id": order_id,
            "product_id": product_id,
            "ordered_quantity": ordered_qty,
            "delivered_quantity": delivered_qty,
            "actual_delivery_date": str(actual_delivery_date) if actual_delivery_date else None,
            "delivery_lead_time_days": lead_time_days,
            "delay_days": delay_days,
            "shipping_mode": shipping_mode,
            "contracted_unit_price": base_price,
            "invoiced_unit_price": invoiced_price,
            "price_variance_amount": price_var_amt,
            "price_variance_pct": price_var_pct,
            "inspected_quantity": inspected_qty,
            "defective_quantity": defective_qty,
            "defect_rate_pct": defect_rate_pct,
            "defect_ppm": defect_ppm,
            "defect_reason_code": defect_reason_code,
            "is_on_time": is_on_time,
            "is_in_full": is_in_full,
            "is_otif": is_otif
        })
        
    orders_records.append({
        "order_id": order_id,
        "po_number": po_number,
        "supplier_id": supplier_id,
        "order_date": str(order_date),
        "promised_delivery_date": str(promised_delivery_date),
        "order_status": order_status,
        "payment_terms": payment_terms,
        "total_po_value": po_line_totals,
        "fiscal_year": fiscal_year,
        "fiscal_quarter": fiscal_quarter
    })

df_purchase_orders = pd.DataFrame(orders_records)
df_deliveries = pd.DataFrame(deliveries_records)

# -------------------------------------------------------------
# 5. Inject Controlled Data Quality Issues (For Cleaning Stage)
# -------------------------------------------------------------
print("Injecting controlled data quality issues for Step 15 cleaning...")

# Issue A: 2 Duplicate rows inserted at the end of deliveries
dup_row1 = df_deliveries.iloc[10].copy()
dup_row1['delivery_id'] = "DEL-0998"  # duplicate content with clone ID
dup_row2 = df_deliveries.iloc[45].copy()
dup_row2['delivery_id'] = "DEL-0999"

df_deliveries = pd.concat([df_deliveries, pd.DataFrame([dup_row1, dup_row2])], ignore_index=True)

# Issue B: Extreme Outliers (Delay Days = 28 to 35 days on 4 records)
outlier_indices = [22, 68, 142, 280]
for idx in outlier_indices:
    if idx < len(df_deliveries):
        df_deliveries.at[idx, 'delay_days'] = int(df_deliveries.at[idx, 'delay_days'] + 30)
        df_deliveries.at[idx, 'delivery_lead_time_days'] = int(df_deliveries.at[idx, 'delivery_lead_time_days'] + 30)

# Issue C: Price Variance Extreme Outliers (2 records with +28% and +35% price creep)
df_deliveries.at[15, 'price_variance_pct'] = 28.50
df_deliveries.at[15, 'invoiced_unit_price'] = round(df_deliveries.at[15, 'contracted_unit_price'] * 1.285, 2)
df_deliveries.at[15, 'price_variance_amount'] = round(df_deliveries.at[15, 'invoiced_unit_price'] - df_deliveries.at[15, 'contracted_unit_price'], 2)

df_deliveries.at[95, 'price_variance_pct'] = 34.00
df_deliveries.at[95, 'invoiced_unit_price'] = round(df_deliveries.at[95, 'contracted_unit_price'] * 1.340, 2)
df_deliveries.at[95, 'price_variance_amount'] = round(df_deliveries.at[95, 'invoiced_unit_price'] - df_deliveries.at[95, 'contracted_unit_price'], 2)

# -------------------------------------------------------------
# 6. Save Excel Workbook (data/supplier_performance_data.xlsx)
# -------------------------------------------------------------
excel_path = os.path.join('data', 'supplier_performance_data.xlsx')
with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
    df_suppliers.to_excel(writer, sheet_name='suppliers', index=False)
    df_categories.to_excel(writer, sheet_name='categories', index=False)
    df_products.to_excel(writer, sheet_name='products', index=False)
    df_purchase_orders.to_excel(writer, sheet_name='purchase_orders', index=False)
    df_deliveries.to_excel(writer, sheet_name='order_deliveries', index=False)

print(f"Saved Excel dataset: {excel_path}")

# -------------------------------------------------------------
# 7. Save Individual CSV Files (data/<table>.csv)
# -------------------------------------------------------------
df_suppliers.to_csv(os.path.join('data', 'suppliers.csv'), index=False)
df_categories.to_csv(os.path.join('data', 'categories.csv'), index=False)
df_products.to_csv(os.path.join('data', 'products.csv'), index=False)
df_purchase_orders.to_csv(os.path.join('data', 'purchase_orders.csv'), index=False)
df_deliveries.to_csv(os.path.join('data', 'order_deliveries.csv'), index=False)
print("Saved CSV files to data/*.csv")

# -------------------------------------------------------------
# 8. Create SQLite Database & Views (data/app.db)
# -------------------------------------------------------------
db_path = os.path.join('data', 'app.db')
if os.path.exists(db_path):
    try:
        os.remove(db_path)
    except:
        pass

conn = sqlite3.connect(db_path)
cursor = conn.cursor()
cursor.execute("PRAGMA foreign_keys = ON;")

# Execute DDL
ddl_statements = """
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

CREATE TABLE categories (
    category_id VARCHAR(10) PRIMARY KEY,
    category_name VARCHAR(50) NOT NULL UNIQUE,
    spend_tier VARCHAR(10) NOT NULL CHECK(spend_tier IN ('Strategic', 'Leverage', 'Bottleneck', 'Routine')),
    risk_weight DECIMAL(3,2) NOT NULL CHECK(risk_weight BETWEEN 0.10 AND 1.00),
    description TEXT
);

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
    defect_reason_code VARCHAR(30) NOT NULL,
    is_on_time INTEGER NOT NULL CHECK(is_on_time IN (0, 1)),
    is_in_full INTEGER NOT NULL CHECK(is_in_full IN (0, 1)),
    is_otif INTEGER NOT NULL CHECK(is_otif IN (0, 1)),
    FOREIGN KEY (order_id) REFERENCES purchase_orders(order_id) ON UPDATE CASCADE ON DELETE CASCADE,
    FOREIGN KEY (product_id) REFERENCES products(product_id) ON UPDATE CASCADE ON DELETE RESTRICT
);

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

cursor.executescript(ddl_statements)
conn.commit()

# Insert data into tables
df_suppliers.to_sql('suppliers', conn, if_exists='append', index=False)
df_categories.to_sql('categories', conn, if_exists='append', index=False)
df_products.to_sql('products', conn, if_exists='append', index=False)
df_purchase_orders.to_sql('purchase_orders', conn, if_exists='append', index=False)
df_deliveries.to_sql('order_deliveries', conn, if_exists='append', index=False)

conn.commit()
print("Populated SQLite database: data/app.db successfully")

# -------------------------------------------------------------
# 9. Automated Validation Report & Sanity Checks
# -------------------------------------------------------------
print("\n" + "="*70)
print("AUTOMATED DATASET VALIDATION REPORT")
print("="*70)

tables = ["suppliers", "categories", "products", "purchase_orders", "order_deliveries"]
for tbl in tables:
    count = cursor.execute(f"SELECT COUNT(*) FROM {tbl}").fetchone()[0]
    print(f"Table '{tbl}': {count} rows")

view_count = cursor.execute("SELECT COUNT(*) FROM v_supplier_performance_analysis").fetchone()[0]
view_cols = len(cursor.execute("PRAGMA table_info(v_supplier_performance_analysis)").fetchall())
print(f"Analytical View 'v_supplier_performance_analysis': {view_count} rows, {view_cols} columns")

# Check Foreign Keys Integrity
fk_check = cursor.execute("PRAGMA foreign_key_check").fetchall()
if len(fk_check) == 0:
    print("Foreign Key Integrity: PASS (0 violations)")
else:
    print(f"Foreign Key Integrity: FAIL ({len(fk_check)} violations)")

# Range sanity checks
otif_check = cursor.execute("SELECT COUNT(*) FROM order_deliveries WHERE is_otif NOT IN (0, 1)").fetchone()[0]
print(f"OTIF Range Check: {'PASS' if otif_check == 0 else 'FAIL'}")

defect_check = cursor.execute("SELECT COUNT(*) FROM order_deliveries WHERE defective_quantity > inspected_quantity").fetchone()[0]
print(f"Defect Quantity <= Inspected Check: {'PASS' if defect_check == 0 else 'FAIL'}")

price_check = cursor.execute("SELECT COUNT(*) FROM order_deliveries WHERE contracted_unit_price <= 0").fetchone()[0]
print(f"Unit Price > 0 Check: {'PASS' if price_check == 0 else 'FAIL'}")

conn.close()
print("="*70 + "\n")
