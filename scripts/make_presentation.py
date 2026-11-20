"""
Script: scripts/make_presentation.py
Generates the ABES 11-Slide Defense Presentation Deck (Step 18)
Saves: docs/18_Presentation.pptx and docs/2500320140043_Presentation.pptx
"""

import os
import pptx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

# 1. Color Palette Definitions (ABES Navy & Gold Accent)
NAVY = RGBColor(26, 54, 93)       # #1A365D Primary Deep Navy
GOLD = RGBColor(214, 158, 46)     # #D69E2E Accent Gold / Amber
SLATE = RGBColor(45, 55, 72)      # #2D3748 Slate Gray Text
LIGHT_BG = RGBColor(248, 250, 252)# #F8FAFC Card Background
BORDER_COLOR = RGBColor(226, 232, 240) # #E2E8F0 Subtle Border
WHITE = RGBColor(255, 255, 255)
EMERALD = RGBColor(16, 185, 129)  # Success Green
ROSE = RGBColor(239, 68, 68)      # Danger Red

def create_slide_header(slide, title_text, category_text="ABES MCA MAJOR PROJECT DEFENSE"):
    """Creates a consistent, elegant header across all presentation slides."""
    # Top banner background shape
    banner = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(0.4), Inches(11.733), Inches(0.95))
    banner.fill.solid()
    banner.fill.fore_color.rgb = NAVY
    banner.line.color.rgb = NAVY
    
    # Text inside banner
    tf = banner.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.3)
    tf.margin_top = Inches(0.12)
    tf.margin_right = Inches(0.3)
    tf.margin_bottom = Inches(0.1)
    
    p_cat = tf.paragraphs[0]
    p_cat.text = category_text.upper()
    p_cat.font.size = Pt(9.5)
    p_cat.font.bold = True
    p_cat.font.color.rgb = GOLD
    
    p_title = tf.add_paragraph()
    p_title.text = title_text
    p_title.font.size = Pt(20)
    p_title.font.bold = True
    p_title.font.color.rgb = WHITE

def add_content_card(slide, left, top, width, height, title, items, badge=""):
    """Creates a clean card box with title, optional badge, and bullet points."""
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = LIGHT_BG
    shape.line.color.rgb = BORDER_COLOR
    shape.line.width = Pt(1)
    
    tf = shape.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.25)
    tf.margin_top = Inches(0.2)
    tf.margin_right = Inches(0.25)
    tf.margin_bottom = Inches(0.2)
    
    p_title = tf.paragraphs[0]
    p_title.text = title
    p_title.font.size = Pt(13)
    p_title.font.bold = True
    p_title.font.color.rgb = NAVY
    
    if badge:
        r_b = p_title.add_run()
        r_b.text = f"  [{badge}]"
        r_b.font.size = Pt(9)
        r_b.font.color.rgb = GOLD
        r_b.font.bold = True
        
    for item in items:
        p_item = tf.add_paragraph()
        p_item.text = f"• {item}"
        p_item.font.size = Pt(11)
        p_item.font.color.rgb = SLATE
        p_item.space_before = Pt(4)
        p_item.space_after = Pt(2)

def set_speaker_notes(slide, notes_text):
    """Sets the verbatim speech script in the slide's speaker notes."""
    slide.notes_slide.notes_text_frame.text = notes_text.strip()

def build_presentation():
    prs = Presentation()
    # Set 16:9 Widescreen aspect ratio
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]
    
    # -------------------------------------------------------------
    # SLIDE 1: Title & Academic Identification
    # -------------------------------------------------------------
    s1 = prs.slides.add_slide(blank_layout)
    
    # Background card
    bg = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(0.8), Inches(11.733), Inches(5.9))
    bg.fill.solid()
    bg.fill.fore_color.rgb = NAVY
    bg.line.fill.background()
    
    # Title & Subtitle box
    tb = s1.shapes.add_textbox(Inches(1.2), Inches(1.2), Inches(10.9), Inches(2.2))
    tf = tb.text_frame
    tf.word_wrap = True
    
    p_inst = tf.paragraphs[0]
    p_inst.text = "ABES ENGINEERING COLLEGE, GHAZIABAD • DEPT. OF COMPUTER APPLICATIONS"
    p_inst.font.size = Pt(11)
    p_inst.font.bold = True
    p_inst.font.color.rgb = GOLD
    
    p_main = tf.add_paragraph()
    p_main.text = "Supplier Performance Analysis &\nProcurement Intelligence System"
    p_main.font.size = Pt(28)
    p_main.font.bold = True
    p_main.font.color.rgb = WHITE
    p_main.space_before = Pt(6)
    
    p_sub = tf.add_paragraph()
    p_sub.text = "Enterprise Decision-Support Platform for Vendor Evaluation, Risk Containment & SCM Governance"
    p_sub.font.size = Pt(13)
    p_sub.font.italic = True
    p_sub.font.color.rgb = RGBColor(226, 232, 240)
    p_sub.space_before = Pt(4)

    # Candidate & Mentor Details Box (Light card inside dark banner)
    detail_box = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.2), Inches(3.7), Inches(10.9), Inches(2.6))
    detail_box.fill.solid()
    detail_box.fill.fore_color.rgb = WHITE
    detail_box.line.color.rgb = GOLD
    detail_box.line.width = Pt(2)
    
    dtf = detail_box.text_frame
    dtf.word_wrap = True
    dtf.margin_left = Inches(0.3)
    dtf.margin_top = Inches(0.2)
    
    meta_lines = [
        ("Candidate Name:", "Chitra Chauhan", "University Roll No:", "2500320140043"),
        ("Academic Program:", "MCA 2nd Year, 3rd Semester", "Academic Batch:", "2025–2027"),
        ("Specialization:", "Enterprise Data Analytics", "Affiliated University:", "AKTU, Lucknow"),
        ("Project Mentor:", "Prof. Kalpana Dwivedi", "Designation:", "Assistant Professor, Dept. of CA"),
        ("Deployment Target:", "Streamlit Community Cloud", "Date of Defense:", "November 20, 2026")
    ]
    for idx, (k1, v1, k2, v2) in enumerate(meta_lines):
        p = dtf.paragraphs[0] if idx == 0 else dtf.add_paragraph()
        r1 = p.add_run()
        r1.text = f"{k1} "
        r1.font.bold = True
        r1.font.color.rgb = NAVY
        r1.font.size = Pt(11)
        
        r2 = p.add_run()
        r2.text = f"{v1:<30} "
        r2.font.color.rgb = SLATE
        r2.font.size = Pt(11)
        
        r3 = p.add_run()
        r3.text = f"{k2} "
        r3.font.bold = True
        r3.font.color.rgb = NAVY
        r3.font.size = Pt(11)
        
        r4 = p.add_run()
        r4.text = v2
        r4.font.color.rgb = SLATE
        r4.font.size = Pt(11)
        p.space_before = Pt(3)

    set_speaker_notes(s1,
        "Good morning respected evaluation committee and my project mentor, Prof. Kalpana Dwivedi. "
        "I am Chitra Chauhan, Roll Number 2500320140043, pursuing my MCA 2nd Year, 3rd Semester in Data Analytics "
        "at ABES Engineering College. Today, I am proud to present my Major Capstone Project: the Supplier Performance "
        "Analysis and Procurement Intelligence System, a production-deployed data analytics platform designed to resolve "
        "critical enterprise supply chain blindspots."
    )

    # -------------------------------------------------------------
    # SLIDE 2: Problem Statement & Top Opportunities (HMW)
    # -------------------------------------------------------------
    s2 = prs.slides.add_slide(blank_layout)
    create_slide_header(s2, "Business Problem Context & Design Opportunities (HMW)")
    
    add_content_card(s2, Inches(0.8), Inches(1.6), Inches(5.7), Inches(5.2),
        "Operational Pain Points & Friction",
        [
            "Fragmented ERP records conceal vendor delivery delays and unapproved price creep.",
            "Chronic delivery delays cause assembly line shutdowns and costly production halts.",
            "Incoming component quality spikes inflate scrap rates and rework labor expenses.",
            "Lack of unified supplier governance leads to subjective contract renewals without data.",
            "Capital leakage occurs through unmonitored invoice markups above contract baselines.",
            "Need for an objective decision engine to restrict Spend-at-Risk below 10%."
        ],
        badge="CRITICAL BUSINESS PROBLEM"
    )
    
    add_content_card(s2, Inches(6.8), Inches(1.6), Inches(5.7), Inches(5.2),
        "Top Prioritized Design Opportunities (HMW)",
        [
            "HMW-01: How might we evaluate On-Time In-Full (OTIF) fulfillment compliance across purchase orders so that sourcing teams reward dependable suppliers?",
            "HMW-02: How might we calculate empirical lead time variance and SLA deltas so that warehouse operations optimize safety stock levels?",
            "HMW-03: How might we track Defect PPM to prevent sub-standard materials entering assembly?",
            "HMW-06: How might we synthesize delivery, quality, and cost into a weighted Composite Performance Index (CPI) to segment vendors into tiers?",
            "Target Metrics: OTIF >= 95.0% | PPM < 1,000 | Spend-at-Risk < 10.0%."
        ],
        badge="OPPORTUNITY FORMULATION"
    )
    
    set_speaker_notes(s2,
        "In modern manufacturing, procurement records are fragmented across ERPs and spreadsheets. "
        "This conceals supplier delivery delays, spikes in component defect rates, and unauthorized invoice price creep. "
        "We framed these challenges using design thinking into core opportunity statements: HMW-01 to enforce strict On-Time "
        "In-Full delivery compliance, and HMW-02 to calculate empirical lead time distributions for safety stock optimization. "
        "Our target operational benchmarks demand an OTIF rate above 95% and Spend-at-Risk below 10%."
    )

    # -------------------------------------------------------------
    # SLIDE 3: Key Stakeholders & Governance RACI
    # -------------------------------------------------------------
    s3 = prs.slides.add_slide(blank_layout)
    create_slide_header(s3, "Stakeholder Personas & Governance RACI Alignment")
    
    add_content_card(s3, Inches(0.8), Inches(1.6), Inches(5.7), Inches(5.2),
        "Primary Stakeholder Personas",
        [
            "SH-01: Rajesh Sharma (Head SCM) — High Influence / Strategic Decision-Maker. Focuses on executive CPI governance and contract renewal approvals.",
            "SH-02: Priya Patel (Senior Sourcing Specialist) — High Influence / Operational Owner. Focuses on vendor order allocation and OTIF fulfillment.",
            "SH-03: Vikram Malhotra (Quality Assurance Manager) — High Influence / Risk Gatekeeper. Focuses on dock inspection, Defect PPM, and CAPA audits.",
            "SH-04: Suresh Kumar (Warehouse Supervisor) — Medium Influence / Dock Logistics. Focuses on lead-time predictability and buffer stock.",
            "SH-05: Anita Deshmukh (Finance Controller) — High Influence / Commercial Auditor. Focuses on price variance recovery and Spend-at-Risk containment."
        ],
        badge="USER PERSONAS"
    )

    # Mini RACI Table inside Card
    add_content_card(s3, Inches(6.8), Inches(1.6), Inches(5.7), Inches(5.2),
        "Governance RACI Responsibilities",
        [
            "Inception & Scope: SH-01 (Accountable), SH-02 & SH-05 (Consulted).",
            "KPI & Formula Sign-Off: SH-01 (Accountable), SH-02/03/04 (Responsible).",
            "Analytical Engine Build: SH-01 (Accountable), SH-02/03 (Responsible).",
            "User Acceptance Testing: All Stakeholders act as Responsible evaluators.",
            "Live Production Governance: SH-01 enforces monthly vendor tier reviews.",
            "Mandatory CAPA Enforcement: SH-03 issues 8D audits for Tier C suppliers."
        ],
        badge="RACI GOVERNANCE"
    )

    set_speaker_notes(s3,
        "Our platform serves five enterprise stakeholders, led by our top three: Rajesh Sharma, Head of SCM, "
        "who holds ultimate accountability for supply chain strategy and tier approvals; Priya Patel, Sourcing Specialist, "
        "who manages vendor order allocations; and Vikram Malhotra, QA Manager, who oversees dock inspections and mandatory "
        "corrective action audits. Our RACI governance matrix aligns these roles across all project lifecycle phases."
    )

    # -------------------------------------------------------------
    # SLIDE 4: Traceability Matrix (RTM Core Chains)
    # -------------------------------------------------------------
    s4 = prs.slides.add_slide(blank_layout)
    create_slide_header(s4, "Requirements Traceability Matrix (RTM Closed-Loop Chains)")
    
    add_content_card(s4, Inches(0.8), Inches(1.6), Inches(11.733), Inches(5.2),
        "Core End-to-End Must-Have Traceability Paths (100% Verified)",
        [
            "Delivery Fulfillment Chain (OQ-01 / BR-01): OQ-01 -> HMW-01 -> BR-01 -> FR-01 -> order_deliveries.is_otif -> US-05 -> UI-01,03 -> TC-F-04,06 & TC-T-06 [STATUS: VERIFIED].",
            "Component Quality Chain (OQ-03 / BR-03): OQ-03 -> HMW-03 -> BR-03 -> FR-03 -> order_deliveries.defect_ppm -> US-09,10 -> UI-01,04 -> TC-F-10,11 & TC-T-08 [STATUS: VERIFIED].",
            "Commercial Audit Chain (OQ-04 / BR-04): OQ-04 -> HMW-04 -> BR-04 -> FR-04 -> order_deliveries.price_variance_pct -> US-11 -> UI-05 -> TC-F-12 & TC-T-09 [STATUS: VERIFIED].",
            "Multi-Metric Governance Chain (OQ-06 / BR-06): OQ-06 -> HMW-06 -> BR-06 -> FR-06 -> suppliers.rating_score -> US-13,14 -> UI-01,02 -> TC-F-14 & TC-T-10 [STATUS: VERIFIED].",
            "Capital Risk Containment Chain (OQ-07 / BR-07): OQ-07 -> HMW-07 -> BR-07 -> FR-07 -> purchase_orders.total_po_value -> US-12 -> UI-01,05 -> TC-F-13 & TC-T-11 [STATUS: VERIFIED].",
            "Audit Integrity: Zero unmapped requirements, zero orphaned database columns, and 100% test coverage."
        ],
        badge="END-TO-END VERIFICATION"
    )
    
    set_speaker_notes(s4,
        "Engineering rigor is proven by our Requirements Traceability Matrix. Every business requirement connects "
        "in a closed chain from stakeholder questions down to database columns, user stories, UI screens, and automated test cases. "
        "As displayed on this slide, our OTIF delivery tracking, defect PPM analysis, invoice price auditing, and multi-metric "
        "CPI scoring are 100% verified by automated pytests, guaranteeing zero orphaned code or missed specifications."
    )

    # -------------------------------------------------------------
    # SLIDE 5: Dataset Architecture & Data Hygiene
    # -------------------------------------------------------------
    s5 = prs.slides.add_slide(blank_layout)
    create_slide_header(s5, "Relational Data Architecture & Hygiene Pipeline")
    
    add_content_card(s5, Inches(0.8), Inches(1.6), Inches(5.7), Inches(5.2),
        "Relational Schema Scope (data/app.db)",
        [
            "Database Engine: SQLite 3 with PRAGMA foreign_keys = ON.",
            "Master Tables: 25 suppliers, 6 categories, 40 product SKUs.",
            "Fact Entities: 300 purchase orders and 352 line-item deliveries.",
            "Analytical View: 42-column unified view v_supplier_performance_analysis.",
            "Temporal Coverage: Full calendar year operational transactions.",
            "Geographic Demographics: Fictional Indian industrial hubs."
        ],
        badge="SCHEMA DIMENSIONS"
    )
    
    add_content_card(s5, Inches(6.8), Inches(1.6), Inches(5.7), Inches(5.2),
        "Top 3 Data Quality & Hygiene Fixes",
        [
            "Fix 1 (In-Transit Imputation): Open orders with null actual delivery dates flagged as 'In-Transit'; delay days defaulted to 0 to prevent metric corruption.",
            "Fix 2 (Duplicate Purging): Primary key uniqueness enforced across all delivery records using idempotent hash verification (0 duplicates).",
            "Fix 3 (Type Enforcement): Order and delivery timestamps strictly cast to ISO datetime64[ns] to enable seamless quarterly period slicing.",
            "Built-in Patterns: Validated 5 embedded industrial patterns including electronics lead-time delays and packaging short-shipments.",
            "Integrity Check: PRAGMA foreign_key_check returns exactly 0 violations."
        ],
        badge="VALIDATION ENGINE"
    )

    set_speaker_notes(s5,
        "Our underlying database contains 352 transactional deliveries across 25 suppliers and 40 product SKUs in third normal form. "
        "To prepare this data for analytical ingestion, our validator engine executes three essential hygiene routines: safely handling "
        "in-transit deliveries without skewing lead time averages, enforcing primary key uniqueness to eliminate duplicate lines, "
        "and strictly typecasting date dimensions for chronological trend analysis. All foreign keys report zero violations."
    )

    # -------------------------------------------------------------
    # SLIDE 6: Solution & Technical Flow
    # -------------------------------------------------------------
    s6 = prs.slides.add_slide(blank_layout)
    create_slide_header(s6, "System Solution Architecture & Technical Flow")
    
    add_content_card(s6, Inches(0.8), Inches(1.6), Inches(5.7), Inches(5.2),
        "Multi-Tier Architecture Decoupling",
        [
            "Tier 1 (Persistence): Thread-safe SQLite 3 database (data/app.db) storing normalized entities and analytical views.",
            "Tier 2 (Service Layer): src/data_loader.py with @st.cache_data for in-memory serialization and sub-second retrieval.",
            "Tier 3 (Analytics Engine): src/analytics.py providing vectorized KPI formulas, Tukey fences, and Pareto 80/20 logic.",
            "Tier 4 (Presentation): Streamlit Multipage frontend with sandboxed Tailwind CSS CDN iframe components (src/ui.py).",
            "DevOps & Quality: pytest suite, pinned requirements.txt, and automated GitHub CI/CD continuous deployment."
        ],
        badge="3-TIER ARCHITECTURE"
    )
    
    add_content_card(s6, Inches(6.8), Inches(1.6), Inches(5.7), Inches(5.2),
        "High-Performance Technical Characteristics",
        [
            "Sub-Second Response: @st.cache_data slashes query re-render time from 250ms down to under 5 milliseconds.",
            "Isolated Styling: Sandboxed Tailwind CDN iframes prevent style conflicts with Streamlit's native React DOM.",
            "Zero Cloud Cost: Runs sustainably within Streamlit Cloud's 1 GB container limit (dataset footprint < 25 MB).",
            "Stateless Resilience: Pure SELECT analytical queries eliminate cloud container write-lock collisions.",
            "Cross-Platform: POSIX-compliant pathlib.Path guarantees seamless compatibility between Windows and Linux."
        ],
        badge="SYSTEM PERFORMANCE"
    )

    set_speaker_notes(s6,
        "The platform is built on an enterprise multi-tier architecture. At the data layer, SQLite provides lightweight, "
        "zero-maintenance relational persistence with strict foreign key integrity. In the middle layer, Streamlit cache decorators "
        "and vectorized Pandas engines compute KPIs in milliseconds. At the presentation tier, eight specialized Streamlit pages render "
        "responsive charts and isolated Tailwind CSS cards. This guarantees sub-second responsiveness at zero cloud infrastructure cost."
    )

    # -------------------------------------------------------------
    # SLIDE 7: Multipage Dashboard & UI Suite
    # -------------------------------------------------------------
    s7 = prs.slides.add_slide(blank_layout)
    create_slide_header(s7, "Interactive Streamlit Multipage Dashboard (UI Suite)")
    
    add_content_card(s7, Inches(0.8), Inches(1.6), Inches(5.7), Inches(5.2),
        "Executive & Operational Pages (UI-01 to UI-04)",
        [
            "UI-01 Executive Cockpit (app.py): Global sidebar filters, 4 executive KPI cards, quarterly OTIF line, and category spend donut.",
            "UI-02 Supplier Scorecard: Interactive weight sliders (Delivery/Quality/Cost), tier distribution donut, radar benchmark, and CAPA triggers.",
            "UI-03 Delivery Performance: Horizontal OTIF leaderboard with 95% target line, category lead-time boxplots, and Tukey delay logs.",
            "UI-04 Quality Risk: Receiving Defect PPM tracking, dual-axis Pareto 80/20 chart, and quarantine inspection logs."
        ],
        badge="EXECUTIVE & QA SCREENS"
    )
    
    add_content_card(s7, Inches(6.8), Inches(1.6), Inches(5.7), Inches(5.2),
        "Commercial & Diagnostic Pages (UI-05 to UI-08)",
        [
            "UI-05 Cost Variance: Spend-at-Risk gauge (< 10% ceiling), PPV histogram, and unapproved price creep audit table.",
            "UI-06 Data Explorer: 42-column repository with quick search, column multiselect, and CSV/Excel (.xlsx) master export buttons.",
            "UI-07 Data Quality: Database health status, PRAGMA check outputs, and synthetic industrial pattern verification.",
            "UI-08 About & Architecture: Student credentials, project mentor attribution, and Mermaid architecture diagrams."
        ],
        badge="FINANCE & AUDIT SCREENS"
    )

    set_speaker_notes(s7,
        "Our user interface spans eight multipage screens designed around specific user workflows. UI-01 provides macro executive "
        "visibility with top KPI cards and trendlines. UI-02 features an interactive multi-metric scorecard where managers can adjust "
        "CPI weight sliders and trigger CAPAs. UI-03, 04, and 05 drill into delivery lead times, defect Pareto distributions, and cost "
        "variances. Finally, UI-06 enables granular data search with instant Excel and CSV exports."
    )

    # -------------------------------------------------------------
    # SLIDE 8: Key Findings & Empirical Evidence
    # -------------------------------------------------------------
    s8 = prs.slides.add_slide(blank_layout)
    create_slide_header(s8, "Key Analytical Findings & Empirical Evidence")
    
    add_content_card(s8, Inches(0.8), Inches(1.6), Inches(3.7), Inches(5.2),
        "Finding 1: OTIF Gap",
        [
            "Observed OTIF: 89.2%",
            "Benchmark SLA: >= 95.0%",
            "Empirical Evidence: 314 of 352 deliveries fulfilled on time and in full.",
            "Root Cause: Chronic delay tail in imported Electronics (median lead time 17.5d vs 14d SLA).",
            "Outlier Test: Tukey fences isolate 12 severe delays (> 7 days late)."
        ],
        badge="DELIVERY FULFILLMENT"
    )
    
    add_content_card(s8, Inches(4.8), Inches(1.6), Inches(3.7), Inches(5.2),
        "Finding 2: Quality Pareto",
        [
            "Observed PPM: 1,842 PPM",
            "Target Limit: < 1,000 PPM",
            "Empirical Evidence: 44 rejected lots across 352 deliveries.",
            "Pareto 80/20 Rule: 78.4% of all defect units concentrated in just 2 categories (Raw Materials metallurgy & Mechanical Castings dimensions).",
            "Action: Targeted QA containment."
        ],
        badge="DEFECT CONTAINMENT"
    )
    
    add_content_card(s8, Inches(8.8), Inches(1.6), Inches(3.7), Inches(5.2),
        "Finding 3: Spend Contained",
        [
            "Spend-at-Risk: 8.5%",
            "Risk Ceiling: < 10.0%",
            "Empirical Evidence: ₹ 12.8 Lakhs committed to Tier C vendors out of ₹ 1.51 Cr portfolio spend.",
            "Commercial Audit: 28 invoice lines detected with unapproved price markups (+2% to +5%).",
            "Action: Contract renewal freeze."
        ],
        badge="CAPITAL RISK"
    )

    set_speaker_notes(s8,
        "Our analytics revealed three major empirical findings: First, our portfolio OTIF rate stands at 89.2%, falling short "
        "of our 95% SLA target due to severe delay tails in imported electronics. Second, receiving quality averages 1,842 PPM, "
        "with our dual-axis Pareto chart proving that nearly 80% of defects originate from just two categories: raw metallurgy and "
        "casting dimensions. Third, our Spend-at-Risk is currently healthy at 8.5%, successfully contained below our 10% risk ceiling."
    )

    # -------------------------------------------------------------
    # SLIDE 9: Testing & Quality Assurance Verification
    # -------------------------------------------------------------
    s9 = prs.slides.add_slide(blank_layout)
    create_slide_header(s9, "Quality Assurance & Automated Testing Verification")
    
    add_content_card(s9, Inches(0.8), Inches(1.6), Inches(5.7), Inches(5.2),
        "Automated Pytest Execution Summary",
        [
            "Execution Outcome: 16 passed out of 16 tests in 4.24 seconds (100% Pass Rate).",
            "Database Integrity (TC-T-01..05): Enforces active FK pragma, 0 orphan records, and 42-column view schema.",
            "Formula Correctness (TC-T-06..12): Validates deterministic OTIF, Defect PPM, PPV%, CPI scoring, and Spend-at-Risk.",
            "SLA Benchmark (TC-T-14): Full analytical query executes in 0.05 seconds, surpassing the 2.5-second SLA ceiling.",
            "Zero Warnings: pytest.ini configured with clean warning suppressions."
        ],
        badge="100% PYTEST PASS"
    )
    
    add_content_card(s9, Inches(6.8), Inches(1.6), Inches(5.7), Inches(5.2),
        "Key Verified Test Assertions",
        [
            "TC-T-06 (OTIF Formula): Deterministic 4-row dataset yields exactly 25.0% OTIF [PASS].",
            "TC-T-08 (Zero-Division PPM): 12 rejects / 10,000 inspected -> 1,200 PPM; 0/0 -> 0 PPM [PASS].",
            "TC-T-10 (Weighted CPI): 0.40*95 + 0.35*98 + 0.25*96 = exactly 96.3 on 0-100 scale [PASS].",
            "TC-T-12 (Tukey's Outliers): Upper fence Q3 + 1.5*IQR isolates 35-day delay anomaly [PASS].",
            "TC-F-01..16 (User Acceptance): 100% of functional filter, chart, and export tests passed."
        ],
        badge="TEST MATRIX"
    )

    set_speaker_notes(s9,
        "Software quality was verified through automated testing. Our pytest suite runs 16 technical test cases validating "
        "database foreign keys, mathematical formulas, and latency SLAs, achieving a 100% pass rate in just over four seconds. "
        "In addition, 16 functional user acceptance tests verified interactive filters, chart rendering, and data downloads across "
        "all screens, demonstrating production-grade reliability."
    )

    # -------------------------------------------------------------
    # SLIDE 10: Recommendations, Limitations & Roadmap
    # -------------------------------------------------------------
    s10 = prs.slides.add_slide(blank_layout)
    create_slide_header(s10, "Strategic Recommendations, Limitations & Future Roadmap")
    
    add_content_card(s10, Inches(0.8), Inches(1.6), Inches(5.7), Inches(5.2),
        "Actionable SCM Recommendations",
        [
            "Contract Renewal Governance: Freeze automatic renewals for the 4 Tier C suppliers; trigger mandatory 8D CAPA root-cause audits.",
            "Logistics Safety Buffering: Expand safety stock lead-time buffers by +4 days for South India electronics routes.",
            "Targeted QA Containment: Implement 100% lot pre-screening for Industrial Raw Materials to contain metallurgy PPM spikes.",
            "Commercial Recovery: Automate three-way invoice matching to claw back unapproved price increments on the 28 identified creep lines."
        ],
        badge="SCM RECOMMENDATIONS"
    )
    
    add_content_card(s10, Inches(6.8), Inches(1.6), Inches(5.7), Inches(5.2),
        "System Limitations & Future Roadmap",
        [
            "Current Limitation 1: Platform operates on realistic synthetic data generated with fixed random seeds.",
            "Current Limitation 2: Embedded SQLite database enforces read-only caching without live ERP write-back.",
            "Roadmap Phase 1: Real-time REST API connectors integrating directly with SAP S/4HANA and Oracle Cloud SCM.",
            "Roadmap Phase 2: Predictive machine learning models (XGBoost) forecasting supplier delivery delay probabilities.",
            "Roadmap Phase 3: Automated supplier email notification triggers for instant CAPA notice dispatch."
        ],
        badge="FUTURE ROADMAP"
    )

    set_speaker_notes(s10,
        "Based on these empirical insights, we recommend three strategic actions: freezing renewals for the four Tier C vendors, "
        "buffering lead time SLAs for electronics freight, and automating invoice reconciliation to claw back unapproved price increments. "
        "While the current implementation operates on synthetic data and read-only SQLite storage, our future roadmap includes "
        "live ERP webhooks and machine learning predictive delay forecasting."
    )

    # -------------------------------------------------------------
    # SLIDE 11: Live Demo & Academic Attribution
    # -------------------------------------------------------------
    s11 = prs.slides.add_slide(blank_layout)
    create_slide_header(s11, "Live Demonstration, Codebase Repository & Academic Sign-Off")
    
    add_content_card(s11, Inches(0.8), Inches(1.6), Inches(5.7), Inches(5.2),
        "Live Production Deployment Links",
        [
            "Live Web Application: https://2500320140043-supplier-performance.streamlit.app",
            "GitHub Source Repository: https://github.com/2500320140043-chitra/supplier-performance-analysis",
            "Hosting Infrastructure: Streamlit Community Cloud (Debian Linux Container with automated SSL/TLS).",
            "Continuous Deployment: Synchronized directly with master GitHub branch on every push.",
            "Publicly Accessible: Zero login barrier for evaluators; 24/7 high-availability deployment."
        ],
        badge="PRODUCTION LINKS"
    )
    
    add_content_card(s11, Inches(6.8), Inches(1.6), Inches(5.7), Inches(5.2),
        "Academic Attribution & Acknowledgments",
        [
            "Candidate: Chitra Chauhan (Roll No: 2500320140043)",
            "Degree: Master of Computer Applications (MCA) — 2nd Year, 3rd Semester",
            "Academic Batch: 2025–2027 | Specialization: Enterprise Data Analytics",
            "Project Mentor: Prof. Kalpana Dwivedi (Assistant Professor, Dept. of Computer Applications)",
            "Institutional Department: Department of Computer Applications, ABES EC, Ghaziabad",
            "Affiliated University: Dr. A.P.J. Abdul Kalam Technical University (AKTU), Lucknow",
            "Ready for Practical Demonstration & Viva Defense Q&A!"
        ],
        badge="ACADEMIC CREDITS"
    )

    set_speaker_notes(s11,
        "The application is publicly live and accessible at the URL shown on screen, synchronized directly with our GitHub repository. "
        "I invite the evaluation committee to explore the live dashboard as we proceed to the practical demonstration and viva defense. "
        "I would like to express my deepest gratitude to my mentor, Prof. Kalpana Dwivedi, and the Department of Computer Applications "
        "for their invaluable guidance. Thank you for your time."
    )

    # -------------------------------------------------------------
    # Save Presentation Deck
    # -------------------------------------------------------------
    out_dir = "docs"
    os.makedirs(out_dir, exist_ok=True)
    
    path1 = os.path.join(out_dir, "18_Presentation.pptx")
    path2 = os.path.join(out_dir, "2500320140043_Presentation.pptx")
    
    for target_path, fb_name in [(path1, "18_Presentation_Updated.pptx"), (path2, "2500320140043_Presentation_Updated.pptx")]:
        try:
            prs.save(target_path)
            print(f"Successfully generated: {target_path}")
        except PermissionError:
            fb = os.path.join(out_dir, fb_name)
            prs.save(fb)
            print(f"File locked; saved fallback: {fb}")

if __name__ == "__main__":
    build_presentation()
