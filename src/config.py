"""
SUPPLIER PERFORMANCE ANALYSIS — APPLICATION CONFIGURATION
Module: src/config.py
Author: Chitra Chauhan (Roll No: 2500320140043)
Mentor: Prof. Kalpana Dwivedi
Institution: ABES Engineering College
"""

import os
from pathlib import Path

# Project Base Directory
BASE_DIR = Path(__file__).resolve().parent.parent

# Database Configuration
DATA_DIR = BASE_DIR / "data"
DB_PATH = DATA_DIR / "app.db"
EXCEL_PATH = DATA_DIR / "supplier_performance_data.xlsx"

# Theme Palettes
PRIMARY_COLOR = "#1A365D"  # Deep Navy
SECONDARY_COLOR = "#D69E2E"  # Amber Accent
BACKGROUND_COLOR = "#F8FAFC"  # Clean Slate 50
TEXT_COLOR = "#0F172A"  # Dark Slate 900
SUCCESS_COLOR = "#10B981"  # Emerald 500
WARNING_COLOR = "#F59E0B"  # Amber 500
DANGER_COLOR = "#EF4444"  # Rose 500
INFO_COLOR = "#3B82F6"  # Blue 500

# Benchmark & KPI SLA Thresholds
TARGET_OTIF_PCT = 95.0
MAX_DEFECT_PPM = 1000
MAX_SPEND_AT_RISK_PCT = 10.0
SEVERE_DELAY_DAYS = 7

# CPI Composite Weights & Tier Definitions
DEFAULT_WEIGHT_DELIVERY = 0.40
DEFAULT_WEIGHT_QUALITY = 0.35
DEFAULT_WEIGHT_COST = 0.25

TIER_A_MIN_SCORE = 85.0
TIER_B_MIN_SCORE = 70.0

# Academic & System Metadata
STUDENT_NAME = "Chitra Chauhan"
STUDENT_ROLL = "2500320140043"
STUDENT_PROGRAM = "MCA 2nd Year (Data Analytics)"
MENTOR_NAME = "Prof. Kalpana Dwivedi"
INSTITUTION = "ABES Engineering College, Ghaziabad"
APP_TITLE = "Supplier Performance Analysis & Procurement Intelligence"
APP_SUBTITLE = "Enterprise Decision-Support Platform for Vendor Evaluation, Risk Mitigation & SCM Analytics"
APP_VERSION = "v1.0.0 (Production)"
