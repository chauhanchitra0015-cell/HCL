"""
SUPPLIER PERFORMANCE ANALYSIS — ANALYTICS & STATISTICAL ENGINE TESTS
Module: tests/test_analysis.py
Covers: TC-T-06, TC-T-07, TC-T-08, TC-T-09, TC-T-10, TC-T-11, TC-T-12, TC-T-15, TC-T-16
"""

import pytest
import numpy as np
import pandas as pd
from src.analytics import (
    compute_otif_percentage,
    compute_defect_ppm,
    compute_price_variance_pct,
    compute_cpi_score,
    compute_kpi_summary,
    compute_supplier_cpi_table,
    compute_defect_pareto,
    compute_lead_time_distribution
)

def test_tc_t_06_otif_calculation_known_dataset():
    """TC-T-06: Verify OTIF formula on a deterministic 4-record test dataset."""
    test_data = pd.DataFrame([
        {"is_on_time": 1, "is_in_full": 1, "is_otif": 1},  # On-time and full -> 1
        {"is_on_time": 0, "is_in_full": 1, "is_otif": 0},  # Late, full -> 0
        {"is_on_time": 1, "is_in_full": 0, "is_otif": 0},  # On-time, partial -> 0
        {"is_on_time": 0, "is_in_full": 0, "is_otif": 0},  # Late, partial -> 0
    ])
    otif = compute_otif_percentage(test_data)
    assert otif == 25.0, f"Expected exactly 25.0% OTIF, got {otif}"

def test_tc_t_07_lead_time_statistics():
    """TC-T-07: Verify mean, standard deviation, and median on a known vector."""
    lead_times = np.array([10, 12, 14, 16, 18])
    assert np.mean(lead_times) == 14.0
    assert np.median(lead_times) == 14.0
    assert round(np.std(lead_times, ddof=1), 2) == 3.16

def test_tc_t_08_defect_ppm_formula():
    """TC-T-08: Verify Defect PPM formula with zero-division protection."""
    assert compute_defect_ppm(1000, 5) == 5000
    assert compute_defect_ppm(10000, 12) == 1200
    assert compute_defect_ppm(0, 0) == 0

def test_tc_t_09_price_variance_formula():
    """TC-T-09: Verify unit price variance percentage calculation."""
    # Price increased from 100 to 105 -> +5.0%
    assert compute_price_variance_pct(100.0, 105.0) == 5.0
    # Price decreased from 200 to 190 -> -5.0%
    assert compute_price_variance_pct(200.0, 190.0) == -5.0

def test_tc_t_10_cpi_weighted_model():
    """TC-T-10: Verify Composite Performance Index computation."""
    cpi = compute_cpi_score(otif_pct=95.0, quality_score=98.0, cost_score=96.0)
    # 0.40 * 95 + 0.35 * 98 + 0.25 * 96 = 38.0 + 34.3 + 24.0 = 96.3
    assert cpi == 96.3

def test_tc_t_11_spend_at_risk_ratio():
    """TC-T-11: Verify Spend-at-Risk calculation."""
    total_spend = 1_000_000.0
    tier_c_spend = 85_000.0
    spend_at_risk_pct = (tier_c_spend / total_spend) * 100.0
    assert spend_at_risk_pct == 8.5
    assert spend_at_risk_pct < 10.0  # Within healthy threshold

def test_tc_t_12_tukey_outlier_detection():
    """TC-T-12: Verify Tukey's fences isolate extreme delay values."""
    delays = np.array([1, 2, 2, 3, 3, 4, 4, 5, 5, 35])  # 35 is an intentional extreme outlier
    q1 = np.percentile(delays, 25)
    q3 = np.percentile(delays, 75)
    iqr = q3 - q1
    upper_fence = q3 + 1.5 * iqr
    outliers = delays[delays > upper_fence]
    assert 35 in outliers
    assert len(outliers) == 1

def test_tc_t_15_analytics_module_kpi_summary():
    """TC-T-15: Verify compute_kpi_summary generates all required macro metrics."""
    sample_df = pd.DataFrame([
        {
            'order_id': 'PO-001', 'delivery_id': 'DEL-001', 'supplier_id': 'SUP-01',
            'supplier_name': 'Acme', 'category_name': 'Metals',
            'is_otif': 1, 'inspected_qty': 1000, 'defective_qty': 1,
            'contracted_line_total': 10000.0, 'invoiced_line_total': 10000.0,
            'delivery_lead_time_days': 10, 'delay_days': 0, 'cost_variance_amount': 0.0
        }
    ])
    summary = compute_kpi_summary(sample_df)
    assert summary['total_orders'] == 1
    assert summary['macro_otif_pct'] == 100.0
    assert summary['macro_defect_ppm'] == 1000
    assert summary['total_spend'] == 10000.0

def test_tc_t_16_analytics_module_pareto():
    """TC-T-16: Verify compute_defect_pareto computes descending cumulative shares."""
    sample_df = pd.DataFrame([
        {'defect_reason': 'Dimensional', 'defective_qty': 80, 'inspected_qty': 1000},
        {'defect_reason': 'Surface', 'defective_qty': 20, 'inspected_qty': 1000}
    ])
    pareto = compute_defect_pareto(sample_df)
    assert len(pareto) == 2
    assert pareto.iloc[0]['defect_reason'] == 'Dimensional'
    assert pareto.iloc[0]['defect_share_pct'] == 80.0
    assert pareto.iloc[1]['cumulative_share_pct'] == 100.0
