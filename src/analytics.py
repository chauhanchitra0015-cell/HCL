"""
SUPPLIER PERFORMANCE ANALYSIS — ANALYTICAL & STATISTICAL ENGINE
Module: src/analytics.py
Implements vectorized KPI computations, multi-metric CPI scoring, Pareto 80/20,
and parametric/non-parametric statistical distributions.
"""

import numpy as np
import pandas as pd
from typing import Dict, Tuple
from src.config import (
    DEFAULT_WEIGHT_DELIVERY,
    DEFAULT_WEIGHT_QUALITY,
    DEFAULT_WEIGHT_COST,
    TIER_A_MIN_SCORE,
    TIER_B_MIN_SCORE
)

def compute_otif_percentage(df: pd.DataFrame) -> float:
    """Calculates OTIF percentage: (On-time and in-full orders / total orders) * 100."""
    if len(df) == 0:
        return 0.0
    return float((df['is_otif'].sum() / len(df)) * 100.0)

def compute_defect_ppm(inspected: int, defective: int) -> int:
    """Calculates Defective Parts Per Million with zero-division safety."""
    if inspected <= 0:
        return 0
    return int((defective / inspected) * 1_000_000)

def compute_price_variance_pct(contracted: float, invoiced: float) -> float:
    """Calculates percentage purchase price variance."""
    if contracted <= 0:
        return 0.0
    return round(((invoiced - contracted) / contracted) * 100.0, 2)

def compute_cpi_score(otif_pct: float, quality_score: float, cost_score: float,
                      w_delivery: float = DEFAULT_WEIGHT_DELIVERY,
                      w_quality: float = DEFAULT_WEIGHT_QUALITY,
                      w_cost: float = DEFAULT_WEIGHT_COST) -> float:
    """Computes weighted Composite Performance Index."""
    score = (w_delivery * otif_pct) + (w_quality * quality_score) + (w_cost * cost_score)
    return round(float(score), 2)

def compute_kpi_summary(df: pd.DataFrame) -> Dict:
    """
    Computes top-level executive macro KPIs across the filtered dataset.
    """
    if len(df) == 0:
        return {
            "total_spend": 0.0,
            "total_orders": 0,
            "total_deliveries": 0,
            "macro_otif_pct": 0.0,
            "macro_defect_ppm": 0,
            "avg_lead_time_days": 0.0,
            "spend_at_risk_pct": 0.0,
            "price_variance_net": 0.0
        }

    total_spend = float(df['invoiced_line_total'].sum()) if 'invoiced_line_total' in df.columns else 0.0
    total_orders = int(df['order_id'].nunique()) if 'order_id' in df.columns else 0
    total_deliveries = len(df)
    
    # OTIF
    otif_count = int(df['is_otif'].sum()) if 'is_otif' in df.columns else 0
    macro_otif_pct = round((otif_count / total_deliveries) * 100.0, 1) if total_deliveries > 0 else 0.0
    
    # Quality PPM
    total_inspected = int(df['inspected_qty'].sum()) if 'inspected_qty' in df.columns else 0
    total_defective = int(df['defective_qty'].sum()) if 'defective_qty' in df.columns else 0
    macro_defect_ppm = compute_defect_ppm(total_inspected, total_defective)
    
    # Lead time
    completed = df.dropna(subset=['delivery_lead_time_days']) if 'delivery_lead_time_days' in df.columns else df
    avg_lead_time = round(float(completed['delivery_lead_time_days'].mean()), 1) if len(completed) > 0 else 0.0
    
    # Price Variance
    price_var_net = float(df['cost_variance_amount'].sum()) if 'cost_variance_amount' in df.columns else 0.0

    # Calculate Spend-at-Risk using supplier level scoring
    cpi_df = compute_supplier_cpi_table(df)
    spend_at_risk_pct = compute_spend_at_risk(cpi_df)

    return {
        "total_spend": total_spend,
        "total_orders": total_orders,
        "total_deliveries": total_deliveries,
        "macro_otif_pct": macro_otif_pct,
        "macro_defect_ppm": macro_defect_ppm,
        "avg_lead_time_days": avg_lead_time,
        "spend_at_risk_pct": spend_at_risk_pct,
        "price_variance_net": price_var_net
    }

def compute_supplier_cpi_table(
    df: pd.DataFrame,
    w_delivery: float = DEFAULT_WEIGHT_DELIVERY,
    w_quality: float = DEFAULT_WEIGHT_QUALITY,
    w_cost: float = DEFAULT_WEIGHT_COST
) -> pd.DataFrame:
    """
    Computes per-supplier multi-metric CPI, Tiers, and CAPA triggers.
    """
    if len(df) == 0:
        return pd.DataFrame()

    grouped = df.groupby(['supplier_id', 'supplier_name', 'category_name']).agg(
        total_deliveries=('delivery_id', 'count'),
        otif_deliveries=('is_otif', 'sum'),
        inspected_qty=('inspected_qty', 'sum'),
        defective_qty=('defective_qty', 'sum'),
        contracted_spend=('contracted_line_total', 'sum'),
        invoiced_spend=('invoiced_line_total', 'sum'),
        avg_lead_time=('delivery_lead_time_days', 'mean'),
        avg_delay_days=('delay_days', 'mean')
    ).reset_index()

    # Delivery Score (OTIF%)
    grouped['otif_pct'] = (grouped['otif_deliveries'] / grouped['total_deliveries'] * 100.0).round(1)
    grouped['delivery_score'] = grouped['otif_pct']

    # Quality Score (normalized 0 - 100 from PPM)
    # 0 PPM -> 100 score; 2500 PPM -> 75 score; >10000 PPM -> 0 score
    def calculate_quality_score(row):
        ppm = compute_defect_ppm(row['inspected_qty'], row['defective_qty'])
        score = 100.0 - (ppm / 100.0)
        return round(float(np.clip(score, 0.0, 100.0)), 1)
        
    grouped['defect_ppm'] = grouped.apply(lambda r: compute_defect_ppm(r['inspected_qty'], r['defective_qty']), axis=1)
    grouped['quality_score'] = grouped.apply(calculate_quality_score, axis=1)

    # Cost Score (normalized 0 - 100 from price variance)
    # <= 0% variance -> 100 score; 5% variance -> 75 score; >= 20% variance -> 0 score
    def calculate_cost_score(row):
        var_pct = compute_price_variance_pct(row['contracted_spend'], row['invoiced_spend'])
        if var_pct <= 0:
            score = 100.0
        else:
            score = 100.0 - (var_pct * 5.0)
        return round(float(np.clip(score, 0.0, 100.0)), 1)
        
    grouped['price_variance_pct'] = grouped.apply(
        lambda r: compute_price_variance_pct(r['contracted_spend'], r['invoiced_spend']), axis=1
    )
    grouped['cost_score'] = grouped.apply(calculate_cost_score, axis=1)

    # CPI Calculation
    grouped['cpi_score'] = (
        (w_delivery * grouped['delivery_score']) +
        (w_quality * grouped['quality_score']) +
        (w_cost * grouped['cost_score'])
    ).round(1)

    # Tier Assignment
    def assign_tier(cpi):
        if cpi >= TIER_A_MIN_SCORE:
            return "Tier A (Preferred)"
        elif cpi >= TIER_B_MIN_SCORE:
            return "Tier B (Conditional)"
        else:
            return "Tier C (At-Risk)"
            
    grouped['performance_tier'] = grouped['cpi_score'].apply(assign_tier)
    grouped['capa_status'] = grouped['cpi_score'].apply(
        lambda x: "MANDATORY CAPA" if x < TIER_B_MIN_SCORE else "COMPLIANT"
    )

    return grouped.sort_values(by='cpi_score', ascending=False)

def compute_spend_at_risk(cpi_df: pd.DataFrame) -> float:
    """
    Computes ratio of invoiced spend going to Tier C suppliers.
    """
    if len(cpi_df) == 0 or 'invoiced_spend' not in cpi_df.columns:
        return 0.0
        
    total_spend = cpi_df['invoiced_spend'].sum()
    if total_spend <= 0:
        return 0.0
        
    tier_c_spend = cpi_df[cpi_df['performance_tier'].str.contains("Tier C", na=False)]['invoiced_spend'].sum()
    return round(float((tier_c_spend / total_spend) * 100.0), 1)

def compute_defect_pareto(df: pd.DataFrame) -> pd.DataFrame:
    """
    Generates Pareto 80/20 distribution table for quality defect root causes and categories.
    """
    if len(df) == 0 or 'defective_qty' not in df.columns:
        return pd.DataFrame()

    pareto = df.groupby('defect_reason').agg(
        defect_count=('defective_qty', 'sum'),
        inspected_count=('inspected_qty', 'sum')
    ).reset_index()

    pareto = pareto[pareto['defect_count'] > 0].sort_values(by='defect_count', ascending=False)
    
    total_defects = pareto['defect_count'].sum()
    if total_defects > 0:
        pareto['defect_share_pct'] = (pareto['defect_count'] / total_defects * 100.0).round(1)
        pareto['cumulative_share_pct'] = pareto['defect_share_pct'].cumsum().round(1)
    else:
        pareto['defect_share_pct'] = 0.0
        pareto['cumulative_share_pct'] = 0.0

    return pareto

def compute_lead_time_distribution(df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Computes Tukey's fences and statistical dispersion for delivery lead times.
    """
    clean = df.dropna(subset=['delivery_lead_time_days']).copy()
    if len(clean) == 0:
        return pd.DataFrame(), pd.DataFrame()

    stats_list = []
    outliers_list = []

    for cat, grp in clean.groupby('category_name'):
        vals = grp['delivery_lead_time_days'].values
        q1 = np.percentile(vals, 25)
        q3 = np.percentile(vals, 75)
        iqr = q3 - q1
        upper_fence = q3 + 1.5 * iqr
        
        cat_outliers = grp[grp['delivery_lead_time_days'] > upper_fence]
        outliers_list.append(cat_outliers)

        stats_list.append({
            "category": cat,
            "mean_days": round(float(np.mean(vals)), 1),
            "median_days": round(float(np.median(vals)), 1),
            "std_days": round(float(np.std(vals, ddof=1)) if len(vals) > 1 else 0.0, 1),
            "iqr_days": round(float(iqr), 1),
            "q1_days": round(float(q1), 1),
            "q3_days": round(float(q3), 1),
            "upper_fence": round(float(upper_fence), 1),
            "outliers_count": len(cat_outliers),
            "sla_baseline": grp['sla_lead_time_days'].iloc[0] if 'sla_lead_time_days' in grp.columns else 14
        })

    summary_df = pd.DataFrame(stats_list)
    outliers_df = pd.concat(outliers_list, ignore_index=True) if outliers_list else pd.DataFrame()

    return summary_df, outliers_df
