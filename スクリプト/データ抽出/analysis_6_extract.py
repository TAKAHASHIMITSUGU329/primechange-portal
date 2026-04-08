#!/usr/bin/env python3
"""
分析6: 品質→売上弾力性分析
Quality → Revenue Elasticity Analysis

19ホテルの口コミスコアと売上指標（稼働率/ADR/RevPAR）の関係を
回帰分析で定量化し、業界ベンチマークと比較する。
"""

import json
import os
from datetime import date
import numpy as np

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Load data
quality_data = json.load(open(os.path.join(BASE_DIR, "primechange_portfolio_analysis.json"), "r"))
revenue_data = json.load(open(os.path.join(BASE_DIR, "hotel_revenue_data.json"), "r"))

overview = quality_data["portfolio_overview"]
hotels_ranked = overview["hotels_ranked"]

KEY_MAP = {"keisei_kinshicho": "keisei_richmond", "comfort_yokohama_kannai": "comfort_yokohama"}


def rev_key(k):
    return KEY_MAP.get(k, k)


# Build integrated dataset
hotels = []
for h in hotels_ranked:
    rk = rev_key(h["key"])
    r = revenue_data.get(rk, {})
    if not r:
        print(f"  ⚠️ Revenue data not found for {h['key']} (mapped: {rk})")
        continue
    hotels.append({
        "key": h["key"],
        "name": h["name"],
        "score": h["avg"],
        "rank": h["rank"],
        "priority": h["priority"],
        "tier": h["tier"],
        "total_reviews": h["total_reviews"],
        "high_rate": h["high_rate"],
        "low_rate": h["low_rate"],
        "cleaning_issue_rate": h.get("cleaning_issue_rate", 0),
        "revenue": r.get("actual_revenue", 0),
        "occupancy": r.get("occupancy_rate", 0),
        "march_revenue": r.get("march_revenue", 0),
        "march_occupancy": r.get("march_occupancy", 0),
        "april_revenue": r.get("april_revenue", 0),
        "april_occupancy": r.get("april_occupancy", 0),
        "profit_rate": r.get("profit_rate", 0),
        "adr": r.get("adr", 0),
        "revpar": r.get("revpar", 0),
        "room_count": r.get("room_count", 0),
        "staff_count": r.get("staff_count", 0),
        "complaint_rate": r.get("complaint_rate", 0),
        "phase": r.get("phase", ""),
    })

print(f"Loaded {len(hotels)} hotels with both quality and revenue data")

# ============================================================
# 1. Simple Linear Regression
# ============================================================

scores = np.array([h["score"] for h in hotels])
occupancies = np.array([h["occupancy"] for h in hotels])
adrs = np.array([h["adr"] for h in hotels])
revpars = np.array([h["revpar"] for h in hotels])
revenues = np.array([h["revenue"] for h in hotels])
room_counts = np.array([h["room_count"] for h in hotels])


def linear_regression(x, y):
    """Simple linear regression using numpy."""
    n = len(x)
    if n < 3:
        return {"slope": 0, "intercept": 0, "r": 0, "r_squared": 0, "n": n}

    slope, intercept = np.polyfit(x, y, 1)
    corr = np.corrcoef(x, y)[0, 1]
    r_squared = corr ** 2

    # Predicted values
    y_pred = slope * x + intercept
    ss_res = np.sum((y - y_pred) ** 2)
    ss_tot = np.sum((y - np.mean(y)) ** 2)

    return {
        "slope": round(float(slope), 6),
        "intercept": round(float(intercept), 6),
        "r": round(float(corr), 4),
        "r_squared": round(float(r_squared), 4),
        "n": n,
    }


# Score vs Occupancy
reg_occ = linear_regression(scores, occupancies)
reg_occ["x_label"] = "口コミスコア"
reg_occ["y_label"] = "稼働率"
reg_occ["interpretation"] = f"スコア1点上昇で稼働率が{reg_occ['slope']*100:.1f}%ポイント変動"

# Score vs ADR
reg_adr = linear_regression(scores, adrs)
reg_adr["x_label"] = "口コミスコア"
reg_adr["y_label"] = "ADR（客室単価）"
reg_adr["interpretation"] = f"スコア1点上昇でADRが{reg_adr['slope']:.0f}円変動"

# Score vs RevPAR
reg_revpar = linear_regression(scores, revpars)
reg_revpar["x_label"] = "口コミスコア"
reg_revpar["y_label"] = "RevPAR"
reg_revpar["interpretation"] = f"スコア1点上昇でRevPARが{reg_revpar['slope']:.0f}円変動"

# Score vs Revenue (controlled for room count: revenue per room)
rev_per_room = revenues / np.maximum(room_counts, 1)
reg_rev_per_room = linear_regression(scores, rev_per_room)
reg_rev_per_room["x_label"] = "口コミスコア"
reg_rev_per_room["y_label"] = "1室あたり月間売上"
reg_rev_per_room["interpretation"] = f"スコア1点上昇で1室あたり月間売上が{reg_rev_per_room['slope']:.0f}円変動"

print(f"\nRegression Results:")
print(f"  Score vs Occupancy: r={reg_occ['r']:.3f}, R²={reg_occ['r_squared']:.3f}")
print(f"  Score vs ADR: r={reg_adr['r']:.3f}, R²={reg_adr['r_squared']:.3f}")
print(f"  Score vs RevPAR: r={reg_revpar['r']:.3f}, R²={reg_revpar['r_squared']:.3f}")
print(f"  Score vs Revenue/Room: r={reg_rev_per_room['r']:.3f}, R²={reg_rev_per_room['r_squared']:.3f}")

# ============================================================
# 2. Threshold Analysis (Score Band Comparison)
# ============================================================

bands = [
    {"range": "7.0-8.0", "min": 7.0, "max": 8.0},
    {"range": "8.0-8.5", "min": 8.0, "max": 8.5},
    {"range": "8.5-9.0", "min": 8.5, "max": 9.0},
    {"range": "9.0+", "min": 9.0, "max": 10.1},
]

threshold_groups = []
for band in bands:
    group_hotels = [h for h in hotels if band["min"] <= h["score"] < band["max"]]
    if group_hotels:
        avg_occ = np.mean([h["occupancy"] for h in group_hotels])
        avg_adr = np.mean([h["adr"] for h in group_hotels])
        avg_revpar = np.mean([h["revpar"] for h in group_hotels])
        avg_rev = np.mean([h["revenue"] for h in group_hotels])
        avg_score = np.mean([h["score"] for h in group_hotels])
        avg_complaint = np.mean([h["complaint_rate"] for h in group_hotels])
        hotel_names = [h["name"] for h in group_hotels]
    else:
        avg_occ = avg_adr = avg_revpar = avg_rev = avg_score = avg_complaint = 0
        hotel_names = []

    threshold_groups.append({
        "range": band["range"],
        "count": len(group_hotels),
        "avg_score": round(avg_score, 2),
        "avg_occupancy": round(float(avg_occ), 4),
        "avg_occupancy_pct": f"{avg_occ*100:.1f}%",
        "avg_adr": round(float(avg_adr), 1),
        "avg_revpar": round(float(avg_revpar), 1),
        "avg_revenue": round(float(avg_rev)),
        "avg_complaint_rate": round(float(avg_complaint), 4),
        "hotels": hotel_names,
    })

# Detect threshold effect
if len(threshold_groups) >= 2:
    below_8 = threshold_groups[0]
    above_8 = threshold_groups[1]
    occ_diff = above_8["avg_occupancy"] - below_8["avg_occupancy"]
    threshold_effect = {
        "threshold_score": 8.0,
        "below_avg_occupancy": below_8["avg_occupancy_pct"],
        "above_avg_occupancy": above_8["avg_occupancy_pct"],
        "occupancy_difference": f"{occ_diff*100:.1f}%ポイント",
        "description": f"スコア8.0を境に稼働率が{occ_diff*100:.1f}%ポイント変動",
    }
else:
    threshold_effect = {"description": "データ不足で閾値検出不可"}

print(f"\nThreshold Analysis:")
for g in threshold_groups:
    print(f"  {g['range']}: {g['count']}ホテル, 稼働率{g['avg_occupancy_pct']}, RevPAR {g['avg_revpar']:.0f}")

# ============================================================
# 3. Revenue Impact Estimation
# ============================================================

total_rev = sum(h["revenue"] for h in hotels)
avg_revpar = float(np.mean(revpars))

# Use our regression slope or industry benchmark
our_revpar_elasticity = reg_revpar["slope"]  # RevPAR change per 1-point score improvement

# Industry benchmark: 0.1 point = 1% RevPAR improvement
# So 1 point = 10% RevPAR improvement
industry_revpar_pct_per_point = 0.10  # 10%

# Our data: RevPAR change per 0.1 point
our_revpar_per_01 = our_revpar_elasticity * 0.1
our_revpar_pct_per_01 = our_revpar_per_01 / avg_revpar if avg_revpar > 0 else 0

impact_scenarios = []
for improvement in [0.1, 0.3, 0.5]:
    revpar_change = our_revpar_elasticity * improvement
    revpar_pct_change = revpar_change / avg_revpar if avg_revpar > 0 else 0

    # Estimate per-hotel revenue impact
    hotel_impacts = []
    for h in hotels:
        rev_change = h["revenue"] * revpar_pct_change if h["revenue"] > 0 else 0
        hotel_impacts.append({
            "name": h["name"],
            "current_score": h["score"],
            "current_revenue": h["revenue"],
            "estimated_revenue_change": round(rev_change),
        })

    total_monthly_impact = sum(hi["estimated_revenue_change"] for hi in hotel_impacts)

    impact_scenarios.append({
        "score_improvement": improvement,
        "revpar_change": round(float(revpar_change), 1),
        "revpar_pct_change": f"{revpar_pct_change*100:.1f}%",
        "total_monthly_revenue_change": round(total_monthly_impact),
        "annual_revenue_change": round(total_monthly_impact * 12),
        "per_hotel_impacts": sorted(hotel_impacts, key=lambda x: x["estimated_revenue_change"], reverse=True),
    })

# Benchmark comparison
benchmark_comparison = {
    "industry_benchmark": "スコア0.1点改善 = RevPAR 1%向上",
    "industry_revpar_pct_per_01": "1.0%",
    "our_data_revpar_pct_per_01": f"{our_revpar_pct_per_01*100:.2f}%",
    "our_revpar_change_per_01": f"{our_revpar_per_01:.1f}円",
    "deviation": f"業界ベンチマーク比 {our_revpar_pct_per_01*100/1.0:.0f}%" if our_revpar_pct_per_01 != 0 else "算出不可",
    "interpretation": "",
}

if our_revpar_pct_per_01 > 0.01:
    benchmark_comparison["interpretation"] = "自社データでも正の相関が確認された。業界ベンチマークと同等以上の弾力性がある可能性。"
elif our_revpar_pct_per_01 > 0:
    benchmark_comparison["interpretation"] = "自社データでは業界ベンチマークより弱い正の相関。立地・ブランド等の要因が大きい可能性。ただしサンプルサイズ(N=19)の制約に注意。"
else:
    benchmark_comparison["interpretation"] = "自社データでは負の相関が見られた。ホテル規模や立地の差異が品質の影響を上回っている可能性。より細かいセグメント分析が必要。"

print(f"\nBenchmark Comparison:")
print(f"  Industry: 0.1pt = RevPAR 1.0%")
print(f"  Our data: 0.1pt = RevPAR {our_revpar_pct_per_01*100:.2f}%")

# ============================================================
# 4. Per-Hotel Improvement Potential
# ============================================================

# For each hotel below average score, estimate improvement potential
avg_score = float(np.mean(scores))
hotel_potentials = []
for h in hotels:
    target_score = min(h["score"] + 0.5, 9.5)  # Target: +0.5 or cap at 9.5
    improvement = target_score - h["score"]
    if improvement <= 0:
        improvement = 0.3  # Even top hotels can improve 0.3
        target_score = h["score"] + 0.3

    revpar_change = our_revpar_elasticity * improvement
    revpar_pct = revpar_change / h["revpar"] if h["revpar"] > 0 else 0
    rev_change = h["revenue"] * revpar_pct if h["revenue"] > 0 else 0

    hotel_potentials.append({
        "key": h["key"],
        "name": h["name"],
        "current_score": h["score"],
        "target_score": round(target_score, 1),
        "improvement": round(improvement, 1),
        "priority": h["priority"],
        "current_revenue": h["revenue"],
        "current_occupancy": round(h["occupancy"] * 100, 1),
        "current_revpar": round(h["revpar"], 1),
        "estimated_revpar_change": round(float(revpar_change), 1),
        "estimated_revenue_change": round(rev_change),
        "estimated_revenue_change_annual": round(rev_change * 12),
    })

hotel_potentials.sort(key=lambda x: x["estimated_revenue_change"], reverse=True)

total_potential = sum(hp["estimated_revenue_change"] for hp in hotel_potentials)
print(f"\nTotal monthly improvement potential: ¥{total_potential:,.0f}")
print(f"Total annual improvement potential: ¥{total_potential*12:,.0f}")

# ============================================================
# 5. Phase-level Analysis
# ============================================================

phases = {}
for h in hotels:
    p = h["phase"]
    if p not in phases:
        phases[p] = {"hotels": [], "scores": [], "occupancies": [], "revenues": []}
    phases[p]["hotels"].append(h["name"])
    phases[p]["scores"].append(h["score"])
    phases[p]["occupancies"].append(h["occupancy"])
    phases[p]["revenues"].append(h["revenue"])

phase_analysis = []
for p in sorted(phases.keys()):
    d = phases[p]
    phase_analysis.append({
        "phase": p,
        "count": len(d["hotels"]),
        "avg_score": round(float(np.mean(d["scores"])), 2),
        "avg_occupancy": f"{float(np.mean(d['occupancies']))*100:.1f}%",
        "avg_revenue": round(float(np.mean(d["revenues"]))),
        "total_revenue": round(float(np.sum(d["revenues"]))),
    })

# ============================================================
# 6. Data Points Table (for pseudo-scatter plot in report)
# ============================================================

data_points = sorted([{
    "name": h["name"],
    "score": h["score"],
    "occupancy": round(h["occupancy"] * 100, 1),
    "adr": round(h["adr"], 0),
    "revpar": round(h["revpar"], 0),
    "revenue": h["revenue"],
    "march_revenue": h["march_revenue"],
    "april_revenue": h["april_revenue"],
    "room_count": h["room_count"],
    "phase": h["phase"],
    "priority": h["priority"],
    "rev_per_room": round(h["revenue"] / max(h["room_count"], 1)),
} for h in hotels], key=lambda x: x["score"])

# ============================================================
# Build Output JSON
# ============================================================

output = {
    "analysis_metadata": {
        "title": "品質→売上弾力性分析",
        "analysis_number": 6,
        "date": date.today().isoformat(),
        "hotels_count": len(hotels),
        "data_sources": [
            "hotel_revenue_data.json（19ホテル売上実績）",
            "primechange_portfolio_analysis.json（19ホテル口コミスコア）",
        ],
        "methodology": "19ホテル横断の単回帰分析（numpy.polyfit）",
        "limitations": [
            "サンプルサイズN=19のため、統計的有意性は限定的",
            "ホテル規模（客室数）、立地、ブランド等の交絡変数を完全には制御できていない",
            "スナップショットデータであり、時系列変動は未反映",
            "因果関係ではなく相関関係の分析である点に注意",
        ],
    },
    "portfolio_summary": {
        "total_hotels": len(hotels),
        "total_monthly_revenue": round(total_rev),
        "avg_score": round(float(np.mean(scores)), 2),
        "median_score": round(float(np.median(scores)), 2),
        "avg_occupancy": f"{float(np.mean(occupancies))*100:.1f}%",
        "avg_adr": round(float(np.mean(adrs)), 1),
        "avg_revpar": round(float(np.mean(revpars)), 1),
        "feb_total_revenue": round(sum(h["revenue"] for h in hotels)),
        "feb_avg_occupancy": f"{float(np.mean([h['occupancy'] for h in hotels]))*100:.1f}%",
        "mar_total_revenue": round(sum(h["march_revenue"] for h in hotels)),
        "mar_avg_occupancy": f"{float(np.mean([h['march_occupancy'] for h in hotels]))*100:.1f}%",
        "apr_total_revenue": round(sum(h["april_revenue"] for h in hotels)),
        "apr_avg_occupancy": f"{float(np.mean([h['april_occupancy'] for h in hotels]))*100:.1f}%",
    },
    "regression_results": {
        "score_vs_occupancy": reg_occ,
        "score_vs_adr": reg_adr,
        "score_vs_revpar": reg_revpar,
        "score_vs_rev_per_room": reg_rev_per_room,
    },
    "threshold_analysis": {
        "groups": threshold_groups,
        "threshold_effect": threshold_effect,
    },
    "revenue_impact_scenarios": impact_scenarios,
    "benchmark_comparison": benchmark_comparison,
    "hotel_improvement_potentials": hotel_potentials,
    "total_improvement_potential": {
        "monthly": round(total_potential),
        "annual": round(total_potential * 12),
    },
    "phase_analysis": phase_analysis,
    "data_points": data_points,
}

# Save
output_path = os.path.join(BASE_DIR, "analysis_6_data.json")
with open(output_path, "w", encoding="utf-8") as f:
    json.dump(output, f, ensure_ascii=False, indent=2)

print(f"\n✅ analysis_6_data.json saved ({os.path.getsize(output_path):,} bytes)")
print(f"   {len(hotels)} hotels, {len(impact_scenarios)} scenarios, {len(hotel_potentials)} hotel potentials")
