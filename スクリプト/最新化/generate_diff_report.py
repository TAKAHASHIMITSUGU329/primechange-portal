#!/usr/bin/env python3
"""旧/新 分析JSONの差分比較レポート生成スクリプト"""

import json
import os
import sys
from datetime import datetime

BASE = "/Users/mitsugutakahashi/ホテル口コミ/データ/分析結果JSON"
BACKUP = os.path.join(BASE, "backup_2026-03-11")
OUTPUT = os.path.join(BASE, "diff_report_2026-03-12.txt")

HOTEL_KEYS = [
    "daiwa_osaki", "chisan", "hearton", "keyakigate", "richmond_mejiro",
    "keisei_kinshicho", "daiichi_ikebukuro", "comfort_roppongi",
    "comfort_suites_tokyobay", "comfort_era_higashikanda",
    "comfort_yokohama_kannai", "comfort_narita", "apa_kamata",
    "apa_sagamihara", "court_shinyokohama", "comment_yokohama",
    "kawasaki_nikko", "henn_na_haneda", "comfort_hakata"
]


def load_json(path):
    if not os.path.exists(path):
        return None
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def fmt_delta(old_val, new_val, fmt=".2f"):
    if old_val is None or new_val is None:
        return "N/A"
    diff = new_val - old_val
    sign = "+" if diff > 0 else ""
    return f"{sign}{diff:{fmt}}"


def compare_hotel(key):
    old = load_json(os.path.join(BACKUP, f"{key}_analysis.json"))
    new = load_json(os.path.join(BASE, f"{key}_analysis.json"))
    if not old or not new:
        return None

    result = {"key": key}
    result["hotel_name"] = new.get("hotel_name", key)

    # Review counts
    old_total = old.get("total_reviews", 0)
    new_total = new.get("total_reviews", 0)
    result["old_reviews"] = old_total
    result["new_reviews"] = new_total
    result["review_delta"] = new_total - old_total

    # Average score
    old_avg = old.get("overall_avg_10pt", 0)
    new_avg = new.get("overall_avg_10pt", 0)
    result["old_avg"] = old_avg
    result["new_avg"] = new_avg
    result["avg_delta"] = new_avg - old_avg

    # Rating distribution
    for label, rate_key in [("high", "high_rate"), ("mid", "mid_rate"), ("low", "low_rate")]:
        old_rate = old.get(rate_key, 0)
        new_rate = new.get(rate_key, 0)
        result[f"old_{label}"] = old_rate
        result[f"new_{label}"] = new_rate
        result[f"{label}_delta"] = new_rate - old_rate

    # Site breakdown
    old_sites = {s["site"]: s for s in old.get("site_stats", [])}
    new_sites = {s["site"]: s for s in new.get("site_stats", [])}
    site_changes = []
    all_site_names = sorted(set(list(old_sites.keys()) + list(new_sites.keys())))
    for site in all_site_names:
        os_ = old_sites.get(site, {})
        ns_ = new_sites.get(site, {})
        site_changes.append({
            "site": site,
            "old_count": os_.get("count", 0),
            "new_count": ns_.get("count", 0),
            "old_avg": os_.get("avg_10pt", 0),
            "new_avg": ns_.get("avg_10pt", 0),
        })
    result["site_changes"] = site_changes

    return result


def main():
    lines = []
    lines.append("=" * 70)
    lines.append("ホテル口コミ分析 差分レポート")
    lines.append(f"比較: 2026-03-11 → 2026-03-12")
    lines.append(f"生成日時: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append("=" * 70)

    total_old = 0
    total_new = 0
    results = []

    for key in HOTEL_KEYS:
        r = compare_hotel(key)
        if r:
            results.append(r)
            total_old += r["old_reviews"]
            total_new += r["new_reviews"]

    # Summary
    lines.append("")
    lines.append("■ 全体サマリー")
    lines.append(f"  対象ホテル数: {len(results)}")
    lines.append(f"  総口コミ数: {total_old} → {total_new} ({'+' if total_new >= total_old else ''}{total_new - total_old}件)")
    lines.append("")

    # Per-hotel table
    lines.append("■ ホテル別比較")
    lines.append(f"{'ホテル名':<30} {'口コミ数':>12} {'平均スコア':>14} {'高評価率':>12} {'低評価率':>12}")
    lines.append("-" * 82)

    for r in results:
        name = r["hotel_name"][:28]
        reviews = f"{r['old_reviews']}→{r['new_reviews']}({'+' if r['review_delta'] >= 0 else ''}{r['review_delta']})"
        avg = f"{r['old_avg']:.2f}→{r['new_avg']:.2f}({fmt_delta(r['old_avg'], r['new_avg'])})"
        high = f"{r['old_high']:.1f}→{r['new_high']:.1f}%"
        low = f"{r['old_low']:.1f}→{r['new_low']:.1f}%"
        lines.append(f"  {name:<28} {reviews:>12} {avg:>14} {high:>12} {low:>12}")

    # Detailed per-hotel
    lines.append("")
    lines.append("■ ホテル別詳細")
    for r in results:
        lines.append("")
        lines.append(f"  【{r['hotel_name']}】")
        lines.append(f"    口コミ数: {r['old_reviews']} → {r['new_reviews']} ({'+' if r['review_delta'] >= 0 else ''}{r['review_delta']})")
        lines.append(f"    平均スコア(10pt): {r['old_avg']:.2f} → {r['new_avg']:.2f} ({fmt_delta(r['old_avg'], r['new_avg'])})")
        lines.append(f"    高評価率: {r['old_high']:.1f}% → {r['new_high']:.1f}% ({fmt_delta(r['old_high'], r['new_high'], '.1f')}%)")
        lines.append(f"    中評価率: {r['old_mid']:.1f}% → {r['new_mid']:.1f}% ({fmt_delta(r['old_mid'], r['new_mid'], '.1f')}%)")
        lines.append(f"    低評価率: {r['old_low']:.1f}% → {r['new_low']:.1f}% ({fmt_delta(r['old_low'], r['new_low'], '.1f')}%)")
        if r["site_changes"]:
            lines.append(f"    サイト別:")
            for sc in r["site_changes"]:
                count_delta = sc["new_count"] - sc["old_count"]
                avg_delta = sc["new_avg"] - sc["old_avg"]
                sign_c = "+" if count_delta >= 0 else ""
                sign_a = "+" if avg_delta >= 0 else ""
                lines.append(f"      {sc['site']:<16} 件数: {sc['old_count']}→{sc['new_count']}({sign_c}{count_delta})  平均: {sc['old_avg']:.2f}→{sc['new_avg']:.2f}({sign_a}{avg_delta:.2f})")

    # Portfolio comparison
    lines.append("")
    lines.append("■ ポートフォリオ分析比較")
    old_pf = load_json(os.path.join(BACKUP, "primechange_portfolio_analysis.json"))
    new_pf = load_json(os.path.join(BASE, "primechange_portfolio_analysis.json"))
    if old_pf and new_pf:
        old_ov = old_pf.get("portfolio_overview", {})
        new_ov = new_pf.get("portfolio_overview", {})
        lines.append(f"  総口コミ数: {old_ov.get('total_reviews', 'N/A')} → {new_ov.get('total_reviews', 'N/A')}")
        lines.append(f"  加重平均: {old_ov.get('avg_score', 0):.2f} → {new_ov.get('avg_score', 0):.2f}")
        best = new_ov.get('best_hotel', {})
        worst = new_ov.get('worst_hotel', {})
        lines.append(f"  最高: {best.get('name', 'N/A')} ({best.get('avg', 'N/A')})")
        lines.append(f"  最低: {worst.get('name', 'N/A')} ({worst.get('avg', 'N/A')})")

    lines.append("")
    lines.append("=" * 70)
    lines.append("レポート終了")
    lines.append("=" * 70)

    report = "\n".join(lines)
    print(report)

    with open(OUTPUT, "w", encoding="utf-8") as f:
        f.write(report)
    print(f"\n保存先: {OUTPUT}")


if __name__ == "__main__":
    main()
