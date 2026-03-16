#!/usr/bin/env python3
"""
分析1: クレーム類型×口コミスコア連動分析
19ホテルのXLSX 🔵クレームシートからクレーム類型別データを抽出し、
口コミスコアとの相関を分析。分析6の弾力性データを活用して収益インパクトを試算する。
"""

import sys, os, json
import numpy as np

sys.path.insert(0, os.path.dirname(__file__))
from hotel_xlsx_utils import HOTEL_FILES, open_workbook, safe_number, revenue_key

# ─── Constants ───────────────────────────────────────────────────────
CLAIM_COLS = {
    4: '巻き込み', 5: '誤入室', 6: 'ドア閉め忘れ', 7: '髪の毛',
    8: '残置', 9: 'セット漏れ', 10: '汚れ', 11: '手配ミス',
    12: '清掃不備', 13: '未清掃', 14: '破損', 16: '私物破棄', 18: 'その他'
}

CLAIM_CATEGORIES = {
    '客室準備系': ['残置', 'セット漏れ', '手配ミス'],
    '清潔性系': ['髪の毛', '汚れ', '清掃不備', '未清掃'],
    '安全・設備系': ['巻き込み', '誤入室', 'ドア閉め忘れ', '破損'],
    'その他': ['私物破棄', 'その他']
}

HOTEL_NAME_MAP = {
    'daiwa_osaki': 'ダイワロイネットホテル東京大崎',
    'chisan': 'チサンホテル浜松町',
    'hearton': 'ハートンホテル東品川',
    'keyakigate': 'ホテルケヤキゲート東京府中',
    'richmond_mejiro': 'リッチモンドホテル東京目白',
    'keisei_richmond': '京成リッチモンドホテル東京錦糸町',
    'daiichi_ikebukuro': '第一イン池袋',
    'comfort_roppongi': 'コンフォートイン六本木',
    'comfort_suites_tokyobay': 'コンフォートスイーツ東京ベイ',
    'comfort_era_higashikanda': 'コンフォートホテルERA東神田',
    'comfort_narita': 'コンフォートホテル成田',
    'comfort_yokohama': 'コンフォートホテル横浜関内',
    'apa_sagamihara': 'アパホテル相模原橋本駅東',
    'apa_kamata': 'アパホテル蒲田駅東',
    'court_shinyokohama': 'コートホテル新横浜',
    'comment_yokohama': 'ホテルコメント横浜関内',
    'henn_na_haneda': '変なホテル東京羽田',
    'kawasaki_nikko': '川崎日航ホテル',
    'comfort_hakata': 'コンフォートホテル博多',
}

OUT_DIR = os.path.dirname(__file__)


def extract_claim_data():
    """Extract claim data from all 19 hotels."""
    all_hotels = {}

    for hk in HOTEL_FILES:
        wb = open_workbook(hk)

        # Find claim sheet
        claim_sheet = None
        for s in wb.sheetnames:
            if 'クレーム' in s:
                claim_sheet = s
                break
        if not claim_sheet:
            print(f"  ⚠ {hk}: no claim sheet found, skipping")
            wb.close()
            continue

        ws = wb[claim_sheet]

        # Collect monthly data
        monthly_data = []
        type_totals = {name: 0 for name in CLAIM_COLS.values()}
        total_claims = 0
        total_rooms = 0

        for row_idx in range(10, 22):
            month_val = ws.cell(row=row_idx, column=3).value
            if not month_val:
                continue

            month_name = str(month_val).strip()
            month_claims = {}
            for col, name in CLAIM_COLS.items():
                v = safe_number(ws.cell(row=row_idx, column=col).value)
                month_claims[name] = v
                type_totals[name] += v

            row_total = safe_number(ws.cell(row=row_idx, column=20).value)
            rooms = safe_number(ws.cell(row=row_idx, column=22).value)
            rate = safe_number(ws.cell(row=row_idx, column=24).value)

            total_claims += row_total
            total_rooms += rooms

            monthly_data.append({
                'month': month_name,
                'claims_by_type': month_claims,
                'total_claims': row_total,
                'rooms_cleaned': rooms,
                'claim_rate': rate
            })

        overall_rate = total_claims / total_rooms if total_rooms > 0 else 0

        all_hotels[hk] = {
            'name': HOTEL_NAME_MAP.get(hk, hk),
            'monthly_data': monthly_data,
            'type_totals': {k: int(v) for k, v in type_totals.items()},
            'total_claims': int(total_claims),
            'total_rooms': int(total_rooms),
            'overall_claim_rate': overall_rate,
            'months_with_data': len([m for m in monthly_data if m['total_claims'] > 0])
        }

        wb.close()

    return all_hotels


def load_quality_scores():
    """Load quality scores from portfolio analysis JSON."""
    path = os.path.join(OUT_DIR, 'primechange_portfolio_analysis.json')
    with open(path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    scores = {}
    # Scores are in portfolio_overview.hotels_ranked (list of dicts)
    ranked = data.get('portfolio_overview', {}).get('hotels_ranked', [])
    for hotel in ranked:
        key = hotel.get('key', '')
        avg = hotel.get('avg', 0)
        name = hotel.get('name', key)
        if key and avg > 0:
            # Store under original key
            scores[key] = {'avg_score': avg, 'name': name}
            # Also store under mapped key (e.g., keisei_kinshicho → keisei_richmond)
            mapped = revenue_key(key)
            if mapped != key:
                scores[mapped] = {'avg_score': avg, 'name': name}
    return scores


def load_elasticity_data():
    """Load elasticity data from Analysis 6."""
    path = os.path.join(OUT_DIR, 'analysis_6_data.json')
    with open(path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    return data


def analyze_type_frequency(all_hotels):
    """Rank claim types by total frequency across all hotels."""
    grand_totals = {name: 0 for name in CLAIM_COLS.values()}
    total_all = 0

    for hk, hdata in all_hotels.items():
        for name, count in hdata['type_totals'].items():
            grand_totals[name] += count
            total_all += count

    ranking = sorted(grand_totals.items(), key=lambda x: x[1], reverse=True)

    result = []
    for name, count in ranking:
        result.append({
            'type': name,
            'total_count': count,
            'share_pct': round(count / total_all * 100, 1) if total_all > 0 else 0,
            'hotels_affected': sum(1 for hdata in all_hotels.values() if hdata['type_totals'].get(name, 0) > 0)
        })

    return result, total_all


def analyze_category_breakdown(all_hotels):
    """Group claim types into categories and analyze."""
    category_totals = {}
    total_all = sum(h['total_claims'] for h in all_hotels.values())

    for cat_name, types in CLAIM_CATEGORIES.items():
        cat_count = 0
        for hdata in all_hotels.values():
            for t in types:
                cat_count += hdata['type_totals'].get(t, 0)

        category_totals[cat_name] = {
            'count': cat_count,
            'share_pct': round(cat_count / total_all * 100, 1) if total_all > 0 else 0,
            'types': types
        }

    return category_totals


def analyze_correlation_with_scores(all_hotels, quality_scores):
    """Correlate claim rates with review scores."""
    # Build paired data: claim_rate vs score
    pairs = []
    type_pairs = {name: [] for name in CLAIM_COLS.values()}

    for hk, hdata in all_hotels.items():
        mapped = revenue_key(hk)
        score_data = quality_scores.get(hk) or quality_scores.get(mapped)
        if not score_data:
            continue

        score = score_data['avg_score']
        rate = hdata['overall_claim_rate']
        rooms = hdata['total_rooms']

        pairs.append({
            'hotel_key': hk,
            'name': hdata['name'],
            'score': score,
            'claim_rate': rate,
            'total_claims': hdata['total_claims'],
            'rooms': rooms
        })

        # Per-type rates
        for name in CLAIM_COLS.values():
            type_count = hdata['type_totals'].get(name, 0)
            type_rate = type_count / rooms * 10000 if rooms > 0 else 0  # per 10,000 rooms
            type_pairs[name].append({'score': score, 'rate': type_rate})

    if len(pairs) < 3:
        return {'error': 'insufficient data'}

    # Overall correlation: claim_rate vs score
    scores = np.array([p['score'] for p in pairs])
    rates = np.array([p['claim_rate'] for p in pairs])

    # Higher claim rate → lower score (expected negative correlation)
    r_overall = float(np.corrcoef(rates, scores)[0, 1]) if len(rates) > 2 else 0

    # Per-type correlations
    type_correlations = []
    for name in CLAIM_COLS.values():
        tp = type_pairs[name]
        s_arr = np.array([p['score'] for p in tp])
        r_arr = np.array([p['rate'] for p in tp])

        # Only compute if there's variance
        if np.std(r_arr) > 0 and len(r_arr) > 2:
            r_val = float(np.corrcoef(r_arr, s_arr)[0, 1])
        else:
            r_val = 0.0

        type_correlations.append({
            'type': name,
            'correlation_r': round(r_val, 4),
            'abs_r': round(abs(r_val), 4),
            'direction': '負（スコア低下）' if r_val < -0.1 else '正（スコア上昇）' if r_val > 0.1 else '弱い/なし',
            'interpretation': interpret_correlation(r_val)
        })

    # Sort by absolute correlation strength
    type_correlations.sort(key=lambda x: x['abs_r'], reverse=True)

    return {
        'overall_correlation': {
            'r': round(r_overall, 4),
            'r_squared': round(r_overall**2, 4),
            'interpretation': interpret_correlation(r_overall),
            'direction': 'クレーム率が高いほどスコアが低い傾向' if r_overall < 0 else 'クレーム率が高いほどスコアが高い傾向（予想外）'
        },
        'type_correlations': type_correlations,
        'data_points': pairs
    }


def interpret_correlation(r):
    """Interpret correlation coefficient."""
    a = abs(r)
    if a >= 0.7:
        return '強い相関'
    elif a >= 0.4:
        return '中程度の相関'
    elif a >= 0.2:
        return '弱い相関'
    else:
        return 'ほぼ無相関'


def analyze_hotel_profiles(all_hotels):
    """Create claim profile for each hotel (composition ratios)."""
    profiles = []

    for hk, hdata in all_hotels.items():
        total = hdata['total_claims']
        if total == 0:
            profiles.append({
                'hotel_key': hk,
                'name': hdata['name'],
                'total_claims': 0,
                'claim_rate': 0,
                'profile': 'クレームなし',
                'top_types': [],
                'category_breakdown': {}
            })
            continue

        # Type composition
        type_share = {}
        for name, count in hdata['type_totals'].items():
            if count > 0:
                type_share[name] = {
                    'count': count,
                    'share_pct': round(count / total * 100, 1)
                }

        top_types = sorted(type_share.items(), key=lambda x: x[1]['count'], reverse=True)[:3]

        # Category breakdown
        cat_breakdown = {}
        for cat_name, types in CLAIM_CATEGORIES.items():
            cat_count = sum(hdata['type_totals'].get(t, 0) for t in types)
            cat_breakdown[cat_name] = {
                'count': cat_count,
                'share_pct': round(cat_count / total * 100, 1)
            }

        # Determine dominant category
        dominant_cat = max(cat_breakdown.items(), key=lambda x: x[1]['count'])

        profiles.append({
            'hotel_key': hk,
            'name': hdata['name'],
            'total_claims': total,
            'claim_rate': round(hdata['overall_claim_rate'] * 10000, 2),  # per 10,000 rooms
            'rooms_cleaned': hdata['total_rooms'],
            'type_composition': type_share,
            'top_types': [{'type': t[0], 'count': t[1]['count'], 'share_pct': t[1]['share_pct']} for t in top_types],
            'category_breakdown': cat_breakdown,
            'dominant_category': dominant_cat[0],
            'profile_label': generate_profile_label(dominant_cat[0], top_types)
        })

    # Sort by claim rate descending
    profiles.sort(key=lambda x: x['claim_rate'], reverse=True)

    return profiles


def generate_profile_label(dominant_cat, top_types):
    """Generate a descriptive label for hotel claim profile."""
    if not top_types:
        return 'クレームなし'
    top_type = top_types[0][0]
    return f'{dominant_cat}集中型（主に{top_type}）'


def estimate_revenue_impact(all_hotels, quality_scores, elasticity_data):
    """Estimate revenue impact of reducing each claim type by 50%."""
    # Get elasticity: score improvement per 0.1pt → RevPAR improvement
    elast = elasticity_data.get('benchmark_comparison', {})
    self_elasticity_pct = elast.get('self_elasticity_pct_per_01', 2.48)  # default from analysis 6

    # Get hotel revenue data
    rev_path = os.path.join(OUT_DIR, 'hotel_revenue_data.json')
    with open(rev_path, 'r', encoding='utf-8') as f:
        rev_data = json.load(f)

    # Calculate correlation-weighted impact
    # Key insight: how much would reducing a claim type improve scores?

    # Overall: total claims=77, total rooms ~119,192
    # Overall claim rate ≈ 0.0646%
    # If we can estimate claim_rate → score relationship, we can chain it with score → revenue

    # Build claim_rate vs score regression
    pairs = []
    for hk, hdata in all_hotels.items():
        mapped = revenue_key(hk)
        score_data = quality_scores.get(hk) or quality_scores.get(mapped)
        if not score_data:
            continue

        rev_info = rev_data.get(hk) or rev_data.get(mapped)
        revpar = rev_info.get('revpar', 0) if rev_info else 0
        monthly_rev = rev_info.get('actual_revenue', 0) if rev_info else 0

        pairs.append({
            'hotel_key': hk,
            'name': hdata['name'],
            'claim_rate': hdata['overall_claim_rate'],
            'score': score_data['avg_score'],
            'revpar': revpar,
            'monthly_revenue': monthly_rev,
            'total_claims': hdata['total_claims'],
            'rooms': hdata['total_rooms']
        })

    if len(pairs) < 3:
        return {'error': 'insufficient data'}

    rates = np.array([p['claim_rate'] for p in pairs])
    scores = np.array([p['score'] for p in pairs])

    # Linear regression: score = a * claim_rate + b
    if np.std(rates) > 0:
        slope, intercept = np.polyfit(rates, scores, 1)
    else:
        slope, intercept = 0, np.mean(scores)

    # Scenario analysis: reduce all claims by 25%, 50%, 75%
    scenarios = []
    for reduction_pct in [25, 50, 75]:
        total_revenue_impact = 0
        hotel_impacts = []

        for p in pairs:
            if p['total_claims'] == 0:
                hotel_impacts.append({
                    'name': p['name'],
                    'current_claims': 0,
                    'reduced_claims': 0,
                    'score_improvement': 0,
                    'revpar_improvement_pct': 0,
                    'revenue_impact_monthly': 0
                })
                continue

            current_rate = p['claim_rate']
            reduced_rate = current_rate * (1 - reduction_pct / 100)
            rate_reduction = current_rate - reduced_rate

            # Score improvement from rate reduction
            score_improvement = abs(slope) * rate_reduction  # slope is negative, so abs

            # Revenue impact: score_improvement → RevPAR improvement
            # self_elasticity: 0.1pt → 2.48% RevPAR
            revpar_improvement_pct = (score_improvement / 0.1) * self_elasticity_pct / 100
            monthly_rev_impact = p['monthly_revenue'] * revpar_improvement_pct

            total_revenue_impact += monthly_rev_impact

            hotel_impacts.append({
                'name': p['name'],
                'current_claims': p['total_claims'],
                'reduced_claims': int(p['total_claims'] * (1 - reduction_pct / 100)),
                'score_improvement': round(score_improvement, 4),
                'revpar_improvement_pct': round(revpar_improvement_pct * 100, 3),
                'revenue_impact_monthly': round(monthly_rev_impact)
            })

        scenarios.append({
            'reduction_pct': reduction_pct,
            'total_monthly_impact': round(total_revenue_impact),
            'total_annual_impact': round(total_revenue_impact * 12),
            'hotel_impacts': sorted(hotel_impacts, key=lambda x: x.get('revenue_impact_monthly', 0), reverse=True)
        })

    # Per-type reduction impact (reduce each type by 50%)
    type_impacts = []
    for type_name in CLAIM_COLS.values():
        type_total = sum(h['type_totals'].get(type_name, 0) for h in all_hotels.values())
        if type_total == 0:
            type_impacts.append({
                'type': type_name,
                'current_count': 0,
                'reduced_count': 0,
                'score_improvement': 0,
                'revenue_impact_monthly': 0
            })
            continue

        total_rooms = sum(h['total_rooms'] for h in all_hotels.values())
        type_rate = type_total / total_rooms if total_rooms > 0 else 0
        reduced_rate = type_rate * 0.5
        rate_reduction = type_rate - reduced_rate

        score_improvement = abs(slope) * rate_reduction
        revpar_improvement_pct = (score_improvement / 0.1) * self_elasticity_pct / 100

        # Apply to total revenue
        total_monthly_rev = sum(p['monthly_revenue'] for p in pairs)
        rev_impact = total_monthly_rev * revpar_improvement_pct

        type_impacts.append({
            'type': type_name,
            'current_count': int(type_total),
            'reduced_count': int(type_total * 0.5),
            'rate_reduction': round(rate_reduction * 10000, 4),
            'score_improvement': round(score_improvement, 4),
            'revenue_impact_monthly': round(rev_impact)
        })

    type_impacts.sort(key=lambda x: x['revenue_impact_monthly'], reverse=True)

    return {
        'regression': {
            'slope': round(slope, 4),
            'intercept': round(intercept, 4),
            'interpretation': f'クレーム率が1%上昇すると、スコアが{abs(slope):.2f}点{"低下" if slope < 0 else "上昇"}する推定'
        },
        'scenarios': scenarios,
        'type_impact_ranking': type_impacts,
        'elasticity_used': {
            'source': '分析6',
            'value': f'0.1点→RevPAR {self_elasticity_pct}%改善'
        }
    }


def compute_improvement_priorities(hotel_profiles, correlation_data):
    """Rank improvement priorities combining frequency, correlation, and spread."""
    type_corrs = {tc['type']: tc for tc in correlation_data.get('type_correlations', [])}

    priorities = []
    for tc in correlation_data.get('type_correlations', []):
        type_name = tc['type']

        # Get frequency data
        total_count = 0
        hotels_affected = 0
        for p in hotel_profiles:
            tc_comp = p.get('type_composition', {})
            if type_name in tc_comp:
                total_count += tc_comp[type_name]['count']
                hotels_affected += 1

        if total_count == 0 and tc['abs_r'] == 0:
            continue

        # Priority score: combine frequency, correlation strength, and spread
        freq_score = min(total_count / 5, 1.0)  # normalize: 5+ claims = max
        corr_score = tc['abs_r']  # 0-1
        spread_score = min(hotels_affected / 10, 1.0)  # 10+ hotels = max

        composite_score = freq_score * 0.4 + corr_score * 0.4 + spread_score * 0.2

        priorities.append({
            'type': type_name,
            'total_count': total_count,
            'hotels_affected': hotels_affected,
            'correlation_r': tc['correlation_r'],
            'abs_correlation': tc['abs_r'],
            'frequency_score': round(freq_score, 3),
            'correlation_score': round(corr_score, 3),
            'spread_score': round(spread_score, 3),
            'composite_priority': round(composite_score, 3),
            'priority_rank': 0  # will be assigned after sorting
        })

    priorities.sort(key=lambda x: x['composite_priority'], reverse=True)
    for i, p in enumerate(priorities):
        p['priority_rank'] = i + 1

    return priorities


def generate_recommendations(priorities, category_totals, scenarios):
    """Generate strategic recommendations based on analysis."""
    recs = []

    # Top 3 priority types
    top3 = priorities[:3] if len(priorities) >= 3 else priorities

    if not top3:
        # Fallback: use type_ranking from frequency data
        recs.append({
            'title': '重点改善クレーム類型（頻度ベース）',
            'rationale': '相関分析データ不足のため、頻度ベースで優先度を設定。',
            'actions': ['クレーム発生頻度の高い類型から順次対策を実施', '改善効果の月次トラッキングとPDCAサイクル確立'],
            'priority': '最優先'
        })
    else:
        # Recommendation 1: Focus on top claim types
        top_types_str = '、'.join(p['type'] for p in top3)
        actions = [f'{top3[0]["type"]}対策: チェックリスト強化と二重確認体制の導入']
        if len(top3) > 1:
            actions.append(f'{top3[1]["type"]}対策: 研修プログラムへの組み込みと月次モニタリング')
        actions.append('改善効果の月次トラッキングとPDCAサイクル確立')
        recs.append({
            'title': f'重点改善クレーム類型: {top_types_str}',
            'rationale': f'頻度・スコア影響度・波及度の複合評価で上位3類型を特定。全クレームの{sum(p["total_count"] for p in top3)}件をカバー。',
            'actions': actions,
            'priority': '最優先'
        })

    # Recommendation 2: Category-based approach
    top_cat = max(category_totals.items(), key=lambda x: x[1]['count'])
    recs.append({
        'title': f'カテゴリ別対策: {top_cat[0]}（全体の{top_cat[1]["share_pct"]}%）',
        'rationale': f'{top_cat[0]}カテゴリが最多で、構成比{top_cat[1]["share_pct"]}%。類型{", ".join(CLAIM_CATEGORIES[top_cat[0]])}への体系的対策が効果的。',
        'actions': [
            f'{top_cat[0]}に関するSOPの見直しと標準化',
            'ベストプラクティスホテルからの手法移植',
            'カテゴリ別KPI設定と週次レビュー'
        ],
        'priority': '高'
    })

    # Recommendation 3: Revenue-linked targets
    if scenarios:
        s50 = next((s for s in scenarios if s['reduction_pct'] == 50), None)
        if s50:
            recs.append({
                'title': f'収益連動目標: クレーム50%削減で年間¥{s50["total_annual_impact"]:,}増収',
                'rationale': '分析6の品質→売上弾力性を活用し、クレーム削減→スコア改善→売上増の連鎖効果を定量化。',
                'actions': [
                    '四半期ごとのクレーム削減目標の設定',
                    'クレーム率と口コミスコアの連動ダッシュボード構築',
                    '改善効果の可視化によるスタッフモチベーション向上'
                ],
                'priority': '高'
            })

    # Recommendation 4: Zero-claim hotels as benchmarks
    zero_hotels = [p for p in hotel_profiles_global if p['total_claims'] == 0]
    if zero_hotels:
        names = '、'.join(h['name'] for h in zero_hotels[:3])
        recs.append({
            'title': f'ゼロクレームホテルの横展開（{len(zero_hotels)}ホテル）',
            'rationale': f'{names}等がクレームゼロを達成。これらのオペレーション手法を他ホテルへ展開。',
            'actions': [
                'ゼロクレームホテルの成功要因インタビュー',
                '標準手順書の作成と全ホテルへの展開',
                'メンター制度による現場レベルでのナレッジ共有'
            ],
            'priority': '中'
        })

    return recs


# ─── Main ────────────────────────────────────────────────────────────
def main():
    print("分析1: クレーム類型×口コミスコア連動分析")
    print("=" * 50)

    # 1. Extract claim data
    print("\n[1/6] クレームデータ抽出中...")
    all_hotels = extract_claim_data()
    total_claims = sum(h['total_claims'] for h in all_hotels.values())
    total_rooms = sum(h['total_rooms'] for h in all_hotels.values())
    print(f"  → 19ホテル、総クレーム{total_claims}件、総清掃客室{total_rooms:,}室")

    # 2. Load quality scores
    print("\n[2/6] 品質スコア読込中...")
    quality_scores = load_quality_scores()
    print(f"  → {len(quality_scores)}ホテルのスコアデータ")

    # 3. Load elasticity data
    print("\n[3/6] 弾力性データ読込中...")
    elasticity_data = load_elasticity_data()
    print(f"  → 分析6データ読込完了")

    # 4. Type frequency analysis
    print("\n[4/6] 類型別頻度分析中...")
    type_ranking, grand_total = analyze_type_frequency(all_hotels)
    category_totals = analyze_category_breakdown(all_hotels)
    print(f"  → 13類型をランキング（全{grand_total}件）")
    for t in type_ranking[:5]:
        print(f"    {t['type']}: {t['total_count']}件 ({t['share_pct']}%)")

    # 5. Correlation analysis
    print("\n[5/6] スコア相関分析中...")
    correlation_data = analyze_correlation_with_scores(all_hotels, quality_scores)
    r = correlation_data.get('overall_correlation', {}).get('r', 0)
    print(f"  → 全体相関: r={r:.4f} ({correlation_data.get('overall_correlation', {}).get('interpretation', '')})")

    # 6. Hotel profiles
    print("\n[6/6] ホテル別プロファイル作成中...")
    global hotel_profiles_global
    hotel_profiles = analyze_hotel_profiles(all_hotels)
    hotel_profiles_global = hotel_profiles
    print(f"  → {len(hotel_profiles)}ホテルのプロファイル生成")

    # Revenue impact estimation
    print("\n[+] 収益インパクト試算中...")
    revenue_impact = estimate_revenue_impact(all_hotels, quality_scores, elasticity_data)

    # Priority ranking
    print("\n[+] 改善優先度算出中...")
    priorities = compute_improvement_priorities(hotel_profiles, correlation_data)

    # Recommendations
    print("\n[+] 提言作成中...")
    scenarios = revenue_impact.get('scenarios', [])
    recommendations = generate_recommendations(priorities, category_totals, scenarios)

    # ─── Build output JSON ───
    output = {
        'analysis_metadata': {
            'analysis_id': 1,
            'title': 'クレーム類型×口コミスコア連動分析',
            'subtitle': 'Claim Type × Review Score Correlation Analysis',
            'total_hotels': 19,
            'total_claims': int(grand_total),
            'total_rooms_cleaned': int(total_rooms),
            'overall_claim_rate': round(grand_total / total_rooms * 100, 4) if total_rooms > 0 else 0,
            'claim_types_analyzed': 13,
            'data_period': 'R8年度（2月〜1月）'
        },
        'type_frequency_ranking': type_ranking,
        'category_breakdown': category_totals,
        'correlation_analysis': correlation_data,
        'hotel_profiles': hotel_profiles,
        'revenue_impact': revenue_impact,
        'improvement_priorities': priorities,
        'recommendations': recommendations,
        'summary_stats': {
            'hotels_with_claims': sum(1 for h in all_hotels.values() if h['total_claims'] > 0),
            'hotels_zero_claims': sum(1 for h in all_hotels.values() if h['total_claims'] == 0),
            'avg_claim_rate': round(np.mean([h['overall_claim_rate'] for h in all_hotels.values()]) * 10000, 2),
            'max_claim_rate_hotel': max(hotel_profiles, key=lambda x: x['claim_rate'])['name'] if hotel_profiles else '',
            'max_claim_rate': max(hotel_profiles, key=lambda x: x['claim_rate'])['claim_rate'] if hotel_profiles else 0,
            'top_claim_type': type_ranking[0]['type'] if type_ranking else '',
            'top_claim_count': type_ranking[0]['total_count'] if type_ranking else 0
        }
    }

    out_path = os.path.join(OUT_DIR, 'analysis_1_data.json')
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(output, f, ensure_ascii=False, indent=2)

    print(f"\n✅ 分析完了 → {out_path} ({os.path.getsize(out_path):,} bytes)")

    # Summary
    print(f"\n===== 分析1 サマリー =====")
    print(f"総クレーム: {grand_total}件 / {total_rooms:,}室 = {grand_total/total_rooms*10000:.1f}件/万室")
    print(f"クレーム有ホテル: {output['summary_stats']['hotels_with_claims']}/19")
    print(f"ゼロクレームホテル: {output['summary_stats']['hotels_zero_claims']}ホテル")
    print(f"最多類型: {type_ranking[0]['type']}（{type_ranking[0]['total_count']}件, {type_ranking[0]['share_pct']}%）")
    print(f"全体相関: r={r:.4f}")
    if scenarios:
        s50 = next((s for s in scenarios if s['reduction_pct'] == 50), None)
        if s50:
            print(f"50%削減→年間インパクト: ¥{s50['total_annual_impact']:,}")


hotel_profiles_global = []

if __name__ == '__main__':
    main()
