#!/usr/bin/env python3
"""
分析3: 人員配置×品質 相関分析
分析4: 清掃完了時間×品質 分析
③日報シートから日次データを抽出し、スタッフ配置・完了時間と品質指標の関連を分析。
"""

import sys, os, json
import numpy as np

sys.path.insert(0, os.path.dirname(__file__))
from hotel_xlsx_utils import HOTEL_FILES, open_workbook, safe_number, safe_time, revenue_key

OUT_DIR = os.path.dirname(__file__)

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


def extract_daily_data():
    """Extract daily report data from all 19 hotels."""
    all_hotels = {}

    for hk in HOTEL_FILES:
        wb = open_workbook(hk)
        daily_sheet = None
        for s in wb.sheetnames:
            if '日報' in s:
                daily_sheet = s
                break
        if not daily_sheet:
            wb.close()
            continue

        ws = wb[daily_sheet]
        daily_entries = []

        for row_idx in range(9, 500):
            date_val = ws.cell(row=row_idx, column=3).value
            if date_val is None:
                continue
            date_str = str(date_val)
            if '202' not in date_str:
                continue

            maid = safe_number(ws.cell(row=row_idx, column=10).value, None)
            checker = safe_number(ws.cell(row=row_idx, column=11).value, None)
            time_v = safe_time(ws.cell(row=row_idx, column=9).value)
            claim = safe_number(ws.cell(row=row_idx, column=8).value, 0)
            workload_raw = ws.cell(row=row_idx, column=14).value
            workload = safe_number(workload_raw, None) if workload_raw else None
            # Some workload values are text like "残 21\n滞在 49\nアウト150\nフル 3"
            if workload is None and workload_raw and isinstance(workload_raw, str):
                # Try to extract total from text
                import re
                nums = re.findall(r'(\d+)', str(workload_raw))
                if nums:
                    workload = sum(int(n) for n in nums)

            entry = {
                'date': date_str[:10],
                'maids': maid if maid and maid > 0 else None,
                'checkers': checker if checker and checker > 0 else None,
                'total_staff': (maid or 0) + (checker or 0) if (maid and maid > 0) or (checker and checker > 0) else None,
                'completion_time': round(time_v, 2) if time_v and time_v > 6 else None,
                'claims': claim,
                'workload': workload if workload and workload > 0 else None,
            }

            # Calculate productivity metrics
            if entry['maids'] and entry['workload'] and entry['workload'] > 0:
                entry['rooms_per_maid'] = round(entry['workload'] / entry['maids'], 1)
            else:
                entry['rooms_per_maid'] = None

            if entry['checkers'] and entry['checkers'] > 0 and entry['workload'] and entry['workload'] > 0:
                entry['rooms_per_checker'] = round(entry['workload'] / entry['checkers'], 1)
            else:
                entry['rooms_per_checker'] = None

            daily_entries.append(entry)

        # Hotel-level aggregates
        maids_data = [e['maids'] for e in daily_entries if e['maids']]
        checkers_data = [e['checkers'] for e in daily_entries if e['checkers']]
        times_data = [e['completion_time'] for e in daily_entries if e['completion_time']]
        claims_data = [e['claims'] for e in daily_entries if e['claims'] > 0]
        rpm_data = [e['rooms_per_maid'] for e in daily_entries if e['rooms_per_maid']]
        workload_data = [e['workload'] for e in daily_entries if e['workload']]

        all_hotels[hk] = {
            'name': HOTEL_NAME_MAP.get(hk, hk),
            'total_days': len(daily_entries),
            'daily_entries': daily_entries,
            'staffing': {
                'maid_data_points': len(maids_data),
                'avg_maids': round(np.mean(maids_data), 1) if maids_data else None,
                'min_maids': min(maids_data) if maids_data else None,
                'max_maids': max(maids_data) if maids_data else None,
                'checker_data_points': len(checkers_data),
                'avg_checkers': round(np.mean(checkers_data), 1) if checkers_data else None,
                'avg_total_staff': round(np.mean(maids_data) + np.mean(checkers_data), 1) if maids_data and checkers_data else None,
                'maid_checker_ratio': round(np.mean(maids_data) / np.mean(checkers_data), 1) if maids_data and checkers_data and np.mean(checkers_data) > 0 else None,
            },
            'completion_time': {
                'data_points': len(times_data),
                'avg_time': round(np.mean(times_data), 2) if times_data else None,
                'min_time': round(min(times_data), 2) if times_data else None,
                'max_time': round(max(times_data), 2) if times_data else None,
                'std_time': round(np.std(times_data), 2) if len(times_data) > 1 else None,
            },
            'productivity': {
                'avg_rooms_per_maid': round(np.mean(rpm_data), 1) if rpm_data else None,
                'avg_workload': round(np.mean(workload_data), 0) if workload_data else None,
            },
            'claims_in_period': sum(e['claims'] for e in daily_entries),
        }

        wb.close()

    return all_hotels


def load_quality_and_revenue():
    """Load quality scores and revenue data."""
    with open(os.path.join(OUT_DIR, 'primechange_portfolio_analysis.json'), 'r') as f:
        portfolio = json.load(f)

    scores = {}
    for hotel in portfolio.get('portfolio_overview', {}).get('hotels_ranked', []):
        key = hotel.get('key', '')
        scores[key] = {'avg_score': hotel.get('avg', 0), 'name': hotel.get('name', key),
                        'cleaning_issue_rate': hotel.get('cleaning_issue_rate', 0)}
        mapped = revenue_key(key)
        if mapped != key:
            scores[mapped] = scores[key]

    with open(os.path.join(OUT_DIR, 'hotel_revenue_data.json'), 'r') as f:
        revenue = json.load(f)

    with open(os.path.join(OUT_DIR, 'analysis_1_data.json'), 'r') as f:
        claims = json.load(f)

    claim_rates = {}
    for hp in claims.get('hotel_profiles', []):
        claim_rates[hp['hotel_key']] = hp['claim_rate']  # per 10,000 rooms

    return scores, revenue, claim_rates


def analyze_staffing_vs_quality(all_hotels, scores, claim_rates):
    """Analysis 3: Staffing levels vs quality metrics."""
    pairs = []

    for hk, hdata in all_hotels.items():
        mapped = revenue_key(hk)
        score_data = scores.get(hk) or scores.get(mapped)
        if not score_data:
            continue

        staffing = hdata['staffing']
        if not staffing['avg_maids']:
            continue

        cr = claim_rates.get(hk, 0)

        pairs.append({
            'hotel_key': hk,
            'name': hdata['name'],
            'avg_maids': staffing['avg_maids'],
            'avg_checkers': staffing['avg_checkers'],
            'avg_total_staff': staffing['avg_total_staff'],
            'maid_checker_ratio': staffing['maid_checker_ratio'],
            'score': score_data['avg_score'],
            'cleaning_issue_rate': score_data['cleaning_issue_rate'],
            'claim_rate': cr,
            'maid_data_points': staffing['maid_data_points'],
        })

    if len(pairs) < 3:
        return {'error': 'insufficient staffing data', 'pairs': pairs}

    # Correlations
    results = {}

    # 1. Total staff vs score
    staff = np.array([p['avg_total_staff'] for p in pairs if p['avg_total_staff']])
    sc = np.array([p['score'] for p in pairs if p['avg_total_staff']])
    if len(staff) > 2 and np.std(staff) > 0:
        r = float(np.corrcoef(staff, sc)[0, 1])
        results['staff_vs_score'] = {'r': round(r, 4), 'r_squared': round(r**2, 4), 'n': len(staff)}

    # 2. Maid count vs score
    maids = np.array([p['avg_maids'] for p in pairs])
    sc_m = np.array([p['score'] for p in pairs])
    if len(maids) > 2 and np.std(maids) > 0:
        r = float(np.corrcoef(maids, sc_m)[0, 1])
        results['maids_vs_score'] = {'r': round(r, 4), 'r_squared': round(r**2, 4), 'n': len(maids)}

    # 3. Checker count vs score
    pairs_with_checkers = [p for p in pairs if p['avg_checkers'] and p['avg_checkers'] > 0]
    if len(pairs_with_checkers) > 2:
        ch = np.array([p['avg_checkers'] for p in pairs_with_checkers])
        sc_c = np.array([p['score'] for p in pairs_with_checkers])
        if np.std(ch) > 0:
            r = float(np.corrcoef(ch, sc_c)[0, 1])
            results['checkers_vs_score'] = {'r': round(r, 4), 'r_squared': round(r**2, 4), 'n': len(ch)}

    # 4. Maid-checker ratio vs score
    pairs_with_ratio = [p for p in pairs if p['maid_checker_ratio'] and p['maid_checker_ratio'] > 0]
    if len(pairs_with_ratio) > 2:
        ratio = np.array([p['maid_checker_ratio'] for p in pairs_with_ratio])
        sc_r = np.array([p['score'] for p in pairs_with_ratio])
        if np.std(ratio) > 0:
            r = float(np.corrcoef(ratio, sc_r)[0, 1])
            results['ratio_vs_score'] = {'r': round(r, 4), 'r_squared': round(r**2, 4), 'n': len(ratio)}

    # 5. Staff vs claim rate
    pairs_with_claims = [p for p in pairs if p['claim_rate'] is not None]
    if len(pairs_with_claims) > 2:
        st = np.array([p['avg_total_staff'] or p['avg_maids'] for p in pairs_with_claims])
        cr = np.array([p['claim_rate'] for p in pairs_with_claims])
        if np.std(st) > 0 and np.std(cr) > 0:
            r = float(np.corrcoef(st, cr)[0, 1])
            results['staff_vs_claims'] = {'r': round(r, 4), 'n': len(st)}

    # Staffing tiers analysis
    pairs_sorted = sorted(pairs, key=lambda x: x['avg_maids'])
    mid = len(pairs_sorted) // 2
    low_staff = pairs_sorted[:mid]
    high_staff = pairs_sorted[mid:]

    tiers = {
        'low_staff': {
            'count': len(low_staff),
            'avg_maids': round(np.mean([p['avg_maids'] for p in low_staff]), 1),
            'avg_score': round(np.mean([p['score'] for p in low_staff]), 2),
            'avg_claim_rate': round(np.mean([p['claim_rate'] for p in low_staff]), 2),
        },
        'high_staff': {
            'count': len(high_staff),
            'avg_maids': round(np.mean([p['avg_maids'] for p in high_staff]), 1),
            'avg_score': round(np.mean([p['score'] for p in high_staff]), 2),
            'avg_claim_rate': round(np.mean([p['claim_rate'] for p in high_staff]), 2),
        }
    }

    # Optimal staffing analysis
    best_hotels = sorted(pairs, key=lambda x: x['score'], reverse=True)[:5]
    worst_hotels = sorted(pairs, key=lambda x: x['score'])[:5]

    optimal = {
        'top5_hotels': [{
            'name': h['name'], 'avg_maids': h['avg_maids'],
            'avg_checkers': h['avg_checkers'], 'ratio': h['maid_checker_ratio'],
            'score': h['score']
        } for h in best_hotels],
        'bottom5_hotels': [{
            'name': h['name'], 'avg_maids': h['avg_maids'],
            'avg_checkers': h['avg_checkers'], 'ratio': h['maid_checker_ratio'],
            'score': h['score']
        } for h in worst_hotels],
        'top5_avg_ratio': round(np.mean([h['maid_checker_ratio'] for h in best_hotels if h['maid_checker_ratio']]), 1) if any(h['maid_checker_ratio'] for h in best_hotels) else None,
        'bottom5_avg_ratio': round(np.mean([h['maid_checker_ratio'] for h in worst_hotels if h['maid_checker_ratio']]), 1) if any(h['maid_checker_ratio'] for h in worst_hotels) else None,
    }

    return {
        'correlations': results,
        'staffing_tiers': tiers,
        'optimal_staffing': optimal,
        'data_points': pairs,
        'total_hotels_analyzed': len(pairs),
    }


def analyze_completion_time_vs_quality(all_hotels, scores, claim_rates):
    """Analysis 4: Completion time vs quality metrics."""
    pairs = []

    for hk, hdata in all_hotels.items():
        mapped = revenue_key(hk)
        score_data = scores.get(hk) or scores.get(mapped)
        if not score_data:
            continue

        ct = hdata['completion_time']
        if not ct['avg_time']:
            continue

        cr = claim_rates.get(hk, 0)

        pairs.append({
            'hotel_key': hk,
            'name': hdata['name'],
            'avg_completion_time': ct['avg_time'],
            'min_time': ct['min_time'],
            'max_time': ct['max_time'],
            'std_time': ct['std_time'],
            'time_data_points': ct['data_points'],
            'score': score_data['avg_score'],
            'cleaning_issue_rate': score_data['cleaning_issue_rate'],
            'claim_rate': cr,
        })

    if len(pairs) < 3:
        return {'error': 'insufficient time data', 'pairs': pairs}

    results = {}

    # 1. Completion time vs score
    times = np.array([p['avg_completion_time'] for p in pairs])
    sc = np.array([p['score'] for p in pairs])
    if np.std(times) > 0:
        r = float(np.corrcoef(times, sc)[0, 1])
        slope, intercept = np.polyfit(times, sc, 1)
        results['time_vs_score'] = {
            'r': round(r, 4), 'r_squared': round(r**2, 4),
            'slope': round(slope, 4), 'intercept': round(intercept, 4),
            'n': len(times),
            'interpretation': f'完了時間が1時間遅くなると、スコアが{abs(slope):.3f}点{"低下" if slope < 0 else "上昇"}'
        }

    # 2. Completion time vs claim rate
    cr_arr = np.array([p['claim_rate'] for p in pairs])
    if np.std(times) > 0 and np.std(cr_arr) > 0:
        r = float(np.corrcoef(times, cr_arr)[0, 1])
        results['time_vs_claims'] = {'r': round(r, 4), 'n': len(times)}

    # 3. Time variability vs quality
    pairs_with_std = [p for p in pairs if p['std_time'] and p['std_time'] > 0]
    if len(pairs_with_std) > 2:
        stds = np.array([p['std_time'] for p in pairs_with_std])
        sc_s = np.array([p['score'] for p in pairs_with_std])
        if np.std(stds) > 0:
            r = float(np.corrcoef(stds, sc_s)[0, 1])
            results['time_variability_vs_score'] = {'r': round(r, 4), 'n': len(stds)}

    # Time tier analysis
    pairs_sorted = sorted(pairs, key=lambda x: x['avg_completion_time'])
    third = max(1, len(pairs_sorted) // 3)

    early = pairs_sorted[:third]
    mid = pairs_sorted[third:third*2]
    late = pairs_sorted[third*2:]

    tiers = {
        'early_finish': {
            'label': '早期完了（〜14:50）',
            'count': len(early),
            'avg_time': round(np.mean([p['avg_completion_time'] for p in early]), 2),
            'avg_score': round(np.mean([p['score'] for p in early]), 2),
            'avg_claim_rate': round(np.mean([p['claim_rate'] for p in early]), 2),
            'hotels': [p['name'] for p in early],
        },
        'mid_finish': {
            'label': '標準完了（14:50〜15:30）',
            'count': len(mid),
            'avg_time': round(np.mean([p['avg_completion_time'] for p in mid]), 2),
            'avg_score': round(np.mean([p['score'] for p in mid]), 2),
            'avg_claim_rate': round(np.mean([p['claim_rate'] for p in mid]), 2),
            'hotels': [p['name'] for p in mid],
        },
        'late_finish': {
            'label': '遅延完了（15:30〜）',
            'count': len(late),
            'avg_time': round(np.mean([p['avg_completion_time'] for p in late]), 2),
            'avg_score': round(np.mean([p['score'] for p in late]), 2),
            'avg_claim_rate': round(np.mean([p['claim_rate'] for p in late]), 2),
            'hotels': [p['name'] for p in late],
        }
    }

    # Benchmark analysis: fastest vs slowest
    fastest_5 = sorted(pairs, key=lambda x: x['avg_completion_time'])[:5]
    slowest_5 = sorted(pairs, key=lambda x: x['avg_completion_time'], reverse=True)[:5]

    benchmark = {
        'fastest_5': [{'name': h['name'], 'time': h['avg_completion_time'], 'score': h['score']} for h in fastest_5],
        'slowest_5': [{'name': h['name'], 'time': h['avg_completion_time'], 'score': h['score']} for h in slowest_5],
        'fastest_avg_score': round(np.mean([h['score'] for h in fastest_5]), 2),
        'slowest_avg_score': round(np.mean([h['score'] for h in slowest_5]), 2),
        'score_difference': round(np.mean([h['score'] for h in fastest_5]) - np.mean([h['score'] for h in slowest_5]), 2),
    }

    return {
        'correlations': results,
        'time_tiers': tiers,
        'benchmark': benchmark,
        'data_points': pairs,
        'total_hotels_analyzed': len(pairs),
        'overall_avg_time': round(np.mean(times), 2),
        'overall_time_range': f'{round(min(times), 1)}〜{round(max(times), 1)}時',
    }


def generate_staffing_recommendations(staffing_analysis):
    """Generate recommendations for Analysis 3."""
    corrs = staffing_analysis.get('correlations', {})
    optimal = staffing_analysis.get('optimal_staffing', {})

    recs = []

    # Checker ratio recommendation
    ratio_r = corrs.get('ratio_vs_score', {}).get('r', 0)
    top_ratio = optimal.get('top5_avg_ratio')
    bottom_ratio = optimal.get('bottom5_avg_ratio')

    recs.append({
        'title': 'メイド・チェッカー比率の最適化',
        'rationale': f'上位5ホテルの平均メイド:チェッカー比率は{top_ratio or "N/A"}:1、下位5ホテルは{bottom_ratio or "N/A"}:1。チェッカー配置の適正化がスコア向上に寄与する可能性。',
        'actions': [
            'メイド3〜5名に対しチェッカー1名の配置基準を設定',
            '高稼働日（150室超）は追加チェッカーの配置を検討',
            'チェッカー不在ホテルへの配置導入'
        ],
        'priority': '高'
    })

    # Staffing level recommendation
    checker_r = corrs.get('checkers_vs_score', {}).get('r', 0)
    recs.append({
        'title': 'チェッカー配置とスコアの連動強化',
        'rationale': f'チェッカー数とスコアの相関r={checker_r}。チェッカーによる品質検査がスコア改善の鍵。',
        'actions': [
            'チェッカー配置率100%を目標に採用・配置計画を策定',
            'チェッカー研修プログラムの統一と強化',
            'チェッカー不在時の代替チェック体制の確立'
        ],
        'priority': '最優先'
    })

    recs.append({
        'title': 'データ入力の標準化と充実',
        'rationale': '日報のメイド数・チェッカー数の入力率がホテルにより大きく異なる。正確な分析のためにデータ品質向上が必要。',
        'actions': [
            '日報必須入力項目の明確化（メイド数・チェッカー数・完了時間）',
            '翌日稼働数の数値入力フォーマット統一',
            '月次での入力率チェックと改善フィードバック'
        ],
        'priority': '中'
    })

    return recs


def generate_time_recommendations(time_analysis):
    """Generate recommendations for Analysis 4."""
    corrs = time_analysis.get('correlations', {})
    benchmark = time_analysis.get('benchmark', {})
    tiers = time_analysis.get('time_tiers', {})

    recs = []

    time_r = corrs.get('time_vs_score', {}).get('r', 0)

    recs.append({
        'title': '清掃完了時間の目標設定',
        'rationale': f'完了時間とスコアの相関r={time_r}。全19ホテル平均{time_analysis.get("overall_avg_time", "N/A")}時。早期完了グループは平均スコア{tiers.get("early_finish", {}).get("avg_score", "N/A")}。',
        'actions': [
            '全ホテル15:00完了を標準目標として設定',
            '遅延傾向ホテルへの原因分析と改善支援',
            '完了時間トラッキングの日次モニタリング導入'
        ],
        'priority': '高'
    })

    recs.append({
        'title': '完了時間のばらつき削減',
        'rationale': '完了時間の標準偏差が大きいホテルは品質の安定性にリスク。安定した完了時間はゲスト体験の向上に寄与。',
        'actions': [
            '清掃手順の標準化とタイムスケジュール管理',
            '高稼働日の事前人員配置計画の策定',
            '完了時間の週次レビューと改善アクション'
        ],
        'priority': '中'
    })

    score_diff = benchmark.get('score_difference', 0)
    recs.append({
        'title': f'ベンチマーク活用（スコア差{score_diff}点）',
        'rationale': f'最速5ホテルの平均スコア{benchmark.get("fastest_avg_score", "N/A")} vs 最遅5ホテル{benchmark.get("slowest_avg_score", "N/A")}（差{score_diff}点）。',
        'actions': [
            '早期完了ホテルのオペレーション手法を調査',
            '効率化ベストプラクティスの横展開',
            '改善効果の月次追跡'
        ],
        'priority': '中'
    })

    return recs


def main():
    print("分析3＆4: 人員配置×品質 / 清掃完了時間×品質")
    print("=" * 55)

    # Extract daily data
    print("\n[1/4] 日報データ抽出中...")
    all_hotels = extract_daily_data()
    total_days = sum(h['total_days'] for h in all_hotels.values())
    print(f"  → {len(all_hotels)}ホテル、{total_days}日分のデータ")

    # Load quality/revenue data
    print("\n[2/4] 品質・売上データ読込中...")
    scores, revenue, claim_rates = load_quality_and_revenue()

    # Analysis 3: Staffing
    print("\n[3/4] 分析3: 人員配置×品質 分析中...")
    staffing_analysis = analyze_staffing_vs_quality(all_hotels, scores, claim_rates)
    n3 = staffing_analysis.get('total_hotels_analyzed', 0)
    print(f"  → {n3}ホテルで分析完了")
    for key, val in staffing_analysis.get('correlations', {}).items():
        print(f"    {key}: r={val.get('r', 'N/A')}")

    staffing_recs = generate_staffing_recommendations(staffing_analysis)

    # Analysis 4: Completion time
    print("\n[4/4] 分析4: 清掃完了時間×品質 分析中...")
    time_analysis = analyze_completion_time_vs_quality(all_hotels, scores, claim_rates)
    n4 = time_analysis.get('total_hotels_analyzed', 0)
    print(f"  → {n4}ホテルで分析完了")
    for key, val in time_analysis.get('correlations', {}).items():
        print(f"    {key}: r={val.get('r', 'N/A')}")

    time_recs = generate_time_recommendations(time_analysis)

    # Hotel summary table
    hotel_summary = []
    for hk, hdata in all_hotels.items():
        mapped = revenue_key(hk)
        score_data = scores.get(hk) or scores.get(mapped) or {}

        hotel_summary.append({
            'hotel_key': hk,
            'name': hdata['name'],
            'days': hdata['total_days'],
            'avg_maids': hdata['staffing']['avg_maids'],
            'avg_checkers': hdata['staffing']['avg_checkers'],
            'maid_checker_ratio': hdata['staffing']['maid_checker_ratio'],
            'avg_completion_time': hdata['completion_time']['avg_time'],
            'time_std': hdata['completion_time']['std_time'],
            'avg_rooms_per_maid': hdata['productivity']['avg_rooms_per_maid'],
            'score': score_data.get('avg_score', None),
            'claim_rate': claim_rates.get(hk, 0),
            'maid_data_points': hdata['staffing']['maid_data_points'],
            'time_data_points': hdata['completion_time']['data_points'],
        })

    hotel_summary.sort(key=lambda x: x['score'] or 0, reverse=True)

    # Build output
    output_3 = {
        'analysis_metadata': {
            'analysis_id': 3,
            'title': '人員配置×品質 相関分析',
            'subtitle': 'Staffing Level × Quality Correlation Analysis',
            'total_hotels': 19,
            'hotels_with_staffing_data': n3,
            'data_period': 'R8年度3月（日報データ）'
        },
        'staffing_analysis': staffing_analysis,
        'hotel_summary': hotel_summary,
        'recommendations': staffing_recs,
    }

    output_4 = {
        'analysis_metadata': {
            'analysis_id': 4,
            'title': '清掃完了時間×品質 分析',
            'subtitle': 'Cleaning Completion Time × Quality Analysis',
            'total_hotels': 19,
            'hotels_with_time_data': n4,
            'data_period': 'R8年度3月（日報データ）'
        },
        'time_analysis': time_analysis,
        'hotel_summary': hotel_summary,
        'recommendations': time_recs,
    }

    # Remove daily entries to reduce JSON size
    for dp in output_3['staffing_analysis'].get('data_points', []):
        pass  # keep pairs only
    for dp in output_4['time_analysis'].get('data_points', []):
        pass  # keep pairs only

    # Save
    path3 = os.path.join(OUT_DIR, 'analysis_3_data.json')
    with open(path3, 'w', encoding='utf-8') as f:
        json.dump(output_3, f, ensure_ascii=False, indent=2)
    print(f"\n✅ 分析3 → {path3} ({os.path.getsize(path3):,} bytes)")

    path4 = os.path.join(OUT_DIR, 'analysis_4_data.json')
    with open(path4, 'w', encoding='utf-8') as f:
        json.dump(output_4, f, ensure_ascii=False, indent=2)
    print(f"✅ 分析4 → {path4} ({os.path.getsize(path4):,} bytes)")

    # Summary
    print(f"\n===== 分析3 サマリー =====")
    for key, val in staffing_analysis.get('correlations', {}).items():
        print(f"  {key}: r={val.get('r')}")
    tiers = staffing_analysis.get('staffing_tiers', {})
    print(f"  低配置グループ: 平均{tiers.get('low_staff',{}).get('avg_maids','N/A')}名 → スコア{tiers.get('low_staff',{}).get('avg_score','N/A')}")
    print(f"  高配置グループ: 平均{tiers.get('high_staff',{}).get('avg_maids','N/A')}名 → スコア{tiers.get('high_staff',{}).get('avg_score','N/A')}")

    print(f"\n===== 分析4 サマリー =====")
    print(f"  平均完了時間: {time_analysis.get('overall_avg_time','N/A')}時")
    print(f"  範囲: {time_analysis.get('overall_time_range','N/A')}")
    for key, val in time_analysis.get('correlations', {}).items():
        print(f"  {key}: r={val.get('r')}")


if __name__ == '__main__':
    main()
