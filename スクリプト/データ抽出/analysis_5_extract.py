#!/usr/bin/env python3
"""分析5: 安全チェック×予兆検出分析"""
import sys, os, json
import numpy as np
sys.path.insert(0, os.path.dirname(__file__))
from hotel_xlsx_utils import HOTEL_FILES, open_workbook, safe_number, revenue_key

OUT_DIR = os.path.dirname(__file__)
HOTEL_NAME_MAP = {
    'daiwa_osaki': 'ダイワロイネットホテル東京大崎', 'chisan': 'チサンホテル浜松町',
    'hearton': 'ハートンホテル東品川', 'keyakigate': 'ホテルケヤキゲート東京府中',
    'richmond_mejiro': 'リッチモンドホテル東京目白', 'keisei_richmond': '京成リッチモンドホテル東京錦糸町',
    'daiichi_ikebukuro': '第一イン池袋', 'comfort_roppongi': 'コンフォートイン六本木',
    'comfort_suites_tokyobay': 'コンフォートスイーツ東京ベイ', 'comfort_era_higashikanda': 'コンフォートホテルERA東神田',
    'comfort_narita': 'コンフォートホテル成田', 'comfort_yokohama': 'コンフォートホテル横浜関内',
    'apa_sagamihara': 'アパホテル相模原橋本駅東', 'apa_kamata': 'アパホテル蒲田駅東',
    'court_shinyokohama': 'コートホテル新横浜', 'comment_yokohama': 'ホテルコメント横浜関内',
    'henn_na_haneda': '変なホテル東京羽田', 'kawasaki_nikko': '川崎日航ホテル',
    'comfort_hakata': 'コンフォートホテル博多',
}

# Safety check items mapping
SAFETY_ITEMS = {
    'seiri_seiton': {'label': '整理整頓', 'rows': [9, 10, 11, 12, 13, 14], 'scoring': 'osx'},
    'anzen_kanri': {'label': '安全管理', 'rows': [16, 17, 18, 19], 'scoring': 'osx'},
    'eisei_kanri': {'label': '衛生管理', 'rows': [21, 22, 23, 24], 'scoring': 'osx'},
    'genba_unei': {'label': '現場運営', 'rows': [27, 28, 29, 30, 31, 32, 33], 'scoring': 'abcde'},
    'genba_rule': {'label': '現場ルール・マナー', 'rows': [35, 36, 37], 'scoring': 'abcde'},
}

def score_osx(val):
    if not val: return None
    s = str(val).strip()
    if s == '〇': return 3
    if s == '△': return 2
    if s == '✖': return 1
    return None

def score_abcde(val):
    if not val: return None
    s = str(val).strip().upper()
    mapping = {'A': 5, 'B': 4, 'C': 3, 'D': 2, 'E': 1}
    return mapping.get(s)

def extract_safety_data():
    all_hotels = {}
    all_inspection_months = []
    for hk in HOTEL_FILES:
        wb = open_workbook(hk)
        safety_sheet = None
        for s in wb.sheetnames:
            if '安全チェック' in s or ('安全' in s and 'チェック' in s):
                safety_sheet = s
                break
        if not safety_sheet:
            wb.close()
            continue

        ws = wb[safety_sheet]
        hotel_data = {'name': HOTEL_NAME_MAP.get(hk, hk), 'inspections': [], 'category_scores': {}}

        # Find inspection date columns (dynamic - check row 5 for dates)
        date_cols = []
        for col in range(4, 30, 3):  # dates at col 4, 7, 10, ...
            month_val = ws.cell(row=5, column=col).value
            day_val = ws.cell(row=5, column=col + 1).value
            inspector = ws.cell(row=6, column=col).value
            if month_val or inspector:
                date_cols.append({'col': col, 'month': str(month_val or ''), 'day': str(day_val or ''), 'inspector': str(inspector or '')})

        # Extract scores for each inspection date
        for dc in date_cols:
            col = dc['col']
            inspection = {'date_info': dc, 'categories': {}, 'items': []}
            all_scores = []

            for cat_key, cat_info in SAFETY_ITEMS.items():
                score_fn = score_osx if cat_info['scoring'] == 'osx' else score_abcde
                cat_scores = []

                for row in cat_info['rows']:
                    # Check both the date column and nearby columns
                    val = ws.cell(row=row, column=col).value
                    if val is None:
                        val = ws.cell(row=row, column=col - 1).value  # try col-1
                    score = score_fn(val)
                    item_text = ws.cell(row=row, column=2).value
                    if score is not None:
                        cat_scores.append(score)
                        all_scores.append(score)
                        inspection['items'].append({
                            'category': cat_info['label'],
                            'item': str(item_text)[:50] if item_text else '',
                            'value': str(val),
                            'score': score,
                            'max_score': 3 if cat_info['scoring'] == 'osx' else 5,
                        })

                if cat_scores:
                    max_s = 3 if cat_info['scoring'] == 'osx' else 5
                    inspection['categories'][cat_info['label']] = {
                        'avg_score': round(np.mean(cat_scores), 2),
                        'max_possible': max_s,
                        'pct': round(np.mean(cat_scores) / max_s * 100, 1),
                        'items_scored': len(cat_scores),
                        'problem_items': len([s for s in cat_scores if s <= (1 if max_s == 3 else 2)]),
                    }

            if all_scores:
                inspection['overall_score'] = round(np.mean(all_scores), 2)
                inspection['total_items'] = len(all_scores)
                inspection['problem_count'] = len([s for s in all_scores if s <= 1])
                hotel_data['inspections'].append(inspection)
                # 実データがある月のみ期間計算に含める
                if dc['month']:
                    import re
                    m_match = re.search(r'(\d+)', dc['month'])
                    if m_match:
                        all_inspection_months.append(int(m_match.group(1)))

        # Hotel-level aggregates
        if hotel_data['inspections']:
            avg_scores = [i['overall_score'] for i in hotel_data['inspections']]
            hotel_data['avg_overall_score'] = round(np.mean(avg_scores), 2)
            hotel_data['inspections_count'] = len(hotel_data['inspections'])

            # Aggregate category scores
            for cat_label in [c['label'] for c in SAFETY_ITEMS.values()]:
                cat_vals = []
                for insp in hotel_data['inspections']:
                    if cat_label in insp['categories']:
                        cat_vals.append(insp['categories'][cat_label]['pct'])
                if cat_vals:
                    hotel_data['category_scores'][cat_label] = round(np.mean(cat_vals), 1)
        else:
            hotel_data['avg_overall_score'] = None
            hotel_data['inspections_count'] = 0

        all_hotels[hk] = hotel_data
        wb.close()

    if all_inspection_months:
        min_m = min(all_inspection_months)
        max_m = max(all_inspection_months)
        safety_period = f'R8年度（{min_m}月〜{max_m}月）' if min_m != max_m else f'R8年度（{min_m}月）'
    else:
        safety_period = 'R8年度'

    return all_hotels, safety_period

def analyze_safety_vs_quality(all_hotels):
    with open(os.path.join(OUT_DIR, 'primechange_portfolio_analysis.json'), 'r') as f:
        portfolio = json.load(f)
    scores = {}
    for h in portfolio.get('portfolio_overview', {}).get('hotels_ranked', []):
        scores[h['key']] = h['avg']
        mapped = revenue_key(h['key'])
        if mapped != h['key']: scores[mapped] = h['avg']

    with open(os.path.join(OUT_DIR, 'analysis_1_data.json'), 'r') as f:
        claims = json.load(f)
    claim_rates = {}
    for hp in claims.get('hotel_profiles', []):
        claim_rates[hp['hotel_key']] = hp['claim_rate']

    # Build paired data
    pairs = []
    for hk, hdata in all_hotels.items():
        if hdata['avg_overall_score'] is None: continue
        mapped = revenue_key(hk)
        score = scores.get(hk) or scores.get(mapped)
        cr = claim_rates.get(hk, 0)
        if not score: continue
        pairs.append({
            'hotel_key': hk, 'name': hdata['name'],
            'safety_score': hdata['avg_overall_score'],
            'review_score': score, 'claim_rate': cr,
            'inspections': hdata['inspections_count'],
            'category_scores': hdata['category_scores'],
        })

    results = {'pairs': pairs, 'correlations': {}}
    if len(pairs) >= 3:
        safety = np.array([p['safety_score'] for p in pairs])
        review = np.array([p['review_score'] for p in pairs])
        if np.std(safety) > 0:
            r = float(np.corrcoef(safety, review)[0, 1])
            results['correlations']['safety_vs_review'] = {'r': round(r, 4), 'n': len(pairs)}

        cr_arr = np.array([p['claim_rate'] for p in pairs])
        if np.std(cr_arr) > 0 and np.std(safety) > 0:
            r = float(np.corrcoef(safety, cr_arr)[0, 1])
            results['correlations']['safety_vs_claims'] = {'r': round(r, 4), 'n': len(pairs)}

    # Problem areas (items with low scores)
    problem_items = []
    for hk, hdata in all_hotels.items():
        for insp in hdata.get('inspections', []):
            for item in insp.get('items', []):
                if item['score'] <= (1 if item['max_score'] == 3 else 2):
                    problem_items.append({
                        'hotel': hdata['name'], 'category': item['category'],
                        'item': item['item'], 'value': item['value'], 'score': item['score'],
                    })

    results['problem_items'] = problem_items
    results['total_problem_count'] = len(problem_items)

    # Hotel ranking by safety score
    results['hotel_ranking'] = sorted(pairs, key=lambda x: x['safety_score'], reverse=True)

    # Recommendations
    results['recommendations'] = [
        {'title': '安全チェック実施率の向上', 'priority': '最優先',
         'rationale': f'19ホテル中{len([h for h in all_hotels.values() if h["inspections_count"] > 0])}ホテルでチェック実施。未実施ホテルへの即時導入が必要。',
         'actions': ['月1回の安全チェック実施を全ホテルに義務化', 'チェックリストのデジタル化による入力率向上', 'チェック結果の本社共有と改善追跡']},
        {'title': '問題項目の重点改善', 'priority': '高',
         'rationale': f'全{len(problem_items)}件の問題項目を特定。整理整頓と現場運営に課題集中。',
         'actions': ['△・✖評価の項目への是正措置と期限設定', '改善前後の写真記録による可視化', '四半期での改善進捗レビュー']},
        {'title': '安全スコアと品質KPIの連動', 'priority': '中',
         'rationale': '安全チェックスコアとクレーム率・口コミスコアの関連を継続的にモニタリングし、予兆検出に活用。',
         'actions': ['安全スコアの低下→クレーム増加の早期警戒指標化', '月次ダッシュボードへの安全チェックスコア統合', 'ベストプラクティスホテルの安全管理手法の横展開']},
    ]

    return results

def main():
    print("分析5: 安全チェック×予兆検出分析")
    print("=" * 50)
    print("\n[1/2] 安全チェックデータ抽出中...")
    all_hotels, safety_period = extract_safety_data()
    checked = sum(1 for h in all_hotels.values() if h['inspections_count'] > 0)
    print(f"  → {len(all_hotels)}ホテル、チェック実施{checked}ホテル")

    print("\n[2/2] 品質連動分析中...")
    results = analyze_safety_vs_quality(all_hotels)
    for k, v in results.get('correlations', {}).items():
        print(f"  {k}: r={v['r']}")
    print(f"  問題項目: {results['total_problem_count']}件")

    output = {
        'analysis_metadata': {
            'analysis_id': 5, 'title': '安全チェック×予兆検出分析',
            'subtitle': 'Safety Inspection × Early Warning Analysis',
            'total_hotels': 19, 'hotels_with_data': checked,
            'data_period': safety_period
        },
        'hotel_details': {hk: {
            'name': h['name'], 'avg_score': h['avg_overall_score'],
            'inspections_count': h['inspections_count'], 'category_scores': h['category_scores'],
        } for hk, h in all_hotels.items()},
        **results,
    }

    path = os.path.join(OUT_DIR, 'analysis_5_data.json')
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    print(f"\n✅ → {path} ({os.path.getsize(path):,} bytes)")

if __name__ == '__main__':
    main()
