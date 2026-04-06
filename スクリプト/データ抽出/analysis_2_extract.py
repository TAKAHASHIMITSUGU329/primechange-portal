#!/usr/bin/env python3
"""分析2: スタッフ個人別パフォーマンス分析"""
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

def extract_staff_data():
    all_hotels = {}
    for hk in HOTEL_FILES:
        wb = open_workbook(hk)
        quality_sheet = award_sheet = None
        for s in wb.sheetnames:
            if '品質' in s and 'データ' in s: quality_sheet = s
            if '皆勤' in s: award_sheet = s

        hotel_data = {'name': HOTEL_NAME_MAP.get(hk, hk), 'maids': [], 'checkers': [], 'roster': []}

        # Extract maid/checker claim data from quality sheet
        if quality_sheet:
            ws = wb[quality_sheet]
            for row_idx in range(24, 50):
                maid_name = ws.cell(row=row_idx, column=2).value
                maid_claims = safe_number(ws.cell(row=row_idx, column=3).value, None)
                if maid_name and maid_name != '-' and maid_claims is not None and maid_claims > 0:
                    hotel_data['maids'].append({'name': str(maid_name).strip(), 'claims': int(maid_claims)})

                checker_name = ws.cell(row=row_idx, column=4).value
                checker_claims = safe_number(ws.cell(row=row_idx, column=5).value, None)
                if checker_name and checker_name != '-' and checker_claims is not None and checker_claims > 0:
                    hotel_data['checkers'].append({'name': str(checker_name).strip(), 'claims': int(checker_claims)})

        # Extract roster from award sheet
        if award_sheet:
            ws2 = wb[award_sheet]
            for row_idx in range(4, 100):
                name = ws2.cell(row=row_idx, column=3).value
                if not name: continue
                position = ws2.cell(row=row_idx, column=4).value
                pay_type = ws2.cell(row=row_idx, column=5).value
                total_days = safe_number(ws2.cell(row=row_idx, column=6).value, None)
                absence = safe_number(ws2.cell(row=row_idx, column=8).value, None)
                hours = safe_number(ws2.cell(row=row_idx, column=9).value, None)

                # Monthly data: Feb=c21(days),c23(hrs),c24(rooms) / Mar=c31(days),c33(hrs),c34(rooms) / Apr=c41(days),c43(hrs),c44(rooms)
                feb_days = safe_number(ws2.cell(row=row_idx, column=21).value, 0)
                feb_hours = safe_number(ws2.cell(row=row_idx, column=23).value, 0)
                feb_rooms = safe_number(ws2.cell(row=row_idx, column=24).value, 0)
                mar_days = safe_number(ws2.cell(row=row_idx, column=31).value, 0)
                mar_hours = safe_number(ws2.cell(row=row_idx, column=33).value, 0)
                mar_rooms = safe_number(ws2.cell(row=row_idx, column=34).value, 0)
                apr_days = safe_number(ws2.cell(row=row_idx, column=41).value, 0)
                apr_hours = safe_number(ws2.cell(row=row_idx, column=43).value, 0)
                apr_rooms = safe_number(ws2.cell(row=row_idx, column=44).value, 0)

                total_rooms = int(feb_rooms + mar_rooms + apr_rooms)
                sum_days = int(feb_days + mar_days + apr_days) if (feb_days + mar_days + apr_days) > 0 else (int(total_days) if total_days else 0)

                hotel_data['roster'].append({
                    'name': str(name).strip(),
                    'position': str(position).strip() if position else 'unknown',
                    'pay_type': str(pay_type).strip() if pay_type else '',
                    'total_days': sum_days,
                    'absence_days': int(absence) if absence else 0,
                    'total_hours': round((feb_hours + mar_hours + apr_hours) or (hours or 0), 1),
                    'rooms_cleaned': total_rooms,
                    'feb_rooms': int(feb_rooms),
                    'mar_rooms': int(mar_rooms),
                    'apr_rooms': int(apr_rooms),
                })

        hotel_data['total_maid_claims'] = sum(m['claims'] for m in hotel_data['maids'])
        hotel_data['total_checker_claims'] = sum(c['claims'] for c in hotel_data['checkers'])
        hotel_data['roster_size'] = len(hotel_data['roster'])
        hotel_data['maid_count'] = sum(1 for r in hotel_data['roster'] if 'メイド' in r['position'])
        hotel_data['checker_count'] = sum(1 for r in hotel_data['roster'] if 'チェッカー' in r['position'])

        all_hotels[hk] = hotel_data
        wb.close()

    return all_hotels

def analyze_performance(all_hotels):
    # Cross-hotel maid claim analysis
    all_maids = []
    for hk, hdata in all_hotels.items():
        for m in hdata['maids']:
            all_maids.append({**m, 'hotel': hdata['name'], 'hotel_key': hk})

    all_checkers = []
    for hk, hdata in all_hotels.items():
        for c in hdata['checkers']:
            all_checkers.append({**c, 'hotel': hdata['name'], 'hotel_key': hk})

    # Roster analysis
    all_staff = []
    for hk, hdata in all_hotels.items():
        for r in hdata['roster']:
            r_copy = {**r, 'hotel': hdata['name'], 'hotel_key': hk}
            if r['total_days'] > 0 and r['rooms_cleaned'] > 0:
                r_copy['rooms_per_day'] = round(r['rooms_cleaned'] / r['total_days'], 1)
            else:
                r_copy['rooms_per_day'] = 0
            all_staff.append(r_copy)

    # Maid productivity stats
    maids_roster = [s for s in all_staff if 'メイド' in s['position'] and s['rooms_cleaned'] > 0]
    if maids_roster:
        rpd_values = [m['rooms_per_day'] for m in maids_roster if m['rooms_per_day'] > 0]
        maid_productivity = {
            'total_maids_with_room_data': len(maids_roster),
            'avg_rooms_per_day': round(np.mean(rpd_values), 1) if rpd_values else 0,
            'median_rooms_per_day': round(np.median(rpd_values), 1) if rpd_values else 0,
            'max_rooms_per_day': round(max(rpd_values), 1) if rpd_values else 0,
            'min_rooms_per_day': round(min(rpd_values), 1) if rpd_values else 0,
            'top_performers': sorted(maids_roster, key=lambda x: x['rooms_per_day'], reverse=True)[:10],
        }
    else:
        maid_productivity = {'total_maids_with_room_data': 0}

    # Attendance analysis
    all_attendance = [s for s in all_staff if s['total_days'] > 0]
    attendance_analysis = {
        'total_staff': len(all_attendance),
        'avg_days': round(np.mean([s['total_days'] for s in all_attendance]), 1) if all_attendance else 0,
        'high_attendance': len([s for s in all_attendance if s['total_days'] >= 20]),
        'low_attendance': len([s for s in all_attendance if s['total_days'] < 10]),
    }

    # Hotel-level summary
    hotel_summaries = []
    for hk, hdata in all_hotels.items():
        maids_with_rooms = [r for r in hdata['roster'] if 'メイド' in r['position'] and r['rooms_cleaned'] > 0]
        avg_rpd = round(np.mean([r['rooms_cleaned'] / r['total_days'] for r in maids_with_rooms if r['total_days'] > 0]), 1) if maids_with_rooms else 0

        hotel_summaries.append({
            'hotel_key': hk,
            'name': hdata['name'],
            'roster_size': hdata['roster_size'],
            'maid_count': hdata['maid_count'],
            'checker_count': hdata['checker_count'],
            'total_claims': hdata['total_maid_claims'],
            'claims_per_maid': round(hdata['total_maid_claims'] / hdata['maid_count'], 2) if hdata['maid_count'] > 0 else 0,
            'avg_rooms_per_day': avg_rpd,
            'avg_attendance_days': round(np.mean([r['total_days'] for r in hdata['roster'] if r['total_days'] > 0]), 1) if any(r['total_days'] > 0 for r in hdata['roster']) else 0,
        })

    hotel_summaries.sort(key=lambda x: x['claims_per_maid'])

    # Recommendations
    recs = [
        {'title': 'クレーム多発スタッフへの個別研修', 'priority': '最優先',
         'rationale': f'全{len(all_maids)}名のメイドからクレームが報告。上位集中是正で全体品質向上が可能。',
         'actions': ['クレーム2件以上のメイドへの個別フィードバック面談', '清掃手順の再確認と実技研修', '改善後のフォローアップ（1ヶ月後再評価）']},
        {'title': 'チェッカー品質管理力の強化', 'priority': '高',
         'rationale': f'チェッカー{len(all_checkers)}名の指摘実績を分析。見逃し傾向のあるチェッカーの底上げが重要。',
         'actions': ['チェッカー間のクレーム発見率比較と是正', 'ベストチェッカーの手法共有会の開催', 'チェック項目の統一と漏れ防止']},
        {'title': '生産性指標の導入と目標管理', 'priority': '中',
         'rationale': f'メイド1日あたり清掃室数の平均{maid_productivity.get("avg_rooms_per_day", 0)}室。個人差が大きく標準化が課題。',
         'actions': ['日あたり清掃室数の個人別トラッキング', '効率と品質のバランス指標の設定', 'トップパフォーマーのベストプラクティス抽出']},
    ]

    return {
        'maid_claims_summary': {'total_maids_with_claims': len(all_maids), 'top_claim_maids': sorted(all_maids, key=lambda x: x['claims'], reverse=True)[:15]},
        'checker_claims_summary': {'total_checkers_with_claims': len(all_checkers), 'top_claim_checkers': sorted(all_checkers, key=lambda x: x['claims'], reverse=True)[:15]},
        'maid_productivity': maid_productivity,
        'attendance_analysis': attendance_analysis,
        'hotel_summaries': hotel_summaries,
        'recommendations': recs,
    }

def main():
    print("分析2: スタッフ個人別パフォーマンス分析")
    print("=" * 50)
    print("\n[1/2] スタッフデータ抽出中...")
    all_hotels = extract_staff_data()
    total_roster = sum(h['roster_size'] for h in all_hotels.values())
    total_maids = sum(len(h['maids']) for h in all_hotels.values())
    print(f"  → {len(all_hotels)}ホテル、在籍{total_roster}名、クレーム関連メイド{total_maids}名")

    print("\n[2/2] パフォーマンス分析中...")
    results = analyze_performance(all_hotels)
    print(f"  → 生産性データ: {results['maid_productivity'].get('total_maids_with_room_data', 0)}名")
    print(f"  → 平均清掃室数/日: {results['maid_productivity'].get('avg_rooms_per_day', 0)}室")

    output = {
        'analysis_metadata': {
            'analysis_id': 2, 'title': 'スタッフ個人別パフォーマンス分析',
            'subtitle': 'Individual Staff Performance Analysis',
            'total_hotels': 19, 'total_staff_analyzed': total_roster,
            'data_period': 'R8年度（2月〜4月）'
        },
        **results,
    }

    path = os.path.join(OUT_DIR, 'analysis_2_data.json')
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    print(f"\n✅ → {path} ({os.path.getsize(path):,} bytes)")

if __name__ == '__main__':
    main()
