#!/usr/bin/env python3
"""XLSXの集計表シートから売上・稼働率・利益データを自動抽出し、hotel_revenue_data.jsonを更新する。"""

import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from hotel_xlsx_utils import HOTEL_FILES, open_workbook, safe_number

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
JSON_DIR = os.path.join(BASE_DIR, '..', '..', 'データ', '分析結果JSON')
OUTPUT_PATH = os.path.join(JSON_DIR, 'hotel_revenue_data.json')

# ホテル名・フェーズマッピング
HOTEL_META = {
    'daiwa_osaki': ('ダイワロイネットホテル東京大崎', 'P1'),
    'chisan': ('チサンホテル浜松町', 'P1'),
    'hearton': ('ハートンホテル東品川', 'P1'),
    'keyakigate': ('ホテルケヤキゲート東京府中', 'P1'),
    'richmond_mejiro': ('リッチモンドホテル東京目白', 'P1'),
    'keisei_richmond': ('京成リッチモンドホテル東京錦糸町', 'P1'),
    'daiichi_ikebukuro': ('第一イン池袋', 'P1'),
    'comfort_roppongi': ('コンフォートイン六本木', 'P2'),
    'comfort_suites_tokyobay': ('コンフォートスイーツ東京ベイ', 'P2'),
    'comfort_era_higashikanda': ('コンフォートホテルERA東京東神田', 'P2'),
    'comfort_narita': ('コンフォートホテル成田', 'P2'),
    'comfort_yokohama': ('コンフォートホテル横浜関内', 'P2'),
    'apa_sagamihara': ('アパホテル相模原橋本駅東', 'P3'),
    'apa_kamata': ('アパホテル蒲田駅東', 'P3'),
    'court_shinyokohama': ('コートホテル新横浜', 'P3'),
    'comment_yokohama': ('ホテルコメント横浜関内', 'P3'),
    'henn_na_haneda': ('変なホテル東京羽田', 'P3'),
    'kawasaki_nikko': ('川崎日航ホテル', 'P3'),
    'comfort_hakata': ('コンフォートホテル博多', 'P4'),
}

# 月別シート名と日数
MONTH_SHEETS = [
    (2, '①R8_2集計', 28),
    (3, '①R8_3集計', 31),
    (4, '①R8_4集計', 30),
]

# セル位置（①R8_N集計シート共通）
ROW_TARGET = 7    # 目標行
ROW_ACTUAL = 8    # 今月（実績）行
COL_PROFIT_RATE = 3       # C: 利益率
COL_VARIABLE_COST = 5     # E: 変動費率
COL_OCCUPANCY = 7         # G: 稼働率
COL_REVENUE = 8           # H: 売上
COL_TARGET_PROFIT = 13    # M: 目標純利益（Row 7）
COL_NET_PROFIT = 13       # M: 純利益（Row 8）
ROW_TOTAL = 44
COL_ROOMS_SOLD = 7        # G: 販売客室数（Row 44）
COL_ROOM_COUNT = 5        # E: 部屋数（Row 13）


def extract_month_data(ws, month_num, days):
    """集計シートから1ヶ月分のデータを抽出する。"""
    revenue = safe_number(ws.cell(row=ROW_ACTUAL, column=COL_REVENUE).value, None)
    occupancy = safe_number(ws.cell(row=ROW_ACTUAL, column=COL_OCCUPANCY).value, None)
    net_profit = safe_number(ws.cell(row=ROW_ACTUAL, column=COL_NET_PROFIT).value, None)
    profit_rate = safe_number(ws.cell(row=ROW_ACTUAL, column=COL_PROFIT_RATE).value, None)
    variable_cost = safe_number(ws.cell(row=ROW_ACTUAL, column=COL_VARIABLE_COST).value, None)
    target_revenue = safe_number(ws.cell(row=ROW_TARGET, column=COL_REVENUE).value, None)
    target_profit = safe_number(ws.cell(row=ROW_TARGET, column=COL_NET_PROFIT).value, None)
    rooms_sold = safe_number(ws.cell(row=ROW_TOTAL, column=COL_ROOMS_SOLD).value, None)
    room_count = safe_number(ws.cell(row=13, column=COL_ROOM_COUNT).value, None)

    # 売上が0またはNoneならデータなしとみなす
    if not revenue or revenue <= 0:
        return None

    result = {
        'revenue': revenue,
        'occupancy': occupancy,
        'net_profit': net_profit,
        'profit_rate': profit_rate,
        'variable_cost_rate': variable_cost,
        'target_revenue': target_revenue,
        'target_net_profit': target_profit,
        'rooms_sold': rooms_sold,
        'room_count': int(room_count) if room_count else None,
        'days': days,
    }

    # ADR / RevPAR 計算
    if rooms_sold and rooms_sold > 0:
        result['adr'] = round(revenue / rooms_sold, 4)
    if room_count and room_count > 0:
        result['revpar'] = round(revenue / (room_count * days), 4)

    return result


def main():
    print("売上データ自動抽出: XLSX → hotel_revenue_data.json")
    print("=" * 50)

    # 既存JSON読み込み
    existing = {}
    if os.path.exists(OUTPUT_PATH):
        with open(OUTPUT_PATH, 'r') as f:
            existing = json.load(f)

    updated = 0
    skipped = 0

    for key in HOTEL_FILES:
        hotel_name, phase = HOTEL_META.get(key, (key, '?'))
        print(f"\n  {hotel_name} ({key}):")

        try:
            wb = open_workbook(key)
        except FileNotFoundError as e:
            print(f"    スキップ: {e}")
            skipped += 1
            continue

        # 既存エントリまたは新規作成
        entry = existing.get(key, {})
        entry['hotel_name'] = hotel_name
        entry['key'] = key
        entry['phase'] = phase

        for month_num, sheet_name, days in MONTH_SHEETS:
            if sheet_name not in wb.sheetnames:
                print(f"    {month_num}月: シート'{sheet_name}'なし")
                continue

            ws = wb[sheet_name]
            data = extract_month_data(ws, month_num, days)

            if not data:
                print(f"    {month_num}月: データなし")
                continue

            if month_num == 2:
                # 2月 = ベース月（直接フィールド）
                entry['period'] = '2026-02-01'
                entry['days_in_month'] = data['days']
                entry['actual_revenue'] = data['revenue']
                entry['occupancy_rate'] = round(data['occupancy'], 4) if data['occupancy'] else 0
                entry['actual_net_profit'] = data['net_profit']
                entry['profit_rate'] = round(data['profit_rate'], 4) if data['profit_rate'] else 0
                entry['variable_cost_rate'] = round(data['variable_cost_rate'], 4) if data['variable_cost_rate'] else 0
                entry['target_revenue'] = data['target_revenue']
                entry['target_net_profit'] = data['target_net_profit']
                entry['rooms_sold'] = data['rooms_sold']
                if data.get('room_count'):
                    entry['room_count'] = data['room_count']
                if data.get('adr'):
                    entry['adr'] = data['adr']
                if data.get('revpar'):
                    entry['revpar'] = data['revpar']
                print(f"    2月: 売上¥{data['revenue']:,.0f} 稼働率{data['occupancy']*100:.1f}%")

            elif month_num == 3:
                entry['march_revenue'] = data['revenue']
                entry['march_occupancy'] = round(data['occupancy'], 4) if data['occupancy'] else 0
                entry['march_net_profit'] = data['net_profit']
                entry['march_days'] = data['days']
                print(f"    3月: 売上¥{data['revenue']:,.0f} 稼働率{data['occupancy']*100:.1f}%")

            elif month_num == 4:
                entry['april_revenue'] = data['revenue']
                entry['april_occupancy'] = round(data['occupancy'], 4) if data['occupancy'] else 0
                entry['april_net_profit'] = data['net_profit']
                entry['april_days'] = data['days']
                print(f"    4月: 売上¥{data['revenue']:,.0f} 稼働率{data['occupancy']*100:.1f}%")

        wb.close()
        existing[key] = entry
        updated += 1

    # JSON出力
    with open(OUTPUT_PATH, 'w', encoding='utf-8') as f:
        json.dump(existing, f, ensure_ascii=False, indent=2)

    print(f"\n{'=' * 50}")
    print(f"完了: {updated}ホテル更新 / {skipped}スキップ")
    print(f"出力: {OUTPUT_PATH}")


if __name__ == '__main__':
    main()
