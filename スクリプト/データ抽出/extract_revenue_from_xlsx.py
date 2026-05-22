#!/usr/bin/env python3
"""XLSXの集計表シートから売上・稼働率・利益データを自動抽出し、hotel_revenue_data.jsonを更新する。"""

import json
import os
import sys
import csv
import io
import urllib.parse
import urllib.request

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from hotel_xlsx_utils import HOTEL_FILES, open_workbook, safe_number

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
JSON_DIR = os.path.join(BASE_DIR, '..', '..', 'データ', '分析結果JSON')
OUTPUT_PATH = os.path.join(JSON_DIR, 'hotel_revenue_data.json')
SCRIPT_OUTPUT_PATH = os.path.join(BASE_DIR, 'hotel_revenue_data.json')

SPREADSHEET_IDS = {
    'daiwa_osaki': '1IIHEn4nAIy9UXzrYptU-RQIiTbKkV0G_CaABF7znVrY',
    'chisan': '1IWigsWTzbRG-juWtIlg4ZchiuWqRhJFpPPczXdQxG6Y',
    'hearton': '1A25mmVRYSnG3ZB8oa0oZVp-vCP2xMkwX-Zqdkk4BIzI',
    'keyakigate': '1srchDxFyv7TJ3IEZXJ19miH04p3jRug5nVtA3BLertQ',
    'richmond_mejiro': '1XWU6925CpT3GMMonAqy4UENKM11gWloUkJIsgGImUts',
    'keisei_richmond': '1jUS_HwTfowG1xIHFtwJbCL5dTj7FrhvUe6d32AevZ2g',
    'daiichi_ikebukuro': '1X2GgFKxTOs7CuJSlPYrpzigraSnWcKh6cMJLfsXhWlU',
    'comfort_roppongi': '1Jtm0rXTigY2OVManNjx1qQ6G9EKQEXuPs_T1BdlOvls',
    'comfort_suites_tokyobay': '1zCFAmzRqvSDbjwvK7qI4cYBHlrmBifTPm0Y-g0rruyE',
    'comfort_era_higashikanda': '1H9jmOVQR4UdEQ5hsxZ2Xz44BT72RJDwNa6BOKFhXxRg',
    'comfort_narita': '1lQ3FRDuE75dkByQRFd0i0F2xcHnl-3-UAOJwhIt3jAU',
    'comfort_yokohama': '1rnQOsyUXuSzBKdqPN_ey_4Iw5VtYWTgSR5Z4nh-1zd4',
    'apa_sagamihara': '1E2ZQJyE6pOJ3jr6GyB56KcYnVVq54m6dO_6h_SQy39A',
    'apa_kamata': '16xuhAdNzdeyAKu-LhU8ATgR8_kZ1JXfa9lT51tAB1Nw',
    'court_shinyokohama': '1Qm5lPPc8m7yutyIH3Pf03YUnF2KpnWjn0SecMzq0CjY',
    'comment_yokohama': '1cVH7khdgh8bDN-wtAw2KVakJqHILo58VOBu0SKmBFrU',
    'henn_na_haneda': '18DkZLJ8UDQ2-4MBrh7B4y28tHaYnoWIQqEoFkvFDNKg',
    'kawasaki_nikko': '1aQ2MaKJmOz7eT53oqszCDO9Fa3UEbfhFSgXfVmVpO9A',
    'comfort_hakata': '1_7xoyIiq1llfO0I2328ZQlB6sD0lMsnpRMp1rMGNPcg',
}

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
    (5, '①R8_5集計', 31),
]

MONTH_FIELD_PREFIX = {
    3: 'march',
    4: 'april',
    5: 'may',
}

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


def parse_number(val, default=None):
    """Convert display/raw spreadsheet value to number."""
    if val is None:
        return default
    if isinstance(val, (int, float)):
        return val
    text = str(val).strip().replace(',', '').replace('¥', '').replace('￥', '')
    if not text:
        return default
    is_pct = text.endswith('%')
    if is_pct:
        text = text[:-1]
    try:
        num = float(text)
        return num / 100 if is_pct else num
    except ValueError:
        return default


def fetch_live_sheet_rows(hotel_key, sheet_name):
    """Fetch a sheet through the Apps Script CSV proxy when configured."""
    base_url = os.environ.get('GAS_CSV_PROXY_URL', '').strip().rstrip('/')
    spreadsheet_id = SPREADSHEET_IDS.get(hotel_key)
    if not base_url or not spreadsheet_id:
        return None

    url = base_url + '?' + urllib.parse.urlencode({
        'id': spreadsheet_id,
        'sheet': sheet_name,
        'values': 'display',
    })
    with urllib.request.urlopen(url, timeout=90) as response:
        text = response.read().decode('utf-8-sig', errors='replace')
    if text.startswith('ERROR:'):
        return None
    return list(csv.reader(io.StringIO(text, newline='')))


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


def cell(rows, row, col):
    """Read 1-based row/column from CSV rows."""
    try:
        return rows[row - 1][col - 1]
    except IndexError:
        return None


def extract_month_data_from_rows(rows, month_num, days):
    revenue = parse_number(cell(rows, ROW_ACTUAL, COL_REVENUE), None)
    occupancy = parse_number(cell(rows, ROW_ACTUAL, COL_OCCUPANCY), None)
    net_profit = parse_number(cell(rows, ROW_ACTUAL, COL_NET_PROFIT), None)
    profit_rate = parse_number(cell(rows, ROW_ACTUAL, COL_PROFIT_RATE), None)
    variable_cost = parse_number(cell(rows, ROW_ACTUAL, COL_VARIABLE_COST), None)
    target_revenue = parse_number(cell(rows, ROW_TARGET, COL_REVENUE), None)
    target_profit = parse_number(cell(rows, ROW_TARGET, COL_NET_PROFIT), None)
    rooms_sold = parse_number(cell(rows, ROW_TOTAL, COL_ROOMS_SOLD), None)
    room_count = parse_number(cell(rows, 13, COL_ROOM_COUNT), None)

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
    if rooms_sold and rooms_sold > 0:
        result['adr'] = round(revenue / rooms_sold, 4)
    if room_count and room_count > 0:
        result['revpar'] = round(revenue / (room_count * days), 4)
    return result


def apply_month_data(entry, month_num, data):
    if month_num == 2:
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
        return

    prefix = MONTH_FIELD_PREFIX.get(month_num)
    if not prefix:
        return
    entry[f'{prefix}_revenue'] = data['revenue']
    entry[f'{prefix}_occupancy'] = round(data['occupancy'], 4) if data['occupancy'] else 0
    entry[f'{prefix}_net_profit'] = data['net_profit']
    entry[f'{prefix}_days'] = data['days']
    entry[f'{prefix}_rooms_sold'] = data['rooms_sold']
    entry[f'{prefix}_target_revenue'] = data['target_revenue']


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

        wb = None
        try:
            wb = open_workbook(key)
        except FileNotFoundError as e:
            print(f"    XLSXなし: {e}")

        # 既存エントリまたは新規作成
        entry = existing.get(key, {})
        entry['hotel_name'] = hotel_name
        entry['key'] = key
        entry['phase'] = phase

        for month_num, sheet_name, days in MONTH_SHEETS:
            data = None
            live_rows = None
            if month_num >= 5:
                try:
                    live_rows = fetch_live_sheet_rows(key, sheet_name)
                except Exception as e:
                    print(f"    {month_num}月: ライブ取得失敗 ({e})")
            if live_rows:
                data = extract_month_data_from_rows(live_rows, month_num, days)
            elif wb and sheet_name in wb.sheetnames:
                data = extract_month_data(wb[sheet_name], month_num, days)
            elif wb:
                print(f"    {month_num}月: シート'{sheet_name}'なし")
                continue

            if not data:
                print(f"    {month_num}月: データなし")
                continue

            apply_month_data(entry, month_num, data)
            print(f"    {month_num}月: 売上¥{data['revenue']:,.0f} 稼働率{data['occupancy']*100:.1f}%")

        if wb:
            wb.close()
        existing[key] = entry
        updated += 1

        if wb is None:
            skipped += 1

    # JSON出力
    with open(OUTPUT_PATH, 'w', encoding='utf-8') as f:
        json.dump(existing, f, ensure_ascii=False, indent=2)
    with open(SCRIPT_OUTPUT_PATH, 'w', encoding='utf-8') as f:
        json.dump(existing, f, ensure_ascii=False, indent=2)

    print(f"\n{'=' * 50}")
    print(f"完了: {updated}ホテル更新 / {skipped}スキップ")
    print(f"出力: {OUTPUT_PATH}")


if __name__ == '__main__':
    main()
