#!/usr/bin/env python3
"""Download live daily cleaning actuals from hotel operations spreadsheets."""

from __future__ import annotations

import csv
import datetime as dt
import io
import json
import os
import re
import sys
import time
import urllib.parse
import urllib.request

ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_SCRIPT_DIR = os.path.join(ROOT_DIR, "スクリプト", "データ抽出")
JSON_DIR = os.path.join(ROOT_DIR, "データ", "分析結果JSON")
OUT_PATH = os.path.join(JSON_DIR, "cleaning_actuals_daily.json")

sys.path.insert(0, DATA_SCRIPT_DIR)
from hotel_xlsx_utils import safe_number, safe_time  # noqa: E402

CURRENT_YEAR = 2026
TODAY = dt.date.today()

HOTELS = [
    ("daiwa_osaki", "ダイワロイネットホテル東京大崎", "1IIHEn4nAIy9UXzrYptU-RQIiTbKkV0G_CaABF7znVrY"),
    ("chisan", "チサンホテル浜松町", "1IWigsWTzbRG-juWtIlg4ZchiuWqRhJFpPPczXdQxG6Y"),
    ("hearton", "ハートンホテル東品川", "1A25mmVRYSnG3ZB8oa0oZVp-vCP2xMkwX-Zqdkk4BIzI"),
    ("keyakigate", "ホテルケヤキゲート東京府中", "1srchDxFyv7TJ3IEZXJ19miH04p3jRug5nVtA3BLertQ"),
    ("richmond_mejiro", "リッチモンドホテル東京目白", "1XWU6925CpT3GMMonAqy4UENKM11gWloUkJIsgGImUts"),
    ("keisei_kinshicho", "京成リッチモンドホテル東京錦糸町", "1jUS_HwTfowG1xIHFtwJbCL5dTj7FrhvUe6d32AevZ2g"),
    ("daiichi_ikebukuro", "第一イン池袋", "1X2GgFKxTOs7CuJSlPYrpzigraSnWcKh6cMJLfsXhWlU"),
    ("comfort_roppongi", "コンフォートイン六本木", "1Jtm0rXTigY2OVManNjx1qQ6G9EKQEXuPs_T1BdlOvls"),
    ("comfort_suites_tokyobay", "コンフォートスイーツ東京ベイ", "1zCFAmzRqvSDbjwvK7qI4cYBHlrmBifTPm0Y-g0rruyE"),
    ("comfort_era_higashikanda", "コンフォートホテルERA東京東神田", "1H9jmOVQR4UdEQ5hsxZ2Xz44BT72RJDwNa6BOKFhXxRg"),
    ("comfort_yokohama_kannai", "コンフォートホテル横浜関内", "1rnQOsyUXuSzBKdqPN_ey_4Iw5VtYWTgSR5Z4nh-1zd4"),
    ("comfort_narita", "コンフォートホテル成田", "1lQ3FRDuE75dkByQRFd0i0F2xcHnl-3-UAOJwhIt3jAU"),
    ("apa_kamata", "アパホテル蒲田駅東", "16xuhAdNzdeyAKu-LhU8ATgR8_kZ1JXfa9lT51tAB1Nw"),
    ("apa_sagamihara", "アパホテル相模原橋本駅東", "1E2ZQJyE6pOJ3jr6GyB56KcYnVVq54m6dO_6h_SQy39A"),
    ("court_shinyokohama", "コートホテル新横浜", "1Qm5lPPc8m7yutyIH3Pf03YUnF2KpnWjn0SecMzq0CjY"),
    ("comment_yokohama", "ホテルコメント横浜関内", "1cVH7khdgh8bDN-wtAw2KVakJqHILo58VOBu0SKmBFrU"),
    ("kawasaki_nikko", "川崎日航ホテル", "1aQ2MaKJmOz7eT53oqszCDO9Fa3UEbfhFSgXfVmVpO9A"),
    ("henn_na_haneda", "変なホテル東京羽田", "18DkZLJ8UDQ2-4MBrh7B4y28tHaYnoWIQqEoFkvFDNKg"),
    ("comfort_hakata", "コンフォートホテル博多", "1_7xoyIiq1llfO0I2328ZQlB6sD0lMsnpRMp1rMGNPcg"),
]


def fetch_text(base_url: str, params: dict[str, str], retries: int = 2) -> str:
    url = base_url.rstrip("/") + "?" + urllib.parse.urlencode(params)
    last_error = None
    for attempt in range(retries + 1):
        try:
            with urllib.request.urlopen(url, timeout=90) as response:
                text = response.read().decode("utf-8-sig", errors="replace")
            if text.startswith("ERROR:"):
                raise RuntimeError(text.strip())
            return text
        except Exception as exc:  # noqa: BLE001
            last_error = exc
            if attempt < retries:
                time.sleep(2 * (attempt + 1))
    raise RuntimeError(str(last_error))


def parse_csv(text: str) -> list[list[str]]:
    return list(csv.reader(io.StringIO(text, newline="")))


def daily_sheet_candidates(month: int) -> list[str]:
    return [f"③R8_{month}日報", f"④R8_{month}日報"]


def parse_day(value: str, month: int) -> str | None:
    value = str(value or "").strip()
    if not value:
        return None
    if re.match(r"^202\d-\d{1,2}-\d{1,2}", value):
        parsed = dt.date.fromisoformat(value[:10])
        return parsed.isoformat() if parsed.month == month else None
    if re.match(r"^\d{1,2}$", value):
        return dt.date(CURRENT_YEAR, month, int(value)).isoformat()
    return None


def parse_time_value(value: str) -> float | None:
    value = str(value or "").strip()
    if not value:
        return None
    match = re.search(r"(\d{1,2}):(\d{2})", value)
    if match:
        return int(match.group(1)) + int(match.group(2)) / 60.0
    parsed = safe_time(value)
    return parsed if parsed and parsed > 6 else None


def parse_workload(value: str) -> float | None:
    num = safe_number(value, None)
    if num:
        return num
    nums = re.findall(r"\d+", str(value or ""))
    return float(sum(int(n) for n in nums)) if nums else None


def row_value(row: list[str], idx: int) -> str:
    return row[idx] if idx < len(row) else ""


def extract_entries(rows: list[list[str]], month: int) -> list[dict]:
    entries = []
    for row in rows:
        date = parse_day(row_value(row, 2), month)
        if not date:
            continue
        parsed_date = dt.date.fromisoformat(date)
        if parsed_date > TODAY:
            continue

        maids = safe_number(row_value(row, 9), None)
        checkers = safe_number(row_value(row, 10), None)
        claims = safe_number(row_value(row, 7), 0) or 0
        workload = parse_workload(row_value(row, 13))
        completion_time = parse_time_value(row_value(row, 8))

        entry = {
            "date": date,
            "claims": int(claims),
            "completion_time": round(completion_time, 2) if completion_time else None,
            "maids": int(maids) if maids and maids > 0 and float(maids).is_integer() else maids,
            "checkers": int(checkers) if checkers and checkers > 0 and float(checkers).is_integer() else checkers,
            "workload": int(workload) if workload and float(workload).is_integer() else workload,
        }
        entry["total_staff"] = (
            (maids or 0) + (checkers or 0)
            if (maids and maids > 0) or (checkers and checkers > 0)
            else None
        )
        entry["rooms_per_maid"] = round(workload / maids, 1) if workload and maids and maids > 0 else None
        entries.append(entry)
    return entries


def extract_monthly_actuals(rows: list[list[str]], month: int) -> list[dict]:
    entries = []
    for row in rows:
        date = parse_day(row_value(row, 2), month)
        if not date:
            continue
        parsed_date = dt.date.fromisoformat(date)
        if parsed_date > TODAY:
            continue

        cleaned_rooms = safe_number(row_value(row, 6), None)
        occupied_rooms = safe_number(row_value(row, 5), None)
        room_count = safe_number(row_value(row, 4), None)
        dd_rooms = safe_number(row_value(row, 7), None)
        single_rooms = safe_number(row_value(row, 8), None)
        twin_rooms = safe_number(row_value(row, 9), None)
        double_rooms = safe_number(row_value(row, 12), None)
        extra_rooms = safe_number(row_value(row, 13), None)
        vcan_checks = safe_number(row_value(row, 14), None)

        if not any([cleaned_rooms, occupied_rooms, dd_rooms, vcan_checks]):
            continue

        entries.append(
            {
                "date": date,
                "room_count": int(room_count) if room_count and float(room_count).is_integer() else room_count,
                "occupied_rooms": int(occupied_rooms) if occupied_rooms and float(occupied_rooms).is_integer() else occupied_rooms,
                "cleaned_rooms": int(cleaned_rooms) if cleaned_rooms and float(cleaned_rooms).is_integer() else cleaned_rooms,
                "dd_rooms": int(dd_rooms) if dd_rooms and float(dd_rooms).is_integer() else dd_rooms,
                "single_rooms": int(single_rooms) if single_rooms and float(single_rooms).is_integer() else single_rooms,
                "twin_rooms": int(twin_rooms) if twin_rooms and float(twin_rooms).is_integer() else twin_rooms,
                "double_rooms": int(double_rooms) if double_rooms and float(double_rooms).is_integer() else double_rooms,
                "extra_rooms": int(extra_rooms) if extra_rooms and float(extra_rooms).is_integer() else extra_rooms,
                "vcan_checks": int(vcan_checks) if vcan_checks and float(vcan_checks).is_integer() else vcan_checks,
            }
        )
    return entries


def merge_entries(daily_entries: list[dict], actual_entries: list[dict]) -> list[dict]:
    merged = {entry["date"]: dict(entry) for entry in actual_entries}
    for entry in daily_entries:
        date = entry["date"]
        merged.setdefault(date, {}).update(entry)
        merged[date]["date"] = date
    return sorted(merged.values(), key=lambda item: item["date"])


def summarize(entries: list[dict]) -> dict:
    cleaned = [e["cleaned_rooms"] for e in entries if e.get("cleaned_rooms")]
    occupied = [e["occupied_rooms"] for e in entries if e.get("occupied_rooms")]
    workloads = [e["workload"] for e in entries if e.get("workload")]
    maids = [e["maids"] for e in entries if e.get("maids")]
    checkers = [e["checkers"] for e in entries if e.get("checkers")]
    times = [e["completion_time"] for e in entries if e.get("completion_time")]
    return {
        "days": len(entries),
        "date_min": min((e["date"] for e in entries), default=None),
        "date_max": max((e["date"] for e in entries), default=None),
        "total_cleaned_rooms": int(sum(cleaned)) if cleaned else 0,
        "total_occupied_rooms": int(sum(occupied)) if occupied else 0,
        "total_workload": int(sum(workloads)) if workloads else 0,
        "total_claims": int(sum(e.get("claims", 0) or 0 for e in entries)),
        "avg_cleaned_rooms": round(sum(cleaned) / len(cleaned), 1) if cleaned else None,
        "avg_occupied_rooms": round(sum(occupied) / len(occupied), 1) if occupied else None,
        "avg_workload": round(sum(workloads) / len(workloads), 1) if workloads else None,
        "avg_maids": round(sum(maids) / len(maids), 1) if maids else None,
        "avg_checkers": round(sum(checkers) / len(checkers), 1) if checkers else None,
        "avg_completion_time": round(sum(times) / len(times), 2) if times else None,
        "time_data_points": len(times),
        "cleaned_room_data_points": len(cleaned),
        "workload_data_points": len(workloads),
    }


def main() -> int:
    base_url = os.environ.get("GAS_CSV_PROXY_URL", "").strip()
    if not base_url:
        print("GAS_CSV_PROXY_URL is not set; skipping cleaning actuals download")
        return 0

    os.makedirs(JSON_DIR, exist_ok=True)
    start_month = int(os.environ.get("CLEANING_START_MONTH", TODAY.month))
    target_month = min(TODAY.month, 12)
    all_hotels = {}
    failures = []

    for key, name, spreadsheet_id in HOTELS:
        print(f"Downloading cleaning actuals: {key}", flush=True)
        try:
            hotel_entries = []
            fetched_sheets = []
            for month in range(start_month, target_month + 1):
                actual_entries = []
                aggregate_sheet = f"①R8_{month}集計"
                csv_text = fetch_text(
                    base_url,
                    {
                        "id": spreadsheet_id,
                        "sheet": aggregate_sheet,
                        "values": "display",
                    },
                )
                actual_entries = extract_monthly_actuals(parse_csv(csv_text), month)
                if actual_entries:
                    fetched_sheets.append(aggregate_sheet)

                month_entries = None
                for sheet_name in daily_sheet_candidates(month):
                    csv_text = fetch_text(
                        base_url,
                        {
                            "id": spreadsheet_id,
                            "sheet": sheet_name,
                            "values": "display",
                        },
                    )
                    entries = extract_entries(parse_csv(csv_text), month)
                    if entries:
                        month_entries = entries
                        fetched_sheets.append(sheet_name)
                        break
                hotel_entries.extend(merge_entries(month_entries or [], actual_entries))
            hotel_entries.sort(key=lambda item: item["date"])
            all_hotels[key] = {
                "name": name,
                "spreadsheet_id": spreadsheet_id,
                "sheets": fetched_sheets,
                "summary": summarize(hotel_entries),
                "daily_entries": hotel_entries,
            }
            print(
                f"  OK {len(hotel_entries)} days, "
                f"max={all_hotels[key]['summary']['date_max']}"
                ,
                flush=True,
            )
        except Exception as exc:  # noqa: BLE001
            print(f"  NG {key}: {exc}", flush=True)
            failures.append({"hotel_key": key, "error": str(exc)})

    all_entries = [
        entry
        for hotel in all_hotels.values()
        for entry in hotel.get("daily_entries", [])
    ]
    no_data = [
        key
        for key, hotel in all_hotels.items()
        if not hotel.get("daily_entries")
    ]
    output = {
        "metadata": {
            "title": "ホテル清掃実績（日報）",
            "source": "Google Sheets daily report tabs via GAS CSV proxy",
            "generated_date": TODAY.isoformat(),
            "target_period": f"2026-{start_month:02d}-01〜{TODAY.isoformat()}",
            "hotels_total": len(HOTELS),
            "hotels_downloaded": len(all_hotels),
            "hotels_with_data": len(all_hotels) - len(no_data),
            "hotels_without_data": no_data,
            "failures": failures,
        },
        "portfolio_summary": summarize(all_entries),
        "hotels": all_hotels,
    }

    with open(OUT_PATH, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    print(f"Saved {OUT_PATH}")
    if failures:
        print(f"Completed with {len(failures)} failures")
    return 0 if all_hotels else 1


if __name__ == "__main__":
    raise SystemExit(main())
