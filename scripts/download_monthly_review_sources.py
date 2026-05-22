#!/usr/bin/env python3
"""Download monthly hotel review source sheets and combine them per hotel."""

import csv
import io
import os
import sys
import urllib.parse
import urllib.request
from datetime import datetime


RAW_SOURCES = {
    "chisan": ("1simGLviaafKfjyb9Dd5Wy58cgaN3mLgNJn-HOWastfw", "chisan_data_converted.csv"),
    "hearton": ("1wVupI9SpDOevnCI8ToI708DDOFqynpKnLhJ-mOGsUb0", "hearton_data.csv"),
    "keyakigate": ("1SDqOPzmT4NFKIkJkXFYebm99yiw8Lt1cxHygDC-7C7Y", "keyakigate_data.csv"),
    "richmond_mejiro": ("1WR5LwMaC62OxjdG8VeuCa0nRzVa1AkXFliEimhITrpQ", "richmond_mejiro_data.csv"),
    "keisei_kinshicho": ("1Fj5yWKsfm_d4H60bslB14v9l4aPtE-PVGtGSIfFZdng", "keisei_kinshicho_data.csv"),
    "daiichi_ikebukuro": ("1xy3vIyyUupB7lJ891h4NFPvfeW5fA5WETSf9a7p374A", "daiichi_ikebukuro_data.csv"),
    "comfort_roppongi": ("1V9UMPmhBQrSesFmoLkiZwv16hww3VI3YlWV6dbHFD04", "comfort_roppongi_data.csv"),
    "comfort_suites_tokyobay": ("1gMzPFEQ8nEaCZ21tAS-ZODtrr1hfwLsQdlAH03A8CHI", "comfort_suites_tokyobay_data.csv"),
    "comfort_era_higashikanda": ("1D3J1Th8bSuQR0yePsG8Ydjg6YueVegw0EBkMxbOmCQo", "comfort_era_higashikanda_data.csv"),
    "comfort_yokohama_kannai": ("1xf_sbSdIm-wiiLV4zl4kkEBR3UPl85IOCttzImI7POc", "comfort_yokohama_kannai_data.csv"),
    "comfort_narita": ("1ybR5mSan4yLf5gpXtX_xRrIZV00CZSCfUru2auXo8mw", "comfort_narita_data.csv"),
    "comfort_hakata": ("1rAomn54IH2Yc1L5dTpwcIVGuzPBWOc69kxHC5VHSVf8", "comfort_hakata_data.csv"),
    "apa_kamata": ("1qmrBuYvMMf0Z6Sfp3Pr2ct1BAU2uJXp3JoTJ5eqA7IA", "apa_kamata_data.csv"),
    "court_shinyokohama": ("1DfE0guo2g1GCJpkq0DKxL7ODr3MVi8aQdiNx9i-gSpk", "court_shinyokohama_data.csv"),
    "comment_yokohama": ("1X1_gL00cx1uHkKQiu25y-ccOR_2MvVL2bZxf2WrSY0g", "comment_yokohama_data.csv"),
    "kawasaki_nikko": ("1JWelo-LNmKlmqqDiVYzMTCklyj9RoUDKtEP68ppmBqE", "kawasaki_nikko_data.csv"),
}


def has_review_header(row):
    text = " ".join(row).lower()
    return ("サイト" in text or "site" in text) and ("評価" in text or "rating" in text) and "投稿日" in text


def fetch_csv(base_url, spreadsheet_id, sheet_name):
    url = base_url + "?" + urllib.parse.urlencode({"id": spreadsheet_id, "sheet": sheet_name})
    with urllib.request.urlopen(url, timeout=90) as response:
        return response.read().decode("utf-8-sig", "replace")


def combine_months(base_url, spreadsheet_id, months):
    output_rows = []
    header = None
    count = 0
    max_date = ""

    for sheet_name in months:
        text = fetch_csv(base_url, spreadsheet_id, sheet_name)
        rows = list(csv.reader(io.StringIO(text)))
        header_index = next((i for i, row in enumerate(rows[:10]) if has_review_header(row)), None)
        if header_index is None:
            continue
        if header is None:
            header = rows[header_index]
            output_rows.append(header)
        for row in rows[header_index + 1:]:
            if len(row) < 4:
                continue
            site = row[1].strip()
            rating = row[2].strip()
            date = row[3].strip()[:10]
            if not site or not rating or site.lower() in ("true", "false"):
                continue
            output_rows.append(row)
            count += 1
            if date > max_date:
                max_date = date

    return output_rows, count, max_date


def main():
    base_url = os.environ.get("GAS_CSV_PROXY_URL", "").rstrip("/")
    if not base_url:
        print("GAS_CSV_PROXY_URL is not set; skipping monthly review sources")
        return 0

    current = datetime.now()
    months = [f"{current.year}年{month:02d}月" for month in range(1, current.month + 1)]
    root = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    out_dir = os.path.join(root, "データ", "ホテル口コミデータ")
    os.makedirs(out_dir, exist_ok=True)

    success = 0
    for key, (spreadsheet_id, filename) in RAW_SOURCES.items():
        try:
            rows, count, max_date = combine_months(base_url, spreadsheet_id, months)
        except Exception as exc:
            print(f"NG  {key}: {exc}", file=sys.stderr)
            continue
        if not rows or count == 0:
            print(f"NG  {key}: no review rows", file=sys.stderr)
            continue
        path = os.path.join(out_dir, filename)
        with open(path, "w", encoding="utf-8-sig", newline="") as handle:
            csv.writer(handle).writerows(rows)
        print(f"OK  {key}: {count} rows, max={max_date}")
        success += 1

    print(f"Monthly review source summary: {success}/{len(RAW_SOURCES)} success")
    return 0 if success else 1


if __name__ == "__main__":
    raise SystemExit(main())
