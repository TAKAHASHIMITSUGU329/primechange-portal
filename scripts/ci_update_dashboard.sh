#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
CSV_DIR="$ROOT_DIR/データ/ホテル口コミデータ"
XLSX_DIR="$ROOT_DIR/データ/ホテル集計表XLSX"
JSON_DIR="$ROOT_DIR/データ/分析結果JSON"
ANALYZE_SCRIPT="$ROOT_DIR/.claude/skills/hotel-review-report/scripts/analyze_reviews.py"
DATA_SCRIPT_DIR="$ROOT_DIR/スクリプト/データ抽出"
BUILD_SCRIPT_DIR="$ROOT_DIR/スクリプト/ホームページ生成"

mkdir -p "$CSV_DIR" "$XLSX_DIR" "$JSON_DIR"

HOTELS=(
  "daiwa_osaki|1IIHEn4nAIy9UXzrYptU-RQIiTbKkV0G_CaABF7znVrY|0|daiwa_osaki_data.csv|ダイワロイネットホテル東京大崎"
  "chisan|1IWigsWTzbRG-juWtIlg4ZchiuWqRhJFpPPczXdQxG6Y|0|chisan_data_converted.csv|チサンホテル浜松町"
  "hearton|1A25mmVRYSnG3ZB8oa0oZVp-vCP2xMkwX-Zqdkk4BIzI|0|hearton_data.csv|ハートンホテル東品川"
  "keyakigate|1srchDxFyv7TJ3IEZXJ19miH04p3jRug5nVtA3BLertQ|605247000|keyakigate_data.csv|ホテルケヤキゲート東京府中"
  "richmond_mejiro|1XWU6925CpT3GMMonAqy4UENKM11gWloUkJIsgGImUts|0|richmond_mejiro_data.csv|リッチモンドホテル東京目白"
  "keisei_kinshicho|1jUS_HwTfowG1xIHFtwJbCL5dTj7FrhvUe6d32AevZ2g|0|keisei_kinshicho_data.csv|京成リッチモンドホテル東京錦糸町"
  "daiichi_ikebukuro|1X2GgFKxTOs7CuJSlPYrpzigraSnWcKh6cMJLfsXhWlU|0|daiichi_ikebukuro_data.csv|第一イン池袋"
  "comfort_roppongi|1Jtm0rXTigY2OVManNjx1qQ6G9EKQEXuPs_T1BdlOvls|0|comfort_roppongi_data.csv|コンフォートイン六本木"
  "comfort_suites_tokyobay|1zCFAmzRqvSDbjwvK7qI4cYBHlrmBifTPm0Y-g0rruyE|0|comfort_suites_tokyobay_data.csv|コンフォートスイーツ東京ベイ"
  "comfort_era_higashikanda|1H9jmOVQR4UdEQ5hsxZ2Xz44BT72RJDwNa6BOKFhXxRg|0|comfort_era_higashikanda_data.csv|コンフォートホテルERA東京東神田"
  "comfort_yokohama_kannai|1rnQOsyUXuSzBKdqPN_ey_4Iw5VtYWTgSR5Z4nh-1zd4|0|comfort_yokohama_kannai_data.csv|コンフォートホテル横浜関内"
  "comfort_narita|1lQ3FRDuE75dkByQRFd0i0F2xcHnl-3-UAOJwhIt3jAU|0|comfort_narita_data.csv|コンフォートホテル成田"
  "apa_kamata|16xuhAdNzdeyAKu-LhU8ATgR8_kZ1JXfa9lT51tAB1Nw|0|apa_kamata_data.csv|アパホテル蒲田駅東"
  "apa_sagamihara|1E2ZQJyE6pOJ3jr6GyB56KcYnVVq54m6dO_6h_SQy39A|0|apa_sagamihara_data.csv|アパホテル相模原橋本駅東"
  "court_shinyokohama|1Qm5lPPc8m7yutyIH3Pf03YUnF2KpnWjn0SecMzq0CjY|0|court_shinyokohama_data.csv|コートホテル新横浜"
  "comment_yokohama|1cVH7khdgh8bDN-wtAw2KVakJqHILo58VOBu0SKmBFrU|0|comment_yokohama_data.csv|ホテルコメント横浜関内"
  "kawasaki_nikko|1aQ2MaKJmOz7eT53oqszCDO9Fa3UEbfhFSgXfVmVpO9A|0|kawasaki_nikko_data.csv|川崎日航ホテル"
  "henn_na_haneda|18DkZLJ8UDQ2-4MBrh7B4y28tHaYnoWIQqEoFkvFDNKg|2026949334|henn_na_haneda_data.csv|変なホテル東京羽田"
  "comfort_hakata|1_7xoyIiq1llfO0I2328ZQlB6sD0lMsnpRMp1rMGNPcg|0|comfort_hakata_data.csv|コンフォートホテル博多"
)

XLSX_FILES=(
  "daiwa_osaki|1IIHEn4nAIy9UXzrYptU-RQIiTbKkV0G_CaABF7znVrY|R8_P1_ダイワロイネットホテル東京大崎_集計表.xlsx"
  "chisan|1IWigsWTzbRG-juWtIlg4ZchiuWqRhJFpPPczXdQxG6Y|R8_P1_チサンホテル浜松町_集計表.xlsx"
  "hearton|1A25mmVRYSnG3ZB8oa0oZVp-vCP2xMkwX-Zqdkk4BIzI|R8_P1_ハートンホテル東品川_集計表.xlsx"
  "keyakigate|1srchDxFyv7TJ3IEZXJ19miH04p3jRug5nVtA3BLertQ|R8_P1_ホテルケヤキゲート東京府中_集計表.xlsx"
  "richmond_mejiro|1XWU6925CpT3GMMonAqy4UENKM11gWloUkJIsgGImUts|R8_P1_リッチモンドホテル東京目白_集計表.xlsx"
  "keisei_kinshicho|1jUS_HwTfowG1xIHFtwJbCL5dTj7FrhvUe6d32AevZ2g|R8_P1_京成リッチモンドホテル東京錦糸町_集計表.xlsx"
  "daiichi_ikebukuro|1X2GgFKxTOs7CuJSlPYrpzigraSnWcKh6cMJLfsXhWlU|R8_P1_第一イン池袋_集計表.xlsx"
  "comfort_roppongi|1Jtm0rXTigY2OVManNjx1qQ6G9EKQEXuPs_T1BdlOvls|R8_P2_コンフォートイン六本木_集計表.xlsx"
  "comfort_suites_tokyobay|1zCFAmzRqvSDbjwvK7qI4cYBHlrmBifTPm0Y-g0rruyE|R8_P2_コンフォートスイーツ東京ベイ_集計表.xlsx"
  "comfort_era_higashikanda|1H9jmOVQR4UdEQ5hsxZ2Xz44BT72RJDwNa6BOKFhXxRg|R8_P2_コンフォートホテルERA東京東神田_集計表.xlsx"
  "comfort_narita|1lQ3FRDuE75dkByQRFd0i0F2xcHnl-3-UAOJwhIt3jAU|R8_P2_コンフォートホテル成田_集計表.xlsx"
  "comfort_yokohama_kannai|1rnQOsyUXuSzBKdqPN_ey_4Iw5VtYWTgSR5Z4nh-1zd4|R8_P2_コンフォートホテル横浜関内_集計表.xlsx"
  "apa_kamata|16xuhAdNzdeyAKu-LhU8ATgR8_kZ1JXfa9lT51tAB1Nw|R8_P3_アパホテル蒲田駅東_集計表.xlsx"
  "apa_sagamihara|1E2ZQJyE6pOJ3jr6GyB56KcYnVVq54m6dO_6h_SQy39A|R8_P3_アパホテル相模原橋本駅東_集計表.xlsx"
  "court_shinyokohama|1Qm5lPPc8m7yutyIH3Pf03YUnF2KpnWjn0SecMzq0CjY|R8_P3_コートホテル新横浜_集計表.xlsx"
  "comment_yokohama|1cVH7khdgh8bDN-wtAw2KVakJqHILo58VOBu0SKmBFrU|R8_P3_ホテルコメント横浜関内_集計表.xlsx"
  "kawasaki_nikko|1aQ2MaKJmOz7eT53oqszCDO9Fa3UEbfhFSgXfVmVpO9A|R8_P3_川崎日航ホテル_集計表.xlsx"
  "henn_na_haneda|18DkZLJ8UDQ2-4MBrh7B4y28tHaYnoWIQqEoFkvFDNKg|R8_P3_変なホテル東京羽田_集計表.xlsx"
  "comfort_hakata|1_7xoyIiq1llfO0I2328ZQlB6sD0lMsnpRMp1rMGNPcg|R8_P4_コンフォートホテル博多_集計表.xlsx"
)

download_file() {
  local url="$1"
  local out="$2"
  local code
  code="$(curl -sL -o "$out" -w "%{http_code}" "$url")"
  printf '%s' "$code"
}

echo "=== Step 1: Download CSV data ==="
csv_success=0
csv_fail=0
for entry in "${HOTELS[@]}"; do
  IFS='|' read -r key sheet_id gid filename hotel_name <<< "$entry"
  out="$CSV_DIR/$filename"
  tmp="${out}.tmp"
  code=""

  if [[ -n "${GAS_CSV_PROXY_URL:-}" ]]; then
    code="$(download_file "${GAS_CSV_PROXY_URL}?key=${key}" "$tmp")"
    if [[ "$code" != "200" ]] || grep -q '^ERROR:' "$tmp"; then
      rm -f "$tmp"
      code=""
    fi
  fi

  if [[ -z "$code" ]]; then
    code="$(download_file "https://docs.google.com/spreadsheets/d/${sheet_id}/gviz/tq?tqx=out:csv&gid=${gid}" "$tmp")"
  fi

  if [[ "$code" == "200" ]] && [[ -s "$tmp" ]] && ! grep -q '^ERROR:' "$tmp"; then
    mv "$tmp" "$out"
    lines="$(wc -l < "$out" | tr -d ' ')"
    echo "OK  ${key}: ${lines} lines"
    csv_success=$((csv_success + 1))
  else
    rm -f "$tmp"
    echo "NG  ${key}: HTTP ${code:-proxy-error}"
    csv_fail=$((csv_fail + 1))
  fi
done
echo "CSV summary: ${csv_success}/19 success, ${csv_fail}/19 failed"

if [[ "$csv_success" -eq 0 ]]; then
  echo "::error::All CSV downloads failed. Check Google Sheets sharing or set GAS_CSV_PROXY_URL."
  exit 1
fi

echo "=== Step 2: Download XLSX data ==="
xlsx_success=0
xlsx_fail=0
for entry in "${XLSX_FILES[@]}"; do
  IFS='|' read -r key sheet_id filename <<< "$entry"
  out="$XLSX_DIR/$filename"
  tmp="${out}.tmp"
  code="$(download_file "https://docs.google.com/spreadsheets/d/${sheet_id}/export?format=xlsx" "$tmp")"
  if [[ "$code" == "200" ]] && [[ -s "$tmp" ]]; then
    mv "$tmp" "$out"
    bytes="$(wc -c < "$out" | tr -d ' ')"
    echo "OK  ${key}: ${bytes} bytes"
    xlsx_success=$((xlsx_success + 1))
  else
    rm -f "$tmp"
    echo "NG  ${key}: HTTP ${code}"
    xlsx_fail=$((xlsx_fail + 1))
  fi
done
echo "XLSX summary: ${xlsx_success}/19 success, ${xlsx_fail}/19 failed"

echo "=== Step 3: Analyze hotel reviews ==="
analysis_success=0
analysis_fail=0
end_month="$(date '+%Y-%m')"
for entry in "${HOTELS[@]}"; do
  IFS='|' read -r key _sheet_id _gid filename hotel_name <<< "$entry"
  csv="$CSV_DIR/$filename"
  out="$JSON_DIR/${key}_analysis.json"
  if [[ ! -s "$csv" ]]; then
    echo "SKIP ${key}: missing CSV"
    analysis_fail=$((analysis_fail + 1))
    continue
  fi
  if python3 "$ANALYZE_SCRIPT" --csv "$csv" --start-month 2025-12 --end-month "$end_month" --output "$out"; then
    echo "OK  ${key}"
    analysis_success=$((analysis_success + 1))
  else
    echo "NG  ${key}"
    analysis_fail=$((analysis_fail + 1))
  fi
done
echo "Analysis summary: ${analysis_success}/19 success, ${analysis_fail}/19 failed"

if [[ "$analysis_success" -eq 0 ]]; then
  echo "::error::All hotel analyses failed."
  exit 1
fi

echo "=== Step 4: Refresh aggregate data ==="
python3 "$DATA_SCRIPT_DIR/extract_revenue_from_xlsx.py" || echo "Revenue extraction skipped"
python3 "$DATA_SCRIPT_DIR/aggregate_portfolio_analysis.py"
for script in analysis_1_extract.py analysis_2_extract.py analysis_3_4_extract.py analysis_5_extract.py analysis_6_extract.py analysis_7_extract.py; do
  python3 "$DATA_SCRIPT_DIR/$script" || echo "$script skipped"
done
for i in 1 2 3 4 5 6 7; do
  src="$DATA_SCRIPT_DIR/analysis_${i}_data.json"
  [[ -f "$src" ]] && cp "$src" "$JSON_DIR/"
done

echo "=== Step 5: Build V1 and V2 dashboards ==="
node "$BUILD_SCRIPT_DIR/build_all.js"
node "$BUILD_SCRIPT_DIR/build_all_v2.js"
cp "$ROOT_DIR/ホームページ/"*.html "$ROOT_DIR/" 2>/dev/null || true
rsync -a "$ROOT_DIR/ホームページV2/" "$ROOT_DIR/v2/"

echo "=== Dashboard update complete ==="
cat "$ROOT_DIR/v2/data/build-meta.json"
