#!/bin/bash
# ホテル口コミデータ最新化スクリプト
# 2026-03-11 実行

set -e
BASE_DIR="/Users/mitsugutakahashi/HotelReview"
DATA_DIR="$BASE_DIR/データ/ホテル口コミデータ"
JSON_DIR="$BASE_DIR/データ/分析結果JSON"
SCRIPT_DIR="$BASE_DIR/スクリプト"
ANALYZE="$BASE_DIR/hotel-review-report/scripts/analyze_reviews.py"

# Hotel list: key, spreadsheet_id, name
declare -a HOTELS=(
  "daiwa_osaki|1IIHEn4nAIy9UXzrYptU-RQIiTbKkV0G_CaABF7znVrY|ダイワロイネットホテル東京大崎"
  "chisan|1IWigsWTzbRG-juWtIlg4ZchiuWqRhJFpPPczXdQxG6Y|チサンホテル浜松町"
  "hearton|1A25mmVRYSnG3ZB8oa0oZVp-vCP2xMkwX-Zqdkk4BIzI|ハートンホテル東品川"
  "keyakigate|1srchDxFyv7TJ3IEZXJ19miH04p3jRug5nVtA3BLertQ|ホテルケヤキゲート東京府中|605247000"
  "richmond_mejiro|1XWU6925CpT3GMMonAqy4UENKM11gWloUkJIsgGImUts|リッチモンドホテル東京目白"
  "keisei_kinshicho|1jUS_HwTfowG1xIHFtwJbCL5dTj7FrhvUe6d32AevZ2g|京成リッチモンドホテル東京錦糸町"
  "daiichi_ikebukuro|1X2GgFKxTOs7CuJSlPYrpzigraSnWcKh6cMJLfsXhWlU|第一イン池袋"
  "comfort_roppongi|1Jtm0rXTigY2OVManNjx1qQ6G9EKQEXuPs_T1BdlOvls|コンフォートイン六本木"
  "comfort_suites_tokyobay|1zCFAmzRqvSDbjwvK7qI4cYBHlrmBifTPm0Y-g0rruyE|コンフォートスイーツ東京ベイ"
  "comfort_era_higashikanda|1H9jmOVQR4UdEQ5hsxZ2Xz44BT72RJDwNa6BOKFhXxRg|コンフォートホテルERA東京東神田"
  "comfort_yokohama_kannai|1rnQOsyUXuSzBKdqPN_ey_4Iw5VtYWTgSR5Z4nh-1zd4|コンフォートホテル横浜関内"
  "comfort_narita|1lQ3FRDuE75dkByQRFd0i0F2xcHnl-3-UAOJwhIt3jAU|コンフォートホテル成田"
  "apa_kamata|16xuhAdNzdeyAKu-LhU8ATgR8_kZ1JXfa9lT51tAB1Nw|アパホテル蒲田駅東"
  "apa_sagamihara|1E2ZQJyE6pOJ3jr6GyB56KcYnVVq54m6dO_6h_SQy39A|アパホテル相模原橋本駅東"
  "court_shinyokohama|1Qm5lPPc8m7yutyIH3Pf03YUnF2KpnWjn0SecMzq0CjY|コートホテル新横浜"
  "comment_yokohama|1cVH7khdgh8bDN-wtAw2KVakJqHILo58VOBu0SKmBFrU|ホテルコメント横浜関内"
  "kawasaki_nikko|1aQ2MaKJmOz7eT53oqszCDO9Fa3UEbfhFSgXfVmVpO9A|川崎日航ホテル"
  "henn_na_haneda|18DkZLJ8UDQ2-4MBrh7B4y28tHaYnoWIQqEoFkvFDNKg|変なホテル東京羽田|2026949334"
  "comfort_hakata|1_7xoyIiq1llfO0I2328ZQlB6sD0lMsnpRMp1rMGNPcg|コンフォートホテル博多"
)

echo "=== ステップ1: 19ホテルのCSVダウンロード ==="
DOWNLOAD_OK=0
DOWNLOAD_FAIL=0
for entry in "${HOTELS[@]}"; do
  IFS='|' read -r key sid name gid <<< "$entry"
  GID="${gid:-0}"
  URL="https://docs.google.com/spreadsheets/d/${sid}/gviz/tq?tqx=out:csv&gid=${GID}"
  OUTFILE="$DATA_DIR/${key}_data.csv"
  echo -n "  ダウンロード: $name ... "
  HTTP_CODE=$(curl -sL -o "$OUTFILE" -w "%{http_code}" "$URL")
  if [ "$HTTP_CODE" = "200" ]; then
    LINES=$(wc -l < "$OUTFILE")
    echo "OK (${LINES}行)"
    DOWNLOAD_OK=$((DOWNLOAD_OK + 1))
  else
    echo "FAIL (HTTP $HTTP_CODE)"
    DOWNLOAD_FAIL=$((DOWNLOAD_FAIL + 1))
  fi
done
echo "  完了: $DOWNLOAD_OK 成功 / $DOWNLOAD_FAIL 失敗"

if [ "$DOWNLOAD_FAIL" -gt 0 ]; then
  echo "WARNING: 一部ダウンロードに失敗しました。続行しますが結果に影響する可能性があります。"
fi

echo ""
echo "=== ステップ2: 各ホテルの口コミ分析 ==="
ANALYZE_OK=0
ANALYZE_FAIL=0
for entry in "${HOTELS[@]}"; do
  IFS='|' read -r key sid name <<< "$entry"
  CSV="$DATA_DIR/${key}_data.csv"
  JSON_OUT="$JSON_DIR/${key}_analysis.json"
  if [ ! -f "$CSV" ]; then
    echo "  スキップ: $name (CSVなし)"
    ANALYZE_FAIL=$((ANALYZE_FAIL + 1))
    continue
  fi
  echo -n "  分析: $name ... "
  if python3 "$ANALYZE" --csv "$CSV" --start-month 2026-03 --end-month 2026-04 --output "$JSON_OUT" 2>/dev/null; then
    echo "OK"
    ANALYZE_OK=$((ANALYZE_OK + 1))
  else
    echo "FAIL"
    ANALYZE_FAIL=$((ANALYZE_FAIL + 1))
  fi
done
echo "  完了: $ANALYZE_OK 成功 / $ANALYZE_FAIL 失敗"

echo ""
echo "=== ステップ3: ポートフォリオ分析再集計 ==="
if python3 "$SCRIPT_DIR/データ抽出/aggregate_portfolio_analysis.py" 2>/dev/null; then
  echo "  ポートフォリオ分析: OK"
else
  echo "  ポートフォリオ分析: スキップ (既存データ使用)"
fi

echo ""
echo "=== ステップ3.5: 深掘り分析データ抽出 ==="
for script in analysis_1_extract.py analysis_2_extract.py analysis_3_4_extract.py \
              analysis_5_extract.py analysis_6_extract.py analysis_7_extract.py; do
  echo -n "  $script ... "
  if python3 "$SCRIPT_DIR/データ抽出/$script" 2>/dev/null; then
    echo "OK"
  else
    echo "スキップ"
  fi
done

echo ""
echo "=== ステップ4: ホームページ再生成 ==="
node "$SCRIPT_DIR/ホームページ生成/build_all.js"

echo ""
echo "=== 最新化完了 ==="
echo "日付: $(date '+%Y-%m-%d %H:%M:%S')"
echo "ホームページ: $BASE_DIR/ホームページ/"
