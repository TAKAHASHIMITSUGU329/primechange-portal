#!/bin/bash
# ホテル口コミ データ最新化 + 納品レポート再生成 統合スクリプト
# 使い方: bash スクリプト/最新化/full_update_with_reports.sh

set -e
BASE_DIR="/Users/mitsugutakahashi/ホテル口コミ"
JSON_DIR="$BASE_DIR/データ/分析結果JSON"
REPORT_DIR="$BASE_DIR/スクリプト/レポート生成"
HOTEL_OUTPUT="$BASE_DIR/納品レポート/ホテル別レポート"
PC_OUTPUT="$BASE_DIR/納品レポート/PRIMECHANGE戦略レポート"
TODAY=$(date '+%Y-%m-%d')
BACKUP_DIR="$JSON_DIR/backup_$(date -v-1d '+%Y-%m-%d' 2>/dev/null || date -d 'yesterday' '+%Y-%m-%d' 2>/dev/null || echo 'prev')"

echo "============================================"
echo "ホテル口コミ 全体最新化パイプライン"
echo "実行日: $TODAY"
echo "============================================"

# ステップ1: バックアップ
echo ""
echo "=== ステップ1: 旧データバックアップ ==="
mkdir -p "$BACKUP_DIR"
cp "$JSON_DIR"/*.json "$BACKUP_DIR/" 2>/dev/null || true
echo "  バックアップ先: $BACKUP_DIR"
echo "  ファイル数: $(ls "$BACKUP_DIR"/*.json 2>/dev/null | wc -l)"

# ステップ2: データ最新化
echo ""
echo "=== ステップ2: データ最新化パイプライン ==="
bash "$BASE_DIR/スクリプト/最新化/update_all.sh"

# ステップ3: 差分レポート
echo ""
echo "=== ステップ3: 差分レポート生成 ==="
python3 "$BASE_DIR/スクリプト/最新化/generate_diff_report.py" || echo "  差分レポート生成スキップ"

# ステップ4: ホテル別レポート再生成
echo ""
echo "=== ステップ4: 19ホテル納品レポート再生成 ==="
REPORT_SCRIPT="$REPORT_DIR/generate_hotel_report.js"

declare -a HOTELS=(
  "daiwa_osaki|ダイワロイネットホテル東京大崎"
  "chisan|チサンホテル浜松町"
  "hearton|ハートンホテル東品川"
  "keyakigate|ホテルケヤキゲート東京府中"
  "richmond_mejiro|リッチモンドホテル東京目白"
  "keisei_kinshicho|京成リッチモンドホテル東京錦糸町"
  "daiichi_ikebukuro|第一イン池袋"
  "comfort_roppongi|コンフォートイン六本木"
  "comfort_suites_tokyobay|コンフォートスイーツ東京ベイ"
  "comfort_era_higashikanda|コンフォートホテルERA東京東神田"
  "comfort_yokohama_kannai|コンフォートホテル横浜関内"
  "comfort_narita|コンフォートホテル成田"
  "apa_kamata|アパホテル蒲田駅東"
  "apa_sagamihara|アパホテル相模原橋本駅東"
  "court_shinyokohama|コートホテル新横浜"
  "comment_yokohama|ホテルコメント横浜関内"
  "kawasaki_nikko|川崎日航ホテル"
  "henn_na_haneda|変なホテル東京羽田"
  "comfort_hakata|コンフォートホテル博多"
)

HOTEL_OK=0
HOTEL_FAIL=0
for entry in "${HOTELS[@]}"; do
  IFS='|' read -r key name <<< "$entry"
  JSON_FILE="$JSON_DIR/${key}_analysis.json"
  echo -n "  生成: $name ... "
  if node "$REPORT_SCRIPT" "$name" "$JSON_FILE" "$key" 2>/dev/null; then
    mv "$JSON_DIR/${key}_口コミ分析改善レポート.docx" "$HOTEL_OUTPUT/" 2>/dev/null
    mv "$JSON_DIR/${key}_口コミ分析レポート.pptx" "$HOTEL_OUTPUT/" 2>/dev/null
    echo "OK"
    HOTEL_OK=$((HOTEL_OK + 1))
  else
    echo "FAIL"
    HOTEL_FAIL=$((HOTEL_FAIL + 1))
  fi
done
echo "  完了: $HOTEL_OK 成功 / $HOTEL_FAIL 失敗"

# ステップ5: PRIMECHANGE戦略レポート再生成
echo ""
echo "=== ステップ5: PRIMECHANGE戦略レポート再生成 ==="

# Symlinks
ln -sf "$JSON_DIR/primechange_portfolio_analysis.json" "$REPORT_DIR/primechange_portfolio_analysis.json"
ln -sf "$JSON_DIR/hotel_revenue_data.json" "$REPORT_DIR/hotel_revenue_data.json"
for i in 1 2 3 4 5 6 7; do
  ln -sf "$JSON_DIR/analysis_${i}_data.json" "$REPORT_DIR/analysis_${i}_data.json"
done

PC_SCRIPTS=(
  "generate_primechange_report.js"
  "generate_cs_strategy_report.js"
  "generate_cs_strategy_pptx.js"
  "generate_analysis_1_report.js"
  "generate_analysis_2_report.js"
  "generate_analysis_3_report.js"
  "generate_analysis_4_report.js"
  "generate_analysis_5_report.js"
  "generate_analysis_6_report.js"
  "generate_analysis_7_report.js"
  "generate_brainstorm_report.js"
  "generate_quality_revenue_report.js"
)

PC_OK=0
PC_FAIL=0
for script in "${PC_SCRIPTS[@]}"; do
  echo -n "  実行: $script ... "
  if node "$REPORT_DIR/$script" 2>/dev/null; then
    echo "OK"
    PC_OK=$((PC_OK + 1))
  else
    echo "FAIL"
    PC_FAIL=$((PC_FAIL + 1))
  fi
done

# Move outputs
for f in "$REPORT_DIR"/PRIMECHANGE_*.docx "$REPORT_DIR"/PRIMECHANGE_*.pptx; do
  [ -f "$f" ] && mv "$f" "$PC_OUTPUT/"
done

# Cleanup symlinks
rm -f "$REPORT_DIR/primechange_portfolio_analysis.json" "$REPORT_DIR/hotel_revenue_data.json"
for i in 1 2 3 4 5 6 7; do
  rm -f "$REPORT_DIR/analysis_${i}_data.json"
done
echo "  完了: $PC_OK 成功 / $PC_FAIL 失敗"

# ステップ6: ダッシュボード & スナップショット生成
echo ""
echo "=== ステップ6: ダッシュボード生成 & スナップショット保存 ==="
BUILD_SCRIPT="$BASE_DIR/スクリプト/ホームページ生成/build_all.js"
if node "$BUILD_SCRIPT"; then
  echo "  ダッシュボード生成: OK"

  # スナップショット検証
  SNAPSHOT_DIR="$BASE_DIR/ホームページ/data/snapshots/$TODAY"
  SNAPSHOT_INDEX="$BASE_DIR/ホームページ/data/snapshot-index.json"
  if [ -d "$SNAPSHOT_DIR" ] && [ -f "$SNAPSHOT_DIR/hotel-reviews-all.json" ] && [ -f "$SNAPSHOT_DIR/hotel-details.json" ]; then
    SNAP_SIZE=$(du -sh "$SNAPSHOT_DIR" | cut -f1)
    echo "  スナップショット保存: OK ($SNAPSHOT_DIR, $SNAP_SIZE)"
  else
    echo "  WARNING: スナップショットが正しく保存されていません: $SNAPSHOT_DIR"
  fi

  if [ -f "$SNAPSHOT_INDEX" ]; then
    SNAP_COUNT=$(node -e "var d=require('$SNAPSHOT_INDEX');console.log(d.length)" 2>/dev/null || echo "?")
    echo "  スナップショット数: $SNAP_COUNT"
    if [ "$SNAP_COUNT" != "?" ] && [ "$SNAP_COUNT" -gt 100 ] 2>/dev/null; then
      echo "  WARNING: スナップショットが100件を超えています。クリーンアップを検討してください。"
    fi
  fi
else
  echo "  ダッシュボード生成: FAIL"
fi

# ステップ7: 最終サマリー
echo ""
echo "============================================"
echo "全体完了: $TODAY"
echo "  ホテル別レポート: $HOTEL_OK/19"
echo "  PRIMECHANGEレポート: $PC_OK/12"
echo "  ダッシュボード: $BASE_DIR/ホームページ/"
echo "  スナップショット: $BASE_DIR/ホームページ/data/snapshots/$TODAY/"
echo "  差分レポート: $JSON_DIR/diff_report_${TODAY}.txt"
echo "============================================"
