# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

PRIMECHANGEホテル口コミ分析ポータル。19ホテルのGoogle Spreadsheetから口コミ・月報データを取得し、分析結果をダッシュボード（HTML）とレポート（DOCX/PPTX）として出力するシステム。

## Key Commands

### データ更新パイプライン（日次）
```bash
# 全パイプライン実行（CSV/XLSX取得→分析→ダッシュボード生成）
bash スクリプト/最新化/update_all.sh

# フル実行（上記＋19ホテルレポート＋12戦略レポート再生成）
bash スクリプト/最新化/full_update_with_reports.sh
```

### 個別実行
```bash
# XLSXダウンロード（19ホテル月報）
# → .github/workflows/download-csvs.yml のXLSX_MAPを参照

# 深掘り分析データ抽出（analysis_1〜7）
python3 スクリプト/データ抽出/analysis_2_extract.py  # スタッフパフォーマンス
python3 スクリプト/データ抽出/analysis_6_extract.py  # 品質×売上相関

# ダッシュボードHTML生成
node スクリプト/ホームページ生成/build_all.js      # V1 → ホームページ/
node スクリプト/ホームページ生成/build_all_v2.js    # V2 → ホームページV2/

# ホテル個別レポート生成
node スクリプト/レポート生成/generate_hotel_report.js "ホテル名" データ/分析結果JSON/{key}_analysis.json {key}
```

### デプロイ
```bash
# ホームページV2/ → v2/ に同期後、git push
rsync -a ホームページV2/ v2/
# ホームページ/ → ルート直下に同期
cp ホームページ/{deep-analysis,hotel-dashboard,index,cleaning-strategy,action-plans,revenue-impact}.html .
```

## Architecture

### データフロー
```
Google Sheets (19ホテル)
  ↓ curl (CSV + XLSX)
データ/ホテル口コミデータ/*.csv          ← 口コミ
データ/ホテル集計表XLSX/*.xlsx          ← 月報（清掃・売上・皆勤）
  ↓
analyze_reviews.py × 19 → データ/分析結果JSON/{key}_analysis.json
aggregate_portfolio_analysis.py → primechange_portfolio_analysis.json
analysis_1〜7_extract.py → analysis_{1-7}_data.json
  ↓
build_all.js    → ホームページ/   (V1: 5ページ)   → ルート直下HTML
build_all_v2.js → ホームページV2/ (V2: 11ページ)  → v2/
```

### ディレクトリ構造の重要ポイント

- **`ホームページ/` と `ホームページV2/`** は .gitignore 対象。ビルド出力先。
- **`v2/`** と **ルート直下HTML** が GitHub Pages 公開対象（git tracked）。
- **`スクリプト/データ抽出/`** にある JSON は analysis_*.py の出力先。`データ/分析結果JSON/` にコピーが必要（`analysis_6_extract.py` 等はローカル出力）。

### analysis_6_extract.py の注意点
- `BASE_DIR` がスクリプトディレクトリ（`スクリプト/データ抽出/`）を指す
- `hotel_revenue_data.json` と `primechange_portfolio_analysis.json` をローカルから読む
- **最新データで実行する場合、先に `データ/分析結果JSON/` からこれらをコピーする必要あり**

### V2ダッシュボード構成（build_all_v2.js）
- `lib/data-loader.js` — 全JSON読み込み・KPIターゲット解析
- `lib/revenue-calc.js` — RevPAR弾力性・改善機会計算・formatYen()
- `lib/delta-engine.js` — 前日比差分計算
- `lib/cs-analyzer.js` — CS 6軸分析
- `lib/deep-analysis-renderers.js` — 7分析タブのHTML生成（renderA1〜A7）
- `lib/page-*.js` — 各ページジェネレータ

### XLSXシート構造（皆勤アワードシート）
月別清掃データのカラム配置:
- 2月: col21(日数), col23(時間), col24(部屋数)
- 3月: col31(日数), col33(時間), col34(部屋数)
- 4月: col41(日数), col43(時間), col44(部屋数)

### 売上データ（hotel_revenue_data.json）
- `actual_revenue` / `occupancy_rate` / `adr` / `revpar` — 2月
- `march_revenue` / `march_occupancy` — 3月
- `april_revenue` / `april_occupancy` — 4月

### ホテルキーマッピング
`hotel_xlsx_utils.py` の `KEY_MAP` でCSVキーとXLSXキーの差異を吸収:
- `keisei_kinshicho` → `keisei_richmond`
- `comfort_yokohama_kannai` → `comfort_yokohama`

## Tech Stack
- **Python 3**: openpyxl, numpy（データ抽出・分析）
- **Node.js**: docx, pptxgenjs, pdfkit（レポート生成・HTML生成）
- **GitHub Actions**: 日次CSV/XLSXダウンロード
- **GitHub Pages**: ダッシュボード公開

## 言語
全コミュニケーション・コメントは日本語で行うこと。
