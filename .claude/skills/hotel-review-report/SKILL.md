---
name: hotel-review-report
description: >
  ホテル口コミ分析レポート（DOCX + PPTX）をGoogle Spreadsheetの口コミデータから自動生成するスキル。
  7章構成のWordレポートと10スライドのPowerPointプレゼンテーションを作成する。
  海外OTA（10点満点）と国内サイト（5点満点）の評価スケール統一も自動処理。
  「ホテル」「口コミ分析」「レビュー分析」「OTA分析」「レポート作成」「改善レポート」
  「スプレッドシートから」「口コミデータ」などのキーワードでトリガーする。
  ホテルの口コミやレビューに関するレポート作成の依頼があれば、このスキルを使う。
---

# ホテル口コミ分析レポート生成

Google Spreadsheetの口コミレビューデータからホテル分析レポート（DOCX + PPTX）を自動生成する。

## ワークフロー概要

```
Google Spreadsheet → CSV エクスポート → Python 分析 → Node.js DOCX/PPTX 生成 → 検証
```

### 前提条件
- 作業ディレクトリに `docx` と `pptxgenjs` の npm パッケージがインストール済みであること
- なければ `npm init -y && npm install docx pptxgenjs` を実行

### ユーザーに確認する項目
1. **ホテル名**（正式名称）
2. **Google Spreadsheet の URL**（口コミデータが入ったシート）
3. **分析対象期間**（例: 2026年2月〜3月）
4. **口コミシートのタブ名**（デフォルト: `💭口コミ`）

---

## Step 1: データ取得

### Google Spreadsheet からの CSV エクスポート

1. Chrome ブラウザでスプレッドシートを開く（`mcp__Claude_in_Chrome__navigate`）
2. 口コミデータのシートタブに移動
3. CSV エクスポート URL にナビゲートしてダウンロードをトリガー:
   ```
   https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/export?format=csv&gid={SHEET_GID}
   ```
   - `gid` はシートタブの URL パラメータから取得
4. ダウンロードされた CSV を `~/Downloads/` から作業ディレクトリにコピー
5. コピー後にファイルの中身を確認（`head` で先頭行を表示）

**注意**: curl では Google 認証エラーになるため、必ずブラウザ経由でエクスポートする。

---

## Step 2: データ分析

### analyze_reviews.py を使用

このスキルに同梱の `scripts/analyze_reviews.py` を実行する:

```bash
python3 <skill-path>/scripts/analyze_reviews.py \
  --csv <CSVファイルパス> \
  --start-month 2026-02 \
  --end-month 2026-03 \
  --output analysis.json
```

スクリプトが自動的に以下を行う:
- カラムの自動検出（ヘッダー名ベースで柔軟にマッピング）
- 評価スケールの自動変換（後述の変換ルール）
- 重複除外（サイト名+評価+日付+コメント先頭30文字）
- サイト別統計計算（件数、平均、10pt換算平均、中央値、判定）
- 評価分布（10pt換算後のスコア別件数）
- High/Mid/Low 集計

### 出力 JSON の構造

```json
{
  "total_reviews": 117,
  "overall_avg_10pt": 8.45,
  "high_rate": 83.8,
  "mid_rate": 14.5,
  "low_rate": 1.7,
  "high_count": 98,
  "mid_count": 17,
  "low_count": 2,
  "site_stats": [
    {"site": "Trip.com", "count": 17, "native_avg": 9.47, "scale": "/10", "avg_10pt": 9.47, "judgment": "優秀"},
    ...
  ],
  "distribution": [
    {"score": 10, "count": 44, "pct": "37.6%"},
    ...
  ],
  "comments": [
    {"site": "...", "rating": 10, "date": "...", "comment": "...", "good": "...", "bad": "..."},
    ...
  ]
}
```

### テキスト分析（Claude が実行）

`analysis.json` の `comments` を読み、以下を分析する:
1. **強み（ポジティブテーマ）**: 6つのカテゴリに分類し、言及数をカウント
   - 典型的カテゴリ: 立地・アクセス、部屋・設備、清潔感、スタッフ対応、朝食、コスパ/リピート
2. **弱み（改善課題）**: 優先度 S/A/B/C に分類
   - S: 評価に直結する最優先課題
   - A: 高頻度の不満
   - B: 中程度の改善課題
   - C: 低頻度だが留意すべき課題
3. **代表的コメントの抽出**: 各テーマの象徴的な引用を選定
4. **改善施策の提案**: Phase 1（即座対応）、Phase 2（短期）、Phase 3（中期）
5. **KPI目標設定**: 現状値 → 目標値 → 期限

---

## Step 3: DOCX レポート生成

### テンプレートの読み込み

1. `assets/docx_template.js` を読む（Read ツール）
2. テンプレート内の `// TODO: CUSTOMIZE` マーカーを確認
3. ホテル固有のデータで置き換えた新しい JS ファイルを作成:
   - ファイル名: `create_{hotel_id}_report.js`
   - 出力ファイル名: `{ホテル名}_口コミ分析改善レポート.docx`

### DOCX 7 章構成

| 章 | 内容 |
|---|---|
| 1 | エグゼクティブサマリー（KPIカード4つ + KEY FINDINGSボックス） |
| 2 | データ概要（サイト別テーブル + 評価分布テーブル + High/Mid/Low集計） |
| 3 | 強み分析（6テーマのテーブル + 最大の強みの詳細） |
| 4 | 弱み分析・優先度マトリクス（S/A/B/C テーブル） |
| 5 | 改善施策提案（Phase 1/2/3） |
| 6 | KPI目標設定（テーブル） |
| 7 | 総括（枠付きボックス） |

詳細は `references/docx_structure.md` を参照。

### カスタマイズポイント

- ホテル名、期間、件数（カバーページ + ヘッダー）
- KPI 値（全体平均、高評価率、低評価率、件数）
- サイト別データテーブル（10pt換算で降順ソート）
- 評価分布テーブル
- 強みテーマ（6件）+ 代表的コメント
- 弱み優先度マトリクス
- 改善施策（Phase 1/2/3）
- KPI目標テーブル
- 総括テキスト

### 実行

```bash
node create_{hotel_id}_report.js
```

---

## Step 4: PPTX プレゼンテーション生成

### テンプレートの読み込み

1. `assets/pptx_template.js` を読む（Read ツール）
2. テンプレート内の `// TODO: CUSTOMIZE` マーカーを確認
3. ホテル固有データで置き換えた新しい JS ファイルを作成:
   - ファイル名: `create_{hotel_id}_pptx.js`
   - 出力ファイル名: `{ホテル名}_口コミ分析レポート.pptx`

### PPTX 10 スライド構成

| Slide | 内容 |
|---|---|
| 1 | タイトル（ホテル名、期間、件数、作成日） |
| 2 | エグゼクティブサマリー（KPIカード4つ + Strengths/Weaknessesパネル） |
| 3 | サイト別評価分析（横棒グラフ + Insightボックス + データテーブル） |
| 4 | 評価分布（ドーナツチャート + 棒グラフ + サマリーカード） |
| 5 | 強み分析（3×2 の 6カードレイアウト） |
| 6 | 弱み分析・優先度マトリクス（テーブル） |
| 7 | 改善施策 Phase 1（2×2 の 4カードレイアウト） |
| 8 | 改善施策 Phase 2・3（左右パネル） |
| 9 | KPI目標設定（テーブル + 注記ボックス） |
| 10 | 総括（テキスト） |

詳細は `references/pptx_structure.md` を参照。

### 実行

```bash
node create_{hotel_id}_pptx.js
```

---

## Step 5: 検証

生成ファイルの検証を行う:

```python
# PPTX 検証
from pptx import Presentation
prs = Presentation("output.pptx")
assert len(prs.slides) == 10  # 10スライド

# DOCX 検証
import zipfile, xml.etree.ElementTree as ET
with zipfile.ZipFile("output.docx") as z:
    tree = ET.parse(z.open("word/document.xml"))
    # テキスト抽出してKPI値を確認

# ファイルフォーマット確認
# file output.docx → "Microsoft Word 2007+"
```

確認項目:
- 10pt換算の数値が正確か
- ホテル名が正しいか
- スライド数が10、DOCX章数が7か
- High/Mid/Low の件数と割合が合っているか

---

## 評価スケール変換ルール

全ホテル共通で以下のルールを適用する:

| サイト | 元の尺度 | 変換方法 |
|---|---|---|
| Booking.com | /10 | そのまま |
| Trip.com | /10 | そのまま |
| Agoda | /10 | そのまま |
| じゃらん | /5 | ×2 で10点換算 |
| 楽天トラベル | /5 | ×2 で10点換算 |
| Google | /5 | ×2 で10点換算 |
| 一休.com | /5 | ×2 で10点換算 |

### 判定基準（10pt換算後）

| 10pt換算 | 判定 | 色 |
|---|---|---|
| 9.0以上 | 優秀 | GREEN |
| 8.0以上 | 良好 | GREEN |
| 7.0以上 | 概ね良好 | ORANGE |
| 7.0未満 | 要改善 | RED |

### 評価カテゴリ

| カテゴリ | スコア範囲 | KPI色 |
|---|---|---|
| 高評価 (High) | 8-10点 | GREEN (80%以上なら) |
| 中評価 (Mid) | 5-7点 | ORANGE |
| 低評価 (Low) | 1-4点 | RED (0%ならGREEN、5%以上ならORANGE/RED) |

---

## デザインガイドライン

### DOCX カラーパレット

| 名前 | コード | 用途 |
|---|---|---|
| NAVY | `1B3A5C` | 見出し、ヘッダーセル |
| ACCENT | `2E75B6` | 小見出し、アクセント |
| GREEN | `27AE60` | ポジティブ値 |
| RED | `C0392B` | ネガティブ値、S優先度 |
| ORANGE | `E67E22` | 中間値、A優先度 |

### PPTX カラーパレット（Midnight Executive テーマ）

| 名前 | コード | 用途 |
|---|---|---|
| navy | `1A2744` | ヘッダー、背景 |
| blue | `3B7DD8` | アクセント、チャート |
| gold | `D4A843` | Insight ボックス、強調 |
| green | `16A34A` | ポジティブ値 |
| red | `DC2626` | ネガティブ値 |

### フォント
- 全体: Arial
- DOCX: 本文 21pt (10.5pt相当)、H1 32pt、H2 26pt
- PPTX: タイトル 42pt、見出し 22pt、本文 11pt、テーブル 9pt

---

## トラブルシューティング

### CSV ダウンロードが失敗する
- curl ではなく必ずブラウザ（認証済み）でアクセスする
- CSVエクスポートURLはファイルダウンロードをトリガーするので、`~/Downloads/` を確認する

### カラム構成が異なる
- `analyze_reviews.py` がヘッダー名ベースで自動検出する
- 見つからない場合はユーザーにカラムマッピングを確認する

### node_modules が見つからない
- 作業ディレクトリで `npm install docx pptxgenjs` を実行
