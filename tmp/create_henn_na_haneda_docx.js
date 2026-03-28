const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, PageNumber, PageBreak, LevelFormat
} = require("docx");

// ============================================================
// 変なホテル東京羽田 固有設定
// ============================================================
const HOTEL_NAME = "変なホテル東京羽田";
const ANALYSIS_PERIOD = "2026年2月〜3月";
const REPORT_DATE = "2026年3月28日";
const OUTPUT_DIR = "納品レポート/ホテル別レポート";

// レビュー基本情報
const REVIEW_COUNT = "338件（重複除外後）";
const TARGET_SITES = "対象サイト：Booking.com / Trip.com / じゃらん / 楽天トラベル / Agoda / Google";

// KPIカードの値
const KPI_CARDS = [
  { label: "全体平均(10pt換算)", value: "8.30", color: "27AE60", bgColor: "E8F5E9" },
  { label: "高評価率(8-10点)", value: "76.3%", color: "27AE60", bgColor: "E8F5E9" },
  { label: "低評価率(1-4点)", value: "3.3%", color: "E67E22", bgColor: "FEF9C3" },
  { label: "レビュー総数", value: "338件", color: "1B3A5C", bgColor: "D5E8F0" },
];

// エグゼクティブサマリー
const EXECUTIVE_SUMMARY_INTRO = "2026年2月〜3月に各OTAサイト・口コミサイトに投稿された338件のレビューを包括的に分析しました（本レポートは先頭50件をサンプルとした傾向分析を含みます）。以下が主要な発見事項です。";
const EXECUTIVE_SUMMARY_EVALUATION_1 = "変なホテル東京羽田は全体平均8.30点（10点満点）と良好な評価を獲得しています。338件という豊富な口コミ件数は他ホテル比で顧客接点の多さを示しており、Trip.comでは9.43点と特に高評価です。";
const EXECUTIVE_SUMMARY_EVALUATION_2 = "最大の特徴は「恐竜ロボットによるチェックイン」「無料空港シャトルバス」という独自のテクノロジー体験であり、国内外ゲストから高い評価を受けています。一方で低評価率3.3%（11件）があり、部屋の狭さや設備に関する課題も指摘されています。羽田空港至近という立地優位性と先進的なロボットテクノロジーの組み合わせが競合差別化の核心です。";

// KEY FINDINGS
const KEY_FINDING_STRENGTH = "テクノロジー活用（ロボットチェックイン・無料シャトルバス・セルフチェックアウト）が最大の差別化ポイントであり、客室・設備（20件）と清潔感（15件）も安定した評価を受けています。羽田空港直結の立地優位性（11件）との組み合わせが強力な競争優位性を形成しています。";
const KEY_FINDING_WEAKNESS = "部屋の狭さ（8件）、設備の老朽化（1件）、アメニティ不足（2件）が改善課題として挙がっています。低評価率3.3%（11件）はGoogle評価での低スコア集中が影響しており、日本語圏ゲストへの体験改善が重要です。";
const KEY_FINDING_OPPORTUNITY = "ロボット・テクノロジー体験という独自性は他ホテルにはない差別化ポイントです。空港アクセスの利便性（無料シャトルバス）と組み合わせたトランジット需要の取り込みを強化することで、インバウンド市場での更なる成長が期待されます。";

// サイト別レビューデータ
const SITE_DATA = [
  ["Trip.com", "23", "9.43", "/10", "9.43", "10.0", "優秀"],
  ["楽天トラベル", "42", "4.29", "/5", "8.57", "8.0", "良好"],
  ["じゃらん", "26", "4.19", "/5", "8.38", "8.0", "良好"],
  ["Agoda", "37", "8.14", "/10", "8.14", "8.0", "良好"],
  ["Booking.com", "145", "8.14", "/10", "8.14", "8.0", "良好"],
  ["Google", "65", "4.06", "/5", "8.12", "10.0", "良好"],
];

// 評価分布データ
const DISTRIBUTION_DATA = [
  [10, 117, "34.6%"],
  [9,  43,  "12.7%"],
  [8,  98,  "29.0%"],
  [7,  31,  "9.2%"],
  [6,  29,  "8.6%"],
  [5,  9,   "2.7%"],
  [4,  2,   "0.6%"],
  [3,  2,   "0.6%"],
  [2,  7,   "2.1%"],
];
const DISTRIBUTION_MAX_COUNT = 117;

// 評価分布サマリー
const HIGH_RATING_SUMMARY = "258件（76.3%）";
const MID_RATING_SUMMARY = "69件（20.4%）";
const LOW_RATING_SUMMARY = "11件（3.3%）";
const LOW_RATING_COLOR = "E67E22";

// データ概要テキスト
const DATA_OVERVIEW_SITE_TEXT = "各予約サイトの評価基準が異なる点に留意が必要です。海外OTA（Booking.com/Trip.com/Agoda）は10点満点、国内サイト（じゃらん/楽天トラベル/Google）は5点満点のため、統一比較のため10点換算（5点満点×2）を行いました。Booking.comが最多件数（145件）であり、インバウンドゲストからのレビューが分析の中核を形成しています。";
const DATA_OVERVIEW_DIST_TEXT = "評価分布はスコア10点が117件（34.6%）で最多、スコア8点が98件（29.0%）と続き、高評価層が全体の76.3%を占めています。スコア2点の低評価が7件（2.1%）あり、これらへの対応が平均スコア向上の鍵を握ります。";

// 強みテーマデータ
const STRENGTH_THEMES = [
  ["テクノロジー・ロボット体験", "20件以上", "「恐竜がチェックインの手続きをしてくれます」「セルフチェックインとチェックアウトもスムーズでした」"],
  ["客室・設備", "20件", "「部屋は小さいですが、必要なものはすべて揃っています。快適なベッド、ベッドサイドの照明、十分な広さのバスルーム」"],
  ["清潔感", "15件", "「とても清潔で快適で、スタッフも親切でした」「羽田空港との無料シャトルバスを提供。清潔で静か」"],
  ["スタッフ対応", "10件", "「スタッフも親切でした」「hotel staff were very helpful in allowing us to take a breakfast box with us」"],
  ["立地・空港アクセス", "11件", "「設備が素晴らしく、サービスも最高で、空港にもとても近い」「空港までのシャトルバスです」"],
  ["コスパ・リピート", "3件", "「2人で一泊14620円と安かったので予約しましたが、部屋がとても綺麗なのとチェックアウトが11時までなのがとてもありがたかった」"],
];

// 強みサブセクション
const STRENGTH_SUB_1_TITLE = "3.1 最大の強み：「テクノロジー活用・ロボット体験」";
const STRENGTH_SUB_1_TEXT = "変なホテル東京羽田の最大の差別化ポイントは「恐竜ロボットによるチェックイン体験」です。国内外ゲストが「A DINOSAUR CHECKS YOU IN」と驚きをもって言及するほどの独自体験は、他ホテルには真似できない競争優位性を形成しています。セルフチェックイン・アウトのスムーズさと組み合わせることで、テクノロジーを活用した効率的なサービスとして高評価を得ています。";
const STRENGTH_SUB_1_BULLETS = [
  "恐竜ロボットによるチェックイン体験が国内外で話題化（SNS訴求力大）",
  "セルフチェックイン・アウトのスムーズな運営（深夜・早朝フライトにも対応）",
  "無料空港シャトルバスの時間通りの運行（トランジット客に最適）",
  "漫画ライブラリの充実（最新漫画揃い・ユニークなゲスト体験の提供）",
];

const STRENGTH_SUB_2_TITLE = "3.2 羽田空港至近という立地優位性";
const STRENGTH_SUB_2_TEXT = "羽田空港からの無料シャトルバス運行は338件の口コミでも繰り返し言及されており、トランジット・早朝/深夜フライト利用者にとって代替不可能な価値を提供しています。この立地優位性とロボット体験の組み合わせが「変なホテル」ブランドの独自ポジショニングを強化しています。";

// 弱み分析データ
const WEAKNESS_PRIORITY_DATA = [
  ["S", "部屋の狭さ", "8件で指摘。2人利用時に狭さを感じるコメントが多い。コンパクト設計の明示的な事前周知が必要", "8件・高頻度言及"],
  ["A", "低評価件数（11件）の集中対策", "低評価率3.3%（11件）。Google評価での低スコアが含まれ、スコア2点が7件存在する", "11件・要個別対応"],
  ["A", "アメニティ不足", "コンプリメンタリーウォーター未提供・基本アメニティへの言及（2件）", "2件"],
  ["B", "設備の老朽化", "一部設備の古さへの指摘（汚いとの言及1件）", "1件"],
  ["C", "朝食選択肢", "朝食バラエティへの期待（概ね満足だが更なる充実を希望）", "3件（好意的言及含む）"],
];

// 改善施策 Phase 1
const PHASE1_DESCRIPTION = "投資不要・オペレーション改善で対応可能な施策";
const PHASE1_ITEMS = [
  {
    title: "(1) 部屋サイズの事前コミュニケーション強化",
    bullets: [
      "予約確認メールにコンパクトルームの写真・サイズを明記",
      "OTAページのルーム説明文に「コンパクト設計（〇〇㎡）」を明示",
      "2人利用の場合のおすすめ客室タイプを案内するポップアップ設置",
    ],
  },
  {
    title: "(2) ロボット・テクノロジー体験のさらなる訴求",
    bullets: [
      "恐竜ロボットチェックインの動画コンテンツを各OTAページに掲載",
      "SNSシェアを促すフォトスポット・ハッシュタグの整備",
      "ロボット体験をレビューに書いてもらうよう促す掲示物設置",
    ],
  },
  {
    title: "(3) 低評価への迅速な対応",
    bullets: [
      "全サイトの口コミへ48時間以内返信（特に低評価への真摯な対応）",
      "スコア2点の低評価コメントを個別分析し再発防止策を実施",
      "Google口コミへの返信強化でスコア底上げを図る",
    ],
  },
];

// 改善施策 Phase 2
const PHASE2_DESCRIPTION = "一定の投資を伴うが、比較的早期に実行可能な施策";
const PHASE2_ITEMS = [
  {
    title: "(1) アメニティサービスの見直し",
    bullets: [
      "コンプリメンタリーウォーターの提供開始（または自販機の充実）",
      "基本アメニティの選択式提供（事前にOTAで選択可能に）",
      "コスパ感を損なわない形でのアメニティグレードアップ",
    ],
  },
  {
    title: "(2) 朝食サービス強化",
    bullets: [
      "バラエティ朝食の充実化（インバウンドを意識した多様なメニュー追加）",
      "早朝フライト向け朝食ボックス（テイクアウト）サービスの標準化",
    ],
  },
];

// 改善施策 Phase 3
const PHASE3_DESCRIPTION = "設備投資を伴う抜本的な改善施策";
const PHASE3_ITEMS = [
  {
    title: "(1) テクノロジーのさらなる進化",
    bullets: [
      "ロボットフロントの機能拡充（多言語対応・観光案内機能追加）",
      "スマートルーム化（音声コントロール・AIコンシェルジュ導入）",
      "客室IoT設備のアップデート（ベッドサイド操作パネルの刷新）",
    ],
  },
  {
    title: "(2) 客室リノベーション",
    bullets: [
      "狭さを感じさせない空間デザインへのリノベーション（収納最適化）",
      "一部客室の設備更新（古さ指摘部分の重点的改善）",
    ],
  },
];

// KPI目標データ
const KPI_TARGET_DATA = [
  ["全体平均(10pt換算)", "8.30点", "8.5点以上", "2026年9月"],
  ["高評価率（8-10点）", "76.3%", "80%以上", "2026年9月"],
  ["低評価率（1-4点）", "3.3%", "2%以下", "2026年9月"],
  ["Trip.com平均評価", "9.43/10点", "9.5点以上", "2026年9月"],
  ["Google平均評価", "4.06/5点", "4.2/5点以上", "2026年9月"],
  ["口コミ返信率", "未計測", "100%（48h以内）", "2026年5月"],
  ["低評価クレーム件数", "11件/2ヶ月", "5件以下/2ヶ月", "2026年9月"],
];

// 総括テキスト
const CONCLUSION_PARAGRAPHS = [
  "変なホテル東京羽田は、2026年2月〜3月において全体平均8.30点・338件という豊富な口コミ数を誇ります。「恐竜ロボットによるチェックイン」「無料空港シャトルバス」という先進的なテクノロジー活用が最大の差別化ポイントであり、国内外ゲストから高い評価と話題性を獲得しています。",
  "Trip.com（9.43点）での高評価に見られるように、インバウンド市場での競争力は非常に強固です。セルフチェックイン・深夜対応・早朝フライト向け朝食ボックス提供など、トランジット需要に最適化されたサービス体制が評価の柱となっています。",
  "改善課題として最も言及される部屋の狭さ（8件）については、事前コミュニケーションの強化（OTAページでの明示）により期待値マネジメントを行うことが最も効果的です。低評価率3.3%（11件）の低減には、低評価口コミへの個別対応と設備老朽化の解消が不可欠です。",
  "ロボット・テクノロジーの先進性をさらに進化させることで、「変なホテル」ブランドの独自性を維持・強化することが長期的な競争優位性確保の鍵です。スマートルーム化・AI コンシェルジュ導入など次世代テクノロジーへの投資が推奨されます。",
];
const CONCLUSION_FINAL_PARAGRAPH = "テクノロジーと羽田空港立地を核とした独自ポジショニングを強化しながら、全体平均8.5点以上・高評価率80%以上・低評価率2%以下の達成を目指してください。338件の豊富な口コミは宝の山であり、ゲストの声を活かした継続的な改善が成長の原動力です。";

// ============================================================
// Color scheme (DO NOT MODIFY)
// ============================================================
const NAVY = "1B3A5C";
const ACCENT = "2E75B6";
const LIGHT_BLUE = "D5E8F0";
const LIGHT_GRAY = "F2F2F2";
const WHITE = "FFFFFF";
const RED_ACCENT = "C0392B";
const GREEN_ACCENT = "27AE60";
const ORANGE_ACCENT = "E67E22";

const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };
const noBorders = {
  top: { style: BorderStyle.NONE, size: 0 },
  bottom: { style: BorderStyle.NONE, size: 0 },
  left: { style: BorderStyle.NONE, size: 0 },
  right: { style: BorderStyle.NONE, size: 0 },
};
const cellMargins = { top: 80, bottom: 80, left: 120, right: 120 };

// ============================================================
// Helper functions (DO NOT MODIFY)
// ============================================================
function heading1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 200 },
    children: [new TextRun({ text, bold: true, size: 32, font: "Arial", color: NAVY })],
  });
}

function heading2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 280, after: 160 },
    children: [new TextRun({ text, bold: true, size: 26, font: "Arial", color: ACCENT })],
  });
}

function heading3(text) {
  return new Paragraph({
    spacing: { before: 200, after: 120 },
    children: [new TextRun({ text, bold: true, size: 22, font: "Arial", color: NAVY })],
  });
}

function para(text, opts = {}) {
  return new Paragraph({
    spacing: { after: opts.afterSpacing || 120, line: 320 },
    alignment: opts.alignment || AlignmentType.LEFT,
    children: [new TextRun({ text, size: 21, font: "Arial", color: opts.color || "333333", ...opts })],
  });
}

function multiRunPara(runs, opts = {}) {
  return new Paragraph({
    spacing: { after: opts.afterSpacing || 120, line: 320 },
    alignment: opts.alignment || AlignmentType.LEFT,
    children: runs.map(r => new TextRun({ size: 21, font: "Arial", color: "333333", ...r })),
  });
}

function bulletItem(text, opts = {}) {
  return new Paragraph({
    numbering: { reference: "bullets", level: opts.level || 0 },
    spacing: { after: 80, line: 300 },
    children: [new TextRun({ text, size: 21, font: "Arial", color: "333333" })],
  });
}

function headerCell(text, width) {
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    shading: { fill: NAVY, type: ShadingType.CLEAR },
    margins: cellMargins,
    verticalAlign: "center",
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text, bold: true, size: 20, font: "Arial", color: WHITE })],
    })],
  });
}

function dataCell(text, width, opts = {}) {
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    shading: opts.fill ? { fill: opts.fill, type: ShadingType.CLEAR } : undefined,
    margins: cellMargins,
    verticalAlign: "center",
    children: [new Paragraph({
      alignment: opts.alignment || AlignmentType.LEFT,
      children: [new TextRun({ text: String(text), size: 20, font: "Arial", color: opts.color || "333333", bold: opts.bold || false })],
    })],
  });
}

function spacer(height = 100) {
  return new Paragraph({ spacing: { after: height }, children: [] });
}

function divider() {
  return new Paragraph({
    spacing: { before: 200, after: 200 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: ACCENT, space: 1 } },
    children: [],
  });
}

// KPI Box helper
function kpiRow(items) {
  const colWidth = Math.floor(9360 / items.length);
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: items.map(() => colWidth),
    rows: [
      new TableRow({
        children: items.map(item =>
          new TableCell({
            borders: { ...noBorders, right: { style: BorderStyle.SINGLE, size: 1, color: "DDDDDD" } },
            width: { size: colWidth, type: WidthType.DXA },
            shading: { fill: item.bgColor || LIGHT_BLUE, type: ShadingType.CLEAR },
            margins: { top: 160, bottom: 160, left: 200, right: 200 },
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 40 },
                children: [new TextRun({ text: item.label, size: 18, font: "Arial", color: "666666" })],
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [new TextRun({ text: item.value, bold: true, size: 36, font: "Arial", color: item.color || NAVY })],
              }),
            ],
          })
        ),
      }),
    ],
  });
}

// Priority matrix helper
function priorityRow(priority, category, detail, impact) {
  const colors = { "S": RED_ACCENT, "A": ORANGE_ACCENT, "B": ACCENT, "C": "888888" };
  return new TableRow({
    children: [
      new TableCell({
        borders,
        width: { size: 800, type: WidthType.DXA },
        margins: cellMargins,
        shading: { fill: LIGHT_GRAY, type: ShadingType.CLEAR },
        verticalAlign: "center",
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: priority, bold: true, size: 22, font: "Arial", color: colors[priority] || NAVY })],
        })],
      }),
      dataCell(category, 2200, { bold: true }),
      dataCell(detail, 4360),
      dataCell(impact, 2000, { alignment: AlignmentType.CENTER }),
    ],
  });
}

// ============================================================
// Build improvement phase sections
// ============================================================
function buildPhaseItems(items) {
  const result = [];
  items.forEach((item, index) => {
    result.push(heading3(item.title));
    item.bullets.forEach(b => result.push(bulletItem(b)));
    if (index < items.length - 1) result.push(spacer(80));
  });
  return result;
}

// ============================================================
// Build site table rows from data array
// ============================================================
function buildSiteTableRows(data) {
  return data.map((row, index) => {
    const [siteName, count, nativeAvg, scale, tenPt, median, verdict] = row;
    const fill = index % 2 === 1 ? LIGHT_GRAY : undefined;
    const verdictColor = verdict === "優秀" ? GREEN_ACCENT : (verdict === "良好" ? GREEN_ACCENT : ORANGE_ACCENT);
    return new TableRow({ children: [
      dataCell(siteName, 1800, { bold: true, fill }),
      dataCell(count, 900, { alignment: AlignmentType.CENTER, fill }),
      dataCell(nativeAvg, 1300, { alignment: AlignmentType.CENTER, bold: true, fill }),
      dataCell(scale, 900, { alignment: AlignmentType.CENTER, fill }),
      dataCell(tenPt, 1300, { alignment: AlignmentType.CENTER, color: verdictColor, bold: true, fill }),
      dataCell(median, 1526, { alignment: AlignmentType.CENTER, fill }),
      dataCell(verdict, 1300, { alignment: AlignmentType.CENTER, color: verdictColor, bold: true, fill }),
    ]});
  });
}

// ============================================================
// Build distribution table rows from data array
// ============================================================
function buildDistributionRows(data, maxCount) {
  return data.map(([rating, count, pct]) => {
    const barWidth = Math.round(count / maxCount * 100);
    const barColor = rating >= 8 ? GREEN_ACCENT : (rating >= 5 ? ORANGE_ACCENT : RED_ACCENT);
    const fill = [10, 8, 6, 4, 2].includes(rating) ? LIGHT_GRAY : WHITE;
    return new TableRow({ children: [
      dataCell(String(rating), 1200, { alignment: AlignmentType.CENTER, bold: true, fill }),
      dataCell(String(count), 1200, { alignment: AlignmentType.CENTER, fill }),
      dataCell(pct, 1500, { alignment: AlignmentType.CENTER, fill }),
      new TableCell({
        borders, width: { size: 5126, type: WidthType.DXA }, margins: cellMargins,
        shading: fill !== WHITE ? { fill, type: ShadingType.CLEAR } : undefined,
        children: [new Paragraph({ children: [
          new TextRun({ text: "\u2588".repeat(Math.max(1, Math.round(barWidth / 5))), size: 20, font: "Arial", color: barColor }),
          new TextRun({ text: ` ${count}件`, size: 18, font: "Arial", color: "666666" }),
        ]})],
      }),
    ]});
  });
}

// ============================================================
// Build strength theme rows from data array
// ============================================================
function buildStrengthRows(data) {
  return data.map((row, index) => {
    const [theme, mentionCount, comments] = row;
    const fill = index % 2 === 0 ? "E8F5E9" : undefined;
    return new TableRow({ children: [
      dataCell(theme, 2600, { bold: true, fill: fill || undefined }),
      dataCell(mentionCount, 1200, { alignment: AlignmentType.CENTER, bold: true, color: GREEN_ACCENT, fill: fill || undefined }),
      dataCell(comments, 5226, { fill: fill || undefined }),
    ]});
  });
}

// ============================================================
// Build KPI target rows from data array
// ============================================================
function buildKPITargetRows(data) {
  return data.map((row, index) => {
    const [item, current, target, deadline] = row;
    const fill = index % 2 === 1 ? LIGHT_GRAY : undefined;
    return new TableRow({ children: [
      dataCell(item, 2800, { bold: true, fill }),
      dataCell(current, 2200, { alignment: AlignmentType.CENTER, fill }),
      dataCell(target, 2200, { alignment: AlignmentType.CENTER, color: GREEN_ACCENT, bold: true, fill }),
      dataCell(deadline, 1826, { alignment: AlignmentType.CENTER, fill }),
    ]});
  });
}

// ============================================================
// Build the document
// ============================================================
const doc = new Document({
  styles: {
    default: { document: { run: { font: "Arial", size: 21 } } },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: "Arial", color: NAVY },
        paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 0 },
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Arial", color: ACCENT },
        paragraph: { spacing: { before: 280, after: 160 }, outlineLevel: 1 },
      },
    ],
  },
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [
          { level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
          { level: 1, format: LevelFormat.BULLET, text: "\u25E6", alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 1440, hanging: 360 } } } },
        ],
      },
      {
        reference: "numbers",
        levels: [
          { level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
        ],
      },
    ],
  },
  sections: [
    // ===== COVER PAGE =====
    {
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
        },
      },
      children: [
        spacer(2000),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 80 },
          border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: ACCENT, space: 8 } },
          children: [new TextRun({ text: "口コミ分析", size: 56, font: "Arial", color: ACCENT })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 200 },
          children: [new TextRun({ text: "改善レポート", size: 56, font: "Arial", bold: true, color: NAVY })],
        }),
        spacer(400),
        // TODO: CUSTOMIZE - ホテル名
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 120 },
          children: [new TextRun({ text: HOTEL_NAME, size: 32, font: "Arial", color: NAVY, bold: true })],
        }),
        spacer(200),
        // TODO: CUSTOMIZE - 分析期間
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 80 },
          children: [new TextRun({ text: `分析対象期間：${ANALYSIS_PERIOD}`, size: 22, font: "Arial", color: "666666" })],
        }),
        // TODO: CUSTOMIZE - レビュー総数
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 80 },
          children: [new TextRun({ text: `レビュー総数：${REVIEW_COUNT}`, size: 22, font: "Arial", color: "666666" })],
        }),
        // TODO: CUSTOMIZE - 対象サイト
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 80 },
          children: [new TextRun({ text: TARGET_SITES, size: 20, font: "Arial", color: "666666" })],
        }),
        spacer(1600),
        // TODO: CUSTOMIZE - 作成日
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 60 },
          children: [new TextRun({ text: `作成日：${REPORT_DATE}`, size: 20, font: "Arial", color: "999999" })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "Confidential - For Internal Use Only", size: 18, font: "Arial", color: "AAAAAA", italics: true })],
        }),
      ],
    },

    // ===== MAIN CONTENT =====
    {
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
        },
      },
      headers: {
        default: new Header({
          children: [new Paragraph({
            alignment: AlignmentType.RIGHT,
            border: { bottom: { style: BorderStyle.SINGLE, size: 2, color: ACCENT, space: 4 } },
            // TODO: CUSTOMIZE - ヘッダーテキスト
            children: [new TextRun({ text: `${HOTEL_NAME}｜口コミ分析改善レポート`, size: 16, font: "Arial", color: "999999" })],
          })],
        }),
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: "Page ", size: 16, font: "Arial", color: "999999" }),
              new TextRun({ children: [PageNumber.CURRENT], size: 16, font: "Arial", color: "999999" })],
          })],
        }),
      },
      children: [
        // ===== 1. EXECUTIVE SUMMARY =====
        heading1("1. エグゼクティブサマリー"),
        // TODO: CUSTOMIZE - サマリーイントロ
        para(EXECUTIVE_SUMMARY_INTRO),
        spacer(100),

        // TODO: CUSTOMIZE - KPIカード
        kpiRow(KPI_CARDS),
        spacer(200),

        heading2("総合評価"),
        // TODO: CUSTOMIZE - 総合評価テキスト
        para(EXECUTIVE_SUMMARY_EVALUATION_1),
        para(EXECUTIVE_SUMMARY_EVALUATION_2),
        spacer(100),

        // Key findings box
        // TODO: CUSTOMIZE - KEY FINDINGS テキスト
        new Table({
          width: { size: 9026, type: WidthType.DXA },
          columnWidths: [9026],
          rows: [new TableRow({
            children: [new TableCell({
              borders: { top: { style: BorderStyle.SINGLE, size: 4, color: ACCENT }, bottom: border, left: { style: BorderStyle.SINGLE, size: 4, color: ACCENT }, right: border },
              width: { size: 9026, type: WidthType.DXA },
              shading: { fill: "F0F7FC", type: ShadingType.CLEAR },
              margins: { top: 200, bottom: 200, left: 300, right: 300 },
              children: [
                new Paragraph({ spacing: { after: 120 }, children: [new TextRun({ text: "KEY FINDINGS", bold: true, size: 22, font: "Arial", color: ACCENT })] }),
                new Paragraph({ spacing: { after: 80 }, children: [
                  new TextRun({ text: "Strength：", bold: true, size: 20, font: "Arial", color: GREEN_ACCENT }),
                  new TextRun({ text: KEY_FINDING_STRENGTH, size: 20, font: "Arial", color: "333333" }),
                ]}),
                new Paragraph({ spacing: { after: 80 }, children: [
                  new TextRun({ text: "Weakness：", bold: true, size: 20, font: "Arial", color: RED_ACCENT }),
                  new TextRun({ text: KEY_FINDING_WEAKNESS, size: 20, font: "Arial", color: "333333" }),
                ]}),
                new Paragraph({ children: [
                  new TextRun({ text: "Opportunity：", bold: true, size: 20, font: "Arial", color: ORANGE_ACCENT }),
                  new TextRun({ text: KEY_FINDING_OPPORTUNITY, size: 20, font: "Arial", color: "333333" }),
                ]}),
              ],
            })],
          })],
        }),

        new Paragraph({ children: [new PageBreak()] }),

        // ===== 2. DATA OVERVIEW =====
        heading1("2. データ概要"),

        heading2("2.1 サイト別レビュー件数・評価"),
        // TODO: CUSTOMIZE - サイト別テキスト
        para(DATA_OVERVIEW_SITE_TEXT),
        spacer(80),

        // TODO: CUSTOMIZE - サイト別テーブル（SITE_DATA配列を編集）
        new Table({
          width: { size: 9026, type: WidthType.DXA },
          columnWidths: [1800, 900, 1300, 900, 1300, 1526, 1300],
          rows: [
            new TableRow({
              children: [
                headerCell("サイト名", 1800),
                headerCell("件数", 900),
                headerCell("ネイティブ平均", 1300),
                headerCell("尺度", 900),
                headerCell("10pt換算", 1300),
                headerCell("中央値(10pt)", 1526),
                headerCell("判定", 1300),
              ],
            }),
            ...buildSiteTableRows(SITE_DATA),
          ],
        }),
        spacer(200),

        heading2("2.2 評価分布（10点換算）"),
        // TODO: CUSTOMIZE - 評価分布テキスト
        para(DATA_OVERVIEW_DIST_TEXT),
        spacer(80),

        // TODO: CUSTOMIZE - 評価分布テーブル（DISTRIBUTION_DATA配列を編集）
        new Table({
          width: { size: 9026, type: WidthType.DXA },
          columnWidths: [1200, 1200, 1500, 5126],
          rows: [
            new TableRow({ children: [headerCell("評価", 1200), headerCell("件数", 1200), headerCell("割合", 1500), headerCell("分布", 5126)] }),
            ...buildDistributionRows(DISTRIBUTION_DATA, DISTRIBUTION_MAX_COUNT),
          ],
        }),
        spacer(100),

        // Rating category summary
        // TODO: CUSTOMIZE - 高/中/低評価サマリー
        new Table({
          width: { size: 9026, type: WidthType.DXA },
          columnWidths: [3008, 3009, 3009],
          rows: [new TableRow({
            children: [
              new TableCell({
                borders, width: { size: 3008, type: WidthType.DXA },
                shading: { fill: "E8F5E9", type: ShadingType.CLEAR },
                margins: { top: 120, bottom: 120, left: 200, right: 200 },
                children: [
                  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [new TextRun({ text: "高評価（8-10点）", bold: true, size: 20, font: "Arial", color: GREEN_ACCENT })] }),
                  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: HIGH_RATING_SUMMARY, bold: true, size: 24, font: "Arial", color: GREEN_ACCENT })] }),
                ],
              }),
              new TableCell({
                borders, width: { size: 3009, type: WidthType.DXA },
                shading: { fill: "FFF3E0", type: ShadingType.CLEAR },
                margins: { top: 120, bottom: 120, left: 200, right: 200 },
                children: [
                  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [new TextRun({ text: "中評価（5-7点）", bold: true, size: 20, font: "Arial", color: ORANGE_ACCENT })] }),
                  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: MID_RATING_SUMMARY, bold: true, size: 24, font: "Arial", color: ORANGE_ACCENT })] }),
                ],
              }),
              new TableCell({
                borders, width: { size: 3009, type: WidthType.DXA },
                shading: { fill: "FDEDEC", type: ShadingType.CLEAR },
                margins: { top: 120, bottom: 120, left: 200, right: 200 },
                children: [
                  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [new TextRun({ text: "低評価（1-4点）", bold: true, size: 20, font: "Arial", color: LOW_RATING_COLOR })] }),
                  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: LOW_RATING_SUMMARY, bold: true, size: 24, font: "Arial", color: LOW_RATING_COLOR })] }),
                ],
              }),
            ],
          })],
        }),

        new Paragraph({ children: [new PageBreak()] }),

        // ===== 3. STRENGTH ANALYSIS =====
        heading1("3. 強み分析（ポジティブ要因）"),
        para("口コミのテキストマイニングにより、以下のポジティブテーマが特定されました。"),
        spacer(100),

        // TODO: CUSTOMIZE - 強みテーマテーブル（STRENGTH_THEMES配列を編集）
        new Table({
          width: { size: 9026, type: WidthType.DXA },
          columnWidths: [2600, 1200, 5226],
          rows: [
            new TableRow({ children: [headerCell("ポジティブテーマ", 2600), headerCell("言及数", 1200), headerCell("代表的なコメント", 5226)] }),
            ...buildStrengthRows(STRENGTH_THEMES),
          ],
        }),
        spacer(200),

        // TODO: CUSTOMIZE - 強みサブセクション
        heading2(STRENGTH_SUB_1_TITLE),
        para(STRENGTH_SUB_1_TEXT),
        ...STRENGTH_SUB_1_BULLETS.map(b => bulletItem(b)),
        spacer(100),

        heading2(STRENGTH_SUB_2_TITLE),
        para(STRENGTH_SUB_2_TEXT),

        new Paragraph({ children: [new PageBreak()] }),

        // ===== 4. WEAKNESS ANALYSIS =====
        heading1("4. 弱み分析（改善課題）"),
        para("ネガティブコメントの分析から、以下の改善課題が抽出されました。影響度と頻度に基づく優先度をS〜Cで設定しています。"),
        spacer(100),

        // TODO: CUSTOMIZE - 弱みテーブル（WEAKNESS_PRIORITY_DATA配列を編集）
        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [800, 2200, 4360, 2000],
          rows: [
            new TableRow({ children: [headerCell("優先度", 800), headerCell("課題カテゴリ", 2200), headerCell("具体的内容", 4360), headerCell("影響度", 2000)] }),
            ...WEAKNESS_PRIORITY_DATA.map(([p, c, d, i]) => priorityRow(p, c, d, i)),
          ],
        }),

        new Paragraph({ children: [new PageBreak()] }),

        // ===== 5. IMPROVEMENT PLAN =====
        heading1("5. 改善施策提案"),
        para("分析結果に基づき、以下の改善施策を「即座対応」「短期」「中期」の3フェーズに分けて提案いたします。"),
        spacer(100),

        // Phase 1
        // TODO: CUSTOMIZE - Phase 1 施策（PHASE1_ITEMS配列を編集）
        heading2("Phase 1：即座対応（今週〜1ヶ月以内）"),
        para(PHASE1_DESCRIPTION),
        spacer(80),
        ...buildPhaseItems(PHASE1_ITEMS),

        new Paragraph({ children: [new PageBreak()] }),

        // Phase 2
        // TODO: CUSTOMIZE - Phase 2 施策（PHASE2_ITEMS配列を編集）
        heading2("Phase 2：短期施策（1〜3ヶ月）"),
        para(PHASE2_DESCRIPTION),
        spacer(80),
        ...buildPhaseItems(PHASE2_ITEMS),

        spacer(200),

        // Phase 3
        // TODO: CUSTOMIZE - Phase 3 施策（PHASE3_ITEMS配列を編集）
        heading2("Phase 3：中期施策（3〜6ヶ月）"),
        para(PHASE3_DESCRIPTION),
        spacer(80),
        ...buildPhaseItems(PHASE3_ITEMS),

        new Paragraph({ children: [new PageBreak()] }),

        // ===== 6. KPI & TARGETS =====
        heading1("6. KPI目標設定"),
        para("以下のKPIを設定し、四半期ごとにモニタリングすることを推奨します。"),
        spacer(100),

        // TODO: CUSTOMIZE - KPI目標テーブル（KPI_TARGET_DATA配列を編集）
        new Table({
          width: { size: 9026, type: WidthType.DXA },
          columnWidths: [2800, 2200, 2200, 1826],
          rows: [
            new TableRow({ children: [headerCell("KPI項目", 2800), headerCell("現状値", 2200), headerCell("目標値（6ヶ月後）", 2200), headerCell("期限", 1826)] }),
            ...buildKPITargetRows(KPI_TARGET_DATA),
          ],
        }),

        new Paragraph({ children: [new PageBreak()] }),

        // ===== 7. CONCLUSION =====
        heading1("7. 総括と今後のアクション"),
        spacer(100),

        // TODO: CUSTOMIZE - 総括テキスト（CONCLUSION_PARAGRAPHS配列を編集）
        new Table({
          width: { size: 9026, type: WidthType.DXA },
          columnWidths: [9026],
          rows: [new TableRow({
            children: [new TableCell({
              borders: { top: { style: BorderStyle.SINGLE, size: 4, color: NAVY }, bottom: { style: BorderStyle.SINGLE, size: 4, color: NAVY }, left: { style: BorderStyle.SINGLE, size: 4, color: NAVY }, right: { style: BorderStyle.SINGLE, size: 4, color: NAVY } },
              width: { size: 9026, type: WidthType.DXA },
              shading: { fill: "F8F9FA", type: ShadingType.CLEAR },
              margins: { top: 300, bottom: 300, left: 400, right: 400 },
              children: [
                ...CONCLUSION_PARAGRAPHS.map(text =>
                  new Paragraph({ spacing: { after: 160 }, children: [new TextRun({ text, size: 21, font: "Arial", color: "333333" })] })
                ),
                new Paragraph({ children: [new TextRun({ text: CONCLUSION_FINAL_PARAGRAPH, size: 21, font: "Arial", color: NAVY, bold: true })] }),
              ],
            })],
          })],
        }),

        spacer(300),
        divider(),
        spacer(100),
        para("本レポートに関するご質問・ご相談がございましたら、お気軽にお問い合わせください。", { alignment: AlignmentType.CENTER, color: "999999" }),
      ],
    },
  ],
});

// TODO: CUSTOMIZE - 出力ファイルパス（HOTEL_NAME, OUTPUT_DIR変数を編集）
Packer.toBuffer(doc).then(buffer => {
  const outputPath = `${OUTPUT_DIR}/${HOTEL_NAME}_口コミ分析改善レポート.docx`;
  fs.writeFileSync(outputPath, buffer);
  console.log("Report created successfully!");
  console.log("Output: " + outputPath);
  console.log("File size: " + (buffer.length / 1024).toFixed(1) + " KB");
});
