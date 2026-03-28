const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, PageNumber, PageBreak, LevelFormat
} = require("docx");

// ============================================================
// ダイワロイネットホテル大崎 固有設定
// ============================================================
const HOTEL_NAME = "ダイワロイネットホテル大崎";
const ANALYSIS_PERIOD = "2026年2月〜3月";
const REPORT_DATE = "2026年3月28日";
const OUTPUT_DIR = "納品レポート/ホテル別レポート";

// レビュー基本情報
const REVIEW_COUNT = "60件（重複除外後）";
const TARGET_SITES = "対象サイト：Booking.com / Trip.com / じゃらん / 楽天トラベル / Agoda / Google";

// KPIカードの値
const KPI_CARDS = [
  { label: "全体平均(10pt換算)", value: "8.87", color: "27AE60", bgColor: "E8F5E9" },
  { label: "高評価率(8-10点)", value: "81.7%", color: "27AE60", bgColor: "E8F5E9" },
  { label: "低評価率(1-4点)", value: "0.0%", color: "27AE60", bgColor: "E8F5E9" },
  { label: "レビュー総数", value: "60件", color: "1B3A5C", bgColor: "D5E8F0" },
];

// エグゼクティブサマリー
const EXECUTIVE_SUMMARY_INTRO = "2026年2月〜3月に各OTAサイト・口コミサイトに投稿された60件のレビューを包括的に分析しました。以下が主要な発見事項です。";
const EXECUTIVE_SUMMARY_EVALUATION_1 = "ダイワロイネットホテル大崎は全体平均8.87点（10点満点）と非常に高い評価を獲得しています。特筆すべきは低評価（1-4点）が0件という圧倒的な安定感であり、全60件中49件（81.7%）が高評価（8-10点）を占めています。";
const EXECUTIVE_SUMMARY_EVALUATION_2 = "立地・アクセスの優秀さ（29件言及）が最大の強みであり、大崎駅直結という好立地が国内外ゲストから高く評価されています。清潔感（16件）や客室・設備（19件）も安定した評価を受けており、ブランドとしての品質が高いレベルで維持されています。";

// KEY FINDINGS
const KEY_FINDING_STRENGTH = "立地・アクセスの卓越性（29件）、客室・設備の充実（19件）、清潔感の高評価（16件）が三大強みとして際立っています。低評価0件は全サイトを通じた圧倒的な顧客満足度の証左です。";
const KEY_FINDING_WEAKNESS = "アメニティの充実度不足（3件：セーフティボックス未設置・コットンボール未提供）、エレベーター関連の言及（2件）、マットレスの劣化懸念（1件）が改善余地として挙がっています。";
const KEY_FINDING_OPPORTUNITY = "低評価0件・高評価率81.7%の強みをさらに伸ばし、Trip.com平均9.67点を全サイトで実現することで、プレミアムブランドとしての地位確立が可能です。インバウンド客からの高評価を活かした海外マーケティング強化も有効です。";

// サイト別レビューデータ
const SITE_DATA = [
  ["Trip.com", "12", "9.67", "/10", "9.67", "10.0", "優秀"],
  ["Agoda", "5", "9.00", "/10", "9.00", "10.0", "優秀"],
  ["Booking.com", "26", "8.96", "/10", "8.96", "9.0", "良好"],
  ["楽天トラベル", "3", "4.33", "/5", "8.67", "8.0", "良好"],
  ["Google", "6", "4.00", "/5", "8.00", "8.0", "良好"],
  ["じゃらん", "8", "4.00", "/5", "8.00", "8.0", "良好"],
];

// 評価分布データ
const DISTRIBUTION_DATA = [
  [10, 30, "50.0%"],
  [9,  8,  "13.3%"],
  [8,  11, "18.3%"],
  [7,  6,  "10.0%"],
  [6,  5,  "8.3%"],
];
const DISTRIBUTION_MAX_COUNT = 30;

// 評価分布サマリー
const HIGH_RATING_SUMMARY = "49件（81.7%）";
const MID_RATING_SUMMARY = "11件（18.3%）";
const LOW_RATING_SUMMARY = "0件（0.0%）";
const LOW_RATING_COLOR = "27AE60";

// データ概要テキスト
const DATA_OVERVIEW_SITE_TEXT = "各予約サイトの評価基準が異なる点に留意が必要です。海外OTA（Booking.com/Trip.com/Agoda）は10点満点、国内サイト（じゃらん/楽天トラベル/Google）は5点満点のため、統一比較のため10点換算（5点満点×2）を行いました。";
const DATA_OVERVIEW_DIST_TEXT = "評価分布はスコア10点が30件（50.0%）と突出して高く、半数のゲストが最高評価を付与しています。スコア8点以上の合計が49件（81.7%）を占め、全体的に高水準の評価分布です。低評価（6点以下）は11件（18.3%）にとどまり、低評価0件という点は特筆に値します。";

// 強みテーマデータ
const STRENGTH_THEMES = [
  ["立地・アクセス", "29件", "「大崎駅直ぐという立地やアクセスしやすいのが便利です」「駅が近くて駐車場も隣にあり良かった」"],
  ["客室・設備", "19件", "「アメニティや設備も充実しており、コンセント、照明設備、ドライヤー、冷暖房も使いやすくとても快適に過ごせました」"],
  ["清潔感", "16件", "「とても清潔」「清潔感のある過ごしやすい環境を提供しているホテルでした」"],
  ["スタッフ対応", "10件", "「チェックイン時に大荷物だったこともあり不安だったのですが、お手伝いしていただき本当に助かりました」"],
  ["朝食", "6件", "「朝食もバラエティに富んだビュッフェで、ホテル料金高騰の中で都心にも関わらず理想的な値段で泊まれた」"],
  ["コスパ・リピート", "1件", "「Amazing value for money!」「このホテル、すごく気に入りました！値段も手頃だし、立地も便利。絶対にまた泊まりたいです！」"],
];

// 強みサブセクション
const STRENGTH_SUB_1_TITLE = "3.1 最大の強み：「立地・アクセス」";
const STRENGTH_SUB_1_TEXT = "大崎駅直結という立地は全60件中29件（48.3%）で言及されており、最大の強みです。国内・海外ゲストともに駅近の利便性を高く評価しており、ビジネス・観光いずれの目的にも対応できる好立地が競合優位性の根幹を形成しています。";
const STRENGTH_SUB_1_BULLETS = [
  "大崎駅から徒歩3分圏内という圧倒的なアクセス優位性",
  "品川・渋谷・新宿など都内主要拠点への高い利便性",
  "インバウンド客からも「非常に良い立地」と多数言及",
  "隣接駐車場の存在により車利用のゲストにも対応",
];

const STRENGTH_SUB_2_TITLE = "3.2 低評価0件の品質安定性";
const STRENGTH_SUB_2_TEXT = "全サイト通じて低評価（1-4点）が0件という事実は、ダイワロイネットホテルブランドとしての品質基準が高水準で維持されていることを示しています。特にTrip.comでの平均9.67点はインバウンド市場での高評価を示しており、海外発信力の強化に繋がっています。";

// 弱み分析データ
const WEAKNESS_PRIORITY_DATA = [
  ["A", "アメニティ不足", "セーフティボックス未設置・コットンボール等基本アメニティの欠如", "3件・複数言及あり"],
  ["A", "設備老朽化（マットレス）", "一部マットレスの沈み込みによる腰痛クレーム（長期滞在に影響）", "1件・評価7点の原因"],
  ["B", "エレベーター関連", "ホテル内のエレベーター・アメニティ関連の言及", "2件"],
  ["B", "部屋の狭さ", "一部客室の狭さについての指摘", "1件"],
  ["B", "水回り品質", "シャワー水圧等への言及（概ね良好だが要継続確認）", "1件"],
  ["C", "コスパ認知", "高評価率に比してコスパ言及が少なく、価値認知の向上余地あり", "1件"],
];

// 改善施策 Phase 1
const PHASE1_DESCRIPTION = "投資不要・オペレーション改善で対応可能な施策";
const PHASE1_ITEMS = [
  {
    title: "(1) アメニティ拡充（即座対応）",
    bullets: [
      "セーフティボックスの全室設置を検討（フロント預かりの手間を解消）",
      "コットンボール・綿棒等基本アメニティをアメニティバスケットに追加",
      "アメニティリストを客室に掲示しフロントへのリクエスト方法を周知",
    ],
  },
  {
    title: "(2) 口コミ返信の強化",
    bullets: [
      "全サイトの口コミへ48時間以内に返信（感謝 + 課題への回答）",
      "高評価レビューへの返信でリピート意向を強化",
      "低評価の懸念に対する真摯な回答で潜在顧客への信頼醸成",
    ],
  },
  {
    title: "(3) インバウンド体験向上",
    bullets: [
      "多言語対応チェックインフローの整備（英語・中国語・韓国語）",
      "観光スポット・交通情報の多言語パンフレット設置",
      "Trip.com/Agodaの高評価を維持するための外国人ゲスト対応強化",
    ],
  },
];

// 改善施策 Phase 2
const PHASE2_DESCRIPTION = "一定の投資を伴うが、比較的早期に実行可能な施策";
const PHASE2_ITEMS = [
  {
    title: "(1) 客室設備の選択的更新",
    bullets: [
      "マットレスの定期点検・劣化品の優先交換（長期滞在クレーム防止）",
      "セーフティボックスの全室設置工事（夏季改装と合わせた実施）",
      "客室内コンセント・USB充電設備の増設",
    ],
  },
  {
    title: "(2) 朝食品質の向上",
    bullets: [
      "ビュッフェメニューの季節感・多様性強化",
      "朝食ルームの混雑緩和（時間帯別入場管理）",
    ],
  },
];

// 改善施策 Phase 3
const PHASE3_DESCRIPTION = "設備投資を伴う抜本的な改善施策";
const PHASE3_ITEMS = [
  {
    title: "(1) 夏季改装を活用した設備刷新",
    bullets: [
      "全客室セーフティボックス設置完了",
      "マットレス全数交換による睡眠品質の均一化",
      "バスルームのシャワーヘッド・水回り設備更新",
    ],
  },
  {
    title: "(2) デジタルサービス強化",
    bullets: [
      "スマートフォンチェックイン・チェックアウト導入",
      "客室内QRコードによるサービスリクエスト対応",
    ],
  },
];

// KPI目標データ
const KPI_TARGET_DATA = [
  ["全体平均(10pt換算)", "8.87点", "9.0点以上", "2026年9月"],
  ["高評価率（8-10点）", "81.7%", "85%以上", "2026年9月"],
  ["低評価率（1-4点）", "0.0%", "0%維持", "継続"],
  ["Trip.com平均評価", "9.67/10点", "9.7点以上", "2026年9月"],
  ["Google平均評価", "4.00/5点", "4.3/5点以上", "2026年9月"],
  ["口コミ返信率", "未計測", "100%（48h以内）", "2026年5月"],
  ["アメニティクレーム", "3件/2ヶ月", "0件/2ヶ月", "2026年6月"],
];

// 総括テキスト
const CONCLUSION_PARAGRAPHS = [
  "ダイワロイネットホテル大崎は、2026年2月〜3月の分析期間において全体平均8.87点・低評価0件という極めて優秀な成績を収めています。大崎駅直結の圧倒的立地優位性と安定した清潔感・設備品質が、国内外ゲストからの高い評価を支えています。",
  "最優先課題はアメニティの充実（セーフティボックス設置・基本アメニティの拡充）であり、夏季改装と連動した客室設備の選択的更新が推奨されます。これらの改善は投資対効果が高く、すでに高い水準にある評価をさらに向上させることが期待されます。",
  "インバウンド市場においてTrip.com平均9.67点・Agoda平均9.00点と特に高い評価を得ており、海外向けマーケティングの強化による客室稼働率向上の余地が大きいといえます。多言語対応の充実と口コミ返信率の向上が次のステップとして有効です。",
  "低評価0件という実績を継続・強固なものとするためには、現状の品質維持に加え、ゲストの期待値を上回るサービス提供（アップグレード・記念日対応等）を積極的に実施することが重要です。",
];
const CONCLUSION_FINAL_PARAGRAPH = "全サイトを通じた低評価0件という圧倒的な品質安定性を維持しながら、アメニティ充実と設備更新を通じて全体平均9.0点以上・高評価率85%以上の達成を目指してください。ダイワロイネットホテル大崎の強みは「行けば必ず満足できる」という確実性にあります。";

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
