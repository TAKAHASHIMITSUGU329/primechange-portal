const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, PageNumber, PageBreak, LevelFormat
} = require("docx");

// ============================================================
// ホテル固有の設定
// ============================================================
const HOTEL_NAME = "コンフォートホテル横浜関内";
const ANALYSIS_PERIOD = "2026年2月〜3月";
const REPORT_DATE = "2026年3月28日";
const OUTPUT_DIR = "納品レポート/ホテル別レポート";

const REVIEW_COUNT = "143件（重複除外後）";
const TARGET_SITES = "対象サイト：Booking.com / Trip.com / じゃらん / 楽天トラベル / Agoda / Google";

const KPI_CARDS = [
  { label: "全体平均(10pt換算)", value: "7.52", color: "E67E22", bgColor: "FFF3E0" },
  { label: "高評価率(8-10点)", value: "67.1%", color: "27AE60", bgColor: "E8F5E9" },
  { label: "低評価率(1-4点)", value: "12.6%", color: "C0392B", bgColor: "FFEBEE" },
  { label: "レビュー総数", value: "143件", color: "1B3A5C", bgColor: "D5E8F0" },
];

const EXECUTIVE_SUMMARY_INTRO = "2026年2月〜3月に各OTAサイト・口コミサイトに投稿された143件のレビューを包括的に分析しました。以下が主要な発見事項です。";
const EXECUTIVE_SUMMARY_EVALUATION_1 = "全体平均スコアは10点換算で7.52点（概ね良好）。高評価率67.1%は一定水準を維持しているものの、低評価率12.6%（18件）が課題であり、朝食クオリティと客室設備の不満が集中しています。";
const EXECUTIVE_SUMMARY_EVALUATION_2 = "最大の強みは「朝食・食事」（82件言及）と「客室・設備」（57件）および「立地・アクセス」（45件）。改装工事中による朝食提供形式の変更が一時的な低評価を招いており、工事完了後の回復が期待されます。";

const KEY_FINDING_STRENGTH = "Strength：立地の利便性（45件）・朝食（82件）・客室清潔感（26件）・スタッフ対応（37件）が高評価の主要因。横浜スタジアム・中華街への徒歩アクセスも評価。";
const KEY_FINDING_WEAKNESS = "Weakness：改装中朝食（弁当提供）への失望（8件）・朝食会場の混雑・料理不足（6件）・エアコン老朽化（2件）・客室の狭さ（5件）が低評価原因。";
const KEY_FINDING_OPPORTUNITY = "Opportunity：改装工事完了後の朝食ビュッフェ復活により低評価率の大幅改善が見込まれる。リピート意向の高さを活用した口コミ向上余地が大きい。";

const SITE_DATA = [
  ["Trip.com", "19", "8.32", "/10", "8.32", "8.0", "良好"],
  ["Agoda", "11", "8.00", "/10", "8.00", "8.0", "良好"],
  ["Google", "24", "3.79", "/5", "7.58", "8.0", "概ね良好"],
  ["楽天トラベル", "42", "3.74", "/5", "7.48", "8.0", "概ね良好"],
  ["じゃらん", "19", "3.63", "/5", "7.26", "7.0", "概ね良好"],
  ["Booking.com", "28", "7.00", "/10", "7.00", "7.0", "概ね良好"],
];

const DISTRIBUTION_DATA = [
  [10, 34, "23.8%"],
  [9,   8, "5.6%"],
  [8,  54, "37.8%"],
  [7,   6, "4.2%"],
  [6,  18, "12.6%"],
  [5,   5, "3.5%"],
  [4,  10, "7.0%"],
  [3,   1, "0.7%"],
  [2,   7, "4.9%"],
];
const DISTRIBUTION_MAX_COUNT = 54;

const HIGH_RATING_SUMMARY = "96件（67.1%）";
const MID_RATING_SUMMARY = "29件（20.3%）";
const LOW_RATING_SUMMARY = "18件（12.6%）";
const LOW_RATING_COLOR = "C0392B";

const DATA_OVERVIEW_SITE_TEXT = "各予約サイトの評価基準が異なる点に留意が必要です。海外OTA（Booking.com/Trip.com/Agoda）は10点満点、国内サイト（じゃらん/楽天トラベル/Google）は5点満点のため、統一比較のため10点換算（5点満点×2）を行いました。";
const DATA_OVERVIEW_DIST_TEXT = "10点・8点帯に集中（10点:23.8%、8点:37.8%）し、高評価層が厚い一方、6点以下が29.5%（42件）存在。改装工事期間中の朝食変更が中低評価の一因と考えられます。";

const STRENGTH_THEMES = [
  ["朝食・食事", "82件", "「朝食も期待以上」「朝食カフェスペースがオシャレでした」「朝食は通常通りです」"],
  ["客室・設備", "57件", "「部屋は清潔で快適」「トリプルベッドの設定がありがたい」「設備は必要最低限」"],
  ["立地・アクセス", "45件", "「最高のロケーション。最寄り駅まで5分」「横浜スタジアムや中華街も徒歩圏内」"],
  ["スタッフ対応", "37件", "「スタッフの方みんな一生懸命でした」「受付のスタッフはとても丁寧でした」"],
  ["清潔感", "26件", "「とても清潔な部屋でした」「清潔で気持ちよく過ごせました」"],
  ["コスパ・リピート", "15件", "「連泊プランでお得に」「また利用したい」「コスパが良い」"],
];

const STRENGTH_SUB_1_TITLE = "3.1 最大の強み：「朝食・食事」";
const STRENGTH_SUB_1_TEXT = "朝食は143件中82件（57.3%）で言及されており、コンフォートホテル横浜関内の圧倒的な強み。カフェ兼ブックルームの活用や無料朝食の提供が宿泊者に高く評価されています。";
const STRENGTH_SUB_1_BULLETS = [
  "朝食ビュッフェの豊富なメニューとカフェスペースの雰囲気が好評",
  "2階のカフェ兼ブックルームはコーヒー・紅茶・水・電子レンジも完備",
  "改装工事中も弁当提供で朝食を無料提供（工事完了後はビュッフェ復活）",
  "立地との相乗効果で観光・ビジネス双方の宿泊者に好評",
];

const STRENGTH_SUB_2_TITLE = "3.2 立地・アクセスの優位性";
const STRENGTH_SUB_2_TEXT = "最寄り駅（関内駅）徒歩5分圏内、横浜スタジアム・中華街・山下公園への徒歩アクセスが可能。ホテル向かいのコンビニや周辺の飲食店の充実度も高評価。";

const WEAKNESS_PRIORITY_DATA = [
  ["S", "朝食提供形式の変更", "改装工事中のビュッフェ→弁当変更が予約時と異なるとして失望。直接的な評価低下要因。", "8件・評価低下直結"],
  ["A", "朝食会場の混雑・料理不足", "朝食会場が狭く料理がなくなる。団体客来訪時は特に悪化。", "6件・高頻度不満"],
  ["A", "客室の狭さ・設備老朽化", "ユニットバスの狭さ、エアコン老朽化、カードキー入れなし等の設備不足。", "7件・高頻度不満"],
  ["B", "スタッフ対応の一部不満", "チェックイン時の対応が不丁寧との報告。一部スタッフへの個別不満。", "3件・中程度"],
  ["B", "エレベーター待ち", "エレベーター2台のみで朝夕のピーク時に混雑。", "3件・中程度"],
  ["C", "シャワーヘッドの水漏れ", "複数部屋でシャワーヘッド付近から水漏れが確認されている。", "2件・要点検"],
];

const PHASE1_DESCRIPTION = "投資不要・オペレーション改善で対応可能な施策（即時実行）";
const PHASE1_ITEMS = [
  {
    title: "(1) 朝食工事期間中の告知強化",
    bullets: [
      "予約ページ・確認メールに改装中朝食形式を明記（期待値調整）",
      "チェックイン時に朝食形式を口頭説明するオペレーション標準化",
      "工事完了日程を口コミ返信で積極的に告知",
    ],
  },
  {
    title: "(2) 朝食会場の混雑対策",
    bullets: [
      "朝食提供時間のピーク分散（時間帯別入場整理）",
      "料理補充頻度の増加と補充担当スタッフの固定",
      "混雑時の並び方を床サインで明示",
    ],
  },
  {
    title: "(3) 口コミへの積極的返信",
    bullets: [
      "低評価コメントへの48時間以内返信を目標設定",
      "改装工事完了・ビュッフェ復活を全レビューで周知",
      "高評価レビューへの感謝返信でエンゲージメント向上",
    ],
  },
];

const PHASE2_DESCRIPTION = "比較的早期に実行可能な施策（1〜3ヶ月）";
const PHASE2_ITEMS = [
  {
    title: "(1) 客室アメニティ・設備改善",
    bullets: [
      "カードキーホルダーの全客室設置",
      "シャワーヘッドの点検・交換（水漏れ報告箇所を優先）",
      "タオルハンガー増設で利便性向上",
    ],
  },
  {
    title: "(2) スタッフ接客研修の強化",
    bullets: [
      "チェックイン・アウト対応の標準手順書整備",
      "繁忙時の対応マナー研修実施",
      "顧客からのフィードバック共有会議（月1回）",
    ],
  },
];

const PHASE3_DESCRIPTION = "設備投資を伴う抜本的改善（3〜6ヶ月）";
const PHASE3_ITEMS = [
  {
    title: "(1) 朝食会場の拡充・設備刷新",
    bullets: [
      "改装完了後の朝食ビュッフェ完全復活",
      "座席数増加と動線改善によるピーク時混雑解消",
    ],
  },
  {
    title: "(2) エアコン・客室設備の更新",
    bullets: [
      "老朽化エアコンの計画的更新（特に長期未交換の客室優先）",
      "ユニットバスの改修または設備グレードアップ",
    ],
  },
];

const KPI_TARGET_DATA = [
  ["全体平均(10pt換算)", "7.52点", "8.0点以上", "2026年9月"],
  ["高評価率（8-10点）", "67.1%", "75%以上", "2026年9月"],
  ["低評価率（1-4点）", "12.6%", "5%以下", "2026年6月"],
  ["楽天トラベル平均", "3.74/5点", "4.0/5点以上", "2026年6月"],
  ["じゃらん平均", "3.63/5点", "4.0/5点以上", "2026年6月"],
  ["口コミ返信率", "未計測", "100%（48h以内）", "2026年4月"],
];

const CONCLUSION_PARAGRAPHS = [
  "コンフォートホテル横浜関内は、横浜中心部という抜群の立地と充実した朝食・カフェスペースを核とした強固な競争力を持つホテルです。関内駅徒歩5分・横浜スタジアム徒歩圏という立地優位性は他社との差別化要因として継続的に機能しています。",
  "直近の低評価の主因は改装工事中の朝食提供形式変更（ビュッフェ→弁当）による期待値ギャップです。工事完了後のビュッフェ復活により、この問題は自然解消される見込みであり、低評価率の改善が期待されます。",
  "短期的にはフロントスタッフの接客品質向上、朝食会場の混雑対策、シャワーヘッド等の設備点検を優先実施してください。口コミへの積極的な返信も評価改善に有効な施策です。",
  "中期的にはエアコン等の客室設備の計画的な更新と、朝食会場の動線・座席数改善により、全体平均8.0点・高評価率75%以上の達成を目指してください。",
];
const CONCLUSION_FINAL_PARAGRAPH = "改装工事完了を転換点として、立地・朝食・清潔感という三大強みを最大限に活かした高評価ホテルへの飛躍を期待します。";

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
