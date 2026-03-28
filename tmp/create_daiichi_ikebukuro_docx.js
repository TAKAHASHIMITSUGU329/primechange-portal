const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, PageNumber, PageBreak, LevelFormat
} = require("docx");

// ============================================================
// ホテル固有の設定
// ============================================================
const HOTEL_NAME = "第一ホテル池袋";
const ANALYSIS_PERIOD = "2026年2月〜3月";
const REPORT_DATE = "2026年3月28日";
const OUTPUT_DIR = "納品レポート/ホテル別レポート";

const REVIEW_COUNT = "130件（重複除外後）";
const TARGET_SITES = "対象サイト：Booking.com / Trip.com / じゃらん / 楽天トラベル / Agoda / Google";

const KPI_CARDS = [
  { label: "全体平均(10pt換算)", value: "9.18", color: "27AE60", bgColor: "E8F5E9" },
  { label: "高評価率(8-10点)", value: "91.5%", color: "27AE60", bgColor: "E8F5E9" },
  { label: "低評価率(1-4点)", value: "2.3%", color: "27AE60", bgColor: "E8F5E9" },
  { label: "レビュー総数", value: "130件", color: "1B3A5C", bgColor: "D5E8F0" },
];

const EXECUTIVE_SUMMARY_INTRO = "2026年2月〜3月に各OTAサイト・口コミサイトに投稿された130件のレビューを包括的に分析しました。以下が主要な発見事項です。";
const EXECUTIVE_SUMMARY_EVALUATION_1 = "全体平均スコアは10点換算で9.18点（優秀水準）。高評価率91.5%・低評価率わずか2.3%という卓越した評価を達成。Trip.com（9.56）・Google（9.21）・Booking.com（8.82）と主要サイトすべてで優秀〜良好水準を記録しています。";
const EXECUTIVE_SUMMARY_EVALUATION_2 = "最大の強みは「立地・アクセス」（84件・64.6%）で、池袋駅東口徒歩数分という抜群のアクセスが圧倒的支持を得ています。加えて「スタッフ対応」（32件）・「客室・設備」（48件）・「清潔感」（22件）が高評価を支えています。建物の古さを超える価値提供が実現されています。";

const KEY_FINDING_STRENGTH = "Strength：立地（84件・64.6%）・スタッフ対応（32件）・客室快適性（48件）・清潔感（22件）がトップクラスの評価。「2つ星とは思えない3つ星クオリティ」と複数評価。";
const KEY_FINDING_WEAKNESS = "Weakness：Agoda評価の低さ（5.5点・4件）・騒音問題（1件）・排水の遅さ（1件）・客室の狭さ（5件）と軽微な課題のみ。";
const KEY_FINDING_OPPORTUNITY = "Opportunity：Agoda評価の改善と口コミ件数のさらなる拡大で、全サイト優秀水準の達成と認知拡大のポテンシャルが大きい。";

const SITE_DATA = [
  ["Trip.com", "66", "9.56", "/10", "9.56", "10.0", "優秀"],
  ["Google", "33", "4.61", "/5", "9.21", "9.0", "優秀"],
  ["Booking.com", "22", "8.82", "/10", "8.82", "9.0", "良好"],
  ["楽天トラベル", "4", "4.25", "/5", "8.50", "9.0", "良好"],
  ["じゃらん", "1", "4.00", "/5", "8.00", "8.0", "良好"],
  ["Agoda", "4", "5.50", "/10", "5.50", "5.5", "要改善"],
];

const DISTRIBUTION_DATA = [
  [10, 82, "63.1%"],
  [9,  19, "14.6%"],
  [8,  18, "13.8%"],
  [7,   3, "2.3%"],
  [6,   2, "1.5%"],
  [5,   3, "2.3%"],
  [4,   2, "1.5%"],
  [2,   1, "0.8%"],
];
const DISTRIBUTION_MAX_COUNT = 82;

const HIGH_RATING_SUMMARY = "119件（91.5%）";
const MID_RATING_SUMMARY = "8件（6.2%）";
const LOW_RATING_SUMMARY = "3件（2.3%）";
const LOW_RATING_COLOR = "27AE60";

const DATA_OVERVIEW_SITE_TEXT = "各予約サイトの評価基準が異なる点に留意が必要です。海外OTA（Booking.com/Trip.com/Agoda）は10点満点、国内サイト（じゃらん/楽天トラベル/Google）は5点満点のため、統一比較のため10点換算（5点満点×2）を行いました。";
const DATA_OVERVIEW_DIST_TEXT = "10点が全体の63.1%と圧倒的な比率を占め、9点・8点合計28.4%と、高評価層が全体の91.5%を占める。これは競合ホテルと比較しても群を抜いた優秀な分布です。";

const STRENGTH_THEMES = [
  ["立地・アクセス", "84件", "「池袋駅東口から徒歩数分の抜群の立地」「パルコ・西武・東武も近くショッピングに便利」「JR・地下鉄・空港バスがすぐ近く」"],
  ["客室・設備", "48件", "「このホテルは2つ星ではなく3つ星クオリティ」「部屋は清潔で快適」「トリプルルームを提供している珍しいホテル」"],
  ["スタッフ対応", "32件", "「スタッフは親切でフレンドリー」「追加料金でレイトチェックアウト可能」「自動チェックインキオスクで便利」"],
  ["清潔感", "22件", "「建物は古いが清潔」「部屋はとても清潔で快適」「清潔さは許容範囲」"],
  ["朝食", "6件", "「朝食がとても美味しかった」「伝統的な日本料理と洋食で味もかなり美味しい」"],
  ["コスパ・リピート", "8件", "「間違いなく3つ星の評価に値する」「立地は非常に便利で再利用したい」「リピート利用」"],
];

const STRENGTH_SUB_1_TITLE = "3.1 最大の強み：「立地・アクセス」";
const STRENGTH_SUB_1_TEXT = "立地・アクセスは130件中84件（64.6%）で言及されており、第一ホテル池袋の圧倒的かつ最大の強み。池袋駅東口から徒歩数分という立地は国内外のゲストから一貫して最高評価を受けており、この強みが91.5%という高評価率の根幹を支えています。";
const STRENGTH_SUB_1_BULLETS = [
  "池袋駅東口から徒歩わずか数分でJR・地下鉄・空港バスへのアクセスが抜群",
  "パルコ・西武・東武といった大型百貨店が徒歩圏内でショッピングに最適",
  "周辺の豊富な飲食店・コンビニが長期滞在・ビジネス利用者に高く評価",
  "「複数回目の利用」とリピーターが多く、立地が選ばれる理由の筆頭に",
];

const STRENGTH_SUB_2_TITLE = "3.2 「2つ星を超える3つ星クオリティ」の評価";
const STRENGTH_SUB_2_TEXT = "複数のゲストから「2つ星ではなく3つ星クオリティ」との評価を受けており、建物の古さを感じさせない清潔感・スタッフのサービス・客室の快適性が期待値を大きく上回っています。この評価は口コミマーケティングにおける強力な訴求ポイントです。";

const WEAKNESS_PRIORITY_DATA = [
  ["S", "Agoda評価の低さ", "全サイト中最低の5.5点（4件）。要改善水準で全体評価を引き下げる潜在リスク。", "4件・最優先改善"],
  ["A", "客室の狭さへの指摘", "部屋が狭い・同時に大きなスーツケース2個を開けられない等の指摘。", "5件・高頻度指摘"],
  ["B", "排水の遅さ", "シャワー・浴槽の排水が追いつかない問題が報告。", "2件・中程度"],
  ["B", "低層階の油煙臭", "周辺飲食店の多さから低層階に油煙臭の報告あり（フロント対応で解消）。", "1件・対応実績あり"],
  ["C", "窓からの騒音", "騒音で眠れないとの報告。立地柄避けられない面もあるが、防音対策余地あり。", "1件・軽微"],
  ["C", "アメニティのロビー取り置き方式", "一部アメニティをロビーに取りに行く必要があり不便との声。", "1件・軽微"],
];

const PHASE1_DESCRIPTION = "投資不要・オペレーション改善で対応可能な施策（即時実行）";
const PHASE1_ITEMS = [
  {
    title: "(1) 高評価水準の維持体制強化",
    bullets: [
      "全レビューへの丁寧な返信を継続（感謝＋施設情報）",
      "Trip.com・Googleの高評価維持のためのゲスト満足度モニタリング",
      "スタッフサービス品質の継続的な教育・評価",
    ],
  },
  {
    title: "(2) Agoda評価の緊急改善",
    bullets: [
      "Agoda掲載情報（写真・説明文）の精度向上と期待値調整",
      "Agoda経由のゲストへの特別ウェルカム対応",
      "過去の低評価コメントへの丁寧な返信と改善アピール",
    ],
  },
  {
    title: "(3) 低層階・騒音対策",
    bullets: [
      "低層階客室への配室前の油煙・騒音リスク告知",
      "希望ゲストへの高層階優先アサイン対応",
      "排水問題の即時点検・修繕依頼",
    ],
  },
];

const PHASE2_DESCRIPTION = "比較的早期に実行可能な施策（1〜3ヶ月）";
const PHASE2_ITEMS = [
  {
    title: "(1) 客室アメニティの改善",
    bullets: [
      "全客室への基本アメニティ完備（ロビー取り置き方式の廃止）",
      "ウェルカムセットのグレードアップ",
      "客室内の収納改善（大型荷物対応）",
    ],
  },
  {
    title: "(2) 口コミ件数の拡大",
    bullets: [
      "チェックアウト時のQRコードを使った口コミ依頼",
      "Agoda・楽天・じゃらんでの口コミ件数増加施策",
      "リピーター向け特典で口コミ促進",
    ],
  },
  {
    title: "(3) 朝食のさらなる充実",
    bullets: [
      "和洋折衷の豊富なメニューの季節対応",
      "朝食満足度のアンケート実施",
    ],
  },
];

const PHASE3_DESCRIPTION = "設備投資を伴う抜本的改善（3〜6ヶ月）";
const PHASE3_ITEMS = [
  {
    title: "(1) 全サイト優秀水準（9.0点以上）の達成",
    bullets: [
      "Agoda評価を要改善（5.5点）から良好（8.0点以上）へ",
      "Booking.com 9.0点以上への引き上げ",
    ],
  },
  {
    title: "(2) バスルーム設備の更新",
    bullets: [
      "排水設備の全面改修でシャワー利用時の水たまり解消",
      "バスルームのリノベーションで体験価値向上",
    ],
  },
  {
    title: "(3) 防音・快適性向上",
    bullets: [
      "客室の防音性能強化（窓・ドアの防音改修）",
      "シャワー音センサーの見直し（ビー音の解消）",
    ],
  },
];

const KPI_TARGET_DATA = [
  ["全体平均(10pt換算)", "9.18点", "9.3点以上", "2026年9月"],
  ["高評価率（8-10点）", "91.5%", "93%以上", "2026年9月"],
  ["低評価率（1-4点）", "2.3%", "1%以下", "2026年9月"],
  ["Agoda平均", "5.50/10点", "8.0/10点以上", "2026年6月"],
  ["Trip.com平均", "9.56/10点", "9.5点以上維持", "2026年9月"],
  ["口コミ総数", "130件/2ヶ月", "160件/2ヶ月以上", "2026年6月"],
];

const CONCLUSION_PARAGRAPHS = [
  "第一ホテル池袋は、全体平均9.18点・高評価率91.5%という業界トップクラスの評価を獲得しており、池袋駅東口徒歩数分という絶対的な立地優位性と、建物の古さを超える清潔感・スタッフサービスの高さが相乗効果を生み出しています。",
  "「建物は古いが、間違いなく2つ星ではなく3つ星クオリティ」という複数のゲストからの評価は、ハードの限界を超えたサービス価値の証明であり、このホテルの本質的な強みを象徴しています。立地・清潔感・スタッフ対応の三位一体が高評価の根幹です。",
  "課題はAgoda評価（5.5点・要改善）の底上げが最優先であり、これが解決されれば全サイト優秀水準達成が視野に入ります。加えて、客室の排水改善・アメニティの充実化などの細部改善により、さらなる評価向上が期待できます。",
  "現在の高評価を基盤として、口コミ件数の拡大と全サイトでの認知向上を進めることが、池袋エリアでのプレゼンス強化につながります。現状の強みを最大限に活かし、次のステージへ挑戦してください。",
];
const CONCLUSION_FINAL_PARAGRAPH = "業界トップクラスの評価9.18点を誇る第一ホテル池袋。Agoda改善と細部の磨き込みで全サイト優秀水準9.3点以上を目指し、池袋エリアのナンバーワンホテルの座を確固たるものにしてください。";

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
