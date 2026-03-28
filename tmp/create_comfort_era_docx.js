const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, PageNumber, PageBreak, LevelFormat
} = require("docx");

const HOTEL_NAME = "コンフォートイン東神田ERA";
const ANALYSIS_PERIOD = "2026年2月〜3月";
const REPORT_DATE = "2026年3月28日";
const OUTPUT_DIR = "納品レポート/ホテル別レポート";

const REVIEW_COUNT = "123件（重複除外後）";
const TARGET_SITES = "対象サイト：Booking.com / Trip.com / Agoda / Google / じゃらん / 楽天トラベル";

const KPI_CARDS = [
  { label: "全体平均(10pt換算)", value: "8.90", color: "27AE60", bgColor: "E8F5E9" },
  { label: "高評価率(8-10点)", value: "86.2%", color: "27AE60", bgColor: "E8F5E9" },
  { label: "低評価率(1-4点)", value: "4.9%", color: "E67E22", bgColor: "FFF3E0" },
  { label: "レビュー総数", value: "123件", color: "1B3A5C", bgColor: "D5E8F0" },
];

const EXECUTIVE_SUMMARY_INTRO = "2026年2月〜3月に各OTAサイト・口コミサイトに投稿された123件のレビューを包括的に分析しました。以下が主要な発見事項です。";
const EXECUTIVE_SUMMARY_EVALUATION_1 = "全体平均は10点換算で8.90点（優秀）。高評価率86.2%という卓越した評価水準を達成しており、4ホテル中最高のレビュー件数（123件）からも高い注目度と満足度が読み取れます。Trip.com（9.79点）、Agoda（10.0点）、じゃらん（10.0点）では優秀評価を獲得。Booking.com（8.95点）でも良好水準を確保しています。";
const EXECUTIVE_SUMMARY_EVALUATION_2 = "立地・アクセス（42件）が最多言及で、複数の地下鉄・JR路線への近接性と秋葉原・浅草へのアクセス良好さが外国人ゲストを中心に圧倒的な評価を獲得しています。一方、Googleでの評価（7.65点）が他サイトより大きく低く、スタッフ対応への不満（3件）が評価を下げる主要因となっています。";

const KEY_FINDING_STRENGTH = "立地・アクセス（42件）、客室・設備（41件）、スタッフ対応（31件）が三大強みとして圧倒的な評価。特に無料朝食（20件）がTrip.comなど海外サイトでの高評価を牽引。";
const KEY_FINDING_WEAKNESS = "Google評価（7.65点）を引き下げているスタッフ対応への不満（3件）と、部屋の狭さ（3件）が改善課題。";
const KEY_FINDING_OPPORTUNITY = "現在の優秀な評価水準を維持しながらGoogle評価を8.0点以上に引き上げることで、全体平均9.0点超の達成が視野に入る。";

const SITE_DATA = [
  ["Agoda", "1", "10.00", "/10", "10.00", "10.0", "優秀"],
  ["じゃらん", "2", "5.00", "/5", "10.00", "10.0", "優秀"],
  ["Trip.com", "28", "9.79", "/10", "9.79", "10.0", "優秀"],
  ["Booking.com", "66", "8.95", "/10", "8.95", "9.5", "良好"],
  ["楽天トラベル", "3", "4.00", "/5", "8.00", "8.0", "良好"],
  ["Google", "23", "3.83", "/5", "7.65", "8.0", "概ね良好"],
];

const DISTRIBUTION_DATA = [
  [10, 71, "57.7%"],
  [9,  16, "13.0%"],
  [8,  19, "15.4%"],
  [7,  7, "5.7%"],
  [6,  3, "2.4%"],
  [5,  1, "0.8%"],
  [4,  2, "1.6%"],
  [3,  1, "0.8%"],
  [2,  3, "2.4%"],
];
const DISTRIBUTION_MAX_COUNT = 71;

const HIGH_RATING_SUMMARY = "106件（86.2%）";
const MID_RATING_SUMMARY = "11件（8.9%）";
const LOW_RATING_SUMMARY = "6件（4.9%）";
const LOW_RATING_COLOR = "C0392B";

const DATA_OVERVIEW_SITE_TEXT = "各予約サイトの評価基準が異なる点に留意が必要です。海外OTA（Booking.com/Trip.com/Agoda）は10点満点、国内サイト（じゃらん/楽天トラベル/Google）は5点満点のため、統一比較のため10点換算（5点満点×2）を行いました。";
const DATA_OVERVIEW_DIST_TEXT = "評価分布では10点満点が71件（57.7%）と圧倒的最多で、全体的に非常に高い満足度を示しています。5点以下の低評価は7件（5.7%）にとどまり、Googleでの低評価が中心です。123件という大量のレビュー数から統計的信頼性も高く、現状評価の安定性が確認できます。";

const STRENGTH_THEMES = [
  ["立地・アクセス", "42件", "「複数の地下鉄・JR路線の交差点に位置し秋葉原・浅草へ2〜3駅」「ロケーションが非常に便利、楼下にはセブン-イレブン」"],
  ["清潔感", "18件", "「清潔で静かで快適、地鉄站から6〜8分で便利」「毎日タオル等の備品を交換してもらえ清潔な滞在ができた」"],
  ["スタッフ対応", "31件", "「スタッフは非常にフレンドリーで効率的。朝食の補充も迅速」「スタッフは親切で英語での対応も問題なかった」"],
  ["客室・設備", "41件", "「部屋は広く友人も良い印象。コンパクトながら必要なものが揃っている」「バスルームも清潔で快適な滞在ができた」"],
  ["朝食", "20件", "「無料朝食が付いており、多種多様なメニューが揃っている」「朝食は精致ではないが応有尽有（何でも揃っている）」"],
  ["コスパ・リピート", "5件", "「価格・立地・サービスのバランスが非常に良く、また利用したい」「東京観光のベースとして最適なコスパホテル」"],
];

const STRENGTH_SUB_1_TITLE = "3.1 最大の強み：「立地・アクセスの圧倒的優位性」";
const STRENGTH_SUB_1_TEXT = "複数の地下鉄・JR路線の交差点に位置し、秋葉原・浅草・清澄白河へ2〜3駅、東京駅・銀座へ30分以内というアクセスが外国人旅行者・ビジネス客双方から圧倒的な評価を得ています。建物1階のセブン-イレブンも利便性評価に貢献しています。";
const STRENGTH_SUB_1_BULLETS = [
  "複数地下鉄・JR路線の交差点に位置（秋葉原・浅草へ2〜3駅）",
  "東京駅・銀座・新宿・渋谷へ30分以内のアクセス",
  "1階セブン-イレブンによる24時間便利な周辺環境",
  "外国人旅行者（Trip.com 9.79点）から特に高い評価を獲得",
];

const STRENGTH_SUB_2_TITLE = "3.2 無料朝食とスタッフ対応の相乗効果";
const STRENGTH_SUB_2_TEXT = "無料朝食（20件）は海外OTAでの高評価（Trip.com 9.79点）の核心的要因であり、スタッフ対応（31件）との相乗効果で「また泊まりたい」というリピート意向を強く創出しています。朝食の迅速な補充対応もスタッフ評価として計上されています。";

const WEAKNESS_PRIORITY_DATA = [
  ["S", "スタッフ対応の一部不備", "「スタッフは全く助けてくれない。申し訳ないができませんと言うだけ」「誕生日旅行なのに全く無関心。最低のサービス」というGoogle低評価の主要因", "高影響・3件"],
  ["A", "朝食の混雑", "「朝食はいつ行っても人が多い。補充を待つ時間がある」という繁忙時間帯の混雑問題", "中影響・2件"],
  ["A", "部屋の狭さ", "「部屋はコンパクト」という指摘（ただし多くはポジティブに言及するコンテキストで登場）", "中影響・3件"],
  ["B", "カーテン・備品の経年劣化", "「汚れたカーテン、シャワーカーテン」など一部備品の劣化が指摘される", "中影響・1件"],
  ["B", "騒音の発生", "「最初の夜は別の部屋からの騒音があった」「カーテンが薄く光が入る」という環境面の指摘", "低影響・2件"],
  ["C", "駅からの距離", "「まあまあ普通、駅から遠い」という一部のゲストからの立地への不満", "低影響・2件"],
];

const PHASE1_DESCRIPTION = "投資不要・オペレーション改善で対応可能な施策";
const PHASE1_ITEMS = [
  {
    title: "(1) スタッフ接客品質の均一化",
    bullets: [
      "「できません」で終わらない代替提案マニュアルの整備",
      "ゲストの特別な日（誕生日・記念日）への配慮プロトコルの導入",
      "月次ロールプレイング研修でゲスト対応の質を統一",
    ],
  },
  {
    title: "(2) 朝食の混雑対策",
    bullets: [
      "繁忙時間帯（7〜9時）の朝食スタッフ増配置",
      "料理補充のモニタリング体制強化（5分以内補充ルール）",
      "朝食時間帯分散のためのインセンティブ案内の導入",
    ],
  },
  {
    title: "(3) 劣化備品の交換",
    bullets: [
      "カーテン・シャワーカーテンの定期交換サイクルの確立",
      "客室備品点検チェックリストの月次実施",
    ],
  },
];

const PHASE2_DESCRIPTION = "一定の投資を伴うが、比較的早期に実行可能な施策";
const PHASE2_ITEMS = [
  {
    title: "(1) Google評価の改善戦略",
    bullets: [
      "Googleでのネガティブ口コミへの積極的な日本語・英語返信",
      "日本人ゲスト向けの接客プロトコルの強化",
      "Googleビジネスプロフィールの最適化とフォト更新",
    ],
  },
  {
    title: "(2) 防音・遮光設備の改善",
    bullets: [
      "遮光カーテンの全室交換（現行カーテンが薄いとの指摘対応）",
      "隣室騒音対策として壁の隙間シールの施工",
    ],
  },
];

const PHASE3_DESCRIPTION = "設備投資を伴う抜本的な改善施策";
const PHASE3_ITEMS = [
  {
    title: "(1) 朝食エリアの拡張・設備強化",
    bullets: [
      "朝食スペースの席数拡大または2部制の導入",
      "朝食メニューの多様化と品質向上",
    ],
  },
  {
    title: "(2) 客室の高グレード化",
    bullets: [
      "上位グレードルームの設定による単価向上",
      "スマートロック・IoT設備の導入",
    ],
  },
];

const KPI_TARGET_DATA = [
  ["全体平均(10pt換算)", "8.90点", "9.2点以上", "2026年9月"],
  ["高評価率（8-10点）", "86.2%", "90%以上", "2026年9月"],
  ["低評価率（1-4点）", "4.9%", "2%以下", "2026年9月"],
  ["Google平均評価", "3.83/5点", "4.3/5点以上", "2026年9月"],
  ["Trip.com平均評価", "9.79/10点", "9.8/10点維持", "2026年9月"],
  ["口コミ返信率", "未計測", "100%（48h以内）", "2026年6月"],
  ["スタッフ対応クレーム", "3件/2ヶ月", "0件/2ヶ月", "2026年6月"],
];

const CONCLUSION_PARAGRAPHS = [
  "コンフォートイン東神田ERAは全体平均8.90点・高評価率86.2%・レビュー123件という圧倒的な評価実績を誇り、4ホテル中最高のパフォーマンスを示しています。複数路線の交差点という立地優位性と無料朝食・充実したスタッフ対応が外国人旅行者を中心に高い満足度を創出しています。",
  "主要な課題はGoogle評価（7.65点）を引き下げているスタッフ対応への一部不満（3件）です。「申し訳ないができません」で終わらない代替提案型の接客マニュアルと、ゲストの特別な日への配慮プロトコルの導入は最小コストで最大効果が期待できるPhase 1施策です。",
  "朝食の混雑（2件）はスタッフの増配置と補充モニタリングの強化で対応可能です。劣化しているカーテン等の備品は定期交換サイクルの確立で継続的な品質維持が実現します。",
  "現在の優秀な評価水準を基盤として、Googleスコアの改善に集中することで2026年9月末に全体平均9.2点以上・高評価率90%以上という更なる高みを目指します。",
];
const CONCLUSION_FINAL_PARAGRAPH = "東京でのトップクラスの口コミ評価を誇るコンフォートイン東神田ERAの強みをさらに磨き、外国人旅行者に最も愛されるホテルとしての地位を確立しましょう。";

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
