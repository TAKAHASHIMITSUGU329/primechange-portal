const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, PageNumber, PageBreak, LevelFormat
} = require("docx");

// ============================================================
// ホテル日航川崎 固有設定
// ============================================================
const HOTEL_NAME = "ホテル日航川崎";
const ANALYSIS_PERIOD = "2026年2月〜3月";
const REPORT_DATE = "2026年3月28日";
const OUTPUT_DIR = "納品レポート/ホテル別レポート";

// レビュー基本情報
const REVIEW_COUNT = "130件（重複除外後）";
const TARGET_SITES = "対象サイト：Booking.com / Trip.com / じゃらん / 楽天トラベル / Agoda / Google";

// KPIカードの値
const KPI_CARDS = [
  { label: "全体平均(10pt換算)", value: "8.68", color: "27AE60", bgColor: "E8F5E9" },
  { label: "高評価率(8-10点)", value: "81.5%", color: "27AE60", bgColor: "E8F5E9" },
  { label: "低評価率(1-4点)", value: "1.5%", color: "27AE60", bgColor: "E8F5E9" },
  { label: "レビュー総数", value: "130件", color: "1B3A5C", bgColor: "D5E8F0" },
];

// エグゼクティブサマリー
const EXECUTIVE_SUMMARY_INTRO = "2026年2月〜3月に各OTAサイト・口コミサイトに投稿された130件のレビューを包括的に分析しました。以下が主要な発見事項です。";
const EXECUTIVE_SUMMARY_EVALUATION_1 = "ホテル日航川崎は全体平均8.68点（10点満点）と高い評価を獲得しています。130件中106件（81.5%）が高評価（8-10点）を占め、Trip.com（9.63点）・じゃらん（9.00点）・楽天トラベル（9.00点）では優秀判定を獲得しています。";
const EXECUTIVE_SUMMARY_EVALUATION_2 = "立地・アクセス（53件言及）が最大の強みとして突出しており、JR川崎駅直結・羽田空港へのアクセス利便性が国内外ゲストから絶賛されています。英語対応スタッフの質が高く評価され、インバウンド需要への対応力も強みです。一方、設備の老朽化（3件）・防音・騒音問題（4件）・エレベーター待ち（6件）が改善課題として挙がっています。";

// KEY FINDINGS
const KEY_FINDING_STRENGTH = "立地・アクセス（53件）・客室設備（46件）・清潔感（22件）が三大強みです。英語対応スタッフへの高評価（20件）とロケーションの圧倒的優位性（川崎駅直結・羽田30分・東京20分）が顧客満足度の基盤を形成しています。";
const KEY_FINDING_WEAKNESS = "エレベーター待ち（6件）・防音・騒音問題（4件）・設備老朽化（3件）・アメニティ（3件）が課題として挙がっています。低評価2件は個人情報管理とアクセス面の指摘に起因しています。";
const KEY_FINDING_OPPORTUNITY = "Trip.com（9.63点）・じゃらん・楽天（9.00点）の優秀評価を他サイトへ横展開し、Google評価（7.16点）をせめて8.0点以上に引き上げることが全体スコア向上の最重要課題です。エレベーター・防音改善により中評価層を高評価層へシフトさせることで高評価率85%以上の達成が見込めます。";

// サイト別レビューデータ
const SITE_DATA = [
  ["Trip.com", "30", "9.63", "/10", "9.63", "10.0", "優秀"],
  ["じゃらん", "8", "4.50", "/5", "9.00", "10.0", "優秀"],
  ["楽天トラベル", "6", "4.50", "/5", "9.00", "10.0", "優秀"],
  ["Booking.com", "66", "8.65", "/10", "8.65", "9.0", "良好"],
  ["Google", "19", "3.58", "/5", "7.16", "8.0", "概ね良好"],
  ["Agoda", "1", "7.00", "/10", "7.00", "7.0", "概ね良好"],
];

// 評価分布データ
const DISTRIBUTION_DATA = [
  [10, 56, "43.1%"],
  [9,  24, "18.5%"],
  [8,  26, "20.0%"],
  [7,  10, "7.7%"],
  [6,  11, "8.5%"],
  [5,  1,  "0.8%"],
  [2,  2,  "1.5%"],
];
const DISTRIBUTION_MAX_COUNT = 56;

// 評価分布サマリー
const HIGH_RATING_SUMMARY = "106件（81.5%）";
const MID_RATING_SUMMARY = "22件（16.9%）";
const LOW_RATING_SUMMARY = "2件（1.5%）";
const LOW_RATING_COLOR = "27AE60";

// データ概要テキスト
const DATA_OVERVIEW_SITE_TEXT = "各予約サイトの評価基準が異なる点に留意が必要です。海外OTA（Booking.com/Trip.com/Agoda）は10点満点、国内サイト（じゃらん/楽天トラベル/Google）は5点満点のため、統一比較のため10点換算（5点満点×2）を行いました。Booking.comが最多件数（66件）でインバウンドゲストの意見が分析の中核を担っています。";
const DATA_OVERVIEW_DIST_TEXT = "評価分布はスコア10点が56件（43.1%）で最多、スコア8点が26件（20.0%）・スコア9点が24件（18.5%）と続き、高評価層が厚く106件（81.5%）が8点以上を占めています。スコア7点（10件）・6点（11件）が中評価層を形成しており、これらを高評価に引き上げる施策が平均向上の鍵です。";

// 強みテーマデータ
const STRENGTH_THEMES = [
  ["立地・アクセス", "53件", "「駅前でアクセスは良い」「JR川崎駅直結」「羽田空港から京急線で簡単にアクセスできました」"],
  ["客室・設備", "46件", "「部屋もとても清潔で、気持ちよく過ごせました」「折り紙などの部屋の飾りには驚いて、カプセルコーヒーマシンは紅茶まで」"],
  ["清潔感", "22件", "「部屋は清潔で整頓されており」「部屋もとても清潔で、気持ちよく過ごせました」"],
  ["スタッフ対応", "20件", "「スタッフは礼儀正しく、ほとんどのスタッフが英語を話します」「スタッフはとても親切で」"],
  ["朝食", "23件", "「朝食を1日追加で購入したのですが、フロントでその旨を伝えたら割引クーポンをもらえました。アイスクリームなど、種類も豊富でした！」"],
  ["コスパ・リピート", "8件", "「立地は便利で、宿泊料金は都心部よりもずっと安いです。川崎は東京の中心部からそれほど遠くなく、宿泊費を大幅に節約できます」"],
];

// 強みサブセクション
const STRENGTH_SUB_1_TITLE = "3.1 最大の強み：「立地・アクセス」";
const STRENGTH_SUB_1_TEXT = "立地・アクセスは全130件中53件（40.8%）で言及されており、圧倒的な最大の強みです。JR川崎駅直結・羽田空港30分・東京駅20分というアクセス優位性が国内外ゲストから絶賛されています。バス停がホテル前という利便性も評価されており、川崎という都市のポテンシャルを最大限に活用したロケーション戦略が機能しています。";
const STRENGTH_SUB_1_BULLETS = [
  "JR川崎駅・京急川崎駅への直結的な近接（徒歩数分）",
  "羽田空港まで電車30分・バスで40分という抜群の空港アクセス",
  "東京駅まで約20分という都心へのアクセス",
  "英語を話すスタッフによる国際旅行者へのサポート体制",
];

const STRENGTH_SUB_2_TITLE = "3.2 英語対応スタッフと朝食の充実";
const STRENGTH_SUB_2_TEXT = "インバウンドゲスト（特にTrip.com利用者）から英語対応スタッフへの高評価が際立っています（20件）。朝食は23件と多数の言及があり、バラエティ豊富なビュッフェと割引クーポン対応など柔軟なサービスが評価の差別化要因となっています。日航グループのサービス品質が多くのゲストに認知されています。";

// 弱み分析データ
const WEAKNESS_PRIORITY_DATA = [
  ["S", "エレベーター待ち", "6件言及。多層ホテルでのエレベーター混雑は特に混雑時間帯に問題化。ゲスト体験に直接影響する課題", "6件・高頻度言及"],
  ["A", "防音・騒音問題", "高層階からの電車音・改装工事騒音（4件）。騒音敏感ゲストへの事前通知が必要", "4件"],
  ["A", "設備の老朽化", "バスルーム設備の古さ・排水不良（3件）。施設年数を考慮した計画的更新が必要", "3件"],
  ["B", "Google評価の低さ", "7.16点（概ね良好）で他サイト比較で低い評価。国内ゲストの不満の反映", "19件・7.16点"],
  ["B", "アメニティ", "一部アメニティへの言及あり。充実度向上の余地あり", "3件"],
  ["C", "個人情報管理", "特定の言及あり（1件）。フロント業務における情報管理強化が必要", "1件"],
];

// 改善施策 Phase 1
const PHASE1_DESCRIPTION = "投資不要・オペレーション改善で対応可能な施策";
const PHASE1_ITEMS = [
  {
    title: "(1) エレベーター混雑の緩和",
    bullets: [
      "チェックイン・チェックアウト時間の分散化（時間帯別案内）",
      "混雑時間帯のスタッフ配置強化でゲスト誘導",
      "階段使用を促す案内の設置（低層階向け）",
    ],
  },
  {
    title: "(2) 防音・騒音の事前案内",
    bullets: [
      "予約確認メールに騒音リスク（電車音・改装工事）を明記",
      "騒音に敏感なゲストへの静音客室（中層階・低層）への案内",
      "改装工事期間中の工事時間帯・騒音レベルの透明な情報開示",
    ],
  },
  {
    title: "(3) Google口コミ対策",
    bullets: [
      "全Google口コミへの丁寧な返信（特に低評価への真摯な対応）",
      "満足したゲストへのGoogle口コミ投稿案内",
      "Google評価の月次モニタリングと改善進捗確認",
    ],
  },
];

// 改善施策 Phase 2
const PHASE2_DESCRIPTION = "一定の投資を伴うが、比較的早期に実行可能な施策";
const PHASE2_ITEMS = [
  {
    title: "(1) バスルーム設備の更新",
    bullets: [
      "排水設備の点検・改修（詰まり・流れの悪さ改善）",
      "シャワーヘッド・水回り設備の選択的更新",
      "老朽化が目立つバスルーム設備の優先交換",
    ],
  },
  {
    title: "(2) エレベーター対策（設備的改善）",
    bullets: [
      "エレベーター定員管理の最適化",
      "エレベーター内・周辺サイン整備によるストレス軽減",
    ],
  },
];

// 改善施策 Phase 3
const PHASE3_DESCRIPTION = "設備投資を伴う抜本的な改善施策";
const PHASE3_ITEMS = [
  {
    title: "(1) 客室・設備の全面刷新",
    bullets: [
      "老朽化した客室設備の計画的全面更新",
      "防音性能の向上（高音帯域の電車音対策）",
      "バスルームの全面リノベーション",
    ],
  },
  {
    title: "(2) 共用施設のリニューアル",
    bullets: [
      "エレベーター増設・速度向上によるオペレーション改善",
      "1階改装工事完了を機にロビー・フロントエリアのグレードアップ",
    ],
  },
];

// KPI目標データ
const KPI_TARGET_DATA = [
  ["全体平均(10pt換算)", "8.68点", "8.9点以上", "2026年9月"],
  ["高評価率（8-10点）", "81.5%", "85%以上", "2026年9月"],
  ["低評価率（1-4点）", "1.5%", "0.5%以下", "2026年9月"],
  ["Trip.com平均評価", "9.63/10点", "9.7点以上", "2026年9月"],
  ["Google平均評価", "3.58/5点", "3.8/5点以上", "2026年9月"],
  ["口コミ返信率", "未計測", "100%（48h以内）", "2026年5月"],
  ["エレベータークレーム", "6件/2ヶ月", "2件以下/2ヶ月", "2026年6月"],
];

// 総括テキスト
const CONCLUSION_PARAGRAPHS = [
  "ホテル日航川崎は、2026年2月〜3月において全体平均8.68点・高評価率81.5%と高水準の評価を獲得しています。JR川崎駅直結という圧倒的立地優位性（53件言及）と英語対応スタッフの質の高さ（20件）が、国内外ゲストからの安定した支持を生み出しています。",
  "Trip.com（9.63点）・じゃらん・楽天トラベル（9.00点）での優秀評価は日航グループとしてのブランド力を示しています。一方でGoogle評価（7.16点・概ね良好）が他サイトと比べ低く、エレベーター待ち（6件）・防音問題（4件）・設備老朽化（3件）が中評価層を形成する主因となっています。",
  "最優先課題はエレベーター混雑の緩和と防音・騒音の事前案内強化です。これらはオペレーション改善で対応可能であり、早期に実施できます。バスルーム設備の更新は中期的な投資として計画的に進め、改装完了後の評価向上を目指すことが推奨されます。",
  "川崎という立地を活かしたコスパの良さ（8件評価）と都心・羽田空港へのアクセス優位性は、ビジネス・トランジット双方の需要を取り込める強みであり、この価値訴求を強化したマーケティング展開が有効です。",
];
const CONCLUSION_FINAL_PARAGRAPH = "JR川崎駅直結という圧倒的立地と日航グループのサービス品質を活かし、エレベーター・防音・設備更新という具体的な課題に対処することで、全体平均8.9点以上・高評価率85%以上の達成が実現可能です。Google評価の3.8/5点以上への引き上げが最重要KPIです。";

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
