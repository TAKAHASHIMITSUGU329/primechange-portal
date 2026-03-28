const pptxgen = require("pptxgenjs");
const pres = new pptxgen();

const HOTEL_NAME = "アパホテル蒲田駅東";
const ANALYSIS_PERIOD = "2026年2月〜3月";
const REVIEW_COUNT = "58件（5サイト）";
const REPORT_DATE = "2026年3月28日";
const OUTPUT_DIR = "納品レポート/ホテル別レポート";

const KPI_AVG = "7.34";
const KPI_HIGH_RATE = "62.1%";
const KPI_LOW_RATE = "10.3%";
const KPI_TOTAL_COUNT = "58件";

const SUMMARY_STRENGTHS = [
  { title: "立地・アクセス", desc: "羽田空港と京急・JR両線利用可能。ビジネス出張・空港乗継に最適（18件）" },
  { title: "清潔感・設備", desc: "客室の清潔感と必要設備の充実が安定して高評価（9件・16件）" },
  { title: "コストパフォーマンス", desc: "立地・価格バランスが良く、出張利用者から支持される価格帯（6件）" },
];

const SUMMARY_WEAKNESSES = [
  { title: "部屋の狭さ", desc: "ダブル・ツイン客室でシングルと変わらない狭さを指摘する声が複数（6件）" },
  { title: "アメニティ不足", desc: "化粧水・乳液等のスキンケアアメニティの未設置が不満要因（4件）" },
  { title: "Agoda評価の低さ", desc: "外国人ゲストからの評価がAgodaで6.69点と要改善水準（16件）" },
];

const SITE_CHART_LABELS = ["Google", "じゃらん", "Booking.com", "楽天トラベル", "Agoda"];
const SITE_CHART_VALUES = [7.88, 7.50, 7.47, 7.00, 6.69];

const SITE_INSIGHT_TEXTS = [
  { text: "サイト別パフォーマンス分析", options: { bold: true, fontSize: 12, breakLine: true } },
  { text: "", options: { fontSize: 6, breakLine: true } },
  { text: "国内サイト強み", options: { bold: true, color: "16A34A", breakLine: true } },
  { text: "Google・じゃらんで7.5点以上を維持", options: { fontSize: 9, color: "64748B", breakLine: true } },
  { text: "", options: { fontSize: 6, breakLine: true } },
  { text: "要改善", options: { bold: true, color: "DC2626", breakLine: true } },
  { text: "Agodaが6.69点と全サイト最低。外国人向け対応強化が急務", options: { fontSize: 9, color: "64748B", breakLine: true } },
  { text: "", options: { fontSize: 6, breakLine: true } },
  { text: "注：", options: { bold: true, fontSize: 9, color: "64748B" } },
  { text: "国内サイト（5点満点）は10点換算で比較", options: { fontSize: 9, color: "64748B" } },
];

const SITE_TABLE_ROWS_DATA = [
  ["Google", "17", "3.94", "/5", "7.88"],
  ["じゃらん", "4", "3.75", "/5", "7.50"],
  ["Booking.com", "17", "7.47", "/10", "7.47"],
  ["楽天トラベル", "4", "3.50", "/5", "7.00"],
  ["Agoda", "16", "6.69", "/10", "6.69"],
];

const DIST_DOUGHNUT_LABELS = ["高評価(8-10)", "中評価(5-7)", "低評価(1-4)"];
const DIST_DOUGHNUT_VALUES = [36, 16, 6];

const DIST_BAR_LABELS = ["10点", "9点", "8点", "7点", "6点", "5点", "4点", "3点", "2点"];
const DIST_BAR_VALUES = [8, 3, 25, 5, 8, 3, 4, 1, 1];

const DIST_SUMMARY_CARDS = [
  { label: "高評価 8-10点", val: "36件（62.1%）", col: "16A34A", bg: "DCFCE7" },
  { label: "中評価 5-7点", val: "16件（27.6%）", col: "D4A843", bg: "FFF7ED" },
  { label: "低評価 1-4点", val: "6件（10.3%）", col: "DC2626", bg: "FEE2E2" },
];

const STRENGTHS_CARDS = [
  { theme: "立地・アクセス", count: "18件", desc: "羽田空港と京急・JR両線利用可能。ビジネス・乗継利用に最適", quote: "「京急とJRの両方からアクセスできる点がお勧めです」" },
  { theme: "清潔感", count: "9件", desc: "客室の清潔さへの言及が多く、日常的な清掃品質への高評価", quote: "「清潔感もありお勧めできます」" },
  { theme: "スタッフ対応", count: "9件", desc: "チェックイン・チェックアウトの丁寧な対応が評価される", quote: "「スムーズで対応が丁寧でした」" },
  { theme: "客室・設備", count: "16件", desc: "ビジネス利用に必要な設備が充実し、快適な睡眠環境を提供", quote: "「ベッドは快適に睡眠をとることができました」" },
  { theme: "朝食", count: "3件", desc: "シンプルながらコスパの高い朝食が一定の評価を得ている", quote: "「朝食付きプランでコストパフォーマンスが高い」" },
  { theme: "コスパ・リピート", count: "6件", desc: "立地と価格のバランスが良く、ビジネス出張での継続利用が多い", quote: "「価格帯から考えると十分な水準のホテルです」" },
];

const WEAKNESS_ITEMS = [
  { pri: "S", cat: "部屋の狭さ", detail: "ダブル・ツインがシングルと変わらない広さ。荷物置き場が不足", count: "6件", color: "DC2626" },
  { pri: "A", cat: "アメニティ不足", detail: "化粧水・乳液・アフターシェーブ未設置。女性・長期滞在者に不評", count: "4件", color: "EA580C" },
  { pri: "A", cat: "WiFi品質", detail: "インターネット接続の不安定さ。ビジネス利用ゲストへの影響大", count: "2件", color: "EA580C" },
  { pri: "B", cat: "スタッフ接客の均一化", detail: "一部フロントスタッフへの「冷たい・愛想がない」という指摘", count: "2件", color: "D4A843" },
  { pri: "B", cat: "清潔感の維持", detail: "「汚くて埃っぽい」「前宿泊者の残留物あり」という清掃不徹底事例", count: "1件", color: "D4A843" },
  { pri: "C", cat: "朝食コスト", detail: "外国人ゲストを中心に朝食の割高感・内容の物足りなさを指摘", count: "1件", color: "94A3B8" },
];

const PHASE1_CARDS = [
  { title: "客室清掃の強化", items: ["清掃後の二重チェック体制の導入", "前宿泊者残留物の確認必須化", "埃・水回り・備品の点検項目標準化", "清掃品質スコアの月次モニタリング"] },
  { title: "フロント接客の均一化", items: ["電話・対面応対マニュアルの整備", "月次ロールプレイング研修の実施", "「共感・寄り添い」トレーニングの導入"] },
  { title: "アメニティラインナップ拡充", items: ["化粧水・乳液・アフターシェーブの常備", "フロント提供型アメニティリストの掲示", "OTA掲載情報へのアメニティ記載追加"] },
  { title: "口コミ返信体制の整備", items: ["全サイトへの48時間以内返信ルール化", "ネガティブ口コミへの具体的改善回答", "Agodaへの重点返信対応の実施"] },
];

const PHASE2_ITEMS = [
  { title: "WiFiインフラアップグレード", items: ["全室安定接続のためのアクセスポイント増設", "ビジネス向け接続速度の保証・SLA設定", "多言語での接続手順案内の掲示"] },
  { title: "枕・寝具の品質向上", items: ["複数種類の枕から選択可能なシステム導入", "高品質マットレストッパーの導入検討", "ベッドリネンの品質グレードアップ"] },
  { title: "多言語対応の強化", items: ["英語対応フロントスタッフの育成", "多言語デジタルサイネージの設置"] },
];

const PHASE3_ITEMS = [
  { title: "客室リノベーション", items: ["ダブル・ツインルームの実質的な広さ確保", "スマート収納設計による体感スペース拡大", "コンパクト客室のインテリア刷新"] },
  { title: "大浴場または共用シャワー設備の検討", items: ["近隣アパホテル大浴場との差別化戦略の検討", "大浴場設置によるAgodaスコア引き上げ効果の試算"] },
  { title: "外国人ゲスト対応インフラ整備", items: ["キャッシュレス・QRチェックイン導入", "多言語コンシェルジュサービスの検討", "外国語口コミサイトへの積極的プロモーション"] },
];

const KPI_TARGET_ROWS = [
  ["全体平均(10pt換算)", "7.34点", "7.8点以上", "2026年9月"],
  ["高評価率（8-10点）", "62.1%", "70%以上", "2026年9月"],
  ["低評価率（1-4点）", "10.3%", "5%以下", "2026年9月"],
  ["Agoda平均評価", "6.69/10点", "7.5/10点以上", "2026年9月"],
  ["Google平均評価", "3.94/5点", "4.2/5点以上", "2026年9月"],
  ["口コミ返信率", "未計測", "100%（48h以内）", "2026年6月"],
  ["清潔感クレーム件数", "1件/2ヶ月", "0件/2ヶ月", "2026年6月"],
];

const KPI_NOTE_BOLD = "モニタリング方針：";
const KPI_NOTE_TEXT = "毎月末に全サイトの口コミを収集・集計し、KPI達成状況を確認。3ヶ月連続で目標未達の項目は優先的に施策を強化します。";

const CLOSING_PARAGRAPHS = [
  { text: "アパホテル蒲田駅東は羽田空港へのアクセスと京急・JR両線利用という強固な立地優位性を持ち、ビジネス出張客から安定した支持を得ています。全体平均7.34点・高評価率62.1%は健全な水準ですが、Agodaでの評価（6.69点）が課題です。", options: { breakLine: true, fontSize: 12 } },
  { text: "", options: { fontSize: 6, breakLine: true } },
  { text: "部屋の狭さ（6件）とアメニティ不足（4件）という繰り返し指摘される課題は、Phase 1のオペレーション改善から段階的に解決できます。特に化粧水・乳液などのアメニティ拡充はコストも低く即座に実施すべき改善です。", options: { breakLine: true, fontSize: 12 } },
  { text: "", options: { fontSize: 6, breakLine: true } },
  { text: "スタッフ接客の均一化とWiFi整備を並行して進めることで、ビジネス利用ゲストのリピート率向上が期待できます。Agodaへの重点的な口コミ返信対応も外国人ゲストへの誠実さを示す有効な施策です。", options: { breakLine: true, fontSize: 12 } },
  { text: "", options: { fontSize: 6, breakLine: true } },
  { text: "立地という揺るぎない強みを活かし、2026年9月末までに全体平均7.8点・高評価率70%の達成を目指しましょう。", options: { bold: true, fontSize: 12, color: "D4A843" } },
];

pres.layout = "LAYOUT_16x9";
pres.author = "Hotel Consulting";
pres.title = HOTEL_NAME + " 口コミ分析改善レポート";

// === Color Palette (Midnight Executive) ===
const C = {
  navy: "1A2744",
  navyLight: "243556",
  blue: "3B7DD8",
  blueLight: "5A9BE6",
  ice: "E8EFF8",
  white: "FFFFFF",
  offWhite: "F5F7FA",
  gray: "64748B",
  grayLight: "94A3B8",
  grayDark: "334155",
  green: "16A34A",
  greenBg: "DCFCE7",
  red: "DC2626",
  redBg: "FEE2E2",
  orange: "EA580C",
  orangeBg: "FFF7ED",
  gold: "D4A843",
};

// === Helpers ===
const shadow = () => ({ type: "outer", blur: 4, offset: 2, angle: 135, color: "000000", opacity: 0.10 });

function addFooter(slide, pageNum) {
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.navy } });
  slide.addText("Confidential", { x: 0.5, y: 5.25, w: 3, h: 0.375, fontSize: 8, color: C.grayLight, fontFace: "Arial", valign: "middle" });
  slide.addText(String(pageNum), { x: 9, y: 5.25, w: 0.5, h: 0.375, fontSize: 8, color: C.grayLight, fontFace: "Arial", align: "right", valign: "middle" });
}

function addContentHeader(slide, title, subtitle) {
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.85, fill: { color: C.navy } });
  slide.addText(title, { x: 0.6, y: 0.08, w: 8, h: 0.5, fontSize: 22, fontFace: "Arial", color: C.white, bold: true, margin: 0 });
  if (subtitle) {
    slide.addText(subtitle, { x: 0.6, y: 0.5, w: 8, h: 0.3, fontSize: 10, fontFace: "Arial", color: C.blueLight, margin: 0 });
  }
}

function kpiCard(slide, x, y, w, h, label, value, color, bgColor) {
  slide.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: bgColor }, shadow: shadow() });
  slide.addShape(pres.shapes.RECTANGLE, { x, y, w, h: 0.05, fill: { color } });
  slide.addText(label, { x, y: y + 0.15, w, h: 0.3, fontSize: 10, fontFace: "Arial", color: C.gray, align: "center", margin: 0 });
  slide.addText(value, { x, y: y + 0.4, w, h: 0.55, fontSize: 32, fontFace: "Arial", color, bold: true, align: "center", margin: 0 });
}

// ==========================================
// SLIDE 1: TITLE
// ==========================================
let s1 = pres.addSlide();
s1.background = { color: C.navy };

// Decorative shapes
s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: C.navyLight, transparency: 40 } });
s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 2.0, w: 0.08, h: 1.8, fill: { color: C.gold } });

s1.addText("口コミ分析", { x: 0.8, y: 1.5, w: 8, h: 0.7, fontSize: 20, fontFace: "Arial", color: C.blueLight, charSpacing: 6, margin: 0 });
s1.addText("改善レポート", { x: 0.8, y: 2.1, w: 8, h: 0.9, fontSize: 42, fontFace: "Arial", color: C.white, bold: true, margin: 0 });
s1.addShape(pres.shapes.LINE, { x: 0.8, y: 3.1, w: 2, h: 0, line: { color: C.gold, width: 2 } });

// TODO: CUSTOMIZE - ホテル名（タイトルスライド）
s1.addText(HOTEL_NAME, { x: 0.8, y: 3.3, w: 8, h: 0.5, fontSize: 16, fontFace: "Arial", color: C.grayLight, margin: 0 });

// TODO: CUSTOMIZE - 分析期間・件数・日付（タイトルスライド）
s1.addText([
  { text: "分析対象期間：" + ANALYSIS_PERIOD, options: { breakLine: true } },
  { text: "レビュー総数：" + REVIEW_COUNT, options: { breakLine: true } },
  { text: "作成日：" + REPORT_DATE },
], { x: 0.8, y: 4.1, w: 5, h: 0.9, fontSize: 10, fontFace: "Arial", color: C.grayLight, margin: 0, paraSpaceAfter: 4 });

// Right side decorative
s1.addShape(pres.shapes.RECTANGLE, { x: 7.5, y: 0, w: 2.5, h: 5.625, fill: { color: C.blue, transparency: 85 } });

// ==========================================
// SLIDE 2: EXECUTIVE SUMMARY
// ==========================================
let s2 = pres.addSlide();
s2.background = { color: C.offWhite };
addContentHeader(s2, "エグゼクティブサマリー", "Executive Summary");

// TODO: CUSTOMIZE - KPI Cards（値はファイル上部の変数で設定）
kpiCard(s2, 0.5, 1.15, 2.05, 1.05, "全体平均(10pt換算)", KPI_AVG, C.blue, C.white);
kpiCard(s2, 2.75, 1.15, 2.05, 1.05, "高評価率(8-10点)", KPI_HIGH_RATE, C.green, C.white);
kpiCard(s2, 5.0, 1.15, 2.05, 1.05, "低評価率(1-4点)", KPI_LOW_RATE, C.green, C.white);
kpiCard(s2, 7.25, 1.15, 2.25, 1.05, "レビュー総数", KPI_TOTAL_COUNT, C.navy, C.white);

// Strengths box
s2.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 2.5, w: 4.3, h: 2.5, fill: { color: C.white }, shadow: shadow() });
s2.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 2.5, w: 4.3, h: 0.05, fill: { color: C.green } });
s2.addText("Strengths", { x: 0.7, y: 2.6, w: 3, h: 0.35, fontSize: 14, fontFace: "Arial", color: C.green, bold: true, margin: 0 });

// TODO: CUSTOMIZE - 強みテキスト（ファイル上部のSUMMARY_STRENGTHSで設定）
const strengthTextItems = [];
SUMMARY_STRENGTHS.forEach((s, i) => {
  strengthTextItems.push({ text: s.title + " ", options: { bold: true, breakLine: true } });
  strengthTextItems.push({ text: "  " + s.desc, options: { fontSize: 9, color: C.gray, breakLine: i < SUMMARY_STRENGTHS.length - 1 } });
});
s2.addText(strengthTextItems, { x: 0.7, y: 3.0, w: 3.8, h: 1.8, fontSize: 11, fontFace: "Arial", color: C.grayDark, margin: 0, paraSpaceAfter: 2 });

// Weaknesses box
s2.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 2.5, w: 4.3, h: 2.5, fill: { color: C.white }, shadow: shadow() });
s2.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 2.5, w: 4.3, h: 0.05, fill: { color: C.red } });
s2.addText("Weaknesses", { x: 5.4, y: 2.6, w: 3, h: 0.35, fontSize: 14, fontFace: "Arial", color: C.red, bold: true, margin: 0 });

// TODO: CUSTOMIZE - 弱みテキスト（ファイル上部のSUMMARY_WEAKNESSESで設定）
const weaknessTextItems = [];
SUMMARY_WEAKNESSES.forEach((w, i) => {
  weaknessTextItems.push({ text: w.title + " ", options: { bold: true, breakLine: true } });
  weaknessTextItems.push({ text: "  " + w.desc, options: { fontSize: 9, color: C.gray, breakLine: i < SUMMARY_WEAKNESSES.length - 1 } });
});
s2.addText(weaknessTextItems, { x: 5.4, y: 3.0, w: 3.8, h: 1.8, fontSize: 11, fontFace: "Arial", color: C.grayDark, margin: 0, paraSpaceAfter: 2 });

addFooter(s2, 2);

// ==========================================
// SLIDE 3: SITE-BY-SITE RATINGS
// ==========================================
let s3 = pres.addSlide();
s3.background = { color: C.offWhite };
addContentHeader(s3, "サイト別評価分析", "Rating Analysis by Platform");

// TODO: CUSTOMIZE - 棒グラフ（ファイル上部のSITE_CHART_*で設定）
s3.addChart(pres.charts.BAR, [
  { name: "10pt換算平均", labels: SITE_CHART_LABELS, values: SITE_CHART_VALUES }
], {
  x: 0.5, y: 1.1, w: 5.5, h: 3.5,
  barDir: "bar",
  chartColors: [C.blue],
  chartArea: { fill: { color: C.white }, roundedCorners: true },
  catAxisLabelColor: C.grayDark, catAxisLabelFontSize: 10,
  valAxisLabelColor: C.gray, valAxisLabelFontSize: 9,
  valGridLine: { color: "E2E8F0", size: 0.5 },
  catGridLine: { style: "none" },
  showValue: true, dataLabelPosition: "outEnd", dataLabelColor: C.grayDark,
  showLegend: false,
  valAxisMaxVal: 10,
});

// Insight box
s3.addShape(pres.shapes.RECTANGLE, { x: 6.3, y: 1.1, w: 3.3, h: 3.5, fill: { color: C.white }, shadow: shadow() });
s3.addShape(pres.shapes.RECTANGLE, { x: 6.3, y: 1.1, w: 3.3, h: 0.05, fill: { color: C.gold } });
s3.addText("Insight", { x: 6.5, y: 1.2, w: 2, h: 0.35, fontSize: 13, fontFace: "Arial", color: C.gold, bold: true, margin: 0 });

// TODO: CUSTOMIZE - Insightテキスト（ファイル上部のSITE_INSIGHT_TEXTSで設定）
s3.addText(SITE_INSIGHT_TEXTS, { x: 6.5, y: 1.55, w: 2.9, h: 2.9, fontSize: 11, fontFace: "Arial", color: C.grayDark, margin: 0, paraSpaceAfter: 2 });

// Table
const tblHeader2 = [
  { text: "サイト名", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 9, fontFace: "Arial", align: "center" } },
  { text: "件数", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 9, fontFace: "Arial", align: "center" } },
  { text: "ネイティブ平均", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 9, fontFace: "Arial", align: "center" } },
  { text: "尺度", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 9, fontFace: "Arial", align: "center" } },
  { text: "10pt換算", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 9, fontFace: "Arial", align: "center" } },
];

// TODO: CUSTOMIZE - テーブルデータ（ファイル上部のSITE_TABLE_ROWS_DATAで設定）
const tblRows = SITE_TABLE_ROWS_DATA.map((r, i) => r.map(cell => ({
  text: cell, options: { fontSize: 9, fontFace: "Arial", align: "center", fill: { color: i % 2 === 0 ? C.ice : C.white } }
})));

s3.addTable([tblHeader2, ...tblRows], {
  x: 0.5, y: 4.7, w: 9, h: 0.1,
  colW: [2.0, 1.0, 2.0, 1.0, 3.0],
  border: { pt: 0.5, color: "CBD5E1" },
  rowH: [0.25, 0.22, 0.22, 0.22, 0.22, 0.22, 0.22],
});

addFooter(s3, 3);

// ==========================================
// SLIDE 4: RATING DISTRIBUTION
// ==========================================
let s4 = pres.addSlide();
s4.background = { color: C.offWhite };
addContentHeader(s4, "評価分布分析（10点換算）", "Rating Distribution");

// TODO: CUSTOMIZE - ドーナツチャート（ファイル上部のDIST_DOUGHNUT_*で設定）
s4.addChart(pres.charts.DOUGHNUT, [{
  name: "評価分布（10pt換算）",
  labels: DIST_DOUGHNUT_LABELS,
  values: DIST_DOUGHNUT_VALUES,
}], {
  x: 0.3, y: 1.2, w: 4.0, h: 3.2,
  chartColors: [C.green, C.gold],
  showPercent: true,
  dataLabelColor: C.white,
  dataLabelFontSize: 11,
  showTitle: false,
  showLegend: true,
  legendPos: "b",
  legendFontSize: 9,
  legendColor: C.gray,
});

// TODO: CUSTOMIZE - 棒グラフ（ファイル上部のDIST_BAR_*で設定）
s4.addChart(pres.charts.BAR, [{
  name: "件数",
  labels: DIST_BAR_LABELS,
  values: DIST_BAR_VALUES,
}], {
  x: 4.5, y: 1.2, w: 5.2, h: 3.2,
  barDir: "col",
  chartColors: [C.green, C.green, C.green, C.gold, C.gold, C.gold],
  chartArea: { fill: { color: C.white }, roundedCorners: true },
  catAxisLabelColor: C.grayDark, catAxisLabelFontSize: 9,
  valAxisLabelColor: C.gray, valAxisLabelFontSize: 8,
  valGridLine: { color: "E2E8F0", size: 0.5 },
  catGridLine: { style: "none" },
  showValue: true, dataLabelPosition: "outEnd", dataLabelColor: C.grayDark, dataLabelFontSize: 9,
  showLegend: false,
});

// TODO: CUSTOMIZE - サマリーカード（ファイル上部のDIST_SUMMARY_CARDSで設定）
DIST_SUMMARY_CARDS.forEach((c, i) => {
  const x = 0.5 + i * 3.1;
  s4.addShape(pres.shapes.RECTANGLE, { x, y: 4.55, w: 2.9, h: 0.55, fill: { color: c.bg } });
  s4.addText(c.label, { x, y: 4.55, w: 1.5, h: 0.55, fontSize: 10, fontFace: "Arial", color: c.col, bold: true, valign: "middle", margin: [0,0,0,8] });
  s4.addText(c.val, { x: x + 1.4, y: 4.55, w: 1.5, h: 0.55, fontSize: 11, fontFace: "Arial", color: c.col, bold: true, align: "right", valign: "middle", margin: [0,8,0,0] });
});

addFooter(s4, 4);

// ==========================================
// SLIDE 5: STRENGTHS
// ==========================================
let s5 = pres.addSlide();
s5.background = { color: C.offWhite };
addContentHeader(s5, "強み分析", "Strength Analysis");

// TODO: CUSTOMIZE - 強みカード（ファイル上部のSTRENGTHS_CARDSで設定）
STRENGTHS_CARDS.forEach((s, i) => {
  const row = Math.floor(i / 3);
  const col = i % 3;
  const x = 0.5 + col * 3.1;
  const y = 1.15 + row * 2.0;

  s5.addShape(pres.shapes.RECTANGLE, { x, y, w: 2.9, h: 1.75, fill: { color: C.white }, shadow: shadow() });
  s5.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.06, h: 1.75, fill: { color: C.green } });

  // Count badge
  s5.addShape(pres.shapes.RECTANGLE, { x: x + 2.0, y: y + 0.08, w: 0.72, h: 0.25, fill: { color: C.greenBg } });
  s5.addText(s.count, { x: x + 2.0, y: y + 0.08, w: 0.72, h: 0.25, fontSize: 9, fontFace: "Arial", color: C.green, bold: true, align: "center", valign: "middle", margin: 0 });

  s5.addText(s.theme, { x: x + 0.15, y: y + 0.08, w: 1.8, h: 0.3, fontSize: 13, fontFace: "Arial", color: C.navy, bold: true, margin: 0 });
  s5.addText(s.desc, { x: x + 0.15, y: y + 0.4, w: 2.6, h: 0.4, fontSize: 9, fontFace: "Arial", color: C.gray, margin: 0 });
  s5.addText(s.quote, { x: x + 0.15, y: y + 0.85, w: 2.6, h: 0.7, fontSize: 9, fontFace: "Arial", color: C.blue, italic: true, margin: 0 });
});

addFooter(s5, 5);

// ==========================================
// SLIDE 6: WEAKNESS / PRIORITY MATRIX
// ==========================================
let s6 = pres.addSlide();
s6.background = { color: C.offWhite };
addContentHeader(s6, "弱み分析・優先度マトリクス", "Weakness Analysis & Priority Matrix");

// TODO: CUSTOMIZE - 弱みテーブル（ファイル上部のWEAKNESS_ITEMSで設定）
const wTblHeader = [
  { text: "優先度", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 9, fontFace: "Arial", align: "center" } },
  { text: "課題カテゴリ", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 9, fontFace: "Arial" } },
  { text: "具体的内容", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 9, fontFace: "Arial" } },
  { text: "件数", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 9, fontFace: "Arial", align: "center" } },
];

const wTblRows = WEAKNESS_ITEMS.map((w, i) => [
  { text: w.pri, options: { fontSize: 12, fontFace: "Arial", align: "center", bold: true, color: w.color, fill: { color: i % 2 === 0 ? C.ice : C.white } } },
  { text: w.cat, options: { fontSize: 9, fontFace: "Arial", bold: true, color: C.grayDark, fill: { color: i % 2 === 0 ? C.ice : C.white } } },
  { text: w.detail, options: { fontSize: 9, fontFace: "Arial", color: C.gray, fill: { color: i % 2 === 0 ? C.ice : C.white } } },
  { text: w.count, options: { fontSize: 9, fontFace: "Arial", align: "center", color: C.grayDark, fill: { color: i % 2 === 0 ? C.ice : C.white } } },
]);

s6.addTable([wTblHeader, ...wTblRows], {
  x: 0.5, y: 1.15, w: 9, h: 0.1,
  colW: [0.8, 2.0, 5.0, 1.2],
  border: { pt: 0.5, color: "CBD5E1" },
  rowH: [0.3, 0.35, 0.35, 0.35, 0.35, 0.35, 0.35, 0.35, 0.35, 0.35],
});

// Legend
s6.addText([
  { text: "S ", options: { bold: true, color: C.red, fontSize: 10 } },
  { text: "= 最優先   ", options: { fontSize: 9, color: C.gray } },
  { text: "A ", options: { bold: true, color: C.orange, fontSize: 10 } },
  { text: "= 高優先   ", options: { fontSize: 9, color: C.gray } },
  { text: "B ", options: { bold: true, color: C.gold, fontSize: 10 } },
  { text: "= 中優先   ", options: { fontSize: 9, color: C.gray } },
  { text: "C ", options: { bold: true, color: C.grayLight, fontSize: 10 } },
  { text: "= 低優先", options: { fontSize: 9, color: C.gray } },
], { x: 0.5, y: 4.8, w: 9, h: 0.3, fontFace: "Arial", margin: 0 });

addFooter(s6, 6);

// ==========================================
// SLIDE 7: IMPROVEMENT PHASE 1
// ==========================================
let s7 = pres.addSlide();
s7.background = { color: C.offWhite };
addContentHeader(s7, "改善施策 Phase 1：即座対応", "Immediate Actions (Today ~ 1 Month)");

s7.addText("投資不要・オペレーション改善で対応可能", { x: 0.6, y: 0.95, w: 8, h: 0.3, fontSize: 10, fontFace: "Arial", color: C.blue, italic: true, margin: 0 });

// TODO: CUSTOMIZE - Phase1カード（ファイル上部のPHASE1_CARDSで設定）
PHASE1_CARDS.forEach((p, i) => {
  const col = i % 2;
  const row = Math.floor(i / 2);
  const x = 0.5 + col * 4.6;
  const y = 1.35 + row * 1.95;

  s7.addShape(pres.shapes.RECTANGLE, { x, y, w: 4.35, h: 1.75, fill: { color: C.white }, shadow: shadow() });
  s7.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.06, h: 1.75, fill: { color: C.blue } });

  s7.addText(p.title, { x: x + 0.15, y: y + 0.05, w: 4, h: 0.35, fontSize: 12, fontFace: "Arial", color: C.navy, bold: true, margin: 0 });

  const bulletItems = p.items.map((item, idx) => ({
    text: item,
    options: { bullet: true, fontSize: 9, color: C.grayDark, breakLine: idx < p.items.length - 1 }
  }));
  s7.addText(bulletItems, { x: x + 0.15, y: y + 0.4, w: 4, h: 1.3, fontFace: "Arial", margin: 0, paraSpaceAfter: 3 });
});

addFooter(s7, 7);

// ==========================================
// SLIDE 8: IMPROVEMENT PHASE 2 & 3
// ==========================================
let s8 = pres.addSlide();
s8.background = { color: C.offWhite };
addContentHeader(s8, "改善施策 Phase 2・3", "Short-term & Mid-term Actions");

// Phase 2
s8.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.1, w: 4.35, h: 3.8, fill: { color: C.white }, shadow: shadow() });
s8.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.1, w: 4.35, h: 0.4, fill: { color: C.blue } });
s8.addText("Phase 2：短期施策（1〜3ヶ月）", { x: 0.65, y: 1.1, w: 4, h: 0.4, fontSize: 12, fontFace: "Arial", color: C.white, bold: true, margin: 0, valign: "middle" });

// TODO: CUSTOMIZE - Phase2アイテム（ファイル上部のPHASE2_ITEMSで設定）
let p2y = 1.6;
PHASE2_ITEMS.forEach((p) => {
  s8.addText(p.title, { x: 0.7, y: p2y, w: 3.8, h: 0.3, fontSize: 10, fontFace: "Arial", color: C.navy, bold: true, margin: 0 });
  const items = p.items.map((item, idx) => ({
    text: item, options: { bullet: true, fontSize: 9, color: C.gray, breakLine: idx < p.items.length - 1 }
  }));
  s8.addText(items, { x: 0.7, y: p2y + 0.28, w: 3.8, h: p.items.length * 0.22 + 0.1, fontFace: "Arial", margin: 0, paraSpaceAfter: 2 });
  p2y += 0.28 + p.items.length * 0.22 + 0.2;
});

// Phase 3
s8.addShape(pres.shapes.RECTANGLE, { x: 5.15, y: 1.1, w: 4.35, h: 3.8, fill: { color: C.white }, shadow: shadow() });
s8.addShape(pres.shapes.RECTANGLE, { x: 5.15, y: 1.1, w: 4.35, h: 0.4, fill: { color: C.navy } });
s8.addText("Phase 3：中期施策（3〜6ヶ月）", { x: 5.3, y: 1.1, w: 4, h: 0.4, fontSize: 12, fontFace: "Arial", color: C.white, bold: true, margin: 0, valign: "middle" });

// TODO: CUSTOMIZE - Phase3アイテム（ファイル上部のPHASE3_ITEMSで設定）
let p3y = 1.6;
PHASE3_ITEMS.forEach((p) => {
  s8.addText(p.title, { x: 5.35, y: p3y, w: 3.8, h: 0.3, fontSize: 10, fontFace: "Arial", color: C.navy, bold: true, margin: 0 });
  const items = p.items.map((item, idx) => ({
    text: item, options: { bullet: true, fontSize: 9, color: C.gray, breakLine: idx < p.items.length - 1 }
  }));
  s8.addText(items, { x: 5.35, y: p3y + 0.28, w: 3.8, h: p.items.length * 0.22 + 0.1, fontFace: "Arial", margin: 0, paraSpaceAfter: 2 });
  p3y += 0.28 + p.items.length * 0.22 + 0.2;
});

addFooter(s8, 8);

// ==========================================
// SLIDE 9: KPI TARGETS
// ==========================================
let s9 = pres.addSlide();
s9.background = { color: C.offWhite };
addContentHeader(s9, "KPI目標設定", "Key Performance Indicators");

const kpiHeader = [
  { text: "KPI項目", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 10, fontFace: "Arial" } },
  { text: "現状値", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 10, fontFace: "Arial", align: "center" } },
  { text: "目標値", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 10, fontFace: "Arial", align: "center" } },
  { text: "期限", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 10, fontFace: "Arial", align: "center" } },
];

// TODO: CUSTOMIZE - KPIデータ行（ファイル上部のKPI_TARGET_ROWSで設定）
const kpiData = KPI_TARGET_ROWS.map((r, i) => [
  { text: r[0], options: { fontSize: 10, fontFace: "Arial", bold: true, color: C.grayDark, fill: { color: i % 2 === 0 ? C.ice : C.white } } },
  { text: r[1], options: { fontSize: 10, fontFace: "Arial", align: "center", color: C.gray, fill: { color: i % 2 === 0 ? C.ice : C.white } } },
  { text: r[2], options: { fontSize: 10, fontFace: "Arial", align: "center", bold: true, color: C.green, fill: { color: i % 2 === 0 ? C.ice : C.white } } },
  { text: r[3], options: { fontSize: 10, fontFace: "Arial", align: "center", color: C.gray, fill: { color: i % 2 === 0 ? C.ice : C.white } } },
]);

s9.addTable([kpiHeader, ...kpiData], {
  x: 0.5, y: 1.2, w: 9, h: 0.1,
  colW: [2.8, 2.0, 2.2, 2.0],
  border: { pt: 0.5, color: "CBD5E1" },
  rowH: [0.35, 0.35, 0.35, 0.35, 0.35, 0.35, 0.35, 0.35],
});

// TODO: CUSTOMIZE - ノートボックス（ファイル上部のKPI_NOTE_*で設定）
s9.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.25, w: 9, h: 0.7, fill: { color: C.white }, shadow: shadow() });
s9.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.25, w: 0.06, h: 0.7, fill: { color: C.blue } });
s9.addText([
  { text: KPI_NOTE_BOLD, options: { bold: true, color: C.navy } },
  { text: KPI_NOTE_TEXT, options: { color: C.gray } },
], { x: 0.7, y: 4.25, w: 8.6, h: 0.7, fontSize: 10, fontFace: "Arial", valign: "middle", margin: 0 });

addFooter(s9, 9);

// ==========================================
// SLIDE 10: CLOSING
// ==========================================
let s10 = pres.addSlide();
s10.background = { color: C.navy };

s10.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: C.navyLight, transparency: 40 } });

s10.addText("総括", { x: 0.8, y: 0.8, w: 8, h: 0.5, fontSize: 14, fontFace: "Arial", color: C.blueLight, charSpacing: 6, margin: 0 });

s10.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 1.5, w: 8.4, h: 2.8, fill: { color: C.white, transparency: 90 } });

// TODO: CUSTOMIZE - 総括テキスト（ファイル上部のCLOSING_PARAGRAPHSで設定）
s10.addText(CLOSING_PARAGRAPHS, { x: 1.0, y: 1.6, w: 8, h: 2.5, fontFace: "Arial", color: C.white, margin: 0, paraSpaceAfter: 4 });

s10.addShape(pres.shapes.LINE, { x: 0.8, y: 4.5, w: 2, h: 0, line: { color: C.gold, width: 2 } });
s10.addText("ご清聴ありがとうございました", { x: 0.8, y: 4.6, w: 8, h: 0.4, fontSize: 14, fontFace: "Arial", color: C.grayLight, margin: 0 });

// === SAVE ===
// TODO: CUSTOMIZE - 出力ファイル名
const outPath = OUTPUT_DIR + "/" + HOTEL_NAME + "_口コミ分析レポート.pptx";
pres.writeFile({ fileName: outPath }).then(() => {
  console.log("PPTX created: " + outPath);
}).catch(err => console.error(err));
