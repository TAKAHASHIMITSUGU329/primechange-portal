const pptxgen = require("pptxgenjs");
const pres = new pptxgen();

// ============================================================
// ホテル固有の設定
// ============================================================
const HOTEL_NAME = "第一ホテル池袋";
const ANALYSIS_PERIOD = "2026年2月〜3月";
const REVIEW_COUNT = "130件（6サイト）";
const REPORT_DATE = "2026年3月28日";
const OUTPUT_DIR = "納品レポート/ホテル別レポート";

// ============================================================
// Slide 2: エグゼクティブサマリー KPI
// ============================================================
const KPI_AVG = "9.18";
const KPI_HIGH_RATE = "91.5%";
const KPI_LOW_RATE = "2.3%";
const KPI_TOTAL_COUNT = "130件";

const SUMMARY_STRENGTHS = [
  { title: "立地・アクセス（最大の強み）", desc: "84件（64.6%）言及。池袋駅東口徒歩数分。JR・地下鉄・空港バスへの完璧なアクセス。" },
  { title: "スタッフ対応・サービス", desc: "32件言及。親切・フレンドリーな対応と柔軟なレイトチェックアウト対応が高評価。" },
  { title: "清潔感・客室品質", desc: "22件言及。「古い建物だが清潔」「2つ星を超える3つ星クオリティ」と複数が高評価。" },
];

const SUMMARY_WEAKNESSES = [
  { title: "Agoda評価の低さ", desc: "5.5点（要改善）。全サイト中唯一の低評価サイトであり最優先改善課題。" },
  { title: "客室の狭さへの指摘", desc: "5件で客室の狭さに言及。スーツケース2個が同時に開けられないとの声。" },
  { title: "バスルーム排水の問題", desc: "シャワー使用時に排水が追いつかない報告。設備改修が必要。" },
];

// ============================================================
// Slide 3: サイト別評価データ
// ============================================================
const SITE_CHART_LABELS = ["Trip.com", "Google", "Booking.com", "楽天トラベル", "じゃらん", "Agoda"];
const SITE_CHART_VALUES = [9.56, 9.21, 8.82, 8.50, 8.00, 5.50];

const SITE_INSIGHT_TEXTS = [
  { text: "サイト別評価インサイト", options: { bold: true, fontSize: 12, breakLine: true } },
  { text: "", options: { fontSize: 6, breakLine: true } },
  { text: "主要5サイトで優秀〜良好水準", options: { bold: true, color: "16A34A", breakLine: true } },
  { text: "Trip.com(9.56)・Google(9.21)・Booking.com(8.82)など5サイトで高評価を達成。130件中126件が良好以上。", options: { fontSize: 9, color: "64748B", breakLine: true } },
  { text: "", options: { fontSize: 6, breakLine: true } },
  { text: "Agodaのみ要改善水準", options: { bold: true, color: "DC2626", breakLine: true } },
  { text: "Agoda(5.5点)が唯一の要改善サイト。4件と件数は少ないが、全体評価の底上げのために早急な対応が必要。", options: { fontSize: 9, color: "64748B", breakLine: true } },
  { text: "", options: { fontSize: 6, breakLine: true } },
  { text: "注：", options: { bold: true, fontSize: 9, color: "64748B" } },
  { text: "国内サイト(5点満点)は×2で10点換算", options: { fontSize: 9, color: "64748B" } },
];

const SITE_TABLE_ROWS_DATA = [
  ["Trip.com", "66", "9.56", "/10", "9.56"],
  ["Google", "33", "4.61", "/5", "9.21"],
  ["Booking.com", "22", "8.82", "/10", "8.82"],
  ["楽天トラベル", "4", "4.25", "/5", "8.50"],
  ["じゃらん", "1", "4.00", "/5", "8.00"],
  ["Agoda", "4", "5.50", "/10", "5.50"],
];

// ============================================================
// Slide 4: 評価分布データ
// ============================================================
const DIST_DOUGHNUT_LABELS = ["高評価(8-10)", "中評価(5-7)", "低評価(1-4)"];
const DIST_DOUGHNUT_VALUES = [119, 8, 3];

const DIST_BAR_LABELS = ["10点", "9点", "8点", "7点", "6点", "5点", "4点", "2点"];
const DIST_BAR_VALUES = [82, 19, 18, 3, 2, 3, 2, 1];

const DIST_SUMMARY_CARDS = [
  { label: "高評価 8-10点", val: "119件（91.5%）", col: "16A34A", bg: "DCFCE7" },
  { label: "中評価 5-7点", val: "8件（6.2%）", col: "16A34A", bg: "DCFCE7" },
  { label: "低評価 1-4点", val: "3件（2.3%）", col: "16A34A", bg: "DCFCE7" },
];

// ============================================================
// Slide 5: 強み分析（6項目）
// ============================================================
const STRENGTHS_CARDS = [
  { theme: "立地・アクセス", count: "84件", desc: "池袋駅東口徒歩数分。JR・地下鉄・空港バス完備。百貨店・飲食店が周辺に集結。64.6%が言及するNo.1の強み。", quote: "「抜群の立地。パルコ・西武・東武も近く最高に便利」" },
  { theme: "客室・設備", count: "48件", desc: "トリプルルームなど多様な客室タイプ。清潔で快適な空間が「2〜3つ星クオリティ」と評価。", quote: "「このホテルは2つ星ではなく3つ星の評価に値する」" },
  { theme: "スタッフ対応", count: "32件", desc: "親切・フレンドリーなスタッフと柔軟なサービス（レイトチェックアウト等）が高評価。", quote: "「スタッフは親切でフレンドリー。レイトチェックアウトも対応可」" },
  { theme: "清潔感", count: "22件", desc: "建物の古さを感じさせない清潔さが継続的に評価。清掃クオリティが高評価率の基盤。", quote: "「古い建物だが清潔。スタッフも親切」" },
  { theme: "朝食", count: "6件", desc: "和洋折衷の朝食メニューが「とても美味しい」と評価。一部はやや高めとの声も。", quote: "「朝食がとても美味しかった」「伝統的な日本料理と洋食で味もかなり美味しい」" },
  { theme: "コスパ・リピート", count: "8件", desc: "立地・サービス対比での高いコスパ評価。リピーター率が高く継続的な支持を獲得。", quote: "「リピート利用。立地は非常に便利で再利用したい」" },
];

// ============================================================
// Slide 6: 弱み優先度マトリクス
// ============================================================
const WEAKNESS_ITEMS = [
  { pri: "S", cat: "Agoda評価の低さ", detail: "全サイト中唯一の要改善水準（5.5点）。掲載情報と実態のギャップ・海外ゲストの期待値差が要因。", count: "4件", color: "DC2626" },
  { pri: "A", cat: "客室の狭さへの指摘", detail: "スーツケース2個が同時に開けられない等。部屋の狭さが一定数のゲストの期待値を下回る。", count: "5件", color: "EA580C" },
  { pri: "B", cat: "バスルーム排水の問題", detail: "シャワー使用時に排水が追いつかずくるぶしまで浸かる状態になるとの報告。", count: "2件", color: "D4A843" },
  { pri: "B", cat: "シャワー起動音の問題", detail: "シャワーを出すと大きなビー音が鳴るとの報告。早急な確認・修繕が必要。", count: "1件", color: "D4A843" },
  { pri: "B", cat: "アメニティの取り置き方式", detail: "一部アメニティをロビーに取りに行く必要があり不便との声。全客室備え付けへの改善が望ましい。", count: "1件", color: "D4A843" },
  { pri: "C", cat: "窓からの騒音", detail: "幹線道路近辺の立地柄、騒音が気になるケースあり。防音対策の検討が必要。", count: "1件", color: "94A3B8" },
];

// ============================================================
// Slide 7: Phase 1 改善施策（4項目）
// ============================================================
const PHASE1_CARDS = [
  { title: "Agoda評価の緊急改善", items: ["掲載情報（写真・説明文）の精度向上", "客室サイズ等の期待値を正確に記載", "過去低評価コメントへの丁寧な返信", "Agoda経由ゲストへの特別ウェルカム対応"] },
  { title: "高評価維持体制の強化", items: ["Trip.com・Googleの高評価維持モニタリング", "全レビューへの丁寧な返信を継続", "スタッフサービス品質の継続的教育・評価", "ゲスト満足度の定期アンケート実施"] },
  { title: "低層階・騒音・設備対策", items: ["低層階客室への油煙・騒音リスクの事前告知", "希望ゲストへの高層階優先アサイン対応", "排水・シャワー音問題の即時点検・修繕", "アメニティの全室完備化（ロビー取り置き廃止）"] },
  { title: "口コミ件数の拡大", items: ["チェックアウト時のQRコード口コミ依頼", "Agoda・楽天・じゃらんの口コミ件数増加", "リピーター向け特典で口コミ促進"] },
];

// ============================================================
// Slide 8: Phase 2 & 3 改善施策
// ============================================================
const PHASE2_ITEMS = [
  { title: "Agoda向けページの全面最適化", items: ["英語・中国語・韓国語対応ページの品質向上", "海外ゲスト向けサービス情報の多言語整備", "Agoda経由ゲストのフィードバック専用収集"] },
  { title: "バスルーム設備の修繕", items: ["排水設備の改修でシャワー時の水たまり解消", "シャワー起動音センサーの見直し・修理", "バスルーム全体の点検と老朽箇所の修繕"] },
  { title: "朝食のさらなる充実", items: ["和洋折衷メニューの季節対応と品質向上", "朝食満足度アンケートによる継続的改善", "朝食価格の見直し（一部からのやや高いとの声への対応）"] },
];

const PHASE3_ITEMS = [
  { title: "全サイト優秀水準の達成", items: ["Agoda 8.0点以上・Booking.com 9.0点以上の達成", "全サイト平均9.3点以上の実現", "口コミ総数160件/2ヶ月以上への拡大"] },
  { title: "バスルームの本格リノベーション", items: ["排水設備の全面改修", "防音性能の強化（窓・ドアの改修）"] },
  { title: "客室の快適性向上", items: ["収納スペースの拡充（大型荷物対応）", "コンセント配置の改善（業務・充電利用者向け）", "独立エアコン・換気設備の維持・強化"] },
];

// ============================================================
// Slide 9: KPI目標設定
// ============================================================
const KPI_TARGET_ROWS = [
  ["全体平均(10pt換算)", "9.18点", "9.3点以上", "2026年9月"],
  ["高評価率（8-10点）", "91.5%", "93%以上", "2026年9月"],
  ["低評価率（1-4点）", "2.3%", "1%以下", "2026年9月"],
  ["Agoda平均", "5.50/10点", "8.0/10点以上", "2026年6月"],
  ["Trip.com平均", "9.56/10点", "9.5点以上維持", "2026年9月"],
  ["口コミ総数", "130件/2ヶ月", "160件/2ヶ月以上", "2026年6月"],
  ["口コミ返信率", "未計測", "100%（48h以内）", "2026年4月"],
];

const KPI_NOTE_BOLD = "モニタリング方針：";
const KPI_NOTE_TEXT = "Agoda改善を最優先指標として月次モニタリング。Trip.com・Googleの高評価水準の維持を継続的にチェックし、全サイト優秀水準達成を2026年9月末を目標に取り組む。";

// ============================================================
// Slide 10: 総括テキスト
// ============================================================
const CLOSING_PARAGRAPHS = [
  { text: "第一ホテル池袋は、全体平均9.18点・高評価率91.5%という業界トップクラスの評価を誇り、池袋駅東口徒歩数分という絶対的な立地優位性と、建物の古さを超えるサービス品質・清潔感が相乗効果を生み出しています。「2つ星ではなく3つ星クオリティ」という複数ゲストの評価はこのホテルの真の価値を示しています。", options: { breakLine: true, fontSize: 12 } },
  { text: "", options: { fontSize: 6, breakLine: true } },
  { text: "唯一の要改善課題であるAgoda評価（5.5点）の底上げが最優先施策です。Agodaの掲載情報最適化・多言語対応強化により、全サイト優秀水準（8.0点以上）への引き上げが実現すれば、全体平均9.3点以上の達成も視野に入ります。", options: { breakLine: true, fontSize: 12 } },
  { text: "", options: { fontSize: 6, breakLine: true } },
  { text: "バスルーム排水の問題やアメニティの取り置き方式など、細部の改善施策は投資規模が小さく効果が直接的です。これらを短期間で対処することで、さらなる評価向上と口コミ件数拡大につながります。", options: { breakLine: true, fontSize: 12 } },
  { text: "", options: { fontSize: 6, breakLine: true } },
  { text: "業界最高水準の評価を誇る第一ホテル池袋が、Agoda改善と細部の磨き込みで全サイト優秀水準9.3点以上を達成し、池袋エリアのナンバーワンホテルとしての地位を確固たるものにしてください。", options: { bold: true, fontSize: 12, color: "D4A843" } },
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
