const pptxgen = require("pptxgenjs");
const pres = new pptxgen();

// ============================================================
// ホテル日航川崎 固有設定
// ============================================================
const HOTEL_NAME = "ホテル日航川崎";
const ANALYSIS_PERIOD = "2026年2月〜3月";
const REVIEW_COUNT = "130件（6サイト）";
const REPORT_DATE = "2026年3月28日";
const OUTPUT_DIR = "納品レポート/ホテル別レポート";

// Slide 2: KPI
const KPI_AVG = "8.68";
const KPI_HIGH_RATE = "81.5%";
const KPI_LOW_RATE = "1.5%";
const KPI_TOTAL_COUNT = "130件";

// Slide 2: 強み
const SUMMARY_STRENGTHS = [
  { title: "立地・アクセス", desc: "53件言及。JR川崎駅直結・羽田30分・東京20分の圧倒的立地優位性" },
  { title: "英語対応スタッフ", desc: "20件言及。英語を話すスタッフへの高評価がインバウンド需要を支える" },
  { title: "朝食の充実", desc: "23件言及。バラエティ豊富なビュッフェ・割引クーポン対応が好評" },
];

// Slide 2: 弱み
const SUMMARY_WEAKNESSES = [
  { title: "エレベーター待ち", desc: "6件言及。混雑時間帯のエレベーター待ちがゲスト体験を損なう最大課題" },
  { title: "防音・騒音問題", desc: "4件言及。電車音・改装工事騒音が騒音敏感ゲストに影響" },
  { title: "設備の老朽化", desc: "3件言及。バスルーム設備の古さ・排水不良が指摘される" },
];

// Slide 3: サイト別
const SITE_CHART_LABELS = ["Trip.com", "じゃらん", "楽天トラベル", "Booking.com", "Google", "Agoda"];
const SITE_CHART_VALUES = [9.63, 9.00, 9.00, 8.65, 7.16, 7.00];

const SITE_INSIGHT_TEXTS = [
  { text: "サイト別インサイト", options: { bold: true, fontSize: 12, breakLine: true } },
  { text: "", options: { fontSize: 6, breakLine: true } },
  { text: "Trip.com・じゃらん・楽天が優秀評価", options: { bold: true, color: "16A34A", breakLine: true } },
  { text: "Trip.com 9.63点・じゃらん/楽天9.00点と国内外3サイトで優秀判定", options: { fontSize: 9, color: "64748B", breakLine: true } },
  { text: "", options: { fontSize: 6, breakLine: true } },
  { text: "Google・Agodaが低め（7.16/7.00点）", options: { bold: true, color: "EA580C", breakLine: true } },
  { text: "「概ね良好」判定。Booking.com（66件・最多）が8.65点で全体を支える構造", options: { fontSize: 9, color: "64748B", breakLine: true } },
  { text: "", options: { fontSize: 6, breakLine: true } },
  { text: "注：", options: { bold: true, fontSize: 9, color: "64748B" } },
  { text: "国内5点満点サイト（じゃらん/楽天/Google）は×2で10点換算", options: { fontSize: 9, color: "64748B" } },
];

const SITE_TABLE_ROWS_DATA = [
  ["Trip.com", "30", "9.63", "/10", "9.63"],
  ["じゃらん", "8", "4.50", "/5", "9.00"],
  ["楽天トラベル", "6", "4.50", "/5", "9.00"],
  ["Booking.com", "66", "8.65", "/10", "8.65"],
  ["Google", "19", "3.58", "/5", "7.16"],
  ["Agoda", "1", "7.00", "/10", "7.00"],
];

// Slide 4: 評価分布
const DIST_DOUGHNUT_LABELS = ["高評価(8-10)", "中評価(5-7)", "低評価(1-4)"];
const DIST_DOUGHNUT_VALUES = [106, 22, 2];

const DIST_BAR_LABELS = ["10点", "9点", "8点", "7点", "6点", "5点以下"];
const DIST_BAR_VALUES = [56, 24, 26, 10, 11, 3];

const DIST_SUMMARY_CARDS = [
  { label: "高評価 8-10点", val: "106件（81.5%）", col: "16A34A", bg: "DCFCE7" },
  { label: "中評価 5-7点", val: "22件（16.9%）", col: "D4A843", bg: "FFF7ED" },
  { label: "低評価 1-4点", val: "2件（1.5%）", col: "16A34A", bg: "DCFCE7" },
];

// Slide 5: 強み分析
const STRENGTHS_CARDS = [
  { theme: "立地・アクセス", count: "53件", desc: "JR川崎駅直結・羽田30分・東京20分の圧倒的アクセス優位性", quote: "「このホテルは一流のサービスを提供しており、JR駅のすぐそばに位置し、東京駅まで徒歩わずか20分、羽田空港まで電車で30分」" },
  { theme: "客室・設備", count: "46件", desc: "清潔な客室・カプセルコーヒーマシン・折り紙等のユニークな演出", quote: "「折り紙などの部屋の飾りには驚いて、カプセルコーヒーマシンは紅茶まで淹れることができます！」" },
  { theme: "朝食", count: "23件", desc: "バラエティ豊富なビュッフェ・割引クーポン・アイスクリーム等が好評", quote: "「朝食を1日追加で購入したのですが、割引クーポンをもらえました。アイスクリームなど、種類も豊富でした！」" },
  { theme: "スタッフ対応", count: "20件", desc: "英語対応力が高く、フライト遅延時の迅速対応等が評価される", quote: "「スタッフは礼儀正しく、ほとんどのスタッフが英語を話します」" },
  { theme: "清潔感", count: "22件", desc: "清潔に保たれた客室がゲストの基本的な安心感を提供", quote: "「部屋もとても清潔で、気持ちよく過ごせました」" },
  { theme: "コスパ・リピート", count: "8件", desc: "都心比較でのリーズナブルな宿泊費と好立地の組み合わせを評価", quote: "「宿泊料金は都心部よりもずっと安いです。川崎は東京の中心部からそれほど遠くなく、宿泊費を大幅に節約できます」" },
];

// Slide 6: 弱み
const WEAKNESS_ITEMS = [
  { pri: "S", cat: "エレベーター待ち", detail: "6件言及。混雑時間帯のエレベーター待ちがゲスト体験に直接影響。オペレーション改善で対応可能", count: "6件", color: "DC2626" },
  { pri: "A", cat: "防音・騒音問題", detail: "電車音（高層階）・改装工事騒音（4件）。騒音敏感ゲストへの事前通知が必要", count: "4件", color: "EA580C" },
  { pri: "A", cat: "設備老朽化（バスルーム）", detail: "バスルーム設備の古さ・排水不良（3件）。施設年数に応じた計画的更新が必要", count: "3件", color: "EA580C" },
  { pri: "B", cat: "Google評価の低さ", detail: "7.16点（概ね良好）で他サイト比較で2点以上の乖離。国内ゲストの不満が集中", count: "19件・7.16点", color: "D4A843" },
  { pri: "B", cat: "アメニティ", detail: "一部アメニティへの充実度向上への期待", count: "3件", color: "D4A843" },
  { pri: "C", cat: "個人情報管理", detail: "フロント業務での個人情報管理に関する指摘（1件）。情報管理強化が必要", count: "1件", color: "94A3B8" },
];

// Slide 7: Phase 1
const PHASE1_CARDS = [
  { title: "エレベーター混雑緩和", items: ["チェックイン・アウト時間分散化（時間帯別案内）", "混雑時間帯のスタッフ配置強化", "階段使用促進案内（低層階向け）", "エレベーター混雑情報の可視化（掲示板）"] },
  { title: "防音・騒音の事前案内", items: ["予約確認メールに騒音リスクを明記", "静音客室への案内（中層階・低層）", "改装工事期間・時間帯の透明な情報開示", "騒音敏感ゲスト向け客室選択サポート"] },
  { title: "Google口コミ対策強化", items: ["全Google口コミへ48h以内返信", "満足ゲストへのGoogle口コミ投稿案内", "低評価への真摯な公開返信", "Google評価の月次モニタリング"] },
  { title: "個人情報管理の強化", items: ["フロント業務での情報管理マニュアル更新", "スタッフへの個人情報保護教育実施", "情報管理チェックリストの定着化"] },
];

// Slide 8: Phase 2 & 3
const PHASE2_ITEMS = [
  { title: "バスルーム設備の更新", items: ["排水設備の点検・改修", "シャワーヘッド・水回り設備の選択的更新", "老朽化が目立つ設備の優先交換"] },
  { title: "朝食品質のさらなる向上", items: ["ビュッフェメニューの多様性強化", "インバウンド向けメニュー追加", "朝食時間帯の混雑緩和策"] },
  { title: "アメニティ充実", items: ["基本アメニティのグレードアップ", "長期滞在ゲスト向けアメニティ追加提供"] },
];

const PHASE3_ITEMS = [
  { title: "客室・設備の全面刷新", items: ["老朽化客室設備の計画的全面更新", "防音性能の向上（電車音対策）", "バスルームの全面リノベーション"] },
  { title: "エレベーター改善", items: ["エレベーター増設・速度向上検討", "1階改装完了後のロビー・フロントのグレードアップ", "スマートエレベーター（待ち時間表示）の導入"] },
  { title: "サービス差別化強化", items: ["日航グループブランドを活かしたプレミアム体験強化", "インバウンド需要を活かした多言語サービス拡充", "川崎立地を活かしたビジネス特化プラン開発"] },
];

// Slide 9: KPI
const KPI_TARGET_ROWS = [
  ["全体平均(10pt換算)", "8.68点", "8.9点以上", "2026年9月"],
  ["高評価率（8-10点）", "81.5%", "85%以上", "2026年9月"],
  ["低評価率（1-4点）", "1.5%", "0.5%以下", "2026年9月"],
  ["Trip.com平均評価", "9.63/10点", "9.7点以上", "2026年9月"],
  ["Google平均評価", "3.58/5点", "3.8/5点以上", "2026年9月"],
  ["口コミ返信率", "未計測", "100%（48h以内）", "2026年5月"],
  ["エレベータークレーム", "6件/2ヶ月", "2件以下/2ヶ月", "2026年6月"],
];

const KPI_NOTE_BOLD = "モニタリング方針：";
const KPI_NOTE_TEXT = "毎月末に全サイトのレビューデータを集計し、エレベーター・騒音クレームの件数推移を重点モニタリング。Google評価の底上げを最重要KPIとし、月次で改善進捗を確認する。1階改装完了後の評価変動も特に注視すること。";

// Slide 10: 総括
const CLOSING_PARAGRAPHS = [
  { text: "ホテル日航川崎は全体平均8.68点・高評価率81.5%と高水準の評価を誇ります。JR川崎駅直結という圧倒的立地優位性（53件）と英語対応スタッフの質の高さ（20件）が、国内外ゲストからの安定した支持を生み出す二本柱となっています。", options: { breakLine: true, fontSize: 12 } },
  { text: "", options: { fontSize: 6, breakLine: true } },
  { text: "Trip.com（9.63点）・じゃらん・楽天トラベル（9.00点）での優秀評価は日航グループとしてのブランド力を示す一方、Google評価（7.16点）の低さとエレベーター待ち（6件）・防音問題（4件）・設備老朽化（3件）が中評価層形成の主因です。", options: { breakLine: true, fontSize: 12 } },
  { text: "", options: { fontSize: 6, breakLine: true } },
  { text: "エレベーター混雑緩和と防音の事前案内強化はオペレーション改善で対応可能な即効性のある施策です。バスルーム設備の更新は中期計画として推進し、1階改装完了と合わせて全体的なグレードアップを図ることで評価のさらなる向上が期待されます。", options: { breakLine: true, fontSize: 12 } },
  { text: "", options: { fontSize: 6, breakLine: true } },
  { text: "川崎という都市のポテンシャルと日航グループの品質基準を融合させ、全体平均8.9点以上・高評価率85%以上を目指してください。Google評価3.8/5点以上への引き上げが最重要KPIです。", options: { bold: true, fontSize: 12, color: "D4A843" } },
];

// ============================================================
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
