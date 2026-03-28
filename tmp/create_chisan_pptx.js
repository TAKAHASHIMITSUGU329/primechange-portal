const pptxgen = require("pptxgenjs");
const pres = new pptxgen();

const HOTEL_NAME = "チサンホテル浜松町";
const ANALYSIS_PERIOD = "2026年2月〜3月";
const REVIEW_COUNT = "47件（5サイト）";
const REPORT_DATE = "2026年3月28日";
const OUTPUT_DIR = "納品レポート/ホテル別レポート";

const KPI_AVG = "7.15";
const KPI_HIGH_RATE = "59.6%";
const KPI_LOW_RATE = "21.3%";
const KPI_TOTAL_COUNT = "47件";

const SUMMARY_STRENGTHS = [
  { title: "立地・アクセス", desc: "浜松町・大門・田町・日の出（ゆりかもめ）の複数路線が徒歩圏内（18件）" },
  { title: "コスパの良さ", desc: "老朽化を補うほどの価格魅力。価格帯に見合った価値を提供（8件）" },
  { title: "朝食の質", desc: "彩りと栄養バランスの良い朝食が一部ゲストから高評価（7件）" },
];

const SUMMARY_WEAKNESSES = [
  { title: "設備の老朽化", desc: "浴槽の不具合・廊下の劣化・ハウスダスト問題など設備劣化の指摘が多数（10件）" },
  { title: "騒音・防音の問題", desc: "線路側客室での電車騒音が睡眠を妨げるレベルで言及（8件）" },
  { title: "アクセスの分かりにくさ", desc: "駅から遠い・わかりにくいとの指摘。案内改善が急務（3件）" },
];

const SITE_CHART_LABELS = ["Agoda", "Booking.com", "楽天トラベル", "じゃらん", "Google"];
const SITE_CHART_VALUES = [8.67, 8.00, 7.06, 6.86, 6.17];

const SITE_INSIGHT_TEXTS = [
  { text: "サイト別パフォーマンス分析", options: { bold: true, fontSize: 12, breakLine: true } },
  { text: "", options: { fontSize: 6, breakLine: true } },
  { text: "海外OTA強み", options: { bold: true, color: "16A34A", breakLine: true } },
  { text: "Agoda（8.67点）とBooking.com（8.0点）では良好水準を確保", options: { fontSize: 9, color: "64748B", breakLine: true } },
  { text: "", options: { fontSize: 6, breakLine: true } },
  { text: "国内サイト要改善", options: { bold: true, color: "DC2626", breakLine: true } },
  { text: "Google（6.17点）・じゃらん（6.86点）が全体スコアを引き下げ", options: { fontSize: 9, color: "64748B", breakLine: true } },
  { text: "", options: { fontSize: 6, breakLine: true } },
  { text: "注：", options: { bold: true, fontSize: 9, color: "64748B" } },
  { text: "国内サイト（5点満点）は10点換算で比較", options: { fontSize: 9, color: "64748B" } },
];

const SITE_TABLE_ROWS_DATA = [
  ["Agoda", "9", "8.67", "/10", "8.67"],
  ["Booking.com", "2", "8.00", "/10", "8.00"],
  ["楽天トラベル", "17", "3.53", "/5", "7.06"],
  ["じゃらん", "7", "3.43", "/5", "6.86"],
  ["Google", "12", "3.08", "/5", "6.17"],
];

const DIST_DOUGHNUT_LABELS = ["高評価(8-10)", "中評価(5-7)", "低評価(1-4)"];
const DIST_DOUGHNUT_VALUES = [28, 9, 10];

const DIST_BAR_LABELS = ["10点", "9点", "8点", "6点", "4点", "2点"];
const DIST_BAR_VALUES = [11, 2, 15, 9, 7, 3];

const DIST_SUMMARY_CARDS = [
  { label: "高評価 8-10点", val: "28件（59.6%）", col: "16A34A", bg: "DCFCE7" },
  { label: "中評価 5-7点", val: "9件（19.1%）", col: "D4A843", bg: "FFF7ED" },
  { label: "低評価 1-4点", val: "10件（21.3%）", col: "DC2626", bg: "FEE2E2" },
];

const STRENGTHS_CARDS = [
  { theme: "立地・アクセス", count: "18件", desc: "複数路線・複数駅へのアクセスでお台場・東京駅・羽田へ便利", quote: "「浜松町・日の出両駅から近く東京駅にも乗り換えなしで行けた」" },
  { theme: "清潔感", count: "6件", desc: "老朽化した設備ながら清掃は行き届いており清潔感を維持", quote: "「部屋も設備に古さを感じるものの、綺麗にされており良かった」" },
  { theme: "スタッフ対応", count: "7件", desc: "チェックイン時間の柔軟な対応など親切な接客が評価される", quote: "「15時チェックインでしたが柔軟に対応いただけて助かりました」" },
  { theme: "客室・設備", count: "22件", desc: "シャワー水圧・無料ランドリーなど基本機能は十分な評価", quote: "「水量やシャワーもしっかり出て困りません」" },
  { theme: "朝食", count: "7件", desc: "彩りと栄養バランスの良い朝食が特に高評価を獲得", quote: "「朝食は彩りも栄養バランスも抜群で満腹に。コスパ最高」" },
  { theme: "コスパ・リピート", count: "8件", desc: "老朽化を補う価格魅力と立地の良さがコスパ感を創出", quote: "「コスパ最高だし、トレインビューで読書するのが気持ちよかった」" },
];

const WEAKNESS_ITEMS = [
  { pri: "S", cat: "設備の老朽化", detail: "浴槽の穴・ヒビ、廊下カーペットの劣化、ハウスダスト問題など設備劣化の具体的指摘多数", count: "10件", color: "DC2626" },
  { pri: "A", cat: "騒音・防音問題", detail: "「ボロアパート並みの防音性能」。電車の音が睡眠を妨げるレベルで線路側客室での問題深刻", count: "8件", color: "EA580C" },
  { pri: "A", cat: "アクセスの分かりにくさ", detail: "「地図がないとたどり着けない」。駅からの道順が不明瞭で特に初回利用者の不満が大きい", count: "3件", color: "EA580C" },
  { pri: "B", cat: "朝食クオリティ", detail: "「おかずが少ない・貧素」という朝食内容への指摘。内容と価格バランスへの不満", count: "2件", color: "D4A843" },
  { pri: "B", cat: "乾燥・加湿器不足", detail: "「老舗のビジネスホテルなのに加湿器ありません」という乾燥問題への不満", count: "1件", color: "D4A843" },
  { pri: "C", cat: "周辺飲食店の少なさ", detail: "「周辺にリーズナブルな飲食店がなく不便」という周辺環境への指摘", count: "1件", color: "94A3B8" },
];

const PHASE1_CARDS = [
  { title: "設備不具合の緊急対応", items: ["浴槽の穴・ヒビなど安全不備の即時修繕", "客室設備点検チェックリストの導入と月次実施", "廊下カーペットの張り替え・固定", "ハウスダスト対策として空気清浄機の設置"] },
  { title: "アクセス案内の改善", items: ["写真付き詳細道順マップをOTA・公式サイトに掲載", "各最寄駅からの案内看板の設置", "チェックイン前のアクセス案内メール自動送信"] },
  { title: "騒音への暫定対策", items: ["線路側客室のOTA掲載での事前告知", "耳栓・防音グッズのフロント無料貸出", "騒音懸念ゲストへの反線路側客室の優先アサイン"] },
  { title: "口コミ返信体制の整備", items: ["全サイト48時間以内返信ルール化", "設備改善計画をネガティブ口コミへの返信に明記", "Google・じゃらんへの重点返信対応"] },
];

const PHASE2_ITEMS = [
  { title: "客室防音対策", items: ["防音カーテン・二重窓の設置（線路側客室優先）", "ドア・窓の隙間テープによる防音強化", "加湿器の客室常備（乾燥対策）"] },
  { title: "客室・共用部の段階的リニューアル", items: ["壁紙・カーペットの張り替えによる清潔感向上", "浴室設備のアップグレード（シャワーヘッド・水栓）", "照明のLED化による明るさ・清潔感の向上"] },
  { title: "朝食メニューの改善", items: ["おかずの種類と質の向上", "季節メニューの定期的な見直し"] },
];

const PHASE3_ITEMS = [
  { title: "全客室の本格リノベーション", items: ["客室全体の内装一新（床・壁・浴室・家具）", "空調設備のアップグレード（温度・乾燥問題の解消）", "高性能空気清浄機の全室設置"] },
  { title: "外壁・構造の防音改修", items: ["線路側客室の外壁防音材の施工", "二重サッシへの全面的な交換"] },
  { title: "共用エリアのリノベーション", items: ["フロント・エントランスの近代化", "朝食スペースの拡充・リデザイン"] },
];

const KPI_TARGET_ROWS = [
  ["全体平均(10pt換算)", "7.15点", "7.8点以上", "2026年9月"],
  ["高評価率（8-10点）", "59.6%", "68%以上", "2026年9月"],
  ["低評価率（1-4点）", "21.3%", "10%以下", "2026年9月"],
  ["Google平均評価", "3.08/5点", "3.8/5点以上", "2026年9月"],
  ["じゃらん平均評価", "3.43/5点", "3.8/5点以上", "2026年9月"],
  ["口コミ返信率", "未計測", "100%（48h以内）", "2026年6月"],
  ["設備クレーム件数", "10件/2ヶ月", "3件以下/2ヶ月", "2026年9月"],
];

const KPI_NOTE_BOLD = "モニタリング方針：";
const KPI_NOTE_TEXT = "毎月末に全サイトの口コミを収集・集計。設備クレーム件数とGoogleスコアの改善状況を重点モニタリング。リノベーション前後の評価変化を追跡します。";

const CLOSING_PARAGRAPHS = [
  { text: "チサンホテル浜松町は複数路線・複数駅へのアクセスという強固な立地優位性を持ちながら、設備の老朽化（10件）と騒音問題（8件）が全体平均7.15点・低評価率21.3%という数字に反映されています。", options: { breakLine: true, fontSize: 12 } },
  { text: "", options: { fontSize: 6, breakLine: true } },
  { text: "最優先すべきはPhase 1の設備不具合緊急修繕とアクセス案内改善です。浴槽の穴・ヒビなど安全に関わる不具合の即時修繕と、写真付き道順マップのOTA掲載は追加投資なしに実施でき、Googleスコアの改善に直結します。", options: { breakLine: true, fontSize: 12 } },
  { text: "", options: { fontSize: 6, breakLine: true } },
  { text: "Phase 2の防音対策（防音カーテン・二重窓）はホテルの評価を最も大きく引き下げている騒音問題への直接的な対処であり、投資優先度が高い施策です。Phase 3のリノベーションは抜本的な評価改善をもたらす長期的投資として計画的に進めてください。", options: { breakLine: true, fontSize: 12 } },
  { text: "", options: { fontSize: 6, breakLine: true } },
  { text: "段階的な改善を着実に実行することで、2026年9月末までに全体平均7.8点・低評価率10%以下の達成を目指しましょう。", options: { bold: true, fontSize: 12, color: "D4A843" } },
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
