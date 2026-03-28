const pptxgen = require("pptxgenjs");
const pres = new pptxgen();

// ============================================================
// ホテル固有の設定
// ============================================================
const HOTEL_NAME = "コンフォートホテル横浜関内";
const ANALYSIS_PERIOD = "2026年2月〜3月";
const REVIEW_COUNT = "143件（6サイト）";
const REPORT_DATE = "2026年3月28日";
const OUTPUT_DIR = "納品レポート/ホテル別レポート";

// ============================================================
// Slide 2: エグゼクティブサマリー KPI
// ============================================================
const KPI_AVG = "7.52";
const KPI_HIGH_RATE = "67.1%";
const KPI_LOW_RATE = "12.6%";
const KPI_TOTAL_COUNT = "143件";

const SUMMARY_STRENGTHS = [
  { title: "朝食・食事", desc: "82件（57%）で言及。無料ビュッフェとカフェスペースが高評価。" },
  { title: "立地・アクセス", desc: "関内駅徒歩5分。横浜スタジアム・中華街・山下公園へ徒歩圏内。" },
  { title: "スタッフ対応", desc: "37件で丁寧・親切な接客を評価。フロント・清掃スタッフ全般。" },
];

const SUMMARY_WEAKNESSES = [
  { title: "朝食提供形式の変更", desc: "改装工事中のビュッフェ→弁当変更が期待値ギャップを招く。" },
  { title: "客室の狭さ・設備老朽化", desc: "ユニットバスの狭さ、エアコン老朽化が不満の主因。" },
  { title: "朝食会場の混雑", desc: "会場が狭く料理不足になるケースが複数報告。" },
];

// ============================================================
// Slide 3: サイト別評価データ
// ============================================================
const SITE_CHART_LABELS = ["Trip.com", "Agoda", "Google", "楽天トラベル", "じゃらん", "Booking.com"];
const SITE_CHART_VALUES = [8.32, 8.00, 7.58, 7.48, 7.26, 7.00];

const SITE_INSIGHT_TEXTS = [
  { text: "サイト別評価インサイト", options: { bold: true, fontSize: 12, breakLine: true } },
  { text: "", options: { fontSize: 6, breakLine: true } },
  { text: "海外OTAが高評価", options: { bold: true, color: "16A34A", breakLine: true } },
  { text: "Trip.com(8.32)・Agoda(8.0)は良好水準。海外ゲストからの評価が安定。", options: { fontSize: 9, color: "64748B", breakLine: true } },
  { text: "", options: { fontSize: 6, breakLine: true } },
  { text: "国内サイトは改善余地", options: { bold: true, color: "D4A843", breakLine: true } },
  { text: "Booking.com(7.0)・じゃらん(7.26)は工事期間の影響が大。工事完了後の回復に期待。", options: { fontSize: 9, color: "64748B", breakLine: true } },
  { text: "", options: { fontSize: 6, breakLine: true } },
  { text: "注：", options: { bold: true, fontSize: 9, color: "64748B" } },
  { text: "国内サイト(5点満点)は×2で10点換算", options: { fontSize: 9, color: "64748B" } },
];

const SITE_TABLE_ROWS_DATA = [
  ["Trip.com", "19", "8.32", "/10", "8.32"],
  ["Agoda", "11", "8.00", "/10", "8.00"],
  ["Google", "24", "3.79", "/5", "7.58"],
  ["楽天トラベル", "42", "3.74", "/5", "7.48"],
  ["じゃらん", "19", "3.63", "/5", "7.26"],
  ["Booking.com", "28", "7.00", "/10", "7.00"],
];

// ============================================================
// Slide 4: 評価分布データ
// ============================================================
const DIST_DOUGHNUT_LABELS = ["高評価(8-10)", "中評価(5-7)", "低評価(1-4)"];
const DIST_DOUGHNUT_VALUES = [96, 29, 18];

const DIST_BAR_LABELS = ["10点", "9点", "8点", "7点", "6点", "5点", "4点", "2点"];
const DIST_BAR_VALUES = [34, 8, 54, 6, 18, 5, 10, 7];

const DIST_SUMMARY_CARDS = [
  { label: "高評価 8-10点", val: "96件（67.1%）", col: "16A34A", bg: "DCFCE7" },
  { label: "中評価 5-7点", val: "29件（20.3%）", col: "D4A843", bg: "FFF7ED" },
  { label: "低評価 1-4点", val: "18件（12.6%）", col: "DC2626", bg: "FEE2E2" },
];

// ============================================================
// Slide 5: 強み分析（6項目）
// ============================================================
const STRENGTHS_CARDS = [
  { theme: "朝食・食事", count: "82件", desc: "無料ビュッフェ・カフェスペースが好評。改装後の本格復活に期待。", quote: "「朝食も期待以上。また利用したい」" },
  { theme: "客室・設備", count: "57件", desc: "トリプルルームなど多様な客室タイプ。清潔感も維持。", quote: "「部屋は清潔で快適」" },
  { theme: "立地・アクセス", count: "45件", desc: "関内駅徒歩5分。横浜観光の拠点として最高の立地。", quote: "「最高のロケーション。駅まで5分」" },
  { theme: "スタッフ対応", count: "37件", desc: "丁寧・親切なスタッフ対応が宿泊者から継続的に高評価。", quote: "「スタッフの方みんな一生懸命でした」" },
  { theme: "清潔感", count: "26件", desc: "客室・共用部ともに清潔さが維持されており宿泊者から安心感。", quote: "「とても清潔な部屋でした」" },
  { theme: "コスパ・リピート", count: "15件", desc: "料金対比での満足度が高く、リピーター・連泊利用者の声あり。", quote: "「連泊プランでお得に泊まりました」" },
];

// ============================================================
// Slide 6: 弱み優先度マトリクス
// ============================================================
const WEAKNESS_ITEMS = [
  { pri: "S", cat: "朝食提供形式の変更", detail: "改装工事中のビュッフェ→弁当変更が予告なく実施。期待値ギャップで低評価多発。", count: "8件", color: "DC2626" },
  { pri: "A", cat: "朝食会場の混雑・料理不足", detail: "会場が狭く、特に団体客来訪時に料理不足。混雑時の動線も不明確。", count: "6件", color: "EA580C" },
  { pri: "A", cat: "客室の狭さ・設備老朽化", detail: "ユニットバス狭小、エアコン老朽化、カードキーホルダーなし等複合的課題。", count: "7件", color: "EA580C" },
  { pri: "B", cat: "スタッフ接客の一部不満", detail: "チェックイン時に急かす等の不丁寧な対応が一部報告。", count: "3件", color: "D4A843" },
  { pri: "B", cat: "エレベーター待ち", detail: "2台のみで朝夕混雑時に待ち時間発生。", count: "3件", color: "D4A843" },
  { pri: "C", cat: "シャワーヘッド水漏れ", detail: "複数客室でシャワーヘッド付近から水漏れ報告。早急な点検が必要。", count: "2件", color: "94A3B8" },
];

// ============================================================
// Slide 7: Phase 1 改善施策（4項目）
// ============================================================
const PHASE1_CARDS = [
  { title: "朝食工事期間の告知強化", items: ["予約確認メールに朝食形式を明記", "チェックイン時に口頭で状況説明", "工事完了日程を口コミ返信で周知", "弁当の内容・品質改善で満足度向上"] },
  { title: "朝食会場の混雑対策", items: ["時間帯別入場整理の実施", "料理補充頻度の増加と担当固定", "床サインで動線を明示", "団体客と個人客の入場分離"] },
  { title: "口コミへの積極的返信", items: ["低評価コメントへの48h以内返信", "改装完了・ビュッフェ復活の告知", "高評価への感謝返信でエンゲージメント向上"] },
  { title: "設備の緊急点検・修繕", items: ["シャワーヘッド水漏れ客室の特定・修理", "カードキーホルダーの全室設置", "タオルハンガー増設"] },
];

// ============================================================
// Slide 8: Phase 2 & 3 改善施策
// ============================================================
const PHASE2_ITEMS = [
  { title: "スタッフ接客研修", items: ["チェックイン・アウト対応マニュアル整備", "繁忙時対応マナー研修実施", "月次フィードバック共有会議の設置"] },
  { title: "朝食ビュッフェ完全復活", items: ["改装完了後の即時ビュッフェ再開", "季節メニュー追加でリピート訴求", "座席数増加と動線改善"] },
  { title: "OTAプロフィール最新化", items: ["改装完了後の写真・説明文の更新", "工事期間中の評価への返信で挽回"] },
];

const PHASE3_ITEMS = [
  { title: "客室設備の計画的更新", items: ["老朽化エアコンの段階的交換", "ユニットバスの改修または設備グレードアップ", "客室内ファシリティの現代化"] },
  { title: "朝食会場の拡充", items: ["座席数の物理的拡大", "調理設備の増強で提供スピード向上"] },
  { title: "プレミアム客室の創設", items: ["一部客室のリノベーションで単価向上", "高評価率95%以上の旗艦客室の確立"] },
];

// ============================================================
// Slide 9: KPI目標設定
// ============================================================
const KPI_TARGET_ROWS = [
  ["全体平均(10pt換算)", "7.52点", "8.0点以上", "2026年9月"],
  ["高評価率（8-10点）", "67.1%", "75%以上", "2026年9月"],
  ["低評価率（1-4点）", "12.6%", "5%以下", "2026年6月"],
  ["楽天トラベル平均", "3.74/5点", "4.0/5点以上", "2026年6月"],
  ["じゃらん平均", "3.63/5点", "4.0/5点以上", "2026年6月"],
  ["口コミ返信率", "未計測", "100%（48h以内）", "2026年4月"],
  ["朝食満足クレーム件数", "8件/2ヶ月", "2件以下/2ヶ月", "2026年6月"],
];

const KPI_NOTE_BOLD = "モニタリング方針：";
const KPI_NOTE_TEXT = "改装工事完了（2026年上期想定）を契機に朝食関連指標が大幅回復する見通し。月次でサイト別平均をモニタリングし、低評価急増時は72時間以内に対策を検討する。";

// ============================================================
// Slide 10: 総括テキスト
// ============================================================
const CLOSING_PARAGRAPHS = [
  { text: "コンフォートホテル横浜関内は、関内駅徒歩5分という抜群の立地と充実した無料朝食を核とした強固な競争力を持つホテルです。高評価率67.1%・平均7.52点は、改装工事中という特殊状況下でも健全な水準を維持しています。", options: { breakLine: true, fontSize: 12 } },
  { text: "", options: { fontSize: 6, breakLine: true } },
  { text: "直近の低評価（12.6%）は改装工事中の朝食形式変更が主因であり、工事完了後の朝食ビュッフェ復活で自然解消が見込まれます。並行して、客室設備の点検・修繕とスタッフ接客品質の向上に取り組むことで、全体評価の底上げが期待できます。", options: { breakLine: true, fontSize: 12 } },
  { text: "", options: { fontSize: 6, breakLine: true } },
  { text: "横浜スタジアム・中華街・山下公園への徒歩アクセスという地域密着型の立地優位性を活かし、観光・ビジネス双方のゲストへの訴求力を高めてください。工事完了後の積極的なOTA情報更新とリピーター向け施策の強化が重要です。", options: { breakLine: true, fontSize: 12 } },
  { text: "", options: { fontSize: 6, breakLine: true } },
  { text: "改装完了を転換点として、全体平均8.0点・高評価率75%達成を目指し、横浜エリアでの競争優位性をさらに高めてください。", options: { bold: true, fontSize: 12, color: "D4A843" } },
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
