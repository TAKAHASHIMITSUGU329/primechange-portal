const pptxgen = require("pptxgenjs");
const pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "Hotel Consulting";
pres.title = "チサンホテル浜松町 口コミ分析改善レポート";

const C = {
  navy: "1A2744", navyLight: "243556", blue: "3B7DD8", blueLight: "5A9BE6",
  ice: "E8EFF8", white: "FFFFFF", offWhite: "F5F7FA",
  gray: "64748B", grayLight: "94A3B8", grayDark: "334155",
  green: "16A34A", greenBg: "DCFCE7", red: "DC2626", redBg: "FEE2E2",
  orange: "EA580C", orangeBg: "FFF7ED", gold: "D4A843",
};

const shadow = () => ({ type: "outer", blur: 4, offset: 2, angle: 135, color: "000000", opacity: 0.10 });

function addFooter(slide, pageNum) {
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.navy } });
  slide.addText("Confidential", { x: 0.5, y: 5.25, w: 3, h: 0.375, fontSize: 8, color: C.grayLight, fontFace: "Arial", valign: "middle" });
  slide.addText(String(pageNum), { x: 9, y: 5.25, w: 0.5, h: 0.375, fontSize: 8, color: C.grayLight, fontFace: "Arial", align: "right", valign: "middle" });
}

function addContentHeader(slide, title, subtitle) {
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.85, fill: { color: C.navy } });
  slide.addText(title, { x: 0.6, y: 0.08, w: 8, h: 0.5, fontSize: 22, fontFace: "Arial", color: C.white, bold: true, margin: 0 });
  if (subtitle) slide.addText(subtitle, { x: 0.6, y: 0.5, w: 8, h: 0.3, fontSize: 10, fontFace: "Arial", color: C.blueLight, margin: 0 });
}

function kpiCard(slide, x, y, w, h, label, value, color, bgColor) {
  slide.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: bgColor }, shadow: shadow() });
  slide.addShape(pres.shapes.RECTANGLE, { x, y, w, h: 0.05, fill: { color } });
  slide.addText(label, { x, y: y + 0.15, w, h: 0.3, fontSize: 10, fontFace: "Arial", color: C.gray, align: "center", margin: 0 });
  slide.addText(value, { x, y: y + 0.4, w, h: 0.55, fontSize: 32, fontFace: "Arial", color, bold: true, align: "center", margin: 0 });
}

// ========== SLIDE 1: TITLE ==========
let s1 = pres.addSlide();
s1.background = { color: C.navy };
s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: C.navyLight, transparency: 40 } });
s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 2.0, w: 0.08, h: 1.8, fill: { color: C.gold } });
s1.addText("口コミ分析", { x: 0.8, y: 1.5, w: 8, h: 0.7, fontSize: 20, fontFace: "Arial", color: C.blueLight, charSpacing: 6, margin: 0 });
s1.addText("改善レポート", { x: 0.8, y: 2.1, w: 8, h: 0.9, fontSize: 42, fontFace: "Arial", color: C.white, bold: true, margin: 0 });
s1.addShape(pres.shapes.LINE, { x: 0.8, y: 3.1, w: 2, h: 0, line: { color: C.gold, width: 2 } });
s1.addText("チサンホテル浜松町", { x: 0.8, y: 3.3, w: 8, h: 0.5, fontSize: 16, fontFace: "Arial", color: C.grayLight, margin: 0 });
s1.addText([
  { text: "分析対象期間：2026年2月〜3月", options: { breakLine: true } },
  { text: "レビュー総数：47件（5サイト）", options: { breakLine: true } },
  { text: "作成日：2026年3月7日" },
], { x: 0.8, y: 4.1, w: 5, h: 0.9, fontSize: 10, fontFace: "Arial", color: C.grayLight, margin: 0, paraSpaceAfter: 4 });
s1.addShape(pres.shapes.RECTANGLE, { x: 7.5, y: 0, w: 2.5, h: 5.625, fill: { color: C.blue, transparency: 85 } });

// ========== SLIDE 2: EXECUTIVE SUMMARY ==========
let s2 = pres.addSlide();
s2.background = { color: C.offWhite };
addContentHeader(s2, "エグゼクティブサマリー", "Executive Summary");

kpiCard(s2, 0.5, 1.15, 2.05, 1.05, "全体平均(10pt換算)", "7.15", C.blue, C.white);
kpiCard(s2, 2.75, 1.15, 2.05, 1.05, "高評価率(8-10点)", "59.6%", C.green, C.white);
kpiCard(s2, 5.0, 1.15, 2.05, 1.05, "低評価率(1-4点)", "21.3%", C.orange, C.white);
kpiCard(s2, 7.25, 1.15, 2.25, 1.05, "レビュー総数", "47件", C.navy, C.white);

// Strengths
s2.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 2.5, w: 4.3, h: 2.5, fill: { color: C.white }, shadow: shadow() });
s2.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 2.5, w: 4.3, h: 0.05, fill: { color: C.green } });
s2.addText("Strengths", { x: 0.7, y: 2.6, w: 3, h: 0.35, fontSize: 14, fontFace: "Arial", color: C.green, bold: true, margin: 0 });
s2.addText([
  { text: "立地・アクセス ", options: { bold: true, breakLine: true } },
  { text: "  34%が言及。浜松町駅近く、羽田直結", options: { fontSize: 9, color: C.gray, breakLine: true } },
  { text: "コストパフォーマンス ", options: { bold: true, breakLine: true } },
  { text: "  25%が言及。リーズナブルな料金設定", options: { fontSize: 9, color: C.gray, breakLine: true } },
  { text: "朝食・清潔感 ", options: { bold: true, breakLine: true } },
  { text: "  朝食7件、清潔感10件。基盤は健在", options: { fontSize: 9, color: C.gray } },
], { x: 0.7, y: 3.0, w: 3.8, h: 1.8, fontSize: 11, fontFace: "Arial", color: C.grayDark, margin: 0, paraSpaceAfter: 2 });

// Weaknesses
s2.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 2.5, w: 4.3, h: 2.5, fill: { color: C.white }, shadow: shadow() });
s2.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 2.5, w: 4.3, h: 0.05, fill: { color: C.red } });
s2.addText("Weaknesses", { x: 5.4, y: 2.6, w: 3, h: 0.35, fontSize: 14, fontFace: "Arial", color: C.red, bold: true, margin: 0 });
s2.addText([
  { text: "設備老朽化 ", options: { bold: true, breakLine: true } },
  { text: "  壁紙・カーペット・家具の経年劣化（10件）", options: { fontSize: 9, color: C.gray, breakLine: true } },
  { text: "騒音・防音 ", options: { bold: true, breakLine: true } },
  { text: "  隣室音・外部騒音・設備音（9件）", options: { fontSize: 9, color: C.gray, breakLine: true } },
  { text: "水回り・清掃 ", options: { bold: true, breakLine: true } },
  { text: "  水圧弱い・排水遅い（6件）/ 清掃不足（5件）", options: { fontSize: 9, color: C.gray } },
], { x: 5.4, y: 3.0, w: 3.8, h: 1.8, fontSize: 11, fontFace: "Arial", color: C.grayDark, margin: 0, paraSpaceAfter: 2 });

addFooter(s2, 2);

// ========== SLIDE 3: SITE-BY-SITE RATINGS ==========
let s3 = pres.addSlide();
s3.background = { color: C.offWhite };
addContentHeader(s3, "サイト別評価分析", "Rating Analysis by Platform");

s3.addChart(pres.charts.BAR, [
  { name: "10pt換算", labels: ["Agoda", "Booking.com", "楽天トラベル", "じゃらん", "Google"], values: [8.67, 8.00, 7.06, 6.86, 6.17] }
], {
  x: 0.5, y: 1.1, w: 5.5, h: 3.5, barDir: "bar",
  chartColors: [C.blue],
  chartArea: { fill: { color: C.white }, roundedCorners: true },
  catAxisLabelColor: C.grayDark, catAxisLabelFontSize: 10,
  valAxisLabelColor: C.gray, valAxisLabelFontSize: 9,
  valGridLine: { color: "E2E8F0", size: 0.5 }, catGridLine: { style: "none" },
  showValue: true, dataLabelPosition: "outEnd", dataLabelColor: C.grayDark,
  showLegend: false, valAxisMaxVal: 10,
});

// Insight box
s3.addShape(pres.shapes.RECTANGLE, { x: 6.3, y: 1.1, w: 3.3, h: 3.5, fill: { color: C.white }, shadow: shadow() });
s3.addShape(pres.shapes.RECTANGLE, { x: 6.3, y: 1.1, w: 3.3, h: 0.05, fill: { color: C.gold } });
s3.addText("Insight", { x: 6.5, y: 1.2, w: 2, h: 0.35, fontSize: 13, fontFace: "Arial", color: C.gold, bold: true, margin: 0 });
s3.addText([
  { text: "改善のポイント", options: { bold: true, fontSize: 12, breakLine: true } },
  { text: "", options: { fontSize: 6, breakLine: true } },
  { text: "海外OTA：安定した高評価", options: { bold: true, color: C.green, breakLine: true } },
  { text: "Agoda 8.67 / Booking 8.00", options: { fontSize: 9, color: C.gray, breakLine: true } },
  { text: "外国人旅行者の満足度は良好", options: { fontSize: 9, color: C.gray, breakLine: true } },
  { text: "", options: { fontSize: 6, breakLine: true } },
  { text: "国内サイト：改善余地あり", options: { bold: true, color: C.orange, breakLine: true } },
  { text: "楽天 3.53/5 (7.06) / じゃらん 3.43/5 (6.86)", options: { fontSize: 9, color: C.gray, breakLine: true } },
  { text: "Google 3.08/5 (6.17) が最優先課題", options: { fontSize: 9, color: C.gray, breakLine: true } },
  { text: "", options: { fontSize: 6, breakLine: true } },
  { text: "注：", options: { bold: true, fontSize: 9, color: C.gray } },
  { text: "チャートは全て10pt換算で表示", options: { fontSize: 9, color: C.gray } },
], { x: 6.5, y: 1.55, w: 2.9, h: 2.9, fontSize: 11, fontFace: "Arial", color: C.grayDark, margin: 0, paraSpaceAfter: 2 });

// Table
const tblHeader = [
  { text: "サイト名", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 9, fontFace: "Arial", align: "center" } },
  { text: "件数", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 9, fontFace: "Arial", align: "center" } },
  { text: "平均(native)", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 9, fontFace: "Arial", align: "center" } },
  { text: "尺度", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 9, fontFace: "Arial", align: "center" } },
  { text: "10pt換算", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 9, fontFace: "Arial", align: "center" } },
];
const tblRows = [
  ["Agoda", "9", "8.67", "/10", "8.67"],
  ["Booking.com", "2", "8.00", "/10", "8.00"],
  ["楽天トラベル", "17", "3.53", "/5", "7.06"],
  ["じゃらん", "7", "3.43", "/5", "6.86"],
  ["Google", "12", "3.08", "/5", "6.17"],
].map((r, i) => r.map(cell => ({
  text: cell, options: { fontSize: 9, fontFace: "Arial", align: "center", fill: { color: i % 2 === 0 ? C.ice : C.white } }
})));

s3.addTable([tblHeader, ...tblRows], {
  x: 0.5, y: 4.65, w: 9, h: 0.1,
  colW: [2.0, 1.2, 1.8, 1.2, 1.8],
  border: { pt: 0.5, color: "CBD5E1" },
  rowH: [0.25, 0.22, 0.22, 0.22, 0.22, 0.22],
});

addFooter(s3, 3);

// ========== SLIDE 4: RATING DISTRIBUTION ==========
let s4 = pres.addSlide();
s4.background = { color: C.offWhite };
addContentHeader(s4, "評価分布分析（10点換算）", "Rating Distribution");

s4.addChart(pres.charts.DOUGHNUT, [{
  name: "評価分布",
  labels: ["高評価(8-10)", "中評価(5-7)", "低評価(1-4)"],
  values: [28, 9, 10],
}], {
  x: 0.3, y: 1.2, w: 4.0, h: 3.2,
  chartColors: [C.green, C.gold, C.red],
  showPercent: true, dataLabelColor: C.white, dataLabelFontSize: 11,
  showTitle: false, showLegend: true, legendPos: "b", legendFontSize: 9, legendColor: C.gray,
});

s4.addChart(pres.charts.BAR, [{
  name: "件数",
  labels: ["10点", "9点", "8点", "6点", "4点", "2点"],
  values: [11, 2, 15, 9, 7, 3],
}], {
  x: 4.5, y: 1.2, w: 5.2, h: 3.2, barDir: "col",
  chartColors: [C.green, C.green, C.green, C.gold, C.red, C.red],
  chartArea: { fill: { color: C.white }, roundedCorners: true },
  catAxisLabelColor: C.grayDark, catAxisLabelFontSize: 9,
  valAxisLabelColor: C.gray, valAxisLabelFontSize: 8,
  valGridLine: { color: "E2E8F0", size: 0.5 }, catGridLine: { style: "none" },
  showValue: true, dataLabelPosition: "outEnd", dataLabelColor: C.grayDark, dataLabelFontSize: 9,
  showLegend: false,
});

const cats = [
  { label: "高評価 8-10点", val: "28件（59.6%）", col: C.green, bg: C.greenBg },
  { label: "中評価 5-7点", val: "9件（19.1%）", col: C.gold, bg: C.orangeBg },
  { label: "低評価 1-4点", val: "10件（21.3%）", col: C.red, bg: C.redBg },
];
cats.forEach((c, i) => {
  const x = 0.5 + i * 3.1;
  s4.addShape(pres.shapes.RECTANGLE, { x, y: 4.55, w: 2.9, h: 0.55, fill: { color: c.bg } });
  s4.addText(c.label, { x, y: 4.55, w: 1.5, h: 0.55, fontSize: 10, fontFace: "Arial", color: c.col, bold: true, valign: "middle", margin: [0,0,0,8] });
  s4.addText(c.val, { x: x + 1.4, y: 4.55, w: 1.5, h: 0.55, fontSize: 11, fontFace: "Arial", color: c.col, bold: true, align: "right", valign: "middle", margin: [0,8,0,0] });
});

addFooter(s4, 4);

// ========== SLIDE 5: STRENGTHS ==========
let s5 = pres.addSlide();
s5.background = { color: C.offWhite };
addContentHeader(s5, "強み分析", "Strength Analysis");

const strengths = [
  { theme: "立地・アクセス", count: "16件", desc: "浜松町駅・大門駅から徒歩圏内、羽田直結", quote: "浜松町から近くて便利、出張に最適な立地" },
  { theme: "コストパフォーマンス", count: "12件", desc: "リーズナブルな料金設定、価格相応", quote: "この値段で立地が良いので満足" },
  { theme: "清潔感", count: "10件", desc: "清掃が行き届いている、きれいな部屋", quote: "部屋がきれいで気持ちよく過ごせた" },
  { theme: "設備・部屋", count: "8件", desc: "部屋が広い、設備が整っている", quote: "思ったより広くて使いやすかった" },
  { theme: "朝食", count: "7件", desc: "朝食が美味しい、和食メニュー充実", quote: "朝食の和食メニューが充実していた" },
  { theme: "スタッフ対応", count: "6件", desc: "親切丁寧、フロントの笑顔", quote: "スタッフの対応が親切で好感が持てた" },
];

strengths.forEach((s, i) => {
  const row = Math.floor(i / 3), col = i % 3;
  const x = 0.5 + col * 3.1, y = 1.15 + row * 2.0;
  s5.addShape(pres.shapes.RECTANGLE, { x, y, w: 2.9, h: 1.75, fill: { color: C.white }, shadow: shadow() });
  s5.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.06, h: 1.75, fill: { color: C.green } });
  s5.addShape(pres.shapes.RECTANGLE, { x: x + 2.0, y: y + 0.08, w: 0.72, h: 0.25, fill: { color: C.greenBg } });
  s5.addText(s.count, { x: x + 2.0, y: y + 0.08, w: 0.72, h: 0.25, fontSize: 9, fontFace: "Arial", color: C.green, bold: true, align: "center", valign: "middle", margin: 0 });
  s5.addText(s.theme, { x: x + 0.15, y: y + 0.08, w: 1.8, h: 0.3, fontSize: 13, fontFace: "Arial", color: C.navy, bold: true, margin: 0 });
  s5.addText(s.desc, { x: x + 0.15, y: y + 0.4, w: 2.6, h: 0.4, fontSize: 9, fontFace: "Arial", color: C.gray, margin: 0 });
  s5.addText(s.quote, { x: x + 0.15, y: y + 0.85, w: 2.6, h: 0.7, fontSize: 9, fontFace: "Arial", color: C.blue, italic: true, margin: 0 });
});

addFooter(s5, 5);

// ========== SLIDE 6: WEAKNESS / PRIORITY MATRIX ==========
let s6 = pres.addSlide();
s6.background = { color: C.offWhite };
addContentHeader(s6, "弱み分析・優先度マトリクス", "Weakness Analysis & Priority Matrix");

const weaknesses = [
  { pri: "S", cat: "設備老朽化", detail: "壁紙・カーペット・家具の経年劣化", count: "10件", color: C.red },
  { pri: "S", cat: "騒音・防音", detail: "隣室音・外部騒音・設備音が問題", count: "9件", color: C.red },
  { pri: "A", cat: "水回り", detail: "水圧弱い、排水遅い、バスルーム狭い", count: "6件", color: C.orange },
  { pri: "A", cat: "清掃品質", detail: "髪の毛、ほこり、ベッド周りの清掃不足", count: "5件", color: C.orange },
  { pri: "B", cat: "コスパ割高感", detail: "設備の古さに対して料金が高い", count: "4件", color: C.gold },
  { pri: "B", cat: "立地（ネガティブ）", detail: "道順不明、夜道が暗い、坂がある", count: "3件", color: C.gold },
  { pri: "C", cat: "タバコ臭", detail: "禁煙室でもタバコの匂い", count: "2件", color: C.grayLight },
  { pri: "C", cat: "空調問題", detail: "エアコン効き悪い、温度調整困難", count: "2件", color: C.grayLight },
  { pri: "C", cat: "加湿器", detail: "加湿器が古い/未設置", count: "1件", color: C.grayLight },
];

const wTblHeader = [
  { text: "優先度", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 9, fontFace: "Arial", align: "center" } },
  { text: "課題カテゴリ", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 9, fontFace: "Arial" } },
  { text: "具体的内容", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 9, fontFace: "Arial" } },
  { text: "件数", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 9, fontFace: "Arial", align: "center" } },
];
const wTblRows = weaknesses.map((w, i) => [
  { text: w.pri, options: { fontSize: 12, fontFace: "Arial", align: "center", bold: true, color: w.color, fill: { color: i % 2 === 0 ? C.ice : C.white } } },
  { text: w.cat, options: { fontSize: 9, fontFace: "Arial", bold: true, color: C.grayDark, fill: { color: i % 2 === 0 ? C.ice : C.white } } },
  { text: w.detail, options: { fontSize: 9, fontFace: "Arial", color: C.gray, fill: { color: i % 2 === 0 ? C.ice : C.white } } },
  { text: w.count, options: { fontSize: 9, fontFace: "Arial", align: "center", color: C.grayDark, fill: { color: i % 2 === 0 ? C.ice : C.white } } },
]);

s6.addTable([wTblHeader, ...wTblRows], {
  x: 0.5, y: 1.15, w: 9, h: 0.1, colW: [0.8, 2.0, 5.0, 1.2],
  border: { pt: 0.5, color: "CBD5E1" },
  rowH: [0.3, 0.35, 0.35, 0.35, 0.35, 0.35, 0.35, 0.35, 0.35, 0.35],
});

s6.addText([
  { text: "S ", options: { bold: true, color: C.red, fontSize: 10 } }, { text: "= 最優先   ", options: { fontSize: 9, color: C.gray } },
  { text: "A ", options: { bold: true, color: C.orange, fontSize: 10 } }, { text: "= 高優先   ", options: { fontSize: 9, color: C.gray } },
  { text: "B ", options: { bold: true, color: C.gold, fontSize: 10 } }, { text: "= 中優先   ", options: { fontSize: 9, color: C.gray } },
  { text: "C ", options: { bold: true, color: C.grayLight, fontSize: 10 } }, { text: "= 低優先", options: { fontSize: 9, color: C.gray } },
], { x: 0.5, y: 4.8, w: 9, h: 0.3, fontFace: "Arial", margin: 0 });

addFooter(s6, 6);

// ========== SLIDE 7: IMPROVEMENT PHASE 1 ==========
let s7 = pres.addSlide();
s7.background = { color: C.offWhite };
addContentHeader(s7, "改善施策 Phase 1：即座対応", "Immediate Actions (Today ~ 1 Month)");

s7.addText("投資を最小限に抑え、オペレーション改善で効果を目指す", { x: 0.6, y: 0.95, w: 8, h: 0.3, fontSize: 10, fontFace: "Arial", color: C.blue, italic: true, margin: 0 });

const phase1 = [
  { title: "清掃品質の強化", items: ["チェックリスト見直し（髪の毛・ほこり・水垢）", "ダブルチェック体制の導入", "清掃後の換気30分以上確保", "週次の抜き打ち品質チェック"] },
  { title: "騒音対策（即効性）", items: ["ドア下部・窓際の簡易防音テープ設置", "耳栓のアメニティ提供を検討", "静かな客室の優先割当てオペレーション", "チェックイン時の希望確認"] },
  { title: "タバコ臭・匂い対策", items: ["禁煙室のオゾン脱臭を月2回実施", "カーペット・カーテンへの消臭処理", "水回りメンテナンス頻度の向上"] },
  { title: "口コミ返信開始", items: ["全サイト48時間以内の返信体制構築", "低評価には謝罪＋具体的改善策を記載", "Google口コミを最優先で対応（3.08/5）", "高評価にも感謝メッセージ返信"] },
];

phase1.forEach((p, i) => {
  const col = i % 2, row = Math.floor(i / 2);
  const x = 0.5 + col * 4.6, y = 1.35 + row * 1.95;
  s7.addShape(pres.shapes.RECTANGLE, { x, y, w: 4.35, h: 1.75, fill: { color: C.white }, shadow: shadow() });
  s7.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.06, h: 1.75, fill: { color: C.blue } });
  s7.addText(p.title, { x: x + 0.15, y: y + 0.05, w: 4, h: 0.35, fontSize: 12, fontFace: "Arial", color: C.navy, bold: true, margin: 0 });
  const bulletItems = p.items.map((item, idx) => ({ text: item, options: { bullet: true, fontSize: 9, color: C.grayDark, breakLine: idx < p.items.length - 1 } }));
  s7.addText(bulletItems, { x: x + 0.15, y: y + 0.4, w: 4, h: 1.3, fontFace: "Arial", margin: 0, paraSpaceAfter: 3 });
});

addFooter(s7, 7);

// ========== SLIDE 8: PHASE 2 & 3 ==========
let s8 = pres.addSlide();
s8.background = { color: C.offWhite };
addContentHeader(s8, "改善施策 Phase 2・3", "Short-term & Mid-term Actions");

s8.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.1, w: 4.35, h: 3.8, fill: { color: C.white }, shadow: shadow() });
s8.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.1, w: 4.35, h: 0.4, fill: { color: C.blue } });
s8.addText("Phase 2：短期施策（1〜3ヶ月）", { x: 0.65, y: 1.1, w: 4, h: 0.4, fontSize: 12, fontFace: "Arial", color: C.white, bold: true, margin: 0, valign: "middle" });

const p2Items = [
  { title: "客室リフレッシュ（部分改修）", items: ["壁紙・カーペット交換（優先フロアから）", "照明のLED化で清潔感演出", "小物の新品交換"] },
  { title: "OTA対策の強化", items: ["施設写真の刷新（プロ撮影）", "Googleビジネスプロフィール最新化", "朝食の魅力をOTAでアピール"] },
  { title: "アクセス・空調改善", items: ["写真付きアクセスガイド作成", "エアコンフィルター定期清掃厳格化", "冬季の全室加湿器設置"] },
];

let p2y = 1.6;
p2Items.forEach((p) => {
  s8.addText(p.title, { x: 0.7, y: p2y, w: 3.8, h: 0.3, fontSize: 10, fontFace: "Arial", color: C.navy, bold: true, margin: 0 });
  const items = p.items.map((item, idx) => ({ text: item, options: { bullet: true, fontSize: 9, color: C.gray, breakLine: idx < p.items.length - 1 } }));
  s8.addText(items, { x: 0.7, y: p2y + 0.28, w: 3.8, h: p.items.length * 0.22 + 0.1, fontFace: "Arial", margin: 0, paraSpaceAfter: 2 });
  p2y += 0.28 + p.items.length * 0.22 + 0.2;
});

s8.addShape(pres.shapes.RECTANGLE, { x: 5.15, y: 1.1, w: 4.35, h: 3.8, fill: { color: C.white }, shadow: shadow() });
s8.addShape(pres.shapes.RECTANGLE, { x: 5.15, y: 1.1, w: 4.35, h: 0.4, fill: { color: C.navy } });
s8.addText("Phase 3：中期施策（3〜6ヶ月）", { x: 5.3, y: 1.1, w: 4, h: 0.4, fontSize: 12, fontFace: "Arial", color: C.white, bold: true, margin: 0, valign: "middle" });

const p3Items = [
  { title: "防音工事", items: ["低評価集中フロアの防音調査・改修", "窓の二重サッシ化（外部騒音対策）", "壁面の遮音材追加", "ドア・窓の気密性向上工事"] },
  { title: "水回りリニューアル", items: ["水圧改善のための給水設備更新", "バスルーム内装の改修", "配管の大規模メンテナンス"] },
  { title: "段階的リノベーション", items: ["最老朽フロアからの改修計画策定", "モダンインテリアへの刷新", "デジタルインフラ（USB/WiFi）整備"] },
];

let p3y = 1.6;
p3Items.forEach((p) => {
  s8.addText(p.title, { x: 5.35, y: p3y, w: 3.8, h: 0.3, fontSize: 10, fontFace: "Arial", color: C.navy, bold: true, margin: 0 });
  const items = p.items.map((item, idx) => ({ text: item, options: { bullet: true, fontSize: 9, color: C.gray, breakLine: idx < p.items.length - 1 } }));
  s8.addText(items, { x: 5.35, y: p3y + 0.28, w: 3.8, h: p.items.length * 0.22 + 0.1, fontFace: "Arial", margin: 0, paraSpaceAfter: 2 });
  p3y += 0.28 + p.items.length * 0.22 + 0.2;
});

addFooter(s8, 8);

// ========== SLIDE 9: KPI TARGETS ==========
let s9 = pres.addSlide();
s9.background = { color: C.offWhite };
addContentHeader(s9, "KPI目標設定", "Key Performance Indicators");

const kpiHeader = [
  { text: "KPI項目", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 10, fontFace: "Arial" } },
  { text: "現状値", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 10, fontFace: "Arial", align: "center" } },
  { text: "目標値", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 10, fontFace: "Arial", align: "center" } },
  { text: "期限", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 10, fontFace: "Arial", align: "center" } },
];

const kpiData = [
  ["全体平均評価(10pt)", "7.15点", "7.8点以上", "2026年9月"],
  ["高評価率（8-10点）", "59.6%", "65%以上", "2026年9月"],
  ["低評価率（1-4点）", "21.3%", "15%以下", "2026年9月"],
  ["楽天トラベル平均", "3.53/5点", "3.8/5点以上", "2026年9月"],
  ["Google評価", "3.08/5点", "3.5/5点以上", "2026年9月"],
  ["口コミ返信率", "未計測", "100%（48h以内）", "2026年6月"],
].map((r, i) => [
  { text: r[0], options: { fontSize: 10, fontFace: "Arial", bold: true, color: C.grayDark, fill: { color: i % 2 === 0 ? C.ice : C.white } } },
  { text: r[1], options: { fontSize: 10, fontFace: "Arial", align: "center", color: C.gray, fill: { color: i % 2 === 0 ? C.ice : C.white } } },
  { text: r[2], options: { fontSize: 10, fontFace: "Arial", align: "center", bold: true, color: C.green, fill: { color: i % 2 === 0 ? C.ice : C.white } } },
  { text: r[3], options: { fontSize: 10, fontFace: "Arial", align: "center", color: C.gray, fill: { color: i % 2 === 0 ? C.ice : C.white } } },
]);

s9.addTable([kpiHeader, ...kpiData], {
  x: 0.5, y: 1.2, w: 9, h: 0.1, colW: [2.8, 2.0, 2.2, 2.0],
  border: { pt: 0.5, color: "CBD5E1" },
  rowH: [0.35, 0.35, 0.35, 0.35, 0.35, 0.35, 0.35],
});

s9.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.85, w: 9, h: 0.6, fill: { color: C.white }, shadow: shadow() });
s9.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.85, w: 0.06, h: 0.6, fill: { color: C.blue } });
s9.addText([
  { text: "モニタリング方針：", options: { bold: true, color: C.navy } },
  { text: "四半期ごとにKPIをレビューし、目標未達の場合は追加施策を検討。月次で口コミモニタリングを実施し、新たな課題の早期発見に努める。", options: { color: C.gray } },
], { x: 0.7, y: 3.85, w: 8.6, h: 0.6, fontSize: 10, fontFace: "Arial", valign: "middle", margin: 0 });

addFooter(s9, 9);

// ========== SLIDE 10: CLOSING ==========
let s10 = pres.addSlide();
s10.background = { color: C.navy };
s10.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: C.navyLight, transparency: 40 } });
s10.addText("総括", { x: 0.8, y: 0.8, w: 8, h: 0.5, fontSize: 14, fontFace: "Arial", color: C.blueLight, charSpacing: 6, margin: 0 });
s10.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 1.5, w: 8.4, h: 2.8, fill: { color: C.white, transparency: 90 } });

s10.addText([
  { text: "当ホテルは全体平均7.15点（10pt換算）、高評価率59.6%と中程度の水準を維持しています。", options: { breakLine: true, fontSize: 12 } },
  { text: "", options: { fontSize: 6, breakLine: true } },
  { text: "「立地」「コスパ」「朝食」の強みは健在であり、海外OTAでの安定した高評価は基本品質の良さを示しています。", options: { breakLine: true, fontSize: 12 } },
  { text: "", options: { fontSize: 6, breakLine: true } },
  { text: "Google（3.08/5）とじゃらん（3.43/5）の国内サイト評価の改善が全体水準向上の鍵です。設備老朽化・騒音・水回りの課題に計画的に対応します。", options: { breakLine: true, fontSize: 12 } },
  { text: "", options: { fontSize: 6, breakLine: true } },
  { text: "Phase 1のオペレーション改善を着実に推進し、月次モニタリングで効果を検証していきましょう。", options: { bold: true, fontSize: 12, color: C.gold } },
], { x: 1.0, y: 1.6, w: 8, h: 2.5, fontFace: "Arial", color: C.white, margin: 0, paraSpaceAfter: 4 });

s10.addShape(pres.shapes.LINE, { x: 0.8, y: 4.5, w: 2, h: 0, line: { color: C.gold, width: 2 } });
s10.addText("ご清聴ありがとうございました", { x: 0.8, y: 4.6, w: 8, h: 0.4, fontSize: 14, fontFace: "Arial", color: C.grayLight, margin: 0 });

// === SAVE ===
const outPath = "/Users/mitsugutakahashi/ホテル口コミ/チサンホテル浜松町_口コミ分析レポート.pptx";
pres.writeFile({ fileName: outPath }).then(() => {
  console.log("PPTX created: " + outPath);
}).catch(err => console.error(err));
