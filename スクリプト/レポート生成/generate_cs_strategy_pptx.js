#!/usr/bin/env node
"use strict";

const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");

// ============================================================
// Data Loading
// ============================================================
const qualityData = JSON.parse(fs.readFileSync(path.resolve(__dirname, "primechange_portfolio_analysis.json"), "utf-8"));
const revenueData = JSON.parse(fs.readFileSync(path.resolve(__dirname, "hotel_revenue_data.json"), "utf-8"));

const overview = qualityData.portfolio_overview;
const deepDive = qualityData.cleaning_deep_dive || {};
const priorityMatrix = qualityData.priority_matrix || {};
const hotelsRanked = overview.hotels_ranked || [];
const categorySummary = deepDive.category_summary || [];

const KEY_MAP = { keisei_kinshicho: "keisei_richmond", comfort_yokohama_kannai: "comfort_yokohama" };
function getRev(qKey) { return revenueData[KEY_MAP[qKey] || qKey] || {}; }

const hotels = hotelsRanked.map(q => {
  const r = getRev(q.key);
  return { ...q, revenue: r.actual_revenue || 0, occupancy: r.occupancy_rate || 0, profit_rate: r.profit_rate || 0, adr: r.adr || 0, room_count: r.room_count || 0 };
});

const totalRevenue = hotels.reduce((s, h) => s + h.revenue, 0);
const medianRev = [...hotels].sort((a, b) => a.revenue - b.revenue)[Math.floor(hotels.length / 2)].revenue;
const qualityThreshold = 8.0;
hotels.forEach(h => {
  const hq = h.avg >= qualityThreshold, hr = h.revenue >= medianRev;
  h.quadrant = hq && hr ? "高品質×高売上" : hq && !hr ? "高品質×低売上" : !hq && hr ? "低品質×高売上" : "低品質×低売上";
});

function fmtY(n) { return n ? "¥" + Number(n).toLocaleString("ja-JP", { maximumFractionDigits: 0 }) : "-"; }
function fmtP(n) { return n ? (Number(n) * 100).toFixed(1) + "%" : "-"; }

// ============================================================
// Color Palette (matching existing reports)
// ============================================================
const C = {
  NAVY: "1B3A5C", ACCENT: "2E75B6", WHITE: "FFFFFF", LIGHT_BG: "F5F7FA",
  TEXT: "333333", SUBTEXT: "666666", GREEN: "27AE60", ORANGE: "FF9800",
  RED: "E74C3C", BLUE: "2196F3", TEAL: "00695C", DARK_GREEN: "1B5E20",
};

const PRIORITY_COLORS = {
  URGENT: { bg: "FFEBEE", text: "E74C3C" },
  HIGH: { bg: "FFF3E0", text: "FF9800" },
  STANDARD: { bg: "E3F2FD", text: "2196F3" },
  MAINTENANCE: { bg: "E8F5E9", text: "27AE60" },
};

// ============================================================
// PPTX Setup
// ============================================================
const pptx = new pptxgen();
pptx.layout = "LAYOUT_16x9";
pptx.author = "PRIMECHANGE";
pptx.title = "PRIMECHANGE CS向上×売上増加 戦略提案書";

const shadow = () => ({ type: "outer", blur: 4, offset: 2, angle: 135, color: "000000", opacity: 0.10 });

function addFooter(slide, pageNum) {
  slide.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.NAVY } });
  slide.addText("Confidential | PRIMECHANGE CS戦略提案書", { x: 0.5, y: 5.25, w: 5, h: 0.375, fontSize: 8, color: "94A3B8", fontFace: "Arial", valign: "middle" });
  slide.addText(String(pageNum), { x: 9, y: 5.25, w: 0.5, h: 0.375, fontSize: 8, color: "94A3B8", fontFace: "Arial", align: "right", valign: "middle" });
}

function addContentHeader(slide, title, subtitle) {
  slide.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.6, fill: { color: C.NAVY } });
  slide.addText(title, { x: 0.5, y: 0.05, w: 9, h: 0.35, fontSize: 20, fontFace: "Arial", color: C.WHITE, bold: true, margin: 0 });
  if (subtitle) {
    slide.addText(subtitle, { x: 0.5, y: 0.35, w: 9, h: 0.22, fontSize: 10, fontFace: "Arial", color: "5A9BE6", margin: 0 });
  }
}

function kpiCard(slide, x, y, w, h, label, value, color, bgColor) {
  slide.addShape(pptx.shapes.RECTANGLE, { x, y, w, h, fill: { color: bgColor }, shadow: shadow() });
  slide.addShape(pptx.shapes.RECTANGLE, { x, y, w, h: 0.04, fill: { color } });
  slide.addText(label, { x, y: y + 0.12, w, h: 0.22, fontSize: 9, fontFace: "Arial", color: C.SUBTEXT, align: "center", margin: 0 });
  slide.addText(value, { x, y: y + 0.32, w, h: 0.45, fontSize: 28, fontFace: "Arial", color, bold: true, align: "center", margin: 0 });
}

// ============================================================
// Slide 1: Title
// ============================================================
(function buildSlide1() {
  const slide = pptx.addSlide();
  slide.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: C.NAVY } });
  // Accent line
  slide.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0, w: 0.08, h: 5.625, fill: { color: C.ACCENT } });
  slide.addText("PRIMECHANGE", { x: 0.8, y: 1.0, w: 8.4, h: 0.7, fontSize: 20, fontFace: "Arial", color: "5A9BE6", margin: 0 });
  slide.addText("CS向上 × 売上増加", { x: 0.8, y: 1.7, w: 8.4, h: 0.8, fontSize: 36, fontFace: "Arial", color: C.WHITE, bold: true, margin: 0 });
  slide.addText("戦略提案書", { x: 0.8, y: 2.5, w: 8.4, h: 0.7, fontSize: 32, fontFace: "Arial", color: C.WHITE, bold: true, margin: 0 });
  slide.addShape(pptx.shapes.RECTANGLE, { x: 0.8, y: 3.3, w: 2.0, h: 0.04, fill: { color: "5A9BE6" } });
  slide.addText("Strategic Proposal for Customer Satisfaction & Revenue Growth", { x: 0.8, y: 3.5, w: 8.4, h: 0.35, fontSize: 13, fontFace: "Arial", color: "94A3B8", italics: true, margin: 0 });
  slide.addText(`対象：${hotels.length}ホテル | 月間総売上：${fmtY(totalRevenue)} | 口コミ：${overview.total_reviews}件`, { x: 0.8, y: 4.1, w: 8, h: 0.3, fontSize: 11, fontFace: "Arial", color: "64748B", margin: 0 });
  slide.addText("2026年3月", { x: 0.8, y: 4.4, w: 4, h: 0.3, fontSize: 11, fontFace: "Arial", color: "64748B", margin: 0 });
  slide.addText("株式会社PRIMECHANGE", { x: 0.8, y: 4.9, w: 8, h: 0.3, fontSize: 10, fontFace: "Arial", color: "475569", margin: 0 });
})();

// ============================================================
// Slide 2: Executive Summary
// ============================================================
(function buildSlide2() {
  const slide = pptx.addSlide();
  addContentHeader(slide, "エグゼクティブサマリー", "CS向上から売上増加へ — 3つの重要な発見");

  kpiCard(slide, 0.3, 0.8, 2.1, 0.85, "平均スコア", String(overview.avg_score || 8.39), C.ACCENT, "E3F2FD");
  kpiCard(slide, 2.65, 0.8, 2.1, 0.85, "月間総売上", "¥1.0億", C.GREEN, "E8F5E9");
  kpiCard(slide, 5.0, 0.8, 2.1, 0.85, "クレーム率", (deepDive.portfolio_cleaning_issue_rate || 4.6) + "%", C.RED, "FFEBEE");
  kpiCard(slide, 7.35, 0.8, 2.1, 0.85, "改善余地", "+3-5%", C.TEAL, "E0F2F1");

  // Three key findings
  const findings = [
    { icon: "🔍", title: "未発掘データの宝庫", body: "月次レポート22シートのうち70%以上が未分析。\nクレーム類型・人員配置・安全チェック等の\nデータが眠っている。", color: C.ACCENT },
    { icon: "⚠️", title: "高リスクの3ホテル", body: "低品質×高売上の3ホテル（博多/蒲田/浜松町）は\n月間売上計約2,000万円。品質悪化→売上急落\nリスクが最も高い。", color: C.RED },
    { icon: "📈", title: "月間+300-500万円の余地", body: "7つの分析テーマで品質を改善し、\nスコア0.5点向上を実現すれば\n月間300-500万円の売上増加。", color: C.GREEN },
  ];

  findings.forEach((f, i) => {
    const x = 0.3 + i * 3.15;
    slide.addShape(pptx.shapes.RECTANGLE, { x, y: 2.0, w: 3.0, h: 2.8, fill: { color: C.LIGHT_BG }, shadow: shadow() });
    slide.addShape(pptx.shapes.RECTANGLE, { x, y: 2.0, w: 3.0, h: 0.04, fill: { color: f.color } });
    slide.addText(f.title, { x: x + 0.15, y: 2.15, w: 2.7, h: 0.3, fontSize: 13, fontFace: "Arial", color: f.color, bold: true, margin: 0 });
    slide.addText(f.body, { x: x + 0.15, y: 2.5, w: 2.7, h: 2.2, fontSize: 10, fontFace: "Arial", color: C.TEXT, valign: "top", margin: 0 });
  });

  addFooter(slide, 2);
})();

// ============================================================
// Slide 3: CS→売上 Causal Mechanism
// ============================================================
(function buildSlide3() {
  const slide = pptx.addSlide();
  addContentHeader(slide, "CS→売上の因果メカニズム", "口コミスコア0.1点改善 ≒ RevPAR約1%向上");

  // Flow diagram using boxes and arrows
  const steps = [
    { text: "分析による\n課題特定", bg: "E3F2FD", color: C.ACCENT },
    { text: "的確な\n改善施策", bg: "E8F5E9", color: C.GREEN },
    { text: "清掃品質向上\nクレーム削減", bg: "E8F5E9", color: C.DARK_GREEN },
    { text: "口コミスコア\n改善", bg: "FFF3E0", color: C.ORANGE },
    { text: "OTAランキング↑\n稼働率向上", bg: "FFF3E0", color: C.ORANGE },
    { text: "ホテル売上\n増加", bg: "FFEBEE", color: C.RED },
  ];

  const startX = 0.3;
  const boxW = 1.35;
  const gap = 0.2;
  const y = 1.0;
  const boxH = 0.9;

  steps.forEach((s, i) => {
    const x = startX + i * (boxW + gap);
    slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x, y, w: boxW, h: boxH, fill: { color: s.bg }, shadow: shadow(), rectRadius: 0.08 });
    slide.addText(s.text, { x, y, w: boxW, h: boxH, fontSize: 9, fontFace: "Arial", color: s.color, bold: true, align: "center", valign: "middle", margin: 0 });
    // Arrow between boxes
    if (i < steps.length - 1) {
      slide.addText("→", { x: x + boxW, y: y + 0.2, w: gap, h: 0.5, fontSize: 16, fontFace: "Arial", color: C.SUBTEXT, align: "center", valign: "middle", margin: 0 });
    }
  });

  // PRIMECHANGE benefit box
  slide.addShape(pptx.shapes.RECTANGLE, { x: 0.3, y: 2.2, w: 9.4, h: 0.5, fill: { color: C.NAVY }, shadow: shadow() });
  slide.addText("→ PRIMECHANGEの契約維持 ・ 単価交渉力向上 ・ 新規受注競争力 → PRIMECHANGE売上増加", {
    x: 0.5, y: 2.2, w: 9.0, h: 0.5, fontSize: 12, fontFace: "Arial", color: C.WHITE, bold: true, align: "center", valign: "middle", margin: 0,
  });

  // Supporting data
  slide.addShape(pptx.shapes.RECTANGLE, { x: 0.3, y: 3.0, w: 4.5, h: 2.0, fill: { color: "F0F7FC" }, shadow: shadow() });
  slide.addText("自社データによる推定", { x: 0.5, y: 3.1, w: 4, h: 0.3, fontSize: 12, fontFace: "Arial", color: C.NAVY, bold: true, margin: 0 });
  slide.addText([
    { text: "月間総売上: ", options: { fontSize: 10, color: C.SUBTEXT } }, { text: fmtY(totalRevenue), options: { fontSize: 10, color: C.TEXT, bold: true } }, { text: "\n", options: { fontSize: 6 } },
    { text: "ポートフォリオ平均: ", options: { fontSize: 10, color: C.SUBTEXT } }, { text: `${overview.avg_score}点`, options: { fontSize: 10, color: C.TEXT, bold: true } }, { text: "\n", options: { fontSize: 6 } },
    { text: "0.5点改善時の推定効果: ", options: { fontSize: 10, color: C.SUBTEXT } }, { text: "+月間300-500万円", options: { fontSize: 10, color: C.GREEN, bold: true } }, { text: "\n", options: { fontSize: 6 } },
    { text: "年間換算: ", options: { fontSize: 10, color: C.SUBTEXT } }, { text: "+3,600-6,000万円", options: { fontSize: 10, color: C.GREEN, bold: true } },
  ], { x: 0.5, y: 3.45, w: 4.1, h: 1.4, valign: "top", margin: 0 });

  slide.addShape(pptx.shapes.RECTANGLE, { x: 5.2, y: 3.0, w: 4.5, h: 2.0, fill: { color: "FFF8E1" }, shadow: shadow() });
  slide.addText("業界ベンチマーク", { x: 5.4, y: 3.1, w: 4, h: 0.3, fontSize: 12, fontFace: "Arial", color: C.ORANGE, bold: true, margin: 0 });
  slide.addText([
    { text: "スコア0.1点改善 → RevPAR +1%\n\n", options: { fontSize: 11, color: C.TEXT, bold: true } },
    { text: "高スコアホテル（9.0+）は低スコアホテル（7.0-8.0）と比べ\n稼働率が平均10-15%高い\n\n", options: { fontSize: 9, color: C.TEXT } },
    { text: "清掃品質は口コミスコアに最も影響する因子の一つ", options: { fontSize: 9, color: C.TEXT } },
  ], { x: 5.4, y: 3.45, w: 4.1, h: 1.4, valign: "top", margin: 0 });

  addFooter(slide, 3);
})();

// ============================================================
// Slide 4: Quality × Revenue Matrix
// ============================================================
(function buildSlide4() {
  const slide = pptx.addSlide();
  addContentHeader(slide, "現状：品質×売上マトリクス", "19ホテルの4象限分類");

  const quadrants = [
    { name: "高品質×高売上", sub: "維持・横展開モデル", color: C.GREEN, bg: "E8F5E9", x: 0.3, y: 0.8, w: 4.55, h: 2.0 },
    { name: "高品質×低売上", sub: "営業強化対象", color: C.ORANGE, bg: "FFF3E0", x: 5.15, y: 0.8, w: 4.55, h: 2.0 },
    { name: "低品質×高売上", sub: "⚠️ 最大リスク", color: C.RED, bg: "FFEBEE", x: 0.3, y: 3.0, w: 4.55, h: 1.9 },
    { name: "低品質×低売上", sub: "根本改革", color: "BF360C", bg: "FBE9E7", x: 5.15, y: 3.0, w: 4.55, h: 1.9 },
  ];

  quadrants.forEach(q => {
    const list = hotels.filter(h => h.quadrant === q.name);
    slide.addShape(pptx.shapes.RECTANGLE, { x: q.x, y: q.y, w: q.w, h: q.h, fill: { color: q.bg }, shadow: shadow() });
    slide.addShape(pptx.shapes.RECTANGLE, { x: q.x, y: q.y, w: q.w, h: 0.04, fill: { color: q.color } });
    slide.addText(`${q.name} (${list.length})`, { x: q.x + 0.15, y: q.y + 0.08, w: q.w - 0.3, h: 0.28, fontSize: 11, fontFace: "Arial", color: q.color, bold: true, margin: 0 });
    slide.addText(q.sub, { x: q.x + 0.15, y: q.y + 0.35, w: q.w - 0.3, h: 0.2, fontSize: 8, fontFace: "Arial", color: C.SUBTEXT, margin: 0 });

    const hotelTexts = list.map(h => `${h.name}（${h.avg}点, ${fmtY(h.revenue)}）`);
    slide.addText(hotelTexts.join("\n"), { x: q.x + 0.15, y: q.y + 0.6, w: q.w - 0.3, h: q.h - 0.7, fontSize: 8, fontFace: "Arial", color: C.TEXT, valign: "top", margin: 0 });
  });

  addFooter(slide, 4);
})();

// ============================================================
// Slide 5: Untapped Data
// ============================================================
(function buildSlide5() {
  const slide = pptx.addSlide();
  addContentHeader(slide, "課題：未発掘データの宝庫", "月次レポートXLSX 22シートのうち70%以上が未分析");

  // Two-column: utilized vs untapped
  slide.addShape(pptx.shapes.RECTANGLE, { x: 0.3, y: 0.85, w: 3.0, h: 3.9, fill: { color: "E8F5E9" }, shadow: shadow() });
  slide.addShape(pptx.shapes.RECTANGLE, { x: 0.3, y: 0.85, w: 3.0, h: 0.04, fill: { color: C.GREEN } });
  slide.addText("✅ 活用済みデータ", { x: 0.45, y: 0.95, w: 2.7, h: 0.3, fontSize: 12, fontFace: "Arial", color: C.GREEN, bold: true, margin: 0 });
  slide.addText([
    { text: "①R8_*集計\n", options: { fontSize: 10, bold: true, color: C.TEXT } },
    { text: "月次KPI（売上/稼働率/利益率等）\n\n", options: { fontSize: 8, color: C.SUBTEXT } },
    { text: "💭口コミ\n", options: { fontSize: 10, bold: true, color: C.TEXT } },
    { text: "6サイトの口コミテキスト\n（1,583件分析済み）\n", options: { fontSize: 8, color: C.SUBTEXT } },
  ], { x: 0.45, y: 1.35, w: 2.7, h: 3.2, valign: "top", margin: 0 });

  slide.addShape(pptx.shapes.RECTANGLE, { x: 3.6, y: 0.85, w: 6.1, h: 3.9, fill: { color: "FFF3E0" }, shadow: shadow() });
  slide.addShape(pptx.shapes.RECTANGLE, { x: 3.6, y: 0.85, w: 6.1, h: 0.04, fill: { color: C.ORANGE } });
  slide.addText("❌ 未活用データ（分析提案対象）", { x: 3.75, y: 0.95, w: 5.8, h: 0.3, fontSize: 12, fontFace: "Arial", color: C.ORANGE, bold: true, margin: 0 });

  const untapped = [
    { sheet: "🔵クレーム", desc: "13類型の月別集計", analysis: "→ 分析1" },
    { sheet: "R8品質データまとめ", desc: "メイド/チェッカー別クレーム", analysis: "→ 分析2" },
    { sheet: "③日報", desc: "出勤人数/完了時間/日次クレーム", analysis: "→ 分析3,4" },
    { sheet: "✅安全チェック", desc: "パトロール◎/△/✖", analysis: "→ 分析5" },
    { sheet: "🏆皆勤アワード", desc: "スタッフ勤務/清掃部屋数", analysis: "→ 分析2" },
    { sheet: "④月報", desc: "改善進捗/スタッフ在籍数", analysis: "" },
    { sheet: "年間集計", desc: "12ヶ月トレンド", analysis: "" },
    { sheet: "🧹特別清掃", desc: "特別清掃請求管理", analysis: "" },
  ];

  const tableHeader = [
    { text: "シート名", options: { fontSize: 8, bold: true, color: C.WHITE, fill: { color: C.NAVY } } },
    { text: "データ内容", options: { fontSize: 8, bold: true, color: C.WHITE, fill: { color: C.NAVY } } },
    { text: "対応分析", options: { fontSize: 8, bold: true, color: C.WHITE, fill: { color: C.NAVY }, align: "center" } },
  ];

  const rows = [tableHeader, ...untapped.map((u, i) => [
    { text: u.sheet, options: { fontSize: 8, bold: true, fill: { color: i % 2 === 0 ? C.WHITE : C.LIGHT_BG } } },
    { text: u.desc, options: { fontSize: 8, fill: { color: i % 2 === 0 ? C.WHITE : C.LIGHT_BG } } },
    { text: u.analysis, options: { fontSize: 8, bold: true, color: C.ACCENT, align: "center", fill: { color: i % 2 === 0 ? C.WHITE : C.LIGHT_BG } } },
  ])];

  slide.addTable(rows, { x: 3.75, y: 1.35, w: 5.8, colW: [2.0, 2.3, 1.2], fontSize: 8, border: { pt: 0.5, color: "CCCCCC" }, autoPage: false });

  addFooter(slide, 5);
})();

// ============================================================
// Slide 6: Analysis Concept 1 - Claim Type × Score
// ============================================================
(function buildSlide6() {
  const slide = pptx.addSlide();
  addContentHeader(slide, "分析1: クレーム類型×口コミスコア連動分析", "最もスコアに効くクレーム類型を特定する");

  // Left: Why & How
  slide.addShape(pptx.shapes.RECTANGLE, { x: 0.3, y: 0.8, w: 4.5, h: 4.15, fill: { color: "F0F7FC" }, shadow: shadow() });
  slide.addShape(pptx.shapes.RECTANGLE, { x: 0.3, y: 0.8, w: 4.5, h: 0.04, fill: { color: C.ACCENT } });

  slide.addText("なぜ必要か", { x: 0.45, y: 0.9, w: 4, h: 0.25, fontSize: 11, fontFace: "Arial", color: C.NAVY, bold: true, margin: 0 });
  slide.addText("13種のクレーム類型のうち、どれがスコア低下に最も影響するかが未特定。限られたリソースを最もインパクトのある類型に集中するため。", {
    x: 0.45, y: 1.2, w: 4.2, h: 0.7, fontSize: 9, fontFace: "Arial", color: C.TEXT, valign: "top", margin: 0,
  });

  slide.addText("使用データ", { x: 0.45, y: 1.95, w: 4, h: 0.25, fontSize: 11, fontFace: "Arial", color: C.NAVY, bold: true, margin: 0 });
  slide.addText("🔵クレームシート: 13類型の月別集計\n（巻き込み/誤入室/髪の毛/残置/セット漏れ/\n  汚れ/清掃不備/未清掃 等）\n口コミスコア: ホテル別月次平均", {
    x: 0.45, y: 2.25, w: 4.2, h: 0.8, fontSize: 9, fontFace: "Arial", color: C.TEXT, valign: "top", margin: 0,
  });

  slide.addText("分析手法", { x: 0.45, y: 3.1, w: 4, h: 0.25, fontSize: 11, fontFace: "Arial", color: C.NAVY, bold: true, margin: 0 });
  slide.addText("- 類型別発生率とスコアの相関係数算出\n- 重回帰分析でインパクトを分離\n- 類型別「スコア弾力性」の推定", {
    x: 0.45, y: 3.4, w: 4.2, h: 0.7, fontSize: 9, fontFace: "Arial", color: C.TEXT, valign: "top", margin: 0,
  });

  // Right: Expected findings & impact
  slide.addShape(pptx.shapes.RECTANGLE, { x: 5.15, y: 0.8, w: 4.55, h: 2.4, fill: { color: C.LIGHT_BG }, shadow: shadow() });
  slide.addShape(pptx.shapes.RECTANGLE, { x: 5.15, y: 0.8, w: 4.55, h: 0.04, fill: { color: C.GREEN } });

  slide.addText("期待される発見", { x: 5.3, y: 0.9, w: 4, h: 0.25, fontSize: 11, fontFace: "Arial", color: C.GREEN, bold: true, margin: 0 });
  slide.addText("✓ 「髪の毛」と「汚れ」のどちらがスコアに\n   より大きく影響するかが判明\n✓ ゲストが最も不快に感じる類型の優先順位\n✓ ホテルごとの「改善すべき類型」の違い", {
    x: 5.3, y: 1.2, w: 4.2, h: 1.9, fontSize: 9, fontFace: "Arial", color: C.TEXT, valign: "top", margin: 0,
  });

  // Impact box
  slide.addShape(pptx.shapes.RECTANGLE, { x: 5.15, y: 3.4, w: 4.55, h: 1.55, fill: { color: "E8F5E9" }, shadow: shadow() });
  slide.addShape(pptx.shapes.RECTANGLE, { x: 5.15, y: 3.4, w: 4.55, h: 0.04, fill: { color: C.DARK_GREEN } });

  slide.addText("CS→売上インパクト", { x: 5.3, y: 3.5, w: 4, h: 0.25, fontSize: 11, fontFace: "Arial", color: C.DARK_GREEN, bold: true, margin: 0 });
  slide.addText("最もスコアに効く類型に集中対策\n→ 効率的にスコア改善\n→ OTAランキング上昇 → 売上増加", {
    x: 5.3, y: 3.8, w: 4.2, h: 0.6, fontSize: 9, fontFace: "Arial", color: C.TEXT, valign: "top", margin: 0,
  });
  slide.addText("推定: スコア+0.3点 → 月間+約300万円", { x: 5.3, y: 4.5, w: 4.2, h: 0.25, fontSize: 11, fontFace: "Arial", color: C.GREEN, bold: true, margin: 0 });

  // Bottom bar: implementation
  slide.addShape(pptx.shapes.RECTANGLE, { x: 0.3, y: 5.05, w: 4.0, h: 0.15, fill: { color: C.ACCENT } });
  slide.addText("難易度: 中 | 期間: 1-2週間", { x: 0.45, y: 4.75, w: 4, h: 0.25, fontSize: 9, fontFace: "Arial", color: C.SUBTEXT, margin: 0 });

  addFooter(slide, 6);
})();

// ============================================================
// Slide 7: Analysis Concept 2 & 3
// ============================================================
(function buildSlide7() {
  const slide = pptx.addSlide();
  addContentHeader(slide, "分析2-3: 人に関する分析", "スタッフパフォーマンスと最適人員配置");

  // Analysis 2: Staff Performance
  slide.addShape(pptx.shapes.RECTANGLE, { x: 0.3, y: 0.8, w: 4.5, h: 4.15, fill: { color: "F0F7FC" }, shadow: shadow() });
  slide.addShape(pptx.shapes.RECTANGLE, { x: 0.3, y: 0.8, w: 4.5, h: 0.04, fill: { color: C.ACCENT } });

  slide.addText("分析2: スタッフ個人別パフォーマンス", { x: 0.45, y: 0.9, w: 4.2, h: 0.25, fontSize: 12, fontFace: "Arial", color: C.ACCENT, bold: true, margin: 0 });

  slide.addText([
    { text: "使用データ: ", options: { fontSize: 9, color: C.NAVY, bold: true } },
    { text: "R8品質データまとめ, 🏆皆勤アワード\n\n", options: { fontSize: 9, color: C.TEXT } },
    { text: "手法:\n", options: { fontSize: 9, color: C.NAVY, bold: true } },
    { text: "- メイド別「1室あたりクレーム率」算出\n- トップ10% / ボトム10%の特定\n- ハイパフォーマーの共通点分析\n\n", options: { fontSize: 9, color: C.TEXT } },
    { text: "期待される発見:\n", options: { fontSize: 9, color: C.NAVY, bold: true } },
    { text: "- クレームの大半が少数スタッフに集中\n  （パレートの法則）\n- チェッカーの「見逃し率」の個人差\n- ハイパフォーマーの清掃手順の特徴\n\n", options: { fontSize: 9, color: C.TEXT } },
    { text: "効果: ", options: { fontSize: 9, color: C.GREEN, bold: true } },
    { text: "ボトム20%を平均まで改善 →\n       クレーム率20-30%削減", options: { fontSize: 9, color: C.GREEN, bold: true } },
  ], { x: 0.45, y: 1.25, w: 4.2, h: 3.6, valign: "top", margin: 0 });

  // Analysis 3: Staff Allocation
  slide.addShape(pptx.shapes.RECTANGLE, { x: 5.15, y: 0.8, w: 4.55, h: 4.15, fill: { color: C.LIGHT_BG }, shadow: shadow() });
  slide.addShape(pptx.shapes.RECTANGLE, { x: 5.15, y: 0.8, w: 4.55, h: 0.04, fill: { color: C.TEAL } });

  slide.addText("分析3: 人員配置×品質 相関分析", { x: 5.3, y: 0.9, w: 4.2, h: 0.25, fontSize: 12, fontFace: "Arial", color: C.TEAL, bold: true, margin: 0 });

  slide.addText([
    { text: "使用データ: ", options: { fontSize: 9, color: C.NAVY, bold: true } },
    { text: "③日報（出勤メイド数/チェッカー数/稼働客室数）\n\n", options: { fontSize: 9, color: C.TEXT } },
    { text: "手法:\n", options: { fontSize: 9, color: C.NAVY, bold: true } },
    { text: "- 「1メイドあたり客室数」vs「クレーム率」\n  の散布図＋回帰分析\n- チェッカー不在日のクレーム率比較\n- ホテル規模別の人員効率分析\n\n", options: { fontSize: 9, color: C.TEXT } },
    { text: "期待される発見:\n", options: { fontSize: 9, color: C.NAVY, bold: true } },
    { text: "- 「15室超でクレーム率急増」等の閾値\n- チェッカー不在日のクレーム増加量\n- ホテル規模別の推奨メイド数\n\n", options: { fontSize: 9, color: C.TEXT } },
    { text: "効果: ", options: { fontSize: 9, color: C.GREEN, bold: true } },
    { text: "人員不足日のクレーム率を50%削減\n       → 全体クレーム率10-15%削減", options: { fontSize: 9, color: C.GREEN, bold: true } },
  ], { x: 5.3, y: 1.25, w: 4.2, h: 3.6, valign: "top", margin: 0 });

  addFooter(slide, 7);
})();

// ============================================================
// Slide 8: Analysis Concept 4 & 5
// ============================================================
(function buildSlide8() {
  const slide = pptx.addSlide();
  addContentHeader(slide, "分析4-5: プロセスと予防", "時間管理の最適化とクレームの未然防止");

  // Analysis 4: Cleaning Time
  slide.addShape(pptx.shapes.RECTANGLE, { x: 0.3, y: 0.8, w: 4.5, h: 4.15, fill: { color: "FFF3E0" }, shadow: shadow() });
  slide.addShape(pptx.shapes.RECTANGLE, { x: 0.3, y: 0.8, w: 4.5, h: 0.04, fill: { color: C.ORANGE } });

  slide.addText("分析4: 清掃完了時間×品質 分析", { x: 0.45, y: 0.9, w: 4.2, h: 0.25, fontSize: 12, fontFace: "Arial", color: C.ORANGE, bold: true, margin: 0 });

  slide.addText([
    { text: "使用データ: ", options: { fontSize: 9, color: C.NAVY, bold: true } },
    { text: "③日報（清掃完了時間/クレーム件数/稼働客室数）\n\n", options: { fontSize: 9, color: C.TEXT } },
    { text: "手法:\n", options: { fontSize: 9, color: C.NAVY, bold: true } },
    { text: "- 完了時間帯別のクレーム率比較\n  （〜13時 / 13-14時 / 14-15時 / 15時〜）\n- 負荷度（客室÷メイド数）と完了時間\n- チェッカー検査の余裕度分析\n\n", options: { fontSize: 9, color: C.TEXT } },
    { text: "期待される発見:\n", options: { fontSize: 9, color: C.NAVY, bold: true } },
    { text: "- 15時以降完了日はクレーム率2倍以上\n- チェッカー検査時間と品質の関係\n\n", options: { fontSize: 9, color: C.TEXT } },
    { text: "効果: ", options: { fontSize: 9, color: C.GREEN, bold: true } },
    { text: "スケジュール最適化で\n時間起因クレームを60%削減", options: { fontSize: 9, color: C.GREEN, bold: true } },
  ], { x: 0.45, y: 1.25, w: 4.2, h: 3.6, valign: "top", margin: 0 });

  slide.addText("難易度: 低〜中 | 期間: 1-2週間", { x: 0.45, y: 4.75, w: 4, h: 0.2, fontSize: 8, fontFace: "Arial", color: C.SUBTEXT, margin: 0 });

  // Analysis 5: Safety Check Prediction
  slide.addShape(pptx.shapes.RECTANGLE, { x: 5.15, y: 0.8, w: 4.55, h: 4.15, fill: { color: "FFEBEE" }, shadow: shadow() });
  slide.addShape(pptx.shapes.RECTANGLE, { x: 5.15, y: 0.8, w: 4.55, h: 0.04, fill: { color: C.RED } });

  slide.addText("分析5: 安全チェック×クレーム予兆検出", { x: 5.3, y: 0.9, w: 4.2, h: 0.25, fontSize: 12, fontFace: "Arial", color: C.RED, bold: true, margin: 0 });

  slide.addText([
    { text: "使用データ: ", options: { fontSize: 9, color: C.NAVY, bold: true } },
    { text: "✅安全チェック（◎/△/✖評価）\n🔵クレーム（月別件数）\n\n", options: { fontSize: 9, color: C.TEXT } },
    { text: "手法:\n", options: { fontSize: 9, color: C.NAVY, bold: true } },
    { text: "- パトロール「△/✖」率と翌月クレームの相関\n- 衛生管理✖ → 翌月クレーム率の検証\n- 予兆スコア開発（翌月予測モデル）\n\n", options: { fontSize: 9, color: C.TEXT } },
    { text: "期待される発見:\n", options: { fontSize: 9, color: C.NAVY, bold: true } },
    { text: "- 「△」3項目以上 → 翌月クレーム有意増加\n- 最も予測力の高い評価項目の特定\n\n", options: { fontSize: 9, color: C.TEXT } },
    { text: "効果: ", options: { fontSize: 9, color: C.GREEN, bold: true } },
    { text: "予兆検知で先手対応 →\nクレーム発生を30-40%回避", options: { fontSize: 9, color: C.GREEN, bold: true } },
  ], { x: 5.3, y: 1.25, w: 4.2, h: 3.6, valign: "top", margin: 0 });

  slide.addText("難易度: 高 | 期間: 3-4週間", { x: 5.3, y: 4.75, w: 4, h: 0.2, fontSize: 8, fontFace: "Arial", color: C.SUBTEXT, margin: 0 });

  addFooter(slide, 8);
})();

// ============================================================
// Slide 9: Analysis 6 & 7
// ============================================================
(function buildSlide9() {
  const slide = pptx.addSlide();
  addContentHeader(slide, "分析6-7: ROI精密化と横展開", "既存データで今すぐ始められる分析");

  // Analysis 6: ROI
  slide.addShape(pptx.shapes.RECTANGLE, { x: 0.3, y: 0.8, w: 4.5, h: 4.15, fill: { color: "E8F5E9" }, shadow: shadow() });
  slide.addShape(pptx.shapes.RECTANGLE, { x: 0.3, y: 0.8, w: 4.5, h: 0.04, fill: { color: C.GREEN } });

  slide.addText("分析6: 品質→売上 弾力性分析", { x: 0.45, y: 0.9, w: 4.2, h: 0.25, fontSize: 12, fontFace: "Arial", color: C.GREEN, bold: true, margin: 0 });

  // Highlight: can do NOW
  slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: 0.45, y: 1.22, w: 2.2, h: 0.25, fill: { color: C.GREEN }, rectRadius: 0.04 });
  slide.addText("✅ 既存データで即実施可能", { x: 0.5, y: 1.22, w: 2.15, h: 0.25, fontSize: 8, fontFace: "Arial", color: C.WHITE, bold: true, margin: 0 });

  slide.addText([
    { text: "\n使用データ: ", options: { fontSize: 9, color: C.NAVY, bold: true } },
    { text: "19ホテルのスコア × 稼働率 × 売上 × ADR\n（既存JSONファイルで全て保有）\n\n", options: { fontSize: 9, color: C.TEXT } },
    { text: "手法:\n", options: { fontSize: 9, color: C.NAVY, bold: true } },
    { text: "- スコア vs 稼働率/ADR/RevPARの回帰分析\n- 「スコア1点あたりの差分」算出\n- ホテル規模で補正した弾力性推定\n\n", options: { fontSize: 9, color: C.TEXT } },
    { text: "期待される発見:\n", options: { fontSize: 9, color: C.NAVY, bold: true } },
    { text: "- 自社データでの弾力性の精密値\n- 「8.0未満で稼働率急落」等の非線形パターン\n\n", options: { fontSize: 9, color: C.TEXT } },
    { text: "効果: ", options: { fontSize: 9, color: C.GREEN, bold: true } },
    { text: "投資配分最適化で同予算から\n20-30%多い効果", options: { fontSize: 9, color: C.GREEN, bold: true } },
  ], { x: 0.45, y: 1.5, w: 4.2, h: 3.3, valign: "top", margin: 0 });

  slide.addText("難易度: 低 | 期間: 1週間", { x: 0.45, y: 4.75, w: 4, h: 0.2, fontSize: 8, fontFace: "Arial", color: C.SUBTEXT, margin: 0 });

  // Analysis 7: Best Practice
  slide.addShape(pptx.shapes.RECTANGLE, { x: 5.15, y: 0.8, w: 4.55, h: 4.15, fill: { color: "F0F7FC" }, shadow: shadow() });
  slide.addShape(pptx.shapes.RECTANGLE, { x: 5.15, y: 0.8, w: 4.55, h: 0.04, fill: { color: C.ACCENT } });

  slide.addText("分析7: ベストプラクティス横展開", { x: 5.3, y: 0.9, w: 4.2, h: 0.25, fontSize: 12, fontFace: "Arial", color: C.ACCENT, bold: true, margin: 0 });

  // Top performers
  const top3 = hotelsRanked.slice(0, 3);
  const bot3 = hotelsRanked.slice(-3);
  slide.addText([
    { text: "高スコアホテルの成功要因を体系化:\n\n", options: { fontSize: 9, color: C.TEXT } },
    { text: "🏆 トップ3:\n", options: { fontSize: 9, color: C.GREEN, bold: true } },
    { text: top3.map(h => `  ${h.name}（${h.avg}点）`).join("\n") + "\n\n", options: { fontSize: 8, color: C.TEXT } },
    { text: "要改善3:\n", options: { fontSize: 9, color: C.RED, bold: true } },
    { text: bot3.map(h => `  ${h.name}（${h.avg}点）`).join("\n") + "\n\n", options: { fontSize: 8, color: C.TEXT } },
    { text: "手法:\n", options: { fontSize: 9, color: C.NAVY, bold: true } },
    { text: "- 口コミ「good」コメントのテキスト分析\n- オペレーション指標の比較\n- メイド/客室比率の違い\n\n", options: { fontSize: 9, color: C.TEXT } },
    { text: "効果: ", options: { fontSize: 9, color: C.GREEN, bold: true } },
    { text: "低スコア5施設 +0.5点改善\n→ 月間+約250万円", options: { fontSize: 9, color: C.GREEN, bold: true } },
  ], { x: 5.3, y: 1.2, w: 4.2, h: 3.6, valign: "top", margin: 0 });

  slide.addText("難易度: 中 | 期間: 3-4週間", { x: 5.3, y: 4.75, w: 4, h: 0.2, fontSize: 8, fontFace: "Arial", color: C.SUBTEXT, margin: 0 });

  addFooter(slide, 9);
})();

// ============================================================
// Slide 10: 7 Themes Summary Table
// ============================================================
(function buildSlide10() {
  const slide = pptx.addSlide();
  addContentHeader(slide, "7テーマ一覧サマリー", "難易度・効果・期間の全体マップ");

  const header = [
    { text: "#", options: { fontSize: 8, bold: true, color: C.WHITE, fill: { color: C.NAVY }, align: "center" } },
    { text: "分析テーマ", options: { fontSize: 8, bold: true, color: C.WHITE, fill: { color: C.NAVY } } },
    { text: "難易度", options: { fontSize: 8, bold: true, color: C.WHITE, fill: { color: C.NAVY }, align: "center" } },
    { text: "効果", options: { fontSize: 8, bold: true, color: C.WHITE, fill: { color: C.NAVY }, align: "center" } },
    { text: "期間", options: { fontSize: 8, bold: true, color: C.WHITE, fill: { color: C.NAVY }, align: "center" } },
    { text: "データ", options: { fontSize: 8, bold: true, color: C.WHITE, fill: { color: C.NAVY }, align: "center" } },
    { text: "推定売上効果", options: { fontSize: 8, bold: true, color: C.WHITE, fill: { color: C.NAVY }, align: "center" } },
  ];

  const themes = [
    ["1", "クレーム類型×スコア連動", "中", "大", "1-2週", "未活用", "+300万/月"],
    ["2", "スタッフ個人パフォーマンス", "中〜高", "大", "2-3週", "未活用", "CR -20-30%"],
    ["3", "人員配置×品質 相関", "中", "大", "2週", "未活用", "CR -10-15%"],
    ["4", "清掃完了時間×品質", "低〜中", "中", "1-2週", "未活用", "CR -60%(時間起因)"],
    ["5", "安全チェック×予兆検出", "高", "大", "3-4週", "未活用", "CR -30-40%"],
    ["6", "品質→売上弾力性（ROI）", "低", "大", "1週", "既存", "ROI精密化"],
    ["7", "ベストプラクティス横展開", "中", "大", "3-4週", "既存+", "+250万/月"],
  ];

  const rows = [header, ...themes.map((t, i) => t.map((c, j) => {
    const fill = i % 2 === 0 ? C.WHITE : C.LIGHT_BG;
    let color = C.TEXT;
    if (j === 2) color = c === "低" || c === "低〜中" ? C.GREEN : c === "高" ? C.RED : C.ORANGE;
    if (j === 3) color = c === "大" ? C.GREEN : C.ORANGE;
    if (j === 6) color = C.GREEN;
    return { text: c, options: { fontSize: 8, bold: j <= 1 || j === 6, color, align: j === 1 ? "left" : "center", fill: { color: fill } } };
  }))];

  slide.addTable(rows, { x: 0.2, y: 0.8, w: 9.6, colW: [0.35, 2.8, 0.7, 0.5, 0.7, 0.7, 1.5], fontSize: 8, border: { pt: 0.5, color: "CCCCCC" }, autoPage: false });

  // Recommendation box
  slide.addShape(pptx.shapes.RECTANGLE, { x: 0.3, y: 3.8, w: 9.4, h: 1.2, fill: { color: "F0F7FC" }, shadow: shadow() });
  slide.addShape(pptx.shapes.RECTANGLE, { x: 0.3, y: 3.8, w: 9.4, h: 0.04, fill: { color: C.ACCENT } });
  slide.addText("推奨実施順序", { x: 0.5, y: 3.9, w: 3, h: 0.25, fontSize: 12, fontFace: "Arial", color: C.NAVY, bold: true, margin: 0 });
  slide.addText([
    { text: "① 分析6（即実施・1週間）", options: { fontSize: 10, bold: true, color: C.GREEN } },
    { text: " → ROIの精密値を確立し、全施策の費用対効果を定量化\n", options: { fontSize: 9, color: C.TEXT } },
    { text: "② 分析1（2週間）", options: { fontSize: 10, bold: true, color: C.ACCENT } },
    { text: " → 最もインパクトのあるクレーム類型を特定\n", options: { fontSize: 9, color: C.TEXT } },
    { text: "③ 分析3,4（並行・2週間）", options: { fontSize: 10, bold: true, color: C.ORANGE } },
    { text: " → 人員配置と時間管理の最適化\n", options: { fontSize: 9, color: C.TEXT } },
    { text: "④ 分析2,5,7（1-3ヶ月）", options: { fontSize: 10, bold: true, color: C.SUBTEXT } },
    { text: " → スタッフ育成・予兆検出・横展開の仕組み化", options: { fontSize: 9, color: C.TEXT } },
  ], { x: 0.5, y: 4.2, w: 9, h: 0.7, valign: "top", margin: 0 });

  addFooter(slide, 10);
})();

// ============================================================
// Slide 11: Action Plan (Phase 1-3)
// ============================================================
(function buildSlide11() {
  const slide = pptx.addSlide();
  addContentHeader(slide, "統合アクションプラン", "3フェーズ改善計画");

  const phases = [
    {
      label: "Phase 1: 即時（2週間）",
      color: C.RED, bg: "FFEBEE",
      items: [
        "品質→売上弾力性の算出（分析6）",
        "URGENTホテル緊急清掃監査",
        "低品質×高売上3ホテルの責任者面談",
        "クレーム類型データ抽出開始（分析1）",
      ],
      effect: "+100-200万円/月\n（売上減少の防止）",
    },
    {
      label: "Phase 2: 短期（1-3ヶ月）",
      color: C.ORANGE, bg: "FFF3E0",
      items: [
        "クレーム類型×スコア分析完了",
        "スタッフパフォーマンス分析実施",
        "人員配置最適化基準の策定",
        "ボトムパフォーマーへのOJT",
        "ベストプラクティス横展開開始",
      ],
      effect: "+200-300万円/月\n（URGENT改善）",
    },
    {
      label: "Phase 3: 中期（3-6ヶ月）",
      color: C.ACCENT, bg: "E3F2FD",
      items: [
        "月次品質ダッシュボード構築",
        "安全チェック×予兆検出の自動化",
        "スタッフ評価制度への品質KPI組込",
        "継続的改善サイクル（PDCA）確立",
      ],
      effect: "+300-500万円/月\n（全体底上げ）",
    },
  ];

  phases.forEach((phase, pi) => {
    const px = 0.3 + pi * 3.15;
    const py = 0.85;
    const ph = 4.1;
    slide.addShape(pptx.shapes.RECTANGLE, { x: px, y: py, w: 3.0, h: ph, fill: { color: phase.bg }, shadow: shadow() });
    slide.addShape(pptx.shapes.RECTANGLE, { x: px, y: py, w: 3.0, h: 0.04, fill: { color: phase.color } });
    slide.addText(phase.label, { x: px + 0.1, y: py + 0.08, w: 2.8, h: 0.3, fontSize: 11, fontFace: "Arial", color: phase.color, bold: true, margin: 0 });

    const itemsText = phase.items.map(item => `• ${item}`).join("\n");
    slide.addText(itemsText, { x: px + 0.1, y: py + 0.45, w: 2.8, h: 2.5, fontSize: 9, fontFace: "Arial", color: C.TEXT, valign: "top", margin: 0 });

    // Effect box at bottom
    slide.addShape(pptx.shapes.RECTANGLE, { x: px + 0.1, y: py + ph - 1.0, w: 2.8, h: 0.85, fill: { color: "E8F5E9" }, shadow: shadow() });
    slide.addText("推定効果", { x: px + 0.2, y: py + ph - 0.95, w: 2.6, h: 0.2, fontSize: 8, fontFace: "Arial", color: C.SUBTEXT, margin: 0 });
    slide.addText(phase.effect, { x: px + 0.2, y: py + ph - 0.7, w: 2.6, h: 0.55, fontSize: 10, fontFace: "Arial", color: C.GREEN, bold: true, margin: 0 });
  });

  addFooter(slide, 11);
})();

// ============================================================
// Slide 12: KPI Framework
// ============================================================
(function buildSlide12() {
  const slide = pptx.addSlide();
  addContentHeader(slide, "KPI・効果測定フレームワーク", "月次追跡KPI一覧");

  const header = [
    { text: "KPI", options: { fontSize: 8, bold: true, color: C.WHITE, fill: { color: C.NAVY } } },
    { text: "現状値", options: { fontSize: 8, bold: true, color: C.WHITE, fill: { color: C.NAVY }, align: "center" } },
    { text: "3ヶ月目標", options: { fontSize: 8, bold: true, color: C.WHITE, fill: { color: C.NAVY }, align: "center" } },
    { text: "6ヶ月目標", options: { fontSize: 8, bold: true, color: C.WHITE, fill: { color: C.NAVY }, align: "center" } },
    { text: "測定方法", options: { fontSize: 8, bold: true, color: C.WHITE, fill: { color: C.NAVY } } },
  ];

  const kpis = [
    ["ポートフォリオ平均スコア", "8.39", "8.55", "8.80", "月次口コミ集計"],
    ["清掃クレーム率", "4.6%", "3.5%", "2.5%", "🔵クレーム月次集計"],
    ["URGENT判定ホテル数", "4", "2", "0", "優先度再判定"],
    ["月間総売上", "¥99.8M", "¥101M", "¥104M", "①集計シート"],
    ["平均稼働率", "73.9%", "75.0%", "77.0%", "①集計シート"],
    ["高評価率（8点以上）", "78.1%", "80.0%", "83.0%", "口コミ分析"],
    ["スタッフ充足率", "未測定", "90%", "95%", "④月報"],
    ["予兆検出対応率", "0%", "50%", "90%", "安全チェック連動"],
  ];

  const rows = [header, ...kpis.map((k, i) => k.map((c, j) => {
    const fill = i % 2 === 0 ? C.WHITE : C.LIGHT_BG;
    const isTarget = j === 2 || j === 3;
    return { text: c, options: { fontSize: 8, bold: j === 0 || isTarget, color: isTarget ? C.GREEN : C.TEXT, align: j === 0 || j === 4 ? "left" : "center", fill: { color: fill } } };
  }))];

  slide.addTable(rows, { x: 0.3, y: 0.8, w: 9.4, colW: [2.4, 1.2, 1.2, 1.2, 2.2], fontSize: 8, border: { pt: 0.5, color: "CCCCCC" }, autoPage: false });

  // Reporting cycle
  slide.addShape(pptx.shapes.RECTANGLE, { x: 0.3, y: 3.6, w: 9.4, h: 1.3, fill: { color: C.LIGHT_BG }, shadow: shadow() });
  slide.addText("レポーティングサイクル", { x: 0.5, y: 3.7, w: 4, h: 0.25, fontSize: 12, fontFace: "Arial", color: C.NAVY, bold: true, margin: 0 });

  const cycles = [
    { label: "週次", desc: "URGENTホテルの\nクレーム件数", color: C.RED, bg: "FFEBEE" },
    { label: "月次", desc: "全KPIの\nダッシュボード", color: C.ORANGE, bg: "FFF3E0" },
    { label: "四半期", desc: "品質監査\n+ベストプラクティス", color: C.ACCENT, bg: "E3F2FD" },
    { label: "半期", desc: "戦略レビュー\n+目標更新", color: C.GREEN, bg: "E8F5E9" },
  ];

  cycles.forEach((c, i) => {
    const cx = 0.5 + i * 2.3;
    slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: cx, y: 4.05, w: 2.0, h: 0.75, fill: { color: c.bg }, shadow: shadow(), rectRadius: 0.06 });
    slide.addText(c.label, { x: cx, y: 4.08, w: 2.0, h: 0.25, fontSize: 10, fontFace: "Arial", color: c.color, bold: true, align: "center", margin: 0 });
    slide.addText(c.desc, { x: cx, y: 4.35, w: 2.0, h: 0.4, fontSize: 8, fontFace: "Arial", color: C.TEXT, align: "center", margin: 0 });
  });

  addFooter(slide, 12);
})();

// ============================================================
// Slide 13: ROI
// ============================================================
(function buildSlide13() {
  const slide = pptx.addSlide();
  addContentHeader(slide, "投資対効果（ROI）試算", "3シナリオでの費用対効果分析");

  const header = [
    { text: "項目", options: { fontSize: 8, bold: true, color: C.WHITE, fill: { color: C.NAVY } } },
    { text: "控えめ", options: { fontSize: 8, bold: true, color: C.WHITE, fill: { color: C.NAVY }, align: "center" } },
    { text: "標準", options: { fontSize: 8, bold: true, color: C.WHITE, fill: { color: C.NAVY }, align: "center" } },
    { text: "楽観", options: { fontSize: 8, bold: true, color: C.WHITE, fill: { color: C.NAVY }, align: "center" } },
  ];

  const roiData = [
    ["対象", "URGENT 4ホテル", "URGENT+HIGH 7ホテル", "全19ホテル"],
    ["分析投資", "100-200万円", "200-400万円", "400-600万円"],
    ["改善施策投資", "200-300万円", "400-600万円", "600-800万円"],
    ["スコア改善幅", "+0.3〜0.5点", "+0.3〜0.7点(対象)", "+0.3〜0.5点(全体)"],
    ["月間売上改善", "+100〜200万円", "+200〜400万円", "+300〜500万円"],
    ["年間売上改善", "+1,200〜2,400万円", "+2,400〜4,800万円", "+3,600〜6,000万円"],
    ["投資回収期間", "3-6ヶ月", "4-8ヶ月", "6-12ヶ月"],
    ["PRIMECHANGE効果", "契約維持確保", "単価交渉材料", "新規受注競争力"],
  ];

  const rows = [header, ...roiData.map((r, i) => r.map((c, j) => {
    const fill = i % 2 === 0 ? C.WHITE : C.LIGHT_BG;
    let color = C.TEXT;
    if (j === 0) color = C.NAVY;
    if (i === 4 || i === 5) color = j > 0 ? C.GREEN : C.NAVY;
    return { text: c, options: { fontSize: 8, bold: j === 0 || i === 4 || i === 5, color, align: j === 0 ? "left" : "center", fill: { color: fill } } };
  }))];

  slide.addTable(rows, { x: 0.2, y: 0.8, w: 9.6, colW: [1.8, 2.4, 2.7, 2.7], fontSize: 8, border: { pt: 0.5, color: "CCCCCC" }, autoPage: false });

  // Recommendation
  slide.addShape(pptx.shapes.RECTANGLE, { x: 0.3, y: 3.7, w: 9.4, h: 1.2, fill: { color: "F0F7FC" }, shadow: shadow() });
  slide.addShape(pptx.shapes.RECTANGLE, { x: 0.3, y: 3.7, w: 9.4, h: 0.04, fill: { color: C.ACCENT } });
  slide.addText("推奨：「控えめシナリオ」から開始 → 成果確認後に段階拡大", { x: 0.5, y: 3.8, w: 8.8, h: 0.3, fontSize: 13, fontFace: "Arial", color: C.NAVY, bold: true, margin: 0 });
  slide.addText("URGENTホテル4施設への集中投資（投資額300-500万円）で最短3ヶ月回収。\n成果を確認後、標準→楽観シナリオへ拡大。最終的にデータドリブン品質管理を全社に展開。", {
    x: 0.5, y: 4.15, w: 8.8, h: 0.6, fontSize: 10, fontFace: "Arial", color: C.TEXT, margin: 0,
  });

  addFooter(slide, 13);
})();

// ============================================================
// Slide 14: Next Steps
// ============================================================
(function buildSlide14() {
  const slide = pptx.addSlide();
  slide.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: C.NAVY } });
  slide.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0, w: 0.08, h: 5.625, fill: { color: C.ACCENT } });

  slide.addText("Next Steps", { x: 0.8, y: 0.5, w: 8, h: 0.5, fontSize: 28, fontFace: "Arial", color: C.WHITE, bold: true, margin: 0 });
  slide.addShape(pptx.shapes.RECTANGLE, { x: 0.8, y: 1.1, w: 2.0, h: 0.04, fill: { color: "5A9BE6" } });

  const steps = [
    { num: "01", title: "本提案書の内容について経営層との合意形成", timeline: "1週間以内", color: C.RED },
    { num: "02", title: "分析6（品質→売上弾力性）の即時実施", timeline: "1週間", color: C.RED },
    { num: "03", title: "URGENTホテルの現地視察・緊急監査", timeline: "2週間以内", color: C.ORANGE },
    { num: "04", title: "クレーム類型データの抽出・分析1の実施", timeline: "2-3週間", color: C.ORANGE },
    { num: "05", title: "Phase 2施策のキックオフ（スタッフ分析・人員配置分析）", timeline: "1ヶ月〜", color: "5A9BE6" },
    { num: "06", title: "月次品質ダッシュボードの設計・構築", timeline: "3ヶ月〜", color: "5A9BE6" },
  ];

  steps.forEach((s, i) => {
    const sy = 1.4 + i * 0.65;
    slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: sy, w: 0.55, h: 0.45, fill: { color: s.color }, rectRadius: 0.06 });
    slide.addText(s.num, { x: 0.8, y: sy, w: 0.55, h: 0.45, fontSize: 14, fontFace: "Arial", color: C.WHITE, bold: true, align: "center", valign: "middle", margin: 0 });
    slide.addText(s.title, { x: 1.5, y: sy + 0.02, w: 5.5, h: 0.22, fontSize: 12, fontFace: "Arial", color: C.WHITE, bold: true, margin: 0 });
    slide.addText(s.timeline, { x: 1.5, y: sy + 0.25, w: 3, h: 0.18, fontSize: 9, fontFace: "Arial", color: "94A3B8", margin: 0 });
  });

  // Bottom message
  slide.addShape(pptx.shapes.RECTANGLE, { x: 0.8, y: 4.6, w: 8.4, h: 0.04, fill: { color: "5A9BE6" } });
  slide.addText("データドリブンな品質管理で、CS向上と売上増加の好循環を実現する", {
    x: 0.8, y: 4.8, w: 8.4, h: 0.35, fontSize: 14, fontFace: "Arial", color: "5A9BE6", bold: true, italics: true, margin: 0,
  });

  addFooter(slide, 14);
})();

// ============================================================
// Save
// ============================================================
const OUTPUT = path.resolve(__dirname, "PRIMECHANGE_CS向上戦略提案書.pptx");
pptx.writeFile({ fileName: OUTPUT }).then(() => {
  const stats = fs.statSync(OUTPUT);
  console.log(`✅ PPTX: ${OUTPUT} (${(stats.size / 1024).toFixed(1)} KB)`);
}).catch(err => {
  console.error("Error:", err);
  process.exit(1);
});
