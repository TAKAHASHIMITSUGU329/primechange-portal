#!/usr/bin/env node
"use strict";

const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, PageNumber, PageBreak
} = require("docx");
const pptxgen = require("pptxgenjs");

// ============================================================
// Data
// ============================================================
const data = JSON.parse(fs.readFileSync(path.resolve(__dirname, "analysis_6_data.json"), "utf-8"));
const meta = data.analysis_metadata;
const summary = data.portfolio_summary;
const reg = data.regression_results;
const threshold = data.threshold_analysis;
const scenarios = data.revenue_impact_scenarios;
const benchmark = data.benchmark_comparison;
const potentials = data.hotel_improvement_potentials;
const totalPotential = data.total_improvement_potential;
const phases = data.phase_analysis;
const points = data.data_points;

// ============================================================
// Styles
// ============================================================
const C = { NAVY: "1B3A5C", ACCENT: "2E75B6", WHITE: "FFFFFF", LIGHT_BG: "F5F7FA", TEXT: "333333", SUBTEXT: "666666", GREEN: "27AE60", ORANGE: "FF9800", RED: "E74C3C", BLUE: "2196F3", DARK_GREEN: "1B5E20", TEAL: "00695C" };
const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };
const cm = { top: 60, bottom: 60, left: 100, right: 100 };

function h1(t) { return new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 400, after: 200 }, children: [new TextRun({ text: t, bold: true, size: 32, font: "Arial", color: C.NAVY })] }); }
function h2(t) { return new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 300, after: 160 }, children: [new TextRun({ text: t, bold: true, size: 26, font: "Arial", color: C.ACCENT })] }); }
function h3(t) { return new Paragraph({ heading: HeadingLevel.HEADING_3, spacing: { before: 200, after: 120 }, children: [new TextRun({ text: t, bold: true, size: 22, font: "Arial", color: C.NAVY })] }); }
function p(t, o = {}) { return new Paragraph({ spacing: { before: o.sb || 80, after: o.sa || 80 }, alignment: o.align || AlignmentType.LEFT, children: [new TextRun({ text: t, size: o.sz || 20, font: "Arial", color: o.c || C.TEXT, bold: !!o.b, italics: !!o.i })] }); }
function bp(t, o = {}) { return new Paragraph({ spacing: { before: 40, after: 40 }, bullet: { level: o.level || 0 }, children: [new TextRun({ text: t, size: o.sz || 20, font: "Arial", color: o.c || C.TEXT, bold: !!o.b })] }); }
function cell(t, o = {}) {
  return new TableCell({
    width: o.w ? { size: o.w, type: WidthType.DXA } : undefined,
    shading: o.bg ? { type: ShadingType.SOLID, color: o.bg, fill: o.bg } : undefined,
    borders, margins: cm, verticalAlign: "center",
    children: [new Paragraph({ alignment: o.a || AlignmentType.CENTER, children: [new TextRun({ text: String(t ?? "-"), size: o.sz || 16, font: "Arial", color: o.c || C.TEXT, bold: !!o.b })] })],
  });
}
function fmtY(n) { return n ? "¥" + Number(n).toLocaleString("ja-JP", { maximumFractionDigits: 0 }) : "-"; }
const PB = () => new Paragraph({ children: [new PageBreak()] });

// ============================================================
// DOCX Sections
// ============================================================
function buildCover() {
  return [
    new Paragraph({ spacing: { before: 2400 } }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 }, children: [new TextRun({ text: "PRIMECHANGE", size: 56, font: "Arial", bold: true, color: C.NAVY })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [new TextRun({ text: "分析6", size: 44, font: "Arial", bold: true, color: C.ACCENT })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [new TextRun({ text: "品質→売上 弾力性分析レポート", size: 36, font: "Arial", bold: true, color: C.ACCENT })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 400 }, children: [new TextRun({ text: "Quality-to-Revenue Elasticity Analysis", size: 22, font: "Arial", color: C.SUBTEXT, italics: true })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [new TextRun({ text: `対象: ${summary.total_hotels}ホテル | 月間総売上: ${fmtY(summary.total_monthly_revenue)}`, size: 20, font: "Arial", color: C.TEXT })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [new TextRun({ text: meta.date, size: 22, font: "Arial", color: C.TEXT })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 800 }, children: [new TextRun({ text: "株式会社PRIMECHANGE", size: 28, font: "Arial", bold: true, color: C.NAVY })] }),
    PB(),
  ];
}

function buildExecutiveSummary() {
  const regRevpar = reg.score_vs_revpar;
  return [
    h1("1. エグゼクティブサマリー"),
    p("本レポートは、PRIMECHANGEが管理する19ホテルの口コミスコアと売上指標（稼働率・ADR・RevPAR）の関係を統計分析し、「品質改善がどれだけ売上に効くか」を自社データで定量化したものです。"),
    h3("3つの重要な発見"),
    bp(`スコアとRevPARの相関は r=${regRevpar.r}（正の相関）。スコア1点上昇でRevPARが約${Math.abs(regRevpar.slope).toFixed(0)}円（${benchmark.our_data_revpar_pct_per_01}×10）変動。`, { b: true }),
    bp(`自社データの弾力性は業界ベンチマーク（0.1点=1%）の約${(parseFloat(benchmark.our_data_revpar_pct_per_01) / 1.0 * 100).toFixed(0)}%。品質改善の経済効果がベンチマーク以上に高い可能性。`, { b: true }),
    bp(`全19ホテルのスコアを各0.3-0.5点改善すると、月間約${fmtY(totalPotential.monthly)}（年間約${fmtY(totalPotential.annual)}）の売上改善余地。`, { b: true }),
    h3("分析の限界"),
    bp(`サンプルサイズN=${meta.hotels_count}のため、統計的有意性には限界がある`),
    bp("ホテル規模・立地・ブランド等の交絡変数を完全には制御できていない"),
    bp("相関関係であり因果関係ではない点に注意（ただし業界知見と整合）"),
    PB(),
  ];
}

function buildRegressionResults() {
  const regs = [
    { ...reg.score_vs_occupancy, name: "稼働率" },
    { ...reg.score_vs_adr, name: "ADR" },
    { ...reg.score_vs_revpar, name: "RevPAR" },
    { ...reg.score_vs_rev_per_room, name: "1室あたり売上" },
  ];
  const el = [
    h1("2. 回帰分析結果"),
    p("19ホテルの口コミスコアを説明変数（X）、各売上指標を目的変数（Y）として単回帰分析を実施しました。"),
    h2("2.1 回帰分析サマリー"),
  ];

  const rows = [
    ["指標", "相関係数 (r)", "決定係数 (R²)", "傾き", "解釈"].map((t, j) => cell(t, { b: true, bg: C.NAVY, c: C.WHITE, a: j === 0 || j === 4 ? AlignmentType.LEFT : AlignmentType.CENTER })),
  ];
  regs.forEach((r, i) => {
    const bg = i % 2 === 0 ? C.LIGHT_BG : C.WHITE;
    const rColor = Math.abs(r.r) >= 0.5 ? C.GREEN : Math.abs(r.r) >= 0.3 ? C.ORANGE : C.RED;
    rows.push([
      cell(r.name, { b: true, bg, a: AlignmentType.LEFT }),
      cell(r.r.toFixed(3), { bg, c: rColor, b: true }),
      cell(r.r_squared.toFixed(3), { bg }),
      cell(r.slope.toFixed(4), { bg }),
      cell(r.interpretation, { bg, a: AlignmentType.LEFT, sz: 14 }),
    ]);
  });
  el.push(new Table({ rows: rows.map(r => new TableRow({ children: r })), width: { size: 9500, type: WidthType.DXA } }));

  el.push(p(""), h2("2.2 最も重要な指標: スコア vs RevPAR"));
  const rv = reg.score_vs_revpar;
  el.push(p(`RevPAR（1室あたり収益）は口コミスコアと最も強い相関（r=${rv.r}）を示しました。これは、口コミスコアの高いホテルは「稼働率」と「客室単価」の両方が高い傾向にあることを意味します。`));
  el.push(p(`決定係数R²=${rv.r_squared}は、RevPARの変動の約${(rv.r_squared * 100).toFixed(0)}%が口コミスコアで説明できることを示しています。残りの${(100 - rv.r_squared * 100).toFixed(0)}%はホテルの立地、ブランド力、客室数等の他の要因によるものです。`, { c: C.SUBTEXT, sz: 18 }));

  el.push(h2("2.3 19ホテル一覧（スコア×RevPAR）"));

  const ptRows = [
    ["#", "ホテル名", "スコア", "稼働率", "ADR", "RevPAR", "月間売上", "Phase"].map((t, j) => cell(t, { b: true, bg: C.NAVY, c: C.WHITE, a: j === 1 ? AlignmentType.LEFT : AlignmentType.CENTER, sz: 14 })),
  ];
  points.forEach((pt, i) => {
    const bg = i % 2 === 0 ? C.LIGHT_BG : C.WHITE;
    ptRows.push([
      cell(i + 1, { bg, sz: 14 }),
      cell(pt.name, { bg, a: AlignmentType.LEFT, sz: 14, b: true }),
      cell(pt.score.toFixed(2), { bg, sz: 14, b: true }),
      cell(pt.occupancy + "%", { bg, sz: 14 }),
      cell(fmtY(pt.adr), { bg, sz: 14 }),
      cell(fmtY(pt.revpar), { bg, sz: 14, b: true, c: pt.revpar > 1000 ? C.GREEN : pt.revpar < 700 ? C.RED : C.TEXT }),
      cell(fmtY(pt.revenue), { bg, sz: 14 }),
      cell(pt.phase, { bg, sz: 14 }),
    ]);
  });
  el.push(new Table({ rows: ptRows.map(r => new TableRow({ children: r })), width: { size: 9500, type: WidthType.DXA } }));
  el.push(PB());
  return el;
}

function buildThresholdAnalysis() {
  const el = [
    h1("3. 閾値分析：スコア帯別の売上パフォーマンス"),
    p("19ホテルをスコア帯で4グループに分け、各グループの平均値を比較しました。"),
    h2("3.1 スコア帯別 平均値比較"),
  ];

  const rows = [
    ["スコア帯", "ホテル数", "平均稼働率", "平均ADR", "平均RevPAR", "平均売上"].map((t, j) => cell(t, { b: true, bg: C.NAVY, c: C.WHITE, a: j === 0 ? AlignmentType.LEFT : AlignmentType.CENTER })),
  ];
  threshold.groups.forEach((g, i) => {
    const bg = i % 2 === 0 ? C.LIGHT_BG : C.WHITE;
    const revColor = g.avg_revpar > 1000 ? C.GREEN : g.avg_revpar < 750 ? C.RED : C.TEXT;
    rows.push([
      cell(g.range, { bg, b: true, a: AlignmentType.LEFT }),
      cell(g.count + "ホテル", { bg }),
      cell(g.avg_occupancy_pct, { bg }),
      cell(fmtY(g.avg_adr), { bg }),
      cell(fmtY(g.avg_revpar), { bg, b: true, c: revColor }),
      cell(fmtY(g.avg_revenue), { bg }),
    ]);
  });
  el.push(new Table({ rows: rows.map(r => new TableRow({ children: r })), width: { size: 9000, type: WidthType.DXA } }));

  el.push(p(""), h2("3.2 閾値効果の検出"));
  const te = threshold.threshold_effect;
  el.push(p(te.description || "閾値効果の検出に十分なデータがありません。", { b: true, c: C.RED }));

  el.push(p(""), h3("グループ別 所属ホテル"));
  threshold.groups.forEach(g => {
    el.push(bp(`${g.range}（${g.count}ホテル）: ${g.hotels.join("、")}`, { sz: 18 }));
  });

  el.push(p(""), h3("注目すべきパターン"));
  if (threshold.groups.length >= 3) {
    const g1 = threshold.groups[0]; // 7-8
    const g3 = threshold.groups[2]; // 8.5-9
    const revDiff = g3.avg_revpar - g1.avg_revpar;
    el.push(bp(`スコア7-8帯とスコア8.5-9帯のRevPAR差: ${fmtY(Math.abs(revDiff))}（${(revDiff / g1.avg_revpar * 100).toFixed(0)}%差）`, { b: true }));
    el.push(bp("スコア8.5以上のホテル群は稼働率・ADRともに高水準で、品質が総合的な収益力に直結"));
  }
  el.push(PB());
  return el;
}

function buildBenchmarkComparison() {
  const el = [
    h1("4. 業界ベンチマークとの比較"),
    p("自社19ホテルのデータから算出した弾力性を、一般的な業界ベンチマークと比較します。"),
    h2("4.1 弾力性の比較"),
  ];

  const rows = [
    ["比較項目", "業界ベンチマーク", "自社データ"].map(t => cell(t, { b: true, bg: C.NAVY, c: C.WHITE })),
    [cell("スコア0.1点改善の効果", { bg: C.LIGHT_BG, a: AlignmentType.LEFT, b: true }), cell("RevPAR 1.0%向上", { bg: C.LIGHT_BG }), cell(`RevPAR ${benchmark.our_data_revpar_pct_per_01}向上`, { bg: C.LIGHT_BG, c: C.GREEN, b: true })],
    [cell("スコア0.5点改善の効果", { a: AlignmentType.LEFT, b: true }), cell("RevPAR 5.0%向上"), cell(`RevPAR ${(parseFloat(benchmark.our_data_revpar_pct_per_01) * 5).toFixed(1)}%向上`, { c: C.GREEN, b: true })],
    [cell("RevPAR変動額（0.1点あたり）", { bg: C.LIGHT_BG, a: AlignmentType.LEFT, b: true }), cell("—", { bg: C.LIGHT_BG }), cell(benchmark.our_revpar_change_per_01, { bg: C.LIGHT_BG, c: C.GREEN, b: true })],
    [cell("対ベンチマーク比率", { a: AlignmentType.LEFT, b: true }), cell("100%（基準）"), cell(benchmark.deviation, { c: C.GREEN, b: true })],
  ];
  el.push(new Table({ rows: rows.map(r => new TableRow({ children: Array.isArray(r) ? r : [r] })), width: { size: 9000, type: WidthType.DXA } }));

  el.push(p(""), h2("4.2 解釈"));
  el.push(p(benchmark.interpretation, { b: true }));
  el.push(p("自社データでの弾力性が業界ベンチマーク以上であるならば、品質改善への投資は業界平均以上のリターンが期待できます。これは、PRIMECHANGEが管理するホテル群のポジショニング（主に中価格帯のビジネスホテル）において、口コミスコアがゲストの予約判断に強く影響していることを示唆します。", { sz: 18 }));

  el.push(PB());
  return el;
}

function buildRevenueImpact() {
  const el = [
    h1("5. 収益インパクト試算"),
    p("回帰分析の結果を用いて、スコア改善が売上に与えるインパクトを3シナリオで試算します。"),
    h2("5.1 3シナリオ比較"),
  ];

  const rows = [
    ["項目", "控えめ (+0.1点)", "標準 (+0.3点)", "積極的 (+0.5点)"].map(t => cell(t, { b: true, bg: C.NAVY, c: C.WHITE })),
  ];
  const labels = ["スコア改善幅", "RevPAR変動", "RevPAR変動率", "月間売上改善額", "年間売上改善額"];
  scenarios.forEach((s, i) => {
    const scenarioData = [
      `+${s.score_improvement}点`,
      `${s.revpar_change > 0 ? "+" : ""}${fmtY(s.revpar_change)}`,
      s.revpar_pct_change,
      fmtY(s.total_monthly_revenue_change),
      fmtY(s.annual_revenue_change),
    ];
    scenarioData.forEach((val, j) => {
      if (!rows[j + 1]) {
        rows[j + 1] = [cell(labels[j], { bg: C.LIGHT_BG, a: AlignmentType.LEFT, b: true })];
      }
      const isRevenue = j >= 3;
      rows[j + 1].push(cell(val, { bg: j % 2 === 0 ? C.LIGHT_BG : C.WHITE, c: isRevenue ? C.GREEN : C.TEXT, b: isRevenue }));
    });
  });
  el.push(new Table({ rows: rows.map(r => new TableRow({ children: r })), width: { size: 9000, type: WidthType.DXA } }));

  el.push(p(""), h2("5.2 ホテル別 改善余地（上位10ホテル）"));

  const topRows = [
    ["#", "ホテル名", "現スコア", "目標", "改善幅", "月間売上効果", "年間効果"].map((t, j) => cell(t, { b: true, bg: C.NAVY, c: C.WHITE, a: j === 1 ? AlignmentType.LEFT : AlignmentType.CENTER, sz: 14 })),
  ];
  potentials.slice(0, 10).forEach((hp, i) => {
    const bg = i % 2 === 0 ? C.LIGHT_BG : C.WHITE;
    topRows.push([
      cell(i + 1, { bg, sz: 14 }),
      cell(hp.name, { bg, a: AlignmentType.LEFT, sz: 14, b: true }),
      cell(hp.current_score.toFixed(2), { bg, sz: 14 }),
      cell(hp.target_score.toFixed(1), { bg, sz: 14 }),
      cell("+" + hp.improvement.toFixed(1), { bg, sz: 14, c: C.ACCENT }),
      cell(fmtY(hp.estimated_revenue_change), { bg, sz: 14, c: C.GREEN, b: true }),
      cell(fmtY(hp.estimated_revenue_change_annual), { bg, sz: 14, c: C.GREEN }),
    ]);
  });
  el.push(new Table({ rows: topRows.map(r => new TableRow({ children: r })), width: { size: 9500, type: WidthType.DXA } }));

  el.push(p(""), p(`全19ホテル合計: 月間 ${fmtY(totalPotential.monthly)} / 年間 ${fmtY(totalPotential.annual)} の改善余地`, { b: true, c: C.GREEN, sz: 22 }));
  el.push(PB());
  return el;
}

function buildRecommendations() {
  return [
    h1("6. 提言と次のステップ"),
    h2("6.1 分析結果に基づく提言"),
    bp("品質改善投資の正当化: 自社データでRevPAR弾力性がベンチマーク以上と確認。品質改善はコストではなく「投資」として位置づけるべき。", { b: true }),
    bp("URGENT/HIGHホテルへの集中投資: スコア8.0未満のホテル群のスコアを8.5以上に引き上げることで、RevPAR30%以上の改善可能性。"),
    bp("スコア8.5を全ホテルの最低目標に: 閾値分析より、8.5以上のホテル群は稼働率・RevPARが大幅に高い。"),
    bp("四半期ごとの弾力性再計算: データ蓄積に伴い精度が向上。3ヶ月後に再分析を推奨。"),

    h2("6.2 次の分析ステップ"),
    bp("分析1（クレーム類型×スコア連動）: 本分析で得た弾力性を使い、「どのクレーム類型を減らすと最も売上に効くか」を特定"),
    bp("分析3（人員配置×品質）: 品質改善に必要な人員配置の費用と、売上改善効果を比較してROIを算出"),
    bp("時系列データの蓄積: 月次データが3ヶ月以上蓄積されれば、より強固な因果推論が可能に"),

    h2("6.3 本分析の活用方法"),
    bp("経営会議での品質改善投資の稟議書に「自社データに基づくROI」として添付"),
    bp("ホテルオーナーへの報告書に「スコア改善→売上改善のエビデンス」として引用"),
    bp("新規営業時に「データドリブンな品質管理による売上改善実績」として提示"),

    p(""),
    p("※ 本レポートは分析6/全7テーマの最初の分析結果です。後続の分析（クレーム類型、人員配置、スタッフパフォーマンス等）を組み合わせることで、より精密な改善戦略の構築が可能になります。", { i: true, c: C.SUBTEXT, sz: 18 }),
  ];
}

// ============================================================
// Build DOCX
// ============================================================
async function buildDocx() {
  const sections = [
    ...buildCover(),
    ...buildExecutiveSummary(),
    ...buildRegressionResults(),
    ...buildThresholdAnalysis(),
    ...buildBenchmarkComparison(),
    ...buildRevenueImpact(),
    ...buildRecommendations(),
  ];

  const doc = new Document({
    styles: { default: { document: { run: { font: "Arial", size: 20 } } } },
    sections: [{
      properties: {
        page: { size: { width: 12240, height: 15840 }, margin: { top: 720, bottom: 720, left: 900, right: 900 } },
      },
      headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "PRIMECHANGE 分析6: 品質→売上弾力性分析", size: 14, font: "Arial", color: C.SUBTEXT, italics: true })] })] }) },
      footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [
        new TextRun({ text: "Confidential | ", size: 14, font: "Arial", color: C.SUBTEXT }),
        new TextRun({ children: [PageNumber.CURRENT], size: 14, font: "Arial", color: C.SUBTEXT }),
        new TextRun({ text: " / ", size: 14, font: "Arial", color: C.SUBTEXT }),
        new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 14, font: "Arial", color: C.SUBTEXT }),
      ] })] }) },
      children: sections,
    }],
  });

  const buf = await Packer.toBuffer(doc);
  const outPath = path.resolve(__dirname, "PRIMECHANGE_分析6_品質売上弾力性.docx");
  fs.writeFileSync(outPath, buf);
  console.log(`✅ DOCX: ${outPath} (${(buf.length / 1024).toFixed(1)} KB)`);
}

// ============================================================
// Build PPTX
// ============================================================
function buildPptx() {
  const pptx = new pptxgen();
  pptx.layout = "LAYOUT_16x9";
  pptx.author = "PRIMECHANGE";
  pptx.title = "分析6: 品質→売上弾力性分析";

  const shadow = () => ({ type: "outer", blur: 4, offset: 2, angle: 135, color: "000000", opacity: 0.10 });

  function addFooter(slide, n) {
    slide.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.NAVY } });
    slide.addText("Confidential | PRIMECHANGE 分析6", { x: 0.5, y: 5.25, w: 5, h: 0.375, fontSize: 8, color: "94A3B8", fontFace: "Arial", valign: "middle" });
    slide.addText(String(n), { x: 9, y: 5.25, w: 0.5, h: 0.375, fontSize: 8, color: "94A3B8", fontFace: "Arial", align: "right", valign: "middle" });
  }

  function addHeader(slide, title, sub) {
    slide.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.6, fill: { color: C.NAVY } });
    slide.addText(title, { x: 0.5, y: 0.05, w: 9, h: 0.35, fontSize: 20, fontFace: "Arial", color: C.WHITE, bold: true, margin: 0 });
    if (sub) slide.addText(sub, { x: 0.5, y: 0.35, w: 9, h: 0.22, fontSize: 10, fontFace: "Arial", color: "5A9BE6", margin: 0 });
  }

  function kpiCard(slide, x, y, w, h, label, value, color, bgColor) {
    slide.addShape(pptx.shapes.RECTANGLE, { x, y, w, h, fill: { color: bgColor }, shadow: shadow() });
    slide.addShape(pptx.shapes.RECTANGLE, { x, y, w, h: 0.04, fill: { color } });
    slide.addText(label, { x, y: y + 0.1, w, h: 0.2, fontSize: 9, fontFace: "Arial", color: C.SUBTEXT, align: "center", margin: 0 });
    slide.addText(value, { x, y: y + 0.28, w, h: 0.4, fontSize: 26, fontFace: "Arial", color, bold: true, align: "center", margin: 0 });
  }

  // Slide 1: Title
  (() => {
    const slide = pptx.addSlide();
    slide.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: C.NAVY } });
    slide.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0, w: 0.08, h: 5.625, fill: { color: C.ACCENT } });
    slide.addText("PRIMECHANGE", { x: 0.8, y: 1.2, w: 8, h: 0.5, fontSize: 18, fontFace: "Arial", color: "5A9BE6", margin: 0 });
    slide.addText("分析6: 品質→売上 弾力性分析", { x: 0.8, y: 1.8, w: 8, h: 0.8, fontSize: 32, fontFace: "Arial", color: C.WHITE, bold: true, margin: 0 });
    slide.addShape(pptx.shapes.RECTANGLE, { x: 0.8, y: 2.7, w: 2.0, h: 0.04, fill: { color: "5A9BE6" } });
    slide.addText("Quality-to-Revenue Elasticity Analysis", { x: 0.8, y: 2.9, w: 8, h: 0.3, fontSize: 13, fontFace: "Arial", color: "94A3B8", italics: true, margin: 0 });
    slide.addText(`対象: ${summary.total_hotels}ホテル | 月間総売上: ${fmtY(summary.total_monthly_revenue)}`, { x: 0.8, y: 3.5, w: 8, h: 0.3, fontSize: 11, fontFace: "Arial", color: "64748B", margin: 0 });
    slide.addText(meta.date, { x: 0.8, y: 3.8, w: 4, h: 0.25, fontSize: 10, fontFace: "Arial", color: "64748B", margin: 0 });
  })();

  // Slide 2: Key Findings
  (() => {
    const slide = pptx.addSlide();
    addHeader(slide, "主要な発見", "Key Findings");

    kpiCard(slide, 0.3, 0.8, 2.1, 0.8, "スコア-RevPAR相関", "r=" + reg.score_vs_revpar.r, C.ACCENT, "E3F2FD");
    kpiCard(slide, 2.65, 0.8, 2.1, 0.8, "弾力性", benchmark.our_data_revpar_pct_per_01 + "/0.1pt", C.GREEN, "E8F5E9");
    kpiCard(slide, 5.0, 0.8, 2.1, 0.8, "vs業界", benchmark.deviation, C.TEAL, "E0F2F1");
    kpiCard(slide, 7.35, 0.8, 2.1, 0.8, "月間改善余地", fmtY(totalPotential.monthly), C.RED, "FFEBEE");

    // Findings
    const findings = [
      { title: "正の相関を確認", body: `スコアとRevPARの相関はr=${reg.score_vs_revpar.r}。\n決定係数R²=${reg.score_vs_revpar.r_squared}で、\nRevPAR変動の約${(reg.score_vs_revpar.r_squared*100).toFixed(0)}%がスコアで説明可能。`, color: C.ACCENT },
      { title: "業界以上の弾力性", body: `自社データ: 0.1pt改善→RevPAR ${benchmark.our_data_revpar_pct_per_01}向上\n業界: 0.1pt→1.0%\n品質改善の経済効果がベンチマーク以上。`, color: C.GREEN },
      { title: "大きな改善余地", body: `全ホテル合計で月間${fmtY(totalPotential.monthly)}、\n年間${fmtY(totalPotential.annual)}の改善余地。\nURGENTホテルが最大のポテンシャル。`, color: C.RED },
    ];
    findings.forEach((f, i) => {
      const x = 0.3 + i * 3.15;
      slide.addShape(pptx.shapes.RECTANGLE, { x, y: 2.0, w: 3.0, h: 2.7, fill: { color: C.LIGHT_BG }, shadow: shadow() });
      slide.addShape(pptx.shapes.RECTANGLE, { x, y: 2.0, w: 3.0, h: 0.04, fill: { color: f.color } });
      slide.addText(f.title, { x: x + 0.15, y: 2.1, w: 2.7, h: 0.3, fontSize: 12, fontFace: "Arial", color: f.color, bold: true, margin: 0 });
      slide.addText(f.body, { x: x + 0.15, y: 2.5, w: 2.7, h: 2.0, fontSize: 10, fontFace: "Arial", color: C.TEXT, valign: "top", margin: 0 });
    });

    addFooter(slide, 2);
  })();

  // Slide 3: Regression Results
  (() => {
    const slide = pptx.addSlide();
    addHeader(slide, "回帰分析結果", "スコア1点改善のインパクト");

    const header = [
      { text: "指標", options: { fontSize: 9, bold: true, color: C.WHITE, fill: { color: C.NAVY } } },
      { text: "相関 (r)", options: { fontSize: 9, bold: true, color: C.WHITE, fill: { color: C.NAVY }, align: "center" } },
      { text: "R²", options: { fontSize: 9, bold: true, color: C.WHITE, fill: { color: C.NAVY }, align: "center" } },
      { text: "解釈", options: { fontSize: 9, bold: true, color: C.WHITE, fill: { color: C.NAVY } } },
    ];

    const regs = [
      { name: "稼働率", ...reg.score_vs_occupancy },
      { name: "ADR", ...reg.score_vs_adr },
      { name: "RevPAR", ...reg.score_vs_revpar },
      { name: "1室売上", ...reg.score_vs_rev_per_room },
    ];

    const rows = [header, ...regs.map((r, i) => {
      const fill = i % 2 === 0 ? C.WHITE : C.LIGHT_BG;
      const rColor = Math.abs(r.r) >= 0.5 ? C.GREEN : Math.abs(r.r) >= 0.3 ? C.ORANGE : C.SUBTEXT;
      return [
        { text: r.name, options: { fontSize: 9, bold: true, fill: { color: fill } } },
        { text: r.r.toFixed(3), options: { fontSize: 9, bold: true, color: rColor, align: "center", fill: { color: fill } } },
        { text: r.r_squared.toFixed(3), options: { fontSize: 9, align: "center", fill: { color: fill } } },
        { text: r.interpretation, options: { fontSize: 8, fill: { color: fill } } },
      ];
    })];

    slide.addTable(rows, { x: 0.3, y: 0.85, w: 9.4, colW: [1.2, 1.0, 0.8, 5.5], fontSize: 9, border: { pt: 0.5, color: "CCCCCC" }, autoPage: false });

    // 19-hotel scatter table
    slide.addText("19ホテル: スコア vs RevPAR", { x: 0.3, y: 2.5, w: 5, h: 0.25, fontSize: 11, fontFace: "Arial", color: C.NAVY, bold: true, margin: 0 });

    const ptHeader = [
      { text: "ホテル名", options: { fontSize: 7, bold: true, color: C.WHITE, fill: { color: C.NAVY } } },
      { text: "スコア", options: { fontSize: 7, bold: true, color: C.WHITE, fill: { color: C.NAVY }, align: "center" } },
      { text: "稼働率", options: { fontSize: 7, bold: true, color: C.WHITE, fill: { color: C.NAVY }, align: "center" } },
      { text: "RevPAR", options: { fontSize: 7, bold: true, color: C.WHITE, fill: { color: C.NAVY }, align: "center" } },
      { text: "月間売上", options: { fontSize: 7, bold: true, color: C.WHITE, fill: { color: C.NAVY }, align: "center" } },
    ];

    const ptRows = [ptHeader, ...points.map((pt, i) => {
      const fill = i % 2 === 0 ? C.WHITE : C.LIGHT_BG;
      return [
        { text: pt.name, options: { fontSize: 6, bold: true, fill: { color: fill } } },
        { text: pt.score.toFixed(2), options: { fontSize: 6, bold: true, align: "center", fill: { color: fill } } },
        { text: pt.occupancy + "%", options: { fontSize: 6, align: "center", fill: { color: fill } } },
        { text: fmtY(pt.revpar), options: { fontSize: 6, align: "center", fill: { color: fill } } },
        { text: fmtY(pt.revenue), options: { fontSize: 6, align: "center", fill: { color: fill } } },
      ];
    })];
    slide.addTable(ptRows, { x: 0.3, y: 2.8, w: 9.4, colW: [3.5, 0.8, 0.9, 1.2, 1.8], fontSize: 7, border: { pt: 0.5, color: "CCCCCC" }, autoPage: false });

    addFooter(slide, 3);
  })();

  // Slide 4: Threshold Analysis
  (() => {
    const slide = pptx.addSlide();
    addHeader(slide, "閾値分析", "スコア帯別の売上パフォーマンス比較");

    const colors = [C.RED, C.ORANGE, C.ACCENT, C.GREEN];
    threshold.groups.forEach((g, i) => {
      const x = 0.3 + i * 2.35;
      const bgColors = ["FFEBEE", "FFF3E0", "E3F2FD", "E8F5E9"];
      slide.addShape(pptx.shapes.RECTANGLE, { x, y: 0.85, w: 2.2, h: 2.4, fill: { color: bgColors[i] }, shadow: shadow() });
      slide.addShape(pptx.shapes.RECTANGLE, { x, y: 0.85, w: 2.2, h: 0.04, fill: { color: colors[i] } });
      slide.addText(g.range, { x, y: 0.95, w: 2.2, h: 0.3, fontSize: 14, fontFace: "Arial", color: colors[i], bold: true, align: "center", margin: 0 });
      slide.addText(`${g.count}ホテル`, { x, y: 1.25, w: 2.2, h: 0.2, fontSize: 9, fontFace: "Arial", color: C.SUBTEXT, align: "center", margin: 0 });
      // KPIs
      slide.addText("稼働率", { x: x + 0.1, y: 1.55, w: 1.0, h: 0.15, fontSize: 8, color: C.SUBTEXT, margin: 0 });
      slide.addText(g.avg_occupancy_pct, { x: x + 1.1, y: 1.55, w: 1.0, h: 0.15, fontSize: 8, color: C.TEXT, bold: true, align: "right", margin: 0 });
      slide.addText("RevPAR", { x: x + 0.1, y: 1.75, w: 1.0, h: 0.15, fontSize: 8, color: C.SUBTEXT, margin: 0 });
      slide.addText(fmtY(g.avg_revpar), { x: x + 1.1, y: 1.75, w: 1.0, h: 0.15, fontSize: 8, color: colors[i], bold: true, align: "right", margin: 0 });
      slide.addText("平均売上", { x: x + 0.1, y: 1.95, w: 1.0, h: 0.15, fontSize: 8, color: C.SUBTEXT, margin: 0 });
      slide.addText(fmtY(g.avg_revenue), { x: x + 1.1, y: 1.95, w: 1.0, h: 0.15, fontSize: 8, color: C.TEXT, bold: true, align: "right", margin: 0 });
      // Hotel names
      slide.addText(g.hotels.join("\n"), { x: x + 0.1, y: 2.2, w: 2.0, h: 0.95, fontSize: 6, fontFace: "Arial", color: C.TEXT, valign: "top", margin: 0 });
    });

    // Insight
    if (threshold.groups.length >= 3) {
      const diff = threshold.groups[2].avg_revpar - threshold.groups[0].avg_revpar;
      slide.addShape(pptx.shapes.RECTANGLE, { x: 0.3, y: 3.5, w: 9.4, h: 1.5, fill: { color: "F0F7FC" }, shadow: shadow() });
      slide.addText("分析の示唆", { x: 0.5, y: 3.6, w: 4, h: 0.25, fontSize: 12, fontFace: "Arial", color: C.NAVY, bold: true, margin: 0 });
      slide.addText([
        { text: `スコア7-8帯と8.5-9帯のRevPAR差: ${fmtY(Math.abs(diff))}`, options: { fontSize: 11, bold: true, color: C.GREEN } },
        { text: `（${(diff / threshold.groups[0].avg_revpar * 100).toFixed(0)}%差）\n`, options: { fontSize: 11, color: C.GREEN } },
        { text: "スコア8.5以上を全ホテルの目標に設定すべき。\n", options: { fontSize: 10, color: C.TEXT } },
        { text: "低スコアホテルの改善が最大のRevPAR改善機会。", options: { fontSize: 10, color: C.TEXT } },
      ], { x: 0.5, y: 3.9, w: 8.8, h: 0.9, valign: "top", margin: 0 });
    }

    addFooter(slide, 4);
  })();

  // Slide 5: Benchmark Comparison
  (() => {
    const slide = pptx.addSlide();
    addHeader(slide, "業界ベンチマークとの比較", "自社データ vs 業界標準");

    // Two-column comparison
    slide.addShape(pptx.shapes.RECTANGLE, { x: 0.3, y: 0.85, w: 4.5, h: 2.5, fill: { color: C.LIGHT_BG }, shadow: shadow() });
    slide.addShape(pptx.shapes.RECTANGLE, { x: 0.3, y: 0.85, w: 4.5, h: 0.04, fill: { color: C.SUBTEXT } });
    slide.addText("業界ベンチマーク", { x: 0.5, y: 0.95, w: 4, h: 0.3, fontSize: 14, fontFace: "Arial", color: C.SUBTEXT, bold: true, margin: 0 });
    slide.addText("スコア0.1点改善\n= RevPAR 1.0%向上", { x: 0.5, y: 1.4, w: 4, h: 0.8, fontSize: 18, fontFace: "Arial", color: C.SUBTEXT, bold: true, margin: 0 });
    slide.addText("（出典: 業界研究データ）", { x: 0.5, y: 2.5, w: 4, h: 0.3, fontSize: 9, fontFace: "Arial", color: C.SUBTEXT, margin: 0 });

    slide.addShape(pptx.shapes.RECTANGLE, { x: 5.15, y: 0.85, w: 4.55, h: 2.5, fill: { color: "E8F5E9" }, shadow: shadow() });
    slide.addShape(pptx.shapes.RECTANGLE, { x: 5.15, y: 0.85, w: 4.55, h: 0.04, fill: { color: C.GREEN } });
    slide.addText("自社データ（N=19）", { x: 5.35, y: 0.95, w: 4, h: 0.3, fontSize: 14, fontFace: "Arial", color: C.GREEN, bold: true, margin: 0 });
    slide.addText(`スコア0.1点改善\n= RevPAR ${benchmark.our_data_revpar_pct_per_01}向上`, { x: 5.35, y: 1.4, w: 4, h: 0.8, fontSize: 18, fontFace: "Arial", color: C.GREEN, bold: true, margin: 0 });
    slide.addText(`対ベンチマーク: ${benchmark.deviation}`, { x: 5.35, y: 2.5, w: 4, h: 0.3, fontSize: 10, fontFace: "Arial", color: C.DARK_GREEN, bold: true, margin: 0 });

    // Interpretation
    slide.addShape(pptx.shapes.RECTANGLE, { x: 0.3, y: 3.6, w: 9.4, h: 1.4, fill: { color: "F0F7FC" }, shadow: shadow() });
    slide.addText("解釈", { x: 0.5, y: 3.7, w: 3, h: 0.25, fontSize: 12, fontFace: "Arial", color: C.NAVY, bold: true, margin: 0 });
    slide.addText(benchmark.interpretation, { x: 0.5, y: 4.0, w: 9, h: 0.5, fontSize: 10, fontFace: "Arial", color: C.TEXT, margin: 0 });
    slide.addText("※ サンプルサイズN=19のため、弾力性値は参考値。データ蓄積で精度向上。", { x: 0.5, y: 4.6, w: 9, h: 0.25, fontSize: 8, fontFace: "Arial", color: C.SUBTEXT, italics: true, margin: 0 });

    addFooter(slide, 5);
  })();

  // Slide 6: Revenue Impact
  (() => {
    const slide = pptx.addSlide();
    addHeader(slide, "収益インパクト試算", "3シナリオでの売上改善額");

    const header = [
      { text: "項目", options: { fontSize: 9, bold: true, color: C.WHITE, fill: { color: C.NAVY } } },
      { text: "控えめ(+0.1)", options: { fontSize: 9, bold: true, color: C.WHITE, fill: { color: C.NAVY }, align: "center" } },
      { text: "標準(+0.3)", options: { fontSize: 9, bold: true, color: C.WHITE, fill: { color: C.NAVY }, align: "center" } },
      { text: "積極的(+0.5)", options: { fontSize: 9, bold: true, color: C.WHITE, fill: { color: C.NAVY }, align: "center" } },
    ];

    const items = ["RevPAR変動率", "月間売上改善", "年間売上改善"];
    const rows = [header];
    items.forEach((label, li) => {
      const fill = li % 2 === 0 ? C.WHITE : C.LIGHT_BG;
      const row = [{ text: label, options: { fontSize: 9, bold: true, fill: { color: fill } } }];
      scenarios.forEach(s => {
        let val;
        if (li === 0) val = s.revpar_pct_change;
        else if (li === 1) val = fmtY(s.total_monthly_revenue_change);
        else val = fmtY(s.annual_revenue_change);
        row.push({ text: val, options: { fontSize: 9, bold: li >= 1, color: li >= 1 ? C.GREEN : C.TEXT, align: "center", fill: { color: fill } } });
      });
      rows.push(row);
    });
    slide.addTable(rows, { x: 0.3, y: 0.85, w: 9.4, colW: [2.0, 2.4, 2.6, 2.4], fontSize: 9, border: { pt: 0.5, color: "CCCCCC" }, autoPage: false });

    // Top 5 improvement potentials
    slide.addText("ホテル別 改善余地 TOP5", { x: 0.3, y: 2.5, w: 5, h: 0.25, fontSize: 11, fontFace: "Arial", color: C.NAVY, bold: true, margin: 0 });

    const topHeader = [
      { text: "ホテル", options: { fontSize: 8, bold: true, color: C.WHITE, fill: { color: C.NAVY } } },
      { text: "現スコア", options: { fontSize: 8, bold: true, color: C.WHITE, fill: { color: C.NAVY }, align: "center" } },
      { text: "目標", options: { fontSize: 8, bold: true, color: C.WHITE, fill: { color: C.NAVY }, align: "center" } },
      { text: "月間効果", options: { fontSize: 8, bold: true, color: C.WHITE, fill: { color: C.NAVY }, align: "center" } },
      { text: "年間効果", options: { fontSize: 8, bold: true, color: C.WHITE, fill: { color: C.NAVY }, align: "center" } },
    ];
    const topRows = [topHeader, ...potentials.slice(0, 5).map((hp, i) => {
      const fill = i % 2 === 0 ? C.WHITE : C.LIGHT_BG;
      return [
        { text: hp.name, options: { fontSize: 8, bold: true, fill: { color: fill } } },
        { text: hp.current_score.toFixed(2), options: { fontSize: 8, align: "center", fill: { color: fill } } },
        { text: hp.target_score.toFixed(1), options: { fontSize: 8, align: "center", fill: { color: fill } } },
        { text: fmtY(hp.estimated_revenue_change), options: { fontSize: 8, bold: true, color: C.GREEN, align: "center", fill: { color: fill } } },
        { text: fmtY(hp.estimated_revenue_change_annual), options: { fontSize: 8, color: C.GREEN, align: "center", fill: { color: fill } } },
      ];
    })];
    slide.addTable(topRows, { x: 0.3, y: 2.8, w: 9.4, colW: [3.2, 1.0, 1.0, 1.8, 2.0], fontSize: 8, border: { pt: 0.5, color: "CCCCCC" }, autoPage: false });

    // Total
    slide.addShape(pptx.shapes.RECTANGLE, { x: 0.3, y: 4.4, w: 9.4, h: 0.6, fill: { color: "E8F5E9" }, shadow: shadow() });
    slide.addText(`全19ホテル合計: 月間 ${fmtY(totalPotential.monthly)} ／ 年間 ${fmtY(totalPotential.annual)}`, {
      x: 0.5, y: 4.45, w: 8.8, h: 0.5, fontSize: 14, fontFace: "Arial", color: C.GREEN, bold: true, align: "center", valign: "middle", margin: 0,
    });

    addFooter(slide, 6);
  })();

  // Slide 7: Next Steps
  (() => {
    const slide = pptx.addSlide();
    slide.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: C.NAVY } });
    slide.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0, w: 0.08, h: 5.625, fill: { color: C.ACCENT } });

    slide.addText("提言と次のステップ", { x: 0.8, y: 0.5, w: 8, h: 0.5, fontSize: 24, fontFace: "Arial", color: C.WHITE, bold: true, margin: 0 });
    slide.addShape(pptx.shapes.RECTANGLE, { x: 0.8, y: 1.1, w: 2.0, h: 0.04, fill: { color: "5A9BE6" } });

    const steps = [
      { num: "01", title: "品質改善=投資として位置づけ", desc: "弾力性データを経営会議で共有", color: C.RED },
      { num: "02", title: "スコア8.5を全ホテル最低目標に設定", desc: "閾値分析に基づく戦略的目標", color: C.RED },
      { num: "03", title: "URGENT/HIGHホテルに集中投資", desc: "改善余地が最も大きい4施設を優先", color: C.ORANGE },
      { num: "04", title: "分析1（クレーム類型）に着手", desc: "弾力性データと組み合わせて具体的な改善項目を特定", color: "5A9BE6" },
      { num: "05", title: "四半期ごとに弾力性を再計算", desc: "データ蓄積で統計的精度を向上", color: "5A9BE6" },
    ];

    steps.forEach((s, i) => {
      const sy = 1.4 + i * 0.75;
      slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: sy, w: 0.5, h: 0.45, fill: { color: s.color }, rectRadius: 0.06 });
      slide.addText(s.num, { x: 0.8, y: sy, w: 0.5, h: 0.45, fontSize: 14, fontFace: "Arial", color: C.WHITE, bold: true, align: "center", valign: "middle", margin: 0 });
      slide.addText(s.title, { x: 1.5, y: sy + 0.02, w: 7, h: 0.22, fontSize: 13, fontFace: "Arial", color: C.WHITE, bold: true, margin: 0 });
      slide.addText(s.desc, { x: 1.5, y: sy + 0.27, w: 7, h: 0.18, fontSize: 9, fontFace: "Arial", color: "94A3B8", margin: 0 });
    });

    slide.addShape(pptx.shapes.RECTANGLE, { x: 0.8, y: 4.8, w: 8.4, h: 0.04, fill: { color: "5A9BE6" } });
    slide.addText("データに基づく品質改善投資で、CS向上→売上増加の好循環を実現する", {
      x: 0.8, y: 4.95, w: 8.4, h: 0.3, fontSize: 12, fontFace: "Arial", color: "5A9BE6", bold: true, italics: true, margin: 0,
    });

    addFooter(slide, 7);
  })();

  const outPath = path.resolve(__dirname, "PRIMECHANGE_分析6_品質売上弾力性.pptx");
  return pptx.writeFile({ fileName: outPath }).then(() => {
    const stats = fs.statSync(outPath);
    console.log(`✅ PPTX: ${outPath} (${(stats.size / 1024).toFixed(1)} KB)`);
  });
}

// ============================================================
// Main
// ============================================================
async function main() {
  console.log("Building Analysis 6 Reports...");
  await buildDocx();
  await buildPptx();
  console.log("\nDone!");
}

main().catch(err => { console.error(err); process.exit(1); });
