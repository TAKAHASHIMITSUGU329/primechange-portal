#!/usr/bin/env node
"use strict";
const fs = require("fs");
const path = require("path");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType, PageNumber, PageBreak } = require("docx");
const pptxgen = require("pptxgenjs");

const data = JSON.parse(fs.readFileSync(path.resolve(__dirname, "analysis_7_data.json"), "utf-8"));
const meta = data.analysis_metadata;
const scorecards = data.hotel_scorecards;
const bp_data = data.best_practices;
const ranked = bp_data.hotels_ranked;
const diffFactors = bp_data.differentiating_factors;
const roadmap = data.implementation_roadmap;
const crossInsights = data.cross_analysis_insights;
const recs = data.recommendations;

const C = { NAVY: "1B3A5C", ACCENT: "2E75B6", WHITE: "FFFFFF", LIGHT_BG: "F5F7FA", TEXT: "333333", SUBTEXT: "666666", GREEN: "27AE60", ORANGE: "FF9800", RED: "E74C3C", TEAL: "00695C" };
const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };
const cm = { top: 60, bottom: 60, left: 100, right: 100 };
function h1(t) { return new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 400, after: 200 }, children: [new TextRun({ text: t, bold: true, size: 32, font: "Arial", color: C.NAVY })] }); }
function h2(t) { return new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 300, after: 160 }, children: [new TextRun({ text: t, bold: true, size: 26, font: "Arial", color: C.ACCENT })] }); }
function h3(t) { return new Paragraph({ heading: HeadingLevel.HEADING_3, spacing: { before: 200, after: 120 }, children: [new TextRun({ text: t, bold: true, size: 22, font: "Arial", color: C.NAVY })] }); }
function p(t, o = {}) { return new Paragraph({ spacing: { before: o.sb || 80, after: o.sa || 80 }, alignment: o.align || AlignmentType.LEFT, children: [new TextRun({ text: t, size: o.sz || 20, font: "Arial", color: o.c || C.TEXT, bold: !!o.b, italics: !!o.i })] }); }
function bp(t, o = {}) { return new Paragraph({ spacing: { before: 40, after: 40 }, bullet: { level: o.level || 0 }, children: [new TextRun({ text: t, size: o.sz || 20, font: "Arial", color: o.c || C.TEXT, bold: !!o.b })] }); }
function cell(t, o = {}) { return new TableCell({ width: o.w ? { size: o.w, type: WidthType.DXA } : undefined, shading: o.bg ? { type: ShadingType.SOLID, color: o.bg, fill: o.bg } : undefined, borders, margins: cm, verticalAlign: "center", children: [new Paragraph({ alignment: o.a || AlignmentType.CENTER, children: [new TextRun({ text: String(t ?? "-"), size: o.sz || 16, font: "Arial", color: o.c || C.TEXT, bold: !!o.b })] })] }); }
function fmtN(n, d = 1) { return n != null ? Number(n).toFixed(d) : "-"; }
const PB = () => new Paragraph({ children: [new PageBreak()] });

async function buildDOCX() {
  // Hotel composite ranking table
  const rkHeader = new TableRow({ children: [
    cell("#", { bg: C.NAVY, c: C.WHITE, b: true, w: 500 }),
    cell("ホテル名", { bg: C.NAVY, c: C.WHITE, b: true, w: 2600 }),
    cell("口コミスコア", { bg: C.NAVY, c: C.WHITE, b: true, w: 1200 }),
    cell("クレーム率", { bg: C.NAVY, c: C.WHITE, b: true, w: 1100 }),
    cell("完了時刻", { bg: C.NAVY, c: C.WHITE, b: true, w: 1000 }),
    cell("安全スコア", { bg: C.NAVY, c: C.WHITE, b: true, w: 1100 }),
    cell("総合スコア", { bg: C.NAVY, c: C.WHITE, b: true, w: 1100 }),
  ] });
  const rkRows = ranked.map((h, i) => {
    const isTop5 = i < 5;
    const isBot5 = i >= ranked.length - 5;
    const scoreColor = isTop5 ? C.GREEN : isBot5 ? C.RED : C.TEXT;
    return new TableRow({ children: [
      cell(i + 1, { bg: i % 2 ? C.WHITE : C.LIGHT_BG, b: true }),
      cell(h.name, { bg: i % 2 ? C.WHITE : C.LIGHT_BG, a: AlignmentType.LEFT, sz: 14 }),
      cell(fmtN(h.review, 2), { bg: i % 2 ? C.WHITE : C.LIGHT_BG }),
      cell(fmtN(h.claims), { bg: i % 2 ? C.WHITE : C.LIGHT_BG }),
      cell(h.time != null ? fmtN(h.time, 2) : "-", { bg: i % 2 ? C.WHITE : C.LIGHT_BG }),
      cell(fmtN(h.safety, 2), { bg: i % 2 ? C.WHITE : C.LIGHT_BG }),
      cell(fmtN(h.composite, 3), { bg: i % 2 ? C.WHITE : C.LIGHT_BG, b: true, c: scoreColor }),
    ] });
  });
  const rkTable = new Table({ rows: [rkHeader, ...rkRows], width: { size: 8600, type: WidthType.DXA } });

  // Roadmap table
  const rmHeader = new TableRow({ children: [
    cell("フェーズ", { bg: C.ACCENT, c: C.WHITE, b: true, w: 2000 }),
    cell("施策", { bg: C.ACCENT, c: C.WHITE, b: true, w: 4500 }),
    cell("期待効果", { bg: C.ACCENT, c: C.WHITE, b: true, w: 2500 }),
  ] });
  const rmRows = roadmap.phases.map((ph, i) => new TableRow({ children: [
    cell(`${ph.phase}\n${ph.title}`, { bg: i % 2 ? C.WHITE : C.LIGHT_BG, a: AlignmentType.LEFT, b: true, sz: 14 }),
    cell(ph.actions.join("\n"), { bg: i % 2 ? C.WHITE : C.LIGHT_BG, a: AlignmentType.LEFT, sz: 14 }),
    cell(ph.expected_impact, { bg: i % 2 ? C.WHITE : C.LIGHT_BG, a: AlignmentType.LEFT, sz: 14 }),
  ] }));
  const rmTable = new Table({ rows: [rmHeader, ...rmRows], width: { size: 9000, type: WidthType.DXA } });

  const zeroClaimHotels = scorecards.filter(h => h.zero_claims).map(h => h.name);
  const avgComposite = ranked.reduce((s, h) => s + h.composite, 0) / ranked.length;

  const doc = new Document({
    styles: { default: { document: { run: { font: "Arial", size: 20 } } } },
    sections: [{
      headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "PRIMECHANGE | 分析7: ベストプラクティス横展開分析", size: 14, color: C.SUBTEXT, font: "Arial" })] })] }) },
      footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Page ", size: 14, color: C.SUBTEXT }), new TextRun({ children: [PageNumber.CURRENT], size: 14, color: C.SUBTEXT })] })] }) },
      children: [
        // Cover
        new Paragraph({ spacing: { before: 2400 } }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 }, children: [new TextRun({ text: "PRIMECHANGE", size: 56, font: "Arial", bold: true, color: C.NAVY })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [new TextRun({ text: "分析7", size: 44, font: "Arial", bold: true, color: C.ACCENT })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [new TextRun({ text: "ベストプラクティス横展開分析レポート", size: 36, font: "Arial", bold: true, color: C.ACCENT })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 400 }, children: [new TextRun({ text: "Best Practice Transfer Analysis", size: 22, font: "Arial", color: C.SUBTEXT, italics: true })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "株式会社PRIMECHANGE", size: 28, font: "Arial", bold: true, color: C.NAVY })] }),
        PB(),

        // 1. Executive Summary
        h1("1. エグゼクティブサマリー"),
        p("本レポートは分析1〜6の全成果を統合し、19ホテルのベストプラクティスを体系的に整理・横展開するための総合分析です。"),
        h3("6分析の統合結果"),
        bp("分析1（クレーム分析）: セット漏れ・残置が全クレームの55.9%を占め、集中対策で半数以上を解消可能。", { b: true }),
        bp("分析2（口コミ分析）: 清掃関連の否定的言及がスコアに直結。清掃品質の可視化が急務。"),
        bp("分析3（人員配置×品質）: メイド「数」よりチェッカー配置比率が品質を左右（r=-0.42）。"),
        bp("分析4（安全チェック）: 安全スコア上位ホテルは現場管理体制が成熟。全社平均は改善余地あり。"),
        bp("分析5（時系列分析）: 清掃完了時刻と品質に相関。早期完了ホテルがゲスト満足度で優位。"),
        bp("分析6（収益インパクト）: スコア0.1点改善でRevPAR 2.48%向上。年間¥1.4億の改善ポテンシャル。"),
        h3("総合スコアサマリー"),
        bp(`全19ホテル平均総合スコア: ${fmtN(avgComposite, 3)}`, { b: true }),
        bp(`トップ5平均: ${fmtN(bp_data.top5_avg_score, 2)}（口コミ）`, { b: true }),
        bp(`ボトム5平均: ${fmtN(bp_data.bottom5_avg_score, 2)}（口コミ）`),
        bp(`ゼロクレーム達成: ${zeroClaimHotels.length}ホテル（${zeroClaimHotels.join("、")}）`, { b: true }),
        PB(),

        // 2. Hotel Composite Ranking
        h1("2. ホテル総合ランキング"),
        p("口コミスコア・クレーム率・完了時刻・安全スコアを統合した総合スコアで19ホテルを順位付けしました。"),
        rkTable,
        p(""),
        p("※ 総合スコアは口コミスコア・クレーム率（逆数）・完了時刻（逆数）・安全スコアを正規化し加重平均したものです。", { sz: 16, c: C.SUBTEXT, i: true }),
        PB(),

        // 3. Differentiating Factors
        h1("3. 差別化要因分析"),
        p("上位ホテルと下位ホテルを分ける4つの差別化要因を特定しました。"),
        ...diffFactors.flatMap((df, i) => [
          h2(`3.${i + 1} ${df.factor}`),
          p(df.description),
          p("該当ホテル:", { b: true }),
          ...df.hotels.map(h => bp(h, { level: 1 })),
          p("横展開アクション:", { b: true }),
          ...df.transferable_actions.map(a => bp(a, { level: 1 })),
        ]),
        PB(),

        // 4. Cross-Analysis Insights
        h1("4. クロス分析インサイト"),
        p("6つの分析を横断して得られた3つの重要な洞察です。"),
        ...crossInsights.flatMap((ins, i) => [
          h2(`4.${i + 1} ${ins.title}`),
          p(`【発見】${ins.finding}`, { b: true }),
          p(`【示唆】${ins.implication}`),
        ]),
        PB(),

        // 5. Implementation Roadmap
        h1("5. 実行ロードマップ"),
        p("12ヶ月間の4フェーズ実行計画です。"),
        rmTable,
        PB(),

        // 6. Recommendations
        h1("6. 提言"),
        ...recs.flatMap((rec, i) => [
          h2(`6.${i + 1} ${rec.title}`),
          p(`【優先度: ${rec.priority}】`, { b: true, c: rec.priority === "最優先" ? C.RED : C.ORANGE }),
          ...rec.actions.map(a => bp(a)),
        ]),
      ]
    }]
  });

  const buf = await Packer.toBuffer(doc);
  const outPath = path.resolve(__dirname, "PRIMECHANGE_分析7_ベストプラクティス横展開.docx");
  fs.writeFileSync(outPath, buf);
  console.log(`DOCX: ${outPath} (${(buf.length / 1024).toFixed(1)} KB)`);
}

function buildPPTX() {
  const pptx = new pptxgen();
  pptx.layout = "LAYOUT_16x9";
  const bgDark = { fill: C.NAVY };
  const titleOpts = { x: 0.5, y: 0.3, w: 9, h: 0.6, fontSize: 22, bold: true, color: C.NAVY, fontFace: "Arial" };

  // Slide 1: Title
  const s1 = pptx.addSlide(); s1.background = bgDark;
  s1.addText("PRIMECHANGE", { x: 0.5, y: 1.2, w: 9, h: 0.8, fontSize: 40, bold: true, color: C.WHITE, fontFace: "Arial", align: "center" });
  s1.addText("分析7: ベストプラクティス横展開分析", { x: 0.5, y: 2.2, w: 9, h: 0.6, fontSize: 24, bold: true, color: C.WHITE, fontFace: "Arial", align: "center" });
  s1.addText(`${meta.total_hotels}ホテル × ${meta.analyses_integrated}分析統合 | ${meta.data_period}`, { x: 0.5, y: 3.4, w: 9, h: 0.3, fontSize: 12, color: "CCCCCC", fontFace: "Arial", align: "center" });

  // Slide 2: KPI Cards
  const s2 = pptx.addSlide();
  s2.addText("主要KPI", { ...titleOpts });
  const kpis = [
    { label: "トップ5 口コミ平均", value: fmtN(bp_data.top5_avg_score, 2), sub: "ベストプラクティス群", color: C.GREEN },
    { label: "ボトム5 口コミ平均", value: fmtN(bp_data.bottom5_avg_score, 2), sub: "改善対象群", color: C.RED },
    { label: "統合分析数", value: String(meta.analyses_integrated), sub: "全分析を横断統合", color: C.ACCENT },
    { label: "ゼロクレーム", value: `${scorecards.filter(h => h.zero_claims).length}ホテル`, sub: "クレームゼロ達成", color: C.TEAL },
  ];
  kpis.forEach((k, i) => {
    const x = 0.3 + i * 2.35;
    s2.addShape(pptx.ShapeType.roundRect, { x, y: 1.4, w: 2.15, h: 1.6, fill: { color: k.color }, rectRadius: 0.1 });
    s2.addText(k.label, { x, y: 1.5, w: 2.15, h: 0.3, fontSize: 10, color: C.WHITE, fontFace: "Arial", align: "center" });
    s2.addText(k.value, { x, y: 1.9, w: 2.15, h: 0.5, fontSize: 24, bold: true, color: C.WHITE, fontFace: "Arial", align: "center" });
    s2.addText(k.sub, { x, y: 2.5, w: 2.15, h: 0.3, fontSize: 9, color: C.WHITE, fontFace: "Arial", align: "center" });
  });
  s2.addText("Top5: " + bp_data.top5_best_practice.join(", "), { x: 0.5, y: 3.4, w: 9, h: 0.3, fontSize: 9, color: C.GREEN, fontFace: "Arial" });
  s2.addText("Bottom5: " + bp_data.bottom5_improvement.join(", "), { x: 0.5, y: 3.8, w: 9, h: 0.3, fontSize: 9, color: C.RED, fontFace: "Arial" });

  // Slide 3: Hotel composite ranking table
  const s3 = pptx.addSlide();
  s3.addText("ホテル総合ランキング", { ...titleOpts });
  const tblRows = [
    [{ text: "#", options: { bold: true, color: C.WHITE, fill: { color: C.NAVY }, fontSize: 7 } },
     { text: "ホテル名", options: { bold: true, color: C.WHITE, fill: { color: C.NAVY }, fontSize: 7 } },
     { text: "口コミ", options: { bold: true, color: C.WHITE, fill: { color: C.NAVY }, fontSize: 7 } },
     { text: "クレーム率", options: { bold: true, color: C.WHITE, fill: { color: C.NAVY }, fontSize: 7 } },
     { text: "完了時刻", options: { bold: true, color: C.WHITE, fill: { color: C.NAVY }, fontSize: 7 } },
     { text: "安全", options: { bold: true, color: C.WHITE, fill: { color: C.NAVY }, fontSize: 7 } },
     { text: "総合", options: { bold: true, color: C.WHITE, fill: { color: C.NAVY }, fontSize: 7 } }],
    ...ranked.map((h, i) => {
      const isTop5 = i < 5;
      const isBot5 = i >= ranked.length - 5;
      const scoreColor = isTop5 ? C.GREEN : isBot5 ? C.RED : C.TEXT;
      return [
        { text: String(i + 1), options: { fontSize: 7, fill: { color: i % 2 ? "FFFFFF" : "F5F7FA" } } },
        { text: h.name.substring(0, 20), options: { fontSize: 6, fill: { color: i % 2 ? "FFFFFF" : "F5F7FA" }, align: "left" } },
        { text: fmtN(h.review, 2), options: { fontSize: 7, fill: { color: i % 2 ? "FFFFFF" : "F5F7FA" } } },
        { text: fmtN(h.claims), options: { fontSize: 7, fill: { color: i % 2 ? "FFFFFF" : "F5F7FA" } } },
        { text: h.time != null ? fmtN(h.time, 1) : "-", options: { fontSize: 7, fill: { color: i % 2 ? "FFFFFF" : "F5F7FA" } } },
        { text: fmtN(h.safety, 2), options: { fontSize: 7, fill: { color: i % 2 ? "FFFFFF" : "F5F7FA" } } },
        { text: fmtN(h.composite, 3), options: { fontSize: 7, bold: true, color: scoreColor, fill: { color: i % 2 ? "FFFFFF" : "F5F7FA" } } },
      ];
    })
  ];
  s3.addTable(tblRows, { x: 0.2, y: 1.0, w: 9.6, h: 4.0, fontSize: 7, fontFace: "Arial", border: { color: "CCCCCC", pt: 0.5 }, colW: [0.4, 2.8, 1.0, 1.0, 1.0, 0.8, 1.0], align: "center" });

  // Slide 4: Differentiating Factors
  const s4 = pptx.addSlide();
  s4.addText("差別化要因", { ...titleOpts });
  const factorColors = [C.GREEN, C.ACCENT, C.ORANGE, C.TEAL];
  diffFactors.forEach((df, i) => {
    const y = 1.2 + i * 1.05;
    s4.addShape(pptx.ShapeType.roundRect, { x: 0.3, y, w: 9.2, h: 0.9, fill: { color: factorColors[i] || C.SUBTEXT }, rectRadius: 0.08 });
    s4.addText(df.factor, { x: 0.5, y: y + 0.05, w: 4.0, h: 0.3, fontSize: 11, bold: true, color: C.WHITE, fontFace: "Arial" });
    s4.addText(df.description, { x: 0.5, y: y + 0.35, w: 4.5, h: 0.3, fontSize: 8, color: C.WHITE, fontFace: "Arial" });
    s4.addText(df.hotels.slice(0, 3).join(", ") + (df.hotels.length > 3 ? " ..." : ""), { x: 5.2, y: y + 0.05, w: 4.0, h: 0.3, fontSize: 8, color: C.WHITE, fontFace: "Arial" });
    s4.addText(df.transferable_actions[0], { x: 5.2, y: y + 0.4, w: 4.0, h: 0.4, fontSize: 7, color: C.WHITE, fontFace: "Arial" });
  });

  // Slide 5: Implementation Roadmap
  const s5 = pptx.addSlide();
  s5.addText("実行ロードマップ (12ヶ月)", { ...titleOpts });
  const phaseColors = [C.RED, C.ORANGE, C.ACCENT, C.TEAL];
  roadmap.phases.forEach((ph, i) => {
    const x = 0.3 + i * 2.35;
    s5.addShape(pptx.ShapeType.roundRect, { x, y: 1.2, w: 2.15, h: 3.2, fill: { color: phaseColors[i] }, rectRadius: 0.1 });
    s5.addText(ph.phase, { x, y: 1.3, w: 2.15, h: 0.3, fontSize: 9, bold: true, color: C.WHITE, fontFace: "Arial", align: "center" });
    s5.addText(ph.title, { x, y: 1.6, w: 2.15, h: 0.3, fontSize: 12, bold: true, color: C.WHITE, fontFace: "Arial", align: "center" });
    ph.actions.forEach((a, ai) => {
      s5.addText("- " + a, { x: x + 0.1, y: 2.1 + ai * 0.4, w: 1.95, h: 0.4, fontSize: 7, color: C.WHITE, fontFace: "Arial" });
    });
    s5.addText(ph.expected_impact, { x: x + 0.1, y: 3.6, w: 1.95, h: 0.5, fontSize: 7, bold: true, color: C.WHITE, fontFace: "Arial", italic: true });
  });

  // Slide 6: Recommendations
  const s6 = pptx.addSlide();
  s6.addText("提言とアクションプラン", { ...titleOpts });
  const recColors = [C.RED, C.ORANGE, C.TEAL];
  recs.forEach((rec, i) => {
    const y = 1.2 + i * 1.2;
    s6.addShape(pptx.ShapeType.roundRect, { x: 0.3, y, w: 9.2, h: 1.0, fill: { color: recColors[i] || C.SUBTEXT }, rectRadius: 0.08 });
    s6.addText(`${rec.title}  [${rec.priority}]`, { x: 0.5, y: y + 0.05, w: 8.8, h: 0.35, fontSize: 12, bold: true, color: C.WHITE, fontFace: "Arial" });
    s6.addText(rec.actions.join(" / "), { x: 0.5, y: y + 0.45, w: 8.8, h: 0.45, fontSize: 9, color: C.WHITE, fontFace: "Arial" });
  });

  const outPath = path.resolve(__dirname, "PRIMECHANGE_分析7_ベストプラクティス横展開.pptx");
  pptx.writeFile({ fileName: outPath }).then(() => {
    console.log(`PPTX: ${outPath} (${(fs.statSync(outPath).size / 1024).toFixed(1)} KB)`);
  });
}

async function main() {
  console.log("Building Analysis 7 Reports...");
  await buildDOCX();
  await buildPPTX();
  console.log("\nDone!");
}
main().catch(e => { console.error(e); process.exit(1); });
