#!/usr/bin/env node
"use strict";
const fs = require("fs");
const path = require("path");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType, PageNumber, PageBreak } = require("docx");
const pptxgen = require("pptxgenjs");

const data = JSON.parse(fs.readFileSync(path.resolve(__dirname, "analysis_3_data.json"), "utf-8"));
const meta = data.analysis_metadata;
const sa = data.staffing_analysis;
const corrs = sa.correlations;
const tiers = sa.staffing_tiers;
const optimal = sa.optimal_staffing;
const hotels = data.hotel_summary;
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
  const hotelsWithData = hotels.filter(h => h.avg_maids);
  const hHeader = new TableRow({ children: [
    cell("ホテル名", { bg: C.NAVY, c: C.WHITE, b: true, w: 2800 }),
    cell("メイド", { bg: C.NAVY, c: C.WHITE, b: true, w: 1000 }),
    cell("チェッカー", { bg: C.NAVY, c: C.WHITE, b: true, w: 1000 }),
    cell("比率(M:C)", { bg: C.NAVY, c: C.WHITE, b: true, w: 1200 }),
    cell("スコア", { bg: C.NAVY, c: C.WHITE, b: true, w: 1000 }),
    cell("クレーム率", { bg: C.NAVY, c: C.WHITE, b: true, w: 1200 }),
    cell("データ日数", { bg: C.NAVY, c: C.WHITE, b: true, w: 900 }),
  ] });
  const hRows = hotelsWithData.map((h, i) => new TableRow({ children: [
    cell(h.name, { bg: i % 2 ? C.WHITE : C.LIGHT_BG, a: AlignmentType.LEFT, sz: 14 }),
    cell(fmtN(h.avg_maids), { bg: i % 2 ? C.WHITE : C.LIGHT_BG }),
    cell(h.avg_checkers ? fmtN(h.avg_checkers) : "-", { bg: i % 2 ? C.WHITE : C.LIGHT_BG }),
    cell(h.maid_checker_ratio ? `${fmtN(h.maid_checker_ratio)}:1` : "-", { bg: i % 2 ? C.WHITE : C.LIGHT_BG }),
    cell(fmtN(h.score, 2), { bg: i % 2 ? C.WHITE : C.LIGHT_BG, b: true }),
    cell(fmtN(h.claim_rate), { bg: i % 2 ? C.WHITE : C.LIGHT_BG }),
    cell(h.maid_data_points, { bg: i % 2 ? C.WHITE : C.LIGHT_BG }),
  ] }));
  const hTable = new Table({ rows: [hHeader, ...hRows], width: { size: 9100, type: WidthType.DXA } });

  // Correlation table
  const corrEntries = Object.entries(corrs);
  const cHeader = new TableRow({ children: [
    cell("相関ペア", { bg: C.ACCENT, c: C.WHITE, b: true, w: 3400 }),
    cell("r", { bg: C.ACCENT, c: C.WHITE, b: true, w: 1200 }),
    cell("R²", { bg: C.ACCENT, c: C.WHITE, b: true, w: 1200 }),
    cell("N", { bg: C.ACCENT, c: C.WHITE, b: true, w: 800 }),
    cell("判定", { bg: C.ACCENT, c: C.WHITE, b: true, w: 2400 }),
  ] });
  const labelMap = { 'staff_vs_score': '総スタッフ vs スコア', 'maids_vs_score': 'メイド数 vs スコア', 'checkers_vs_score': 'チェッカー数 vs スコア', 'ratio_vs_score': 'M:C比率 vs スコア', 'staff_vs_claims': '総スタッフ vs クレーム率' };
  const cRows = corrEntries.map(([k, v], i) => {
    const absR = Math.abs(v.r);
    const judge = absR >= 0.4 ? '中程度の相関' : absR >= 0.2 ? '弱い相関' : 'ほぼ無相関';
    const jColor = absR >= 0.4 ? C.RED : absR >= 0.2 ? C.ORANGE : C.TEXT;
    return new TableRow({ children: [
      cell(labelMap[k] || k, { bg: i % 2 ? C.WHITE : C.LIGHT_BG, a: AlignmentType.LEFT }),
      cell(fmtN(v.r, 4), { bg: i % 2 ? C.WHITE : C.LIGHT_BG, b: true }),
      cell(v.r_squared != null ? fmtN(v.r_squared, 4) : "-", { bg: i % 2 ? C.WHITE : C.LIGHT_BG }),
      cell(v.n, { bg: i % 2 ? C.WHITE : C.LIGHT_BG }),
      cell(judge, { bg: i % 2 ? C.WHITE : C.LIGHT_BG, c: jColor }),
    ] });
  });
  const cTable = new Table({ rows: [cHeader, ...cRows], width: { size: 9000, type: WidthType.DXA } });

  const doc = new Document({
    styles: { default: { document: { run: { font: "Arial", size: 20 } } } },
    sections: [{
      headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "PRIMECHANGE | 分析3: 人員配置×品質", size: 14, color: C.SUBTEXT, font: "Arial" })] })] }) },
      footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Page ", size: 14, color: C.SUBTEXT }), new TextRun({ children: [PageNumber.CURRENT], size: 14, color: C.SUBTEXT })] })] }) },
      children: [
        // Cover
        new Paragraph({ spacing: { before: 2400 } }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 }, children: [new TextRun({ text: "PRIMECHANGE", size: 56, font: "Arial", bold: true, color: C.NAVY })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [new TextRun({ text: "分析3", size: 44, font: "Arial", bold: true, color: C.ACCENT })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [new TextRun({ text: "人員配置×品質 相関分析レポート", size: 36, font: "Arial", bold: true, color: C.ACCENT })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 400 }, children: [new TextRun({ text: "Staffing Level × Quality Correlation Analysis", size: 22, font: "Arial", color: C.SUBTEXT, italics: true })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "株式会社PRIMECHANGE", size: 28, font: "Arial", bold: true, color: C.NAVY })] }),
        PB(),

        // Executive Summary
        h1("1. エグゼクティブサマリー"),
        p("19ホテルの日報データからメイド数・チェッカー数と口コミスコア・クレーム率の関係を統計分析しました。"),
        h3("主要な発見"),
        bp(`メイド:チェッカー比率とスコアに中程度の負の相関（r=${fmtN(corrs.ratio_vs_score?.r, 4)}）。比率が低い（チェッカーが相対的に多い）ほどスコアが高い傾向。`, { b: true }),
        bp(`チェッカー数とスコアに弱い正の相関（r=${fmtN(corrs.checkers_vs_score?.r, 4)}）。チェッカー配置がスコア向上に寄与する可能性。`, { b: true }),
        bp(`スタッフ総数とクレーム率に中程度の正の相関（r=${fmtN(corrs.staff_vs_claims?.r, 4)}）。大規模ホテルほど検知力が高い傾向を反映。`, { b: true }),
        bp(`メイド数そのものとスコアの相関はほぼゼロ（r=${fmtN(corrs.maids_vs_score?.r, 4)}）。「数」よりも「質」と「チェック体制」が品質を左右。`),
        PB(),

        // Correlation Results
        h1("2. 相関分析結果"),
        p("5つの相関ペアを分析し、人員配置と品質指標の関連を定量化しました。"),
        cTable,
        p(""),
        h3("注目すべき相関"),
        bp(`M:C比率 vs スコア（r=${fmtN(corrs.ratio_vs_score?.r, 4)}）: メイド1人あたりのチェッカー数が多いほどスコアが高い。チェッカー体制の充実が品質向上の鍵。`, { b: true }),
        bp(`総スタッフ vs クレーム率（r=${fmtN(corrs.staff_vs_claims?.r, 4)}）: スタッフが多いほどクレーム検出率が高い。品質管理体制の成熟度を反映。`),
        PB(),

        // Hotel Comparison
        h1("3. ホテル別人員配置一覧"),
        p("全19ホテルのメイド数・チェッカー数・比率と品質指標の一覧です。"),
        hTable,
        PB(),

        // Staffing Tier Analysis
        h1("4. 配置水準別の品質比較"),
        p("メイド数の中央値で2グループに分け、品質指標を比較しました。"),
        bp(`低配置グループ（${tiers.low_staff?.count}ホテル）: 平均${fmtN(tiers.low_staff?.avg_maids)}名 → スコア${fmtN(tiers.low_staff?.avg_score, 2)}`, { b: true }),
        bp(`高配置グループ（${tiers.high_staff?.count}ホテル）: 平均${fmtN(tiers.high_staff?.avg_maids)}名 → スコア${fmtN(tiers.high_staff?.avg_score, 2)}`, { b: true }),
        p(`両グループのスコア差は${Math.abs((tiers.high_staff?.avg_score||0) - (tiers.low_staff?.avg_score||0)).toFixed(2)}点と小さく、メイド「数」よりも「質」や「チェック体制」が品質を決定する重要な要因であることを示唆しています。`),

        h2("4.1 上位ホテル vs 下位ホテルの配置比較"),
        bp("スコア上位5ホテル:", { b: true }),
        ...optimal.top5_hotels.map(h => bp(`  ${h.name}: メイド${fmtN(h.avg_maids)}名, チェッカー${h.avg_checkers ? fmtN(h.avg_checkers) : "-"}名, 比率${h.ratio ? fmtN(h.ratio) + ":1" : "-"}, スコア${fmtN(h.score, 2)}`, { level: 1 })),
        bp("スコア下位5ホテル:", { b: true }),
        ...optimal.bottom5_hotels.map(h => bp(`  ${h.name}: メイド${fmtN(h.avg_maids)}名, チェッカー${h.avg_checkers ? fmtN(h.avg_checkers) : "-"}名, 比率${h.ratio ? fmtN(h.ratio) + ":1" : "-"}, スコア${fmtN(h.score, 2)}`, { level: 1 })),
        PB(),

        // Recommendations
        h1("5. 提言"),
        ...recs.flatMap((rec, i) => [
          h2(`5.${i + 1} ${rec.title}`),
          p(`【優先度: ${rec.priority}】`, { b: true, c: rec.priority === "最優先" ? C.RED : C.ORANGE }),
          p(rec.rationale),
          ...rec.actions.map(a => bp(a)),
        ]),
      ]
    }]
  });

  const buf = await Packer.toBuffer(doc);
  const outPath = path.resolve(__dirname, "PRIMECHANGE_分析3_人員配置品質相関.docx");
  fs.writeFileSync(outPath, buf);
  console.log(`✅ DOCX: ${outPath} (${(buf.length / 1024).toFixed(1)} KB)`);
}

function buildPPTX() {
  const pptx = new pptxgen();
  pptx.layout = "LAYOUT_16x9";
  const bgDark = { fill: C.NAVY };
  const titleOpts = { x: 0.5, y: 0.3, w: 9, h: 0.6, fontSize: 22, bold: true, color: C.NAVY, fontFace: "Arial" };

  // Slide 1: Title
  const s1 = pptx.addSlide(); s1.background = bgDark;
  s1.addText("PRIMECHANGE", { x: 0.5, y: 1.2, w: 9, h: 0.8, fontSize: 40, bold: true, color: C.WHITE, fontFace: "Arial", align: "center" });
  s1.addText("分析3: 人員配置×品質 相関分析", { x: 0.5, y: 2.2, w: 9, h: 0.6, fontSize: 24, bold: true, color: C.WHITE, fontFace: "Arial", align: "center" });
  s1.addText(`${meta.total_hotels}ホテル日報データに基づくスタッフ配置と品質の関連分析`, { x: 0.5, y: 3.4, w: 9, h: 0.3, fontSize: 12, color: "CCCCCC", fontFace: "Arial", align: "center" });

  // Slide 2: KPIs
  const s2 = pptx.addSlide();
  s2.addText("主要な発見", { ...titleOpts });
  const kpis = [
    { label: "M:C比率 vs スコア", value: `r=${fmtN(corrs.ratio_vs_score?.r, 3)}`, sub: "中程度の負の相関", color: C.RED },
    { label: "チェッカー vs スコア", value: `r=${fmtN(corrs.checkers_vs_score?.r, 3)}`, sub: "弱い正の相関", color: C.ACCENT },
    { label: "メイド vs スコア", value: `r=${fmtN(corrs.maids_vs_score?.r, 3)}`, sub: "ほぼ無相関", color: C.SUBTEXT },
    { label: "スタッフ vs クレーム", value: `r=${fmtN(corrs.staff_vs_claims?.r, 3)}`, sub: "中程度の正の相関", color: C.ORANGE },
  ];
  kpis.forEach((k, i) => {
    const x = 0.3 + i * 2.35;
    s2.addShape(pptx.ShapeType.roundRect, { x, y: 1.4, w: 2.15, h: 1.6, fill: { color: k.color }, rectRadius: 0.1 });
    s2.addText(k.label, { x, y: 1.5, w: 2.15, h: 0.3, fontSize: 10, color: C.WHITE, fontFace: "Arial", align: "center" });
    s2.addText(k.value, { x, y: 1.9, w: 2.15, h: 0.5, fontSize: 24, bold: true, color: C.WHITE, fontFace: "Arial", align: "center" });
    s2.addText(k.sub, { x, y: 2.5, w: 2.15, h: 0.3, fontSize: 9, color: C.WHITE, fontFace: "Arial", align: "center" });
  });
  s2.addText("• 「数」よりも「質」と「チェック体制」が品質を決定する重要な要因", { x: 0.5, y: 3.5, w: 9, h: 0.3, fontSize: 11, color: C.TEXT, fontFace: "Arial" });
  s2.addText("• メイド:チェッカー比率の最適化がスコア向上の最も効果的なレバー", { x: 0.5, y: 3.9, w: 9, h: 0.3, fontSize: 11, color: C.TEXT, fontFace: "Arial" });

  // Slide 3: Hotel comparison table
  const s3 = pptx.addSlide();
  s3.addText("ホテル別人員配置と品質", { ...titleOpts });
  const topH = hotels.filter(h => h.avg_maids).slice(0, 12);
  const tblRows = [
    [{ text: "ホテル", options: { bold: true, color: C.WHITE, fill: { color: C.NAVY }, fontSize: 8 } },
     { text: "メイド", options: { bold: true, color: C.WHITE, fill: { color: C.NAVY }, fontSize: 8 } },
     { text: "チェッカー", options: { bold: true, color: C.WHITE, fill: { color: C.NAVY }, fontSize: 8 } },
     { text: "M:C比", options: { bold: true, color: C.WHITE, fill: { color: C.NAVY }, fontSize: 8 } },
     { text: "スコア", options: { bold: true, color: C.WHITE, fill: { color: C.NAVY }, fontSize: 8 } }],
    ...topH.map((h, i) => [
      { text: h.name.replace(/ホテル/g, "H.").substring(0, 18), options: { fontSize: 7, fill: { color: i % 2 ? "FFFFFF" : "F5F7FA" } } },
      { text: fmtN(h.avg_maids), options: { fontSize: 8, fill: { color: i % 2 ? "FFFFFF" : "F5F7FA" } } },
      { text: h.avg_checkers ? fmtN(h.avg_checkers) : "-", options: { fontSize: 8, fill: { color: i % 2 ? "FFFFFF" : "F5F7FA" } } },
      { text: h.maid_checker_ratio ? `${fmtN(h.maid_checker_ratio)}:1` : "-", options: { fontSize: 8, fill: { color: i % 2 ? "FFFFFF" : "F5F7FA" } } },
      { text: fmtN(h.score, 2), options: { fontSize: 8, bold: true, fill: { color: i % 2 ? "FFFFFF" : "F5F7FA" } } },
    ])
  ];
  s3.addTable(tblRows, { x: 0.3, y: 1.2, w: 9, h: 3.5, fontSize: 8, fontFace: "Arial", border: { color: "CCCCCC", pt: 0.5 }, colW: [3.0, 1.2, 1.2, 1.2, 1.2], align: "center" });

  // Slide 4: Tier comparison
  const s4 = pptx.addSlide();
  s4.addText("配置水準別の品質比較", { ...titleOpts });
  [['low_staff', '低配置', C.ORANGE, 0.5], ['high_staff', '高配置', C.ACCENT, 5.0]].forEach(([k, label, color, x]) => {
    const t = tiers[k];
    s4.addShape(pptx.ShapeType.roundRect, { x, y: 1.4, w: 4.2, h: 2.2, fill: { color }, rectRadius: 0.1 });
    s4.addText(`${label}グループ (${t?.count || 0}ホテル)`, { x, y: 1.5, w: 4.2, h: 0.4, fontSize: 14, bold: true, color: C.WHITE, fontFace: "Arial", align: "center" });
    s4.addText(`平均メイド: ${fmtN(t?.avg_maids)}名`, { x: x + 0.3, y: 2.0, w: 3.6, h: 0.35, fontSize: 12, color: C.WHITE, fontFace: "Arial" });
    s4.addText(`平均スコア: ${fmtN(t?.avg_score, 2)}`, { x: x + 0.3, y: 2.4, w: 3.6, h: 0.35, fontSize: 16, bold: true, color: C.WHITE, fontFace: "Arial" });
    s4.addText(`クレーム率: ${fmtN(t?.avg_claim_rate)}/万室`, { x: x + 0.3, y: 2.9, w: 3.6, h: 0.3, fontSize: 11, color: C.WHITE, fontFace: "Arial" });
  });
  s4.addText("→ メイド「数」の多寡よりも、チェッカー配置と比率の最適化が品質を左右", { x: 0.5, y: 4.0, w: 9, h: 0.3, fontSize: 12, bold: true, color: C.NAVY, fontFace: "Arial" });

  // Slide 5: Recommendations
  const s5 = pptx.addSlide();
  s5.addText("提言とアクションプラン", { ...titleOpts });
  const recColors = [C.RED, C.ORANGE, C.TEAL];
  recs.forEach((rec, i) => {
    const y = 1.2 + i * 1.2;
    s5.addShape(pptx.ShapeType.roundRect, { x: 0.3, y, w: 9.2, h: 1.0, fill: { color: recColors[i] || C.SUBTEXT }, rectRadius: 0.08 });
    s5.addText(`${rec.title}  [${rec.priority}]`, { x: 0.5, y: y + 0.05, w: 8.8, h: 0.35, fontSize: 12, bold: true, color: C.WHITE, fontFace: "Arial" });
    s5.addText(rec.actions.join(" / "), { x: 0.5, y: y + 0.45, w: 8.8, h: 0.45, fontSize: 9, color: C.WHITE, fontFace: "Arial" });
  });

  const outPath = path.resolve(__dirname, "PRIMECHANGE_分析3_人員配置品質相関.pptx");
  pptx.writeFile({ fileName: outPath }).then(() => {
    console.log(`✅ PPTX: ${outPath} (${(fs.statSync(outPath).size / 1024).toFixed(1)} KB)`);
  });
}

async function main() {
  console.log("Building Analysis 3 Reports...");
  await buildDOCX();
  await buildPPTX();
  console.log("\nDone!");
}
main().catch(e => { console.error(e); process.exit(1); });
