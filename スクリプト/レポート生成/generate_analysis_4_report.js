#!/usr/bin/env node
"use strict";
const fs = require("fs");
const path = require("path");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType, PageNumber, PageBreak } = require("docx");
const pptxgen = require("pptxgenjs");

const data = JSON.parse(fs.readFileSync(path.resolve(__dirname, "analysis_4_data.json"), "utf-8"));
const meta = data.analysis_metadata;
const ta = data.time_analysis;
const corrs = ta.correlations;
const tiers = ta.time_tiers;
const bench = ta.benchmark;
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
function fmtTime(h) { if (!h) return "-"; const hrs = Math.floor(h); const mins = Math.round((h - hrs) * 60); return `${hrs}:${mins.toString().padStart(2, '0')}`; }
const PB = () => new Paragraph({ children: [new PageBreak()] });

async function buildDOCX() {
  const hotelsWithData = hotels.filter(h => h.avg_completion_time);
  const hHeader = new TableRow({ children: [
    cell("ホテル名", { bg: C.NAVY, c: C.WHITE, b: true, w: 3000 }),
    cell("平均完了時間", { bg: C.NAVY, c: C.WHITE, b: true, w: 1400 }),
    cell("最早", { bg: C.NAVY, c: C.WHITE, b: true, w: 900 }),
    cell("最遅", { bg: C.NAVY, c: C.WHITE, b: true, w: 900 }),
    cell("スコア", { bg: C.NAVY, c: C.WHITE, b: true, w: 1000 }),
    cell("クレーム率", { bg: C.NAVY, c: C.WHITE, b: true, w: 1200 }),
    cell("データ日数", { bg: C.NAVY, c: C.WHITE, b: true, w: 900 }),
  ] });
  hotelsWithData.sort((a, b) => a.avg_completion_time - b.avg_completion_time);
  const hRows = hotelsWithData.map((h, i) => {
    const timeColor = h.avg_completion_time <= 15 ? C.GREEN : h.avg_completion_time <= 15.5 ? C.ORANGE : C.RED;
    return new TableRow({ children: [
      cell(h.name, { bg: i % 2 ? C.WHITE : C.LIGHT_BG, a: AlignmentType.LEFT, sz: 14 }),
      cell(fmtTime(h.avg_completion_time), { bg: i % 2 ? C.WHITE : C.LIGHT_BG, b: true, c: timeColor }),
      cell(fmtTime(h.avg_completion_time - (h.time_std || 0)), { bg: i % 2 ? C.WHITE : C.LIGHT_BG, sz: 14 }),
      cell(fmtTime(h.avg_completion_time + (h.time_std || 0)), { bg: i % 2 ? C.WHITE : C.LIGHT_BG, sz: 14 }),
      cell(fmtN(h.score, 2), { bg: i % 2 ? C.WHITE : C.LIGHT_BG, b: true }),
      cell(fmtN(h.claim_rate), { bg: i % 2 ? C.WHITE : C.LIGHT_BG }),
      cell(h.time_data_points, { bg: i % 2 ? C.WHITE : C.LIGHT_BG }),
    ] });
  });
  const hTable = new Table({ rows: [hHeader, ...hRows], width: { size: 9300, type: WidthType.DXA } });

  const corrEntries = Object.entries(corrs);
  const labelMap = { 'time_vs_score': '完了時間 vs スコア', 'time_vs_claims': '完了時間 vs クレーム率', 'time_variability_vs_score': '時間ばらつき vs スコア' };
  const cHeader = new TableRow({ children: [
    cell("相関ペア", { bg: C.ACCENT, c: C.WHITE, b: true, w: 3400 }),
    cell("r", { bg: C.ACCENT, c: C.WHITE, b: true, w: 1400 }),
    cell("R²", { bg: C.ACCENT, c: C.WHITE, b: true, w: 1400 }),
    cell("判定", { bg: C.ACCENT, c: C.WHITE, b: true, w: 2800 }),
  ] });
  const cRows = corrEntries.map(([k, v], i) => {
    const absR = Math.abs(v.r);
    const judge = absR >= 0.4 ? '中程度の相関' : absR >= 0.2 ? '弱い相関' : 'ほぼ無相関';
    return new TableRow({ children: [
      cell(labelMap[k] || k, { bg: i % 2 ? C.WHITE : C.LIGHT_BG, a: AlignmentType.LEFT }),
      cell(fmtN(v.r, 4), { bg: i % 2 ? C.WHITE : C.LIGHT_BG, b: true }),
      cell(v.r_squared != null ? fmtN(v.r_squared, 4) : "-", { bg: i % 2 ? C.WHITE : C.LIGHT_BG }),
      cell(judge, { bg: i % 2 ? C.WHITE : C.LIGHT_BG }),
    ] });
  });
  const cTable = new Table({ rows: [cHeader, ...cRows], width: { size: 9000, type: WidthType.DXA } });

  const doc = new Document({
    styles: { default: { document: { run: { font: "Arial", size: 20 } } } },
    sections: [{
      headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "PRIMECHANGE | 分析4: 清掃完了時間×品質", size: 14, color: C.SUBTEXT, font: "Arial" })] })] }) },
      footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Page ", size: 14, color: C.SUBTEXT }), new TextRun({ children: [PageNumber.CURRENT], size: 14, color: C.SUBTEXT })] })] }) },
      children: [
        new Paragraph({ spacing: { before: 2400 } }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 }, children: [new TextRun({ text: "PRIMECHANGE", size: 56, font: "Arial", bold: true, color: C.NAVY })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [new TextRun({ text: "分析4", size: 44, font: "Arial", bold: true, color: C.ACCENT })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [new TextRun({ text: "清掃完了時間×品質 分析レポート", size: 36, font: "Arial", bold: true, color: C.ACCENT })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 400 }, children: [new TextRun({ text: "Cleaning Completion Time × Quality Analysis", size: 22, font: "Arial", color: C.SUBTEXT, italics: true })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "株式会社PRIMECHANGE", size: 28, font: "Arial", bold: true, color: C.NAVY })] }),
        PB(),

        h1("1. エグゼクティブサマリー"),
        p(`18ホテルの日報データから清掃完了時間と口コミスコア・クレーム率の関係を分析しました。全ホテル平均完了時間は${fmtTime(ta.overall_avg_time)}、範囲は${ta.overall_time_range}です。`),
        h3("主要な発見"),
        bp(`完了時間とスコアに弱い正の相関（r=${fmtN(corrs.time_vs_score?.r, 4)}）。やや意外だが、丁寧な清掃に時間をかけるホテルのスコアが若干高い傾向。`, { b: true }),
        bp(`完了時間とクレーム率に弱い正の相関（r=${fmtN(corrs.time_vs_claims?.r, 4)}）。完了が遅いホテルでクレームがやや多い傾向。`, { b: true }),
        bp(`最速5ホテル平均スコア${fmtN(bench.fastest_avg_score, 2)} vs 最遅5ホテル${fmtN(bench.slowest_avg_score, 2)}（差${fmtN(bench.score_difference, 2)}点）。`),
        bp(`時間のばらつき（標準偏差）とスコアの相関はほぼゼロ（r=${fmtN(corrs.time_variability_vs_score?.r, 4)}）。安定性よりも基本完了時間が重要。`),
        PB(),

        h1("2. 相関分析結果"),
        cTable,
        p(""),
        p(`${corrs.time_vs_score?.interpretation || ""}。ただし相関は弱く、完了時間だけでスコアを予測することは困難です。人員配置（分析3）やチェック体制との組み合わせで品質が決定されると考えられます。`),
        PB(),

        h1("3. ホテル別完了時間一覧"),
        hTable,
        PB(),

        h1("4. 完了時間帯別の品質比較"),
        bp(`早期完了グループ（${tiers.early_finish?.count}ホテル）: 平均${fmtTime(tiers.early_finish?.avg_time)} → スコア${fmtN(tiers.early_finish?.avg_score, 2)}`, { b: true }),
        bp(`標準完了グループ（${tiers.mid_finish?.count}ホテル）: 平均${fmtTime(tiers.mid_finish?.avg_time)} → スコア${fmtN(tiers.mid_finish?.avg_score, 2)}`, { b: true }),
        bp(`遅延完了グループ（${tiers.late_finish?.count}ホテル）: 平均${fmtTime(tiers.late_finish?.avg_time)} → スコア${fmtN(tiers.late_finish?.avg_score, 2)}`, { b: true }),

        h2("4.1 最速 vs 最遅ベンチマーク"),
        bp("最速5ホテル:", { b: true }),
        ...bench.fastest_5.map(h => bp(`  ${h.name}: ${fmtTime(h.time)}, スコア${fmtN(h.score, 2)}`, { level: 1 })),
        bp("最遅5ホテル:", { b: true }),
        ...bench.slowest_5.map(h => bp(`  ${h.name}: ${fmtTime(h.time)}, スコア${fmtN(h.score, 2)}`, { level: 1 })),
        PB(),

        h1("5. 提言"),
        ...recs.flatMap((rec, i) => [
          h2(`5.${i + 1} ${rec.title}`),
          p(`【優先度: ${rec.priority}】`, { b: true, c: rec.priority === "高" ? C.ORANGE : C.TEAL }),
          p(rec.rationale),
          ...rec.actions.map(a => bp(a)),
        ]),
      ]
    }]
  });

  const buf = await Packer.toBuffer(doc);
  const outPath = path.resolve(__dirname, "PRIMECHANGE_分析4_清掃完了時間品質.docx");
  fs.writeFileSync(outPath, buf);
  console.log(`✅ DOCX: ${outPath} (${(buf.length / 1024).toFixed(1)} KB)`);
}

function buildPPTX() {
  const pptx = new pptxgen();
  pptx.layout = "LAYOUT_16x9";
  const bgDark = { fill: C.NAVY };
  const titleOpts = { x: 0.5, y: 0.3, w: 9, h: 0.6, fontSize: 22, bold: true, color: C.NAVY, fontFace: "Arial" };

  const s1 = pptx.addSlide(); s1.background = bgDark;
  s1.addText("PRIMECHANGE", { x: 0.5, y: 1.2, w: 9, h: 0.8, fontSize: 40, bold: true, color: C.WHITE, fontFace: "Arial", align: "center" });
  s1.addText("分析4: 清掃完了時間×品質 分析", { x: 0.5, y: 2.2, w: 9, h: 0.6, fontSize: 24, bold: true, color: C.WHITE, fontFace: "Arial", align: "center" });
  s1.addText(`平均完了時間: ${fmtTime(ta.overall_avg_time)} | 対象: ${meta.hotels_with_time_data}ホテル`, { x: 0.5, y: 3.4, w: 9, h: 0.3, fontSize: 12, color: "CCCCCC", fontFace: "Arial", align: "center" });

  // KPIs
  const s2 = pptx.addSlide();
  s2.addText("主要な発見", { ...titleOpts });
  const kpis = [
    { label: "平均完了時間", value: fmtTime(ta.overall_avg_time), sub: `範囲: ${ta.overall_time_range}`, color: C.ACCENT },
    { label: "時間 vs スコア", value: `r=${fmtN(corrs.time_vs_score?.r, 3)}`, sub: "弱い正の相関", color: C.ORANGE },
    { label: "時間 vs クレーム", value: `r=${fmtN(corrs.time_vs_claims?.r, 3)}`, sub: "弱い正の相関", color: C.RED },
    { label: "最速vs最遅", value: `${fmtN(bench.score_difference, 2)}点差`, sub: `${fmtN(bench.fastest_avg_score, 2)} vs ${fmtN(bench.slowest_avg_score, 2)}`, color: C.TEAL },
  ];
  kpis.forEach((k, i) => {
    const x = 0.3 + i * 2.35;
    s2.addShape(pptx.ShapeType.roundRect, { x, y: 1.4, w: 2.15, h: 1.6, fill: { color: k.color }, rectRadius: 0.1 });
    s2.addText(k.label, { x, y: 1.5, w: 2.15, h: 0.3, fontSize: 10, color: C.WHITE, fontFace: "Arial", align: "center" });
    s2.addText(k.value, { x, y: 1.9, w: 2.15, h: 0.5, fontSize: 24, bold: true, color: C.WHITE, fontFace: "Arial", align: "center" });
    s2.addText(k.sub, { x, y: 2.5, w: 2.15, h: 0.3, fontSize: 9, color: C.WHITE, fontFace: "Arial", align: "center" });
  });

  // Tier comparison
  const s3 = pptx.addSlide();
  s3.addText("完了時間帯別の品質比較", { ...titleOpts });
  const tierColors = [C.GREEN, C.ORANGE, C.RED];
  ['early_finish', 'mid_finish', 'late_finish'].forEach((k, i) => {
    const t = tiers[k];
    const x = 0.3 + i * 3.15;
    s3.addShape(pptx.ShapeType.roundRect, { x, y: 1.4, w: 2.9, h: 2.5, fill: { color: tierColors[i] }, rectRadius: 0.1 });
    s3.addText(t?.label || k, { x, y: 1.5, w: 2.9, h: 0.35, fontSize: 11, bold: true, color: C.WHITE, fontFace: "Arial", align: "center" });
    s3.addText(fmtTime(t?.avg_time), { x, y: 1.9, w: 2.9, h: 0.5, fontSize: 28, bold: true, color: C.WHITE, fontFace: "Arial", align: "center" });
    s3.addText(`スコア: ${fmtN(t?.avg_score, 2)}`, { x, y: 2.5, w: 2.9, h: 0.35, fontSize: 13, color: C.WHITE, fontFace: "Arial", align: "center" });
    s3.addText(`クレーム率: ${fmtN(t?.avg_claim_rate)}/万室`, { x, y: 2.9, w: 2.9, h: 0.3, fontSize: 10, color: C.WHITE, fontFace: "Arial", align: "center" });
    s3.addText(`${t?.count || 0}ホテル`, { x, y: 3.3, w: 2.9, h: 0.3, fontSize: 10, color: C.WHITE, fontFace: "Arial", align: "center" });
  });

  // Recommendations
  const s4 = pptx.addSlide();
  s4.addText("提言とアクションプラン", { ...titleOpts });
  recs.forEach((rec, i) => {
    const y = 1.2 + i * 1.2;
    const colors = [C.ORANGE, C.TEAL, C.ACCENT];
    s4.addShape(pptx.ShapeType.roundRect, { x: 0.3, y, w: 9.2, h: 1.0, fill: { color: colors[i] || C.SUBTEXT }, rectRadius: 0.08 });
    s4.addText(`${rec.title}  [${rec.priority}]`, { x: 0.5, y: y + 0.05, w: 8.8, h: 0.35, fontSize: 12, bold: true, color: C.WHITE, fontFace: "Arial" });
    s4.addText(rec.actions.join(" / "), { x: 0.5, y: y + 0.45, w: 8.8, h: 0.45, fontSize: 9, color: C.WHITE, fontFace: "Arial" });
  });

  const outPath = path.resolve(__dirname, "PRIMECHANGE_分析4_清掃完了時間品質.pptx");
  pptx.writeFile({ fileName: outPath }).then(() => {
    console.log(`✅ PPTX: ${outPath} (${(fs.statSync(outPath).size / 1024).toFixed(1)} KB)`);
  });
}

async function main() {
  console.log("Building Analysis 4 Reports...");
  await buildDOCX();
  await buildPPTX();
  console.log("\nDone!");
}
main().catch(e => { console.error(e); process.exit(1); });
