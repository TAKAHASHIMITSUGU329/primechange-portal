#!/usr/bin/env node
"use strict";
const fs = require("fs");
const path = require("path");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType, PageNumber, PageBreak } = require("docx");
const pptxgen = require("pptxgenjs");

const data = JSON.parse(fs.readFileSync(path.resolve(__dirname, "analysis_5_data.json"), "utf-8"));
const meta = data.analysis_metadata;
const pairs = data.pairs;
const ranking = data.hotel_ranking;
const corrs = data.correlations;
const problems = data.problem_items;
const recs = data.recommendations;
const hotelDetails = data.hotel_details;

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

// Compute category averages across all hotels
const categories = ["整理整頓", "安全管理", "衛生管理", "現場運営", "現場ルール・マナー"];
function computeCategoryAvgs() {
  const result = {};
  categories.forEach(cat => {
    const vals = Object.values(hotelDetails).map(h => h.category_scores[cat]).filter(v => v != null);
    result[cat] = vals.length ? vals.reduce((a, b) => a + b, 0) / vals.length : null;
  });
  return result;
}
const catAvgs = computeCategoryAvgs();

// Compute overall safety score average
const avgSafetyScore = ranking.length ? ranking.reduce((s, h) => s + h.safety_score, 0) / ranking.length : 0;

// Count problems by category
function countProblemsByCategory() {
  const counts = {};
  problems.forEach(pi => { counts[pi.category] = (counts[pi.category] || 0) + 1; });
  return counts;
}
const problemCounts = countProblemsByCategory();

// Count problems by hotel
function countProblemsByHotel() {
  const counts = {};
  problems.forEach(pi => { counts[pi.hotel] = (counts[pi.hotel] || 0) + 1; });
  return counts;
}
const problemByHotel = countProblemsByHotel();

async function buildDOCX() {
  // Hotel ranking table
  const rkHeader = new TableRow({ children: [
    cell("順位", { bg: C.NAVY, c: C.WHITE, b: true, w: 600 }),
    cell("ホテル名", { bg: C.NAVY, c: C.WHITE, b: true, w: 2800 }),
    cell("安全スコア", { bg: C.NAVY, c: C.WHITE, b: true, w: 1000 }),
    cell("口コミ", { bg: C.NAVY, c: C.WHITE, b: true, w: 900 }),
    cell("クレーム率", { bg: C.NAVY, c: C.WHITE, b: true, w: 1000 }),
    cell("検査回数", { bg: C.NAVY, c: C.WHITE, b: true, w: 900 }),
    cell("問題件数", { bg: C.NAVY, c: C.WHITE, b: true, w: 900 }),
  ] });
  const rkRows = ranking.map((h, i) => new TableRow({ children: [
    cell(i + 1, { bg: i % 2 ? C.WHITE : C.LIGHT_BG, b: true }),
    cell(h.name, { bg: i % 2 ? C.WHITE : C.LIGHT_BG, a: AlignmentType.LEFT, sz: 14 }),
    cell(fmtN(h.safety_score, 2), { bg: i % 2 ? C.WHITE : C.LIGHT_BG, b: true, c: h.safety_score >= 3.0 ? C.GREEN : h.safety_score >= 2.8 ? C.ORANGE : C.RED }),
    cell(fmtN(h.review_score, 2), { bg: i % 2 ? C.WHITE : C.LIGHT_BG }),
    cell(fmtN(h.claim_rate, 2), { bg: i % 2 ? C.WHITE : C.LIGHT_BG }),
    cell(h.inspections, { bg: i % 2 ? C.WHITE : C.LIGHT_BG }),
    cell(problemByHotel[h.name] || 0, { bg: i % 2 ? C.WHITE : C.LIGHT_BG }),
  ] }));
  const rkTable = new Table({ rows: [rkHeader, ...rkRows], width: { size: 8100, type: WidthType.DXA } });

  // Category analysis table
  const catHeader = new TableRow({ children: [
    cell("カテゴリ", { bg: C.ACCENT, c: C.WHITE, b: true, w: 2400 }),
    cell("平均スコア(%)", { bg: C.ACCENT, c: C.WHITE, b: true, w: 1400 }),
    cell("問題件数", { bg: C.ACCENT, c: C.WHITE, b: true, w: 1200 }),
    cell("評価対象ホテル数", { bg: C.ACCENT, c: C.WHITE, b: true, w: 1600 }),
    cell("判定", { bg: C.ACCENT, c: C.WHITE, b: true, w: 1800 }),
  ] });
  const catRows = categories.map((cat, i) => {
    const avg = catAvgs[cat];
    const cnt = problemCounts[cat] || 0;
    const hotelsWithCat = Object.values(hotelDetails).filter(h => h.category_scores[cat] != null).length;
    const judge = avg == null ? "-" : avg >= 90 ? "良好" : avg >= 75 ? "要注意" : "要改善";
    const jColor = avg == null ? C.TEXT : avg >= 90 ? C.GREEN : avg >= 75 ? C.ORANGE : C.RED;
    return new TableRow({ children: [
      cell(cat, { bg: i % 2 ? C.WHITE : C.LIGHT_BG, a: AlignmentType.LEFT }),
      cell(avg != null ? fmtN(avg) : "-", { bg: i % 2 ? C.WHITE : C.LIGHT_BG, b: true }),
      cell(cnt, { bg: i % 2 ? C.WHITE : C.LIGHT_BG }),
      cell(hotelsWithCat, { bg: i % 2 ? C.WHITE : C.LIGHT_BG }),
      cell(judge, { bg: i % 2 ? C.WHITE : C.LIGHT_BG, c: jColor, b: true }),
    ] });
  });
  const catTable = new Table({ rows: [catHeader, ...catRows], width: { size: 8400, type: WidthType.DXA } });

  // Problem items table (show unique items grouped)
  const uniqueProblems = {};
  problems.forEach(pi => {
    const key = `${pi.category}|${pi.item}`;
    if (!uniqueProblems[key]) uniqueProblems[key] = { category: pi.category, item: pi.item, hotels: [], score: pi.score };
    uniqueProblems[key].hotels.push(pi.hotel);
  });
  const problemList = Object.values(uniqueProblems).sort((a, b) => b.hotels.length - a.hotels.length);
  const prHeader = new TableRow({ children: [
    cell("カテゴリ", { bg: C.NAVY, c: C.WHITE, b: true, w: 1400 }),
    cell("チェック項目", { bg: C.NAVY, c: C.WHITE, b: true, w: 4200 }),
    cell("該当ホテル数", { bg: C.NAVY, c: C.WHITE, b: true, w: 1200 }),
    cell("深刻度", { bg: C.NAVY, c: C.WHITE, b: true, w: 1000 }),
  ] });
  const prRows = problemList.map((pi, i) => new TableRow({ children: [
    cell(pi.category, { bg: i % 2 ? C.WHITE : C.LIGHT_BG, a: AlignmentType.LEFT, sz: 14 }),
    cell(pi.item, { bg: i % 2 ? C.WHITE : C.LIGHT_BG, a: AlignmentType.LEFT, sz: 14 }),
    cell(`${pi.hotels.length}/${meta.total_hotels}`, { bg: i % 2 ? C.WHITE : C.LIGHT_BG, b: true, c: pi.hotels.length >= 15 ? C.RED : pi.hotels.length >= 10 ? C.ORANGE : C.TEXT }),
    cell(pi.score <= 1 ? "高" : "中", { bg: i % 2 ? C.WHITE : C.LIGHT_BG, c: pi.score <= 1 ? C.RED : C.ORANGE, b: true }),
  ] }));
  const prTable = new Table({ rows: [prHeader, ...prRows], width: { size: 7800, type: WidthType.DXA } });

  // Correlation table
  const corrEntries = Object.entries(corrs);
  const corrLabelMap = { 'safety_vs_review': '安全スコア vs 口コミスコア', 'safety_vs_claims': '安全スコア vs クレーム率' };
  const corrHeader = new TableRow({ children: [
    cell("相関ペア", { bg: C.ACCENT, c: C.WHITE, b: true, w: 3400 }),
    cell("相関係数 (r)", { bg: C.ACCENT, c: C.WHITE, b: true, w: 1600 }),
    cell("N", { bg: C.ACCENT, c: C.WHITE, b: true, w: 800 }),
    cell("判定", { bg: C.ACCENT, c: C.WHITE, b: true, w: 2400 }),
  ] });
  const corrRows = corrEntries.map(([k, v], i) => {
    const absR = Math.abs(v.r);
    const judge = absR >= 0.4 ? '中程度の相関' : absR >= 0.2 ? '弱い相関' : 'ほぼ無相関';
    const jColor = absR >= 0.4 ? C.RED : absR >= 0.2 ? C.ORANGE : C.TEXT;
    return new TableRow({ children: [
      cell(corrLabelMap[k] || k, { bg: i % 2 ? C.WHITE : C.LIGHT_BG, a: AlignmentType.LEFT }),
      cell(fmtN(v.r, 4), { bg: i % 2 ? C.WHITE : C.LIGHT_BG, b: true }),
      cell(v.n, { bg: i % 2 ? C.WHITE : C.LIGHT_BG }),
      cell(judge, { bg: i % 2 ? C.WHITE : C.LIGHT_BG, c: jColor }),
    ] });
  });
  const corrTable = new Table({ rows: [corrHeader, ...corrRows], width: { size: 8200, type: WidthType.DXA } });

  // Top / Bottom hotels
  const top5 = ranking.slice(0, 5);
  const bottom5 = ranking.slice(-5);

  const doc = new Document({
    styles: { default: { document: { run: { font: "Arial", size: 20 } } } },
    sections: [{
      headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "PRIMECHANGE | 分析5: 安全チェック×予兆検出", size: 14, color: C.SUBTEXT, font: "Arial" })] })] }) },
      footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Page ", size: 14, color: C.SUBTEXT }), new TextRun({ children: [PageNumber.CURRENT], size: 14, color: C.SUBTEXT })] })] }) },
      children: [
        // Cover
        new Paragraph({ spacing: { before: 2400 } }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 }, children: [new TextRun({ text: "PRIMECHANGE", size: 56, font: "Arial", bold: true, color: C.NAVY })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [new TextRun({ text: "分析5", size: 44, font: "Arial", bold: true, color: C.ACCENT })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [new TextRun({ text: "安全チェック×予兆検出分析レポート", size: 36, font: "Arial", bold: true, color: C.ACCENT })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 400 }, children: [new TextRun({ text: "Safety Inspection × Early Warning Analysis", size: 22, font: "Arial", color: C.SUBTEXT, italics: true })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "株式会社PRIMECHANGE", size: 28, font: "Arial", bold: true, color: C.NAVY })] }),
        PB(),

        // 1. Executive Summary
        h1("1. エグゼクティブサマリー"),
        p(`${meta.total_hotels}ホテルの安全チェックデータと口コミ・クレーム指標を統合分析し、安全管理状況と品質低下の予兆を検出しました。`),
        h3("主要な発見"),
        bp(`全${meta.total_hotels}ホテルで安全チェック実施済み。平均安全スコアは${fmtN(avgSafetyScore, 2)}点。`, { b: true }),
        bp(`問題項目は合計${data.total_problem_count}件を検出。「整理整頓」と「現場運営」にカテゴリが集中。`, { b: true }),
        bp(`安全スコアと口コミスコアの相関はr=${fmtN(corrs.safety_vs_review?.r, 4)}（弱い正の相関）。安全管理が高いほど口コミも良好な傾向。`),
        bp(`安全スコアとクレーム率の相関はr=${fmtN(corrs.safety_vs_claims?.r, 4)}。安全管理水準とクレーム率の関連は限定的。`),
        bp(`上位3ホテル（変なホテル東京羽田、アパホテル相模原橋本駅東、川崎日航ホテル）は安全スコア3.1以上を達成。`, { b: true }),
        PB(),

        // 2. Hotel Ranking
        h1("2. ホテル別安全スコアランキング"),
        p("安全チェックの総合スコアに基づくホテルランキングです。スコアが高いほど安全管理水準が高いことを示します。"),
        rkTable,
        p(""),
        h3("上位5ホテル"),
        ...top5.map((h, i) => bp(`${i + 1}位: ${h.name}（安全スコア${fmtN(h.safety_score, 2)}, 口コミ${fmtN(h.review_score, 2)}）`, { b: true })),
        h3("下位5ホテル"),
        ...bottom5.map((h, i) => bp(`${ranking.length - 4 + i}位: ${h.name}（安全スコア${fmtN(h.safety_score, 2)}, 口コミ${fmtN(h.review_score, 2)}）`)),
        PB(),

        // 3. Category Analysis
        h1("3. カテゴリ別分析"),
        p("安全チェックの5つのカテゴリ（整理整頓、安全管理、衛生管理、現場運営、現場ルール・マナー）ごとの分析結果です。"),
        catTable,
        p(""),
        ...categories.flatMap(cat => {
          const avg = catAvgs[cat];
          const cnt = problemCounts[cat] || 0;
          if (avg == null && cnt === 0) return [];
          return [
            h3(cat),
            p(`平均スコア: ${avg != null ? fmtN(avg) + "%" : "データなし"} / 問題件数: ${cnt}件`),
            ...(cat === "整理整頓" ? [bp("全19ホテルで「棚に物が積みあがっていないか」が問題項目として検出。最も広範な課題。", { b: true })] : []),
            ...(cat === "現場運営" ? [bp("「スタッフ・社員同士の関係性」「フロントとの連絡・関係値」が大多数のホテルで課題。組織的な改善が必要。", { b: true })] : []),
            ...(cat === "衛生管理" ? [bp("「温度計・湿度計の設置」が一部ホテルで未実施。基本的な環境管理の徹底が必要。")] : []),
            ...(cat === "安全管理" ? [bp("対象ホテルの安全管理スコアは83.3〜100.0%と比較的高水準。継続維持が重要。", { c: C.GREEN })] : []),
            ...(cat === "現場ルール・マナー" ? [bp("73.3〜100.0%の範囲。一部ホテルでルール遵守に課題あり。")] : []),
          ];
        }),
        PB(),

        // 4. Problem Items
        h1("4. 問題項目一覧"),
        p(`安全チェックで検出された問題項目（△・✖・D・E評価）の一覧です。合計${data.total_problem_count}件。`),
        prTable,
        p(""),
        h3("問題項目の傾向"),
        bp("「棚に物が積みあがっていないか」: 19ホテル全てで問題検出。全社的な整理整頓の改善が必要。", { b: true, c: C.RED }),
        bp("「スタッフ・社員同士の関係性」: 多数のホテルで課題。現場コミュニケーション改善が急務。", { b: true, c: C.RED }),
        bp("「フロントとの連絡・関係値」: 現場とフロントの連携強化が品質向上の鍵。", { b: true, c: C.ORANGE }),
        bp("「温度計・湿度計の設置」: 4ホテルで未実施。即座に対応可能な改善項目。", { c: C.ORANGE }),
        PB(),

        // 5. Correlation with Quality
        h1("5. 品質指標との相関分析"),
        p("安全チェックスコアと口コミスコア・クレーム率の相関を分析し、安全管理と品質の関連を検証しました。"),
        corrTable,
        p(""),
        bp(`安全スコア vs 口コミスコア（r=${fmtN(corrs.safety_vs_review?.r, 4)}）: 弱い正の相関。安全管理水準が高いほど口コミスコアもやや高い傾向。`, { b: true }),
        bp(`安全スコア vs クレーム率（r=${fmtN(corrs.safety_vs_claims?.r, 4)}）: ほぼ無相関。クレーム率は安全チェック以外の要因にも大きく依存。`),
        p("安全チェックスコアの低下は品質低下の予兆指標として活用可能ですが、単独指標としては限界があり、他の分析との組み合わせが重要です。", { sb: 160 }),
        PB(),

        // 6. Recommendations
        h1("6. 提言"),
        ...recs.flatMap((rec, i) => [
          h2(`6.${i + 1} ${rec.title}`),
          p(`【優先度: ${rec.priority}】`, { b: true, c: rec.priority === "最優先" ? C.RED : rec.priority === "高" ? C.ORANGE : C.TEAL }),
          p(rec.rationale),
          ...rec.actions.map(a => bp(a)),
        ]),
      ]
    }]
  });

  const buf = await Packer.toBuffer(doc);
  const outPath = path.resolve(__dirname, "PRIMECHANGE_分析5_安全チェック予兆検出.docx");
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
  s1.addText("分析5: 安全チェック×予兆検出分析", { x: 0.5, y: 2.2, w: 9, h: 0.6, fontSize: 24, bold: true, color: C.WHITE, fontFace: "Arial", align: "center" });
  s1.addText(`${meta.total_hotels}ホテルの安全チェックデータに基づく予兆検出レポート`, { x: 0.5, y: 3.4, w: 9, h: 0.3, fontSize: 12, color: "CCCCCC", fontFace: "Arial", align: "center" });

  // Slide 2: KPI cards
  const s2 = pptx.addSlide();
  s2.addText("主要KPI", { ...titleOpts });
  const kpis = [
    { label: "対象ホテル数", value: `${meta.total_hotels}`, sub: "全ホテル実施済み", color: C.GREEN },
    { label: "平均安全スコア", value: fmtN(avgSafetyScore, 2), sub: "5点満点中", color: C.ACCENT },
    { label: "問題項目数", value: `${data.total_problem_count}`, sub: "要改善項目", color: C.RED },
    { label: "安全vs口コミ相関", value: `r=${fmtN(corrs.safety_vs_review?.r, 3)}`, sub: "弱い正の相関", color: C.ORANGE },
  ];
  kpis.forEach((k, i) => {
    const x = 0.3 + i * 2.35;
    s2.addShape(pptx.ShapeType.roundRect, { x, y: 1.4, w: 2.15, h: 1.6, fill: { color: k.color }, rectRadius: 0.1 });
    s2.addText(k.label, { x, y: 1.5, w: 2.15, h: 0.3, fontSize: 10, color: C.WHITE, fontFace: "Arial", align: "center" });
    s2.addText(k.value, { x, y: 1.9, w: 2.15, h: 0.5, fontSize: 24, bold: true, color: C.WHITE, fontFace: "Arial", align: "center" });
    s2.addText(k.sub, { x, y: 2.5, w: 2.15, h: 0.3, fontSize: 9, color: C.WHITE, fontFace: "Arial", align: "center" });
  });
  s2.addText("• 全ホテルで安全チェック実施。「整理整頓」「現場運営」に課題が集中", { x: 0.5, y: 3.5, w: 9, h: 0.3, fontSize: 11, color: C.TEXT, fontFace: "Arial" });
  s2.addText("• 安全スコアと口コミに弱い正の相関。安全管理の充実が品質向上に寄与する可能性", { x: 0.5, y: 3.9, w: 9, h: 0.3, fontSize: 11, color: C.TEXT, fontFace: "Arial" });

  // Slide 3: Hotel safety ranking table
  const s3 = pptx.addSlide();
  s3.addText("ホテル別安全スコアランキング", { ...titleOpts });
  const topH = ranking.slice(0, 12);
  const tblRows = [
    [{ text: "順位", options: { bold: true, color: C.WHITE, fill: { color: C.NAVY }, fontSize: 8 } },
     { text: "ホテル名", options: { bold: true, color: C.WHITE, fill: { color: C.NAVY }, fontSize: 8 } },
     { text: "安全スコア", options: { bold: true, color: C.WHITE, fill: { color: C.NAVY }, fontSize: 8 } },
     { text: "口コミ", options: { bold: true, color: C.WHITE, fill: { color: C.NAVY }, fontSize: 8 } },
     { text: "クレーム率", options: { bold: true, color: C.WHITE, fill: { color: C.NAVY }, fontSize: 8 } },
     { text: "問題件数", options: { bold: true, color: C.WHITE, fill: { color: C.NAVY }, fontSize: 8 } }],
    ...topH.map((h, i) => [
      { text: `${i + 1}`, options: { fontSize: 8, fill: { color: i % 2 ? "FFFFFF" : "F5F7FA" } } },
      { text: h.name.replace(/ホテル/g, "H.").substring(0, 18), options: { fontSize: 7, fill: { color: i % 2 ? "FFFFFF" : "F5F7FA" } } },
      { text: fmtN(h.safety_score, 2), options: { fontSize: 8, bold: true, fill: { color: i % 2 ? "FFFFFF" : "F5F7FA" } } },
      { text: fmtN(h.review_score, 2), options: { fontSize: 8, fill: { color: i % 2 ? "FFFFFF" : "F5F7FA" } } },
      { text: fmtN(h.claim_rate, 2), options: { fontSize: 8, fill: { color: i % 2 ? "FFFFFF" : "F5F7FA" } } },
      { text: `${problemByHotel[h.name] || 0}`, options: { fontSize: 8, fill: { color: i % 2 ? "FFFFFF" : "F5F7FA" } } },
    ])
  ];
  s3.addTable(tblRows, { x: 0.3, y: 1.2, w: 9.2, h: 3.5, fontSize: 8, fontFace: "Arial", border: { color: "CCCCCC", pt: 0.5 }, colW: [0.6, 3.2, 1.2, 1.0, 1.2, 1.0], align: "center" });

  // Slide 4: Problem areas
  const s4 = pptx.addSlide();
  s4.addText("問題領域の分析", { ...titleOpts });
  // Category problem boxes
  const catColors = [C.RED, C.ORANGE, C.TEAL, C.RED, C.ACCENT];
  const catDisplay = categories.filter(cat => (problemCounts[cat] || 0) > 0 || catAvgs[cat] != null);
  catDisplay.forEach((cat, i) => {
    const cnt = problemCounts[cat] || 0;
    const avg = catAvgs[cat];
    const col = i < catColors.length ? catColors[i] : C.SUBTEXT;
    if (i < 3) {
      const x = 0.3 + i * 3.1;
      s4.addShape(pptx.ShapeType.roundRect, { x, y: 1.2, w: 2.9, h: 1.4, fill: { color: col }, rectRadius: 0.1 });
      s4.addText(cat, { x, y: 1.3, w: 2.9, h: 0.3, fontSize: 11, bold: true, color: C.WHITE, fontFace: "Arial", align: "center" });
      s4.addText(`${cnt}件`, { x, y: 1.65, w: 2.9, h: 0.4, fontSize: 22, bold: true, color: C.WHITE, fontFace: "Arial", align: "center" });
      s4.addText(avg != null ? `平均 ${fmtN(avg)}%` : "データなし", { x, y: 2.15, w: 2.9, h: 0.3, fontSize: 9, color: C.WHITE, fontFace: "Arial", align: "center" });
    } else {
      const x = 0.3 + (i - 3) * 3.1;
      s4.addShape(pptx.ShapeType.roundRect, { x, y: 2.85, w: 2.9, h: 1.4, fill: { color: col }, rectRadius: 0.1 });
      s4.addText(cat, { x, y: 2.95, w: 2.9, h: 0.3, fontSize: 11, bold: true, color: C.WHITE, fontFace: "Arial", align: "center" });
      s4.addText(`${cnt}件`, { x, y: 3.3, w: 2.9, h: 0.4, fontSize: 22, bold: true, color: C.WHITE, fontFace: "Arial", align: "center" });
      s4.addText(avg != null ? `平均 ${fmtN(avg)}%` : "データなし", { x, y: 3.8, w: 2.9, h: 0.3, fontSize: 9, color: C.WHITE, fontFace: "Arial", align: "center" });
    }
  });
  s4.addText(`問題項目合計: ${data.total_problem_count}件 / 最頻出: 「棚に物が積みあがっていないか」(全${meta.total_hotels}ホテル)`, { x: 0.5, y: 4.5, w: 9, h: 0.3, fontSize: 11, bold: true, color: C.NAVY, fontFace: "Arial" });

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

  const outPath = path.resolve(__dirname, "PRIMECHANGE_分析5_安全チェック予兆検出.pptx");
  pptx.writeFile({ fileName: outPath }).then(() => {
    console.log(`PPTX: ${outPath} (${(fs.statSync(outPath).size / 1024).toFixed(1)} KB)`);
  });
}

async function main() {
  console.log("Building Analysis 5 Reports...");
  await buildDOCX();
  await buildPPTX();
  console.log("\nDone!");
}
main().catch(e => { console.error(e); process.exit(1); });
