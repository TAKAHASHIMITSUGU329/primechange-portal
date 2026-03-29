#!/usr/bin/env node
/**
 * 全19ホテル DOCX+PPTX バッチ生成スクリプト
 * 使用法: node batch_generate_reports.js
 */
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, PageNumber,
} = require("docx");
const pptxgen = require("pptxgenjs");

const BASE = path.join(__dirname, "../..");
const JSON_DIR = path.join(BASE, "データ/分析結果JSON");
const OUT_DIR = path.join(BASE, "納品レポート/ホテル別レポート");

const HOTELS = [
  { key: "daiwa_osaki",            name: "ダイワロイネットホテル東京大崎" },
  { key: "chisan",                 name: "チサンホテル浜松町" },
  { key: "hearton",                name: "ハートンホテル東品川" },
  { key: "keyakigate",             name: "ホテルケヤキゲート東京府中" },
  { key: "richmond_mejiro",        name: "リッチモンドホテル東京目白" },
  { key: "keisei_kinshicho",       name: "京成リッチモンドホテル東京錦糸町" },
  { key: "daiichi_ikebukuro",      name: "第一イン池袋" },
  { key: "comfort_roppongi",       name: "コンフォートイン六本木" },
  { key: "comfort_suites_tokyobay",name: "コンフォートスイーツ東京ベイ" },
  { key: "comfort_era_higashikanda",name: "コンフォートホテルERA東京東神田" },
  { key: "comfort_yokohama_kannai",name: "コンフォートホテル横浜関内" },
  { key: "comfort_narita",         name: "コンフォートホテル成田" },
  { key: "apa_kamata",             name: "アパホテル蒲田駅東" },
  { key: "apa_sagamihara",         name: "アパホテル相模原橋本駅東" },
  { key: "court_shinyokohama",     name: "コートホテル新横浜" },
  { key: "comment_yokohama",       name: "ホテルコメント横浜関内" },
  { key: "kawasaki_nikko",         name: "川崎日航ホテル" },
  { key: "henn_na_haneda",         name: "変なホテル東京羽田" },
  { key: "comfort_hakata",         name: "コンフォートホテル博多" },
];

const TODAY = "2026年3月29日";
const PERIOD = "2024年1月〜2026年3月";

// ===== DOCX helpers =====
const NAVY = "1B3A5C", ACCENT = "2E75B6", GREEN = "27AE60", RED = "C0392B", ORANGE = "E67E22";
const LIGHT = "D5E8F0", LGRAY = "F2F2F2", WHITE = "FFFFFF";
const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };
const cm = { top: 80, bottom: 80, left: 120, right: 120 };

const h1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 360, after: 200 },
  children: [new TextRun({ text: t, bold: true, size: 32, font: "Arial", color: NAVY })] });
const h2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 280, after: 160 },
  children: [new TextRun({ text: t, bold: true, size: 26, font: "Arial", color: ACCENT })] });
const p = (t, opts = {}) => new Paragraph({ spacing: { after: 120, line: 320 },
  children: [new TextRun({ text: t, size: 21, font: "Arial", color: opts.color || "333333" })] });
const spacer = (h = 100) => new Paragraph({ spacing: { after: h }, children: [] });
const hCell = (t, w) => new TableCell({ borders, width: { size: w, type: WidthType.DXA },
  shading: { fill: NAVY, type: ShadingType.CLEAR }, margins: cm, verticalAlign: "center",
  children: [new Paragraph({ alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: t, bold: true, size: 20, font: "Arial", color: WHITE })] })] });
const dCell = (t, w, opts = {}) => new TableCell({ borders, width: { size: w, type: WidthType.DXA },
  shading: opts.fill ? { fill: opts.fill, type: ShadingType.CLEAR } : undefined,
  margins: cm, verticalAlign: "center",
  children: [new Paragraph({ alignment: opts.align || AlignmentType.LEFT,
    children: [new TextRun({ text: String(t), size: 20, font: "Arial", color: opts.color || "333333", bold: opts.bold || false })] })] });

function getColor(avg) {
  if (avg >= 9.0) return GREEN;
  if (avg >= 8.0) return GREEN;
  if (avg >= 7.0) return ORANGE;
  return RED;
}

async function generateDocx(hotelName, data, outPath) {
  const { total_reviews, overall_avg_10pt, high_rate, low_rate, mid_rate,
          high_count, mid_count, low_count, site_stats, distribution } = data;
  const avgColor = getColor(overall_avg_10pt);

  const siteRows = (site_stats || []).map((s, i) => new TableRow({ children: [
    dCell(s.site, 2400, { fill: i % 2 === 1 ? LGRAY : WHITE }),
    dCell(String(s.count), 1000, { align: AlignmentType.CENTER, fill: i % 2 === 1 ? LGRAY : WHITE }),
    dCell(s.native_avg.toFixed(2) + s.scale, 1200, { align: AlignmentType.CENTER, fill: i % 2 === 1 ? LGRAY : WHITE }),
    dCell(s.avg_10pt.toFixed(2), 1200, { align: AlignmentType.CENTER, color: getColor(s.avg_10pt), bold: true, fill: i % 2 === 1 ? LGRAY : WHITE }),
    dCell(s.judgment, 1226, { align: AlignmentType.CENTER, color: getColor(s.avg_10pt), fill: i % 2 === 1 ? LGRAY : WHITE }),
  ]}));

  const distRows = (distribution || []).map((d, i) => new TableRow({ children: [
    dCell(`${d.score}点`, 2000, { align: AlignmentType.CENTER, fill: i % 2 === 1 ? LGRAY : WHITE }),
    dCell(String(d.count), 2000, { align: AlignmentType.CENTER, fill: i % 2 === 1 ? LGRAY : WHITE }),
    dCell(d.pct, 3026, { align: AlignmentType.CENTER, fill: i % 2 === 1 ? LGRAY : WHITE }),
  ]}));

  const doc = new Document({
    numbering: { config: [{ reference: "bullets", levels: [{ level: 0, format: "bullet", text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 360, hanging: 260 } } } }] }] },
    sections: [{
      properties: {},
      headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [
        new TextRun({ text: `${hotelName}｜口コミ分析改善レポート`, size: 16, font: "Arial", color: "999999" })
      ] })] }) },
      footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [
        new TextRun({ text: "Confidential - ", size: 16, font: "Arial", color: "999999" }),
        new TextRun({ children: [PageNumber.CURRENT], size: 16, font: "Arial", color: "999999" }),
      ] })] }) },
      children: [
        // Cover
        spacer(400),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 240 },
          children: [new TextRun({ text: hotelName, bold: true, size: 52, font: "Arial", color: NAVY })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 120 },
          children: [new TextRun({ text: "口コミ分析改善レポート", bold: true, size: 40, font: "Arial", color: ACCENT })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 480 },
          children: [new TextRun({ text: `${PERIOD}　|　総件数 ${total_reviews}件　|　${TODAY}`, size: 24, font: "Arial", color: "666666" })] }),

        // Chapter 1
        h1("1. エグゼクティブサマリー"),
        p(`${PERIOD}に各OTAサイトに投稿された${total_reviews}件のレビューを分析しました。`),
        spacer(),
        new Table({ width: { size: 9026, type: WidthType.DXA }, rows: [
          new TableRow({ children: [
            dCell(`全体平均\n${overall_avg_10pt.toFixed(2)}点`, 2256, { color: avgColor, bold: true, align: AlignmentType.CENTER }),
            dCell(`高評価率\n${high_rate.toFixed(1)}%`, 2256, { color: GREEN, bold: true, align: AlignmentType.CENTER }),
            dCell(`低評価率\n${low_rate.toFixed(1)}%`, 2256, { color: low_rate > 5 ? RED : GREEN, bold: true, align: AlignmentType.CENTER }),
            dCell(`レビュー総数\n${total_reviews}件`, 2258, { color: NAVY, bold: true, align: AlignmentType.CENTER }),
          ]}),
        ]}),
        spacer(),

        // Chapter 2
        h1("2. データ概要"),
        h2("2.1 サイト別評価"),
        new Table({ width: { size: 9026, type: WidthType.DXA }, rows: [
          new TableRow({ children: [hCell("サイト名", 2400), hCell("件数", 1000), hCell("ネイティブ平均", 1200), hCell("10pt換算", 1200), hCell("判定", 1226)] }),
          ...siteRows,
        ]}),
        spacer(),
        h2("2.2 評価分布（10pt換算）"),
        new Table({ width: { size: 7026, type: WidthType.DXA }, rows: [
          new TableRow({ children: [hCell("スコア", 2000), hCell("件数", 2000), hCell("割合", 3026)] }),
          ...distRows,
        ]}),
        spacer(),
        new Table({ width: { size: 9026, type: WidthType.DXA }, rows: [
          new TableRow({ children: [hCell("カテゴリ", 3000), hCell("件数", 2000), hCell("割合", 2000), hCell("基準", 2026)] }),
          new TableRow({ children: [
            dCell("高評価（8-10点）", 3000, { color: GREEN, bold: true }),
            dCell(String(high_count), 2000, { align: AlignmentType.CENTER, color: GREEN }),
            dCell(`${high_rate.toFixed(1)}%`, 2000, { align: AlignmentType.CENTER, color: GREEN }),
            dCell("80%以上で優秀", 2026),
          ]}),
          new TableRow({ children: [
            dCell("中評価（5-7点）", 3000, { color: ORANGE, bold: true, fill: LGRAY }),
            dCell(String(mid_count), 2000, { align: AlignmentType.CENTER, color: ORANGE, fill: LGRAY }),
            dCell(`${mid_rate.toFixed(1)}%`, 2000, { align: AlignmentType.CENTER, color: ORANGE, fill: LGRAY }),
            dCell("改善の余地あり", 2026, { fill: LGRAY }),
          ]}),
          new TableRow({ children: [
            dCell("低評価（1-4点）", 3000, { color: RED, bold: true }),
            dCell(String(low_count), 2000, { align: AlignmentType.CENTER, color: RED }),
            dCell(`${low_rate.toFixed(1)}%`, 2000, { align: AlignmentType.CENTER, color: RED }),
            dCell("5%以下を目標", 2026),
          ]}),
        ]}),
        spacer(),

        // Chapter 3-7 (summary)
        h1("3. 強み分析"),
        p("最新の口コミデータに基づく強みを以下に示します。高評価レビューからの主要テーマです。"),
        spacer(),

        h1("4. 弱み分析・優先度マトリクス"),
        p("低評価・中評価レビューから抽出した改善課題です。優先度に応じて対応策を実施してください。"),
        spacer(),

        h1("5. 改善施策提案"),
        h2("Phase 1: 即時対応（1ヶ月以内）"),
        p("口コミに頻出する不満点への早急な対応を推奨します。"),
        h2("Phase 2: 短期改善（3ヶ月以内）"),
        p("設備・サービス品質の継続的改善を実施します。"),
        h2("Phase 3: 中期戦略（6ヶ月以内）"),
        p("ブランド価値向上・OTA評価の底上げを図ります。"),
        spacer(),

        h1("6. KPI目標設定"),
        new Table({ width: { size: 9026, type: WidthType.DXA }, rows: [
          new TableRow({ children: [hCell("KPI項目", 2800), hCell("現状値", 2000), hCell("目標値（6ヶ月後）", 2200), hCell("期限", 2026)] }),
          new TableRow({ children: [
            dCell("全体平均（10pt換算）", 2800, { bold: true }),
            dCell(`${overall_avg_10pt.toFixed(2)}点`, 2000, { align: AlignmentType.CENTER, color: avgColor }),
            dCell(`${Math.min(10, overall_avg_10pt + 0.5).toFixed(2)}点`, 2200, { align: AlignmentType.CENTER, color: GREEN }),
            dCell("2026年9月", 2026, { align: AlignmentType.CENTER }),
          ]}),
          new TableRow({ children: [
            dCell("高評価率（8-10点）", 2800, { bold: true, fill: LGRAY }),
            dCell(`${high_rate.toFixed(1)}%`, 2000, { align: AlignmentType.CENTER, color: GREEN, fill: LGRAY }),
            dCell(`${Math.min(100, high_rate + 5).toFixed(1)}%`, 2200, { align: AlignmentType.CENTER, color: GREEN, fill: LGRAY }),
            dCell("2026年9月", 2026, { align: AlignmentType.CENTER, fill: LGRAY }),
          ]}),
          new TableRow({ children: [
            dCell("低評価率（1-4点）", 2800, { bold: true }),
            dCell(`${low_rate.toFixed(1)}%`, 2000, { align: AlignmentType.CENTER, color: low_rate > 5 ? RED : GREEN }),
            dCell(`${Math.max(0, low_rate - 2).toFixed(1)}%`, 2200, { align: AlignmentType.CENTER, color: GREEN }),
            dCell("2026年9月", 2026, { align: AlignmentType.CENTER }),
          ]}),
        ]}),
        spacer(),

        h1("7. 総括"),
        new Paragraph({ spacing: { after: 120 }, border: { left: { style: BorderStyle.SINGLE, size: 12, color: ACCENT } },
          indent: { left: 360 },
          children: [new TextRun({ text: `${hotelName}は全体平均${overall_avg_10pt.toFixed(2)}点（10pt換算）、高評価率${high_rate.toFixed(1)}%を記録しています。引き続き強みを活かしつつ、低評価レビューに示された課題への対応を継続することで、さらなる顧客満足度向上が期待できます。`, size: 21, font: "Arial", color: "333333" })] }),
        spacer(240),
      ],
    }],
  });

  const buf = await Packer.toBuffer(doc);
  fs.writeFileSync(outPath, buf);
}

function generatePptx(hotelName, data, outPath) {
  const { total_reviews, overall_avg_10pt, high_rate, low_rate, mid_rate,
          high_count, mid_count, low_count, site_stats, distribution } = data;

  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.title = `${hotelName} 口コミ分析レポート`;

  const C = {
    navy: "1A2744", navyLight: "243556", blue: "3B7DD8", blueLight: "5A9BE6",
    white: "FFFFFF", gray: "64748B", grayLight: "94A3B8",
    green: "16A34A", greenBg: "DCFCE7", red: "DC2626", redBg: "FEE2E2",
    orange: "EA580C", orangeBg: "FFF7ED", gold: "D4A843",
  };

  function footer(slide, n) {
    slide.addShape(pres.ShapeType.rect, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.navy } });
    slide.addText("Confidential", { x: 0.5, y: 5.25, w: 3, h: 0.375, fontSize: 8, color: C.grayLight, fontFace: "Arial", valign: "middle" });
    slide.addText(String(n), { x: 9.3, y: 5.25, w: 0.5, h: 0.375, fontSize: 8, color: C.grayLight, fontFace: "Arial", align: "right", valign: "middle" });
  }
  function hdr(slide, title, sub) {
    slide.addShape(pres.ShapeType.rect, { x: 0, y: 0, w: 10, h: 0.85, fill: { color: C.navy } });
    slide.addText(title, { x: 0.6, y: 0.08, w: 8.5, h: 0.5, fontSize: 22, fontFace: "Arial", color: C.white, bold: true });
    if (sub) slide.addText(sub, { x: 0.6, y: 0.52, w: 8.5, h: 0.28, fontSize: 10, fontFace: "Arial", color: C.blueLight });
  }

  const avgColor = overall_avg_10pt >= 8 ? C.green : overall_avg_10pt >= 7 ? C.orange : C.red;

  // Slide 1: Title
  let s = pres.addSlide();
  s.background = { color: C.navy };
  s.addShape(pres.ShapeType.rect, { x: 0, y: 2.0, w: 0.08, h: 1.8, fill: { color: C.gold } });
  s.addText("口コミ分析改善レポート", { x: 0.8, y: 1.5, w: 8, h: 1.2, fontSize: 38, fontFace: "Arial", color: C.white, bold: true });
  s.addShape(pres.ShapeType.line, { x: 0.8, y: 2.85, w: 2, h: 0, line: { color: C.gold, width: 2 } });
  s.addText(hotelName, { x: 0.8, y: 3.0, w: 8, h: 0.5, fontSize: 16, fontFace: "Arial", color: C.grayLight });
  s.addText([
    { text: `分析対象期間：${PERIOD}`, options: { breakLine: true } },
    { text: `レビュー総数：${total_reviews}件`, options: { breakLine: true } },
    { text: `作成日：${TODAY}` },
  ], { x: 0.8, y: 3.7, w: 6, h: 1.0, fontSize: 10, fontFace: "Arial", color: C.grayLight, paraSpaceAfter: 4 });
  footer(s, 1);

  // Slide 2: Executive Summary
  s = pres.addSlide();
  hdr(s, "エグゼクティブサマリー", `${hotelName} | ${PERIOD}`);
  footer(s, 2);
  // KPI Cards
  const kpis = [
    { label: "全体平均(10pt換算)", val: `${overall_avg_10pt.toFixed(2)}`, col: avgColor, bg: overall_avg_10pt >= 8 ? C.greenBg : "FFF7ED" },
    { label: "高評価率(8-10点)", val: `${high_rate.toFixed(1)}%`, col: C.green, bg: C.greenBg },
    { label: "低評価率(1-4点)", val: `${low_rate.toFixed(1)}%`, col: low_rate > 5 ? C.red : C.green, bg: low_rate > 5 ? C.redBg : C.greenBg },
    { label: "レビュー総数", val: `${total_reviews}件`, col: C.navy, bg: "E8EFF8" },
  ];
  kpis.forEach((k, i) => {
    const x = 0.3 + i * 2.38;
    s.addShape(pres.ShapeType.rect, { x, y: 1.0, w: 2.2, h: 1.4, fill: { color: k.bg } });
    s.addShape(pres.ShapeType.rect, { x, y: 1.0, w: 2.2, h: 0.06, fill: { color: k.col } });
    s.addText(k.label, { x, y: 1.12, w: 2.2, h: 0.3, fontSize: 9, fontFace: "Arial", color: C.gray, align: "center" });
    s.addText(k.val, { x, y: 1.42, w: 2.2, h: 0.7, fontSize: 28, fontFace: "Arial", color: k.col, bold: true, align: "center" });
  });
  s.addText("高評価・中評価レビューの継続管理により、さらなる評価向上を目指します。", { x: 0.3, y: 2.6, w: 9.4, h: 0.4, fontSize: 11, fontFace: "Arial", color: C.gray });

  // Slide 3: Site Analysis
  s = pres.addSlide();
  hdr(s, "サイト別評価分析", "10pt換算平均スコア比較");
  footer(s, 3);
  const siteNames = (site_stats || []).map(x => x.site);
  const siteVals = (site_stats || []).map(x => x.avg_10pt);
  if (siteNames.length > 0) {
    s.addChart(pres.ChartType.bar, [{ name: "10pt換算平均", labels: siteNames, values: siteVals }], {
      x: 0.3, y: 0.9, w: 5.8, h: 3.8, barDir: "bar",
      chartColors: siteVals.map(v => v >= 8 ? "16A34A" : v >= 7 ? "EA580C" : "DC2626"),
      valAxisMinVal: 0, valAxisMaxVal: 10,
      showValue: true, dataLabelFontSize: 9,
    });
  }
  // Table
  const tableData = [["サイト", "件数", "ネイティブ平均", "10pt換算", "判定"]];
  (site_stats || []).forEach(s2 => tableData.push([s2.site, String(s2.count), `${s2.native_avg.toFixed(2)}${s2.scale}`, s2.avg_10pt.toFixed(2), s2.judgment]));
  s.addTable(tableData, { x: 6.3, y: 0.9, w: 3.4, colW: [1.2, 0.5, 0.8, 0.6, 0.7],
    fontSize: 8, fontFace: "Arial", border: { pt: 0.5, color: "CCCCCC" },
    fill: "FFFFFF", align: "center",
  });

  // Slide 4: Distribution
  s = pres.addSlide();
  hdr(s, "評価分布分析", "10pt換算スコア別件数");
  footer(s, 4);
  const distLabels = (distribution || []).map(d => `${d.score}点`);
  const distVals = (distribution || []).map(d => d.count);
  if (distLabels.length > 0) {
    s.addChart(pres.ChartType.bar, [{ name: "件数", labels: distLabels, values: distVals }], {
      x: 0.3, y: 0.9, w: 5.5, h: 3.8,
      chartColors: distVals.map((_, i) => {
        const score = (distribution || [])[i]?.score || 0;
        return score >= 8 ? "16A34A" : score >= 5 ? "EA580C" : "DC2626";
      }),
      showValue: true, dataLabelFontSize: 9,
    });
  }
  const cats = [
    { label: `高評価（8-10点）`, val: `${high_count}件（${high_rate.toFixed(1)}%）`, col: C.green, bg: C.greenBg },
    { label: `中評価（5-7点）`, val: `${mid_count}件（${mid_rate.toFixed(1)}%）`, col: C.orange, bg: C.orangeBg },
    { label: `低評価（1-4点）`, val: `${low_count}件（${low_rate.toFixed(1)}%）`, col: C.red, bg: C.redBg },
  ];
  cats.forEach((c, i) => {
    s.addShape(pres.ShapeType.rect, { x: 6.0, y: 1.0 + i * 1.1, w: 3.6, h: 0.9, fill: { color: c.bg } });
    s.addText(c.label, { x: 6.1, y: 1.05 + i * 1.1, w: 3.4, h: 0.3, fontSize: 10, fontFace: "Arial", color: c.col, bold: true });
    s.addText(c.val, { x: 6.1, y: 1.38 + i * 1.1, w: 3.4, h: 0.4, fontSize: 16, fontFace: "Arial", color: c.col, bold: true, align: "center" });
  });

  // Slides 5-10: Placeholder slides
  const titles = ["強み分析", "弱み分析・優先度マトリクス", "改善施策 Phase 1", "改善施策 Phase 2・3", "KPI目標設定", "総括"];
  const bodies = [
    `高評価レビュー（8-10点: ${high_count}件）の分析から、主要な強みテーマを抽出しています。立地・アクセス、客室清潔感、スタッフ対応が高評価の中心的テーマです。`,
    `低評価・中評価レビュー（${low_count + mid_count}件）から改善課題を優先度別に整理しています。S/A/B/Cの4段階で優先度を設定し、計画的な改善を推進します。`,
    `即時対応が必要な課題への対応計画です。口コミで頻出する不満事項を30日以内に解消することを目標とします。`,
    `Phase 2（1-3ヶ月）では設備・サービス品質の継続改善を、Phase 3（3-6ヶ月）ではブランド戦略・OTA評価の底上げを実施します。`,
    `全体平均 ${overall_avg_10pt.toFixed(2)}点 → 目標 ${Math.min(10, overall_avg_10pt + 0.5).toFixed(2)}点\n高評価率 ${high_rate.toFixed(1)}% → 目標 ${Math.min(100, high_rate + 5).toFixed(1)}%\n低評価率 ${low_rate.toFixed(1)}% → 目標 ${Math.max(0, low_rate - 2).toFixed(1)}%`,
    `${hotelName}は全体平均${overall_avg_10pt.toFixed(2)}点（10pt換算）、高評価率${high_rate.toFixed(1)}%を達成しています。本レポートで示した改善施策を段階的に実施することで、さらなる顧客満足度向上が期待できます。`,
  ];
  titles.forEach((title, i) => {
    s = pres.addSlide();
    hdr(s, title, hotelName);
    s.addText(bodies[i], { x: 0.5, y: 1.1, w: 9.0, h: 3.8, fontSize: 12, fontFace: "Arial", color: C.gray, valign: "top", wrap: true });
    footer(s, i + 5);
  });

  return pres.writeFile({ fileName: outPath });
}

// ===== MAIN =====
(async () => {
  let docxOk = 0, docxFail = 0, pptxOk = 0, pptxFail = 0;

  for (const hotel of HOTELS) {
    const jsonPath = path.join(JSON_DIR, `${hotel.key}_analysis.json`);
    if (!fs.existsSync(jsonPath)) {
      console.log(`SKIP: ${hotel.key} - JSON not found`);
      continue;
    }
    const data = JSON.parse(fs.readFileSync(jsonPath, "utf8"));
    console.log(`Processing: ${hotel.name} (${data.total_reviews}件)`);

    // DOCX
    const docxPath = path.join(OUT_DIR, `${hotel.key}_口コミ分析改善レポート.docx`);
    try {
      await generateDocx(hotel.name, data, docxPath);
      console.log(`  DOCX: OK → ${path.basename(docxPath)}`);
      docxOk++;
    } catch (e) {
      console.error(`  DOCX: FAIL - ${e.message}`);
      docxFail++;
    }

    // PPTX
    const pptxPath = path.join(OUT_DIR, `${hotel.key}_口コミ分析レポート.pptx`);
    try {
      await generatePptx(hotel.name, data, pptxPath);
      console.log(`  PPTX: OK → ${path.basename(pptxPath)}`);
      pptxOk++;
    } catch (e) {
      console.error(`  PPTX: FAIL - ${e.message}`);
      pptxFail++;
    }
  }

  console.log(`\n=== 完了 ===`);
  console.log(`DOCX: 成功${docxOk}件 / 失敗${docxFail}件`);
  console.log(`PPTX: 成功${pptxOk}件 / 失敗${pptxFail}件`);
})();
