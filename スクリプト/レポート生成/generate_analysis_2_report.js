#!/usr/bin/env node
"use strict";
const fs = require("fs");
const path = require("path");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType, PageNumber, PageBreak } = require("docx");
const pptxgen = require("pptxgenjs");

const data = JSON.parse(fs.readFileSync(path.resolve(__dirname, "analysis_2_data.json"), "utf-8"));
const meta = data.analysis_metadata;
const maidClaims = data.maid_claims_summary;
const checkerClaims = data.checker_claims_summary;
const prod = data.maid_productivity;
const attend = data.attendance_analysis;
const hotels = data.hotel_summaries;
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
  // Maid claims ranking table (top 15)
  const maidHeader = new TableRow({ children: [
    cell("順位", { bg: C.NAVY, c: C.WHITE, b: true, w: 700 }),
    cell("名前", { bg: C.NAVY, c: C.WHITE, b: true, w: 1800 }),
    cell("ホテル", { bg: C.NAVY, c: C.WHITE, b: true, w: 4000 }),
    cell("クレーム件数", { bg: C.NAVY, c: C.WHITE, b: true, w: 1400 }),
  ] });
  const maidRows = maidClaims.top_claim_maids.map((m, i) => new TableRow({ children: [
    cell(i + 1, { bg: i % 2 ? C.WHITE : C.LIGHT_BG }),
    cell(m.name, { bg: i % 2 ? C.WHITE : C.LIGHT_BG, a: AlignmentType.LEFT }),
    cell(m.hotel, { bg: i % 2 ? C.WHITE : C.LIGHT_BG, a: AlignmentType.LEFT, sz: 14 }),
    cell(m.claims, { bg: i % 2 ? C.WHITE : C.LIGHT_BG, b: true, c: m.claims >= 2 ? C.RED : C.TEXT }),
  ] }));
  const maidTable = new Table({ rows: [maidHeader, ...maidRows], width: { size: 7900, type: WidthType.DXA } });

  // Checker claims ranking table (top 15)
  const checkerHeader = new TableRow({ children: [
    cell("順位", { bg: C.ACCENT, c: C.WHITE, b: true, w: 700 }),
    cell("名前", { bg: C.ACCENT, c: C.WHITE, b: true, w: 1800 }),
    cell("ホテル", { bg: C.ACCENT, c: C.WHITE, b: true, w: 4000 }),
    cell("クレーム件数", { bg: C.ACCENT, c: C.WHITE, b: true, w: 1400 }),
  ] });
  const checkerRows = checkerClaims.top_claim_checkers.map((c, i) => new TableRow({ children: [
    cell(i + 1, { bg: i % 2 ? C.WHITE : C.LIGHT_BG }),
    cell(c.name, { bg: i % 2 ? C.WHITE : C.LIGHT_BG, a: AlignmentType.LEFT }),
    cell(c.hotel, { bg: i % 2 ? C.WHITE : C.LIGHT_BG, a: AlignmentType.LEFT, sz: 14 }),
    cell(c.claims, { bg: i % 2 ? C.WHITE : C.LIGHT_BG, b: true, c: c.claims >= 3 ? C.RED : c.claims >= 2 ? C.ORANGE : C.TEXT }),
  ] }));
  const checkerTable = new Table({ rows: [checkerHeader, ...checkerRows], width: { size: 7900, type: WidthType.DXA } });

  // Productivity top performers table
  const prodHeader = new TableRow({ children: [
    cell("順位", { bg: C.TEAL, c: C.WHITE, b: true, w: 600 }),
    cell("名前", { bg: C.TEAL, c: C.WHITE, b: true, w: 1600 }),
    cell("ホテル", { bg: C.TEAL, c: C.WHITE, b: true, w: 2800 }),
    cell("出勤日数", { bg: C.TEAL, c: C.WHITE, b: true, w: 900 }),
    cell("清掃室数", { bg: C.TEAL, c: C.WHITE, b: true, w: 900 }),
    cell("室/日", { bg: C.TEAL, c: C.WHITE, b: true, w: 900 }),
    cell("報酬形態", { bg: C.TEAL, c: C.WHITE, b: true, w: 900 }),
  ] });
  const prodRows = prod.top_performers.map((tp, i) => new TableRow({ children: [
    cell(i + 1, { bg: i % 2 ? C.WHITE : C.LIGHT_BG }),
    cell(tp.name, { bg: i % 2 ? C.WHITE : C.LIGHT_BG, a: AlignmentType.LEFT }),
    cell(tp.hotel, { bg: i % 2 ? C.WHITE : C.LIGHT_BG, a: AlignmentType.LEFT, sz: 14 }),
    cell(tp.total_days, { bg: i % 2 ? C.WHITE : C.LIGHT_BG }),
    cell(tp.rooms_cleaned, { bg: i % 2 ? C.WHITE : C.LIGHT_BG }),
    cell(fmtN(tp.rooms_per_day), { bg: i % 2 ? C.WHITE : C.LIGHT_BG, b: true, c: C.TEAL }),
    cell(tp.pay_type || "-", { bg: i % 2 ? C.WHITE : C.LIGHT_BG }),
  ] }));
  const prodTable = new Table({ rows: [prodHeader, ...prodRows], width: { size: 8600, type: WidthType.DXA } });

  // Hotel-level comparison table
  const hotelHeader = new TableRow({ children: [
    cell("ホテル名", { bg: C.NAVY, c: C.WHITE, b: true, w: 2800 }),
    cell("人数", { bg: C.NAVY, c: C.WHITE, b: true, w: 800 }),
    cell("メイド", { bg: C.NAVY, c: C.WHITE, b: true, w: 800 }),
    cell("チェッカー", { bg: C.NAVY, c: C.WHITE, b: true, w: 900 }),
    cell("クレーム", { bg: C.NAVY, c: C.WHITE, b: true, w: 900 }),
    cell("件/メイド", { bg: C.NAVY, c: C.WHITE, b: true, w: 900 }),
    cell("室/日", { bg: C.NAVY, c: C.WHITE, b: true, w: 800 }),
    cell("出勤日", { bg: C.NAVY, c: C.WHITE, b: true, w: 800 }),
  ] });
  const hotelRows = hotels.map((h, i) => {
    const claimColor = h.claims_per_maid >= 0.5 ? C.RED : h.claims_per_maid >= 0.2 ? C.ORANGE : C.TEXT;
    return new TableRow({ children: [
      cell(h.name, { bg: i % 2 ? C.WHITE : C.LIGHT_BG, a: AlignmentType.LEFT, sz: 13 }),
      cell(h.roster_size, { bg: i % 2 ? C.WHITE : C.LIGHT_BG }),
      cell(h.maid_count || "-", { bg: i % 2 ? C.WHITE : C.LIGHT_BG }),
      cell(h.checker_count || "-", { bg: i % 2 ? C.WHITE : C.LIGHT_BG }),
      cell(h.total_claims, { bg: i % 2 ? C.WHITE : C.LIGHT_BG }),
      cell(h.claims_per_maid != null ? fmtN(h.claims_per_maid, 2) : "-", { bg: i % 2 ? C.WHITE : C.LIGHT_BG, b: true, c: claimColor }),
      cell(h.avg_rooms_per_day ? fmtN(h.avg_rooms_per_day) : "-", { bg: i % 2 ? C.WHITE : C.LIGHT_BG }),
      cell(fmtN(h.avg_attendance_days), { bg: i % 2 ? C.WHITE : C.LIGHT_BG }),
    ] });
  });
  const hotelTable = new Table({ rows: [hotelHeader, ...hotelRows], width: { size: 8700, type: WidthType.DXA } });

  const doc = new Document({
    styles: { default: { document: { run: { font: "Arial", size: 20 } } } },
    sections: [{
      headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "PRIMECHANGE | 分析2: スタッフ個人別パフォーマンス", size: 14, color: C.SUBTEXT, font: "Arial" })] })] }) },
      footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Page ", size: 14, color: C.SUBTEXT }), new TextRun({ children: [PageNumber.CURRENT], size: 14, color: C.SUBTEXT })] })] }) },
      children: [
        // Cover
        new Paragraph({ spacing: { before: 2400 } }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 }, children: [new TextRun({ text: "PRIMECHANGE", size: 56, font: "Arial", bold: true, color: C.NAVY })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [new TextRun({ text: "分析2", size: 44, font: "Arial", bold: true, color: C.ACCENT })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [new TextRun({ text: "スタッフ個人別パフォーマンス分析", size: 36, font: "Arial", bold: true, color: C.ACCENT })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 400 }, children: [new TextRun({ text: "Individual Staff Performance Analysis", size: 22, font: "Arial", color: C.SUBTEXT, italics: true })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "株式会社PRIMECHANGE", size: 28, font: "Arial", bold: true, color: C.NAVY })] }),
        PB(),

        // Executive Summary
        h1("1. エグゼクティブサマリー"),
        p(`${meta.total_hotels}ホテルのスタッフ個人別パフォーマンスデータを分析しました。対象スタッフ${meta.total_staff_analyzed}名、データ期間は${meta.data_period}です。`),
        h3("主要な発見"),
        bp(`分析対象スタッフ総数: ${meta.total_staff_analyzed}名（${meta.total_hotels}ホテル）`, { b: true }),
        bp(`クレーム発生メイド: ${maidClaims.total_maids_with_claims}名。上位13名が各2件以上のクレームを記録。`, { b: true }),
        bp(`メイド1日あたり平均清掃室数: ${fmtN(prod.avg_rooms_per_day)}室（中央値${fmtN(prod.median_rooms_per_day)}室）。最高${fmtN(prod.max_rooms_per_day)}室から最低${fmtN(prod.min_rooms_per_day)}室まで個人差が極めて大きい。`, { b: true }),
        bp(`チェッカークレーム: ${checkerClaims.total_checkers_with_claims}名のチェッカーがクレーム関連。「なし」（チェッカー不在時）が最多の9件。`),
        bp(`出勤分析: 全${attend.total_staff}名の平均出勤日数${fmtN(attend.avg_days)}日。高出勤者${attend.high_attendance}名、低出勤者${attend.low_attendance}名。`),
        PB(),

        // Maid Claims Ranking
        h1("2. メイド別クレームランキング"),
        p(`クレームが報告された${maidClaims.total_maids_with_claims}名のメイドのうち、上位15名を以下に示します。`),
        h3("上位15名（クレーム件数順）"),
        maidTable,
        p(""),
        bp("ダイワロイネットホテル東京大崎、アパホテル蒲田駅東からそれぞれ複数名がランクイン。ホテル固有の構造的要因の可能性。", { b: true }),
        bp("クレーム2件以上のメイドは全13名。個別面談による原因分析と再発防止が急務。"),
        PB(),

        // Checker Claims Analysis
        h1("3. チェッカー別クレーム分析"),
        p(`クレーム関連のチェッカー${checkerClaims.total_checkers_with_claims}名の上位15名を以下に示します。`),
        h3("チェッカー別クレーム上位15名"),
        checkerTable,
        p(""),
        bp("「なし」（チェッカー不在時）が9件で最多。チェッカー不在がクレーム発生の最大リスク要因。", { b: true, c: C.RED }),
        bp("結城（ダイワロイネット大崎）が4件、ピョー（同）が3件。特定チェッカーの見逃し傾向を要分析。"),
        bp("ハートンホテル東品川から4名がランクイン。チェック体制の見直しが必要。"),
        PB(),

        // Productivity Analysis
        h1("4. 生産性分析（清掃室数/日）"),
        p(`室数データのあるメイド${prod.total_maids_with_room_data}名の1日あたり清掃室数を分析しました。`),
        h3("生産性サマリー"),
        bp(`平均: ${fmtN(prod.avg_rooms_per_day)}室/日`, { b: true }),
        bp(`中央値: ${fmtN(prod.median_rooms_per_day)}室/日`),
        bp(`最高: ${fmtN(prod.max_rooms_per_day)}室/日（楊 - チサンホテル浜松町）`, { b: true, c: C.GREEN }),
        bp(`最低: ${fmtN(prod.min_rooms_per_day)}室/日`),
        p(""),
        h3("トップパフォーマー10名"),
        prodTable,
        p(""),
        bp("チサンホテル浜松町のメイドが上位10名中5名を占有。歩合制メイドの生産性が突出。", { b: true }),
        bp("ハートンホテル東品川からも4名がランクイン。高稼働ホテルでの効率的清掃が特徴。"),
        bp("トップパフォーマーは全員欠勤0日。高い出勤率と生産性に正の相関。"),
        PB(),

        // Attendance Analysis
        h1("5. 出勤分析"),
        p(`全${attend.total_staff}名のスタッフ出勤状況を分析しました。`),
        h3("出勤サマリー"),
        bp(`平均出勤日数: ${fmtN(attend.avg_days)}日`, { b: true }),
        bp(`高出勤者（20日以上）: ${attend.high_attendance}名（${fmtN(attend.high_attendance / attend.total_staff * 100)}%）`, { c: C.GREEN }),
        bp(`低出勤者（10日未満）: ${attend.low_attendance}名（${fmtN(attend.low_attendance / attend.total_staff * 100)}%）`, { c: C.ORANGE }),
        p(""),
        bp("チサンホテル浜松町、リッチモンドホテル東京目白が平均25.6日と最高出勤率。"),
        bp("コンフォートイン六本木、コンフォートホテル成田が13日台と低出勤傾向。パートタイム比率の影響。"),
        PB(),

        // Hotel-level comparison table
        h1("6. ホテル別スタッフパフォーマンス比較"),
        p("全19ホテルのスタッフ配置・クレーム・生産性・出勤状況の一覧です。"),
        hotelTable,
        p(""),
        bp("川崎日航ホテル: メイド1人あたりクレーム0.83件で最高。重点改善対象。", { b: true, c: C.RED }),
        bp("ダイワロイネットホテル東京大崎: 0.67件/メイドで第2位。少人数体制での負荷集中が要因の可能性。", { c: C.RED }),
        bp("コンフォートホテル成田・横浜関内: クレーム0件で模範ホテル。", { c: C.GREEN }),
        PB(),

        // Recommendations
        h1("7. 提言"),
        ...recs.flatMap((rec, i) => [
          h2(`7.${i + 1} ${rec.title}`),
          p(`【優先度: ${rec.priority}】`, { b: true, c: rec.priority === "最優先" ? C.RED : rec.priority === "高" ? C.ORANGE : C.TEAL }),
          p(rec.rationale),
          ...rec.actions.map(a => bp(a)),
        ]),
      ]
    }]
  });

  const buf = await Packer.toBuffer(doc);
  const outPath = path.resolve(__dirname, "PRIMECHANGE_分析2_スタッフパフォーマンス.docx");
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
  s1.addText("分析2: スタッフ個人別パフォーマンス分析", { x: 0.5, y: 2.2, w: 9, h: 0.6, fontSize: 24, bold: true, color: C.WHITE, fontFace: "Arial", align: "center" });
  s1.addText(`${meta.total_hotels}ホテル・${meta.total_staff_analyzed}名のスタッフパフォーマンスデータに基づく分析`, { x: 0.5, y: 3.4, w: 9, h: 0.3, fontSize: 12, color: "CCCCCC", fontFace: "Arial", align: "center" });

  // Slide 2: KPI cards
  const s2 = pptx.addSlide();
  s2.addText("主要KPI", { ...titleOpts });
  const kpis = [
    { label: "分析対象スタッフ", value: `${meta.total_staff_analyzed}名`, sub: `${meta.total_hotels}ホテル`, color: C.ACCENT },
    { label: "クレーム発生メイド", value: `${maidClaims.total_maids_with_claims}名`, sub: "個別研修対象", color: C.RED },
    { label: "平均清掃室数/日", value: `${fmtN(prod.avg_rooms_per_day)}室`, sub: `最高${fmtN(prod.max_rooms_per_day)}室`, color: C.TEAL },
    { label: "平均出勤日数", value: `${fmtN(attend.avg_days)}日`, sub: `高出勤${attend.high_attendance}名`, color: C.ORANGE },
  ];
  kpis.forEach((k, i) => {
    const x = 0.3 + i * 2.35;
    s2.addShape(pptx.ShapeType.roundRect, { x, y: 1.4, w: 2.15, h: 1.6, fill: { color: k.color }, rectRadius: 0.1 });
    s2.addText(k.label, { x, y: 1.5, w: 2.15, h: 0.3, fontSize: 10, color: C.WHITE, fontFace: "Arial", align: "center" });
    s2.addText(k.value, { x, y: 1.9, w: 2.15, h: 0.5, fontSize: 24, bold: true, color: C.WHITE, fontFace: "Arial", align: "center" });
    s2.addText(k.sub, { x, y: 2.5, w: 2.15, h: 0.3, fontSize: 9, color: C.WHITE, fontFace: "Arial", align: "center" });
  });
  s2.addText(`• クレーム2件以上のメイド13名に個別研修が急務`, { x: 0.5, y: 3.5, w: 9, h: 0.3, fontSize: 11, color: C.TEXT, fontFace: "Arial" });
  s2.addText(`• チェッカー不在時のクレーム9件 — チェック体制の見直しが必要`, { x: 0.5, y: 3.9, w: 9, h: 0.3, fontSize: 11, color: C.TEXT, fontFace: "Arial" });

  // Slide 3: Claims analysis by hotel
  const s3 = pptx.addSlide();
  s3.addText("ホテル別クレーム分析", { ...titleOpts });
  const claimHotels = hotels.filter(h => h.total_claims > 0).sort((a, b) => b.total_claims - a.total_claims).slice(0, 12);
  const tblRows = [
    [{ text: "ホテル", options: { bold: true, color: C.WHITE, fill: { color: C.NAVY }, fontSize: 8 } },
     { text: "人数", options: { bold: true, color: C.WHITE, fill: { color: C.NAVY }, fontSize: 8 } },
     { text: "メイド", options: { bold: true, color: C.WHITE, fill: { color: C.NAVY }, fontSize: 8 } },
     { text: "クレーム", options: { bold: true, color: C.WHITE, fill: { color: C.NAVY }, fontSize: 8 } },
     { text: "件/メイド", options: { bold: true, color: C.WHITE, fill: { color: C.NAVY }, fontSize: 8 } }],
    ...claimHotels.map((h, i) => [
      { text: h.name.replace(/ホテル/g, "H.").substring(0, 20), options: { fontSize: 7, fill: { color: i % 2 ? "FFFFFF" : "F5F7FA" } } },
      { text: String(h.roster_size), options: { fontSize: 8, fill: { color: i % 2 ? "FFFFFF" : "F5F7FA" } } },
      { text: String(h.maid_count || "-"), options: { fontSize: 8, fill: { color: i % 2 ? "FFFFFF" : "F5F7FA" } } },
      { text: String(h.total_claims), options: { fontSize: 8, bold: true, fill: { color: i % 2 ? "FFFFFF" : "F5F7FA" } } },
      { text: h.claims_per_maid != null ? fmtN(h.claims_per_maid, 2) : "-", options: { fontSize: 8, bold: true, color: h.claims_per_maid >= 0.5 ? C.RED : C.TEXT, fill: { color: i % 2 ? "FFFFFF" : "F5F7FA" } } },
    ])
  ];
  s3.addTable(tblRows, { x: 0.3, y: 1.2, w: 9, h: 3.5, fontSize: 8, fontFace: "Arial", border: { color: "CCCCCC", pt: 0.5 }, colW: [3.2, 1.0, 1.0, 1.2, 1.2], align: "center" });
  s3.addText("川崎日航H.・ダイワロイネットH.東京大崎が件/メイド比で要改善", { x: 0.5, y: 4.6, w: 9, h: 0.3, fontSize: 10, bold: true, color: C.RED, fontFace: "Arial" });

  // Slide 4: Productivity comparison
  const s4 = pptx.addSlide();
  s4.addText("生産性比較（清掃室数/日）", { ...titleOpts });
  const topProd = prod.top_performers.slice(0, 5);
  const prodTblRows = [
    [{ text: "名前", options: { bold: true, color: C.WHITE, fill: { color: C.TEAL }, fontSize: 9 } },
     { text: "ホテル", options: { bold: true, color: C.WHITE, fill: { color: C.TEAL }, fontSize: 9 } },
     { text: "出勤日", options: { bold: true, color: C.WHITE, fill: { color: C.TEAL }, fontSize: 9 } },
     { text: "清掃室数", options: { bold: true, color: C.WHITE, fill: { color: C.TEAL }, fontSize: 9 } },
     { text: "室/日", options: { bold: true, color: C.WHITE, fill: { color: C.TEAL }, fontSize: 9 } }],
    ...topProd.map((tp, i) => [
      { text: tp.name, options: { fontSize: 9, fill: { color: i % 2 ? "FFFFFF" : "F5F7FA" } } },
      { text: tp.hotel.replace(/ホテル/g, "H.").substring(0, 18), options: { fontSize: 8, fill: { color: i % 2 ? "FFFFFF" : "F5F7FA" } } },
      { text: String(tp.total_days), options: { fontSize: 9, fill: { color: i % 2 ? "FFFFFF" : "F5F7FA" } } },
      { text: String(tp.rooms_cleaned), options: { fontSize: 9, fill: { color: i % 2 ? "FFFFFF" : "F5F7FA" } } },
      { text: fmtN(tp.rooms_per_day), options: { fontSize: 9, bold: true, color: C.TEAL, fill: { color: i % 2 ? "FFFFFF" : "F5F7FA" } } },
    ])
  ];
  s4.addTable(prodTblRows, { x: 0.3, y: 1.2, w: 9, h: 2.0, fontSize: 9, fontFace: "Arial", border: { color: "CCCCCC", pt: 0.5 }, colW: [1.8, 3.0, 1.0, 1.2, 1.0], align: "center" });
  // Summary boxes
  [['平均', `${fmtN(prod.avg_rooms_per_day)}室/日`, C.ACCENT, 0.5], ['中央値', `${fmtN(prod.median_rooms_per_day)}室/日`, C.TEAL, 3.2], ['最高', `${fmtN(prod.max_rooms_per_day)}室/日`, C.GREEN, 5.9]].forEach(([label, value, color, x]) => {
    s4.addShape(pptx.ShapeType.roundRect, { x, y: 3.6, w: 2.4, h: 0.9, fill: { color }, rectRadius: 0.08 });
    s4.addText(label, { x, y: 3.65, w: 2.4, h: 0.3, fontSize: 10, color: C.WHITE, fontFace: "Arial", align: "center" });
    s4.addText(value, { x, y: 3.95, w: 2.4, h: 0.4, fontSize: 18, bold: true, color: C.WHITE, fontFace: "Arial", align: "center" });
  });

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

  const outPath = path.resolve(__dirname, "PRIMECHANGE_分析2_スタッフパフォーマンス.pptx");
  pptx.writeFile({ fileName: outPath }).then(() => {
    console.log(`PPTX: ${outPath} (${(fs.statSync(outPath).size / 1024).toFixed(1)} KB)`);
  });
}

async function main() {
  console.log("Building Analysis 2 Reports...");
  await buildDOCX();
  await buildPPTX();
  console.log("\nDone!");
}
main().catch(e => { console.error(e); process.exit(1); });
