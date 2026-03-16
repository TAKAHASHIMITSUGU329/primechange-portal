#!/usr/bin/env node
"use strict";

const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, PageNumber, PageBreak, LevelFormat
} = require("docx");
const pptxgen = require("pptxgenjs");

// ============================================================
// Paths
// ============================================================
const JSON_PATH = path.resolve(__dirname, "primechange_portfolio_analysis.json");
const OUTPUT_DIR = __dirname;
const DOCX_NAME = "PRIMECHANGE_清掃戦略レポート.docx";
const PPTX_NAME = "PRIMECHANGE_清掃戦略レポート.pptx";

// ============================================================
// Load data
// ============================================================
if (!fs.existsSync(JSON_PATH)) {
  console.error("Error: JSON file not found:", JSON_PATH);
  process.exit(1);
}
const data = JSON.parse(fs.readFileSync(JSON_PATH, "utf-8"));

const meta = data.report_metadata || {};
const overview = data.portfolio_overview || {};
const deepDive = data.cleaning_deep_dive || {};
const priorityMatrix = data.priority_matrix || {};
const actionPlans = data.action_plans || [];
const crossCutting = data.cross_cutting_recommendations || [];
const kpiFramework = data.kpi_framework || {};
const roiEstimation = data.roi_estimation || {};

const hotelsRanked = overview.hotels_ranked || [];
const categorySummary = deepDive.category_summary || [];
const cleaningMatrix = deepDive.hotel_cleaning_matrix || [];
const portfolioTargets = kpiFramework.portfolio_targets || [];
const perHotelTargets = kpiFramework.per_hotel_targets || [];
const scenarios = roiEstimation.scenarios || [];

// ============================================================
// Color Palette
// ============================================================
const C = {
  NAVY: "1B3A5C",
  ACCENT: "2E75B6",
  WHITE: "FFFFFF",
  LIGHT_BG: "F5F7FA",
  TEXT: "333333",
  SUBTEXT: "666666",
  GREEN: "27AE60",
  ORANGE: "FF9800",
  RED: "E74C3C",
  BLUE: "2196F3",
  YELLOW: "FFC107",
};

const PRIORITY_COLORS = {
  URGENT:      { bg: "FFEBEE", text: "E74C3C" },
  HIGH:        { bg: "FFF3E0", text: "FF9800" },
  STANDARD:    { bg: "E3F2FD", text: "2196F3" },
  MAINTENANCE: { bg: "E8F5E9", text: "27AE60" },
};

function priorityLabel(p) {
  const labels = { URGENT: "緊急", HIGH: "高", STANDARD: "標準", MAINTENANCE: "維持" };
  return labels[p] || p;
}

// ============================================================
// DOCX Helpers
// ============================================================
const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };
const noBorders = {
  top: { style: BorderStyle.NONE, size: 0 },
  bottom: { style: BorderStyle.NONE, size: 0 },
  left: { style: BorderStyle.NONE, size: 0 },
  right: { style: BorderStyle.NONE, size: 0 },
};
const cellMargins = { top: 80, bottom: 80, left: 120, right: 120 };

function heading1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 200 },
    children: [new TextRun({ text, bold: true, size: 32, font: "Arial", color: C.NAVY })],
  });
}

function heading2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 280, after: 160 },
    children: [new TextRun({ text, bold: true, size: 26, font: "Arial", color: C.ACCENT })],
  });
}

function heading3(text) {
  return new Paragraph({
    spacing: { before: 200, after: 120 },
    children: [new TextRun({ text, bold: true, size: 22, font: "Arial", color: C.NAVY })],
  });
}

function para(text, opts = {}) {
  return new Paragraph({
    spacing: { after: opts.afterSpacing || 120, line: 320 },
    alignment: opts.alignment || AlignmentType.LEFT,
    children: [new TextRun({ text, size: 21, font: "Arial", color: opts.color || C.TEXT, ...opts })],
  });
}

function multiRunPara(runs, opts = {}) {
  return new Paragraph({
    spacing: { after: opts.afterSpacing || 120, line: 320 },
    alignment: opts.alignment || AlignmentType.LEFT,
    children: runs.map(r => new TextRun({ size: 21, font: "Arial", color: C.TEXT, ...r })),
  });
}

function bulletItem(text, opts = {}) {
  return new Paragraph({
    numbering: { reference: "bullets", level: opts.level || 0 },
    spacing: { after: 80, line: 300 },
    children: [new TextRun({ text, size: 21, font: "Arial", color: C.TEXT })],
  });
}

function headerCell(text, width, opts = {}) {
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.PERCENTAGE },
    shading: { fill: C.NAVY, type: ShadingType.CLEAR },
    margins: cellMargins,
    verticalAlign: "center",
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text, bold: true, size: 18, font: "Arial", color: C.WHITE })],
    })],
  });
}

function dataCell(text, width, opts = {}) {
  const shadingObj = opts.fill ? { fill: opts.fill, type: ShadingType.CLEAR } : undefined;
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.PERCENTAGE },
    shading: shadingObj,
    margins: cellMargins,
    verticalAlign: "center",
    children: [new Paragraph({
      alignment: opts.alignment || AlignmentType.LEFT,
      children: [new TextRun({ text: String(text != null ? text : ""), size: 18, font: "Arial", color: opts.color || C.TEXT, bold: opts.bold || false })],
    })],
  });
}

function spacer(height = 100) {
  return new Paragraph({ spacing: { after: height }, children: [] });
}

function divider() {
  return new Paragraph({
    spacing: { before: 200, after: 200 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: C.ACCENT, space: 1 } },
    children: [],
  });
}

function kpiCardRow(items) {
  const colWidth = Math.floor(100 / items.length);
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({
        children: items.map(item =>
          new TableCell({
            borders: { ...noBorders, right: { style: BorderStyle.SINGLE, size: 1, color: "DDDDDD" } },
            width: { size: colWidth, type: WidthType.PERCENTAGE },
            shading: { fill: item.bgColor || C.LIGHT_BG, type: ShadingType.CLEAR },
            margins: { top: 160, bottom: 160, left: 200, right: 200 },
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 40 },
                children: [new TextRun({ text: item.label, size: 18, font: "Arial", color: C.SUBTEXT })],
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [new TextRun({ text: item.value, bold: true, size: 36, font: "Arial", color: item.color || C.NAVY })],
              }),
            ],
          })
        ),
      }),
    ],
  });
}

// ============================================================
// Chapter Builders - DOCX
// ============================================================

// --- COVER PAGE ---
function buildCoverPage() {
  return {
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
      },
    },
    children: [
      spacer(2000),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 80 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: C.ACCENT, space: 8 } },
        children: [new TextRun({ text: "PRIMECHANGE", size: 56, font: "Arial", color: C.ACCENT })],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 200 },
        children: [new TextRun({ text: "ホテル清掃戦略レポート", size: 56, font: "Arial", bold: true, color: C.NAVY })],
      }),
      spacer(200),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 120 },
        children: [new TextRun({ text: meta.subtitle || "ゲスト口コミデータに基づく清掃品質改善提案書", size: 24, font: "Arial", color: C.SUBTEXT })],
      }),
      spacer(200),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 80 },
        children: [new TextRun({ text: `作成日：${meta.date || "2026年3月8日"}`, size: 22, font: "Arial", color: C.SUBTEXT })],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 80 },
        children: [new TextRun({ text: `対象：${meta.total_hotels || 19}ホテル`, size: 22, font: "Arial", color: C.SUBTEXT })],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 80 },
        children: [new TextRun({ text: `分析期間：${meta.analysis_period || ""}`, size: 20, font: "Arial", color: C.SUBTEXT })],
      }),
      spacer(1600),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 80 },
        children: [new TextRun({ text: `提出先：${meta.prepared_for || ""}`, size: 20, font: "Arial", color: "999999" })],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 80 },
        children: [new TextRun({ text: `作成：${meta.prepared_by || ""}`, size: 20, font: "Arial", color: "999999" })],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "Confidential - For Internal Use Only", size: 18, font: "Arial", color: "AAAAAA", italics: true })],
      }),
    ],
  };
}

// --- CHAPTER 1: EXECUTIVE SUMMARY ---
function buildChapter1() {
  const urgentCount = (priorityMatrix.urgent || []).length;
  const highCount = (priorityMatrix.high || []).length;
  const standardCount = (priorityMatrix.standard || []).length;
  const maintenanceCount = (priorityMatrix.maintenance || []).length;

  const urgentAvgs = (priorityMatrix.urgent || []).map(h => h.avg);
  const highAvgs = (priorityMatrix.high || []).map(h => h.avg);
  const standardAvgs = (priorityMatrix.standard || []).map(h => h.avg);
  const maintenanceAvgs = (priorityMatrix.maintenance || []).map(h => h.avg);

  function rangeStr(avgs) {
    if (avgs.length === 0) return "-";
    if (avgs.length === 1) return avgs[0].toFixed(2);
    return `${Math.min(...avgs).toFixed(2)}-${Math.max(...avgs).toFixed(2)}`;
  }

  function mainIssues(hotels) {
    const issues = {};
    hotels.forEach(h => (h.key_problems || []).forEach(p => { issues[p] = (issues[p] || 0) + 1; }));
    return Object.entries(issues).sort((a, b) => b[1] - a[1]).slice(0, 3).map(([k]) => k).join("、") || "-";
  }

  const priorityTableRows = [
    ["URGENT", String(urgentCount), rangeStr(urgentAvgs), mainIssues(priorityMatrix.urgent || [])],
    ["HIGH", String(highCount), rangeStr(highAvgs), mainIssues(priorityMatrix.high || [])],
    ["STANDARD", String(standardCount), rangeStr(standardAvgs), mainIssues(priorityMatrix.standard || [])],
    ["MAINTENANCE", String(maintenanceCount), rangeStr(maintenanceAvgs), mainIssues(priorityMatrix.maintenance || [])],
  ];

  const cleaningRate = deepDive.portfolio_cleaning_issue_rate || 4.6;

  const children = [
    heading1("1. エグゼクティブサマリー"),
    para(`ポートフォリオ全体（${meta.total_hotels || 19}ホテル、${meta.total_reviews || 1583}件のレビュー）を分析した結果、ポートフォリオ平均スコアは${overview.avg_score || 8.39}点（10点換算）、清掃関連クレーム率は${cleaningRate}%でした。以下に主要な発見事項と提案の概要をまとめます。`),
    spacer(100),

    kpiCardRow([
      { label: "ポートフォリオ平均", value: String(overview.avg_score || 8.39), color: C.ACCENT, bgColor: "E3F2FD" },
      { label: "高評価率", value: (overview.portfolio_high_rate || 78.1) + "%", color: C.GREEN, bgColor: "E8F5E9" },
      { label: "清掃クレーム率", value: cleaningRate + "%", color: C.RED, bgColor: "FFEBEE" },
      { label: "レビュー総数", value: String(meta.total_reviews || 1583) + "件", color: C.NAVY, bgColor: C.LIGHT_BG },
    ]),
    spacer(200),

    // KEY FINDINGS box
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [new TableRow({
        children: [new TableCell({
          borders: { top: { style: BorderStyle.SINGLE, size: 4, color: C.ACCENT }, bottom: border, left: { style: BorderStyle.SINGLE, size: 4, color: C.ACCENT }, right: border },
          width: { size: 100, type: WidthType.PERCENTAGE },
          shading: { fill: "F0F7FC", type: ShadingType.CLEAR },
          margins: { top: 200, bottom: 200, left: 300, right: 300 },
          children: [
            new Paragraph({ spacing: { after: 120 }, children: [new TextRun({ text: "KEY FINDINGS", bold: true, size: 22, font: "Arial", color: C.ACCENT })] }),
            new Paragraph({ spacing: { after: 80 }, children: [
              new TextRun({ text: "Strength：", bold: true, size: 20, font: "Arial", color: C.GREEN }),
              new TextRun({ text: `上位5ホテルの平均スコアは8.7点超。${overview.best_hotel ? overview.best_hotel.name + "(" + overview.best_hotel.avg + "点)" : ""}がポートフォリオ最高評価`, size: 20, font: "Arial", color: C.TEXT }),
            ]}),
            new Paragraph({ spacing: { after: 80 }, children: [
              new TextRun({ text: "Challenge：", bold: true, size: 20, font: "Arial", color: C.RED }),
              new TextRun({ text: `${urgentCount}ホテルがURGENT（緊急対応）に分類。清掃関連クレーム率${cleaningRate}%（業界平均2-3%を超過）`, size: 20, font: "Arial", color: C.TEXT }),
            ]}),
            new Paragraph({ children: [
              new TextRun({ text: "Opportunity：", bold: true, size: 20, font: "Arial", color: C.ORANGE }),
              new TextRun({ text: "清掃品質の標準化により、URGENTホテルのスコアを+0.5-1.0pt改善し、RevPAR 5-10%向上が見込まれます", size: 20, font: "Arial", color: C.TEXT }),
            ]}),
          ],
        })],
      })],
    }),
    spacer(200),

    heading2("優先度概要"),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({ children: [
          headerCell("優先度", 15),
          headerCell("ホテル数", 15),
          headerCell("スコアレンジ", 25),
          headerCell("主要課題", 45),
        ]}),
        ...priorityTableRows.map((row, idx) => {
          const p = row[0];
          const pc = PRIORITY_COLORS[p] || PRIORITY_COLORS.STANDARD;
          return new TableRow({ children: [
            dataCell(p, 15, { alignment: AlignmentType.CENTER, color: pc.text, bold: true, fill: pc.bg }),
            dataCell(row[1], 15, { alignment: AlignmentType.CENTER }),
            dataCell(row[2], 25, { alignment: AlignmentType.CENTER }),
            dataCell(row[3], 45),
          ]});
        }),
      ],
    }),
  ];
  return children;
}

// --- CHAPTER 2: PORTFOLIO OVERVIEW ---
function buildChapter2() {
  const children = [
    new Paragraph({ children: [new PageBreak()] }),
    heading1("2. ポートフォリオ概況分析"),

    heading2("2.1 全ホテルスコアランキング"),
    para(`${meta.total_hotels || 19}ホテルの口コミスコアを10点換算で比較しました。ポートフォリオ平均は${overview.avg_score}点、中央値は${overview.median_score}点です。`),
    spacer(80),

    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({ children: [
          headerCell("#", 5),
          headerCell("ホテル名", 30),
          headerCell("平均", 10),
          headerCell("高評価率", 12),
          headerCell("低評価率", 12),
          headerCell("件数", 10),
          headerCell("ティア", 10),
          headerCell("優先度", 11),
        ]}),
        ...hotelsRanked.map((h, idx) => {
          const fill = idx % 2 === 1 ? C.LIGHT_BG : undefined;
          const pc = PRIORITY_COLORS[h.priority] || PRIORITY_COLORS.STANDARD;
          return new TableRow({ children: [
            dataCell(String(h.rank), 5, { alignment: AlignmentType.CENTER, fill }),
            dataCell(h.name, 30, { bold: true, fill }),
            dataCell(h.avg.toFixed(2), 10, { alignment: AlignmentType.CENTER, bold: true, fill }),
            dataCell(h.high_rate.toFixed(1) + "%", 12, { alignment: AlignmentType.CENTER, fill }),
            dataCell(h.low_rate.toFixed(1) + "%", 12, { alignment: AlignmentType.CENTER, color: h.low_rate >= 10 ? C.RED : C.TEXT, fill }),
            dataCell(String(h.total_reviews), 10, { alignment: AlignmentType.CENTER, fill }),
            dataCell(h.tier, 10, { alignment: AlignmentType.CENTER, fill }),
            dataCell(h.priority, 11, { alignment: AlignmentType.CENTER, color: pc.text, bold: true, fill: pc.bg }),
          ]});
        }),
      ],
    }),
    spacer(200),

    heading2("2.2 スコア分布分析"),
    para(`ポートフォリオ内のスコア分布は以下の通りです。`),
    bulletItem(`優秀（9.0以上）：${hotelsRanked.filter(h => h.avg >= 9.0).length}ホテル - 高い顧客満足度を維持`),
    bulletItem(`良好（8.5-8.99）：${hotelsRanked.filter(h => h.avg >= 8.5 && h.avg < 9.0).length}ホテル - 安定した運営`),
    bulletItem(`概ね良好（8.0-8.49）：${hotelsRanked.filter(h => h.avg >= 8.0 && h.avg < 8.5).length}ホテル - 改善余地あり`),
    bulletItem(`要改善（8.0未満）：${hotelsRanked.filter(h => h.avg < 8.0).length}ホテル - 早急な対策が必要`),
    spacer(100),

    heading2("2.3 ポートフォリオ全体の特徴"),
    para(`最高スコアは${overview.best_hotel ? overview.best_hotel.name + "（" + overview.best_hotel.avg + "点）" : "N/A"}、最低スコアは${overview.worst_hotel ? overview.worst_hotel.name + "（" + overview.worst_hotel.avg + "点）" : "N/A"}です。スコアレンジは${overview.best_hotel && overview.worst_hotel ? (overview.best_hotel.avg - overview.worst_hotel.avg).toFixed(2) : "N/A"}点と大きく、ホテル間の品質格差が見られます。清掃品質の標準化が、ポートフォリオ全体の底上げに重要です。`),
  ];
  return children;
}

// --- CHAPTER 3: CLEANING DEEP DIVE ---
function buildChapter3() {
  const totalMentions = deepDive.total_cleaning_mentions || 73;
  const issueRate = deepDive.portfolio_cleaning_issue_rate || 4.6;

  // Top 6 categories for matrix
  const topCategories = categorySummary.slice(0, 6).map(c => c.category);

  // Only hotels with cleaning issues
  const matrixHotels = cleaningMatrix.filter(h => h.cleaning_issue_count > 0);

  const children = [
    new Paragraph({ children: [new PageBreak()] }),
    heading1("3. 清掃品質 Deep Dive"),

    heading2("3.1 清掃関連クレーム概要"),
    para(`分析対象${meta.total_reviews || 1583}件のレビューのうち、清掃に関するネガティブな言及は合計${totalMentions}件（クレーム率${issueRate}%）でした。業界のベンチマーク（2-3%）を超過しており、組織的な改善が必要です。`),
    spacer(100),

    heading2("3.2 カテゴリ別分析"),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({ children: [
          headerCell("カテゴリ", 30),
          headerCell("言及数", 15),
          headerCell("該当ホテル数", 20),
          headerCell("深刻度", 15),
          headerCell("割合", 20),
        ]}),
        ...categorySummary.map((cat, idx) => {
          const fill = idx % 2 === 1 ? C.LIGHT_BG : undefined;
          const sevColor = cat.severity === "CRITICAL" ? C.RED : cat.severity === "HIGH" ? C.ORANGE : cat.severity === "MEDIUM" ? C.YELLOW : C.SUBTEXT;
          const pct = totalMentions > 0 ? ((cat.total_mentions / totalMentions) * 100).toFixed(1) + "%" : "-";
          return new TableRow({ children: [
            dataCell(cat.category, 30, { bold: true, fill }),
            dataCell(String(cat.total_mentions), 15, { alignment: AlignmentType.CENTER, bold: true, fill }),
            dataCell(cat.hotels_affected + "/" + (meta.total_hotels || 19), 20, { alignment: AlignmentType.CENTER, fill }),
            dataCell(cat.severity, 15, { alignment: AlignmentType.CENTER, color: sevColor, bold: true, fill }),
            dataCell(pct, 20, { alignment: AlignmentType.CENTER, fill }),
          ]});
        }),
      ],
    }),
    spacer(200),

    heading2("3.3 ホテル別清掃課題マトリクス"),
    para("清掃関連クレームのあるホテルについて、カテゴリ別の件数を示します。セルの色は件数の多さを表します。"),
    spacer(80),
  ];

  // Build matrix table
  if (matrixHotels.length > 0 && topCategories.length > 0) {
    const catColWidth = Math.floor(55 / topCategories.length);
    const headerRow = new TableRow({ children: [
      headerCell("ホテル名", 25),
      headerCell("合計", 8),
      headerCell("率(%)", 12),
      ...topCategories.map(cat => headerCell(cat.length > 4 ? cat.substring(0, 4) : cat, catColWidth)),
    ]});

    const dataRows = matrixHotels.map((h, idx) => {
      const fill = idx % 2 === 1 ? C.LIGHT_BG : undefined;
      return new TableRow({ children: [
        dataCell(h.name, 25, { bold: true, fill }),
        dataCell(String(h.cleaning_issue_count), 8, { alignment: AlignmentType.CENTER, bold: true, fill }),
        dataCell(h.cleaning_issue_rate.toFixed(1), 12, { alignment: AlignmentType.CENTER, color: h.cleaning_issue_rate >= 10 ? C.RED : h.cleaning_issue_rate >= 5 ? C.ORANGE : C.TEXT, fill }),
        ...topCategories.map(cat => {
          const count = (h.categories || {})[cat] || 0;
          let cellFill = fill;
          if (count >= 6) cellFill = "FFCDD2";
          else if (count >= 3) cellFill = "FFE0B2";
          else if (count >= 1) cellFill = "FFF9C4";
          return dataCell(count > 0 ? String(count) : "-", catColWidth, { alignment: AlignmentType.CENTER, fill: cellFill, bold: count >= 3 });
        }),
      ]});
    });

    children.push(new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [headerRow, ...dataRows],
    }));
  }

  return children;
}

// --- CHAPTER 4: PRIORITY MATRIX & ACTION PLANS ---
function buildChapter4() {
  const children = [
    new Paragraph({ children: [new PageBreak()] }),
    heading1("4. 優先度マトリクスとアクションプラン"),

    heading2("4.1 優先度分類基準"),
    para("各ホテルは、口コミスコア・清掃クレーム率・低評価率を総合的に評価し、以下の4段階に分類しています。"),
    bulletItem("URGENT（緊急）：平均スコア8.0未満、または清掃クレーム率10%超。即時対応が必要"),
    bulletItem("HIGH（高）：清掃クレーム率5%超、または特定カテゴリで顕著な問題あり"),
    bulletItem("STANDARD（標準）：清掃クレームが散発的。定期的な改善で対応可能"),
    bulletItem("MAINTENANCE（維持）：清掃関連クレームなし。現状維持で良好"),
    spacer(200),
  ];

  // 4.2 URGENT hotels detailed
  const urgentPlans = actionPlans.filter(p => p.priority_level === "URGENT");
  if (urgentPlans.length > 0) {
    children.push(heading2("4.2 URGENTホテル詳細アクションプラン"));
    urgentPlans.forEach(plan => {
      children.push(new Paragraph({ children: [new PageBreak()] }));
      children.push(heading3(`【URGENT】${plan.hotel}`));
      children.push(para(`現状スコア：${plan.current_avg}点 → 目標スコア：${plan.target_avg}点`));

      // Find hotel in matrix for issues
      const matrixEntry = cleaningMatrix.find(h => h.name === plan.hotel);
      if (matrixEntry) {
        const topIssues = Object.entries(matrixEntry.categories || {}).sort((a, b) => b[1] - a[1]).slice(0, 3);
        if (topIssues.length > 0) {
          children.push(para(`主要課題：${topIssues.map(([k, v]) => k + "(" + v + "件)").join("、")}`));
        }
      }
      spacer(100);

      // Phase 1
      if (plan.phase1_immediate) {
        children.push(heading3(`Phase 1：即時対応（${plan.phase1_immediate.timeline}）`));
        (plan.phase1_immediate.actions || []).slice(0, 6).forEach(a => {
          children.push(bulletItem(`${a.action}（${a.category}）`));
        });
      }
      children.push(spacer(100));

      // Phase 2
      if (plan.phase2_short_term) {
        children.push(heading3(`Phase 2：短期施策（${plan.phase2_short_term.timeline}）`));
        (plan.phase2_short_term.actions || []).slice(0, 5).forEach(a => {
          children.push(bulletItem(`${a.action}（${a.category}）`));
        });
      }
      children.push(spacer(100));

      // Phase 3
      if (plan.phase3_medium_term) {
        children.push(heading3(`Phase 3：中期施策（${plan.phase3_medium_term.timeline}）`));
        (plan.phase3_medium_term.actions || []).slice(0, 4).forEach(a => {
          children.push(bulletItem(`${a.action}（${a.category}）`));
        });
      }
    });
  }

  // 4.3 HIGH hotels summary
  const highPlans = actionPlans.filter(p => p.priority_level === "HIGH");
  if (highPlans.length > 0) {
    children.push(new Paragraph({ children: [new PageBreak()] }));
    children.push(heading2("4.3 HIGHホテル改善サマリー"));
    highPlans.forEach(plan => {
      children.push(heading3(`【HIGH】${plan.hotel}（${plan.current_avg}点 → ${plan.target_avg}点）`));
      const matrixEntry = cleaningMatrix.find(h => h.name === plan.hotel);
      if (matrixEntry) {
        const topIssues = Object.entries(matrixEntry.categories || {}).sort((a, b) => b[1] - a[1]).slice(0, 3);
        children.push(para(`主要課題：${topIssues.map(([k, v]) => k + "(" + v + "件)").join("、")}`));
      }
      if (plan.phase1_immediate) {
        children.push(para("即時対応策："));
        (plan.phase1_immediate.actions || []).slice(0, 4).forEach(a => {
          children.push(bulletItem(a.action));
        });
      }
      children.push(spacer(120));
    });
  }

  // 4.4 STANDARD/MAINTENANCE brief
  children.push(heading2("4.4 STANDARD・MAINTENANCEホテル"));
  const stdHotels = (priorityMatrix.standard || []).map(h => h.hotel).join("、");
  const maintHotels = (priorityMatrix.maintenance || []).map(h => h.hotel).join("、");
  children.push(para(`STANDARDホテル（${(priorityMatrix.standard || []).length}件）：${stdHotels}`));
  children.push(para("これらのホテルは散発的な清掃クレームが見られますが、定期的な品質管理の徹底で対応可能です。"));
  if (maintHotels) {
    children.push(para(`MAINTENANCEホテル（${(priorityMatrix.maintenance || []).length}件）：${maintHotels}`));
    children.push(para("清掃関連クレームがなく、現状の品質水準を維持することが推奨されます。"));
  }

  return children;
}

// --- CHAPTER 5: CROSS-CUTTING RECOMMENDATIONS ---
function buildChapter5() {
  const children = [
    new Paragraph({ children: [new PageBreak()] }),
    heading1("5. 全社横断施策提案"),
    para("個別ホテルの改善に加え、ポートフォリオ全体の品質底上げのため、以下の横断的施策を提案いたします。"),
    spacer(100),
  ];

  crossCutting.forEach((rec, idx) => {
    children.push(heading2(`5.${idx + 1} ${rec.theme}`));
    children.push(multiRunPara([
      { text: "対象：", bold: true },
      { text: rec.applicable_hotels || "全ホテル" },
      { text: "　優先度：", bold: true },
      { text: rec.priority || "" },
    ]));
    children.push(para(rec.description || ""));
    (rec.items || []).forEach(item => {
      children.push(bulletItem(item));
    });
    children.push(spacer(120));
  });

  return children;
}

// --- CHAPTER 6: KPI FRAMEWORK ---
function buildChapter6() {
  const children = [
    new Paragraph({ children: [new PageBreak()] }),
    heading1("6. KPIフレームワーク"),

    heading2("6.1 ポートフォリオ全体KPI目標"),
    para("以下のKPIを設定し、四半期ごとにモニタリングすることを推奨します。"),
    spacer(80),

    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({ children: [
          headerCell("KPI項目", 30),
          headerCell("現状値", 20),
          headerCell("目標値", 25),
          headerCell("期限", 25),
        ]}),
        ...portfolioTargets.map((t, idx) => {
          const fill = idx % 2 === 1 ? C.LIGHT_BG : undefined;
          return new TableRow({ children: [
            dataCell(t.kpi, 30, { bold: true, fill }),
            dataCell(t.current, 20, { alignment: AlignmentType.CENTER, fill }),
            dataCell(t.target, 25, { alignment: AlignmentType.CENTER, color: C.GREEN, bold: true, fill }),
            dataCell(t.deadline, 25, { alignment: AlignmentType.CENTER, fill }),
          ]});
        }),
      ],
    }),
    spacer(200),

    heading2("6.2 ホテル別KPI目標（URGENT・HIGH）"),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({ children: [
          headerCell("ホテル名", 28),
          headerCell("優先度", 12),
          headerCell("現スコア", 12),
          headerCell("目標スコア", 12),
          headerCell("現クレーム率", 18),
          headerCell("目標クレーム率", 18),
        ]}),
        ...perHotelTargets.map((t, idx) => {
          const fill = idx % 2 === 1 ? C.LIGHT_BG : undefined;
          const pc = PRIORITY_COLORS[t.priority] || PRIORITY_COLORS.STANDARD;
          return new TableRow({ children: [
            dataCell(t.hotel, 28, { bold: true, fill }),
            dataCell(t.priority, 12, { alignment: AlignmentType.CENTER, color: pc.text, bold: true, fill: pc.bg }),
            dataCell(t.current_avg.toFixed(2), 12, { alignment: AlignmentType.CENTER, fill }),
            dataCell(t.target_avg.toFixed(2), 12, { alignment: AlignmentType.CENTER, color: C.GREEN, bold: true, fill }),
            dataCell(t.current_cleaning_rate.toFixed(1) + "%", 18, { alignment: AlignmentType.CENTER, fill }),
            dataCell(t.target_cleaning_rate.toFixed(1) + "%", 18, { alignment: AlignmentType.CENTER, color: C.GREEN, bold: true, fill }),
          ]});
        }),
      ],
    }),
  ];
  return children;
}

// --- CHAPTER 7: ROI ESTIMATION ---
function buildChapter7() {
  const children = [
    new Paragraph({ children: [new PageBreak()] }),
    heading1("7. ROI試算"),

    heading2("試算方法"),
    para(roiEstimation.methodology || "口コミスコア0.1pt改善 ≒ RevPAR約1%向上（業界ベンチマーク）"),
    spacer(100),

    heading2("3つのシナリオ"),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({ children: [
          headerCell("シナリオ", 25),
          headerCell("対象ホテル数", 12),
          headerCell("概算コスト", 18),
          headerCell("期待改善幅", 18),
          headerCell("収益インパクト", 15),
          headerCell("ROI回収", 12),
        ]}),
        ...scenarios.map((s, idx) => {
          const fill = idx % 2 === 1 ? C.LIGHT_BG : undefined;
          return new TableRow({ children: [
            dataCell(s.scenario, 25, { bold: true, fill }),
            dataCell(String(s.target_hotels), 12, { alignment: AlignmentType.CENTER, fill }),
            dataCell(s.estimated_cost, 18, { alignment: AlignmentType.CENTER, fill }),
            dataCell(s.expected_improvement, 18, { alignment: AlignmentType.CENTER, fill }),
            dataCell(s.revenue_impact, 15, { alignment: AlignmentType.CENTER, color: C.GREEN, bold: true, fill }),
            dataCell(s.roi_period, 12, { alignment: AlignmentType.CENTER, fill }),
          ]});
        }),
      ],
    }),
  ];
  return children;
}

// --- CHAPTER 8: CONCLUSION ---
function buildChapter8() {
  const urgentNames = (priorityMatrix.urgent || []).map(h => h.hotel);
  const children = [
    new Paragraph({ children: [new PageBreak()] }),
    heading1("8. 結論と次のステップ"),

    heading2("トップ3優先事項"),
    bulletItem(`URGENTホテル（${urgentNames.join("、")}）の清掃品質緊急改善。Phase 1対応を今週中に開始。`),
    bulletItem("全社統一の清掃チェックリスト（50項目）の導入と品質管理（QC）体制の構築。"),
    bulletItem("清掃関連KPIの月次モニタリング体制の確立と、四半期レビューの実施。"),
    spacer(200),

    heading2("次のステップ"),
    bulletItem("本レポートの内容について経営層との合意形成（1週間以内）"),
    bulletItem("URGENTホテルの現地視察と清掃責任者との改善ミーティング実施（2週間以内）"),
    bulletItem("清掃チェックリストの最終版作成と全ホテルへの展開（1ヶ月以内）"),
    bulletItem("改善効果の初回レビュー（3ヶ月後）"),
    spacer(300),

    divider(),
    spacer(100),

    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [new TableRow({
        children: [new TableCell({
          borders: { top: { style: BorderStyle.SINGLE, size: 4, color: C.NAVY }, bottom: { style: BorderStyle.SINGLE, size: 4, color: C.NAVY }, left: { style: BorderStyle.SINGLE, size: 4, color: C.NAVY }, right: { style: BorderStyle.SINGLE, size: 4, color: C.NAVY } },
          width: { size: 100, type: WidthType.PERCENTAGE },
          shading: { fill: "F8F9FA", type: ShadingType.CLEAR },
          margins: { top: 300, bottom: 300, left: 400, right: 400 },
          children: [
            new Paragraph({ spacing: { after: 160 }, children: [new TextRun({ text: `本レポートは、${meta.total_reviews || 1583}件のゲスト口コミデータに基づき、PRIMECHANGEポートフォリオ全${meta.total_hotels || 19}ホテルの清掃品質を分析したものです。データに基づく優先順位付けと段階的な改善施策により、ポートフォリオ全体の顧客満足度向上とRevPAR改善を実現することが可能です。`, size: 21, font: "Arial", color: C.TEXT })] }),
            new Paragraph({ children: [new TextRun({ text: "清掃品質の標準化は、ゲスト満足度向上の最も費用対効果の高い施策です。本提案の実行により、ポートフォリオ全体の競争力強化に寄与いたします。", size: 21, font: "Arial", color: C.NAVY, bold: true })] }),
          ],
        })],
      })],
    }),

    spacer(300),
    para("本レポートに関するご質問・ご相談がございましたら、お気軽にお問い合わせください。", { alignment: AlignmentType.CENTER, color: "999999" }),
  ];
  return children;
}

// ============================================================
// Build DOCX Document
// ============================================================
const doc = new Document({
  styles: {
    default: { document: { run: { font: "Arial", size: 21 } } },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: "Arial", color: C.NAVY },
        paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 0 },
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Arial", color: C.ACCENT },
        paragraph: { spacing: { before: 280, after: 160 }, outlineLevel: 1 },
      },
    ],
  },
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [
          { level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
          { level: 1, format: LevelFormat.BULLET, text: "\u25E6", alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 1440, hanging: 360 } } } },
        ],
      },
    ],
  },
  sections: [
    // Cover page
    buildCoverPage(),
    // Main content
    {
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
        },
      },
      headers: {
        default: new Header({
          children: [new Paragraph({
            alignment: AlignmentType.RIGHT,
            border: { bottom: { style: BorderStyle.SINGLE, size: 2, color: C.ACCENT, space: 4 } },
            children: [new TextRun({ text: "PRIMECHANGE｜ホテル清掃戦略レポート", size: 16, font: "Arial", color: "999999" })],
          })],
        }),
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({ text: "Confidential | Page ", size: 16, font: "Arial", color: "999999" }),
              new TextRun({ children: [PageNumber.CURRENT], size: 16, font: "Arial", color: "999999" }),
            ],
          })],
        }),
      },
      children: [
        ...buildChapter1(),
        ...buildChapter2(),
        ...buildChapter3(),
        ...buildChapter4(),
        ...buildChapter5(),
        ...buildChapter6(),
        ...buildChapter7(),
        ...buildChapter8(),
      ],
    },
  ],
});


// ============================================================
// PPTX Presentation
// ============================================================
const pptx = new pptxgen();
pptx.layout = "LAYOUT_16x9";
pptx.author = "PRIMECHANGE Consulting";
pptx.title = "PRIMECHANGE ホテル清掃戦略レポート";

const shadow = () => ({ type: "outer", blur: 4, offset: 2, angle: 135, color: "000000", opacity: 0.10 });

function addFooter(slide, pageNum) {
  slide.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.NAVY } });
  slide.addText("Confidential | PRIMECHANGE", { x: 0.5, y: 5.25, w: 4, h: 0.375, fontSize: 8, color: "94A3B8", fontFace: "Arial", valign: "middle" });
  slide.addText(String(pageNum), { x: 9, y: 5.25, w: 0.5, h: 0.375, fontSize: 8, color: "94A3B8", fontFace: "Arial", align: "right", valign: "middle" });
}

function addContentHeader(slide, title, subtitle) {
  slide.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.6, fill: { color: C.NAVY } });
  slide.addText(title, { x: 0.5, y: 0.05, w: 9, h: 0.35, fontSize: 20, fontFace: "Arial", color: C.WHITE, bold: true, margin: 0 });
  if (subtitle) {
    slide.addText(subtitle, { x: 0.5, y: 0.35, w: 9, h: 0.22, fontSize: 10, fontFace: "Arial", color: "5A9BE6", margin: 0 });
  }
}

function pptxKpiCard(slide, x, y, w, h, label, value, color, bgColor) {
  slide.addShape(pptx.shapes.RECTANGLE, { x, y, w, h, fill: { color: bgColor }, shadow: shadow() });
  slide.addShape(pptx.shapes.RECTANGLE, { x, y, w, h: 0.04, fill: { color } });
  slide.addText(label, { x, y: y + 0.12, w, h: 0.22, fontSize: 9, fontFace: "Arial", color: C.SUBTEXT, align: "center", margin: 0 });
  slide.addText(value, { x, y: y + 0.32, w, h: 0.45, fontSize: 28, fontFace: "Arial", color, bold: true, align: "center", margin: 0 });
}

// --- Slide 1: Title ---
(function buildSlide1() {
  const slide = pptx.addSlide();
  slide.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: C.NAVY } });
  slide.addText("PRIMECHANGE", { x: 0.8, y: 1.0, w: 8.4, h: 0.7, fontSize: 20, fontFace: "Arial", color: "5A9BE6", margin: 0 });
  slide.addText("ホテル清掃戦略レポート", { x: 0.8, y: 1.7, w: 8.4, h: 1.0, fontSize: 36, fontFace: "Arial", color: C.WHITE, bold: true, margin: 0 });
  slide.addShape(pptx.shapes.RECTANGLE, { x: 0.8, y: 2.8, w: 2.0, h: 0.04, fill: { color: "5A9BE6" } });
  slide.addText(meta.subtitle || "ゲスト口コミデータに基づく清掃品質改善提案書", { x: 0.8, y: 3.0, w: 8.4, h: 0.4, fontSize: 14, fontFace: "Arial", color: "94A3B8", margin: 0 });
  slide.addText(meta.date || "2026年3月8日", { x: 0.8, y: 3.6, w: 4, h: 0.3, fontSize: 12, fontFace: "Arial", color: "64748B", margin: 0 });
  slide.addText(`対象：${meta.total_hotels || 19}ホテル | レビュー：${meta.total_reviews || 1583}件`, { x: 0.8, y: 3.9, w: 6, h: 0.3, fontSize: 11, fontFace: "Arial", color: "64748B", margin: 0 });
  slide.addText(`提出先：${meta.prepared_for || ""}`, { x: 0.8, y: 4.5, w: 8, h: 0.25, fontSize: 10, fontFace: "Arial", color: "64748B", margin: 0 });
  slide.addText("Confidential", { x: 0.8, y: 4.9, w: 3, h: 0.25, fontSize: 9, fontFace: "Arial", color: "475569", italics: true, margin: 0 });
})();

// --- Slide 2: Executive Summary ---
(function buildSlide2() {
  const slide = pptx.addSlide();
  addContentHeader(slide, "エグゼクティブサマリー", "Executive Summary");

  pptxKpiCard(slide, 0.3, 0.8, 2.1, 0.9, "ポートフォリオ平均", String(overview.avg_score || 8.39), C.ACCENT, "E3F2FD");
  pptxKpiCard(slide, 2.65, 0.8, 2.1, 0.9, "高評価率", (overview.portfolio_high_rate || 78.1) + "%", C.GREEN, "E8F5E9");
  pptxKpiCard(slide, 5.0, 0.8, 2.1, 0.9, "清掃クレーム率", (deepDive.portfolio_cleaning_issue_rate || 4.6) + "%", C.RED, "FFEBEE");
  pptxKpiCard(slide, 7.35, 0.8, 2.1, 0.9, "レビュー総数", (meta.total_reviews || 1583) + "件", C.NAVY, C.LIGHT_BG);

  // Key findings box
  slide.addShape(pptx.shapes.RECTANGLE, { x: 0.3, y: 2.0, w: 9.1, h: 2.8, fill: { color: "F0F7FC" }, line: { color: C.ACCENT, width: 1.5 }, shadow: shadow() });
  slide.addText("KEY FINDINGS", { x: 0.5, y: 2.1, w: 3, h: 0.3, fontSize: 14, fontFace: "Arial", color: C.ACCENT, bold: true, margin: 0 });

  const findings = [
    [{ text: "Strength: ", options: { fontSize: 10, fontFace: "Arial", color: C.GREEN, bold: true } },
     { text: `上位5ホテル平均8.7点超。${overview.best_hotel ? overview.best_hotel.name + "が" + overview.best_hotel.avg + "点で最高評価" : ""}`, options: { fontSize: 10, fontFace: "Arial", color: C.TEXT } }],
    [{ text: "Challenge: ", options: { fontSize: 10, fontFace: "Arial", color: C.RED, bold: true } },
     { text: `${(priorityMatrix.urgent || []).length}ホテルがURGENT。清掃クレーム率${deepDive.portfolio_cleaning_issue_rate || 4.6}%（業界平均2-3%超過）`, options: { fontSize: 10, fontFace: "Arial", color: C.TEXT } }],
    [{ text: "Opportunity: ", options: { fontSize: 10, fontFace: "Arial", color: C.ORANGE, bold: true } },
     { text: "清掃品質標準化によりURGENTホテル+0.5-1.0pt改善、RevPAR 5-10%向上見込み", options: { fontSize: 10, fontFace: "Arial", color: C.TEXT } }],
  ];
  findings.forEach((f, i) => {
    slide.addText(f, { x: 0.5, y: 2.5 + i * 0.7, w: 8.8, h: 0.6, valign: "top", margin: 0 });
  });

  addFooter(slide, 2);
})();

// --- Slide 3: Portfolio Overview ---
(function buildSlide3() {
  const slide = pptx.addSlide();
  addContentHeader(slide, "ポートフォリオ概況", "全19ホテルスコアランキング");

  const headerRow = [
    { text: "#", options: { fontSize: 7, bold: true, color: C.WHITE, fill: { color: C.NAVY }, align: "center" } },
    { text: "ホテル名", options: { fontSize: 7, bold: true, color: C.WHITE, fill: { color: C.NAVY }, align: "left" } },
    { text: "平均", options: { fontSize: 7, bold: true, color: C.WHITE, fill: { color: C.NAVY }, align: "center" } },
    { text: "高評価率", options: { fontSize: 7, bold: true, color: C.WHITE, fill: { color: C.NAVY }, align: "center" } },
    { text: "優先度", options: { fontSize: 7, bold: true, color: C.WHITE, fill: { color: C.NAVY }, align: "center" } },
    { text: "ティア", options: { fontSize: 7, bold: true, color: C.WHITE, fill: { color: C.NAVY }, align: "center" } },
  ];

  const rows = [headerRow];
  hotelsRanked.forEach((h, idx) => {
    const pc = PRIORITY_COLORS[h.priority] || PRIORITY_COLORS.STANDARD;
    const fillColor = idx % 2 === 1 ? C.LIGHT_BG : C.WHITE;
    rows.push([
      { text: String(h.rank), options: { fontSize: 7, align: "center", fill: { color: fillColor } } },
      { text: h.name, options: { fontSize: 7, bold: true, fill: { color: fillColor } } },
      { text: h.avg.toFixed(2), options: { fontSize: 7, bold: true, align: "center", fill: { color: fillColor } } },
      { text: h.high_rate.toFixed(1) + "%", options: { fontSize: 7, align: "center", fill: { color: fillColor } } },
      { text: h.priority, options: { fontSize: 7, bold: true, color: pc.text, align: "center", fill: { color: pc.bg } } },
      { text: h.tier, options: { fontSize: 7, align: "center", fill: { color: fillColor } } },
    ]);
  });

  slide.addTable(rows, { x: 0.2, y: 0.75, w: 9.6, colW: [0.4, 3.4, 0.8, 1.0, 1.2, 0.8], fontSize: 7, border: { pt: 0.5, color: "CCCCCC" }, autoPage: false });
  addFooter(slide, 3);
})();

// --- Slide 4: Cleaning Overview ---
(function buildSlide4() {
  const slide = pptx.addSlide();
  addContentHeader(slide, "清掃品質 Deep Dive", "カテゴリ別清掃クレーム分析");

  // Stats cards
  pptxKpiCard(slide, 0.3, 0.8, 2.8, 0.85, "清掃クレーム総数", String(deepDive.total_cleaning_mentions || 73) + "件", C.RED, "FFEBEE");
  pptxKpiCard(slide, 3.4, 0.8, 2.8, 0.85, "クレーム率", (deepDive.portfolio_cleaning_issue_rate || 4.6) + "%", C.ORANGE, "FFF3E0");
  pptxKpiCard(slide, 6.5, 0.8, 2.8, 0.85, "カテゴリ数", String(categorySummary.length), C.ACCENT, "E3F2FD");

  // Category table
  const headerRow = [
    { text: "カテゴリ", options: { fontSize: 8, bold: true, color: C.WHITE, fill: { color: C.NAVY } } },
    { text: "言及数", options: { fontSize: 8, bold: true, color: C.WHITE, fill: { color: C.NAVY }, align: "center" } },
    { text: "該当ホテル", options: { fontSize: 8, bold: true, color: C.WHITE, fill: { color: C.NAVY }, align: "center" } },
    { text: "深刻度", options: { fontSize: 8, bold: true, color: C.WHITE, fill: { color: C.NAVY }, align: "center" } },
  ];

  const rows = [headerRow];
  categorySummary.forEach((cat, idx) => {
    const fillColor = idx % 2 === 1 ? C.LIGHT_BG : C.WHITE;
    const sevColor = cat.severity === "CRITICAL" ? C.RED : cat.severity === "HIGH" ? C.ORANGE : C.SUBTEXT;
    rows.push([
      { text: cat.category, options: { fontSize: 8, bold: true, fill: { color: fillColor } } },
      { text: String(cat.total_mentions), options: { fontSize: 8, bold: true, align: "center", fill: { color: fillColor } } },
      { text: cat.hotels_affected + "/" + (meta.total_hotels || 19), options: { fontSize: 8, align: "center", fill: { color: fillColor } } },
      { text: cat.severity, options: { fontSize: 8, bold: true, color: sevColor, align: "center", fill: { color: fillColor } } },
    ]);
  });

  slide.addTable(rows, { x: 0.3, y: 2.0, w: 9.4, colW: [3.5, 1.5, 1.8, 1.5], fontSize: 8, border: { pt: 0.5, color: "CCCCCC" }, autoPage: false });
  addFooter(slide, 4);
})();

// --- Slide 5: Hotel x Cleaning Matrix ---
(function buildSlide5() {
  const slide = pptx.addSlide();
  addContentHeader(slide, "ホテル別清掃課題マトリクス", "ヒートマップ：色が濃いほど件数が多い");

  const topCats = categorySummary.slice(0, 6).map(c => c.category);
  const matrixHotels = cleaningMatrix.filter(h => h.cleaning_issue_count > 0 && h.priority !== "MAINTENANCE");

  const headerRow = [
    { text: "ホテル名", options: { fontSize: 7, bold: true, color: C.WHITE, fill: { color: C.NAVY } } },
    { text: "合計", options: { fontSize: 7, bold: true, color: C.WHITE, fill: { color: C.NAVY }, align: "center" } },
    ...topCats.map(cat => ({ text: cat.length > 5 ? cat.substring(0, 5) : cat, options: { fontSize: 6, bold: true, color: C.WHITE, fill: { color: C.NAVY }, align: "center" } })),
  ];

  const rows = [headerRow];
  matrixHotels.forEach(h => {
    const cells = [
      { text: h.name, options: { fontSize: 7, bold: true } },
      { text: String(h.cleaning_issue_count), options: { fontSize: 7, bold: true, align: "center" } },
    ];
    topCats.forEach(cat => {
      const count = (h.categories || {})[cat] || 0;
      let bgColor = C.WHITE;
      if (count >= 6) bgColor = "FFCDD2";
      else if (count >= 3) bgColor = "FFE0B2";
      else if (count >= 1) bgColor = "FFF9C4";
      cells.push({ text: count > 0 ? String(count) : "-", options: { fontSize: 7, align: "center", fill: { color: bgColor }, color: count >= 3 ? C.RED : C.TEXT } });
    });
    rows.push(cells);
  });

  const catColW = topCats.map(() => (9.4 - 2.8 - 0.6) / topCats.length);
  slide.addTable(rows, { x: 0.3, y: 0.8, w: 9.4, colW: [2.8, 0.6, ...catColW], fontSize: 7, border: { pt: 0.5, color: "CCCCCC" }, autoPage: false });

  // Legend
  slide.addText([
    { text: "凡例: ", options: { fontSize: 8, bold: true } },
    { text: " 1-2件 ", options: { fontSize: 8, fill: { color: "FFF9C4" } } },
    { text: "  ", options: { fontSize: 8 } },
    { text: " 3-5件 ", options: { fontSize: 8, fill: { color: "FFE0B2" } } },
    { text: "  ", options: { fontSize: 8 } },
    { text: " 6件以上 ", options: { fontSize: 8, fill: { color: "FFCDD2" } } },
  ], { x: 0.3, y: 4.8, w: 5, h: 0.3, margin: 0 });

  addFooter(slide, 5);
})();

// --- Slide 6: Priority Matrix ---
(function buildSlide6() {
  const slide = pptx.addSlide();
  addContentHeader(slide, "優先度マトリクス", "4象限分類");

  const quadrants = [
    { label: "URGENT（緊急）", hotels: priorityMatrix.urgent || [], color: C.RED, bg: "FFEBEE", x: 0.3, y: 0.8, w: 4.5, h: 2.1 },
    { label: "HIGH（高）", hotels: priorityMatrix.high || [], color: C.ORANGE, bg: "FFF3E0", x: 5.1, y: 0.8, w: 4.5, h: 2.1 },
    { label: "STANDARD（標準）", hotels: priorityMatrix.standard || [], color: C.BLUE, bg: "E3F2FD", x: 0.3, y: 3.1, w: 4.5, h: 1.8 },
    { label: "MAINTENANCE（維持）", hotels: priorityMatrix.maintenance || [], color: C.GREEN, bg: "E8F5E9", x: 5.1, y: 3.1, w: 4.5, h: 1.8 },
  ];

  quadrants.forEach(q => {
    slide.addShape(pptx.shapes.RECTANGLE, { x: q.x, y: q.y, w: q.w, h: q.h, fill: { color: q.bg }, shadow: shadow() });
    slide.addShape(pptx.shapes.RECTANGLE, { x: q.x, y: q.y, w: q.w, h: 0.04, fill: { color: q.color } });
    slide.addText(q.label + ` (${q.hotels.length})`, { x: q.x + 0.15, y: q.y + 0.1, w: q.w - 0.3, h: 0.3, fontSize: 11, fontFace: "Arial", color: q.color, bold: true, margin: 0 });

    const hotelTexts = q.hotels.map(h => `${h.hotel}（${h.avg}点）`);
    slide.addText(hotelTexts.join("\n"), { x: q.x + 0.15, y: q.y + 0.45, w: q.w - 0.3, h: q.h - 0.6, fontSize: 8, fontFace: "Arial", color: C.TEXT, valign: "top", margin: 0 });
  });

  addFooter(slide, 6);
})();

// --- Slides 7-9: URGENT Hotel Deep Dives ---
(function buildSlides7to9() {
  const urgentPlans = actionPlans.filter(p => p.priority_level === "URGENT");
  urgentPlans.forEach((plan, idx) => {
    const slide = pptx.addSlide();
    addContentHeader(slide, `【URGENT】${plan.hotel}`, `優先度 #${plan.priority_rank || idx + 1} | 現状 ${plan.current_avg}点 → 目標 ${plan.target_avg}点`);

    // Current state
    pptxKpiCard(slide, 0.3, 0.8, 2.0, 0.8, "現状スコア", String(plan.current_avg), C.RED, "FFEBEE");
    pptxKpiCard(slide, 2.5, 0.8, 2.0, 0.8, "目標スコア", String(plan.target_avg), C.GREEN, "E8F5E9");

    // Top issues
    const matrixEntry = cleaningMatrix.find(h => h.name === plan.hotel);
    let issueText = "";
    if (matrixEntry) {
      const topIssues = Object.entries(matrixEntry.categories || {}).sort((a, b) => b[1] - a[1]).slice(0, 4);
      issueText = topIssues.map(([k, v]) => `${k}: ${v}件`).join("\n");
    }
    slide.addShape(pptx.shapes.RECTANGLE, { x: 4.8, y: 0.8, w: 4.8, h: 0.8, fill: { color: C.LIGHT_BG }, shadow: shadow() });
    slide.addText("主要課題", { x: 4.9, y: 0.82, w: 2, h: 0.2, fontSize: 9, fontFace: "Arial", color: C.SUBTEXT, bold: true, margin: 0 });
    slide.addText(issueText, { x: 4.9, y: 1.0, w: 4.6, h: 0.55, fontSize: 8, fontFace: "Arial", color: C.TEXT, margin: 0 });

    // Action phases
    const phases = [
      { label: "Phase 1: 即時対応", timeline: plan.phase1_immediate?.timeline || "1-2週間", actions: plan.phase1_immediate?.actions || [], color: C.RED, bg: "FFEBEE" },
      { label: "Phase 2: 短期施策", timeline: plan.phase2_short_term?.timeline || "1-3ヶ月", actions: plan.phase2_short_term?.actions || [], color: C.ORANGE, bg: "FFF3E0" },
      { label: "Phase 3: 中期施策", timeline: plan.phase3_medium_term?.timeline || "3-6ヶ月", actions: plan.phase3_medium_term?.actions || [], color: C.BLUE, bg: "E3F2FD" },
    ];

    phases.forEach((phase, pi) => {
      const px = 0.3 + pi * 3.15;
      const py = 1.85;
      slide.addShape(pptx.shapes.RECTANGLE, { x: px, y: py, w: 3.0, h: 2.95, fill: { color: phase.bg }, shadow: shadow() });
      slide.addShape(pptx.shapes.RECTANGLE, { x: px, y: py, w: 3.0, h: 0.04, fill: { color: phase.color } });
      slide.addText(`${phase.label}（${phase.timeline}）`, { x: px + 0.1, y: py + 0.08, w: 2.8, h: 0.25, fontSize: 9, fontFace: "Arial", color: phase.color, bold: true, margin: 0 });

      const actionTexts = phase.actions.slice(0, 5).map(a => `- ${a.action}`).join("\n");
      slide.addText(actionTexts, { x: px + 0.1, y: py + 0.38, w: 2.8, h: 2.5, fontSize: 7, fontFace: "Arial", color: C.TEXT, valign: "top", margin: 0 });
    });

    addFooter(slide, 7 + idx);
  });
})();

// --- Slides 10-11: HIGH Hotels ---
(function buildSlides10to11() {
  const highPlans = actionPlans.filter(p => p.priority_level === "HIGH");

  // Slide 10: HIGH Hotels Overview
  const slide10 = pptx.addSlide();
  addContentHeader(slide10, "HIGHホテル改善サマリー", `対象：${highPlans.length}ホテル`);

  highPlans.forEach((plan, idx) => {
    const py = 0.8 + idx * 1.45;
    const matrixEntry = cleaningMatrix.find(h => h.name === plan.hotel);
    const topIssues = matrixEntry ? Object.entries(matrixEntry.categories || {}).sort((a, b) => b[1] - a[1]).slice(0, 3) : [];

    slide10.addShape(pptx.shapes.RECTANGLE, { x: 0.3, y: py, w: 9.4, h: 1.3, fill: { color: "FFF3E0" }, shadow: shadow() });
    slide10.addShape(pptx.shapes.RECTANGLE, { x: 0.3, y: py, w: 9.4, h: 0.04, fill: { color: C.ORANGE } });
    slide10.addText(`${plan.hotel}`, { x: 0.5, y: py + 0.08, w: 4, h: 0.25, fontSize: 12, fontFace: "Arial", color: C.NAVY, bold: true, margin: 0 });
    slide10.addText(`${plan.current_avg}点 → ${plan.target_avg}点`, { x: 5.0, y: py + 0.08, w: 3, h: 0.25, fontSize: 11, fontFace: "Arial", color: C.GREEN, bold: true, margin: 0 });
    slide10.addText(`主要課題：${topIssues.map(([k, v]) => k + "(" + v + "件)").join("、")}`, { x: 0.5, y: py + 0.35, w: 8, h: 0.2, fontSize: 9, fontFace: "Arial", color: C.SUBTEXT, margin: 0 });

    const keyActions = (plan.phase1_immediate?.actions || []).slice(0, 3).map(a => `- ${a.action}`).join("\n");
    slide10.addText(keyActions, { x: 0.5, y: py + 0.6, w: 8.5, h: 0.65, fontSize: 8, fontFace: "Arial", color: C.TEXT, valign: "top", margin: 0 });
  });

  addFooter(slide10, 10);

  // Slide 11: HIGH Hotels Action Details
  const slide11 = pptx.addSlide();
  addContentHeader(slide11, "HIGHホテル推奨アクション", "Phase別主要施策");

  const headerRow = [
    { text: "ホテル", options: { fontSize: 8, bold: true, color: C.WHITE, fill: { color: C.NAVY } } },
    { text: "Phase 1（即時）", options: { fontSize: 8, bold: true, color: C.WHITE, fill: { color: C.NAVY } } },
    { text: "Phase 2（短期）", options: { fontSize: 8, bold: true, color: C.WHITE, fill: { color: C.NAVY } } },
    { text: "Phase 3（中期）", options: { fontSize: 8, bold: true, color: C.WHITE, fill: { color: C.NAVY } } },
  ];

  const rows = [headerRow];
  highPlans.forEach((plan, idx) => {
    const fillColor = idx % 2 === 1 ? C.LIGHT_BG : C.WHITE;
    rows.push([
      { text: plan.hotel, options: { fontSize: 7, bold: true, fill: { color: fillColor } } },
      { text: (plan.phase1_immediate?.actions || []).slice(0, 2).map(a => a.action).join("\n"), options: { fontSize: 7, fill: { color: fillColor } } },
      { text: (plan.phase2_short_term?.actions || []).slice(0, 2).map(a => a.action).join("\n"), options: { fontSize: 7, fill: { color: fillColor } } },
      { text: (plan.phase3_medium_term?.actions || []).slice(0, 2).map(a => a.action).join("\n"), options: { fontSize: 7, fill: { color: fillColor } } },
    ]);
  });

  slide11.addTable(rows, { x: 0.3, y: 0.8, w: 9.4, colW: [2.0, 2.5, 2.5, 2.4], fontSize: 7, border: { pt: 0.5, color: "CCCCCC" }, autoPage: false });
  addFooter(slide11, 11);
})();

// --- Slides 12-14: Cross-Cutting Recommendations ---
(function buildSlides12to14() {
  const recsPerSlide = 2;
  for (let si = 0; si < 3; si++) {
    const slide = pptx.addSlide();
    const startIdx = si * recsPerSlide;
    const slideRecs = crossCutting.slice(startIdx, startIdx + recsPerSlide);
    if (slideRecs.length === 0) {
      addContentHeader(slide, "全社横断施策提案", "（続き）");
      addFooter(slide, 12 + si);
      continue;
    }

    addContentHeader(slide, "全社横断施策提案", `施策 ${startIdx + 1}${slideRecs.length > 1 ? "-" + (startIdx + slideRecs.length) : ""}`);

    slideRecs.forEach((rec, ri) => {
      const py = 0.8 + ri * 2.2;
      const bgColor = ri === 0 ? "F0F7FC" : C.LIGHT_BG;

      slide.addShape(pptx.shapes.RECTANGLE, { x: 0.3, y: py, w: 9.4, h: 2.0, fill: { color: bgColor }, shadow: shadow() });
      slide.addShape(pptx.shapes.RECTANGLE, { x: 0.3, y: py, w: 9.4, h: 0.04, fill: { color: C.ACCENT } });
      slide.addText(rec.theme, { x: 0.5, y: py + 0.08, w: 5, h: 0.3, fontSize: 14, fontFace: "Arial", color: C.NAVY, bold: true, margin: 0 });
      slide.addText(`対象：${rec.applicable_hotels || ""} | 優先度：${rec.priority || ""}`, { x: 5.5, y: py + 0.08, w: 4, h: 0.3, fontSize: 9, fontFace: "Arial", color: C.SUBTEXT, align: "right", margin: 0 });
      slide.addText(rec.description || "", { x: 0.5, y: py + 0.4, w: 9, h: 0.35, fontSize: 9, fontFace: "Arial", color: C.TEXT, margin: 0 });

      const items = (rec.items || []).slice(0, 5).map(item => `- ${item}`).join("\n");
      slide.addText(items, { x: 0.5, y: py + 0.8, w: 9, h: 1.1, fontSize: 8, fontFace: "Arial", color: C.TEXT, valign: "top", margin: 0 });
    });

    addFooter(slide, 12 + si);
  }
})();

// --- Slide 15: KPI Framework ---
(function buildSlide15() {
  const slide = pptx.addSlide();
  addContentHeader(slide, "KPIフレームワーク", "ポートフォリオ目標 & ホテル別目標");

  // Portfolio targets table
  slide.addText("ポートフォリオ全体KPI", { x: 0.3, y: 0.75, w: 4, h: 0.25, fontSize: 11, fontFace: "Arial", color: C.NAVY, bold: true, margin: 0 });

  const ptHeader = [
    { text: "KPI項目", options: { fontSize: 8, bold: true, color: C.WHITE, fill: { color: C.NAVY } } },
    { text: "現状値", options: { fontSize: 8, bold: true, color: C.WHITE, fill: { color: C.NAVY }, align: "center" } },
    { text: "目標値", options: { fontSize: 8, bold: true, color: C.WHITE, fill: { color: C.NAVY }, align: "center" } },
    { text: "期限", options: { fontSize: 8, bold: true, color: C.WHITE, fill: { color: C.NAVY }, align: "center" } },
  ];
  const ptRows = [ptHeader];
  portfolioTargets.forEach((t, idx) => {
    const fillColor = idx % 2 === 1 ? C.LIGHT_BG : C.WHITE;
    ptRows.push([
      { text: t.kpi, options: { fontSize: 8, bold: true, fill: { color: fillColor } } },
      { text: t.current, options: { fontSize: 8, align: "center", fill: { color: fillColor } } },
      { text: t.target, options: { fontSize: 8, bold: true, color: C.GREEN, align: "center", fill: { color: fillColor } } },
      { text: t.deadline, options: { fontSize: 8, align: "center", fill: { color: fillColor } } },
    ]);
  });
  slide.addTable(ptRows, { x: 0.3, y: 1.05, w: 9.4, colW: [3.0, 1.8, 2.0, 1.6], fontSize: 8, border: { pt: 0.5, color: "CCCCCC" }, autoPage: false });

  // Per-hotel targets
  slide.addText("ホテル別KPI目標（URGENT/HIGH）", { x: 0.3, y: 2.7, w: 5, h: 0.25, fontSize: 11, fontFace: "Arial", color: C.NAVY, bold: true, margin: 0 });

  const htHeader = [
    { text: "ホテル", options: { fontSize: 7, bold: true, color: C.WHITE, fill: { color: C.NAVY } } },
    { text: "優先度", options: { fontSize: 7, bold: true, color: C.WHITE, fill: { color: C.NAVY }, align: "center" } },
    { text: "現スコア", options: { fontSize: 7, bold: true, color: C.WHITE, fill: { color: C.NAVY }, align: "center" } },
    { text: "目標", options: { fontSize: 7, bold: true, color: C.WHITE, fill: { color: C.NAVY }, align: "center" } },
    { text: "現クレーム率", options: { fontSize: 7, bold: true, color: C.WHITE, fill: { color: C.NAVY }, align: "center" } },
    { text: "目標率", options: { fontSize: 7, bold: true, color: C.WHITE, fill: { color: C.NAVY }, align: "center" } },
  ];
  const htRows = [htHeader];
  perHotelTargets.forEach((t, idx) => {
    const fillColor = idx % 2 === 1 ? C.LIGHT_BG : C.WHITE;
    const pc = PRIORITY_COLORS[t.priority] || PRIORITY_COLORS.STANDARD;
    htRows.push([
      { text: t.hotel, options: { fontSize: 7, bold: true, fill: { color: fillColor } } },
      { text: t.priority, options: { fontSize: 7, bold: true, color: pc.text, align: "center", fill: { color: pc.bg } } },
      { text: t.current_avg.toFixed(2), options: { fontSize: 7, align: "center", fill: { color: fillColor } } },
      { text: t.target_avg.toFixed(2), options: { fontSize: 7, bold: true, color: C.GREEN, align: "center", fill: { color: fillColor } } },
      { text: t.current_cleaning_rate.toFixed(1) + "%", options: { fontSize: 7, align: "center", fill: { color: fillColor } } },
      { text: t.target_cleaning_rate.toFixed(1) + "%", options: { fontSize: 7, bold: true, color: C.GREEN, align: "center", fill: { color: fillColor } } },
    ]);
  });
  slide.addTable(htRows, { x: 0.3, y: 3.0, w: 9.4, colW: [2.6, 0.9, 1.0, 1.0, 1.3, 1.0], fontSize: 7, border: { pt: 0.5, color: "CCCCCC" }, autoPage: false });

  addFooter(slide, 15);
})();

// --- Slide 16: ROI Estimation ---
(function buildSlide16() {
  const slide = pptx.addSlide();
  addContentHeader(slide, "ROI試算", roiEstimation.methodology || "");

  const headerRow = [
    { text: "シナリオ", options: { fontSize: 8, bold: true, color: C.WHITE, fill: { color: C.NAVY } } },
    { text: "対象", options: { fontSize: 8, bold: true, color: C.WHITE, fill: { color: C.NAVY }, align: "center" } },
    { text: "概算コスト", options: { fontSize: 8, bold: true, color: C.WHITE, fill: { color: C.NAVY }, align: "center" } },
    { text: "期待改善", options: { fontSize: 8, bold: true, color: C.WHITE, fill: { color: C.NAVY }, align: "center" } },
    { text: "収益効果", options: { fontSize: 8, bold: true, color: C.WHITE, fill: { color: C.NAVY }, align: "center" } },
    { text: "ROI回収", options: { fontSize: 8, bold: true, color: C.WHITE, fill: { color: C.NAVY }, align: "center" } },
  ];

  const rows = [headerRow];
  scenarios.forEach((s, idx) => {
    const fillColor = idx % 2 === 1 ? C.LIGHT_BG : C.WHITE;
    rows.push([
      { text: s.scenario, options: { fontSize: 8, bold: true, fill: { color: fillColor } } },
      { text: String(s.target_hotels) + "ホテル", options: { fontSize: 8, align: "center", fill: { color: fillColor } } },
      { text: s.estimated_cost, options: { fontSize: 8, align: "center", fill: { color: fillColor } } },
      { text: s.expected_improvement, options: { fontSize: 8, align: "center", fill: { color: fillColor } } },
      { text: s.revenue_impact, options: { fontSize: 8, bold: true, color: C.GREEN, align: "center", fill: { color: fillColor } } },
      { text: s.roi_period, options: { fontSize: 8, align: "center", fill: { color: fillColor } } },
    ]);
  });

  slide.addTable(rows, { x: 0.3, y: 1.0, w: 9.4, colW: [2.5, 0.9, 1.5, 1.6, 1.6, 1.0], fontSize: 8, border: { pt: 0.5, color: "CCCCCC" }, autoPage: false });

  // Visual highlight
  slide.addShape(pptx.shapes.RECTANGLE, { x: 0.3, y: 3.0, w: 9.4, h: 1.5, fill: { color: "F0F7FC" }, shadow: shadow() });
  slide.addText("推奨：シナリオA（URGENTホテル集中改善）から開始", { x: 0.5, y: 3.1, w: 8.8, h: 0.35, fontSize: 14, fontFace: "Arial", color: C.NAVY, bold: true, margin: 0 });
  slide.addText("最も費用対効果が高く、短期間でROI回収が見込めます。\n成果を確認後、シナリオBへ拡大し、最終的に全社品質管理システム（シナリオC）へ移行することを推奨いたします。", { x: 0.5, y: 3.5, w: 8.8, h: 0.9, fontSize: 10, fontFace: "Arial", color: C.TEXT, margin: 0 });

  addFooter(slide, 16);
})();

// --- Slide 17: Implementation Roadmap ---
(function buildSlide17() {
  const slide = pptx.addSlide();
  addContentHeader(slide, "実施ロードマップ", "3フェーズ改善計画");

  const phases = [
    { label: "Phase 1", sub: "即時対応", timeline: "1-2週間", color: C.RED, bg: "FFEBEE",
      items: ["清掃品質基準の明確化", "重点箇所の清掃手順見直し", "臭気発生源の特定と対処", "汚損箇所のスポットクリーニング"] },
    { label: "Phase 2", sub: "短期施策", timeline: "1-3ヶ月", color: C.ORANGE, bg: "FFF3E0",
      items: ["品質管理チェックリスト導入", "QCインスペクション体制構築", "定期メンテナンスサイクル導入", "カーペットディープクリーニング"] },
    { label: "Phase 3", sub: "中期施策", timeline: "3-6ヶ月", color: C.BLUE, bg: "E3F2FD",
      items: ["清掃スタッフ研修プログラム", "デジタル清掃記録システム", "設備更新計画の策定", "空気品質モニタリング導入"] },
  ];

  // Timeline arrow
  slide.addShape(pptx.shapes.RECTANGLE, { x: 0.5, y: 1.05, w: 9.0, h: 0.06, fill: { color: C.ACCENT } });

  // Phase markers
  phases.forEach((phase, pi) => {
    const px = 0.5 + pi * 3.15;

    // Circle marker on timeline
    slide.addShape(pptx.shapes.OVAL, { x: px + 1.2, y: 0.9, w: 0.35, h: 0.35, fill: { color: phase.color } });
    slide.addText(String(pi + 1), { x: px + 1.2, y: 0.9, w: 0.35, h: 0.35, fontSize: 10, fontFace: "Arial", color: C.WHITE, bold: true, align: "center", valign: "middle" });

    // Phase card
    slide.addShape(pptx.shapes.RECTANGLE, { x: px, y: 1.5, w: 2.9, h: 3.2, fill: { color: phase.bg }, shadow: shadow() });
    slide.addShape(pptx.shapes.RECTANGLE, { x: px, y: 1.5, w: 2.9, h: 0.04, fill: { color: phase.color } });

    slide.addText(phase.label, { x: px + 0.1, y: 1.58, w: 2.7, h: 0.28, fontSize: 14, fontFace: "Arial", color: phase.color, bold: true, margin: 0 });
    slide.addText(`${phase.sub}（${phase.timeline}）`, { x: px + 0.1, y: 1.88, w: 2.7, h: 0.22, fontSize: 9, fontFace: "Arial", color: C.SUBTEXT, margin: 0 });

    const itemTexts = phase.items.map(item => `- ${item}`).join("\n");
    slide.addText(itemTexts, { x: px + 0.1, y: 2.2, w: 2.7, h: 2.4, fontSize: 8, fontFace: "Arial", color: C.TEXT, valign: "top", margin: 0 });
  });

  addFooter(slide, 17);
})();

// --- Slide 18: Next Steps ---
(function buildSlide18() {
  const slide = pptx.addSlide();
  slide.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: C.NAVY } });

  slide.addText("Next Steps", { x: 0.8, y: 0.8, w: 8.4, h: 0.5, fontSize: 28, fontFace: "Arial", color: C.WHITE, bold: true, margin: 0 });
  slide.addShape(pptx.shapes.RECTANGLE, { x: 0.8, y: 1.4, w: 2.0, h: 0.04, fill: { color: "5A9BE6" } });

  const steps = [
    { num: "01", text: "本レポートの内容について経営層との合意形成", timeline: "1週間以内" },
    { num: "02", text: "URGENTホテルの現地視察と清掃責任者との改善ミーティング実施", timeline: "2週間以内" },
    { num: "03", text: "清掃チェックリストの最終版作成と全ホテルへの展開", timeline: "1ヶ月以内" },
    { num: "04", text: "改善効果の初回レビューとKPIモニタリング開始", timeline: "3ヶ月後" },
  ];

  steps.forEach((step, i) => {
    const py = 1.8 + i * 0.85;
    slide.addText(step.num, { x: 0.8, y: py, w: 0.6, h: 0.5, fontSize: 24, fontFace: "Arial", color: "5A9BE6", bold: true, margin: 0 });
    slide.addText(step.text, { x: 1.5, y: py, w: 6.5, h: 0.3, fontSize: 13, fontFace: "Arial", color: C.WHITE, margin: 0 });
    slide.addText(step.timeline, { x: 1.5, y: py + 0.3, w: 3, h: 0.25, fontSize: 10, fontFace: "Arial", color: "94A3B8", margin: 0 });
  });

  slide.addText("Thank you", { x: 0.8, y: 4.8, w: 3, h: 0.4, fontSize: 16, fontFace: "Arial", color: "64748B", italics: true, margin: 0 });
  slide.addText("Confidential | PRIMECHANGE", { x: 6, y: 5.0, w: 3.5, h: 0.3, fontSize: 9, fontFace: "Arial", color: "475569", align: "right", margin: 0 });
})();


// ============================================================
// Generate files
// ============================================================
async function generate() {
  console.log("Generating PRIMECHANGE reports...");
  console.log("Input: " + JSON_PATH);
  console.log("");

  // DOCX
  try {
    const docxPath = path.join(OUTPUT_DIR, DOCX_NAME);
    const buffer = await Packer.toBuffer(doc);
    fs.writeFileSync(docxPath, buffer);
    const docxSize = (fs.statSync(docxPath).size / 1024).toFixed(1);
    console.log(`DOCX: ${docxPath}`);
    console.log(`  Size: ${docxSize} KB`);
  } catch (err) {
    console.error("DOCX generation error:", err.message);
  }

  // PPTX
  try {
    const pptxPath = path.join(OUTPUT_DIR, PPTX_NAME);
    await pptx.writeFile({ fileName: pptxPath });
    const pptxSize = (fs.statSync(pptxPath).size / 1024).toFixed(1);
    console.log(`PPTX: ${pptxPath}`);
    console.log(`  Size: ${pptxSize} KB`);
  } catch (err) {
    console.error("PPTX generation error:", err.message);
  }

  console.log("\nDone.");
}

generate().catch(err => {
  console.error("Fatal error:", err);
  process.exit(1);
});
