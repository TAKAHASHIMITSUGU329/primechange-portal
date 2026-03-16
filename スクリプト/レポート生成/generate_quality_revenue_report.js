#!/usr/bin/env node
"use strict";

const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, PageNumber, PageBreak
} = require("docx");

// ============================================================
// Paths & Data Loading
// ============================================================
const QUALITY_JSON = path.resolve(__dirname, "primechange_portfolio_analysis.json");
const REVENUE_JSON = path.resolve(__dirname, "hotel_revenue_data.json");
const OUTPUT_DOCX = path.resolve(__dirname, "PRIMECHANGE_品質売上総合レポート.docx");

const qualityData = JSON.parse(fs.readFileSync(QUALITY_JSON, "utf-8"));
const revenueData = JSON.parse(fs.readFileSync(REVENUE_JSON, "utf-8"));

const hotelsRanked = qualityData.portfolio_overview.hotels_ranked || [];
const actionPlans = qualityData.action_plans || [];

// Key mapping: quality keys → revenue keys
const KEY_MAP = {
  keisei_kinshicho: "keisei_richmond",
  comfort_yokohama_kannai: "comfort_yokohama",
};

function getRevenue(qualityKey) {
  const mappedKey = KEY_MAP[qualityKey] || qualityKey;
  return revenueData[mappedKey] || null;
}

// ============================================================
// Color Palette
// ============================================================
const C = {
  NAVY: "1B3A5C", ACCENT: "2E75B6", WHITE: "FFFFFF",
  LIGHT_BG: "F5F7FA", TEXT: "333333", SUBTEXT: "666666",
  GREEN: "27AE60", ORANGE: "FF9800", RED: "E74C3C",
  BLUE: "2196F3", YELLOW: "FFC107", DARK_GREEN: "1B5E20",
};

const PRIORITY_COLORS = {
  URGENT:      { bg: "FFEBEE", text: "C62828" },
  HIGH:        { bg: "FFF3E0", text: "E65100" },
  STANDARD:    { bg: "E3F2FD", text: "1565C0" },
  MAINTENANCE: { bg: "E8F5E9", text: "2E7D32" },
};

const QUADRANT_COLORS = {
  "高品質×高売上": { bg: "E8F5E9", text: "1B5E20" },
  "高品質×低売上": { bg: "FFF8E1", text: "F57F17" },
  "低品質×高売上": { bg: "FFEBEE", text: "B71C1C" },
  "低品質×低売上": { bg: "FBE9E7", text: "BF360C" },
};

function priorityLabel(p) {
  return { URGENT: "🔴 緊急", HIGH: "🟠 高", STANDARD: "🔵 標準", MAINTENANCE: "🟢 維持" }[p] || p;
}

// ============================================================
// DOCX Helpers
// ============================================================
const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };
const cellMargins = { top: 60, bottom: 60, left: 80, right: 80 };

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
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 200, after: 120 },
    children: [new TextRun({ text, bold: true, size: 22, font: "Arial", color: C.NAVY })],
  });
}

function para(text, opts = {}) {
  return new Paragraph({
    spacing: { before: opts.spaceBefore || 80, after: opts.spaceAfter || 80 },
    alignment: opts.align || AlignmentType.LEFT,
    children: [new TextRun({
      text, size: opts.size || 20, font: "Arial",
      color: opts.color || C.TEXT, bold: opts.bold || false,
      italics: opts.italics || false,
    })],
  });
}

function bulletPara(text, opts = {}) {
  return new Paragraph({
    spacing: { before: 40, after: 40 },
    bullet: { level: opts.level || 0 },
    children: [new TextRun({
      text, size: opts.size || 20, font: "Arial",
      color: opts.color || C.TEXT, bold: opts.bold || false,
    })],
  });
}

function makeCell(text, opts = {}) {
  return new TableCell({
    width: opts.width ? { size: opts.width, type: WidthType.DXA } : undefined,
    shading: opts.bgColor ? { type: ShadingType.SOLID, color: opts.bgColor, fill: opts.bgColor } : undefined,
    borders,
    margins: cellMargins,
    verticalAlign: "center",
    children: [
      new Paragraph({
        alignment: opts.align || AlignmentType.CENTER,
        children: [new TextRun({
          text: String(text || "-"),
          size: opts.size || 16, font: "Arial",
          color: opts.color || C.TEXT,
          bold: opts.bold || false,
        })],
      }),
    ],
  });
}

function fmtNum(n) {
  if (n === null || n === undefined) return "-";
  return Number(n).toLocaleString("ja-JP", { maximumFractionDigits: 0 });
}

function fmtPct(n) {
  if (n === null || n === undefined) return "-";
  return (Number(n) * 100).toFixed(1) + "%";
}

function fmtYen(n) {
  if (n === null || n === undefined) return "-";
  return "¥" + Number(n).toLocaleString("ja-JP", { maximumFractionDigits: 0 });
}

// ============================================================
// Build Integrated Dataset
// ============================================================
const integratedHotels = hotelsRanked.map(q => {
  const rev = getRevenue(q.key);
  return {
    name: q.name,
    key: q.key,
    // Quality metrics
    avg_score: q.avg,
    total_reviews: q.total_reviews,
    high_rate: q.high_rate,
    low_rate: q.low_rate,
    cleaning_issue_rate: q.cleaning_issue_rate,
    cleaning_issue_count: q.cleaning_issue_count,
    priority: q.priority,
    tier: q.tier,
    // Revenue metrics
    actual_revenue: rev ? rev.actual_revenue : null,
    target_revenue: rev ? rev.target_revenue : null,
    occupancy_rate: rev ? rev.occupancy_rate : null,
    profit_rate: rev ? rev.profit_rate : null,
    net_profit: rev ? rev.actual_net_profit : null,
    room_count: rev ? rev.room_count : null,
    staff_count: rev ? rev.staff_count : null,
    adr: rev ? rev.adr : null,
    revpar: rev ? rev.revpar : null,
    phase: rev ? rev.phase : null,
    variable_cost_rate: rev ? rev.variable_cost_rate : null,
    complaint_rate: rev ? rev.complaint_rate : null,
  };
});

// Revenue median for quadrant classification
const revenues = integratedHotels.map(h => h.actual_revenue).filter(Boolean).sort((a, b) => a - b);
const medianRevenue = revenues[Math.floor(revenues.length / 2)];
const qualityThreshold = 8.0; // Score threshold for "high quality"

// Classify into quadrants
integratedHotels.forEach(h => {
  const highQuality = h.avg_score >= qualityThreshold;
  const highRevenue = h.actual_revenue >= medianRevenue;
  if (highQuality && highRevenue) h.quadrant = "高品質×高売上";
  else if (highQuality && !highRevenue) h.quadrant = "高品質×低売上";
  else if (!highQuality && highRevenue) h.quadrant = "低品質×高売上";
  else h.quadrant = "低品質×低売上";
});

// Sort by revenue descending for main table
const hotelsByRevenue = [...integratedHotels].sort((a, b) => (b.actual_revenue || 0) - (a.actual_revenue || 0));

// ============================================================
// Section 1: Cover Page
// ============================================================
function buildCoverPage() {
  return [
    new Paragraph({ spacing: { before: 2000 } }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 200 },
      children: [new TextRun({
        text: "PRIMECHANGE", size: 56, font: "Arial",
        bold: true, color: C.NAVY,
      })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 100 },
      children: [new TextRun({
        text: "ホテル品質×売上 総合分析レポート", size: 40, font: "Arial",
        bold: true, color: C.ACCENT,
      })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 600 },
      children: [new TextRun({
        text: "Quality × Revenue Integrated Analysis", size: 24, font: "Arial",
        color: C.SUBTEXT, italics: true,
      })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 100 },
      children: [new TextRun({
        text: `対象期間: 2026年2月（売上実績）`, size: 22, font: "Arial", color: C.TEXT,
      })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 100 },
      children: [new TextRun({
        text: `対象ホテル数: ${integratedHotels.length}ホテル`, size: 22, font: "Arial", color: C.TEXT,
      })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 100 },
      children: [new TextRun({
        text: `作成日: 2026年3月9日`, size: 22, font: "Arial", color: C.TEXT,
      })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 800 },
      children: [new TextRun({
        text: "株式会社PRIMECHANGE", size: 28, font: "Arial", bold: true, color: C.NAVY,
      })],
    }),
    new Paragraph({ children: [new PageBreak()] }),
  ];
}

// ============================================================
// Section 2: Executive Summary
// ============================================================
function buildExecutiveSummary() {
  const totalRevenue = integratedHotels.reduce((sum, h) => sum + (h.actual_revenue || 0), 0);
  const avgOccupancy = integratedHotels.reduce((sum, h) => sum + (h.occupancy_rate || 0), 0) / integratedHotels.length;
  const avgScore = integratedHotels.reduce((sum, h) => sum + h.avg_score, 0) / integratedHotels.length;
  const urgentCount = integratedHotels.filter(h => h.priority === "URGENT").length;
  const highCount = integratedHotels.filter(h => h.priority === "HIGH").length;

  const elements = [
    heading1("1. エグゼクティブサマリー"),
    para("PRIMECHANGEが管理する全19ホテルの品質（口コミ分析）と売上（月次実績）を統合分析した結果を報告します。"),
    heading3("ポートフォリオ全体概況"),
  ];

  // KPI summary table
  const kpiData = [
    ["指標", "数値"],
    ["管理ホテル数", `${integratedHotels.length}ホテル`],
    ["月間総売上（2月実績）", fmtYen(totalRevenue)],
    ["平均稼働率", fmtPct(avgOccupancy)],
    ["平均品質スコア", `${avgScore.toFixed(2)} / 10.0`],
    ["緊急対応ホテル数", `${urgentCount}ホテル`],
    ["要注意ホテル数", `${highCount}ホテル`],
  ];

  const kpiRows = kpiData.map((row, i) =>
    new TableRow({
      children: [
        makeCell(row[0], { width: 4000, bold: i === 0, bgColor: i === 0 ? C.NAVY : (i % 2 === 0 ? C.LIGHT_BG : C.WHITE), color: i === 0 ? C.WHITE : C.TEXT, align: AlignmentType.LEFT }),
        makeCell(row[1], { width: 5000, bold: i === 0, bgColor: i === 0 ? C.NAVY : (i % 2 === 0 ? C.LIGHT_BG : C.WHITE), color: i === 0 ? C.WHITE : C.TEXT }),
      ],
    })
  );

  elements.push(new Table({ rows: kpiRows, width: { size: 9000, type: WidthType.DXA } }));

  elements.push(
    para(""),
    heading3("重要な発見"),
    bulletPara("低品質×高売上のホテル（コートホテル新横浜、チサンホテル浜松町など）は、品質低下による売上減少リスクが最も大きい", { bold: true }),
    bulletPara("コンフォートスイーツ東京ベイは品質・売上ともに最高水準で、成功モデルとして横展開すべき"),
    bulletPara("URGENT判定の4ホテル（博多・蒲田・浜松町・新横浜）の合計月間売上は約1,973万円 ─ 品質改善による売上維持が急務"),
    bulletPara("全体の平均稼働率72.7%に対し、品質上位ホテルは平均75%超と、品質と稼働率の正相関が確認される"),
    new Paragraph({ children: [new PageBreak()] }),
  );

  return elements;
}

// ============================================================
// Section 3: All Hotels Overview Table
// ============================================================
function buildOverviewTable() {
  const elements = [
    heading1("2. 全ホテル一覧（品質×売上）"),
    para("19ホテルを売上高順に一覧表示。品質スコア・稼働率・売上・純利益を横断的に比較します。"),
    para(""),
  ];

  const headers = ["#", "ホテル名", "品質\nスコア", "レビュー\n件数", "稼働率", "月間売上\n（2月）", "純利益", "利益率", "ADR", "優先度"];
  const colWidths = [400, 2800, 700, 700, 700, 1300, 1200, 700, 900, 900];

  const headerRow = new TableRow({
    tableHeader: true,
    children: headers.map((h, i) =>
      makeCell(h, { width: colWidths[i], bold: true, bgColor: C.NAVY, color: C.WHITE, size: 14 })
    ),
  });

  const dataRows = hotelsByRevenue.map((h, idx) => {
    const pc = PRIORITY_COLORS[h.priority] || PRIORITY_COLORS.STANDARD;
    return new TableRow({
      children: [
        makeCell(idx + 1, { width: colWidths[0], size: 14, bgColor: idx % 2 === 0 ? C.WHITE : C.LIGHT_BG }),
        makeCell(h.name, { width: colWidths[1], size: 13, align: AlignmentType.LEFT, bgColor: idx % 2 === 0 ? C.WHITE : C.LIGHT_BG }),
        makeCell(h.avg_score.toFixed(2), { width: colWidths[2], size: 14, bgColor: idx % 2 === 0 ? C.WHITE : C.LIGHT_BG, color: h.avg_score < 7.5 ? C.RED : (h.avg_score >= 8.5 ? C.GREEN : C.TEXT) }),
        makeCell(h.total_reviews, { width: colWidths[3], size: 14, bgColor: idx % 2 === 0 ? C.WHITE : C.LIGHT_BG }),
        makeCell(fmtPct(h.occupancy_rate), { width: colWidths[4], size: 14, bgColor: idx % 2 === 0 ? C.WHITE : C.LIGHT_BG }),
        makeCell(fmtYen(h.actual_revenue), { width: colWidths[5], size: 14, bgColor: idx % 2 === 0 ? C.WHITE : C.LIGHT_BG }),
        makeCell(fmtYen(h.net_profit), { width: colWidths[6], size: 14, bgColor: idx % 2 === 0 ? C.WHITE : C.LIGHT_BG }),
        makeCell(fmtPct(h.profit_rate), { width: colWidths[7], size: 14, bgColor: idx % 2 === 0 ? C.WHITE : C.LIGHT_BG }),
        makeCell(h.adr ? fmtYen(h.adr) : "-", { width: colWidths[8], size: 14, bgColor: idx % 2 === 0 ? C.WHITE : C.LIGHT_BG }),
        makeCell(priorityLabel(h.priority), { width: colWidths[9], size: 13, bgColor: pc.bg, color: pc.text, bold: true }),
      ],
    });
  });

  elements.push(new Table({ rows: [headerRow, ...dataRows], width: { size: 10300, type: WidthType.DXA } }));
  elements.push(
    para(""),
    para("※ 品質スコアは海外OTA口コミの10点満点評価平均。売上データは2026年2月実績。", { size: 16, color: C.SUBTEXT, italics: true }),
    para("※ 優先度: 🔴緊急=清掃クレーム率10%超or低スコア, 🟠高=改善余地大, 🔵標準=良好, 🟢維持=最良", { size: 16, color: C.SUBTEXT, italics: true }),
    new Paragraph({ children: [new PageBreak()] }),
  );

  return elements;
}

// ============================================================
// Section 4: Quality × Revenue Matrix
// ============================================================
function buildQuadrantAnalysis() {
  const elements = [
    heading1("3. 品質×売上マトリクス分析"),
    para(`品質スコア${qualityThreshold}点を基準に「高品質/低品質」、月間売上の中央値（${fmtYen(medianRevenue)}）を基準に「高売上/低売上」で4象限に分類します。`),
    para(""),
  ];

  const quadrants = [
    { name: "高品質×高売上", label: "維持・横展開モデル", desc: "品質・売上ともに高水準。成功要因を他ホテルに横展開すべき。" },
    { name: "高品質×低売上", label: "営業強化対象", desc: "品質は高いが売上が伸びていない。マーケティング・価格戦略の見直しが必要。" },
    { name: "低品質×高売上", label: "緊急改善対象（損失リスク大）", desc: "売上は高いが品質が低い。口コミ悪化による売上急落リスクが最も高い。最優先で改善。" },
    { name: "低品質×低売上", label: "根本改革対象", desc: "品質・売上ともに課題。抜本的な改革プログラムが必要。" },
  ];

  for (const q of quadrants) {
    const qc = QUADRANT_COLORS[q.name] || { bg: C.LIGHT_BG, text: C.TEXT };
    const hotels = integratedHotels.filter(h => h.quadrant === q.name);

    elements.push(heading2(`${q.name} ─ ${q.label}`));
    elements.push(para(q.desc));

    if (hotels.length === 0) {
      elements.push(para("（該当ホテルなし）", { italics: true, color: C.SUBTEXT }));
    } else {
      const rows = [
        new TableRow({
          children: [
            makeCell("ホテル名", { width: 3000, bold: true, bgColor: qc.bg, color: qc.text, size: 15, align: AlignmentType.LEFT }),
            makeCell("品質スコア", { width: 1200, bold: true, bgColor: qc.bg, color: qc.text, size: 15 }),
            makeCell("月間売上", { width: 1800, bold: true, bgColor: qc.bg, color: qc.text, size: 15 }),
            makeCell("稼働率", { width: 1000, bold: true, bgColor: qc.bg, color: qc.text, size: 15 }),
            makeCell("優先度", { width: 1200, bold: true, bgColor: qc.bg, color: qc.text, size: 15 }),
          ],
        }),
      ];
      for (const h of hotels.sort((a, b) => (b.actual_revenue || 0) - (a.actual_revenue || 0))) {
        const pc = PRIORITY_COLORS[h.priority];
        rows.push(new TableRow({
          children: [
            makeCell(h.name, { width: 3000, size: 15, align: AlignmentType.LEFT }),
            makeCell(h.avg_score.toFixed(2), { width: 1200, size: 15 }),
            makeCell(fmtYen(h.actual_revenue), { width: 1800, size: 15 }),
            makeCell(fmtPct(h.occupancy_rate), { width: 1000, size: 15 }),
            makeCell(priorityLabel(h.priority), { width: 1200, size: 14, bgColor: pc.bg, color: pc.text, bold: true }),
          ],
        }));
      }
      elements.push(new Table({ rows, width: { size: 8200, type: WidthType.DXA } }));
    }
    elements.push(para(""));
  }

  elements.push(new Paragraph({ children: [new PageBreak()] }));
  return elements;
}

// ============================================================
// Section 5: PRIMECHANGE Action Plan
// ============================================================
function buildActionPlan() {
  const elements = [
    heading1("4. PRIMECHANGEアクションプラン"),
    para("品質×売上の統合分析に基づき、PRIMECHANGEが具体的に何をすべきかを優先度・時間軸別に明示します。売上インパクトの大きいホテルから優先的に対応します。"),
    para(""),
  ];

  // Phase 1: Immediate (This Week)
  elements.push(heading2("Phase 1: 最優先対応（今週中）"));
  elements.push(para("低品質×高売上ホテルの緊急対応 ─ 口コミ悪化による売上減少を防ぐ", { bold: true, color: C.RED }));

  const urgentHighRevenue = integratedHotels
    .filter(h => h.quadrant === "低品質×高売上" || (h.priority === "URGENT" && h.actual_revenue >= medianRevenue))
    .sort((a, b) => (b.actual_revenue || 0) - (a.actual_revenue || 0));

  for (const h of urgentHighRevenue) {
    elements.push(heading3(`${h.name}（月間売上: ${fmtYen(h.actual_revenue)}, スコア: ${h.avg_score}）`));
    elements.push(bulletPara(`現状: 品質スコア${h.avg_score}（${h.priority}判定）、稼働率${fmtPct(h.occupancy_rate)}`));
    elements.push(bulletPara("リスク: 口コミ悪化が継続すれば、OTAランキング低下→予約減→売上減の悪循環に陥る可能性", { color: C.RED }));
    elements.push(bulletPara("アクション①: 清掃品質の緊急監査 ─ PRIMECHANGEスタッフによる抜き打ち検査を今週中に実施"));
    elements.push(bulletPara("アクション②: 低評価レビューの分析→具体的改善ポイントの特定（3日以内）"));
    elements.push(bulletPara("アクション③: ホテル責任者との緊急ミーティング設定（週内）"));
    elements.push(bulletPara(`成功指標: 1ヶ月後に清掃クレーム率${h.cleaning_issue_rate ? (h.cleaning_issue_rate * 100).toFixed(0) : "N/A"}%→5%以下`));
    elements.push(para(""));
  }

  // Phase 2: Short-term (1 month)
  elements.push(heading2("Phase 2: 短期改善（1ヶ月以内）"));
  elements.push(para("URGENTホテル全体の改善プログラム実施", { bold: true, color: C.ORANGE }));

  const urgentHotels = integratedHotels.filter(h => h.priority === "URGENT").sort((a, b) => (b.actual_revenue || 0) - (a.actual_revenue || 0));
  const urgentTotalRevenue = urgentHotels.reduce((s, h) => s + (h.actual_revenue || 0), 0);

  elements.push(para(`対象: ${urgentHotels.length}ホテル（合計月間売上: ${fmtYen(urgentTotalRevenue)}）`));

  const urgentTable = [
    new TableRow({
      children: [
        makeCell("ホテル名", { width: 2500, bold: true, bgColor: C.RED, color: C.WHITE, size: 15, align: AlignmentType.LEFT }),
        makeCell("品質スコア", { width: 1000, bold: true, bgColor: C.RED, color: C.WHITE, size: 15 }),
        makeCell("月間売上", { width: 1500, bold: true, bgColor: C.RED, color: C.WHITE, size: 15 }),
        makeCell("主要課題", { width: 2500, bold: true, bgColor: C.RED, color: C.WHITE, size: 15, align: AlignmentType.LEFT }),
        makeCell("目標スコア", { width: 1000, bold: true, bgColor: C.RED, color: C.WHITE, size: 15 }),
      ],
    }),
  ];

  const urgentIssues = {
    comfort_hakata: "清掃品質・設備老朽化",
    apa_kamata: "清掃品質・騒音対策",
    chisan: "清掃品質・アメニティ不足",
    court_shinyokohama: "清掃品質・接客対応",
  };

  for (const h of urgentHotels) {
    urgentTable.push(new TableRow({
      children: [
        makeCell(h.name, { width: 2500, size: 15, align: AlignmentType.LEFT }),
        makeCell(h.avg_score.toFixed(2), { width: 1000, size: 15, color: C.RED }),
        makeCell(fmtYen(h.actual_revenue), { width: 1500, size: 15 }),
        makeCell(urgentIssues[h.key] || "清掃品質の総合改善", { width: 2500, size: 15, align: AlignmentType.LEFT }),
        makeCell((h.avg_score + 0.5).toFixed(1) + "以上", { width: 1000, size: 15, color: C.GREEN }),
      ],
    }));
  }

  elements.push(new Table({ rows: urgentTable, width: { size: 8500, type: WidthType.DXA } }));
  elements.push(para(""));

  elements.push(bulletPara("共通アクション①: 清掃SOP（標準作業手順）の統一版を作成し、4ホテル全てに導入", { bold: true }));
  elements.push(bulletPara("共通アクション②: PRIMECHANGEの清掃チームリーダーが各ホテルを週次巡回"));
  elements.push(bulletPara("共通アクション③: 口コミ返信体制の構築 ─ 低評価レビューへの24時間以内の対応ルール"));
  elements.push(bulletPara("共通アクション④: スタッフ研修（清掃基準・接客マナー）を月2回実施"));

  elements.push(para(""));

  // Phase 3: Medium-term (3 months)
  elements.push(heading2("Phase 3: 中期改善（3ヶ月）"));
  elements.push(para("HIGHホテルをSTANDARDに引き上げ + 営業強化対象ホテルの売上改善", { bold: true, color: C.ACCENT }));

  const highHotels = integratedHotels.filter(h => h.priority === "HIGH").sort((a, b) => (b.actual_revenue || 0) - (a.actual_revenue || 0));

  for (const h of highHotels) {
    elements.push(heading3(`${h.name}（月間売上: ${fmtYen(h.actual_revenue)}, スコア: ${h.avg_score}）`));
    elements.push(bulletPara(`現状分析: ${h.priority}判定、稼働率${fmtPct(h.occupancy_rate)}、利益率${fmtPct(h.profit_rate)}`));
    if (h.key === "comment_yokohama") {
      elements.push(bulletPara("改善ポイント: レビュー件数95件と多く、高スコア（8.88）だがHIGH判定 → 清掃クレーム対応の強化で9.0超を目指す"));
      elements.push(bulletPara("アクション: 清掃チェックリストの徹底、特別清掃の定期実施"));
    } else if (h.key === "kawasaki_nikko") {
      elements.push(bulletPara("改善ポイント: 高売上（¥7.2M）だがスコア8.60 → 清掃品質の微調整でSTANDARDへ"));
      elements.push(bulletPara("アクション: ゲストフィードバックの即日共有体制、アメニティグレードの見直し"));
    } else if (h.key === "comfort_roppongi") {
      elements.push(bulletPara("改善ポイント: スコア8.08でHIGH判定 → 清掃品質と設備メンテナンスの改善"));
      elements.push(bulletPara("アクション: 水回り設備の重点清掃、リネン品質のアップグレード"));
    }
    elements.push(para(""));
  }

  // High quality + Low revenue hotels
  const salesBoost = integratedHotels.filter(h => h.quadrant === "高品質×低売上").sort((a, b) => a.avg_score - b.avg_score);
  if (salesBoost.length > 0) {
    elements.push(heading3("営業強化対象ホテル（高品質×低売上）"));
    for (const h of salesBoost) {
      elements.push(bulletPara(`${h.name}: スコア${h.avg_score}（品質良好）だが売上${fmtYen(h.actual_revenue)} → OTA掲載最適化、価格戦略の見直し、プロモーション強化`));
    }
    elements.push(para(""));
  }

  // Phase 4: Long-term (6 months)
  elements.push(heading2("Phase 4: 長期目標（6ヶ月）"));
  elements.push(para("全ポートフォリオの底上げとスケール", { bold: true, color: C.DARK_GREEN }));
  elements.push(bulletPara("目標①: URGENT判定ホテル → 0件（全ホテルSTANDARD以上）", { bold: true }));
  elements.push(bulletPara("目標②: ポートフォリオ平均スコア 8.39 → 8.80以上"));
  elements.push(bulletPara("目標③: 清掃クレーム率 全ホテル5%以下"));
  elements.push(bulletPara("目標④: 月間総売上の3%改善（品質改善による稼働率・ADR向上）"));
  elements.push(para(""));
  elements.push(bulletPara("施策①: PRIMECHANGEベストプラクティス集の作成 ─ コンフォートスイーツ東京ベイのモデルを体系化"));
  elements.push(bulletPara("施策②: ホテル間ベンチマーキング制度の導入（月次品質レポートの共有会）"));
  elements.push(bulletPara("施策③: 清掃品質のAI活用（チェックリストアプリ導入、写真による清掃完了確認）"));
  elements.push(bulletPara("施策④: スタッフ評価制度の改革（品質KPIを評価基準に組み込み）"));

  elements.push(new Paragraph({ children: [new PageBreak()] }));
  return elements;
}

// ============================================================
// Section 6: Hotel Detail Cards
// ============================================================
function buildHotelCards() {
  const elements = [
    heading1("5. ホテル別詳細カード"),
    para("各ホテルの品質・売上のサマリーと推奨アクションを一覧します。"),
    para(""),
  ];

  // URGENT and HIGH hotels get detailed cards
  const detailedHotels = integratedHotels
    .filter(h => h.priority === "URGENT" || h.priority === "HIGH")
    .sort((a, b) => {
      const order = { URGENT: 0, HIGH: 1, STANDARD: 2, MAINTENANCE: 3 };
      return (order[a.priority] - order[b.priority]) || ((b.actual_revenue || 0) - (a.actual_revenue || 0));
    });

  const otherHotels = integratedHotels
    .filter(h => h.priority !== "URGENT" && h.priority !== "HIGH")
    .sort((a, b) => (b.actual_revenue || 0) - (a.actual_revenue || 0));

  for (const h of detailedHotels) {
    const pc = PRIORITY_COLORS[h.priority];

    elements.push(heading2(`${h.name}`));

    // Card table
    const rows = [
      new TableRow({
        children: [
          makeCell("指標", { width: 2500, bold: true, bgColor: pc.bg, color: pc.text, size: 15, align: AlignmentType.LEFT }),
          makeCell("数値", { width: 2500, bold: true, bgColor: pc.bg, color: pc.text, size: 15 }),
          makeCell("指標", { width: 2500, bold: true, bgColor: pc.bg, color: pc.text, size: 15, align: AlignmentType.LEFT }),
          makeCell("数値", { width: 2500, bold: true, bgColor: pc.bg, color: pc.text, size: 15 }),
        ],
      }),
      new TableRow({
        children: [
          makeCell("品質スコア", { width: 2500, size: 15, align: AlignmentType.LEFT }),
          makeCell(`${h.avg_score} / 10.0`, { width: 2500, size: 15, bold: true, color: h.avg_score < 7.5 ? C.RED : C.TEXT }),
          makeCell("月間売上（2月）", { width: 2500, size: 15, align: AlignmentType.LEFT }),
          makeCell(fmtYen(h.actual_revenue), { width: 2500, size: 15, bold: true }),
        ],
      }),
      new TableRow({
        children: [
          makeCell("レビュー件数", { width: 2500, size: 15, align: AlignmentType.LEFT }),
          makeCell(`${h.total_reviews}件`, { width: 2500, size: 15 }),
          makeCell("純利益", { width: 2500, size: 15, align: AlignmentType.LEFT }),
          makeCell(fmtYen(h.net_profit), { width: 2500, size: 15 }),
        ],
      }),
      new TableRow({
        children: [
          makeCell("清掃クレーム率", { width: 2500, size: 15, align: AlignmentType.LEFT }),
          makeCell(h.cleaning_issue_rate ? (h.cleaning_issue_rate * 100).toFixed(1) + "%" : "-", { width: 2500, size: 15, color: h.cleaning_issue_rate > 0.1 ? C.RED : C.TEXT }),
          makeCell("稼働率", { width: 2500, size: 15, align: AlignmentType.LEFT }),
          makeCell(fmtPct(h.occupancy_rate), { width: 2500, size: 15 }),
        ],
      }),
      new TableRow({
        children: [
          makeCell("優先度", { width: 2500, size: 15, align: AlignmentType.LEFT }),
          makeCell(priorityLabel(h.priority), { width: 2500, size: 15, bold: true, bgColor: pc.bg, color: pc.text }),
          makeCell("利益率", { width: 2500, size: 15, align: AlignmentType.LEFT }),
          makeCell(fmtPct(h.profit_rate), { width: 2500, size: 15 }),
        ],
      }),
      new TableRow({
        children: [
          makeCell("マトリクス分類", { width: 2500, size: 15, align: AlignmentType.LEFT }),
          makeCell(h.quadrant, { width: 2500, size: 15, bold: true }),
          makeCell("客室数 / スタッフ数", { width: 2500, size: 15, align: AlignmentType.LEFT }),
          makeCell(`${h.room_count || "-"}室 / ${h.staff_count || "-"}名`, { width: 2500, size: 15 }),
        ],
      }),
    ];

    elements.push(new Table({ rows, width: { size: 10000, type: WidthType.DXA } }));
    elements.push(para(""));

    // Find action plan data
    const ap = actionPlans.find(a => a.hotel_key === h.key);
    if (ap) {
      elements.push(para("推奨アクション:", { bold: true, color: C.NAVY }));
      if (ap.immediate_actions) {
        for (const action of ap.immediate_actions.slice(0, 3)) {
          elements.push(bulletPara(`${action.action || action}`, { size: 18 }));
        }
      }
    }
    elements.push(para(""));
  }

  // STANDARD and MAINTENANCE hotels - compact format
  elements.push(heading2("STANDARD / MAINTENANCEホテル一覧"));
  elements.push(para("品質が良好なホテル群。現状の品質維持と売上最大化を目指します。"));

  const compactRows = [
    new TableRow({
      children: [
        makeCell("ホテル名", { width: 2800, bold: true, bgColor: C.ACCENT, color: C.WHITE, size: 15, align: AlignmentType.LEFT }),
        makeCell("スコア", { width: 900, bold: true, bgColor: C.ACCENT, color: C.WHITE, size: 15 }),
        makeCell("月間売上", { width: 1500, bold: true, bgColor: C.ACCENT, color: C.WHITE, size: 15 }),
        makeCell("稼働率", { width: 900, bold: true, bgColor: C.ACCENT, color: C.WHITE, size: 15 }),
        makeCell("象限", { width: 2000, bold: true, bgColor: C.ACCENT, color: C.WHITE, size: 15 }),
        makeCell("優先度", { width: 1000, bold: true, bgColor: C.ACCENT, color: C.WHITE, size: 15 }),
      ],
    }),
  ];

  for (const h of otherHotels) {
    const pc = PRIORITY_COLORS[h.priority];
    compactRows.push(new TableRow({
      children: [
        makeCell(h.name, { width: 2800, size: 14, align: AlignmentType.LEFT }),
        makeCell(h.avg_score.toFixed(2), { width: 900, size: 14 }),
        makeCell(fmtYen(h.actual_revenue), { width: 1500, size: 14 }),
        makeCell(fmtPct(h.occupancy_rate), { width: 900, size: 14 }),
        makeCell(h.quadrant, { width: 2000, size: 14 }),
        makeCell(priorityLabel(h.priority), { width: 1000, size: 13, bgColor: pc.bg, color: pc.text, bold: true }),
      ],
    }));
  }

  elements.push(new Table({ rows: compactRows, width: { size: 9100, type: WidthType.DXA } }));
  elements.push(para(""));

  return elements;
}

// ============================================================
// Build & Save Document
// ============================================================
async function main() {
  console.log("Building integrated Quality × Revenue DOCX report...");

  const sections = [
    ...buildCoverPage(),
    ...buildExecutiveSummary(),
    ...buildOverviewTable(),
    ...buildQuadrantAnalysis(),
    ...buildActionPlan(),
    ...buildHotelCards(),
  ];

  const doc = new Document({
    styles: {
      default: {
        document: { run: { font: "Arial", size: 20 } },
      },
    },
    sections: [
      {
        properties: {
          page: {
            size: { width: 12240, height: 15840 },
            margin: { top: 720, bottom: 720, left: 900, right: 900 },
          },
        },
        headers: {
          default: new Header({
            children: [
              new Paragraph({
                alignment: AlignmentType.RIGHT,
                children: [new TextRun({
                  text: "PRIMECHANGE ホテル品質×売上 総合分析レポート",
                  size: 14, font: "Arial", color: C.SUBTEXT, italics: true,
                })],
              }),
            ],
          }),
        },
        footers: {
          default: new Footer({
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                  new TextRun({ text: "© 2026 PRIMECHANGE Inc. | ", size: 14, font: "Arial", color: C.SUBTEXT }),
                  new TextRun({ children: [PageNumber.CURRENT], size: 14, font: "Arial", color: C.SUBTEXT }),
                  new TextRun({ text: " / ", size: 14, font: "Arial", color: C.SUBTEXT }),
                  new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 14, font: "Arial", color: C.SUBTEXT }),
                ],
              }),
            ],
          }),
        },
        children: sections,
      },
    ],
  });

  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(OUTPUT_DOCX, buffer);
  console.log(`✅ DOCX saved: ${OUTPUT_DOCX} (${(buffer.length / 1024).toFixed(1)} KB)`);

  // Print summary stats
  console.log(`\n📊 Summary:`);
  console.log(`   Hotels: ${integratedHotels.length}`);
  console.log(`   Revenue median: ${fmtYen(medianRevenue)}`);
  console.log(`   Quadrants:`);
  for (const q of ["高品質×高売上", "高品質×低売上", "低品質×高売上", "低品質×低売上"]) {
    const list = integratedHotels.filter(h => h.quadrant === q);
    console.log(`     ${q}: ${list.length}ホテル - ${list.map(h => h.name).join(", ")}`);
  }
}

main().catch(e => { console.error(e); process.exit(1); });
