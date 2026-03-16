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
const data = JSON.parse(fs.readFileSync(path.resolve(__dirname, "analysis_1_data.json"), "utf-8"));
const meta = data.analysis_metadata;
const typeRanking = data.type_frequency_ranking;
const catBreakdown = data.category_breakdown;
const corrAnalysis = data.correlation_analysis;
const hotelProfiles = data.hotel_profiles;
const revImpact = data.revenue_impact;
const priorities = data.improvement_priorities;
const recs = data.recommendations;
const stats = data.summary_stats;

// ============================================================
// Styles
// ============================================================
const C = { NAVY: "1B3A5C", ACCENT: "2E75B6", WHITE: "FFFFFF", LIGHT_BG: "F5F7FA", TEXT: "333333", SUBTEXT: "666666", GREEN: "27AE60", ORANGE: "FF9800", RED: "E74C3C", BLUE: "2196F3", TEAL: "00695C", DARK_GREEN: "1B5E20", AMBER: "F57F17", PURPLE: "7B1FA2" };
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
function fmtN(n, d = 1) { return n != null ? Number(n).toFixed(d) : "-"; }
function fmtPct(n) { return n != null ? Number(n).toFixed(1) + "%" : "-"; }
const PB = () => new Paragraph({ children: [new PageBreak()] });

// ============================================================
// DOCX Sections
// ============================================================
function buildCover() {
  return [
    new Paragraph({ spacing: { before: 2400 } }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 }, children: [new TextRun({ text: "PRIMECHANGE", size: 56, font: "Arial", bold: true, color: C.NAVY })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [new TextRun({ text: "分析1", size: 44, font: "Arial", bold: true, color: C.ACCENT })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [new TextRun({ text: "クレーム類型×口コミスコア連動分析", size: 36, font: "Arial", bold: true, color: C.ACCENT })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 400 }, children: [new TextRun({ text: "Claim Type × Review Score Correlation Analysis", size: 22, font: "Arial", color: C.SUBTEXT, italics: true })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [new TextRun({ text: `対象: ${meta.total_hotels}ホテル | 総クレーム: ${meta.total_claims}件 | 期間: ${meta.data_period}`, size: 20, font: "Arial", color: C.TEXT })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [new TextRun({ text: "2025年3月", size: 22, font: "Arial", color: C.TEXT })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 800 }, children: [new TextRun({ text: "株式会社PRIMECHANGE", size: 28, font: "Arial", bold: true, color: C.NAVY })] }),
    PB(),
  ];
}

function buildExecutiveSummary() {
  const overallR = corrAnalysis.overall_correlation.r;
  const s50 = revImpact.scenarios.find(s => s.reduction_pct === 50);
  return [
    h1("1. エグゼクティブサマリー"),
    p("本レポートは、PRIMECHANGEが管理する19ホテルの月次クレームデータ（13類型）を分析し、口コミスコアおよび売上との関連を定量化したものです。クレーム類型ごとの「スコアへの影響度」と「収益インパクト」を明らかにし、優先的に取り組むべき改善テーマを特定しています。"),

    h3("重要な発見"),
    bp(`総クレーム${meta.total_claims}件を13類型に分類。「セット漏れ」が${typeRanking[0].share_pct}%、「残置」が${typeRanking[1].share_pct}%と上位2類型で全体の55.9%を占有。`, { b: true }),
    bp(`カテゴリ別では「客室準備系」が全体の${catBreakdown["客室準備系"].share_pct}%と最多。セット漏れ・残置・手配ミスなど、清掃後のチェック体制で防止可能な類型が中心。`, { b: true }),
    bp(`3ホテル（目白・成田・横浜関内）がゼロクレームを達成。これらのオペレーション手法の横展開が有効。`, { b: true }),
    bp(`クレーム50%削減で年間約${fmtY(s50 ? s50.total_annual_impact : 0)}の収益改善ポテンシャル（分析6の弾力性を適用）。`, { b: true }),

    h3("分析の特徴"),
    bp("全クレーム率は6.7件/万室と極めて低水準。高品質オペレーション基盤の上でのさらなる改善余地を探索。"),
    bp("相関分析は正の弱い相関（r=0.20）を示すが、これは高品質ホテルほど検知・報告体制が整っているためと推察。"),
    bp("改善優先度は「頻度×スコア影響度×波及度」の複合評価で算出し、現場で即実行可能な提言を導出。"),
    PB(),
  ];
}

function buildTypeFrequency() {
  const activeTypes = typeRanking.filter(t => t.total_count > 0);
  const headerRow = new TableRow({
    children: [
      cell("順位", { bg: C.NAVY, c: C.WHITE, b: true, w: 800 }),
      cell("クレーム類型", { bg: C.NAVY, c: C.WHITE, b: true, w: 2400 }),
      cell("件数", { bg: C.NAVY, c: C.WHITE, b: true, w: 1200 }),
      cell("構成比", { bg: C.NAVY, c: C.WHITE, b: true, w: 1200 }),
      cell("影響ホテル数", { bg: C.NAVY, c: C.WHITE, b: true, w: 1600 }),
      cell("重要度", { bg: C.NAVY, c: C.WHITE, b: true, w: 1600 }),
    ]
  });
  const dataRows = activeTypes.map((t, i) => {
    const severity = t.total_count >= 10 ? "高" : t.total_count >= 5 ? "中" : "低";
    const sevColor = severity === "高" ? C.RED : severity === "中" ? C.ORANGE : C.TEXT;
    return new TableRow({
      children: [
        cell(i + 1, { bg: i % 2 ? C.WHITE : C.LIGHT_BG }),
        cell(t.type, { bg: i % 2 ? C.WHITE : C.LIGHT_BG, a: AlignmentType.LEFT }),
        cell(t.total_count, { bg: i % 2 ? C.WHITE : C.LIGHT_BG, b: true }),
        cell(fmtPct(t.share_pct), { bg: i % 2 ? C.WHITE : C.LIGHT_BG }),
        cell(`${t.hotels_affected}/19`, { bg: i % 2 ? C.WHITE : C.LIGHT_BG }),
        cell(severity, { bg: i % 2 ? C.WHITE : C.LIGHT_BG, c: sevColor, b: true }),
      ]
    });
  });
  const freqTable = new Table({ rows: [headerRow, ...dataRows], width: { size: 8800, type: WidthType.DXA } });

  // Category breakdown table
  const catHeader = new TableRow({
    children: [
      cell("カテゴリ", { bg: C.ACCENT, c: C.WHITE, b: true, w: 2400 }),
      cell("含まれる類型", { bg: C.ACCENT, c: C.WHITE, b: true, w: 3600 }),
      cell("件数", { bg: C.ACCENT, c: C.WHITE, b: true, w: 1200 }),
      cell("構成比", { bg: C.ACCENT, c: C.WHITE, b: true, w: 1600 }),
    ]
  });
  const catRows = Object.entries(catBreakdown).sort((a, b) => b[1].count - a[1].count).map(([name, d], i) => {
    return new TableRow({
      children: [
        cell(name, { bg: i % 2 ? C.WHITE : C.LIGHT_BG, b: true, a: AlignmentType.LEFT }),
        cell(d.types.join("、"), { bg: i % 2 ? C.WHITE : C.LIGHT_BG, a: AlignmentType.LEFT, sz: 14 }),
        cell(d.count, { bg: i % 2 ? C.WHITE : C.LIGHT_BG, b: true }),
        cell(fmtPct(d.share_pct), { bg: i % 2 ? C.WHITE : C.LIGHT_BG }),
      ]
    });
  });
  const catTable = new Table({ rows: [catHeader, ...catRows], width: { size: 8800, type: WidthType.DXA } });

  return [
    h1("2. クレーム類型別頻度分析"),
    h2("2.1 類型別発生件数ランキング"),
    p(`${meta.data_period}の全19ホテルで発生した総クレーム${meta.total_claims}件を13類型に分類し、頻度順にランキングしました。`),
    freqTable,
    p(`上位2類型（セット漏れ・残置）だけで全体の55.9%を占め、改善の焦点が明確です。特に「セット漏れ」は10ホテルで発生しており、横断的な対策が求められます。`, { sb: 200 }),

    h2("2.2 カテゴリ別分析"),
    p("13類型を4カテゴリに集約し、構造的な傾向を分析しました。"),
    catTable,
    p(`「客室準備系」が過半数を超えており、これは清掃完了後のチェック工程で防止可能なクレームが大半を占めていることを示しています。チェッカー体制の強化が最も効率的な改善施策と言えます。`, { sb: 200 }),
    PB(),
  ];
}

function buildCorrelation() {
  const overall = corrAnalysis.overall_correlation;
  const typCorrs = corrAnalysis.type_correlations.filter(tc => tc.abs_r > 0);

  // Correlation table
  const corrHeader = new TableRow({
    children: [
      cell("クレーム類型", { bg: C.NAVY, c: C.WHITE, b: true, w: 2400 }),
      cell("相関係数 r", { bg: C.NAVY, c: C.WHITE, b: true, w: 1600 }),
      cell("絶対値 |r|", { bg: C.NAVY, c: C.WHITE, b: true, w: 1400 }),
      cell("判定", { bg: C.NAVY, c: C.WHITE, b: true, w: 1800 }),
      cell("方向性", { bg: C.NAVY, c: C.WHITE, b: true, w: 1600 }),
    ]
  });
  const corrRows = typCorrs.slice(0, 10).map((tc, i) => {
    const judgeColor = tc.abs_r >= 0.4 ? C.RED : tc.abs_r >= 0.2 ? C.ORANGE : C.TEXT;
    return new TableRow({
      children: [
        cell(tc.type, { bg: i % 2 ? C.WHITE : C.LIGHT_BG, a: AlignmentType.LEFT }),
        cell(fmtN(tc.correlation_r, 4), { bg: i % 2 ? C.WHITE : C.LIGHT_BG }),
        cell(fmtN(tc.abs_r, 4), { bg: i % 2 ? C.WHITE : C.LIGHT_BG, b: true }),
        cell(tc.interpretation, { bg: i % 2 ? C.WHITE : C.LIGHT_BG, c: judgeColor }),
        cell(tc.direction, { bg: i % 2 ? C.WHITE : C.LIGHT_BG, sz: 14 }),
      ]
    });
  });
  const corrTable = new Table({ rows: [corrHeader, ...corrRows], width: { size: 8800, type: WidthType.DXA } });

  return [
    h1("3. クレーム率×口コミスコア相関分析"),
    h2("3.1 全体の相関"),
    p(`全19ホテルのクレーム率と口コミスコアの相関係数は r=${fmtN(overall.r, 4)}（${overall.interpretation}）です。`),
    p("一般的にはクレーム率が高いほどスコアが低下するという負の相関が期待されますが、今回のデータでは弱い正の相関を示しました。これには以下の理由が考えられます："),
    bp("全ホテルのクレーム率が0.00〜0.15%と極めて低水準にあり、スコアへの直接的影響が限定的"),
    bp("高品質なホテルほどクレーム検知・報告体制が整っており、「報告件数＝品質意識の高さ」を反映"),
    bp("サンプルサイズ（N=19）が統計的推論には限定的"),

    h2("3.2 類型別のスコア相関"),
    p("各クレーム類型の発生率（万室あたり）と口コミスコアの相関を分析しました。"),
    corrTable,
    p(`「汚れ」が最も高い相関（|r|=${fmtN(typCorrs[0]?.abs_r, 4)}）を示しており、宿泊者の口コミ評価に最も敏感に反映される類型と推察されます。`, { sb: 200 }),

    h3("分析上の注意"),
    bp("クレーム率が極低水準のため、相関係数は参考値として解釈が必要"),
    bp("正の相関は「検知能力の高いホテル＝高品質」という間接的因果で説明可能"),
    bp("今後のデータ蓄積により、より精緻な分析が可能になる見込み"),
    PB(),
  ];
}

function buildHotelProfiles() {
  const withClaims = hotelProfiles.filter(h => h.total_claims > 0);
  const zeroClaims = hotelProfiles.filter(h => h.total_claims === 0);

  // Hotel table with claim rates
  const hHeader = new TableRow({
    children: [
      cell("ホテル名", { bg: C.NAVY, c: C.WHITE, b: true, w: 3200 }),
      cell("クレーム数", { bg: C.NAVY, c: C.WHITE, b: true, w: 1200 }),
      cell("清掃室数", { bg: C.NAVY, c: C.WHITE, b: true, w: 1400 }),
      cell("万室あたり", { bg: C.NAVY, c: C.WHITE, b: true, w: 1400 }),
      cell("主要類型", { bg: C.NAVY, c: C.WHITE, b: true, w: 1600 }),
    ]
  });
  const hRows = hotelProfiles.map((h, i) => {
    const rateColor = h.claim_rate > 10 ? C.RED : h.claim_rate > 5 ? C.ORANGE : C.GREEN;
    const topType = h.top_types && h.top_types.length > 0 ? h.top_types[0].type : "-";
    return new TableRow({
      children: [
        cell(h.name, { bg: i % 2 ? C.WHITE : C.LIGHT_BG, a: AlignmentType.LEFT, sz: 14 }),
        cell(h.total_claims, { bg: i % 2 ? C.WHITE : C.LIGHT_BG, b: true }),
        cell(h.rooms_cleaned ? h.rooms_cleaned.toLocaleString() : "-", { bg: i % 2 ? C.WHITE : C.LIGHT_BG }),
        cell(fmtN(h.claim_rate, 1), { bg: i % 2 ? C.WHITE : C.LIGHT_BG, c: rateColor, b: true }),
        cell(topType, { bg: i % 2 ? C.WHITE : C.LIGHT_BG, sz: 14 }),
      ]
    });
  });
  const hTable = new Table({ rows: [hHeader, ...hRows], width: { size: 8800, type: WidthType.DXA } });

  const sections = [
    h1("4. ホテル別クレームプロファイル"),
    h2("4.1 ホテル別一覧"),
    p(`全19ホテルのクレーム発生状況を万室あたりの発生率で比較しました。`),
    hTable,
    p(`クレーム率の最高は${stats.max_claim_rate_hotel}（${fmtN(stats.max_claim_rate, 1)}件/万室）、平均は${fmtN(stats.avg_claim_rate, 1)}件/万室です。`, { sb: 200 }),
  ];

  // Zero claim hotels
  if (zeroClaims.length > 0) {
    sections.push(h2("4.2 ゼロクレーム達成ホテル"));
    sections.push(p(`以下の${zeroClaims.length}ホテルが期間中クレームゼロを達成しています。`));
    zeroClaims.forEach(h => {
      sections.push(bp(`${h.name}（清掃${h.rooms_cleaned ? h.rooms_cleaned.toLocaleString() : "-"}室）`, { b: true }));
    });
    sections.push(p("これらのホテルのオペレーション手法（チェック体制・研修内容・スタッフ配置等）を体系的に調査し、他ホテルへの横展開を推奨します。"));
  }

  // Top claim hotels detail
  sections.push(h2("4.3 要注意ホテルの類型分析"));
  const topHotels = withClaims.slice(0, 5);
  topHotels.forEach(h => {
    if (h.top_types && h.top_types.length > 0) {
      const typesStr = h.top_types.map(t => `${t.type}(${t.count}件/${fmtPct(t.share_pct)})`).join("、");
      sections.push(bp(`${h.name}: ${typesStr} — ${h.profile_label || h.dominant_category + "集中型"}`, { b: false }));
    }
  });
  sections.push(PB());

  return sections;
}

function buildRevenueImpact() {
  const scenarios = revImpact.scenarios;
  const regression = revImpact.regression;
  const typeImpacts = revImpact.type_impact_ranking.filter(t => t.current_count > 0);

  // Scenario table
  const sHeader = new TableRow({
    children: [
      cell("削減率", { bg: C.ACCENT, c: C.WHITE, b: true, w: 1600 }),
      cell("月間収益インパクト", { bg: C.ACCENT, c: C.WHITE, b: true, w: 2800 }),
      cell("年間収益インパクト", { bg: C.ACCENT, c: C.WHITE, b: true, w: 2800 }),
      cell("実現難易度", { bg: C.ACCENT, c: C.WHITE, b: true, w: 1600 }),
    ]
  });
  const sRows = scenarios.map((s, i) => {
    const diff = s.reduction_pct <= 25 ? "低" : s.reduction_pct <= 50 ? "中" : "高";
    const diffColor = diff === "低" ? C.GREEN : diff === "中" ? C.ORANGE : C.RED;
    return new TableRow({
      children: [
        cell(`${s.reduction_pct}%削減`, { bg: i % 2 ? C.WHITE : C.LIGHT_BG, b: true }),
        cell(fmtY(s.total_monthly_impact), { bg: i % 2 ? C.WHITE : C.LIGHT_BG, b: true, c: C.GREEN }),
        cell(fmtY(s.total_annual_impact), { bg: i % 2 ? C.WHITE : C.LIGHT_BG, b: true, c: C.GREEN }),
        cell(diff, { bg: i % 2 ? C.WHITE : C.LIGHT_BG, c: diffColor }),
      ]
    });
  });
  const sTable = new Table({ rows: [sHeader, ...sRows], width: { size: 8800, type: WidthType.DXA } });

  // Type impact table
  const tHeader = new TableRow({
    children: [
      cell("クレーム類型", { bg: C.NAVY, c: C.WHITE, b: true, w: 2400 }),
      cell("現在件数", { bg: C.NAVY, c: C.WHITE, b: true, w: 1400 }),
      cell("50%削減後", { bg: C.NAVY, c: C.WHITE, b: true, w: 1400 }),
      cell("スコア改善", { bg: C.NAVY, c: C.WHITE, b: true, w: 1600 }),
      cell("月間収益効果", { bg: C.NAVY, c: C.WHITE, b: true, w: 2000 }),
    ]
  });
  const tRows = typeImpacts.map((t, i) => {
    return new TableRow({
      children: [
        cell(t.type, { bg: i % 2 ? C.WHITE : C.LIGHT_BG, a: AlignmentType.LEFT }),
        cell(t.current_count, { bg: i % 2 ? C.WHITE : C.LIGHT_BG }),
        cell(t.reduced_count, { bg: i % 2 ? C.WHITE : C.LIGHT_BG }),
        cell(t.score_improvement > 0 ? `+${fmtN(t.score_improvement, 4)}` : "-", { bg: i % 2 ? C.WHITE : C.LIGHT_BG }),
        cell(t.revenue_impact_monthly > 0 ? fmtY(t.revenue_impact_monthly) : "-", { bg: i % 2 ? C.WHITE : C.LIGHT_BG, c: C.GREEN }),
      ]
    });
  });
  const tTable = new Table({ rows: [tHeader, ...tRows], width: { size: 8800, type: WidthType.DXA } });

  return [
    h1("5. 収益インパクト試算"),
    h2("5.1 クレーム削減シナリオ"),
    p(`分析6で算出した品質→売上弾力性（${revImpact.elasticity_used.value}）を基に、クレーム削減が収益に与える影響を試算しました。`),
    sTable,
    p(`現実的な目標として「50%削減」を目指した場合、年間約${fmtY(scenarios[1]?.total_annual_impact)}の収益改善が見込まれます。`, { sb: 200 }),

    h2("5.2 類型別の収益インパクト"),
    p("各クレーム類型を50%削減した場合の個別収益効果です。"),
    tTable,
    p(`最も収益インパクトが大きいのは上位の頻出類型（セット漏れ・残置）であり、これらの集中改善が費用対効果の観点で最も有効です。`, { sb: 200 }),

    h3("前提条件"),
    bp(`回帰式: スコア = ${fmtN(regression.slope, 2)} × クレーム率 + ${fmtN(regression.intercept, 2)}`),
    bp(`弾力性: ${revImpact.elasticity_used.value}（分析6より）`),
    bp("試算は現時点のデータに基づく推定値であり、実際の効果は施策実行内容に依存"),
    PB(),
  ];
}

function buildPriorities() {
  const top5 = priorities.slice(0, Math.min(8, priorities.length));

  const pHeader = new TableRow({
    children: [
      cell("順位", { bg: C.NAVY, c: C.WHITE, b: true, w: 800 }),
      cell("クレーム類型", { bg: C.NAVY, c: C.WHITE, b: true, w: 2000 }),
      cell("件数", { bg: C.NAVY, c: C.WHITE, b: true, w: 1000 }),
      cell("影響ホテル", { bg: C.NAVY, c: C.WHITE, b: true, w: 1200 }),
      cell("相関|r|", { bg: C.NAVY, c: C.WHITE, b: true, w: 1200 }),
      cell("複合スコア", { bg: C.NAVY, c: C.WHITE, b: true, w: 1200 }),
      cell("優先度", { bg: C.NAVY, c: C.WHITE, b: true, w: 1200 }),
    ]
  });
  const pRows = top5.map((pr, i) => {
    const priority = i < 2 ? "最優先" : i < 4 ? "高" : "中";
    const prColor = priority === "最優先" ? C.RED : priority === "高" ? C.ORANGE : C.TEXT;
    return new TableRow({
      children: [
        cell(pr.priority_rank, { bg: i % 2 ? C.WHITE : C.LIGHT_BG, b: true }),
        cell(pr.type, { bg: i % 2 ? C.WHITE : C.LIGHT_BG, a: AlignmentType.LEFT }),
        cell(pr.total_count, { bg: i % 2 ? C.WHITE : C.LIGHT_BG }),
        cell(pr.hotels_affected, { bg: i % 2 ? C.WHITE : C.LIGHT_BG }),
        cell(fmtN(pr.abs_correlation, 3), { bg: i % 2 ? C.WHITE : C.LIGHT_BG }),
        cell(fmtN(pr.composite_priority, 3), { bg: i % 2 ? C.WHITE : C.LIGHT_BG, b: true }),
        cell(priority, { bg: i % 2 ? C.WHITE : C.LIGHT_BG, c: prColor, b: true }),
      ]
    });
  });
  const pTable = new Table({ rows: [pHeader, ...pRows], width: { size: 8600, type: WidthType.DXA } });

  return [
    h1("6. 改善優先度ランキング"),
    p("「頻度（40%）」「スコア影響度（40%）」「波及度（20%）」の3軸で複合スコアを算出し、改善の優先順位を決定しました。"),
    pTable,
    p(""),
    bp(`第1優先: ${top5[0]?.type ?? "-"} — 発生頻度が高く、スコアへの影響も大きいため、最優先で対策を講じるべき類型`, { b: true }),
    top5.length > 1 ? bp(`第2優先: ${top5[1]?.type ?? "-"} — スコアとの相関が比較的高く、改善による効果が期待できる`, { b: true }) : p(""),
    PB(),
  ];
}

function buildRecommendations() {
  const sections = [
    h1("7. 提言とアクションプラン"),
  ];

  recs.forEach((rec, i) => {
    sections.push(h2(`7.${i + 1} ${rec.title}`));
    sections.push(p(`【優先度: ${rec.priority}】`, { b: true, c: rec.priority === "最優先" ? C.RED : C.ORANGE }));
    sections.push(p(rec.rationale));
    sections.push(h3("具体的アクション"));
    rec.actions.filter(a => a).forEach(a => {
      sections.push(bp(a));
    });
  });

  // Implementation timeline
  sections.push(h2("7.5 実施ロードマップ"));
  sections.push(bp("【Phase 1（1-2ヶ月）】上位2類型（残置・セット漏れ）のチェックリスト強化", { b: true }));
  sections.push(bp("【Phase 2（3-4ヶ月）】清潔性系クレーム（汚れ・髪の毛）対策の研修実施", { b: true }));
  sections.push(bp("【Phase 3（5-6ヶ月）】ゼロクレームホテルのベストプラクティス横展開", { b: true }));
  sections.push(bp("【Phase 4（7-12ヶ月）】全ホテル統一のクレーム予防体制確立と効果測定", { b: true }));

  return sections;
}

// ============================================================
// Build DOCX
// ============================================================
async function buildDOCX() {
  const doc = new Document({
    styles: { default: { document: { run: { font: "Arial", size: 20 } } } },
    sections: [{
      headers: {
        default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "PRIMECHANGE | 分析1: クレーム類型×スコア連動分析", size: 14, color: C.SUBTEXT, font: "Arial" })] })] })
      },
      footers: {
        default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Page ", size: 14, color: C.SUBTEXT }), new TextRun({ children: [PageNumber.CURRENT], size: 14, color: C.SUBTEXT })] })] })
      },
      children: [
        ...buildCover(),
        ...buildExecutiveSummary(),
        ...buildTypeFrequency(),
        ...buildCorrelation(),
        ...buildHotelProfiles(),
        ...buildRevenueImpact(),
        ...buildPriorities(),
        ...buildRecommendations(),
      ]
    }]
  });

  const buf = await Packer.toBuffer(doc);
  const outPath = path.resolve(__dirname, "PRIMECHANGE_分析1_クレーム類型スコア連動.docx");
  fs.writeFileSync(outPath, buf);
  console.log(`✅ DOCX: ${outPath} (${(buf.length / 1024).toFixed(1)} KB)`);
}

// ============================================================
// Build PPTX
// ============================================================
function buildPPTX() {
  const pptx = new pptxgen();
  pptx.layout = "LAYOUT_16x9";
  pptx.author = "PRIMECHANGE";
  pptx.subject = "分析1: クレーム類型×スコア連動分析";

  const bgDark = { fill: C.NAVY };
  const bgLight = { fill: C.WHITE };
  const titleOpts = { x: 0.5, y: 0.3, w: 9, h: 0.6, fontSize: 22, bold: true, color: C.NAVY, fontFace: "Arial" };
  const subtitleOpts = { x: 0.5, y: 0.9, w: 9, h: 0.4, fontSize: 13, color: C.SUBTEXT, fontFace: "Arial" };

  // Slide 1: Title
  const s1 = pptx.addSlide();
  s1.background = bgDark;
  s1.addText("PRIMECHANGE", { x: 0.5, y: 1.2, w: 9, h: 0.8, fontSize: 40, bold: true, color: C.WHITE, fontFace: "Arial", align: "center" });
  s1.addText("分析1: クレーム類型×口コミスコア連動分析", { x: 0.5, y: 2.2, w: 9, h: 0.6, fontSize: 24, bold: true, color: C.WHITE, fontFace: "Arial", align: "center" });
  s1.addText("Claim Type × Review Score Correlation Analysis", { x: 0.5, y: 2.9, w: 9, h: 0.4, fontSize: 14, color: "AAAAAA", fontFace: "Arial", italic: true, align: "center" });
  s1.addText(`対象: ${meta.total_hotels}ホテル | 総クレーム: ${meta.total_claims}件 | 13類型分析`, { x: 0.5, y: 3.8, w: 9, h: 0.3, fontSize: 12, color: "CCCCCC", fontFace: "Arial", align: "center" });
  s1.addText("2025年3月", { x: 0.5, y: 4.3, w: 9, h: 0.3, fontSize: 14, color: C.WHITE, fontFace: "Arial", align: "center" });

  // Slide 2: Key Findings (KPI Cards)
  const s2 = pptx.addSlide();
  s2.background = bgLight;
  s2.addText("主要な発見", { ...titleOpts });
  s2.addText("19ホテル×13クレーム類型の統合分析から得られた重要な洞察", { ...subtitleOpts });

  const kpis = [
    { label: "総クレーム件数", value: `${meta.total_claims}件`, sub: `${meta.total_rooms_cleaned.toLocaleString()}室中`, color: C.ACCENT },
    { label: "万室あたりクレーム", value: `${fmtN(stats.avg_claim_rate, 1)}件`, sub: "19ホテル平均", color: C.GREEN },
    { label: "最多類型", value: `セット漏れ`, sub: `${typeRanking[0].share_pct}% (${typeRanking[0].total_count}件)`, color: C.ORANGE },
    { label: "ゼロクレーム", value: `${stats.hotels_zero_claims}ホテル`, sub: "達成率 15.8%", color: C.TEAL },
  ];

  kpis.forEach((kpi, i) => {
    const x = 0.3 + i * 2.35;
    s2.addShape(pptx.ShapeType.roundRect, { x, y: 1.6, w: 2.15, h: 1.6, fill: { color: kpi.color }, rectRadius: 0.1 });
    s2.addText(kpi.label, { x, y: 1.7, w: 2.15, h: 0.3, fontSize: 10, color: C.WHITE, fontFace: "Arial", align: "center" });
    s2.addText(kpi.value, { x, y: 2.0, w: 2.15, h: 0.6, fontSize: 26, bold: true, color: C.WHITE, fontFace: "Arial", align: "center" });
    s2.addText(kpi.sub, { x, y: 2.7, w: 2.15, h: 0.3, fontSize: 9, color: C.WHITE, fontFace: "Arial", align: "center" });
  });

  // Bottom insights
  const insights = [
    `「客室準備系」クレームが全体の${catBreakdown["客室準備系"].share_pct}%を占有`,
    `上位2類型（セット漏れ・残置）で全体の55.9%`,
    `クレーム50%削減で年間約${fmtY(revImpact.scenarios[1]?.total_annual_impact)}の改善ポテンシャル`
  ];
  insights.forEach((ins, i) => {
    s2.addText(`• ${ins}`, { x: 0.5, y: 3.6 + i * 0.35, w: 9, h: 0.35, fontSize: 11, color: C.TEXT, fontFace: "Arial" });
  });

  // Slide 3: Type Frequency Ranking
  const s3 = pptx.addSlide();
  s3.background = bgLight;
  s3.addText("クレーム類型別頻度ランキング", { ...titleOpts });

  const activeTypes = typeRanking.filter(t => t.total_count > 0);
  const freqRows = [
    [{ text: "類型", options: { bold: true, color: C.WHITE, fill: { color: C.NAVY }, fontSize: 9 } },
     { text: "件数", options: { bold: true, color: C.WHITE, fill: { color: C.NAVY }, fontSize: 9 } },
     { text: "構成比", options: { bold: true, color: C.WHITE, fill: { color: C.NAVY }, fontSize: 9 } },
     { text: "影響ホテル", options: { bold: true, color: C.WHITE, fill: { color: C.NAVY }, fontSize: 9 } }],
    ...activeTypes.map((t, i) => [
      { text: t.type, options: { fontSize: 9, fill: { color: i % 2 ? "FFFFFF" : "F5F7FA" } } },
      { text: String(t.total_count), options: { fontSize: 9, bold: true, fill: { color: i % 2 ? "FFFFFF" : "F5F7FA" } } },
      { text: fmtPct(t.share_pct), options: { fontSize: 9, fill: { color: i % 2 ? "FFFFFF" : "F5F7FA" } } },
      { text: `${t.hotels_affected}/19`, options: { fontSize: 9, fill: { color: i % 2 ? "FFFFFF" : "F5F7FA" } } },
    ])
  ];
  s3.addTable(freqRows, { x: 0.3, y: 1.2, w: 4.5, h: 3.5, fontSize: 9, fontFace: "Arial", border: { color: "CCCCCC", pt: 0.5 }, colW: [1.5, 0.8, 1.0, 1.2], align: "center" });

  // Category breakdown on right
  const cats = Object.entries(catBreakdown).sort((a, b) => b[1].count - a[1].count);
  const catColors = [C.ACCENT, C.ORANGE, C.TEAL, C.SUBTEXT];
  cats.forEach((c, i) => {
    const y = 1.4 + i * 0.9;
    s3.addShape(pptx.ShapeType.roundRect, { x: 5.3, y, w: 4.3, h: 0.75, fill: { color: catColors[i] || C.SUBTEXT }, rectRadius: 0.05 });
    s3.addText(`${c[0]}  ${c[1].count}件 (${fmtPct(c[1].share_pct)})`, { x: 5.5, y: y + 0.05, w: 3.9, h: 0.35, fontSize: 12, bold: true, color: C.WHITE, fontFace: "Arial" });
    s3.addText(c[1].types.join("、"), { x: 5.5, y: y + 0.4, w: 3.9, h: 0.25, fontSize: 8, color: C.WHITE, fontFace: "Arial" });
  });

  // Slide 4: Correlation Analysis
  const s4 = pptx.addSlide();
  s4.background = bgLight;
  s4.addText("クレーム率×スコア相関分析", { ...titleOpts });
  s4.addText(`全体相関: r=${fmtN(corrAnalysis.overall_correlation.r, 4)} (${corrAnalysis.overall_correlation.interpretation})`, { ...subtitleOpts });

  const topCorrs = corrAnalysis.type_correlations.filter(tc => tc.abs_r > 0).slice(0, 8);
  const corrTblRows = [
    [{ text: "類型", options: { bold: true, color: C.WHITE, fill: { color: C.NAVY }, fontSize: 9 } },
     { text: "相関 r", options: { bold: true, color: C.WHITE, fill: { color: C.NAVY }, fontSize: 9 } },
     { text: "|r|", options: { bold: true, color: C.WHITE, fill: { color: C.NAVY }, fontSize: 9 } },
     { text: "判定", options: { bold: true, color: C.WHITE, fill: { color: C.NAVY }, fontSize: 9 } }],
    ...topCorrs.map((tc, i) => [
      { text: tc.type, options: { fontSize: 9, fill: { color: i % 2 ? "FFFFFF" : "F5F7FA" } } },
      { text: fmtN(tc.correlation_r, 4), options: { fontSize: 9, fill: { color: i % 2 ? "FFFFFF" : "F5F7FA" } } },
      { text: fmtN(tc.abs_r, 4), options: { fontSize: 9, bold: true, fill: { color: i % 2 ? "FFFFFF" : "F5F7FA" } } },
      { text: tc.interpretation, options: { fontSize: 9, color: tc.abs_r >= 0.4 ? C.RED : tc.abs_r >= 0.2 ? C.ORANGE : C.TEXT, fill: { color: i % 2 ? "FFFFFF" : "F5F7FA" } } },
    ])
  ];
  s4.addTable(corrTblRows, { x: 0.3, y: 1.4, w: 5.0, h: 3.0, fontSize: 9, fontFace: "Arial", border: { color: "CCCCCC", pt: 0.5 }, colW: [1.5, 0.9, 0.8, 1.2], align: "center" });

  // Interpretation box
  s4.addShape(pptx.ShapeType.roundRect, { x: 5.6, y: 1.4, w: 4.0, h: 2.5, fill: { color: "F0F4F8" }, rectRadius: 0.1, line: { color: C.ACCENT, width: 1 } });
  s4.addText("分析上の考察", { x: 5.8, y: 1.5, w: 3.6, h: 0.35, fontSize: 12, bold: true, color: C.NAVY, fontFace: "Arial" });
  const considerations = [
    "• 全ホテルのクレーム率が極めて低水準（0〜0.15%）",
    "• 高品質ホテルほど検知力が高い可能性",
    "• 「汚れ」が最高相関（r=0.42）",
    "• データ蓄積で分析精度向上が見込まれる",
    "• N=19のため参考値として活用"
  ];
  considerations.forEach((c, i) => {
    s4.addText(c, { x: 5.8, y: 1.9 + i * 0.35, w: 3.6, h: 0.3, fontSize: 9, color: C.TEXT, fontFace: "Arial" });
  });

  // Slide 5: Hotel Profiles
  const s5 = pptx.addSlide();
  s5.background = bgLight;
  s5.addText("ホテル別クレームプロファイル", { ...titleOpts });

  const top10Hotels = hotelProfiles.slice(0, 10);
  const hotelTblRows = [
    [{ text: "ホテル名", options: { bold: true, color: C.WHITE, fill: { color: C.NAVY }, fontSize: 8 } },
     { text: "件数", options: { bold: true, color: C.WHITE, fill: { color: C.NAVY }, fontSize: 8 } },
     { text: "万室あたり", options: { bold: true, color: C.WHITE, fill: { color: C.NAVY }, fontSize: 8 } },
     { text: "主要類型", options: { bold: true, color: C.WHITE, fill: { color: C.NAVY }, fontSize: 8 } }],
    ...top10Hotels.map((h, i) => {
      const topType = h.top_types && h.top_types.length > 0 ? h.top_types[0].type : "-";
      const rateColor = h.claim_rate > 10 ? C.RED : h.claim_rate > 5 ? C.ORANGE : h.total_claims === 0 ? C.GREEN : C.TEXT;
      return [
        { text: h.name.replace(/ホテル/g, 'H.'), options: { fontSize: 7, fill: { color: i % 2 ? "FFFFFF" : "F5F7FA" } } },
        { text: String(h.total_claims), options: { fontSize: 8, bold: true, fill: { color: i % 2 ? "FFFFFF" : "F5F7FA" } } },
        { text: fmtN(h.claim_rate, 1), options: { fontSize: 8, bold: true, color: rateColor, fill: { color: i % 2 ? "FFFFFF" : "F5F7FA" } } },
        { text: topType, options: { fontSize: 7, fill: { color: i % 2 ? "FFFFFF" : "F5F7FA" } } },
      ];
    })
  ];
  s5.addTable(hotelTblRows, { x: 0.3, y: 1.2, w: 5.5, h: 3.5, fontSize: 8, fontFace: "Arial", border: { color: "CCCCCC", pt: 0.5 }, colW: [2.2, 0.7, 1.0, 1.2], align: "center" });

  // Zero claims highlight
  const zeroHotels = hotelProfiles.filter(h => h.total_claims === 0);
  s5.addShape(pptx.ShapeType.roundRect, { x: 6.0, y: 1.2, w: 3.6, h: 1.8, fill: { color: C.GREEN }, rectRadius: 0.1 });
  s5.addText("🏆 ゼロクレーム達成", { x: 6.2, y: 1.3, w: 3.2, h: 0.35, fontSize: 14, bold: true, color: C.WHITE, fontFace: "Arial", align: "center" });
  zeroHotels.forEach((h, i) => {
    s5.addText(`• ${h.name}`, { x: 6.3, y: 1.75 + i * 0.3, w: 3.0, h: 0.28, fontSize: 10, color: C.WHITE, fontFace: "Arial" });
  });
  s5.addText("→ 成功手法の横展開推奨", { x: 6.3, y: 1.75 + zeroHotels.length * 0.3 + 0.1, w: 3.0, h: 0.25, fontSize: 9, bold: true, color: C.WHITE, fontFace: "Arial" });

  // Slide 6: Revenue Impact
  const s6 = pptx.addSlide();
  s6.background = bgLight;
  s6.addText("収益インパクト試算", { ...titleOpts });
  s6.addText("分析6の弾力性を活用したクレーム削減→売上改善の連鎖効果", { ...subtitleOpts });

  const scenarioData = revImpact.scenarios;
  const scenColors = [C.GREEN, C.ACCENT, C.ORANGE];
  scenarioData.forEach((s, i) => {
    const x = 0.5 + i * 3.1;
    s6.addShape(pptx.ShapeType.roundRect, { x, y: 1.5, w: 2.8, h: 2.0, fill: { color: scenColors[i] }, rectRadius: 0.1 });
    s6.addText(`${s.reduction_pct}%削減`, { x, y: 1.6, w: 2.8, h: 0.3, fontSize: 14, bold: true, color: C.WHITE, fontFace: "Arial", align: "center" });
    s6.addText(`月間 ${fmtY(s.total_monthly_impact)}`, { x, y: 2.0, w: 2.8, h: 0.4, fontSize: 11, color: C.WHITE, fontFace: "Arial", align: "center" });
    s6.addText(`年間 ${fmtY(s.total_annual_impact)}`, { x, y: 2.5, w: 2.8, h: 0.5, fontSize: 20, bold: true, color: C.WHITE, fontFace: "Arial", align: "center" });
  });

  // Bottom: priority types
  const topImpacts = revImpact.type_impact_ranking.filter(t => t.current_count > 0).slice(0, 5);
  s6.addText("類型別インパクト（50%削減時）:", { x: 0.5, y: 3.8, w: 9, h: 0.3, fontSize: 11, bold: true, color: C.NAVY, fontFace: "Arial" });
  topImpacts.forEach((t, i) => {
    s6.addText(`${t.type}: ${t.current_count}件→${t.reduced_count}件 ${t.revenue_impact_monthly > 0 ? fmtY(t.revenue_impact_monthly) + "/月" : ""}`, { x: 0.5 + (i % 3) * 3.2, y: 4.2 + Math.floor(i / 3) * 0.3, w: 3.0, h: 0.28, fontSize: 9, color: C.TEXT, fontFace: "Arial" });
  });

  // Slide 7: Improvement Priorities
  const s7 = pptx.addSlide();
  s7.background = bgLight;
  s7.addText("改善優先度と提言", { ...titleOpts });

  // Priority ranking
  const topPriorities = priorities.slice(0, 5);
  const prioColors = [C.RED, C.ORANGE, C.AMBER, C.ACCENT, C.TEAL];
  topPriorities.forEach((pr, i) => {
    const y = 1.3 + i * 0.55;
    s7.addShape(pptx.ShapeType.roundRect, { x: 0.3, y, w: 4.5, h: 0.45, fill: { color: prioColors[i] || C.SUBTEXT }, rectRadius: 0.05 });
    s7.addText(`#${pr.priority_rank} ${pr.type}  (${pr.total_count}件, |r|=${fmtN(pr.abs_correlation, 3)})`, { x: 0.5, y: y + 0.05, w: 4.1, h: 0.35, fontSize: 11, bold: true, color: C.WHITE, fontFace: "Arial" });
  });

  // Action plan
  s7.addShape(pptx.ShapeType.roundRect, { x: 5.2, y: 1.3, w: 4.4, h: 3.5, fill: { color: "F0F4F8" }, rectRadius: 0.1, line: { color: C.NAVY, width: 1 } });
  s7.addText("実施ロードマップ", { x: 5.4, y: 1.4, w: 4.0, h: 0.35, fontSize: 13, bold: true, color: C.NAVY, fontFace: "Arial" });

  const phases = [
    { phase: "Phase 1 (1-2月)", action: "残置・セット漏れ対策\nチェックリスト強化", color: C.RED },
    { phase: "Phase 2 (3-4月)", action: "清潔性系対策\n研修プログラム導入", color: C.ORANGE },
    { phase: "Phase 3 (5-6月)", action: "ゼロクレームホテル\nベストプラクティス横展開", color: C.GREEN },
    { phase: "Phase 4 (7-12月)", action: "統一予防体制確立\n効果測定・PDCA", color: C.TEAL },
  ];
  phases.forEach((ph, i) => {
    const y = 1.85 + i * 0.72;
    s7.addShape(pptx.ShapeType.roundRect, { x: 5.4, y, w: 1.6, h: 0.6, fill: { color: ph.color }, rectRadius: 0.05 });
    s7.addText(ph.phase, { x: 5.45, y: y + 0.05, w: 1.5, h: 0.5, fontSize: 8, bold: true, color: C.WHITE, fontFace: "Arial", align: "center", valign: "middle" });
    s7.addText(ph.action, { x: 7.1, y, w: 2.3, h: 0.6, fontSize: 9, color: C.TEXT, fontFace: "Arial", valign: "middle" });
  });

  const outPath = path.resolve(__dirname, "PRIMECHANGE_分析1_クレーム類型スコア連動.pptx");
  pptx.writeFile({ fileName: outPath }).then(() => {
    console.log(`✅ PPTX: ${outPath} (${(fs.statSync(outPath).size / 1024).toFixed(1)} KB)`);
  });
}

// ============================================================
// Main
// ============================================================
async function main() {
  console.log("Building Analysis 1 Reports...");
  await buildDOCX();
  await buildPPTX();
  console.log("\nDone!");
}

main().catch(e => { console.error(e); process.exit(1); });
