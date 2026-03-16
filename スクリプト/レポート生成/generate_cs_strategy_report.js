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
// Data Loading
// ============================================================
const qualityData = JSON.parse(fs.readFileSync(path.resolve(__dirname, "primechange_portfolio_analysis.json"), "utf-8"));
const revenueData = JSON.parse(fs.readFileSync(path.resolve(__dirname, "hotel_revenue_data.json"), "utf-8"));
const OUTPUT = path.resolve(__dirname, "PRIMECHANGE_CS向上戦略提案書.docx");

const overview = qualityData.portfolio_overview;
const deepDive = qualityData.cleaning_deep_dive || {};
const priorityMatrix = qualityData.priority_matrix || {};
const hotelsRanked = overview.hotels_ranked || [];

const KEY_MAP = { keisei_kinshicho: "keisei_richmond", comfort_yokohama_kannai: "comfort_yokohama" };
function getRev(qKey) { return revenueData[KEY_MAP[qKey] || qKey] || {}; }

// Integrated hotel list
const hotels = hotelsRanked.map(q => {
  const r = getRev(q.key);
  return { ...q, revenue: r.actual_revenue || 0, occupancy: r.occupancy_rate || 0, profit_rate: r.profit_rate || 0, net_profit: r.actual_net_profit || 0, adr: r.adr || 0, room_count: r.room_count || 0, staff_count: r.staff_count || 0, phase: r.phase || "" };
});

const totalRevenue = hotels.reduce((s, h) => s + h.revenue, 0);
const medianRev = [...hotels].sort((a, b) => a.revenue - b.revenue)[Math.floor(hotels.length / 2)].revenue;
const qualityThreshold = 8.0;

hotels.forEach(h => {
  const hq = h.avg >= qualityThreshold, hr = h.revenue >= medianRev;
  h.quadrant = hq && hr ? "高品質×高売上" : hq && !hr ? "高品質×低売上" : !hq && hr ? "低品質×高売上" : "低品質×低売上";
});

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
function fmtP(n) { return n ? (Number(n) * 100).toFixed(1) + "%" : "-"; }
const PB = () => new Paragraph({ children: [new PageBreak()] });

// ============================================================
// Section 1: Cover
// ============================================================
function buildCover() {
  return [
    new Paragraph({ spacing: { before: 2400 } }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 }, children: [new TextRun({ text: "PRIMECHANGE", size: 56, font: "Arial", bold: true, color: C.NAVY })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [new TextRun({ text: "CS向上 × 売上増加", size: 44, font: "Arial", bold: true, color: C.ACCENT })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [new TextRun({ text: "戦略提案書", size: 40, font: "Arial", bold: true, color: C.ACCENT })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 600 }, children: [new TextRun({ text: "Strategic Proposal for Customer Satisfaction & Revenue Growth", size: 22, font: "Arial", color: C.SUBTEXT, italics: true })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [new TextRun({ text: `対象: ${hotels.length}ホテル | 月間総売上: ${fmtY(totalRevenue)} | 口コミ: ${overview.total_reviews}件`, size: 20, font: "Arial", color: C.TEXT })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [new TextRun({ text: "2026年3月", size: 22, font: "Arial", color: C.TEXT })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 800 }, children: [new TextRun({ text: "株式会社PRIMECHANGE", size: 28, font: "Arial", bold: true, color: C.NAVY })] }),
    PB(),
  ];
}

// ============================================================
// Section 2: Executive Summary
// ============================================================
function buildExecSummary() {
  const urgentHotels = (priorityMatrix.URGENT || []).map(h => h.name);
  return [
    h1("1. エグゼクティブサマリー"),
    p("本提案書は、PRIMECHANGEが管理する19ホテルの「CS（顧客満足度）向上」を起点に「売上増加」を実現するための戦略フレームワークを提示します。"),
    h3("CS向上 → 売上増加の因果メカニズム"),
    bp("清掃品質向上 → クレーム削減 → 口コミスコア改善", { b: true }),
    bp("口コミスコア改善 → OTAランキング上昇 → 稼働率向上 → ホテル売上増加", { level: 1 }),
    bp("ホテル売上増加 → PRIMECHANGEの契約維持・単価交渉力向上・新規受注", { level: 1 }),
    p("業界データでは口コミスコア0.1点改善がRevPAR（客室あたり収益）を約1%向上させるとされています。当社19ホテルの月間総売上約1億円に対し、ポートフォリオ平均0.5点改善で月間300-500万円の売上改善余地があると推定します。", { sb: 120 }),
    h3("7つの分析コンセプト"),
    p("月次レポートXLSX（19ホテル×22シート）に眠る未活用データを活用し、以下7テーマの分析を提案します:"),
    bp("分析1: クレーム類型×口コミスコア連動分析 — 最もスコアに効くクレーム類型を特定"),
    bp("分析2: スタッフ個人別パフォーマンス分析 — 人の問題を可視化・解決"),
    bp("分析3: 人員配置×品質相関分析 — 最適配置基準を策定"),
    bp("分析4: 清掃完了時間×品質分析 — 時間プレッシャーの影響を定量化"),
    bp("分析5: 安全チェック×クレーム予兆検出 — 問題発生前の予防"),
    bp("分析6: 品質→売上弾力性分析 — 自社データでROIを精密化"),
    bp("分析7: ベストプラクティス横展開 — 成功モデルの体系化"),
    h3("緊急アクション"),
    p(`現在URGENT判定の4ホテル（${urgentHotels.join("、")}）は合計月間売上約1,973万円。品質悪化による売上減少リスクが最も高く、最優先で対応が必要です。`, { c: C.RED }),
    PB(),
  ];
}

// ============================================================
// Section 3: Current State Dashboard
// ============================================================
function buildCurrentState() {
  const el = [
    h1("2. 現状分析：データが語る19ホテルの実態"),
    h2("2.1 ポートフォリオ全体KPI"),
  ];

  const kpiData = [
    ["指標", "現状値", "備考"],
    ["管理ホテル数", `${hotels.length}ホテル`, "P1:7, P2:5, P3:6, P4:1"],
    ["月間総売上（2月）", fmtY(totalRevenue), "最大:スイーツ東京ベイ¥13.2M"],
    ["平均口コミスコア", `${overview.avg_score} / 10.0`, "中央値: " + overview.median_score],
    ["総レビュー件数", `${overview.total_reviews}件`, "6つのOTAサイトから収集"],
    ["平均稼働率", fmtP(hotels.reduce((s, h) => s + h.occupancy, 0) / hotels.length), "最低:横浜関内55.8%"],
    ["清掃クレーム率", `${deepDive.cleaning_issue_rate || 4.6}%`, `${deepDive.total_cleaning_issues || 73}件/${overview.total_reviews}件`],
    ["URGENT判定ホテル", `${(priorityMatrix.URGENT || []).length}ホテル`, "緊急改善が必要"],
  ];

  el.push(new Table({ rows: kpiData.map((r, i) => new TableRow({
    children: r.map((c, j) => cell(c, { w: j === 0 ? 2500 : j === 1 ? 2500 : 4000, b: i === 0, bg: i === 0 ? C.NAVY : i % 2 === 0 ? C.LIGHT_BG : C.WHITE, c: i === 0 ? C.WHITE : C.TEXT, a: j === 0 ? AlignmentType.LEFT : AlignmentType.CENTER, sz: 15 }))
  })), width: { size: 9000, type: WidthType.DXA } }));

  // Quadrant summary
  el.push(p(""), h2("2.2 品質×売上マトリクス"));
  el.push(p(`品質スコア${qualityThreshold}点以上を「高品質」、月間売上の中央値（${fmtY(medianRev)}）以上を「高売上」として4象限に分類:`));

  const quadrants = [
    { name: "高品質×高売上（維持・横展開モデル）", color: C.GREEN, desc: "品質・売上ともに高水準。成功要因を他ホテルに横展開すべき。" },
    { name: "高品質×低売上（営業強化対象）", color: C.ORANGE, desc: "品質は高いが売上が伸びていない。マーケティング・価格戦略の見直しが必要。" },
    { name: "低品質×高売上（最大リスク）", color: C.RED, desc: "売上は高いが品質が低い。口コミ悪化→売上急落のリスクが最も大きい。最優先で改善。" },
    { name: "低品質×低売上（根本改革）", color: "BF360C", desc: "品質・売上ともに課題。抜本的な改革プログラムが必要。" },
  ];

  for (const q of quadrants) {
    const list = hotels.filter(h => h.quadrant === q.name.split("（")[0]);
    el.push(h3(q.name));
    el.push(p(q.desc));
    if (list.length) {
      el.push(p(list.map(h => `${h.name}（${h.avg}点, ${fmtY(h.revenue)}）`).join(" / "), { sz: 18 }));
    } else {
      el.push(p("（該当なし）", { i: true, c: C.SUBTEXT }));
    }
  }

  // Untapped data
  el.push(p(""), h2("2.3 未発掘データの宝庫"));
  el.push(p("各ホテルの月次レポートXLSXには22のシートが含まれますが、現在の分析で活用されているのはごく一部です:"));

  const dataStatus = [
    ["シート名", "データ内容", "活用状況"],
    ["①R8_*集計", "月次KPI（売上/稼働率/利益率等）", "✅ 活用済み"],
    ["💭口コミ", "各サイトの口コミテキスト", "✅ 活用済み"],
    ["🔵クレーム", "クレーム類型13種の月別集計", "❌ 未活用"],
    ["③R8_*日報", "日別クレーム数/完了時間/出勤メイド数", "❌ 未活用"],
    ["④R8_*月報", "改善進捗テキスト/スタッフ在籍・不足数", "❌ 未活用"],
    ["R8品質データまとめ", "メイド/チェッカー別クレーム頻度", "❌ 未活用"],
    ["✅安全チェック", "パトロール結果（◎/△/✖）", "❌ 未活用"],
    ["🏆皆勤アワード", "スタッフ出勤/労働時間/清掃部屋数", "❌ 未活用"],
    ["年間集計", "12ヶ月の売上/利益/稼働率推移", "❌ 未活用"],
    ["🌟目標数値", "各ポジション時給/稼働率目標", "❌ 未活用"],
    ["🧹特別清掃", "特別清掃の請求・施工管理", "❌ 未活用"],
  ];

  el.push(new Table({ rows: dataStatus.map((r, i) => new TableRow({
    children: r.map((c, j) => cell(c, { w: j === 0 ? 2500 : j === 1 ? 4500 : 2000, b: i === 0, bg: i === 0 ? C.NAVY : c.includes("❌") ? "FFF3E0" : c.includes("✅") ? "E8F5E9" : i % 2 === 0 ? C.LIGHT_BG : C.WHITE, c: i === 0 ? C.WHITE : C.TEXT, a: j === 2 ? AlignmentType.CENTER : AlignmentType.LEFT, sz: 14 }))
  })), width: { size: 9000, type: WidthType.DXA } }));

  el.push(p(""), p("上記の未活用データを分析することで、「なぜクレームが発生するのか」「どうすれば効率的に品質を改善できるのか」という根本的な問いに答えることが可能になります。", { b: true, c: C.NAVY }));
  el.push(PB());
  return el;
}

// ============================================================
// Section 4: 7 Analysis Concepts
// ============================================================
function buildAnalysisConcepts() {
  const el = [h1("3. 7つの分析コンセプト")];
  el.push(p("以下の7テーマは、月次レポートXLSXの未活用データを分析することで、CS向上→売上増加を実現するための具体的な分析提案です。各テーマの「なぜ必要か」「どのデータを使うか」「何が分かるか」「CS/売上にどう効くか」を明示します。"));

  // Analysis 1
  el.push(p(""), h2("分析1: クレーム類型×口コミスコア連動分析"));
  el.push(h3("なぜ必要か"));
  el.push(p("現在の清掃クレーム分析では「クレーム全体の件数」は把握していますが、13種のクレーム類型（髪の毛/汚れ/セット漏れ等）のどれが口コミスコア低下に最も影響するかは特定できていません。限られた改善リソースを最も効果的な類型に集中投資するためには、この「インパクト分析」が不可欠です。"));
  el.push(h3("使用データ"));
  el.push(bp("🔵クレームシート: 13類型（巻き込み/誤入室/ドア閉め忘れ/髪の毛/残置/セット漏れ/汚れ/手配ミス/清掃不備/未清掃/破損/私物破棄/その他）× 月別集計"));
  el.push(bp("口コミスコア（analysis.json）: ホテル別月次平均スコア"));
  el.push(h3("分析手法"));
  el.push(bp("各クレーム類型の月間発生率と口コミスコアの相関係数を算出"));
  el.push(bp("重回帰分析: 複数の類型を説明変数としてスコアへの影響度を分離"));
  el.push(bp("類型別の「スコア弾力性」を推定（例: 髪の毛クレーム1件 = スコア-0.05点）"));
  el.push(h3("期待される発見"));
  el.push(bp("「髪の毛」と「汚れ」のどちらがスコアに与える影響が大きいかが定量的に判明", { b: true }));
  el.push(bp("ゲストが最も不快に感じるクレーム類型の優先順位が確定"));
  el.push(bp("ホテルごとに「改善すべき類型」が異なる可能性を発見"));
  el.push(h3("CS→売上への因果ロジック"));
  el.push(p("最もスコアに効くクレーム類型に集中対策 → 効率的にスコア改善 → OTAランキング上昇 → 稼働率向上 → 売上増加"));
  el.push(p("推定効果: スコア0.3点改善で月間売上約300万円（3%）の改善余地", { b: true, c: C.GREEN }));
  el.push(h3("実装"));
  el.push(bp("難易度: 中（openpyxlでXLSXから類型別件数を抽出 → Pythonで相関分析）"));
  el.push(bp("所要期間: 1-2週間"));

  // Analysis 2
  el.push(p(""), h2("分析2: スタッフ個人別パフォーマンス分析"));
  el.push(h3("なぜ必要か"));
  el.push(p("クレームの発生原因が「仕組み」の問題なのか「個人」の問題なのかが区別できていません。R8品質データまとめシートにはメイド/チェッカー別のクレーム頻度データが存在しますが、未分析です。個人レベルのパフォーマンス差を可視化することで、「誰に何を教えるべきか」が明確になります。"));
  el.push(h3("使用データ"));
  el.push(bp("R8品質データまとめシート: メイド別クレーム発生件数、チェッカー別クレーム発生件数"));
  el.push(bp("🏆皆勤アワードシート: スタッフ名/ポジション/労働時間/年間清掃部屋数"));
  el.push(bp("③日報: 出勤メイド数、チェッカー数"));
  el.push(h3("分析手法"));
  el.push(bp("メイド別「清掃1室あたりクレーム率」の算出"));
  el.push(bp("パフォーマンス分布のヒストグラム作成 → トップ10%とボトム10%の特定"));
  el.push(bp("ハイパフォーマーの共通点分析（経験年数/担当フロア/シフトパターン等）"));
  el.push(h3("期待される発見"));
  el.push(bp("クレームの大半が少数のスタッフに集中している可能性（パレートの法則）", { b: true }));
  el.push(bp("チェッカーの検出力にも個人差があり、「見逃し率」の高いチェッカーの特定"));
  el.push(bp("ハイパフォーマーの清掃手順や時間配分の特徴"));
  el.push(h3("CS→売上への因果ロジック"));
  el.push(p("ボトムパフォーマーへの集中OJT + ハイパフォーマーのノウハウ横展開 → クレーム率の均質化・全体底上げ → スコア改善 → 売上増加"));
  el.push(p("推定効果: ボトム20%のスタッフのクレーム率を平均まで改善 → 全体クレーム率20-30%削減", { b: true, c: C.GREEN }));
  el.push(h3("実装"));
  el.push(bp("難易度: 中〜高（DUMMYFUNCTION列の解決が必要な場合あり）"));
  el.push(bp("所要期間: 2-3週間"));

  // Analysis 3
  el.push(p(""), h2("分析3: 人員配置×品質 相関分析"));
  el.push(h3("なぜ必要か"));
  el.push(p("「人手が足りない日にクレームが増える」という現場の感覚はあっても、定量的な裏付けがありません。日報データには日別の出勤メイド数・チェッカー数・稼働客室数が記録されており、「1メイドあたり客室数」と「クレーム発生率」の関係を明らかにできます。"));
  el.push(h3("使用データ"));
  el.push(bp("③日報: 日別の出勤メイド数、出勤チェッカー数、稼働客室数、クレーム件数"));
  el.push(bp("①集計: 月間の稼働率、売上"));
  el.push(h3("分析手法"));
  el.push(bp("「1メイドあたり客室数」vs「日次クレーム率」の散布図＋回帰分析"));
  el.push(bp("チェッカー1人あたり担当客室数と見逃し率の関係"));
  el.push(bp("ホテル間比較: 同規模ホテルでの人員効率の違い"));
  el.push(h3("期待される発見"));
  el.push(bp("「1メイドあたり15室を超えるとクレーム率が急増する」等の閾値の特定", { b: true }));
  el.push(bp("チェッカー不在日のクレーム率との差分"));
  el.push(bp("最適人員配置基準（ホテル規模別の推奨メイド数/チェッカー数）"));
  el.push(h3("CS→売上への因果ロジック"));
  el.push(p("データに基づく最適人員配置 → 繁忙日の品質低下防止 → クレーム率安定 → スコア維持・改善 → 売上安定"));
  el.push(p("推定効果: 人員不足日のクレーム率を50%削減 → 全体クレーム率10-15%削減", { b: true, c: C.GREEN }));
  el.push(h3("実装"));
  el.push(bp("難易度: 中（日報の日次データをopenpyxlで抽出）"));
  el.push(bp("所要期間: 2週間"));

  // Analysis 4
  el.push(p(""), h2("分析4: 清掃完了時間×品質 分析"));
  el.push(h3("なぜ必要か"));
  el.push(p("チェックイン時刻までに全客室の清掃を完了させるプレッシャーが、清掃品質に悪影響を与えている可能性があります。日報に記録された「清掃完了時間」とクレーム件数の関係を分析することで、時間管理の改善がCS向上に直結するかを検証できます。"));
  el.push(h3("使用データ"));
  el.push(bp("③日報: 清掃完了時間（時刻）、クレーム件数、稼働客室数"));
  el.push(h3("分析手法"));
  el.push(bp("完了時間帯別（〜13時/13-14時/14-15時/15時以降）のクレーム率比較"));
  el.push(bp("「稼働客室数÷メイド数」（負荷度）と完了時間の関係"));
  el.push(bp("完了時間とチェッカー検査時間の余裕度分析"));
  el.push(h3("期待される発見"));
  el.push(bp("完了が15時以降にずれ込む日はクレーム率が2倍以上になる等のパターン", { b: true }));
  el.push(bp("チェッカー検査の時間的余裕が品質に与える影響度"));
  el.push(h3("CS→売上への因果ロジック"));
  el.push(p("適正なスケジュール設計 → 時間的余裕の確保 → 丁寧な清掃・確実な検査 → クレーム削減 → スコア改善"));
  el.push(p("推定効果: スケジュール最適化で時間起因クレームを60%削減", { b: true, c: C.GREEN }));
  el.push(h3("実装"));
  el.push(bp("難易度: 低〜中"));
  el.push(bp("所要期間: 1-2週間"));

  // Analysis 5
  el.push(p(""), h2("分析5: 安全チェック×クレーム予兆検出"));
  el.push(h3("なぜ必要か"));
  el.push(p("安全衛生パトロール（✅安全チェック）では整理整頓・安全管理・衛生管理の3項目で◎/△/✖の評価を行っていますが、この結果はクレームデータと紐づけられていません。パトロール評価と翌月のクレーム数を紐づけることで、「予兆検出モデル」を構築できます。"));
  el.push(h3("使用データ"));
  el.push(bp("✅安全チェックシート: パトロール日、評価結果（◎/△/✖）、実施者"));
  el.push(bp("🔵クレームシート: 月別クレーム件数"));
  el.push(h3("分析手法"));
  el.push(bp("パトロール評価の「△/✖」率と翌月クレーム数の相関分析"));
  el.push(bp("「衛生管理✖」がある月の翌月クレーム率の変化を検証"));
  el.push(bp("予兆スコアの開発: パトロール結果から翌月のクレーム件数を予測するモデル"));
  el.push(h3("期待される発見"));
  el.push(bp("安全チェックで「△」以下が3項目以上ある月は、翌月のクレームが有意に増加する等のパターン", { b: true }));
  el.push(bp("どの評価項目（整理整頓/安全管理/衛生管理）がクレーム予測に最も有効か"));
  el.push(h3("CS→売上への因果ロジック"));
  el.push(p("予兆検出 → クレーム発生「前」に予防対策 → クレームの未然防止 → スコア安定 → 売上維持"));
  el.push(p("推定効果: 予兆が検知された月に先手を打つことで、クレーム発生を30-40%回避", { b: true, c: C.GREEN }));
  el.push(h3("実装"));
  el.push(bp("難易度: 高（データの正規化とモデル構築が必要）"));
  el.push(bp("所要期間: 3-4週間"));

  // Analysis 6
  el.push(p(""), h2("分析6: 品質→売上 弾力性分析（ROI精密化）"));
  el.push(h3("なぜ必要か"));
  el.push(p("現在のROI試算は業界ベンチマーク（0.1点改善=RevPAR1%向上）に依存しています。しかし、PRIMECHANGE管理の19ホテルの実データで「スコアと売上の関係」を直接算出すれば、より精度の高い投資判断が可能になります。"));
  el.push(h3("使用データ（既存データで実施可能）"));
  el.push(bp("口コミスコア（19ホテル）× 稼働率 × 売上 × ADR × RevPAR"));
  el.push(bp("hotel_revenue_data.json + primechange_portfolio_analysis.json"));
  el.push(h3("分析手法"));
  el.push(bp("19ホテルの横断データでスコア vs 稼働率/ADR/RevPARの回帰分析"));
  el.push(bp("「スコア1点あたりの稼働率差分」「スコア1点あたりのADR差分」を算出"));
  el.push(bp("ホテル規模（客室数）で補正した弾力性の推定"));
  el.push(h3("期待される発見"));
  el.push(bp("自社データでの弾力性が業界ベンチマークと一致するか乖離するか", { b: true }));
  el.push(bp("ホテル規模や立地による弾力性の違い（大型ホテルは弾力性が高い等）"));
  el.push(bp("「スコア8.0を下回ると稼働率が急落する」等の非線形パターン"));
  el.push(h3("CS→売上への因果ロジック"));
  el.push(p("精密なROI算出 → 経営判断の精度向上 → 最適な投資配分 → 効率的なCS改善 → 売上最大化"));
  el.push(p("推定効果: 投資配分の最適化で同じ予算から20-30%多い効果を引き出す", { b: true, c: C.GREEN }));
  el.push(h3("実装"));
  el.push(bp("難易度: 低（既存データのみで実施可能）", { b: true }));
  el.push(bp("所要期間: 1週間"));

  // Analysis 7
  el.push(p(""), h2("分析7: ベストプラクティス横展開分析"));
  el.push(h3("なぜ必要か"));
  el.push(p("コンフォートスイーツ東京ベイ（9.17点）やERA東京東神田（9.10点）は、同じPRIMECHANGE管理下で突出した品質を達成しています。これらの成功要因を体系的に抽出し、低スコアホテルに横展開することが、ポートフォリオ全体の底上げの最短ルートです。"));
  el.push(h3("使用データ（既存データで部分的に実施可能）"));
  el.push(bp("口コミコメント（analysis.json）: 「good」フィールドの頻出キーワード比較"));
  el.push(bp("売上/スタッフデータ: 客室あたりスタッフ比率の比較"));
  el.push(bp("日報/月報: 高スコアホテルの業務フロー・改善活動の記録"));
  el.push(h3("分析手法"));
  el.push(bp("高スコアホテル（Top5: 9.17, 9.10, 9.03, 8.96, 8.88）の口コミ「good」コメントのテキスト分析"));
  el.push(bp("低スコアホテル（Bottom5: 7.03, 7.15, 7.28, 7.63, 8.00）との比較"));
  el.push(bp("オペレーション指標の比較: メイド1人あたり客室数、チェッカー配置率、清掃完了時間等"));
  el.push(h3("期待される発見"));
  el.push(bp("高スコアホテルに共通する口コミキーワード（「清潔」「丁寧」「対応が早い」等）", { b: true }));
  el.push(bp("高スコアホテルのメイド1人あたり客室数が低スコアホテルより少ない等の構造的要因"));
  el.push(bp("チェッカーの配置パターン（全室チェック vs 抜き打ち）の違い"));
  el.push(h3("CS→売上への因果ロジック"));
  el.push(p("成功要因の体系化 → 標準化マニュアル作成 → 全ホテルへの導入 → ポートフォリオ全体のスコア底上げ → 売上改善"));
  el.push(p("推定効果: 低スコアホテル5施設のスコアを平均0.5点改善 → 月間売上約250万円増", { b: true, c: C.GREEN }));
  el.push(h3("実装"));
  el.push(bp("難易度: 中（テキスト分析＋現場ヒアリング）"));
  el.push(bp("所要期間: 3-4週間"));

  // Summary table
  el.push(p(""), h2("7テーマ一覧サマリー"));

  const summaryData = [
    ["#", "分析テーマ", "難易度", "効果", "期間", "データ"],
    ["1", "クレーム類型×スコア連動", "中", "大", "1-2週", "未活用"],
    ["2", "スタッフ個人パフォーマンス", "中〜高", "大", "2-3週", "未活用"],
    ["3", "人員配置×品質相関", "中", "大", "2週", "未活用"],
    ["4", "清掃完了時間×品質", "低〜中", "中", "1-2週", "未活用"],
    ["5", "安全チェック×予兆検出", "高", "大", "3-4週", "未活用"],
    ["6", "品質→売上弾力性（ROI）", "低", "大", "1週", "既存"],
    ["7", "ベストプラクティス横展開", "中", "大", "3-4週", "既存+"],
  ];

  el.push(new Table({ rows: summaryData.map((r, i) => new TableRow({
    children: r.map((c, j) => cell(c, { w: [400, 3200, 900, 800, 1000, 1200][j], b: i === 0, bg: i === 0 ? C.NAVY : i % 2 === 0 ? C.LIGHT_BG : C.WHITE, c: i === 0 ? C.WHITE : c === "大" ? C.GREEN : c === "低" ? C.GREEN : c === "高" ? C.RED : C.TEXT, a: j <= 1 ? AlignmentType.LEFT : AlignmentType.CENTER, sz: 15 }))
  })), width: { size: 7500, type: WidthType.DXA } }));

  el.push(PB());
  return el;
}

// ============================================================
// Section 5: Action Plan
// ============================================================
function buildActionPlan() {
  const el = [
    h1("4. 統合アクションプラン"),
    p("7つの分析コンセプトに基づき、PRIMECHANGEが具体的に取るべきアクションを3フェーズで提案します。売上インパクトと実装の容易さのバランスを考慮し、「最も効果が高く、すぐにできること」から優先的に実施します。"),
  ];

  // Phase 1
  el.push(p(""), h2("Phase 1: 即時アクション（2週間以内）"));
  el.push(p("既存データで今すぐ実行可能な施策。投資は最小限で最大の効果を狙う。", { b: true, c: C.RED }));

  const phase1 = [
    ["施策", "担当", "期限", "成功指標", "対象ホテル"],
    ["品質→売上弾力性の算出（分析6）", "データ分析チーム", "1週間", "自社データでのROI式確立", "全19ホテル"],
    ["URGENTホテル緊急清掃監査", "品質管理部", "3日以内", "全室抜き打ち検査完了", "博多/蒲田/浜松町/新横浜"],
    ["低品質×高売上3ホテルの責任者面談", "経営層", "1週間", "改善コミット合意", "博多/蒲田/浜松町"],
    ["クレーム類型データの抽出開始（分析1）", "データ分析チーム", "2週間", "19ホテル分の抽出完了", "全19ホテル"],
  ];

  el.push(new Table({ rows: phase1.map((r, i) => new TableRow({
    children: r.map((c, j) => cell(c, { w: [2500, 1500, 1000, 2000, 2500][j], b: i === 0, bg: i === 0 ? C.RED : C.WHITE, c: i === 0 ? C.WHITE : C.TEXT, a: j === 0 || j === 4 ? AlignmentType.LEFT : AlignmentType.CENTER, sz: 14 }))
  })), width: { size: 9500, type: WidthType.DXA } }));

  el.push(p("推定効果: Phase 1だけで月間売上100-200万円の減少防止（低品質ホテルの急落回避）", { b: true, c: C.GREEN, sb: 120 }));

  // Phase 2
  el.push(p(""), h2("Phase 2: 短期改善プログラム（1-3ヶ月）"));
  el.push(p("データ分析基盤を構築し、分析結果に基づく現場改善を開始。", { b: true, c: C.ORANGE }));

  const phase2 = [
    ["施策", "担当", "期限", "成功指標"],
    ["クレーム類型×スコア連動分析の完了（分析1）", "データ分析", "1ヶ月", "類型別インパクトの定量化"],
    ["スタッフパフォーマンス分析の実施（分析2）", "データ分析+現場", "6週間", "全メイドのクレーム率算出"],
    ["人員配置最適化基準の策定（分析3）", "運営管理", "2ヶ月", "ホテル別推奨人員表の完成"],
    ["ボトムパフォーマーへの集中OJT開始", "教育担当", "1ヶ月〜", "対象スタッフの特定と研修開始"],
    ["ベストプラクティス横展開プログラム（分析7）", "品質管理+現場", "3ヶ月", "成功事例集の作成・共有"],
    ["清掃完了時間の分析と改善（分析4）", "運営管理", "6週間", "適正スケジュールの導入"],
  ];

  el.push(new Table({ rows: phase2.map((r, i) => new TableRow({
    children: r.map((c, j) => cell(c, { w: [3200, 1500, 1200, 3100][j], b: i === 0, bg: i === 0 ? C.ORANGE : i % 2 === 0 ? C.LIGHT_BG : C.WHITE, c: i === 0 ? C.WHITE : C.TEXT, a: j === 0 || j === 3 ? AlignmentType.LEFT : AlignmentType.CENTER, sz: 14 }))
  })), width: { size: 9000, type: WidthType.DXA } }));

  el.push(p("推定効果: URGENTホテルのスコア平均0.3-0.5点改善、月間売上200-300万円の改善", { b: true, c: C.GREEN, sb: 120 }));

  // Phase 3
  el.push(p(""), h2("Phase 3: 中期 仕組み化・PDCA定着（3-6ヶ月）"));
  el.push(p("分析を「一度きり」ではなく「継続的な改善サイクル」として仕組み化する。", { b: true, c: C.ACCENT }));

  el.push(h3("月次品質ダッシュボードの構築"));
  el.push(bp("19ホテルの品質KPI（口コミスコア/クレーム率/人員充足率/完了時間）を自動集計"));
  el.push(bp("異常値アラート機能: スコア急落やクレーム急増を自動検知"));
  el.push(bp("ホテル間ベンチマーキング: 同規模ホテルとの比較を可視化"));

  el.push(h3("安全チェック×予兆検出の自動化（分析5）"));
  el.push(bp("パトロール結果から翌月のクレーム予測スコアを自動算出"));
  el.push(bp("「要注意」判定が出たホテルへの先手対応プロトコルの整備"));

  el.push(h3("スタッフ評価制度への品質KPI組み込み"));
  el.push(bp("個人別クレーム率を評価基準に追加（分析2の結果を活用）"));
  el.push(bp("クレームゼロ月間達成のインセンティブ制度"));
  el.push(bp("ハイパフォーマーのメンター制度の導入"));

  el.push(h3("継続的改善サイクルの確立"));
  el.push(bp("月次: 品質ダッシュボードレビュー（経営会議）"));
  el.push(bp("週次: URGENTホテルの進捗確認"));
  el.push(bp("四半期: 全ホテルの品質監査＋ベストプラクティス共有会"));

  el.push(p("推定効果: ポートフォリオ平均スコア8.39→8.80以上、URGENT判定0件、月間売上3-5%（300-500万円）改善", { b: true, c: C.GREEN, sb: 120 }));

  el.push(PB());
  return el;
}

// ============================================================
// Section 6: KPI Framework
// ============================================================
function buildKPIFramework() {
  const el = [
    h1("5. KPI・効果測定フレームワーク"),
    p("各施策の成果を定量的に追跡するためのKPIを設定します。"),
  ];

  const kpis = [
    ["KPI", "現状値", "3ヶ月目標", "6ヶ月目標", "測定方法"],
    ["ポートフォリオ平均スコア", "8.39", "8.55", "8.80", "月次口コミ集計"],
    ["清掃クレーム率", "4.6%", "3.5%", "2.5%", "🔵クレーム月次集計"],
    ["URGENT判定ホテル数", "4", "2", "0", "優先度再判定"],
    ["月間総売上", "¥99.8M", "¥101M", "¥104M", "①集計シート"],
    ["平均稼働率", "73.9%", "75.0%", "77.0%", "①集計シート"],
    ["高評価率（8点以上）", "78.1%", "80.0%", "83.0%", "口コミ分析"],
    ["スタッフ充足率", "未測定", "90%", "95%", "④月報"],
    ["予兆検出対応率", "0%", "50%", "90%", "安全チェック連動"],
  ];

  el.push(new Table({ rows: kpis.map((r, i) => new TableRow({
    children: r.map((c, j) => cell(c, { w: [2200, 1200, 1200, 1200, 2200][j], b: i === 0, bg: i === 0 ? C.NAVY : i % 2 === 0 ? C.LIGHT_BG : C.WHITE, c: i === 0 ? C.WHITE : C.TEXT, a: j === 0 || j === 4 ? AlignmentType.LEFT : AlignmentType.CENTER, sz: 15 }))
  })), width: { size: 8000, type: WidthType.DXA } }));

  el.push(p(""), h3("レポーティングサイクル"));
  el.push(bp("週次: URGENTホテルのクレーム件数・対応状況"));
  el.push(bp("月次: 全KPIのダッシュボードレビュー（経営会議にて）"));
  el.push(bp("四半期: 品質監査結果＋ベストプラクティス共有会"));
  el.push(bp("半期: 戦略レビュー＋次期目標設定"));

  el.push(PB());
  return el;
}

// ============================================================
// Section 7: ROI
// ============================================================
function buildROI() {
  const el = [
    h1("6. 投資対効果（ROI）試算"),
    p("7つの分析テーマの実施と、それに基づく改善施策への投資の費用対効果を3シナリオで試算します。"),
  ];

  const roi = [
    ["項目", "控えめシナリオ", "標準シナリオ", "楽観シナリオ"],
    ["対象", "URGENT 4ホテル集中", "URGENT+HIGH 7ホテル", "全19ホテル"],
    ["分析投資", "100-200万円", "200-400万円", "400-600万円"],
    ["改善施策投資", "200-300万円", "400-600万円", "600-800万円"],
    ["スコア改善幅", "+0.3〜0.5点", "+0.3〜0.7点(対象)", "+0.3〜0.5点(全体)"],
    ["月間売上改善", "+100〜200万円", "+200〜400万円", "+300〜500万円"],
    ["年間売上改善", "+1,200〜2,400万円", "+2,400〜4,800万円", "+3,600〜6,000万円"],
    ["投資回収期間", "3-6ヶ月", "4-8ヶ月", "6-12ヶ月"],
    ["PRIMECHANGE効果", "契約維持確保", "単価交渉材料", "新規受注競争力"],
  ];

  el.push(new Table({ rows: roi.map((r, i) => new TableRow({
    children: r.map((c, j) => cell(c, { w: [2000, 2300, 2300, 2400][j], b: i === 0 || j === 0, bg: i === 0 ? C.NAVY : j === 0 ? C.LIGHT_BG : C.WHITE, c: i === 0 ? C.WHITE : j === 0 ? C.NAVY : C.TEXT, a: j === 0 ? AlignmentType.LEFT : AlignmentType.CENTER, sz: 15 }))
  })), width: { size: 9000, type: WidthType.DXA } }));

  el.push(p(""), p("※ 試算の前提: 業界データ「スコア0.1点改善 ≒ RevPAR 1%向上」を基本に、自社19ホテルの規模で補正。分析6（弾力性分析）の結果で精密化予定。", { sz: 17, c: C.SUBTEXT, i: true }));

  el.push(p(""), h3("PRIMECHANGEにとっての戦略的価値"));
  el.push(bp("契約維持: クレーム削減・スコア改善の実績がホテルオーナーとの関係強化に直結", { b: true }));
  el.push(bp("単価交渉: 「品質改善で稼働率が向上した」というデータは、清掃単価の引き上げ交渉の最強の武器"));
  el.push(bp("新規受注: 「スコアを0.5点改善した実績」は、新規ホテルへの営業で最も説得力のあるエビデンス"));
  el.push(bp("差別化: データドリブンな品質管理は、清掃管理業界での明確な差別化要因"));

  return el;
}

// ============================================================
// Build & Save
// ============================================================
async function main() {
  console.log("Building CS Strategy Report DOCX...");

  const sections = [
    ...buildCover(),
    ...buildExecSummary(),
    ...buildCurrentState(),
    ...buildAnalysisConcepts(),
    ...buildActionPlan(),
    ...buildKPIFramework(),
    ...buildROI(),
  ];

  const doc = new Document({
    styles: { default: { document: { run: { font: "Arial", size: 20 } } } },
    sections: [{
      properties: {
        page: { size: { width: 12240, height: 15840 }, margin: { top: 720, bottom: 720, left: 900, right: 900 } },
      },
      headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "PRIMECHANGE CS向上×売上増加 戦略提案書", size: 14, font: "Arial", color: C.SUBTEXT, italics: true })] })] }) },
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
  fs.writeFileSync(OUTPUT, buf);
  console.log(`✅ DOCX: ${OUTPUT} (${(buf.length / 1024).toFixed(1)} KB)`);
}

main().catch(e => { console.error(e); process.exit(1); });
