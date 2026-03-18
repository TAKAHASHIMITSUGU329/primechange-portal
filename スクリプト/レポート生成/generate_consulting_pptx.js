#!/usr/bin/env node
"use strict";

const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");

// ============================================================
// Data
// ============================================================
const ROOT = path.resolve(__dirname, "../..");
const hotelsRanked = JSON.parse(fs.readFileSync(path.join(ROOT, "hotel-ranked.json"), "utf-8"));
const OUTPUT = path.resolve(ROOT, "納品レポート/PRIMECHANGE_経営コンサルティングレポート.pptx");

// ============================================================
// Colors & Helpers
// ============================================================
const C = {
  NAVY: "1B3A5C", ACCENT: "C2333A", WHITE: "FFFFFF", LIGHT_BG: "F5F7FA",
  TEXT: "333333", SUBTEXT: "666666", GREEN: "27AE60", ORANGE: "FF9800",
  RED: "E74C3C", BLUE: "2E75B6", TEAL: "00695C", DARK_BG: "1E293B",
};

const pptx = new pptxgen();
pptx.layout = "LAYOUT_16x9";
pptx.author = "PRIMECHANGE";
pptx.title = "PRIMECHANGE 経営コンサルティングレポート";

const shadow = () => ({ type: "outer", blur: 4, offset: 2, angle: 135, color: "000000", opacity: 0.10 });

function addFooter(slide, num) {
  slide.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.NAVY } });
  slide.addText("Confidential | PRIMECHANGE 経営コンサルティングレポート", { x: 0.5, y: 5.25, w: 6, h: 0.375, fontSize: 8, color: "94A3B8", fontFace: "Arial", valign: "middle" });
  slide.addText(String(num), { x: 9, y: 5.25, w: 0.5, h: 0.375, fontSize: 8, color: "94A3B8", fontFace: "Arial", align: "right", valign: "middle" });
}

function addHeader(slide, title, subtitle) {
  slide.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.6, fill: { color: C.NAVY } });
  slide.addText(title, { x: 0.5, y: 0.05, w: 9, h: 0.35, fontSize: 20, fontFace: "Arial", color: C.WHITE, bold: true });
  if (subtitle) slide.addText(subtitle, { x: 0.5, y: 0.35, w: 9, h: 0.22, fontSize: 10, fontFace: "Arial", color: "5A9BE6" });
}

function kpiCard(slide, x, y, w, h, label, value, color, bgColor) {
  slide.addShape(pptx.shapes.RECTANGLE, { x, y, w, h, fill: { color: bgColor }, shadow: shadow() });
  slide.addShape(pptx.shapes.RECTANGLE, { x, y, w, h: 0.04, fill: { color } });
  slide.addText(label, { x, y: y + 0.1, w, h: 0.2, fontSize: 9, fontFace: "Arial", color: C.SUBTEXT, align: "center" });
  slide.addText(value, { x, y: y + 0.3, w, h: 0.4, fontSize: 26, fontFace: "Arial", color, bold: true, align: "center" });
}

// ============================================================
// Slide 1: Title
// ============================================================
(function () {
  const s = pptx.addSlide();
  s.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: C.NAVY } });
  s.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0, w: 0.08, h: 5.625, fill: { color: C.ACCENT } });
  s.addText("PRIMECHANGE", { x: 0.8, y: 1.0, w: 8.4, h: 0.6, fontSize: 20, fontFace: "Arial", color: "5A9BE6" });
  s.addText("経営コンサルティング", { x: 0.8, y: 1.6, w: 8.4, h: 0.8, fontSize: 36, fontFace: "Arial", color: C.WHITE, bold: true });
  s.addText("レポート", { x: 0.8, y: 2.4, w: 8.4, h: 0.7, fontSize: 32, fontFace: "Arial", color: C.WHITE, bold: true });
  s.addShape(pptx.shapes.RECTANGLE, { x: 0.8, y: 3.2, w: 2.0, h: 0.04, fill: { color: C.ACCENT } });
  s.addText("Management Consulting Report — Data-Driven Growth Strategy", { x: 0.8, y: 3.4, w: 8.4, h: 0.35, fontSize: 13, fontFace: "Arial", color: "94A3B8", italics: true });
  s.addText("対象：19ホテル | 口コミ：2,070件 | 月間総売上：¥9,985万", { x: 0.8, y: 4.0, w: 8, h: 0.3, fontSize: 11, fontFace: "Arial", color: "64748B" });
  s.addText("2026年3月17日", { x: 0.8, y: 4.3, w: 4, h: 0.3, fontSize: 11, fontFace: "Arial", color: "64748B" });
  s.addText("株式会社PRIMECHANGE", { x: 0.8, y: 4.9, w: 8, h: 0.3, fontSize: 10, fontFace: "Arial", color: "475569" });
})();

// ============================================================
// Slide 2: Executive Summary
// ============================================================
(function () {
  const s = pptx.addSlide();
  addHeader(s, "エグゼクティブサマリー", "5つの提言領域 — データ資産を最大限に活用する成長戦略");

  kpiCard(s, 0.3, 0.8, 2.1, 0.85, "管理ホテル", "19", C.BLUE, "E3F2FD");
  kpiCard(s, 2.65, 0.8, 2.1, 0.85, "口コミ分析", "2,070件", C.TEAL, "E0F2F1");
  kpiCard(s, 5.0, 0.8, 2.1, 0.85, "平均スコア", "8.39", C.GREEN, "E8F5E9");
  kpiCard(s, 7.35, 0.8, 2.1, 0.85, "月間総売上", "¥1.0億", C.ACCENT, "FFEBEE");

  s.addText("核心的発見", { x: 0.4, y: 1.85, w: 9, h: 0.3, fontSize: 14, fontFace: "Arial", color: C.NAVY, bold: true });
  s.addText(
    "PRIMECHANGEは清掃会社でありながら、業界で極めて稀な「データ分析基盤」を保有。\n" +
    "6 OTAから2,070件の口コミを構造化分析し、7種の深掘り分析と79本の自動レポートを生成。\n" +
    "しかしこの強みはHPにも営業にも活かされておらず、最大の成長機会を逃している。",
    { x: 0.4, y: 2.15, w: 9.2, h: 0.85, fontSize: 10, fontFace: "Arial", color: C.TEXT, lineSpacingMultiple: 1.5 }
  );

  const areas = [
    { n: "1", t: "データ資産の事業化", d: "営業ツール→顧客ポータル→SaaS", c: C.BLUE },
    { n: "2", t: "品質→売上ROI", d: "年間+1.47億円の売上改善余地", c: C.GREEN },
    { n: "3", t: "HP・営業戦略刷新", d: "データドリブンなブランディング", c: C.ACCENT },
    { n: "4", t: "ポートフォリオ最適化", d: "19ホテルのティア別戦略", c: C.ORANGE },
    { n: "5", t: "組織・オペ強化", d: "KPI運用・自動化・V3完成", c: C.TEAL },
  ];
  areas.forEach((a, i) => {
    const x = 0.3 + i * 1.88;
    s.addShape(pptx.shapes.RECTANGLE, { x, y: 3.15, w: 1.72, h: 1.8, fill: { color: C.WHITE }, shadow: shadow() });
    s.addShape(pptx.shapes.RECTANGLE, { x, y: 3.15, w: 1.72, h: 0.04, fill: { color: a.c } });
    s.addShape(pptx.shapes.OVAL, { x: x + 0.6, y: 3.3, w: 0.5, h: 0.5, fill: { color: a.c } });
    s.addText(a.n, { x: x + 0.6, y: 3.3, w: 0.5, h: 0.5, fontSize: 18, fontFace: "Arial", color: C.WHITE, bold: true, align: "center", valign: "middle" });
    s.addText(a.t, { x: x + 0.05, y: 3.9, w: 1.62, h: 0.35, fontSize: 10, fontFace: "Arial", color: C.NAVY, bold: true, align: "center" });
    s.addText(a.d, { x: x + 0.05, y: 4.25, w: 1.62, h: 0.55, fontSize: 8, fontFace: "Arial", color: C.SUBTEXT, align: "center" });
  });

  addFooter(s, 2);
})();

// ============================================================
// Slide 3: Data Asset Monetization
// ============================================================
(function () {
  const s = pptx.addSlide();
  addHeader(s, "領域1：データ資産の事業化", "清掃業界で類を見ない知的資産を事業に転換する");

  s.addText("保有するデジタル資産", { x: 0.4, y: 0.75, w: 9, h: 0.3, fontSize: 13, fontFace: "Arial", color: C.NAVY, bold: true });

  const assets = [
    { t: "口コミ分析基盤", d: "6 OTAから2,070件を\n構造化分析", c: C.BLUE },
    { t: "3世代ダッシュボード", d: "V1(6P)→V2(9P)→V3(10P)\nの進化", c: C.TEAL },
    { t: "自動レポート生成", d: "79本のDOCX/PPTXを\n自動生成", c: C.GREEN },
    { t: "7種深掘り分析", d: "クレーム分類・人員配置\n売上弾性等", c: C.ACCENT },
  ];
  assets.forEach((a, i) => {
    const x = 0.3 + i * 2.35;
    s.addShape(pptx.shapes.RECTANGLE, { x, y: 1.1, w: 2.15, h: 1.1, fill: { color: C.WHITE }, shadow: shadow() });
    s.addShape(pptx.shapes.RECTANGLE, { x, y: 1.1, w: 2.15, h: 0.04, fill: { color: a.c } });
    s.addText(a.t, { x: x + 0.05, y: 1.2, w: 2.05, h: 0.3, fontSize: 10, fontFace: "Arial", color: a.c, bold: true, align: "center" });
    s.addText(a.d, { x: x + 0.05, y: 1.5, w: 2.05, h: 0.55, fontSize: 9, fontFace: "Arial", color: C.TEXT, align: "center" });
  });

  s.addText("3段階の事業化ロードマップ", { x: 0.4, y: 2.4, w: 9, h: 0.3, fontSize: 13, fontFace: "Arial", color: C.NAVY, bold: true });

  // Stage cards
  const stages = [
    { n: "Stage 1", t: "営業ツール活用", when: "即時", items: "・ダッシュボードデモで新規提案\n・79本レポートを匿名サンプルに\n・競合との明確な差別化", effect: "新規受注率向上", c: C.BLUE },
    { n: "Stage 2", t: "顧客ポータル提供", when: "3-6ヶ月", items: "・各ホテルに専用ログイン\n・月次レポート自動配信\n・月額5-10万円アップセル", effect: "解約率低下", c: C.TEAL },
    { n: "Stage 3", t: "SaaS独立事業", when: "6-12ヶ月", items: "・清掃契約なしでも分析提供\n・チェーン本部向け管理ツール\n・月額3-8万円/ホテル", effect: "ストック収入確立", c: C.GREEN },
  ];
  stages.forEach((st, i) => {
    const x = 0.3 + i * 3.15;
    s.addShape(pptx.shapes.RECTANGLE, { x, y: 2.8, w: 2.95, h: 2.2, fill: { color: C.WHITE }, shadow: shadow() });
    s.addShape(pptx.shapes.RECTANGLE, { x, y: 2.8, w: 2.95, h: 0.35, fill: { color: st.c } });
    s.addText(`${st.n}（${st.when}）`, { x, y: 2.8, w: 2.95, h: 0.35, fontSize: 11, fontFace: "Arial", color: C.WHITE, bold: true, align: "center", valign: "middle" });
    s.addText(st.t, { x: x + 0.1, y: 3.2, w: 2.75, h: 0.3, fontSize: 11, fontFace: "Arial", color: st.c, bold: true });
    s.addText(st.items, { x: x + 0.1, y: 3.5, w: 2.75, h: 1.0, fontSize: 8.5, fontFace: "Arial", color: C.TEXT, lineSpacingMultiple: 1.4 });
    s.addShape(pptx.shapes.RECTANGLE, { x: x + 0.1, y: 4.55, w: 2.75, h: 0.3, fill: { color: "F0F4F8" } });
    s.addText(`期待効果：${st.effect}`, { x: x + 0.1, y: 4.55, w: 2.75, h: 0.3, fontSize: 9, fontFace: "Arial", color: st.c, bold: true, align: "center", valign: "middle" });
  });

  addFooter(s, 3);
})();

// ============================================================
// Slide 4: Quality → Revenue ROI
// ============================================================
(function () {
  const s = pptx.addSlide();
  addHeader(s, "領域2：清掃品質 → 売上のROIストーリー", "スコアと売上の統計的相関をホテルオーナーへの価値提案に転換");

  // Regression table
  s.addText("回帰分析結果", { x: 0.4, y: 0.75, w: 4, h: 0.25, fontSize: 12, fontFace: "Arial", color: C.NAVY, bold: true });
  const regRows = [
    ["指標", "スコア1点あたり", "相関 r", "解釈"],
    ["稼働率", "+3.9%pt", "0.34", "弱〜中"],
    ["ADR", "+¥255", "0.56", "中"],
    ["RevPAR", "+¥230", "0.60", "中〜強"],
    ["月間売上/室", "+¥6,453", "0.60", "中〜強"],
  ];
  s.addTable(regRows, {
    x: 0.3, y: 1.05, w: 4.5, fontSize: 9, fontFace: "Arial",
    border: { type: "solid", pt: 0.5, color: "CCCCCC" },
    colW: [1.3, 1.2, 0.8, 1.2],
    rowH: [0.28, 0.25, 0.25, 0.25, 0.25],
    autoPage: false,
    color: C.TEXT,
  });
  // Make header row styled
  s.addShape(pptx.shapes.RECTANGLE, { x: 0.3, y: 1.05, w: 4.5, h: 0.28, fill: { color: C.NAVY } });
  s.addText("指標        スコア1点あたり   相関 r     解釈", { x: 0.35, y: 1.05, w: 4.4, h: 0.28, fontSize: 8, fontFace: "Arial", color: C.WHITE, bold: true, valign: "middle" });

  // Score band performance
  s.addText("スコア帯別パフォーマンス", { x: 5.2, y: 0.75, w: 4, h: 0.25, fontSize: 12, fontFace: "Arial", color: C.NAVY, bold: true });
  const bandRows = [
    ["スコア帯", "ホテル数", "平均RevPAR", "最高帯との差"],
    ["7.0-8.0", "5", "¥711", "-35%"],
    ["8.0-8.5", "3", "¥890", "-18%"],
    ["8.5-9.0", "7", "¥1,029", "-6%"],
    ["9.0+", "4", "¥1,092", "基準"],
  ];
  s.addTable(bandRows, {
    x: 5.1, y: 1.05, w: 4.5, fontSize: 9, fontFace: "Arial",
    border: { type: "solid", pt: 0.5, color: "CCCCCC" },
    colW: [1.0, 0.9, 1.2, 1.4],
    rowH: [0.28, 0.25, 0.25, 0.25, 0.25],
    autoPage: false, color: C.TEXT,
  });
  s.addShape(pptx.shapes.RECTANGLE, { x: 5.1, y: 1.05, w: 4.5, h: 0.28, fill: { color: C.NAVY } });
  s.addText("スコア帯    ホテル数   平均RevPAR   最高帯との差", { x: 5.15, y: 1.05, w: 4.4, h: 0.28, fontSize: 8, fontFace: "Arial", color: C.WHITE, bold: true, valign: "middle" });

  // Improvement scenarios
  s.addText("改善シナリオ別インパクト", { x: 0.4, y: 2.6, w: 9, h: 0.3, fontSize: 13, fontFace: "Arial", color: C.NAVY, bold: true });

  const scenarios = [
    { label: "シナリオA\n+0.1点", rev: "+¥23/RevPAR", monthly: "+¥246万/月", annual: "+¥2,949万/年", c: C.BLUE, bg: "E3F2FD" },
    { label: "シナリオB\n+0.3点", rev: "+¥69/RevPAR", monthly: "+¥737万/月", annual: "+¥8,846万/年", c: C.ORANGE, bg: "FFF3E0" },
    { label: "シナリオC\n+0.5点", rev: "+¥115/RevPAR", monthly: "+¥1,229万/月", annual: "+¥1.47億/年", c: C.GREEN, bg: "E8F5E9" },
  ];
  scenarios.forEach((sc, i) => {
    const x = 0.3 + i * 3.15;
    s.addShape(pptx.shapes.RECTANGLE, { x, y: 3.0, w: 2.95, h: 1.9, fill: { color: sc.bg }, shadow: shadow() });
    s.addShape(pptx.shapes.RECTANGLE, { x, y: 3.0, w: 2.95, h: 0.04, fill: { color: sc.c } });
    s.addText(sc.label, { x, y: 3.1, w: 2.95, h: 0.5, fontSize: 11, fontFace: "Arial", color: sc.c, bold: true, align: "center" });
    s.addText(sc.rev, { x, y: 3.6, w: 2.95, h: 0.3, fontSize: 10, fontFace: "Arial", color: C.TEXT, align: "center" });
    s.addText(sc.monthly, { x, y: 3.9, w: 2.95, h: 0.3, fontSize: 10, fontFace: "Arial", color: C.TEXT, align: "center" });
    s.addText(sc.annual, { x, y: 4.25, w: 2.95, h: 0.4, fontSize: 18, fontFace: "Arial", color: sc.c, bold: true, align: "center" });
  });

  s.addShape(pptx.shapes.RECTANGLE, { x: 0.3, y: 5.0, w: 9.4, h: 0.02, fill: { color: C.ACCENT } });

  addFooter(s, 4);
})();

// ============================================================
// Slide 5: HP Gap Analysis
// ============================================================
(function () {
  const s = pptx.addSlide();
  addHeader(s, "領域3：ホームページと営業戦略のギャップ", "最大の強み（データ分析力）が対外的に全く伝わっていない");

  s.addText("現状の問題", { x: 0.4, y: 0.75, w: 4, h: 0.25, fontSize: 13, fontFace: "Arial", color: C.RED, bold: true });

  const problems = [
    { item: "メッセージ", status: "「本気の清掃」— 抽象的" },
    { item: "実績・データ", status: "記載なし" },
    { item: "ブログ", status: "自己啓発的で事業無関連" },
    { item: "FC募集", status: "バナーのみ・詳細なし" },
    { item: "会社概要", status: "住所・電話のみ" },
  ];
  problems.forEach((pr, i) => {
    const y = 1.1 + i * 0.32;
    s.addShape(pptx.shapes.RECTANGLE, { x: 0.4, y, w: 0.06, h: 0.25, fill: { color: C.RED } });
    s.addText(pr.item, { x: 0.6, y, w: 1.5, h: 0.25, fontSize: 9, fontFace: "Arial", color: C.NAVY, bold: true });
    s.addText(pr.status, { x: 2.1, y, w: 2.8, h: 0.25, fontSize: 9, fontFace: "Arial", color: C.SUBTEXT });
  });

  s.addText("改善提案（優先順位順）", { x: 5.2, y: 0.75, w: 4.5, h: 0.25, fontSize: 13, fontFace: "Arial", color: C.GREEN, bold: true });

  const proposals = [
    { n: "1", t: "トップに数値実績", d: "19ホテル / 2,070件 / 8.39点", pri: "最優先" },
    { n: "2", t: "事例ページ新設", d: "Before/After・デモ動画・サンプルDL", pri: "高" },
    { n: "3", t: "ブログ戦略転換", d: "業界分析記事でSEO強化", pri: "中" },
    { n: "4", t: "FC募集ページ充実", d: "収益モデル・支援内容・声", pri: "中" },
    { n: "5", t: "10周年ブランディング", d: "記念キャンペーン・PR", pri: "中" },
  ];
  proposals.forEach((pr, i) => {
    const y = 1.1 + i * 0.52;
    s.addShape(pptx.shapes.OVAL, { x: 5.3, y: y + 0.05, w: 0.3, h: 0.3, fill: { color: C.GREEN } });
    s.addText(pr.n, { x: 5.3, y: y + 0.05, w: 0.3, h: 0.3, fontSize: 12, fontFace: "Arial", color: C.WHITE, bold: true, align: "center", valign: "middle" });
    s.addText(pr.t, { x: 5.75, y, w: 2, h: 0.25, fontSize: 10, fontFace: "Arial", color: C.NAVY, bold: true });
    s.addText(pr.d, { x: 5.75, y: y + 0.22, w: 3.5, h: 0.22, fontSize: 8, fontFace: "Arial", color: C.SUBTEXT });
  });

  // Key message
  s.addShape(pptx.shapes.RECTANGLE, { x: 0.3, y: 3.9, w: 9.4, h: 0.9, fill: { color: "FFF8E1" }, shadow: shadow() });
  s.addShape(pptx.shapes.RECTANGLE, { x: 0.3, y: 3.9, w: 0.06, h: 0.9, fill: { color: C.ORANGE } });
  s.addText("核心メッセージ", { x: 0.5, y: 3.95, w: 9, h: 0.3, fontSize: 11, fontFace: "Arial", color: C.ORANGE, bold: true });
  s.addText(
    "「清掃品質の見える化で、口コミスコア向上 → 売上向上を実現する」— このメッセージをHPのファーストビューに。\n" +
    "ダッシュボードのスクリーンショットと、「スコア+0.5で年間+1.47億円」の数字を出すだけで、競合と決定的に差別化できる。",
    { x: 0.5, y: 4.25, w: 9, h: 0.5, fontSize: 9, fontFace: "Arial", color: C.TEXT, lineSpacingMultiple: 1.4 }
  );

  addFooter(s, 5);
})();

// ============================================================
// Slide 6: Portfolio Strategy
// ============================================================
(function () {
  const s = pptx.addSlide();
  addHeader(s, "領域4：ポートフォリオ戦略", "19ホテルのティア別リソース配分と成長戦略");

  // Tier summary
  const tiers = [
    { name: "優秀 (9.0+)", count: 3, avg: "9.10", issue: "0.6%", strategy: "ベストプラクティス\n抽出・横展開", c: C.GREEN },
    { name: "良好 (8.5-9.0)", count: 7, avg: "8.74", issue: "4.0%", strategy: "維持・微調整", c: C.BLUE },
    { name: "概ね良好\n(8.0-8.5)", count: 4, avg: "8.27", issue: "3.6%", strategy: "重点監視\n予防的改善", c: C.ORANGE },
    { name: "要改善 (<8.0)", count: 5, avg: "7.41", issue: "11.1%", strategy: "緊急集中改善\nプログラム", c: C.RED },
  ];
  tiers.forEach((t, i) => {
    const x = 0.3 + i * 2.4;
    s.addShape(pptx.shapes.RECTANGLE, { x, y: 0.75, w: 2.2, h: 1.6, fill: { color: C.WHITE }, shadow: shadow() });
    s.addShape(pptx.shapes.RECTANGLE, { x, y: 0.75, w: 2.2, h: 0.04, fill: { color: t.c } });
    s.addText(t.name, { x, y: 0.85, w: 2.2, h: 0.35, fontSize: 10, fontFace: "Arial", color: t.c, bold: true, align: "center" });
    s.addText(String(t.count) + "ホテル", { x, y: 1.2, w: 2.2, h: 0.3, fontSize: 22, fontFace: "Arial", color: t.c, bold: true, align: "center" });
    s.addText(`平均: ${t.avg} | 課題率: ${t.issue}`, { x, y: 1.5, w: 2.2, h: 0.22, fontSize: 8, fontFace: "Arial", color: C.SUBTEXT, align: "center" });
    s.addText(t.strategy, { x, y: 1.75, w: 2.2, h: 0.5, fontSize: 8.5, fontFace: "Arial", color: C.TEXT, align: "center" });
  });

  // Strategic actions
  s.addText("4つの戦略的アクション", { x: 0.4, y: 2.55, w: 9, h: 0.3, fontSize: 13, fontFace: "Arial", color: C.NAVY, bold: true });

  const actions = [
    { letter: "A", title: "要改善5ホテル集中改善", detail: "博多(14.9%)、浜松町(12.5%)、新横浜(9.2%)、\n蒲田(9.8%)、横浜関内(9.2%)\n→ 6ヶ月でスコア8.0超え目標\n→ 月額+292万円（+0.5点時）", c: C.RED },
    { letter: "B", title: "ベストプラクティス横展開", detail: "ERA東神田(9.16)、スイーツ東京ベイ(9.15)の\n清掃手順を「PRIMECHANGE清掃スタンダード」\nとして文書化し全ホテルに展開", c: C.GREEN },
    { letter: "C", title: "新規獲得ターゲティング", detail: "理想：スコア7.0-8.0の中価格帯BH\n営業ピッチ：スコア分析→売上増シミュレーション\n回避：超高級/スコア6.0未満", c: C.BLUE },
    { letter: "D", title: "チェーン一括提案", detail: "コンフォート系6ホテルが既にポートフォリオに\n→ チェーン本部への一括受注提案\n「全系列にデータ分析+清掃を包括」", c: C.TEAL },
  ];
  actions.forEach((a, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = 0.3 + col * 4.85, y = 2.95 + row * 1.15;
    s.addShape(pptx.shapes.RECTANGLE, { x, y, w: 4.65, h: 1.05, fill: { color: C.WHITE }, shadow: shadow() });
    s.addShape(pptx.shapes.RECTANGLE, { x, y, w: 0.04, h: 1.05, fill: { color: a.c } });
    s.addShape(pptx.shapes.OVAL, { x: x + 0.15, y: y + 0.15, w: 0.4, h: 0.4, fill: { color: a.c } });
    s.addText(a.letter, { x: x + 0.15, y: y + 0.15, w: 0.4, h: 0.4, fontSize: 16, fontFace: "Arial", color: C.WHITE, bold: true, align: "center", valign: "middle" });
    s.addText(a.title, { x: x + 0.65, y: y + 0.08, w: 3.8, h: 0.25, fontSize: 11, fontFace: "Arial", color: a.c, bold: true });
    s.addText(a.detail, { x: x + 0.65, y: y + 0.33, w: 3.8, h: 0.65, fontSize: 8, fontFace: "Arial", color: C.TEXT, lineSpacingMultiple: 1.3 });
  });

  addFooter(s, 6);
})();

// ============================================================
// Slide 7: Org & Operations
// ============================================================
(function () {
  const s = pptx.addSlide();
  addHeader(s, "領域5：組織・オペレーション強化", "KPI運用の確立、自動化の深化、V3ダッシュボード完成");

  // Current assessment
  s.addText("技術基盤の評価", { x: 0.4, y: 0.75, w: 4, h: 0.25, fontSize: 12, fontFace: "Arial", color: C.NAVY, bold: true });

  const assess = [
    { item: "自動化パイプライン", status: "良好", c: C.GREEN },
    { item: "スナップショット管理", status: "良好", c: C.GREEN },
    { item: "レポート品質", status: "良好", c: C.GREEN },
    { item: "KPI目標設定", status: "問題あり", c: C.RED },
    { item: "V3完成度", status: "未完", c: C.ORANGE },
    { item: "データ更新頻度", status: "要検討", c: C.ORANGE },
  ];
  assess.forEach((a, i) => {
    const y = 1.1 + i * 0.28;
    s.addShape(pptx.shapes.RECTANGLE, { x: 0.5, y: y + 0.05, w: 0.15, h: 0.15, fill: { color: a.c } });
    s.addText(a.item, { x: 0.8, y, w: 2, h: 0.25, fontSize: 9, fontFace: "Arial", color: C.TEXT });
    s.addText(a.status, { x: 2.8, y, w: 1.2, h: 0.25, fontSize: 9, fontFace: "Arial", color: a.c, bold: true });
  });

  // KPI targets
  s.addText("提案するKPI目標（2026年9月）", { x: 5.2, y: 0.75, w: 4.5, h: 0.25, fontSize: 12, fontFace: "Arial", color: C.NAVY, bold: true });

  const kpis = [
    { label: "平均スコア", now: "8.39", target: "8.60", delta: "+0.21" },
    { label: "清掃課題率", now: "4.7%", target: "3.5%", delta: "-1.2pt" },
    { label: "高評価率", now: "78.5%", target: "82.0%", delta: "+3.5pt" },
    { label: "低評価率", now: "4.5%", target: "3.0%", delta: "-1.5pt" },
  ];
  kpis.forEach((k, i) => {
    const y = 1.1 + i * 0.42;
    s.addShape(pptx.shapes.RECTANGLE, { x: 5.3, y, w: 4.2, h: 0.36, fill: { color: i % 2 === 0 ? "F8FAFC" : C.WHITE }, shadow: shadow() });
    s.addText(k.label, { x: 5.35, y, w: 1.2, h: 0.36, fontSize: 9, fontFace: "Arial", color: C.NAVY, bold: true, valign: "middle" });
    s.addText(k.now, { x: 6.55, y, w: 0.8, h: 0.36, fontSize: 9, fontFace: "Arial", color: C.SUBTEXT, align: "center", valign: "middle" });
    s.addText("→", { x: 7.35, y, w: 0.3, h: 0.36, fontSize: 10, fontFace: "Arial", color: C.TEXT, align: "center", valign: "middle" });
    s.addText(k.target, { x: 7.65, y, w: 0.8, h: 0.36, fontSize: 10, fontFace: "Arial", color: C.GREEN, bold: true, align: "center", valign: "middle" });
    s.addText(k.delta, { x: 8.45, y, w: 0.9, h: 0.36, fontSize: 9, fontFace: "Arial", color: C.GREEN, align: "center", valign: "middle" });
  });

  // Action items
  s.addText("4つの強化施策", { x: 0.4, y: 2.95, w: 9, h: 0.3, fontSize: 13, fontFace: "Arial", color: C.NAVY, bold: true });

  const ops = [
    { n: "1", t: "月次レビュー会議", d: "ダッシュボードで15分レビュー\n閾値下回りアラート＋対応協議\n四半期ごとの目標見直し", c: C.BLUE },
    { n: "2", t: "データ更新自動化", d: "cronで週次/月次自動更新\n口コミ自動取得パイプライン\nSlack/メール通知アラート", c: C.TEAL },
    { n: "3", t: "V3ダッシュボード完成", d: "ES Dashboard（スタッフ管理）\nスキルマッピング\n人員配置最適化", c: C.GREEN },
    { n: "4", t: "差分レポート活用", d: "前月比の成果可視化\n投資効果の証明\n月次自動送信で契約継続促進", c: C.ACCENT },
  ];
  ops.forEach((o, i) => {
    const x = 0.3 + i * 2.4;
    s.addShape(pptx.shapes.RECTANGLE, { x, y: 3.3, w: 2.2, h: 1.7, fill: { color: C.WHITE }, shadow: shadow() });
    s.addShape(pptx.shapes.RECTANGLE, { x, y: 3.3, w: 2.2, h: 0.04, fill: { color: o.c } });
    s.addShape(pptx.shapes.OVAL, { x: x + 0.85, y: 3.4, w: 0.45, h: 0.45, fill: { color: o.c } });
    s.addText(o.n, { x: x + 0.85, y: 3.4, w: 0.45, h: 0.45, fontSize: 18, fontFace: "Arial", color: C.WHITE, bold: true, align: "center", valign: "middle" });
    s.addText(o.t, { x: x + 0.05, y: 3.9, w: 2.1, h: 0.3, fontSize: 10, fontFace: "Arial", color: o.c, bold: true, align: "center" });
    s.addText(o.d, { x: x + 0.1, y: 4.2, w: 2.0, h: 0.7, fontSize: 8, fontFace: "Arial", color: C.TEXT, align: "center", lineSpacingMultiple: 1.3 });
  });

  addFooter(s, 7);
})();

// ============================================================
// Slide 8: Priority Matrix
// ============================================================
(function () {
  const s = pptx.addSlide();
  addHeader(s, "総合ロードマップ：優先順位マトリクス", "インパクト × 実現容易性で施策を分類");

  const items = [
    { t: "HPにデータ実績掲載", impact: "高", ease: "高", pri: "最優先", c: C.RED },
    { t: "KPI目標値の修正", impact: "中", ease: "高", pri: "最優先", c: C.RED },
    { t: "要改善5ホテル集中改善", impact: "高", ease: "中", pri: "高", c: C.ORANGE },
    { t: "営業にダッシュボード活用", impact: "高", ease: "高", pri: "高", c: C.ORANGE },
    { t: "ベストプラクティス横展開", impact: "中", ease: "中", pri: "中", c: C.BLUE },
    { t: "顧客向けポータル提供", impact: "高", ease: "低", pri: "中", c: C.BLUE },
    { t: "ブログの業界専門化", impact: "中", ease: "中", pri: "中", c: C.BLUE },
    { t: "FC募集ページ充実", impact: "中", ease: "中", pri: "中", c: C.BLUE },
    { t: "SaaS化", impact: "最高", ease: "低", pri: "長期", c: C.TEAL },
  ];

  const tblRows = [["施策", "インパクト", "容易性", "優先度"]];
  items.forEach(it => tblRows.push([it.t, it.impact, it.ease, it.pri]));

  s.addTable(tblRows, {
    x: 0.3, y: 0.8, w: 9.4, fontSize: 9.5, fontFace: "Arial",
    border: { type: "solid", pt: 0.5, color: "CCCCCC" },
    colW: [4.0, 1.5, 1.5, 2.4],
    rowH: [0.3, 0.28, 0.28, 0.28, 0.28, 0.28, 0.28, 0.28, 0.28, 0.28],
    autoPage: false, color: C.TEXT,
  });
  s.addShape(pptx.shapes.RECTANGLE, { x: 0.3, y: 0.8, w: 9.4, h: 0.3, fill: { color: C.NAVY } });
  s.addText("施策                                        インパクト        容易性           優先度", { x: 0.4, y: 0.8, w: 9.2, h: 0.3, fontSize: 9, fontFace: "Arial", color: C.WHITE, bold: true, valign: "middle" });

  addFooter(s, 8);
})();

// ============================================================
// Slide 9: Timeline
// ============================================================
(function () {
  const s = pptx.addSlide();
  addHeader(s, "実行タイムライン", "4フェーズで段階的に実行");

  const phases = [
    {
      n: "Phase 1", when: "即時〜1ヶ月", label: "Quick Wins",
      items: "・HPトップに数値実績掲載\n・KPI目標値の設定・反映\n・4 URGENTホテルにQCマネージャー配置\n・営業資料にダッシュボード追加",
      c: C.RED
    },
    {
      n: "Phase 2", when: "1〜3ヶ月", label: "Foundation",
      items: "・HP事例ページ新設\n・ブログ戦略転換（業界分析）\n・月次レビュー会議導入\n・ベストプラクティス文書化・横展開\n・FC募集ページ充実",
      c: C.ORANGE
    },
    {
      n: "Phase 3", when: "3〜6ヶ月", label: "Scale",
      items: "・顧客向けポータル開発・提供開始\n・データ更新の完全自動化\n・V3ダッシュボード完成\n・要改善ホテルの8.0超え検証\n・10周年ブランディング",
      c: C.BLUE
    },
    {
      n: "Phase 4", when: "6〜12ヶ月", label: "Transform",
      items: "・SaaS型口コミ分析サービスMVP\n・ホテルチェーン本部へ包括提案\n・ストック型収入モデルの確立\n・新規市場（旅館・民泊）への展開検討",
      c: C.GREEN
    },
  ];

  // Timeline line
  s.addShape(pptx.shapes.RECTANGLE, { x: 0.5, y: 1.55, w: 9.0, h: 0.04, fill: { color: C.NAVY } });

  phases.forEach((ph, i) => {
    const x = 0.3 + i * 2.4;
    // Timeline dot
    s.addShape(pptx.shapes.OVAL, { x: x + 0.95, y: 1.4, w: 0.3, h: 0.3, fill: { color: ph.c } });

    // Phase header
    s.addText(ph.n, { x, y: 0.8, w: 2.2, h: 0.25, fontSize: 12, fontFace: "Arial", color: ph.c, bold: true, align: "center" });
    s.addText(ph.when, { x, y: 1.05, w: 2.2, h: 0.2, fontSize: 9, fontFace: "Arial", color: C.SUBTEXT, align: "center" });
    s.addText(ph.label, { x: x + 0.8, y: 1.38, w: 0.6, h: 0.3, fontSize: 7, fontFace: "Arial", color: C.WHITE, bold: true, align: "center", valign: "middle" });

    // Phase card
    s.addShape(pptx.shapes.RECTANGLE, { x, y: 1.85, w: 2.2, h: 2.8, fill: { color: C.WHITE }, shadow: shadow() });
    s.addShape(pptx.shapes.RECTANGLE, { x, y: 1.85, w: 2.2, h: 0.04, fill: { color: ph.c } });
    s.addText(ph.items, { x: x + 0.1, y: 1.95, w: 2.0, h: 2.5, fontSize: 8.5, fontFace: "Arial", color: C.TEXT, lineSpacingMultiple: 1.5, valign: "top" });
  });

  addFooter(s, 9);
})();

// ============================================================
// Slide 10: Closing
// ============================================================
(function () {
  const s = pptx.addSlide();
  s.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: C.NAVY } });
  s.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0, w: 0.08, h: 5.625, fill: { color: C.ACCENT } });

  s.addText("まとめ", { x: 0.8, y: 1.0, w: 8.4, h: 0.5, fontSize: 16, fontFace: "Arial", color: "5A9BE6" });
  s.addText("PRIMECHANGEの最大の武器は、", { x: 0.8, y: 1.6, w: 8.4, h: 0.5, fontSize: 24, fontFace: "Arial", color: C.WHITE, bold: true });
  s.addText("清掃業界で類を見ない", { x: 0.8, y: 2.1, w: 8.4, h: 0.5, fontSize: 24, fontFace: "Arial", color: C.WHITE, bold: true });
  s.addText("「データ分析基盤」です。", { x: 0.8, y: 2.6, w: 8.4, h: 0.6, fontSize: 28, fontFace: "Arial", color: C.ACCENT, bold: true });

  s.addShape(pptx.shapes.RECTANGLE, { x: 0.8, y: 3.3, w: 4.0, h: 0.04, fill: { color: C.ACCENT } });

  s.addText(
    "これを内部ツールに留めず、\n営業・顧客維持・新規事業の全てに活用する\n戦略にシフトすべきです。",
    { x: 0.8, y: 3.5, w: 8.4, h: 1.0, fontSize: 16, fontFace: "Arial", color: "CBD5E1", lineSpacingMultiple: 1.6 }
  );

  s.addText("株式会社PRIMECHANGE", { x: 0.8, y: 4.8, w: 8, h: 0.3, fontSize: 10, fontFace: "Arial", color: "475569" });
  s.addText("2026年3月17日", { x: 0.8, y: 5.1, w: 4, h: 0.2, fontSize: 9, fontFace: "Arial", color: "475569" });
})();

// ============================================================
// Save
// ============================================================
pptx.writeFile({ fileName: OUTPUT }).then(() => {
  console.log(`✓ PPTX saved: ${OUTPUT}`);
}).catch(e => { console.error(e); process.exit(1); });
