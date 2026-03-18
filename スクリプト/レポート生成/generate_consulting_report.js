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
// Data
// ============================================================
const ROOT = path.resolve(__dirname, "../..");
const hotelsRanked = JSON.parse(fs.readFileSync(path.join(ROOT, "hotel-ranked.json"), "utf-8"));
const OUTPUT = path.resolve(ROOT, "納品レポート/PRIMECHANGE_経営コンサルティングレポート.docx");

const totalReviews = 2070;
const avgScore = 8.39;
const highRate = 78.5;
const cleaningIssueRate = 4.7;
const monthlyRevenue = 99850000;

// ============================================================
// Styles
// ============================================================
const C = { NAVY: "1B3A5C", ACCENT: "C2333A", WHITE: "FFFFFF", LIGHT_BG: "F5F7FA", TEXT: "333333", SUBTEXT: "666666", GREEN: "27AE60", ORANGE: "FF9800", RED: "E74C3C", BLUE: "2E75B6" };
const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };
const cm = { top: 60, bottom: 60, left: 100, right: 100 };

function h1(t) { return new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 400, after: 200 }, children: [new TextRun({ text: t, bold: true, size: 32, font: "Arial", color: C.NAVY })] }); }
function h2(t) { return new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 300, after: 160 }, children: [new TextRun({ text: t, bold: true, size: 26, font: "Arial", color: C.BLUE })] }); }
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
function headerRow(cols) { return new TableRow({ children: cols.map(c => cell(c, { bg: C.NAVY, c: C.WHITE, b: true })) }); }
const PB = () => new Paragraph({ children: [new PageBreak()] });

// ============================================================
// Cover
// ============================================================
function buildCover() {
  return [
    new Paragraph({ spacing: { before: 2400 } }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 }, children: [new TextRun({ text: "PRIMECHANGE", size: 56, font: "Arial", bold: true, color: C.NAVY })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [new TextRun({ text: "経営コンサルティングレポート", size: 44, font: "Arial", bold: true, color: C.ACCENT })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 600 }, children: [new TextRun({ text: "Management Consulting Report — Data-Driven Growth Strategy", size: 22, font: "Arial", color: C.SUBTEXT, italics: true })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [new TextRun({ text: `対象: ${hotelsRanked.length}ホテル | 口コミ: ${totalReviews}件 | 月間総売上: ¥9,985万`, size: 20, font: "Arial", color: C.TEXT })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [new TextRun({ text: "2026年3月17日", size: 22, font: "Arial", color: C.TEXT })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 800 }, children: [new TextRun({ text: "株式会社PRIMECHANGE", size: 28, font: "Arial", bold: true, color: C.NAVY })] }),
    PB(),
  ];
}

// ============================================================
// Section 1: Executive Summary
// ============================================================
function buildExecSummary() {
  return [
    h1("1. エグゼクティブサマリー"),
    p("本レポートは、株式会社PRIMECHANGEの事業データ・ダッシュボード資産・ホームページを経営コンサルティングの観点から分析し、中長期的な成長戦略を提言するものです。"),
    h3("核心的発見"),
    bp("PRIMECHANGEは清掃会社でありながら、業界で極めて稀な「データ分析基盤」を保有している。6つのOTAから2,070件の口コミを構造化分析し、7種類の深掘り分析と79本の自動レポートを生成するパイプラインは、そのまま新規事業になり得る知的資産である。", { b: true }),
    bp("しかし、この強みはホームページにも営業活動にもほとんど活かされておらず、対外的な訴求は「本気の清掃」「ベンチャーの柔軟性」という定性的メッセージに留まっている。"),
    bp("口コミスコアと売上には統計的に有意な相関（RevPAR相関 r=0.60）があり、スコア0.5点改善でポートフォリオ全体の年間売上が+1.47億円増加する試算。清掃は「コスト」ではなく「売上向上投資」として再定義すべきである。", { b: true }),
    h3("5つの提言領域"),
    bp("領域1：データ資産の事業化 — 内部ツールを営業武器・顧客向けサービス・SaaSへ段階的に事業化"),
    bp("領域2：品質→売上のROIストーリー — ホテルオーナーへの定量的な価値提案"),
    bp("領域3：ホームページ・営業戦略の刷新 — データドリブンなブランディングへの転換"),
    bp("領域4：ポートフォリオ最適化 — 19ホテルのティア別リソース配分戦略"),
    bp("領域5：組織・オペレーション強化 — KPI運用・自動化・V3完成"),
    PB(),
  ];
}

// ============================================================
// Section 2: Data Asset Monetization
// ============================================================
function buildDataAsset() {
  return [
    h1("2. 領域1：データ資産の事業化"),
    p("PRIMECHANGEが保有するデジタル資産は、清掃業界において極めてユニークな競争優位の源泉です。"),
    h2("2.1 保有資産の棚卸し"),
    new Table({ rows: [
      headerRow(["資産", "内容", "競合との比較"]),
      new TableRow({ children: [cell("口コミ分析基盤"), cell("6 OTAから2,070件を構造化分析"), cell("業界で類を見ない", { c: C.GREEN, b: true })] }),
      new TableRow({ children: [cell("ダッシュボード"), cell("V1(6P)→V2(9P)→V3(10P)の3世代"), cell("通常の清掃会社は未保有")] }),
      new TableRow({ children: [cell("自動レポート生成"), cell("79本のDOCX/PPTXを自動生成"), cell("一般に手作業で月1本程度")] }),
      new TableRow({ children: [cell("7種深掘り分析"), cell("クレーム分類・人員配置・売上弾性等"), cell("コンサルファームレベル")] }),
    ]}),
    h2("2.2 3段階の事業化ロードマップ"),
    h3("Stage 1（即時）：営業ツールとして活用"),
    bp("新規ホテル獲得の提案時にダッシュボードのデモを実施"),
    bp("既存79本のレポートを匿名化サンプルとして営業資料に活用"),
    bp("「御社のホテルも、このレベルの分析が可能になります」という具体的価値提案"),
    p("期待効果：新規受注率の向上、競合清掃会社との明確な差別化", { b: true, c: C.GREEN }),
    h3("Stage 2（3-6ヶ月）：顧客向けポータルとして提供"),
    bp("各ホテルに専用ログインを付与し、自社のスコア推移をリアルタイム確認可能に"),
    bp("月次自動レポートのメール配信機能"),
    bp("価格設定例：清掃契約にバンドルし月額5-10万円のアップセル"),
    p("期待効果：顧客のスイッチングコスト向上（解約するとデータが見れなくなる）→ 解約率低下", { b: true, c: C.GREEN }),
    h3("Stage 3（6-12ヶ月）：SaaS型サービスとして独立事業化"),
    bp("清掃契約がないホテルにも口コミ分析サービスを単体提供"),
    bp("ホテルチェーン本部向けポートフォリオ管理ツール"),
    bp("価格設定例：月額3-8万円/ホテル（ボリュームディスカウント）"),
    p("期待効果：ストック型収入の獲得、清掃事業の景気変動リスクのヘッジ", { b: true, c: C.GREEN }),
    PB(),
  ];
}

// ============================================================
// Section 3: Quality → Revenue ROI
// ============================================================
function buildROI() {
  const urgentHotels = hotelsRanked.filter(h => h.priority === "URGENT");
  return [
    h1("3. 領域2：清掃品質 → 売上のROIストーリー"),
    p("回帰分析により、口コミスコアと売上指標には統計的に有意な相関関係が確認されています。"),
    h2("3.1 回帰分析結果"),
    new Table({ rows: [
      headerRow(["指標", "スコア1点あたり変動", "相関係数 r", "R²", "解釈"]),
      new TableRow({ children: [cell("稼働率"), cell("+3.9%pt"), cell("0.34"), cell("0.12"), cell("弱〜中程度の相関")] }),
      new TableRow({ children: [cell("ADR（客室単価）"), cell("+¥255"), cell("0.56"), cell("0.31"), cell("中程度の相関")] }),
      new TableRow({ children: [cell("RevPAR"), cell("+¥230"), cell("0.60"), cell("0.36"), cell("中〜強の相関", { b: true })] }),
      new TableRow({ children: [cell("1室あたり月間売上"), cell("+¥6,453"), cell("0.60"), cell("0.36"), cell("中〜強の相関", { b: true })] }),
    ]}),
    h2("3.2 スコア帯別パフォーマンス"),
    new Table({ rows: [
      headerRow(["スコア帯", "ホテル数", "平均稼働率", "平均ADR", "平均RevPAR", "最高スコア帯との差"]),
      new TableRow({ children: [cell("7.0-8.0"), cell("5"), cell("69.4%"), cell("¥1,024"), cell("¥711"), cell("-35%", { c: C.RED, b: true })] }),
      new TableRow({ children: [cell("8.0-8.5"), cell("3"), cell("72.5%"), cell("¥1,229"), cell("¥890"), cell("-18%", { c: C.ORANGE })] }),
      new TableRow({ children: [cell("8.5-9.0"), cell("7"), cell("76.9%"), cell("¥1,331"), cell("¥1,029"), cell("-6%")] }),
      new TableRow({ children: [cell("9.0+", { b: true }), cell("4"), cell("75.3%"), cell("¥1,461"), cell("¥1,092", { b: true }), cell("基準", { c: C.GREEN, b: true })] }),
    ]}),
    h2("3.3 改善シナリオ別インパクト"),
    new Table({ rows: [
      headerRow(["シナリオ", "スコア改善幅", "RevPAR変動", "月間売上変動", "年間売上変動"]),
      new TableRow({ children: [cell("A：最小投資"), cell("+0.1点"), cell("+¥23 (+2.5%)"), cell("+¥246万"), cell("+¥2,949万")] }),
      new TableRow({ children: [cell("B：中程度投資"), cell("+0.3点"), cell("+¥69 (+7.4%)"), cell("+¥737万"), cell("+¥8,846万", { b: true })] }),
      new TableRow({ children: [cell("C：本格投資", { b: true }), cell("+0.5点"), cell("+¥115 (+12.3%)"), cell("+¥1,229万"), cell("+¥1億4,743万", { c: C.GREEN, b: true })] }),
    ]}),
    h2("3.4 核心メッセージ"),
    p("「PRIMECHANGEに清掃を任せると、口コミスコアが上がり、それが直接売上に反映される」", { b: true, sz: 24, c: C.ACCENT }),
    p("これは単なる「清掃外注」ではなく、「売上向上投資」としてのポジショニングです。ホテルオーナーへの提案時、このROIデータを前面に出すことで、価格競争から脱却し、価値ベースの契約交渉が可能になります。"),
    PB(),
  ];
}

// ============================================================
// Section 4: Homepage & Sales Strategy Gap
// ============================================================
function buildHPGap() {
  return [
    h1("4. 領域3：ホームページと営業戦略のギャップ"),
    p("現在のホームページを分析した結果、PRIMECHANGEの最大の強みであるデータ分析力が対外的に全く伝わっていないことが判明しました。"),
    h2("4.1 現状の問題点"),
    new Table({ rows: [
      headerRow(["要素", "現状", "問題"]),
      new TableRow({ children: [cell("メッセージ"), cell("「本気の清掃」「ベンチャーの柔軟性」"), cell("抽象的で競合と差別化できない")] }),
      new TableRow({ children: [cell("実績"), cell("記載なし"), cell("データ分析力が見えない")] }),
      new TableRow({ children: [cell("ブログ"), cell("「人生は逆」「GIVE and TAKE」等"), cell("自己啓発的で事業との関連が薄い")] }),
      new TableRow({ children: [cell("FC募集"), cell("バナーのみ"), cell("収益モデルや支援内容なし")] }),
      new TableRow({ children: [cell("会社概要"), cell("住所・電話のみ"), cell("設立年・代表・従業員数なし")] }),
    ]}),
    h2("4.2 改善提案（優先順位順）"),
    h3("提案1（最優先）：トップページにデータ実績を掲載"),
    bp("「19ホテル管理」「2,070件の口コミ分析」「平均スコア8.39/10」を数値で訴求"),
    bp("ダッシュボードのスクリーンショットを掲載"),
    bp("「清掃品質の見える化で、口コミスコア向上 → 売上向上を実現」というメッセージ"),
    h3("提案2：事例ページの新設"),
    bp("匿名化したBefore/After事例（スコア改善実績）"),
    bp("ダッシュボードの画面デモ動画"),
    bp("レポートサンプルのダウンロード"),
    h3("提案3：ブログの戦略的活用"),
    bp("自己啓発記事 → ホテル業界の口コミトレンド分析記事に転換"),
    bp("SEOキーワード：「ホテル清掃 品質管理」「OTA口コミ改善」「ホテル清掃 外注」"),
    bp("月1-2本で業界専門メディアとしてのポジション構築"),
    h3("提案4：FC募集ページの充実"),
    bp("収益モデル（初期投資、月商予測、利益率）の開示"),
    bp("本部支援内容（データ分析ツール、研修、品質管理ノウハウ）"),
    bp("既存FCオーナーの声（テスティモニアル）"),
    h3("提案5：10周年ブランディング"),
    bp("2026年で10周年を迎えるタイミングを活用"),
    bp("記念キャンペーン、PR活動、メディア露出の企画"),
    PB(),
  ];
}

// ============================================================
// Section 5: Portfolio Strategy
// ============================================================
function buildPortfolio() {
  const tiers = { "優秀": [], "良好": [], "概ね良好": [], "要改善": [] };
  hotelsRanked.forEach(h => { if (tiers[h.tier]) tiers[h.tier].push(h); });

  const els = [
    h1("5. 領域4：ポートフォリオ戦略（19ホテルの最適化）"),
    h2("5.1 ポートフォリオ構造"),
    new Table({ rows: [
      headerRow(["ティア", "ホテル数", "平均スコア", "平均清掃課題率", "戦略方針"]),
      new TableRow({ children: [cell("優秀（9.0+）", { c: C.GREEN, b: true }), cell(String(tiers["優秀"].length)), cell("9.10"), cell("0.6%"), cell("ベストプラクティス抽出・横展開")] }),
      new TableRow({ children: [cell("良好（8.5-9.0）"), cell(String(tiers["良好"].length)), cell("8.74"), cell("4.0%"), cell("維持・微調整")] }),
      new TableRow({ children: [cell("概ね良好（8.0-8.5）"), cell(String(tiers["概ね良好"].length)), cell("8.27"), cell("3.6%"), cell("重点監視・予防的改善")] }),
      new TableRow({ children: [cell("要改善（<8.0）", { c: C.RED, b: true }), cell(String(tiers["要改善"].length)), cell("7.41"), cell("11.1%"), cell("緊急集中改善プログラム", { c: C.RED, b: true })] }),
    ]}),
    h2("5.2 19ホテル全一覧"),
  ];

  const hotelRows = [headerRow(["順位", "ホテル名", "スコア", "口コミ数", "高評価率", "清掃課題率", "優先度", "ティア"])];
  hotelsRanked.forEach(h => {
    const prioColor = h.priority === "URGENT" ? C.RED : h.priority === "HIGH" ? C.ORANGE : C.TEXT;
    hotelRows.push(new TableRow({ children: [
      cell(String(h.rank)), cell(h.name, { a: AlignmentType.LEFT }),
      cell(h.avg.toFixed(2)), cell(String(h.total_reviews)),
      cell(h.high_rate + "%"), cell(h.cleaning_issue_rate + "%", { c: h.cleaning_issue_rate >= 9 ? C.RED : C.TEXT }),
      cell(h.priority, { c: prioColor, b: true }), cell(h.tier),
    ]}));
  });
  els.push(new Table({ rows: hotelRows }));

  els.push(
    h2("5.3 戦略的アクション"),
    h3("A. 要改善5ホテルへの集中改善プログラム"),
    p("特にコンフォートホテル博多（清掃課題率14.9%）は全ポートフォリオ最悪。口コミでの悪影響がブランド全体に波及するリスクがあります。"),
    bp("即時アクション：専任QCマネージャーの配置（2週間以内）"),
    bp("清掃課題の11カテゴリ（臭気、カビ、排水、害虫等）への個別対策"),
    bp("目標：6ヶ月以内にスコア8.0超えを達成"),
    bp("投資対効果：5ホテル合計で月額+292万円（スコア+0.5の場合）", { b: true }),
    h3("B. 優秀ホテルのベストプラクティス横展開"),
    p("ERA東神田（9.16）とスイーツ東京ベイ（9.15）の清掃課題率はわずか0.8-0.9%。"),
    bp("これらのホテルの清掃手順・チェックリスト・スタッフ配置を文書化"),
    bp("「PRIMECHANGE清掃スタンダード」として全ホテルに展開"),
    bp("優秀ホテルのチームリーダーによる要改善ホテルへの巡回指導"),
    h3("C. 新規ホテル獲得のターゲティング"),
    bp("理想ターゲット：現在スコア7.0-8.0の中価格帯ビジネスホテル（改善余地が大きい）"),
    bp("営業ピッチ：現在のスコアを分析し、清掃改善による売上増をシミュレーション提示"),
    bp("避けるべきターゲット：超高級ホテル（清掃以外の要素が大きい）、スコア6.0未満（設備老朽化の可能性）"),
    h3("D. ホテルブランド別の戦略的アプローチ"),
    p("ポートフォリオにコンフォート系6ホテルが含まれている点に注目。チェーン本部への一括提案により「全コンフォート系列にデータ分析+清掃を包括受注」する戦略が有効です。"),
    PB(),
  );
  return els;
}

// ============================================================
// Section 6: Org & Operations
// ============================================================
function buildOrgOps() {
  return [
    h1("6. 領域5：組織・オペレーション強化"),
    h2("6.1 技術基盤の評価"),
    new Table({ rows: [
      headerRow(["項目", "現状評価", "コメント"]),
      new TableRow({ children: [cell("自動化パイプライン"), cell("良好", { c: C.GREEN }), cell("スプレッドシート→分析→レポート生成が自動化済")] }),
      new TableRow({ children: [cell("スナップショット管理"), cell("良好", { c: C.GREEN }), cell("時系列でデータ変化を追跡")] }),
      new TableRow({ children: [cell("レポート品質"), cell("良好", { c: C.GREEN }), cell("7章DOCX + 10スライドPPTXの自動生成")] }),
      new TableRow({ children: [cell("KPI目標設定"), cell("問題あり", { c: C.RED, b: true }), cell("ポートフォリオ目標値がundefinedのまま")] }),
      new TableRow({ children: [cell("V3完成度"), cell("未完", { c: C.ORANGE }), cell("ES Dashboard等の一部ページが未実装")] }),
      new TableRow({ children: [cell("データ更新頻度"), cell("要検討", { c: C.ORANGE }), cell("現在は手動実行")] }),
    ]}),
    h2("6.2 KPI目標の設定（即時対応）"),
    p("action-plans-content.json のKPI目標が全て undefined になっているのは重大な問題です。以下の目標値を提案します："),
    new Table({ rows: [
      headerRow(["指標", "現在値", "6ヶ月目標（2026年9月）", "改善幅"]),
      new TableRow({ children: [cell("平均スコア"), cell("8.39"), cell("8.60", { c: C.GREEN, b: true }), cell("+0.21")] }),
      new TableRow({ children: [cell("清掃課題率"), cell("4.7%"), cell("3.5%", { c: C.GREEN, b: true }), cell("-1.2pt")] }),
      new TableRow({ children: [cell("高評価率"), cell("78.5%"), cell("82.0%", { c: C.GREEN, b: true }), cell("+3.5pt")] }),
      new TableRow({ children: [cell("低評価率"), cell("4.5%"), cell("3.0%", { c: C.GREEN, b: true }), cell("-1.5pt")] }),
    ]}),
    h2("6.3 オペレーション強化施策"),
    h3("施策1：月次レビュー会議の導入"),
    bp("ダッシュボードを使って15分でポートフォリオ全体をレビュー"),
    bp("スコアが閾値を下回ったホテルのアラートと対応協議"),
    bp("四半期ごとの目標見直し"),
    h3("施策2：データ更新の自動化・定期化"),
    bp("full_update_with_reports.sh のcron設定（週次または月次）"),
    bp("新しい口コミの自動取得パイプライン構築"),
    bp("Slack/メール通知：スコアが閾値を下回った場合のアラート"),
    h3("施策3：V3ダッシュボードの完成"),
    bp("ES Dashboard（スタッフ管理）：スキルマッピング、最適人員配置"),
    bp("研修履歴と品質スコアの相関分析"),
    h3("施策4：差分レポートの戦略的活用"),
    bp("「前月比でスコアが+0.2向上しました」→ 成果の可視化"),
    bp("「清掃課題が3件減少しました」→ 投資効果の証明"),
    bp("ホテルオーナーに月次で自動送信 → 契約継続の動機付け"),
    PB(),
  ];
}

// ============================================================
// Section 7: Priority Matrix & Roadmap
// ============================================================
function buildRoadmap() {
  return [
    h1("7. 総合ロードマップ：優先順位マトリクス"),
    h2("7.1 施策の優先順位"),
    new Table({ rows: [
      headerRow(["施策", "インパクト", "実現容易性", "優先度"]),
      new TableRow({ children: [cell("HPにデータ実績を掲載", { a: AlignmentType.LEFT }), cell("高"), cell("高"), cell("最優先", { c: C.RED, b: true })] }),
      new TableRow({ children: [cell("KPI目標値のundefined修正", { a: AlignmentType.LEFT }), cell("中"), cell("高"), cell("最優先", { c: C.RED, b: true })] }),
      new TableRow({ children: [cell("要改善5ホテルへの集中改善", { a: AlignmentType.LEFT }), cell("高"), cell("中"), cell("高", { c: C.ORANGE, b: true })] }),
      new TableRow({ children: [cell("営業提案にダッシュボード活用", { a: AlignmentType.LEFT }), cell("高"), cell("高"), cell("高", { c: C.ORANGE, b: true })] }),
      new TableRow({ children: [cell("ベストプラクティス横展開", { a: AlignmentType.LEFT }), cell("中"), cell("中"), cell("中")] }),
      new TableRow({ children: [cell("顧客向けポータル提供", { a: AlignmentType.LEFT }), cell("高"), cell("低"), cell("中")] }),
      new TableRow({ children: [cell("ブログの業界専門化", { a: AlignmentType.LEFT }), cell("中"), cell("中"), cell("中")] }),
      new TableRow({ children: [cell("FC募集ページ充実", { a: AlignmentType.LEFT }), cell("中"), cell("中"), cell("中")] }),
      new TableRow({ children: [cell("SaaS化", { a: AlignmentType.LEFT }), cell("最高"), cell("低"), cell("長期", { c: C.BLUE })] }),
    ]}),
    h2("7.2 タイムライン"),
    h3("Phase 1：即時〜1ヶ月（Quick Wins）"),
    bp("ホームページのトップに数値実績（19ホテル/2,070件/8.39点）を掲載"),
    bp("KPI目標値の設定とダッシュボードへの反映"),
    bp("博多・浜松町・新横浜・蒲田の4 URGENTホテルに専任QCマネージャー配置"),
    bp("営業提案資料にダッシュボードのスクリーンショット・レポートサンプルを追加"),
    h3("Phase 2：1〜3ヶ月（Foundation）"),
    bp("HP事例ページの新設（匿名化Before/After事例）"),
    bp("ブログのテーマ転換（自己啓発 → 業界分析）"),
    bp("月次レビュー会議の導入"),
    bp("ベストプラクティスの文書化と横展開開始"),
    bp("FC募集ページの充実"),
    h3("Phase 3：3〜6ヶ月（Scale）"),
    bp("顧客向けポータルの開発・提供開始"),
    bp("データ更新の完全自動化（cron + アラート通知）"),
    bp("V3ダッシュボードの完成"),
    bp("要改善ホテルのスコア8.0超え達成を検証"),
    bp("10周年ブランディングキャンペーンの実施"),
    h3("Phase 4：6〜12ヶ月（Transform）"),
    bp("SaaS型口コミ分析サービスの企画・MVP開発"),
    bp("ホテルチェーン本部への包括提案"),
    bp("ストック型収入モデルの確立"),
  ];
}

// ============================================================
// Build & Output
// ============================================================
async function main() {
  const children = [
    ...buildCover(),
    ...buildExecSummary(),
    ...buildDataAsset(),
    ...buildROI(),
    ...buildHPGap(),
    ...buildPortfolio(),
    ...buildOrgOps(),
    ...buildRoadmap(),
  ];

  const doc = new Document({
    creator: "PRIMECHANGE",
    title: "PRIMECHANGE 経営コンサルティングレポート",
    description: "データ駆動型成長戦略の提言",
    styles: {
      default: { document: { run: { font: "Arial", size: 20, color: C.TEXT } } },
    },
    sections: [{
      properties: {
        page: { margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } },
      },
      headers: {
        default: new Header({
          children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [
            new TextRun({ text: "PRIMECHANGE 経営コンサルティングレポート | Confidential", size: 14, font: "Arial", color: C.SUBTEXT, italics: true }),
          ]})]
        }),
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [
            new TextRun({ text: "- ", size: 14, font: "Arial", color: C.SUBTEXT }),
            new TextRun({ children: [PageNumber.CURRENT], size: 14, font: "Arial", color: C.SUBTEXT }),
            new TextRun({ text: " -", size: 14, font: "Arial", color: C.SUBTEXT }),
          ]})]
        }),
      },
      children,
    }],
  });

  const buf = await Packer.toBuffer(doc);
  fs.mkdirSync(path.dirname(OUTPUT), { recursive: true });
  fs.writeFileSync(OUTPUT, buf);
  console.log(`✓ DOCX saved: ${OUTPUT}`);
}

main().catch(e => { console.error(e); process.exit(1); });
