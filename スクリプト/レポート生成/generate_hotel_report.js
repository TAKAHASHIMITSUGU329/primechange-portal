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
// CLI Arguments
// ============================================================
const args = process.argv.slice(2);
if (args.length < 3) {
  console.error("Usage: node generate_hotel_report.js <hotel_name> <analysis_json> <output_prefix>");
  console.error('Example: node generate_hotel_report.js "コンフォートイン六本木" comfort_roppongi_analysis.json comfort_roppongi');
  process.exit(1);
}

const HOTEL_NAME = args[0];
const ANALYSIS_JSON_PATH = path.resolve(args[1]);
const OUTPUT_PREFIX = args[2];
const OUTPUT_DIR = path.dirname(ANALYSIS_JSON_PATH);

// ============================================================
// Load analysis data
// ============================================================
if (!fs.existsSync(ANALYSIS_JSON_PATH)) {
  console.error("Error: Analysis JSON file not found: " + ANALYSIS_JSON_PATH);
  process.exit(1);
}

const data = JSON.parse(fs.readFileSync(ANALYSIS_JSON_PATH, "utf-8"));

// ============================================================
// Date and period
// ============================================================
const now = new Date();
const REPORT_DATE = `${now.getFullYear()}年${now.getMonth() + 1}月${now.getDate()}日`;

// Determine analysis period from comments
function getAnalysisPeriod(comments) {
  if (!comments || comments.length === 0) return `${now.getFullYear()}年`;
  const dates = comments
    .filter(c => c.date)
    .map(c => new Date(c.date))
    .filter(d => !isNaN(d.getTime()))
    .sort((a, b) => a - b);
  if (dates.length === 0) return `${now.getFullYear()}年`;
  const first = dates[0];
  const last = dates[dates.length - 1];
  const months = new Set();
  dates.forEach(d => months.add(`${d.getFullYear()}年${d.getMonth() + 1}月`));
  const sortedMonths = Array.from(months).sort();
  if (sortedMonths.length === 1) return sortedMonths[0];
  return `${sortedMonths[0]}〜${sortedMonths[sortedMonths.length - 1].replace(/^\d+年/, "")}`;
}
const ANALYSIS_PERIOD = getAnalysisPeriod(data.comments);

// ============================================================
// Extract core metrics
// ============================================================
const totalReviews = data.total_reviews || 0;
const overallAvg = data.overall_avg_10pt || 0;
const highCount = data.high_count || 0;
const highRate = data.high_rate || 0;
const midCount = data.mid_count || 0;
const midRate = data.mid_rate || 0;
const lowCount = data.low_count || 0;
const lowRate = data.low_rate || 0;
const siteStats = data.site_stats || [];
const distribution = data.distribution || [];
const comments = data.comments || [];

const siteNames = siteStats.map(s => s.site);
const REVIEW_COUNT = `${totalReviews}件（${siteStats.length}サイト）`;
const TARGET_SITES = "対象サイト：" + siteNames.join(" / ");

// ============================================================
// Color helpers based on score
// ============================================================
function scoreColor(score) {
  if (score >= 9.0) return "27AE60";
  if (score >= 8.0) return "2196F3";
  if (score >= 7.0) return "FF9800";
  return "E74C3C";
}
function scoreBgColor(score) {
  if (score >= 9.0) return "E8F5E9";
  if (score >= 8.0) return "E3F2FD";
  if (score >= 7.0) return "FFF3E0";
  return "FFEBEE";
}
function scoreJudgment(score) {
  if (score >= 9.0) return "優秀";
  if (score >= 8.0) return "良好";
  if (score >= 7.0) return "概ね良好";
  return "要改善";
}

// ============================================================
// Keyword Analysis Engine
// ============================================================
const KEYWORD_CATEGORIES = {
  positive: {
    "立地・アクセス": { keywords: ["立地", "駅", "アクセス", "近い", "便利", "location", "close", "metro", "station", "subway", "地下鉄", "徒歩", "近く"], mentions: 0, samples: [] },
    "清潔さ・清掃": { keywords: ["清潔", "清掃", "綺麗", "きれい", "キレイ", "clean", "掃除", "行き届", "衛生"], mentions: 0, samples: [] },
    "スタッフの対応": { keywords: ["スタッフ", "フロント", "対応", "親切", "丁寧", "笑顔", "フレンドリー", "friendly", "staff", "helpful", "接客", "サービス"], mentions: 0, samples: [] },
    "部屋の広さ": { keywords: ["広い", "広め", "スペース", "spacious", "large", "big", "huge", "大きい"], mentions: 0, samples: [] },
    "コストパフォーマンス": { keywords: ["コスパ", "コストパフォーマンス", "リーズナブル", "お得", "安い", "手頃", "value", "reasonable", "affordable", "价格", "划算"], mentions: 0, samples: [] },
    "朝食": { keywords: ["朝食", "朝ごはん", "breakfast", "ビュッフェ", "美味し"], mentions: 0, samples: [] },
    "快適さ・設備": { keywords: ["快適", "comfortable", "ベッド", "bed", "アメニティ", "amenities", "Wi-Fi", "wifi", "設備"], mentions: 0, samples: [] },
    "コンビニ・周辺施設": { keywords: ["コンビニ", "ファミリーマート", "ファミマ", "セブン", "ローソン", "ドンキ", "convenience", "スーパー", "買い物", "レストラン", "飲食"], mentions: 0, samples: [] },
  },
  negative: {
    "設備の老朽化": { keywords: ["古い", "古さ", "老朽", "old", "年季", "年数", "オンボロ", "古め", "经年"], mentions: 0, samples: [] },
    "防音性能": { keywords: ["防音", "騒音", "音", "うるさい", "noise", "noisy", "sound", "壁が薄", "音漏れ", "소음", "진동"], mentions: 0, samples: [] },
    "部屋の狭さ": { keywords: ["狭い", "狭め", "小さい", "small", "tiny", "narrow", "窮屈", "좁", "cramped"], mentions: 0, samples: [] },
    "浴室・水回り": { keywords: ["浴室", "バス", "シャワー", "排水", "水回り", "ユニットバス", "bathroom", "shower", "drain", "toilet", "トイレ", "화장실"], mentions: 0, samples: [] },
    "臭い・匂い": { keywords: ["臭い", "匂い", "臭", "smell", "odor", "냄새", "タバコ", "煙草", "smoke"], mentions: 0, samples: [] },
    "朝食の改善要望": { keywords: ["朝食.*同じ", "メニュー.*少", "品数", "品切", "食べれるもの.*限", "breakfast.*same", "エスニック", "タイ料理.*(?:辛|不満)"], mentions: 0, samples: [] },
    "エアコン・空調": { keywords: ["エアコン", "空調", "暑い", "寒い", "温度", "AC", "air con", "冷房", "暖房", "蒸し暑", "stuffy", "hot"], mentions: 0, samples: [] },
    "窓・採光": { keywords: ["窓", "開けられない", "frosted", "曇りガラス", "自然光", "暗い", "dreary", "창밖", "明るく"], mentions: 0, samples: [] },
    "対応・サービスの不満": { keywords: ["態度", "不愉快", "返金", "refund", "残念", "謝罪", "rude"], mentions: 0, samples: [] },
    "清掃不備": { keywords: ["ゴミ", "不潔", "汚", "dirty", "unclean", "掃除.*不"], mentions: 0, samples: [] },
  }
};

function getAllCommentText(comment) {
  const fields = [
    comment.comment, comment.good, comment.bad,
    comment.translated, comment.translated_good, comment.translated_bad
  ];
  return fields.filter(Boolean).join(" ");
}

function analyzeKeywords() {
  comments.forEach(c => {
    const text = getAllCommentText(c);
    if (!text.trim()) return;

    for (const [category, catData] of Object.entries(KEYWORD_CATEGORIES.positive)) {
      for (const kw of catData.keywords) {
        try {
          if (new RegExp(kw, "i").test(text)) {
            catData.mentions++;
            if (catData.samples.length < 3) {
              const snippet = extractSnippet(text, kw);
              if (snippet && !catData.samples.includes(snippet)) catData.samples.push(snippet);
            }
            break;
          }
        } catch(e) {
          if (text.toLowerCase().includes(kw.toLowerCase())) {
            catData.mentions++;
            break;
          }
        }
      }
    }

    for (const [category, catData] of Object.entries(KEYWORD_CATEGORIES.negative)) {
      for (const kw of catData.keywords) {
        try {
          if (new RegExp(kw, "i").test(text)) {
            catData.mentions++;
            if (catData.samples.length < 3) {
              const snippet = extractSnippet(text, kw);
              if (snippet && !catData.samples.includes(snippet)) catData.samples.push(snippet);
            }
            break;
          }
        } catch(e) {
          if (text.toLowerCase().includes(kw.toLowerCase())) {
            catData.mentions++;
            break;
          }
        }
      }
    }
  });
}

function extractSnippet(text, keyword) {
  try {
    const regex = new RegExp(`[^。！？.!?]*${keyword}[^。！？.!?]*`, "i");
    const match = text.match(regex);
    if (match) {
      let snippet = match[0].trim();
      if (snippet.length > 60) snippet = snippet.substring(0, 57) + "...";
      return snippet;
    }
  } catch(e) {}
  return null;
}

analyzeKeywords();

// Sort strengths and weaknesses by mention count
const strengthsSorted = Object.entries(KEYWORD_CATEGORIES.positive)
  .filter(([_, d]) => d.mentions > 0)
  .sort((a, b) => b[1].mentions - a[1].mentions);

const weaknessesSorted = Object.entries(KEYWORD_CATEGORIES.negative)
  .filter(([_, d]) => d.mentions > 0)
  .sort((a, b) => b[1].mentions - a[1].mentions);

// ============================================================
// Generate content from data
// ============================================================

// Site data formatted for tables
const SITE_DATA = siteStats.map(s => {
  const scaleText = s.scale === "/5" ? "/5 (×2)" : s.scale;
  return [s.site, String(s.count), String(s.native_avg), scaleText, String(s.avg_10pt), String(s.median_10pt), s.judgment || scoreJudgment(s.avg_10pt)];
});

// Distribution data
const DISTRIBUTION_DATA = distribution.map(d => [d.score, d.count, d.pct]);
const DISTRIBUTION_MAX_COUNT = Math.max(...distribution.map(d => d.count), 1);

// KPI Cards
const KPI_CARDS = [
  { label: "全体平均(10pt換算)", value: overallAvg.toFixed(2), color: scoreColor(overallAvg), bgColor: scoreBgColor(overallAvg) },
  { label: "高評価率(8-10点)", value: highRate.toFixed(1) + "%", color: highRate >= 80 ? "27AE60" : (highRate >= 60 ? "FF9800" : "E74C3C"), bgColor: highRate >= 80 ? "E8F5E9" : (highRate >= 60 ? "FFF3E0" : "FFEBEE") },
  { label: "低評価率(1-4点)", value: lowRate.toFixed(1) + "%", color: lowRate <= 5 ? "27AE60" : (lowRate <= 15 ? "FF9800" : "E74C3C"), bgColor: lowRate <= 5 ? "E8F5E9" : (lowRate <= 15 ? "FFF3E0" : "FFEBEE") },
  { label: "レビュー総数", value: totalReviews + "件", color: "1B3A5C", bgColor: "D5E8F0" },
];

// Executive Summary Text Generation
function generateExecutiveSummaryIntro() {
  return `${ANALYSIS_PERIOD}に各OTAサイト・口コミサイトに投稿された${totalReviews}件のレビューを包括的に分析しました。以下が主要な発見事項です。`;
}

function generateEvaluation1() {
  const bestSite = siteStats.length > 0 ? siteStats.reduce((a, b) => a.avg_10pt > b.avg_10pt ? a : b) : null;
  const judgment = scoreJudgment(overallAvg);
  let text = `全体平均スコア${overallAvg.toFixed(2)}点（10点換算）は`;
  if (overallAvg >= 9.0) text += "非常に高い水準であり、";
  else if (overallAvg >= 8.0) text += "良好な水準であり、";
  else if (overallAvg >= 7.0) text += "概ね標準的な水準であり、";
  else text += "改善の余地がある水準であり、";
  text += `${highRate.toFixed(1)}%が高評価（8-10点）に分類されます。`;
  text += `低評価（1-4点）は${lowCount}件（${lowRate.toFixed(1)}%）`;
  if (lowRate <= 5) text += "で、ゲスト満足度は良好です。";
  else if (lowRate <= 15) text += "で、一定の改善課題があります。";
  else text += "で、早急な改善対応が必要です。";
  if (bestSite) {
    text += `${bestSite.site}（${bestSite.avg_10pt.toFixed(2)}点）で特に高い評価を獲得しています。`;
  }
  return text;
}

function generateEvaluation2() {
  const worstSite = siteStats.length > 0 ? siteStats.reduce((a, b) => a.avg_10pt < b.avg_10pt ? a : b) : null;
  let text = "";
  if (worstSite && worstSite.avg_10pt < overallAvg - 1.0) {
    text += `一方で、${worstSite.site}（${worstSite.avg_10pt.toFixed(1)}点）では相対的に低いスコアとなっており、改善の余地があります。`;
  }
  const topStrengths = strengthsSorted.slice(0, 2).map(([name]) => name);
  if (topStrengths.length > 0) {
    text += `全体としては「${topStrengths.join("」と「")}」がホテルの主な強みとして際立っています。`;
  }
  return text;
}

// Key findings
function generateKeyFindingStrength() {
  return strengthsSorted.slice(0, 3).map(([name, d]) => {
    const countLabel = d.mentions >= 50 ? "50件以上" : d.mentions >= 20 ? "20件以上" : d.mentions >= 10 ? "10件以上" : `${d.mentions}件`;
    return `${name}（${countLabel}で言及）`;
  }).join("、") || "十分なデータなし";
}

function generateKeyFindingWeakness() {
  return weaknessesSorted.slice(0, 4).map(([name, d]) => `${name}（${d.mentions}件）`).join("、") || "特筆すべき弱みなし";
}

function generateKeyFindingOpportunity() {
  if (overallAvg >= 8.5) return "全体的に高評価を維持しており、設備面の小改善で高評価率をさらに安定させるポテンシャルが大きい";
  if (overallAvg >= 7.0) return "基本的なサービス品質は確保されており、弱み項目の改善により高評価率の向上が期待できる";
  return "複数の改善課題が存在するが、重点的な対策により顧客満足度の大幅な向上が可能";
}

// Strength themes for table
function generateStrengthThemes() {
  return strengthsSorted.slice(0, 6).map(([name, d]) => {
    const countLabel = d.mentions >= 50 ? "50件以上" : d.mentions >= 20 ? "20件以上" : d.mentions >= 10 ? "10件以上" : `${d.mentions}件`;
    const sampleText = d.samples.length > 0 ? d.samples.slice(0, 2).map(s => `「${s}」`).join("") : "（コメントなし）";
    return [name, countLabel, sampleText];
  });
}

// Strength sub-sections
function generateStrengthSub1() {
  const top = strengthsSorted[0];
  if (!top) return { title: "3.1 主要な強み", text: "口コミから明確な強みテーマが抽出されました。", bullets: [] };
  const [name, d] = top;
  return {
    title: `3.1 最大の強み：「${name}」`,
    text: `口コミの多数で「${name}」に関するポジティブな言及が見られ、${HOTEL_NAME}の最大の差別化要因となっています。${d.samples.length > 0 ? d.samples[0] + "といった声が多数寄せられています。" : ""}`,
    bullets: d.samples.slice(0, 4).map(s => s),
  };
}

function generateStrengthSub2() {
  const second = strengthsSorted[1];
  if (!second) return { title: "3.2 その他の強み", text: "追加の強みテーマは限定的です。" };
  const [name, d] = second;
  return {
    title: `3.2 ${name}`,
    text: `${name}に関するポジティブな評価も多く寄せられています。${d.samples.length > 0 ? "「" + d.samples[0] + "」といったコメントが見られます。" : ""}`,
  };
}

// Weakness priority data
function generateWeaknessPriorityData() {
  const priorityLabels = ["S", "A", "A", "B", "B", "C", "C", "C", "C"];
  return weaknessesSorted.slice(0, 6).map(([name, d], i) => {
    const priority = i === 0 ? "S" : i <= 2 ? "A" : i <= 4 ? "B" : "C";
    const impact = d.mentions >= 5 ? `重大・${d.mentions}件` : d.mentions >= 3 ? `中度・${d.mentions}件` : `軽度・${d.mentions}件`;
    const detail = d.samples.length > 0 ? d.samples[0] : `${name}に関する指摘あり`;
    return [priority, name, detail, impact];
  });
}

// Improvement phases auto-generation based on weaknesses
function generatePhase1Items() {
  const items = [];
  const weakTopics = weaknessesSorted.map(([name]) => name);

  if (weakTopics.some(t => t.includes("臭") || t.includes("匂"))) {
    items.push({
      title: "(1) 客室の消臭対策の強化",
      bullets: [
        "チェックアウト後の換気・消臭スプレー処理の徹底",
        "定期的な脱臭剤の設置と交換フローの確立",
        "喫煙・臭いに関するクレーム発生時の即座対応マニュアル作成",
      ],
    });
  }
  if (weakTopics.some(t => t.includes("清掃"))) {
    items.push({
      title: `(${items.length + 1}) 清掃品質の向上`,
      bullets: [
        "清掃チェックリストの見直しと強化（ベッド下・隅の重点確認）",
        "清掃後のダブルチェック体制の導入",
        "清掃スタッフへの再研修実施",
      ],
    });
  }
  if (weakTopics.some(t => t.includes("朝食"))) {
    items.push({
      title: `(${items.length + 1}) 朝食サービスの改善`,
      bullets: [
        "メニューの定期的な見直しとバリエーション追加",
        "補充頻度の向上と品切れ防止",
        "宿泊客数に応じた朝食準備量の最適化",
      ],
    });
  }
  if (weakTopics.some(t => t.includes("エアコン") || t.includes("空調"))) {
    items.push({
      title: `(${items.length + 1}) エアコン・空調の定期点検強化`,
      bullets: [
        "全室のエアコン設定を季節に応じてデフォルト確認する運用フロー導入",
        "冷暖房の正常動作を定期チェックリストに追加",
        "異常報告があった客室の優先メンテナンス実施",
      ],
    });
  }
  if (weakTopics.some(t => t.includes("対応") || t.includes("サービス"))) {
    items.push({
      title: `(${items.length + 1}) スタッフ対応品質の統一`,
      bullets: [
        "接客マニュアルの見直しと全スタッフへの周知",
        "ゲストからのクレーム対応フローの明確化",
        "定期的なCS研修の実施",
      ],
    });
  }

  // Fallback items if not enough
  if (items.length < 2) {
    items.push({
      title: `(${items.length + 1}) 口コミモニタリング体制の構築`,
      bullets: [
        "全OTAサイトの口コミを週次で確認するフロー導入",
        "ネガティブコメントへの48時間以内の返信を目標設定",
        "スタッフ間での口コミ共有ミーティングの定例化",
      ],
    });
  }
  if (items.length < 3) {
    items.push({
      title: `(${items.length + 1}) ゲストコミュニケーションの強化`,
      bullets: [
        "チェックイン時の挨拶・案内の標準化",
        "チェックアウト時の簡易アンケート実施",
        "リピーターへの感謝メッセージの導入",
      ],
    });
  }

  return items.slice(0, 4);
}

function generatePhase2Items() {
  const items = [];
  const weakTopics = weaknessesSorted.map(([name]) => name);

  if (weakTopics.some(t => t.includes("浴室") || t.includes("水回"))) {
    items.push({
      title: "(1) 浴室の快適性向上",
      bullets: [
        "排水口の清掃・修繕を全室で実施し、排水速度を改善",
        "シャワーヘッドの交換（水圧改善・節水タイプ導入）",
        "浴室換気・暖房設備の点検と改善",
      ],
    });
  }
  if (weakTopics.some(t => t.includes("防音"))) {
    items.push({
      title: `(${items.length + 1}) 防音対策の強化`,
      bullets: [
        "窓際への防音カーテン導入（外部騒音対策）",
        "ドア下部への隙間テープ設置（廊下・隣室からの音漏れ軽減）",
        "騒音の少ない客室への優先案内の仕組み化",
      ],
    });
  }
  if (weakTopics.some(t => t.includes("窓") || t.includes("採光"))) {
    items.push({
      title: `(${items.length + 1}) 窓・採光環境の改善`,
      bullets: [
        "遮光カーテンとレースカーテンの品質見直し",
        "窓の開閉可否に関する事前案内の徹底",
        "採光の良い客室の優先割り当てルール整備",
      ],
    });
  }

  items.push({
    title: `(${items.length + 1}) 口コミ返信体制の構築`,
    bullets: [
      "全サイトの口コミに48時間以内の返信を目標に設定",
      "特に低評価レビューへのフォローアップ強化",
    ],
  });

  return items.slice(0, 3);
}

function generatePhase3Items() {
  const items = [];
  const weakTopics = weaknessesSorted.map(([name]) => name);

  if (weakTopics.some(t => t.includes("老朽") || t.includes("設備"))) {
    items.push({
      title: "(1) 客室リニューアル計画",
      bullets: [
        "壁紙・カーペットの張替えによる「古さ」印象の払拭",
        "照明のLED化・調光機能の追加で客室の雰囲気向上",
        "USB充電ポート付きコンセントの設置",
      ],
    });
  } else {
    items.push({
      title: "(1) 客室設備のアップグレード",
      bullets: [
        "内装の定期的なリフレッシュ計画の策定",
        "照明・電源周りの設備更新",
        "快適性向上のための設備投資計画の立案",
      ],
    });
  }

  if (weakTopics.some(t => t.includes("浴室") || t.includes("水回"))) {
    items.push({
      title: "(2) 浴室の抜本的改修",
      bullets: [
        "配管更新による排水問題の根本解決",
        "換気設備の強化による浴室内の温度・湿度改善",
      ],
    });
  }

  if (weakTopics.some(t => t.includes("防音"))) {
    items.push({
      title: `(${items.length + 1}) 防音工事`,
      bullets: [
        "騒音源側客室の窓を二重サッシに交換",
        "隣室との壁面に遮音材を追加する改修工事の検討",
      ],
    });
  }

  if (weakTopics.some(t => t.includes("部屋") && t.includes("狭"))) {
    items.push({
      title: `(${items.length + 1}) 客室レイアウトの見直し`,
      bullets: [
        "家具配置の最適化による体感スペースの拡大",
        "コンパクト家具への入替検討",
        "収納スペースの効率化",
      ],
    });
  }

  if (items.length < 2) {
    items.push({
      title: `(${items.length + 1}) ブランド価値向上施策`,
      bullets: [
        "ホテル独自のサービスコンセプトの策定",
        "差別化ポイントを活かしたプロモーション強化",
        "リピーター向け特典プログラムの導入",
      ],
    });
  }

  return items.slice(0, 3);
}

// KPI targets generation
function generateKPITargets() {
  const targets = [];
  const targetAvg = overallAvg < 8.0 ? (overallAvg + 0.5).toFixed(1) + "点以上" : overallAvg < 9.0 ? (overallAvg + 0.3).toFixed(1) + "点以上" : (Math.min(overallAvg + 0.2, 9.8)).toFixed(1) + "点以上";
  targets.push(["全体平均(10pt換算)", overallAvg.toFixed(2) + "点", targetAvg, "2026年9月"]);

  const targetHighRate = highRate >= 90 ? (Math.min(highRate + 2, 98)).toFixed(0) + "%以上" : highRate >= 70 ? (highRate + 5).toFixed(0) + "%以上" : (highRate + 10).toFixed(0) + "%以上";
  targets.push(["高評価率（8-10点）", highRate.toFixed(1) + "%", targetHighRate, "2026年9月"]);

  const targetLowRate = lowRate <= 2 ? "0%維持" : lowRate <= 10 ? (Math.max(lowRate - 3, 0)).toFixed(0) + "%以下" : (Math.max(lowRate - 5, 2)).toFixed(0) + "%以下";
  targets.push(["低評価率（1-4点）", lowRate.toFixed(1) + "%", targetLowRate, "2026年9月"]);

  // Worst site improvement target
  const worstSite = siteStats.length > 0 ? siteStats.reduce((a, b) => a.avg_10pt < b.avg_10pt ? a : b) : null;
  if (worstSite && worstSite.avg_10pt < overallAvg - 0.5) {
    const targetScore = (worstSite.avg_10pt + 1.0).toFixed(1) + "点以上";
    targets.push([`${worstSite.site}平均評価`, worstSite.avg_10pt.toFixed(1) + "点", targetScore, "2026年9月"]);
  }

  targets.push(["口コミ返信率", "未計測", "100%（48h以内）", "2026年6月"]);

  if (weaknessesSorted.length > 0) {
    const topWeakCount = weaknessesSorted[0][1].mentions;
    targets.push([weaknessesSorted[0][0] + "クレーム", `約${topWeakCount}件/期間`, `${Math.max(Math.floor(topWeakCount / 2), 1)}件以下`, "2026年9月"]);
  }

  return targets;
}

// Conclusion text generation
function generateConclusionParagraphs() {
  const paragraphs = [];
  const topStrengths = strengthsSorted.slice(0, 2).map(([name]) => name);
  const topWeaknesses = weaknessesSorted.slice(0, 3).map(([name]) => name);

  paragraphs.push(
    `${HOTEL_NAME}は、全体平均${overallAvg.toFixed(2)}点（10点換算）、高評価率${highRate.toFixed(1)}%という${overallAvg >= 9.0 ? "非常に高い" : overallAvg >= 8.0 ? "良好な" : "一定の"}顧客満足度を${overallAvg >= 8.0 ? "維持している" : "示している"}ホテルです。${topStrengths.length > 0 ? "特に「" + topStrengths.join("」「") + "」が多くのゲストから高く評価されています。" : ""}`
  );

  if (topWeaknesses.length > 0) {
    paragraphs.push(
      `改善課題としては、${topWeaknesses.join("、")}が主要なテーマとして浮上しています。${overallAvg >= 8.0 ? "ただし、これらの課題を指摘するゲストの多くも他の面では高い評価を付けており、ネガティブ要素が全体評価を大きく押し下げているわけではありません。" : "これらの課題への対応が、顧客満足度向上の鍵となります。"}`
    );
  }

  paragraphs.push(
    `Phase 1（即座対応）のオペレーション改善は、低コストで即座に実行可能であり、ゲスト体験の底上げに直結します。Phase 2（短期）の設備改善は、中評価ゲストを高評価に引き上げる効果が期待できます。Phase 3（中期）の抜本的改修は、ブランド価値の向上に寄与します。`
  );

  const worstSite = siteStats.length > 0 ? siteStats.reduce((a, b) => a.avg_10pt < b.avg_10pt ? a : b) : null;
  if (worstSite && worstSite.avg_10pt < overallAvg - 0.5) {
    paragraphs.push(
      `${worstSite.site}での低スコア（${worstSite.avg_10pt.toFixed(1)}点）は、口コミ返信の強化とサービス品質の一貫性確保で改善余地があります。全サイトでの口コミ返信率100%を目指すことで、ホテルの姿勢をゲストに伝え、リピート率の向上につなげることが重要です。`
    );
  }

  return paragraphs;
}

function generateConclusionFinal() {
  const topStrength = strengthsSorted.length > 0 ? strengthsSorted[0][0] : "サービス品質";
  return `${topStrength}という強みを活かしつつ、段階的な改善を実行することで、${HOTEL_NAME}はさらなる顧客満足度の向上とブランド価値の強化を実現することができます。`;
}

// Appendix - sample comments
function generateAppendixComments() {
  const withText = comments.filter(c => {
    const text = getAllCommentText(c);
    return text.trim().length > 10;
  });
  // Return up to 10 representative comments
  const highRated = withText.filter(c => c.rating_10pt >= 8).slice(0, 5);
  const lowRated = withText.filter(c => c.rating_10pt <= 6).slice(0, 5);
  return [...highRated, ...lowRated].slice(0, 10);
}

// ============================================================
// Generate all content
// ============================================================
const EXEC_SUMMARY_INTRO = generateExecutiveSummaryIntro();
const EXEC_EVAL_1 = generateEvaluation1();
const EXEC_EVAL_2 = generateEvaluation2();
const KEY_FINDING_STRENGTH = generateKeyFindingStrength();
const KEY_FINDING_WEAKNESS = generateKeyFindingWeakness();
const KEY_FINDING_OPPORTUNITY = generateKeyFindingOpportunity();
const STRENGTH_THEMES = generateStrengthThemes();
const STRENGTH_SUB_1 = generateStrengthSub1();
const STRENGTH_SUB_2 = generateStrengthSub2();
const WEAKNESS_PRIORITY_DATA = generateWeaknessPriorityData();
const PHASE1_ITEMS = generatePhase1Items();
const PHASE2_ITEMS = generatePhase2Items();
const PHASE3_ITEMS = generatePhase3Items();
const KPI_TARGET_DATA = generateKPITargets();
const CONCLUSION_PARAGRAPHS = generateConclusionParagraphs();
const CONCLUSION_FINAL = generateConclusionFinal();
const APPENDIX_COMMENTS = generateAppendixComments();

const HIGH_RATING_SUMMARY = `${highCount}件（${highRate.toFixed(1)}%）`;
const MID_RATING_SUMMARY = `${midCount}件（${midRate.toFixed(1)}%）`;
const LOW_RATING_SUMMARY = `${lowCount}件（${lowRate.toFixed(1)}%）`;
const LOW_RATING_COLOR = lowRate <= 5 ? "27AE60" : lowRate <= 15 ? "FF9800" : "E74C3C";

const DATA_OVERVIEW_SITE_TEXT = "各予約サイトの評価基準が異なる点に留意が必要です。海外OTA（Booking.com/Trip.com/Agoda）は10点満点、国内サイト（じゃらん/楽天トラベル）およびGoogleは5点満点のため、統一比較のため10点換算（5点満点×2）を行いました。";

function generateDistOverviewText() {
  const topScore = distribution.length > 0 ? distribution.reduce((a, b) => a.count > b.count ? a : b) : null;
  let text = "";
  if (topScore) {
    text += `評価分布は${topScore.score}点（${topScore.pct}）に最も多く、`;
  }
  text += `8点以上の高評価が${highRate.toFixed(1)}%を占め`;
  if (highRate >= 80) text += "、高い顧客満足度を示しています。";
  else if (highRate >= 60) text += "ています。";
  else text += "る一方、改善の余地も見られます。";
  if (lowCount > 0) {
    text += `低評価（1-4点）は計${lowCount}件（${lowRate.toFixed(1)}%）です。`;
  }
  return text;
}
const DATA_OVERVIEW_DIST_TEXT = generateDistOverviewText();


// ============================================================
// DOCX COLORS (same as reference)
// ============================================================
const NAVY = "1B3A5C";
const ACCENT = "2E75B6";
const LIGHT_BLUE = "D5E8F0";
const LIGHT_GRAY = "F2F2F2";
const WHITE = "FFFFFF";
const RED_ACCENT = "C0392B";
const GREEN_ACCENT = "27AE60";
const ORANGE_ACCENT = "E67E22";

const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };
const noBorders = {
  top: { style: BorderStyle.NONE, size: 0 },
  bottom: { style: BorderStyle.NONE, size: 0 },
  left: { style: BorderStyle.NONE, size: 0 },
  right: { style: BorderStyle.NONE, size: 0 },
};
const cellMargins = { top: 80, bottom: 80, left: 120, right: 120 };

// ============================================================
// DOCX Helper functions (identical to reference)
// ============================================================
function heading1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 200 },
    children: [new TextRun({ text, bold: true, size: 32, font: "Arial", color: NAVY })],
  });
}

function heading2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 280, after: 160 },
    children: [new TextRun({ text, bold: true, size: 26, font: "Arial", color: ACCENT })],
  });
}

function heading3(text) {
  return new Paragraph({
    spacing: { before: 200, after: 120 },
    children: [new TextRun({ text, bold: true, size: 22, font: "Arial", color: NAVY })],
  });
}

function para(text, opts = {}) {
  return new Paragraph({
    spacing: { after: opts.afterSpacing || 120, line: 320 },
    alignment: opts.alignment || AlignmentType.LEFT,
    children: [new TextRun({ text, size: 21, font: "Arial", color: opts.color || "333333", ...opts })],
  });
}

function multiRunPara(runs, opts = {}) {
  return new Paragraph({
    spacing: { after: opts.afterSpacing || 120, line: 320 },
    alignment: opts.alignment || AlignmentType.LEFT,
    children: runs.map(r => new TextRun({ size: 21, font: "Arial", color: "333333", ...r })),
  });
}

function bulletItem(text, opts = {}) {
  return new Paragraph({
    numbering: { reference: "bullets", level: opts.level || 0 },
    spacing: { after: 80, line: 300 },
    children: [new TextRun({ text, size: 21, font: "Arial", color: "333333" })],
  });
}

function headerCell(text, width) {
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    shading: { fill: NAVY, type: ShadingType.CLEAR },
    margins: cellMargins,
    verticalAlign: "center",
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text, bold: true, size: 20, font: "Arial", color: WHITE })],
    })],
  });
}

function dataCell(text, width, opts = {}) {
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    shading: opts.fill ? { fill: opts.fill, type: ShadingType.CLEAR } : undefined,
    margins: cellMargins,
    verticalAlign: "center",
    children: [new Paragraph({
      alignment: opts.alignment || AlignmentType.LEFT,
      children: [new TextRun({ text: String(text), size: 20, font: "Arial", color: opts.color || "333333", bold: opts.bold || false })],
    })],
  });
}

function spacer(height = 100) {
  return new Paragraph({ spacing: { after: height }, children: [] });
}

function divider() {
  return new Paragraph({
    spacing: { before: 200, after: 200 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: ACCENT, space: 1 } },
    children: [],
  });
}

function kpiRow(items) {
  const colWidth = Math.floor(9360 / items.length);
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: items.map(() => colWidth),
    rows: [
      new TableRow({
        children: items.map(item =>
          new TableCell({
            borders: { ...noBorders, right: { style: BorderStyle.SINGLE, size: 1, color: "DDDDDD" } },
            width: { size: colWidth, type: WidthType.DXA },
            shading: { fill: item.bgColor || LIGHT_BLUE, type: ShadingType.CLEAR },
            margins: { top: 160, bottom: 160, left: 200, right: 200 },
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 40 },
                children: [new TextRun({ text: item.label, size: 18, font: "Arial", color: "666666" })],
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [new TextRun({ text: item.value, bold: true, size: 36, font: "Arial", color: item.color || NAVY })],
              }),
            ],
          })
        ),
      }),
    ],
  });
}

function priorityRow(priority, category, detail, impact) {
  const colors = { "S": RED_ACCENT, "A": ORANGE_ACCENT, "B": ACCENT, "C": "888888" };
  return new TableRow({
    children: [
      new TableCell({
        borders,
        width: { size: 800, type: WidthType.DXA },
        margins: cellMargins,
        shading: { fill: LIGHT_GRAY, type: ShadingType.CLEAR },
        verticalAlign: "center",
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: priority, bold: true, size: 22, font: "Arial", color: colors[priority] || NAVY })],
        })],
      }),
      dataCell(category, 2200, { bold: true }),
      dataCell(detail, 4360),
      dataCell(impact, 2000, { alignment: AlignmentType.CENTER }),
    ],
  });
}

function buildPhaseItems(items) {
  const result = [];
  items.forEach((item, index) => {
    result.push(heading3(item.title));
    item.bullets.forEach(b => result.push(bulletItem(b)));
    if (index < items.length - 1) result.push(spacer(80));
  });
  return result;
}

function buildSiteTableRows(siteData) {
  return siteData.map((row, index) => {
    const [siteName, count, nativeAvg, scale, tenPt, median, verdict] = row;
    const fill = index % 2 === 1 ? LIGHT_GRAY : undefined;
    const verdictColor = verdict === "優秀" ? GREEN_ACCENT : (verdict === "良好" ? GREEN_ACCENT : (verdict === "概ね良好" ? ORANGE_ACCENT : RED_ACCENT));
    return new TableRow({ children: [
      dataCell(siteName, 1800, { bold: true, fill }),
      dataCell(count, 900, { alignment: AlignmentType.CENTER, fill }),
      dataCell(nativeAvg, 1300, { alignment: AlignmentType.CENTER, bold: true, fill }),
      dataCell(scale, 900, { alignment: AlignmentType.CENTER, fill }),
      dataCell(tenPt, 1300, { alignment: AlignmentType.CENTER, color: verdictColor, bold: true, fill }),
      dataCell(median, 1526, { alignment: AlignmentType.CENTER, fill }),
      dataCell(verdict, 1300, { alignment: AlignmentType.CENTER, color: verdictColor, bold: true, fill }),
    ]});
  });
}

function buildDistributionRows(distData, maxCount) {
  return distData.map(([rating, count, pct]) => {
    const barWidth = Math.round(count / maxCount * 100);
    const barColor = rating >= 8 ? GREEN_ACCENT : (rating >= 5 ? ORANGE_ACCENT : RED_ACCENT);
    const fill = [10, 8, 6, 4, 2].includes(rating) ? LIGHT_GRAY : WHITE;
    return new TableRow({ children: [
      dataCell(String(rating), 1200, { alignment: AlignmentType.CENTER, bold: true, fill }),
      dataCell(String(count), 1200, { alignment: AlignmentType.CENTER, fill }),
      dataCell(pct, 1500, { alignment: AlignmentType.CENTER, fill }),
      new TableCell({
        borders, width: { size: 5126, type: WidthType.DXA }, margins: cellMargins,
        shading: fill !== WHITE ? { fill, type: ShadingType.CLEAR } : undefined,
        children: [new Paragraph({ children: [
          new TextRun({ text: "\u2588".repeat(Math.max(1, Math.round(barWidth / 5))), size: 20, font: "Arial", color: barColor }),
          new TextRun({ text: ` ${count}件`, size: 18, font: "Arial", color: "666666" }),
        ]})],
      }),
    ]});
  });
}

function buildStrengthRows(themeData) {
  return themeData.map((row, index) => {
    const [theme, mentionCount, comments] = row;
    const fill = index % 2 === 0 ? "E8F5E9" : undefined;
    return new TableRow({ children: [
      dataCell(theme, 2600, { bold: true, fill: fill || undefined }),
      dataCell(mentionCount, 1200, { alignment: AlignmentType.CENTER, bold: true, color: GREEN_ACCENT, fill: fill || undefined }),
      dataCell(comments, 5226, { fill: fill || undefined }),
    ]});
  });
}

function buildKPITargetRows(kpiData) {
  return kpiData.map((row, index) => {
    const [item, current, target, deadline] = row;
    const fill = index % 2 === 1 ? LIGHT_GRAY : undefined;
    return new TableRow({ children: [
      dataCell(item, 2800, { bold: true, fill }),
      dataCell(current, 2200, { alignment: AlignmentType.CENTER, fill }),
      dataCell(target, 2200, { alignment: AlignmentType.CENTER, color: GREEN_ACCENT, bold: true, fill }),
      dataCell(deadline, 1826, { alignment: AlignmentType.CENTER, fill }),
    ]});
  });
}

// ============================================================
// Build appendix section
// ============================================================
function buildAppendixSection() {
  const result = [];
  result.push(heading1("7. 付録（口コミ原文抜粋）"));
  result.push(para("以下に、分析に使用した代表的な口コミの原文を抜粋します。"));
  result.push(spacer(100));

  APPENDIX_COMMENTS.forEach((c, i) => {
    const text = c.translated || c.comment || c.translated_good || c.good || "";
    const badText = c.translated_bad || c.bad || "";
    if (!text && !badText) return;

    const ratingLabel = `${c.rating_10pt}点`;
    const siteLabel = c.site;
    const dateLabel = c.date || "";

    result.push(heading3(`[${i + 1}] ${siteLabel} / ${ratingLabel} / ${dateLabel}`));
    if (text) {
      const displayText = text.length > 200 ? text.substring(0, 197) + "..." : text;
      result.push(para(displayText));
    }
    if (badText) {
      result.push(multiRunPara([
        { text: "【改善点】", bold: true, color: RED_ACCENT },
        { text: badText.length > 150 ? badText.substring(0, 147) + "..." : badText },
      ]));
    }
    result.push(spacer(60));
  });

  return result;
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
        run: { size: 32, bold: true, font: "Arial", color: NAVY },
        paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 0 },
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Arial", color: ACCENT },
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
      {
        reference: "numbers",
        levels: [
          { level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
        ],
      },
    ],
  },
  sections: [
    // ===== COVER PAGE =====
    {
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
          border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: ACCENT, space: 8 } },
          children: [new TextRun({ text: "口コミ分析", size: 56, font: "Arial", color: ACCENT })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 200 },
          children: [new TextRun({ text: "改善レポート", size: 56, font: "Arial", bold: true, color: NAVY })],
        }),
        spacer(400),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 120 },
          children: [new TextRun({ text: HOTEL_NAME, size: 32, font: "Arial", color: NAVY, bold: true })],
        }),
        spacer(200),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 80 },
          children: [new TextRun({ text: `分析対象期間：${ANALYSIS_PERIOD}`, size: 22, font: "Arial", color: "666666" })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 80 },
          children: [new TextRun({ text: `レビュー総数：${REVIEW_COUNT}`, size: 22, font: "Arial", color: "666666" })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 80 },
          children: [new TextRun({ text: TARGET_SITES, size: 20, font: "Arial", color: "666666" })],
        }),
        spacer(1600),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 60 },
          children: [new TextRun({ text: `作成日：${REPORT_DATE}`, size: 20, font: "Arial", color: "999999" })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "Confidential - For Internal Use Only", size: 18, font: "Arial", color: "AAAAAA", italics: true })],
        }),
      ],
    },

    // ===== MAIN CONTENT =====
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
            border: { bottom: { style: BorderStyle.SINGLE, size: 2, color: ACCENT, space: 4 } },
            children: [new TextRun({ text: `${HOTEL_NAME}｜口コミ分析改善レポート`, size: 16, font: "Arial", color: "999999" })],
          })],
        }),
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: "Page ", size: 16, font: "Arial", color: "999999" }),
              new TextRun({ children: [PageNumber.CURRENT], size: 16, font: "Arial", color: "999999" })],
          })],
        }),
      },
      children: [
        // ===== 1. EXECUTIVE SUMMARY =====
        heading1("1. エグゼクティブサマリー"),
        para(EXEC_SUMMARY_INTRO),
        spacer(100),

        kpiRow(KPI_CARDS),
        spacer(200),

        heading2("総合評価"),
        para(EXEC_EVAL_1),
        para(EXEC_EVAL_2),
        spacer(100),

        // Key findings box
        new Table({
          width: { size: 9026, type: WidthType.DXA },
          columnWidths: [9026],
          rows: [new TableRow({
            children: [new TableCell({
              borders: { top: { style: BorderStyle.SINGLE, size: 4, color: ACCENT }, bottom: border, left: { style: BorderStyle.SINGLE, size: 4, color: ACCENT }, right: border },
              width: { size: 9026, type: WidthType.DXA },
              shading: { fill: "F0F7FC", type: ShadingType.CLEAR },
              margins: { top: 200, bottom: 200, left: 300, right: 300 },
              children: [
                new Paragraph({ spacing: { after: 120 }, children: [new TextRun({ text: "KEY FINDINGS", bold: true, size: 22, font: "Arial", color: ACCENT })] }),
                new Paragraph({ spacing: { after: 80 }, children: [
                  new TextRun({ text: "Strength：", bold: true, size: 20, font: "Arial", color: GREEN_ACCENT }),
                  new TextRun({ text: KEY_FINDING_STRENGTH, size: 20, font: "Arial", color: "333333" }),
                ]}),
                new Paragraph({ spacing: { after: 80 }, children: [
                  new TextRun({ text: "Weakness：", bold: true, size: 20, font: "Arial", color: RED_ACCENT }),
                  new TextRun({ text: KEY_FINDING_WEAKNESS, size: 20, font: "Arial", color: "333333" }),
                ]}),
                new Paragraph({ children: [
                  new TextRun({ text: "Opportunity：", bold: true, size: 20, font: "Arial", color: ORANGE_ACCENT }),
                  new TextRun({ text: KEY_FINDING_OPPORTUNITY, size: 20, font: "Arial", color: "333333" }),
                ]}),
              ],
            })],
          })],
        }),

        new Paragraph({ children: [new PageBreak()] }),

        // ===== 2. DATA OVERVIEW =====
        heading1("2. データ概要"),

        heading2("2.1 サイト別レビュー件数・評価"),
        para(DATA_OVERVIEW_SITE_TEXT),
        spacer(80),

        new Table({
          width: { size: 9026, type: WidthType.DXA },
          columnWidths: [1800, 900, 1300, 900, 1300, 1526, 1300],
          rows: [
            new TableRow({
              children: [
                headerCell("サイト名", 1800),
                headerCell("件数", 900),
                headerCell("ネイティブ平均", 1300),
                headerCell("尺度", 900),
                headerCell("10pt換算", 1300),
                headerCell("中央値(10pt)", 1526),
                headerCell("判定", 1300),
              ],
            }),
            ...buildSiteTableRows(SITE_DATA),
          ],
        }),
        spacer(200),

        heading2("2.2 評価分布（10点換算）"),
        para(DATA_OVERVIEW_DIST_TEXT),
        spacer(80),

        new Table({
          width: { size: 9026, type: WidthType.DXA },
          columnWidths: [1200, 1200, 1500, 5126],
          rows: [
            new TableRow({ children: [headerCell("評価", 1200), headerCell("件数", 1200), headerCell("割合", 1500), headerCell("分布", 5126)] }),
            ...buildDistributionRows(DISTRIBUTION_DATA, DISTRIBUTION_MAX_COUNT),
          ],
        }),
        spacer(100),

        // Rating category summary
        new Table({
          width: { size: 9026, type: WidthType.DXA },
          columnWidths: [3008, 3009, 3009],
          rows: [new TableRow({
            children: [
              new TableCell({
                borders, width: { size: 3008, type: WidthType.DXA },
                shading: { fill: "E8F5E9", type: ShadingType.CLEAR },
                margins: { top: 120, bottom: 120, left: 200, right: 200 },
                children: [
                  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [new TextRun({ text: "高評価（8-10点）", bold: true, size: 20, font: "Arial", color: GREEN_ACCENT })] }),
                  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: HIGH_RATING_SUMMARY, bold: true, size: 24, font: "Arial", color: GREEN_ACCENT })] }),
                ],
              }),
              new TableCell({
                borders, width: { size: 3009, type: WidthType.DXA },
                shading: { fill: "FFF3E0", type: ShadingType.CLEAR },
                margins: { top: 120, bottom: 120, left: 200, right: 200 },
                children: [
                  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [new TextRun({ text: "中評価（5-7点）", bold: true, size: 20, font: "Arial", color: ORANGE_ACCENT })] }),
                  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: MID_RATING_SUMMARY, bold: true, size: 24, font: "Arial", color: ORANGE_ACCENT })] }),
                ],
              }),
              new TableCell({
                borders, width: { size: 3009, type: WidthType.DXA },
                shading: { fill: "FDEDEC", type: ShadingType.CLEAR },
                margins: { top: 120, bottom: 120, left: 200, right: 200 },
                children: [
                  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [new TextRun({ text: "低評価（1-4点）", bold: true, size: 20, font: "Arial", color: LOW_RATING_COLOR })] }),
                  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: LOW_RATING_SUMMARY, bold: true, size: 24, font: "Arial", color: LOW_RATING_COLOR })] }),
                ],
              }),
            ],
          })],
        }),

        new Paragraph({ children: [new PageBreak()] }),

        // ===== 3. STRENGTH ANALYSIS =====
        heading1("3. 強み分析（ポジティブ要因）"),
        para("口コミのテキストマイニングにより、以下のポジティブテーマが特定されました。"),
        spacer(100),

        new Table({
          width: { size: 9026, type: WidthType.DXA },
          columnWidths: [2600, 1200, 5226],
          rows: [
            new TableRow({ children: [headerCell("ポジティブテーマ", 2600), headerCell("言及数", 1200), headerCell("代表的なコメント", 5226)] }),
            ...buildStrengthRows(STRENGTH_THEMES),
          ],
        }),
        spacer(200),

        heading2(STRENGTH_SUB_1.title),
        para(STRENGTH_SUB_1.text),
        ...(STRENGTH_SUB_1.bullets || []).map(b => bulletItem(b)),
        spacer(100),

        heading2(STRENGTH_SUB_2.title),
        para(STRENGTH_SUB_2.text),

        new Paragraph({ children: [new PageBreak()] }),

        // ===== 4. WEAKNESS ANALYSIS =====
        heading1("4. 弱み分析（改善課題）"),
        para("ネガティブコメントの分析から、以下の改善課題が抽出されました。影響度と頻度に基づく優先度をS〜Cで設定しています。"),
        spacer(100),

        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [800, 2200, 4360, 2000],
          rows: [
            new TableRow({ children: [headerCell("優先度", 800), headerCell("課題カテゴリ", 2200), headerCell("具体的内容", 4360), headerCell("影響度", 2000)] }),
            ...WEAKNESS_PRIORITY_DATA.map(([p, c, d, i]) => priorityRow(p, c, d, i)),
          ],
        }),

        new Paragraph({ children: [new PageBreak()] }),

        // ===== 5. IMPROVEMENT PLAN =====
        heading1("5. 改善施策提案"),
        para("分析結果に基づき、以下の改善施策を「即座対応」「短期」「中期」の3フェーズに分けて提案いたします。"),
        spacer(100),

        // Phase 1
        heading2("Phase 1：即座対応（今週〜1ヶ月以内）"),
        para("投資不要・オペレーション改善で対応可能な施策"),
        spacer(80),
        ...buildPhaseItems(PHASE1_ITEMS),

        new Paragraph({ children: [new PageBreak()] }),

        // Phase 2
        heading2("Phase 2：短期施策（1〜3ヶ月）"),
        para("一定の投資を伴うが、比較的早期に実行可能な施策"),
        spacer(80),
        ...buildPhaseItems(PHASE2_ITEMS),

        spacer(200),

        // Phase 3
        heading2("Phase 3：中期施策（3〜6ヶ月）"),
        para("設備投資を伴う抜本的な改善施策"),
        spacer(80),
        ...buildPhaseItems(PHASE3_ITEMS),

        new Paragraph({ children: [new PageBreak()] }),

        // ===== 6. KPI & TARGETS =====
        heading1("6. KPI目標設定"),
        para("以下のKPIを設定し、四半期ごとにモニタリングすることを推奨します。"),
        spacer(100),

        new Table({
          width: { size: 9026, type: WidthType.DXA },
          columnWidths: [2800, 2200, 2200, 1826],
          rows: [
            new TableRow({ children: [headerCell("KPI項目", 2800), headerCell("現状値", 2200), headerCell("目標値（6ヶ月後）", 2200), headerCell("期限", 1826)] }),
            ...buildKPITargetRows(KPI_TARGET_DATA),
          ],
        }),

        new Paragraph({ children: [new PageBreak()] }),

        // ===== 7. APPENDIX & CONCLUSION =====
        ...buildAppendixSection(),

        spacer(300),
        divider(),

        // Conclusion box
        new Table({
          width: { size: 9026, type: WidthType.DXA },
          columnWidths: [9026],
          rows: [new TableRow({
            children: [new TableCell({
              borders: { top: { style: BorderStyle.SINGLE, size: 4, color: NAVY }, bottom: { style: BorderStyle.SINGLE, size: 4, color: NAVY }, left: { style: BorderStyle.SINGLE, size: 4, color: NAVY }, right: { style: BorderStyle.SINGLE, size: 4, color: NAVY } },
              width: { size: 9026, type: WidthType.DXA },
              shading: { fill: "F8F9FA", type: ShadingType.CLEAR },
              margins: { top: 300, bottom: 300, left: 400, right: 400 },
              children: [
                ...CONCLUSION_PARAGRAPHS.map(text =>
                  new Paragraph({ spacing: { after: 160 }, children: [new TextRun({ text, size: 21, font: "Arial", color: "333333" })] })
                ),
                new Paragraph({ children: [new TextRun({ text: CONCLUSION_FINAL, size: 21, font: "Arial", color: NAVY, bold: true })] }),
              ],
            })],
          })],
        }),

        spacer(300),
        divider(),
        spacer(100),
        para("本レポートに関するご質問・ご相談がございましたら、お気軽にお問い合わせください。", { alignment: AlignmentType.CENTER, color: "999999" }),
      ],
    },
  ],
});


// ============================================================
// Build PPTX Presentation
// ============================================================
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.author = "Hotel Consulting";
pres.title = HOTEL_NAME + " 口コミ分析改善レポート";

// === Color Palette (Midnight Executive) ===
const C = {
  navy: "1A2744",
  navyLight: "243556",
  blue: "3B7DD8",
  blueLight: "5A9BE6",
  ice: "E8EFF8",
  white: "FFFFFF",
  offWhite: "F5F7FA",
  gray: "64748B",
  grayLight: "94A3B8",
  grayDark: "334155",
  green: "16A34A",
  greenBg: "DCFCE7",
  red: "DC2626",
  redBg: "FEE2E2",
  orange: "EA580C",
  orangeBg: "FFF7ED",
  gold: "D4A843",
};

const shadow = () => ({ type: "outer", blur: 4, offset: 2, angle: 135, color: "000000", opacity: 0.10 });

function addFooter(slide, pageNum) {
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.navy } });
  slide.addText("Confidential", { x: 0.5, y: 5.25, w: 3, h: 0.375, fontSize: 8, color: C.grayLight, fontFace: "Arial", valign: "middle" });
  slide.addText(String(pageNum), { x: 9, y: 5.25, w: 0.5, h: 0.375, fontSize: 8, color: C.grayLight, fontFace: "Arial", align: "right", valign: "middle" });
}

function addContentHeader(slide, title, subtitle) {
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.85, fill: { color: C.navy } });
  slide.addText(title, { x: 0.6, y: 0.08, w: 8, h: 0.5, fontSize: 22, fontFace: "Arial", color: C.white, bold: true, margin: 0 });
  if (subtitle) {
    slide.addText(subtitle, { x: 0.6, y: 0.5, w: 8, h: 0.3, fontSize: 10, fontFace: "Arial", color: C.blueLight, margin: 0 });
  }
}

function pptxKpiCard(slide, x, y, w, h, label, value, color, bgColor) {
  slide.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: bgColor }, shadow: shadow() });
  slide.addShape(pres.shapes.RECTANGLE, { x, y, w, h: 0.05, fill: { color } });
  slide.addText(label, { x, y: y + 0.15, w, h: 0.3, fontSize: 10, fontFace: "Arial", color: C.gray, align: "center", margin: 0 });
  slide.addText(value, { x, y: y + 0.4, w, h: 0.55, fontSize: 32, fontFace: "Arial", color, bold: true, align: "center", margin: 0 });
}

// PPTX data preparation
const pptxKpiAvg = overallAvg.toFixed(2);
const pptxKpiHighRate = highRate.toFixed(1) + "%";
const pptxKpiLowRate = lowRate.toFixed(1) + "%";
const pptxKpiTotal = totalReviews + "件";

// Summary strengths for PPTX (top 3)
const SUMMARY_STRENGTHS = strengthsSorted.slice(0, 3).map(([name, d]) => ({
  title: name,
  desc: d.samples.length > 0 ? d.samples[0] : `${name}に関する高評価が多数`,
}));
// Pad if fewer than 3
while (SUMMARY_STRENGTHS.length < 3) {
  SUMMARY_STRENGTHS.push({ title: "全体的な満足度", desc: "ゲストから概ね好意的な評価を受けています" });
}

const SUMMARY_WEAKNESSES = weaknessesSorted.slice(0, 3).map(([name, d]) => ({
  title: name,
  desc: d.samples.length > 0 ? d.samples[0] : `${name}に関する改善要望あり`,
}));
while (SUMMARY_WEAKNESSES.length < 3) {
  SUMMARY_WEAKNESSES.push({ title: "軽微な改善点", desc: "大きな問題はないが、さらなる品質向上の余地あり" });
}

// Site chart data for PPTX
const SITE_CHART_LABELS = siteStats.map(s => s.site);
const SITE_CHART_VALUES = siteStats.map(s => s.avg_10pt);

// Generate site insight text
function generateSiteInsightTexts() {
  const bestSite = siteStats.length > 0 ? siteStats.reduce((a, b) => a.avg_10pt > b.avg_10pt ? a : b) : null;
  const worstSite = siteStats.length > 0 ? siteStats.reduce((a, b) => a.avg_10pt < b.avg_10pt ? a : b) : null;

  const texts = [
    { text: "サイト別評価の傾向", options: { bold: true, fontSize: 12, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "高評価グループ", options: { bold: true, color: "16A34A", breakLine: true } },
  ];

  if (bestSite) {
    texts.push({ text: `${bestSite.site}(${bestSite.avg_10pt.toFixed(2)})が最高スコア。`, options: { fontSize: 9, color: "64748B", breakLine: true } });
  }

  texts.push({ text: "", options: { fontSize: 6, breakLine: true } });

  if (worstSite && worstSite.avg_10pt < overallAvg - 0.5) {
    texts.push({ text: "改善対象サイト", options: { bold: true, color: "DC2626", breakLine: true } });
    texts.push({ text: `${worstSite.site}(${worstSite.avg_10pt.toFixed(1)})が相対的に低い。口コミ返信の強化が急務`, options: { fontSize: 9, color: "64748B", breakLine: true } });
    texts.push({ text: "", options: { fontSize: 6, breakLine: true } });
  }

  texts.push({ text: "注：", options: { bold: true, fontSize: 9, color: "64748B" } });
  texts.push({ text: "Google/じゃらん/楽天は5点満点のため×2で10pt換算", options: { fontSize: 9, color: "64748B" } });

  return texts;
}

// Site table rows for PPTX
const SITE_TABLE_ROWS_DATA = siteStats.map(s => {
  const scaleText = s.scale === "/5" ? "/5 (×2)" : s.scale;
  return [s.site, String(s.count), String(s.native_avg), scaleText, String(s.avg_10pt)];
});

// Distribution data for PPTX
const DIST_DOUGHNUT_LABELS = ["高評価(8-10)", "中評価(5-7)", "低評価(1-4)"];
const DIST_DOUGHNUT_VALUES = [highCount, midCount, lowCount];

const DIST_BAR_LABELS = distribution.map(d => d.score + "点");
const DIST_BAR_VALUES = distribution.map(d => d.count);

const DIST_SUMMARY_CARDS = [
  { label: "高評価 8-10点", val: `${highCount}件（${highRate.toFixed(1)}%）`, col: "16A34A", bg: "DCFCE7" },
  { label: "中評価 5-7点", val: `${midCount}件（${midRate.toFixed(1)}%）`, col: "D4A843", bg: "FFF7ED" },
  { label: "低評価 1-4点", val: `${lowCount}件（${lowRate.toFixed(1)}%）`, col: lowRate <= 5 ? "16A34A" : "DC2626", bg: lowRate <= 5 ? "DCFCE7" : "FEE2E2" },
];

// Chart colors for distribution bar chart
function getDistBarColors() {
  return distribution.map(d => {
    if (d.score >= 8) return C.green;
    if (d.score >= 5) return C.gold;
    return C.red;
  });
}

// Strengths cards for PPTX
const STRENGTHS_CARDS = strengthsSorted.slice(0, 6).map(([name, d]) => {
  const countLabel = d.mentions >= 50 ? "50件+" : d.mentions >= 20 ? "20件+" : d.mentions >= 10 ? "10件+" : `${d.mentions}件`;
  return {
    theme: name,
    count: countLabel,
    desc: d.samples.length > 0 ? d.samples[0] : `${name}に関する好評価`,
    quote: d.samples.length > 1 ? `「${d.samples[1]}」` : (d.samples.length > 0 ? `「${d.samples[0]}」` : ""),
  };
});
while (STRENGTHS_CARDS.length < 6) {
  STRENGTHS_CARDS.push({ theme: "その他", count: "少数", desc: "その他のポジティブな評価", quote: "" });
}

// Weakness items for PPTX
const WEAKNESS_ITEMS = weaknessesSorted.slice(0, 9).map(([name, d], i) => {
  const priority = i === 0 ? "S" : i <= 2 ? "A" : i <= 4 ? "B" : "C";
  const color = priority === "S" ? "DC2626" : priority === "A" ? "EA580C" : priority === "B" ? "D4A843" : "94A3B8";
  return {
    pri: priority,
    cat: name,
    detail: d.samples.length > 0 ? d.samples[0] : `${name}に関する指摘`,
    count: `${d.mentions}件`,
    color,
  };
});

// Phase cards for PPTX
const PHASE1_CARDS = PHASE1_ITEMS.map(item => ({
  title: item.title.replace(/^\(\d+\)\s*/, ""),
  items: item.bullets,
}));

const PHASE2_ITEMS_PPTX = PHASE2_ITEMS.map(item => ({
  title: item.title.replace(/^\(\d+\)\s*/, ""),
  items: item.bullets,
}));

const PHASE3_ITEMS_PPTX = PHASE3_ITEMS.map(item => ({
  title: item.title.replace(/^\(\d+\)\s*/, ""),
  items: item.bullets,
}));

// KPI rows for PPTX
const KPI_TARGET_ROWS = KPI_TARGET_DATA;

// Closing paragraphs for PPTX
function generateClosingParagraphs() {
  const paragraphs = [];
  paragraphs.push({ text: CONCLUSION_PARAGRAPHS[0] || "", options: { breakLine: true, fontSize: 12 } });
  paragraphs.push({ text: "", options: { fontSize: 6, breakLine: true } });
  if (CONCLUSION_PARAGRAPHS[1]) {
    paragraphs.push({ text: CONCLUSION_PARAGRAPHS[1], options: { breakLine: true, fontSize: 12 } });
    paragraphs.push({ text: "", options: { fontSize: 6, breakLine: true } });
  }
  if (CONCLUSION_PARAGRAPHS[2]) {
    paragraphs.push({ text: CONCLUSION_PARAGRAPHS[2], options: { breakLine: true, fontSize: 12 } });
    paragraphs.push({ text: "", options: { fontSize: 6, breakLine: true } });
  }
  paragraphs.push({ text: CONCLUSION_FINAL, options: { bold: true, fontSize: 12, color: "D4A843" } });
  return paragraphs;
}

// ==========================================
// SLIDE 1: TITLE
// ==========================================
let s1 = pres.addSlide();
s1.background = { color: C.navy };

s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: C.navyLight, transparency: 40 } });
s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 2.0, w: 0.08, h: 1.8, fill: { color: C.gold } });

s1.addText("口コミ分析", { x: 0.8, y: 1.5, w: 8, h: 0.7, fontSize: 20, fontFace: "Arial", color: C.blueLight, charSpacing: 6, margin: 0 });
s1.addText("改善レポート", { x: 0.8, y: 2.1, w: 8, h: 0.9, fontSize: 42, fontFace: "Arial", color: C.white, bold: true, margin: 0 });
s1.addShape(pres.shapes.LINE, { x: 0.8, y: 3.1, w: 2, h: 0, line: { color: C.gold, width: 2 } });

s1.addText(HOTEL_NAME, { x: 0.8, y: 3.3, w: 8, h: 0.5, fontSize: 16, fontFace: "Arial", color: C.grayLight, margin: 0 });

s1.addText([
  { text: "分析対象期間：" + ANALYSIS_PERIOD, options: { breakLine: true } },
  { text: "レビュー総数：" + REVIEW_COUNT, options: { breakLine: true } },
  { text: "作成日：" + REPORT_DATE },
], { x: 0.8, y: 4.1, w: 5, h: 0.9, fontSize: 10, fontFace: "Arial", color: C.grayLight, margin: 0, paraSpaceAfter: 4 });

s1.addShape(pres.shapes.RECTANGLE, { x: 7.5, y: 0, w: 2.5, h: 5.625, fill: { color: C.blue, transparency: 85 } });

// ==========================================
// SLIDE 2: EXECUTIVE SUMMARY
// ==========================================
let s2 = pres.addSlide();
s2.background = { color: C.offWhite };
addContentHeader(s2, "エグゼクティブサマリー", "Executive Summary");

pptxKpiCard(s2, 0.5, 1.15, 2.05, 1.05, "全体平均(10pt換算)", pptxKpiAvg, C.blue, C.white);
pptxKpiCard(s2, 2.75, 1.15, 2.05, 1.05, "高評価率(8-10点)", pptxKpiHighRate, C.green, C.white);
pptxKpiCard(s2, 5.0, 1.15, 2.05, 1.05, "低評価率(1-4点)", pptxKpiLowRate, lowRate <= 5 ? C.green : C.red, C.white);
pptxKpiCard(s2, 7.25, 1.15, 2.25, 1.05, "レビュー総数", pptxKpiTotal, C.navy, C.white);

// Strengths box
s2.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 2.5, w: 4.3, h: 2.5, fill: { color: C.white }, shadow: shadow() });
s2.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 2.5, w: 4.3, h: 0.05, fill: { color: C.green } });
s2.addText("Strengths", { x: 0.7, y: 2.6, w: 3, h: 0.35, fontSize: 14, fontFace: "Arial", color: C.green, bold: true, margin: 0 });

const strengthTextItems = [];
SUMMARY_STRENGTHS.forEach((s, i) => {
  strengthTextItems.push({ text: s.title + " ", options: { bold: true, breakLine: true } });
  strengthTextItems.push({ text: "  " + s.desc, options: { fontSize: 9, color: C.gray, breakLine: i < SUMMARY_STRENGTHS.length - 1 } });
});
s2.addText(strengthTextItems, { x: 0.7, y: 3.0, w: 3.8, h: 1.8, fontSize: 11, fontFace: "Arial", color: C.grayDark, margin: 0, paraSpaceAfter: 2 });

// Weaknesses box
s2.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 2.5, w: 4.3, h: 2.5, fill: { color: C.white }, shadow: shadow() });
s2.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 2.5, w: 4.3, h: 0.05, fill: { color: C.red } });
s2.addText("Weaknesses", { x: 5.4, y: 2.6, w: 3, h: 0.35, fontSize: 14, fontFace: "Arial", color: C.red, bold: true, margin: 0 });

const weaknessTextItems = [];
SUMMARY_WEAKNESSES.forEach((w, i) => {
  weaknessTextItems.push({ text: w.title + " ", options: { bold: true, breakLine: true } });
  weaknessTextItems.push({ text: "  " + w.desc, options: { fontSize: 9, color: C.gray, breakLine: i < SUMMARY_WEAKNESSES.length - 1 } });
});
s2.addText(weaknessTextItems, { x: 5.4, y: 3.0, w: 3.8, h: 1.8, fontSize: 11, fontFace: "Arial", color: C.grayDark, margin: 0, paraSpaceAfter: 2 });

addFooter(s2, 2);

// ==========================================
// SLIDE 3: KPI CARDS (dedicated slide)
// ==========================================
let s3kpi = pres.addSlide();
s3kpi.background = { color: C.offWhite };
addContentHeader(s3kpi, "主要KPI", "Key Performance Indicators");

// Large KPI cards
pptxKpiCard(s3kpi, 0.5, 1.3, 4.3, 1.8, "全体平均スコア（10点換算）", pptxKpiAvg, scoreColor(overallAvg), C.white);
pptxKpiCard(s3kpi, 5.2, 1.3, 4.3, 1.8, "高評価率（8-10点）", pptxKpiHighRate, highRate >= 80 ? C.green : C.gold, C.white);
pptxKpiCard(s3kpi, 0.5, 3.4, 4.3, 1.5, "低評価率（1-4点）", pptxKpiLowRate, lowRate <= 5 ? C.green : C.red, C.white);
pptxKpiCard(s3kpi, 5.2, 3.4, 4.3, 1.5, "レビュー総数", pptxKpiTotal, C.navy, C.white);

addFooter(s3kpi, 3);

// ==========================================
// SLIDE 4: SITE-BY-SITE RATINGS
// ==========================================
let s3 = pres.addSlide();
s3.background = { color: C.offWhite };
addContentHeader(s3, "サイト別評価分析", "Rating Analysis by Platform");

if (SITE_CHART_LABELS.length > 0) {
  s3.addChart(pres.charts.BAR, [
    { name: "10pt換算平均", labels: SITE_CHART_LABELS, values: SITE_CHART_VALUES }
  ], {
    x: 0.5, y: 1.1, w: 5.5, h: 3.5,
    barDir: "bar",
    chartColors: [C.blue],
    chartArea: { fill: { color: C.white }, roundedCorners: true },
    catAxisLabelColor: C.grayDark, catAxisLabelFontSize: 10,
    valAxisLabelColor: C.gray, valAxisLabelFontSize: 9,
    valGridLine: { color: "E2E8F0", size: 0.5 },
    catGridLine: { style: "none" },
    showValue: true, dataLabelPosition: "outEnd", dataLabelColor: C.grayDark,
    showLegend: false,
    valAxisMaxVal: 10,
  });
}

// Insight box
s3.addShape(pres.shapes.RECTANGLE, { x: 6.3, y: 1.1, w: 3.3, h: 3.5, fill: { color: C.white }, shadow: shadow() });
s3.addShape(pres.shapes.RECTANGLE, { x: 6.3, y: 1.1, w: 3.3, h: 0.05, fill: { color: C.gold } });
s3.addText("Insight", { x: 6.5, y: 1.2, w: 2, h: 0.35, fontSize: 13, fontFace: "Arial", color: C.gold, bold: true, margin: 0 });

s3.addText(generateSiteInsightTexts(), { x: 6.5, y: 1.55, w: 2.9, h: 2.9, fontSize: 11, fontFace: "Arial", color: C.grayDark, margin: 0, paraSpaceAfter: 2 });

// Table
const tblHeader2 = [
  { text: "サイト名", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 9, fontFace: "Arial", align: "center" } },
  { text: "件数", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 9, fontFace: "Arial", align: "center" } },
  { text: "ネイティブ平均", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 9, fontFace: "Arial", align: "center" } },
  { text: "尺度", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 9, fontFace: "Arial", align: "center" } },
  { text: "10pt換算", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 9, fontFace: "Arial", align: "center" } },
];

const tblRows = SITE_TABLE_ROWS_DATA.map((r, i) => r.map(cell => ({
  text: cell, options: { fontSize: 9, fontFace: "Arial", align: "center", fill: { color: i % 2 === 0 ? C.ice : C.white } }
})));

const rowHeights = [0.25, ...SITE_TABLE_ROWS_DATA.map(() => 0.22)];
s3.addTable([tblHeader2, ...tblRows], {
  x: 0.5, y: 4.7, w: 9, h: 0.1,
  colW: [2.0, 1.0, 2.0, 1.0, 3.0],
  border: { pt: 0.5, color: "CBD5E1" },
  rowH: rowHeights,
});

addFooter(s3, 4);

// ==========================================
// SLIDE 5: RATING DISTRIBUTION
// ==========================================
let s4 = pres.addSlide();
s4.background = { color: C.offWhite };
addContentHeader(s4, "評価分布分析（10点換算）", "Rating Distribution");

s4.addChart(pres.charts.DOUGHNUT, [{
  name: "評価分布（10pt換算）",
  labels: DIST_DOUGHNUT_LABELS,
  values: DIST_DOUGHNUT_VALUES,
}], {
  x: 0.3, y: 1.2, w: 4.0, h: 3.2,
  chartColors: [C.green, C.gold, C.red],
  showPercent: true,
  dataLabelColor: C.white,
  dataLabelFontSize: 11,
  showTitle: false,
  showLegend: true,
  legendPos: "b",
  legendFontSize: 9,
  legendColor: C.gray,
});

if (DIST_BAR_LABELS.length > 0) {
  s4.addChart(pres.charts.BAR, [{
    name: "件数",
    labels: DIST_BAR_LABELS,
    values: DIST_BAR_VALUES,
  }], {
    x: 4.5, y: 1.2, w: 5.2, h: 3.2,
    barDir: "col",
    chartColors: getDistBarColors(),
    chartArea: { fill: { color: C.white }, roundedCorners: true },
    catAxisLabelColor: C.grayDark, catAxisLabelFontSize: 9,
    valAxisLabelColor: C.gray, valAxisLabelFontSize: 8,
    valGridLine: { color: "E2E8F0", size: 0.5 },
    catGridLine: { style: "none" },
    showValue: true, dataLabelPosition: "outEnd", dataLabelColor: C.grayDark, dataLabelFontSize: 9,
    showLegend: false,
  });
}

DIST_SUMMARY_CARDS.forEach((c, i) => {
  const x = 0.5 + i * 3.1;
  s4.addShape(pres.shapes.RECTANGLE, { x, y: 4.55, w: 2.9, h: 0.55, fill: { color: c.bg } });
  s4.addText(c.label, { x, y: 4.55, w: 1.5, h: 0.55, fontSize: 10, fontFace: "Arial", color: c.col, bold: true, valign: "middle", margin: [0,0,0,8] });
  s4.addText(c.val, { x: x + 1.4, y: 4.55, w: 1.5, h: 0.55, fontSize: 11, fontFace: "Arial", color: c.col, bold: true, align: "right", valign: "middle", margin: [0,8,0,0] });
});

addFooter(s4, 5);

// ==========================================
// SLIDE 6: STRENGTHS
// ==========================================
let s5 = pres.addSlide();
s5.background = { color: C.offWhite };
addContentHeader(s5, "強み分析", "Strength Analysis");

STRENGTHS_CARDS.forEach((s, i) => {
  const row = Math.floor(i / 3);
  const col = i % 3;
  const x = 0.5 + col * 3.1;
  const y = 1.15 + row * 2.0;

  s5.addShape(pres.shapes.RECTANGLE, { x, y, w: 2.9, h: 1.75, fill: { color: C.white }, shadow: shadow() });
  s5.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.06, h: 1.75, fill: { color: C.green } });

  s5.addShape(pres.shapes.RECTANGLE, { x: x + 2.0, y: y + 0.08, w: 0.72, h: 0.25, fill: { color: C.greenBg } });
  s5.addText(s.count, { x: x + 2.0, y: y + 0.08, w: 0.72, h: 0.25, fontSize: 9, fontFace: "Arial", color: C.green, bold: true, align: "center", valign: "middle", margin: 0 });

  s5.addText(s.theme, { x: x + 0.15, y: y + 0.08, w: 1.8, h: 0.3, fontSize: 13, fontFace: "Arial", color: C.navy, bold: true, margin: 0 });
  s5.addText(s.desc, { x: x + 0.15, y: y + 0.4, w: 2.6, h: 0.4, fontSize: 9, fontFace: "Arial", color: C.gray, margin: 0 });
  s5.addText(s.quote, { x: x + 0.15, y: y + 0.85, w: 2.6, h: 0.7, fontSize: 9, fontFace: "Arial", color: C.blue, italic: true, margin: 0 });
});

addFooter(s5, 6);

// ==========================================
// SLIDE 7: WEAKNESS / PRIORITY MATRIX
// ==========================================
let s6 = pres.addSlide();
s6.background = { color: C.offWhite };
addContentHeader(s6, "弱み分析・優先度マトリクス", "Weakness Analysis & Priority Matrix");

const wTblHeader = [
  { text: "優先度", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 9, fontFace: "Arial", align: "center" } },
  { text: "課題カテゴリ", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 9, fontFace: "Arial" } },
  { text: "具体的内容", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 9, fontFace: "Arial" } },
  { text: "件数", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 9, fontFace: "Arial", align: "center" } },
];

const wTblRows = WEAKNESS_ITEMS.map((w, i) => [
  { text: w.pri, options: { fontSize: 12, fontFace: "Arial", align: "center", bold: true, color: w.color, fill: { color: i % 2 === 0 ? C.ice : C.white } } },
  { text: w.cat, options: { fontSize: 9, fontFace: "Arial", bold: true, color: C.grayDark, fill: { color: i % 2 === 0 ? C.ice : C.white } } },
  { text: w.detail, options: { fontSize: 9, fontFace: "Arial", color: C.gray, fill: { color: i % 2 === 0 ? C.ice : C.white } } },
  { text: w.count, options: { fontSize: 9, fontFace: "Arial", align: "center", color: C.grayDark, fill: { color: i % 2 === 0 ? C.ice : C.white } } },
]);

const weakRowHeights = [0.3, ...WEAKNESS_ITEMS.map(() => 0.35)];
s6.addTable([wTblHeader, ...wTblRows], {
  x: 0.5, y: 1.15, w: 9, h: 0.1,
  colW: [0.8, 2.0, 5.0, 1.2],
  border: { pt: 0.5, color: "CBD5E1" },
  rowH: weakRowHeights,
});

// Legend
s6.addText([
  { text: "S ", options: { bold: true, color: C.red, fontSize: 10 } },
  { text: "= 最優先   ", options: { fontSize: 9, color: C.gray } },
  { text: "A ", options: { bold: true, color: C.orange, fontSize: 10 } },
  { text: "= 高優先   ", options: { fontSize: 9, color: C.gray } },
  { text: "B ", options: { bold: true, color: C.gold, fontSize: 10 } },
  { text: "= 中優先   ", options: { fontSize: 9, color: C.gray } },
  { text: "C ", options: { bold: true, color: C.grayLight, fontSize: 10 } },
  { text: "= 低優先", options: { fontSize: 9, color: C.gray } },
], { x: 0.5, y: 4.8, w: 9, h: 0.3, fontFace: "Arial", margin: 0 });

addFooter(s6, 7);

// ==========================================
// SLIDE 8: IMPROVEMENT PHASE 1
// ==========================================
let s7 = pres.addSlide();
s7.background = { color: C.offWhite };
addContentHeader(s7, "改善施策 Phase 1：即座対応", "Immediate Actions (Today ~ 1 Month)");

s7.addText("投資不要・オペレーション改善で対応可能", { x: 0.6, y: 0.95, w: 8, h: 0.3, fontSize: 10, fontFace: "Arial", color: C.blue, italic: true, margin: 0 });

PHASE1_CARDS.forEach((p, i) => {
  const col = i % 2;
  const row = Math.floor(i / 2);
  const x = 0.5 + col * 4.6;
  const y = 1.35 + row * 1.95;

  s7.addShape(pres.shapes.RECTANGLE, { x, y, w: 4.35, h: 1.75, fill: { color: C.white }, shadow: shadow() });
  s7.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.06, h: 1.75, fill: { color: C.blue } });

  s7.addText(p.title, { x: x + 0.15, y: y + 0.05, w: 4, h: 0.35, fontSize: 12, fontFace: "Arial", color: C.navy, bold: true, margin: 0 });

  const bulletItems = p.items.map((item, idx) => ({
    text: item,
    options: { bullet: true, fontSize: 9, color: C.grayDark, breakLine: idx < p.items.length - 1 }
  }));
  s7.addText(bulletItems, { x: x + 0.15, y: y + 0.4, w: 4, h: 1.3, fontFace: "Arial", margin: 0, paraSpaceAfter: 3 });
});

addFooter(s7, 8);

// ==========================================
// SLIDE 9: IMPROVEMENT PHASE 2 & 3
// ==========================================
let s8 = pres.addSlide();
s8.background = { color: C.offWhite };
addContentHeader(s8, "改善施策 Phase 2・3", "Short-term & Mid-term Actions");

// Phase 2
s8.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.1, w: 4.35, h: 3.8, fill: { color: C.white }, shadow: shadow() });
s8.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.1, w: 4.35, h: 0.4, fill: { color: C.blue } });
s8.addText("Phase 2：短期施策（1〜3ヶ月）", { x: 0.65, y: 1.1, w: 4, h: 0.4, fontSize: 12, fontFace: "Arial", color: C.white, bold: true, margin: 0, valign: "middle" });

let p2y = 1.6;
PHASE2_ITEMS_PPTX.forEach((p) => {
  s8.addText(p.title, { x: 0.7, y: p2y, w: 3.8, h: 0.3, fontSize: 10, fontFace: "Arial", color: C.navy, bold: true, margin: 0 });
  const items = p.items.map((item, idx) => ({
    text: item, options: { bullet: true, fontSize: 9, color: C.gray, breakLine: idx < p.items.length - 1 }
  }));
  s8.addText(items, { x: 0.7, y: p2y + 0.28, w: 3.8, h: p.items.length * 0.22 + 0.1, fontFace: "Arial", margin: 0, paraSpaceAfter: 2 });
  p2y += 0.28 + p.items.length * 0.22 + 0.2;
});

// Phase 3
s8.addShape(pres.shapes.RECTANGLE, { x: 5.15, y: 1.1, w: 4.35, h: 3.8, fill: { color: C.white }, shadow: shadow() });
s8.addShape(pres.shapes.RECTANGLE, { x: 5.15, y: 1.1, w: 4.35, h: 0.4, fill: { color: C.navy } });
s8.addText("Phase 3：中期施策（3〜6ヶ月）", { x: 5.3, y: 1.1, w: 4, h: 0.4, fontSize: 12, fontFace: "Arial", color: C.white, bold: true, margin: 0, valign: "middle" });

let p3y = 1.6;
PHASE3_ITEMS_PPTX.forEach((p) => {
  s8.addText(p.title, { x: 5.35, y: p3y, w: 3.8, h: 0.3, fontSize: 10, fontFace: "Arial", color: C.navy, bold: true, margin: 0 });
  const items = p.items.map((item, idx) => ({
    text: item, options: { bullet: true, fontSize: 9, color: C.gray, breakLine: idx < p.items.length - 1 }
  }));
  s8.addText(items, { x: 5.35, y: p3y + 0.28, w: 3.8, h: p.items.length * 0.22 + 0.1, fontFace: "Arial", margin: 0, paraSpaceAfter: 2 });
  p3y += 0.28 + p.items.length * 0.22 + 0.2;
});

addFooter(s8, 9);

// ==========================================
// SLIDE 10: KPI TARGETS & CLOSING
// ==========================================
let s9 = pres.addSlide();
s9.background = { color: C.offWhite };
addContentHeader(s9, "KPI目標設定・まとめ", "Key Performance Indicators & Summary");

const kpiHeader = [
  { text: "KPI項目", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 10, fontFace: "Arial" } },
  { text: "現状値", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 10, fontFace: "Arial", align: "center" } },
  { text: "目標値", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 10, fontFace: "Arial", align: "center" } },
  { text: "期限", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 10, fontFace: "Arial", align: "center" } },
];

const kpiDataPptx = KPI_TARGET_ROWS.map((r, i) => [
  { text: r[0], options: { fontSize: 10, fontFace: "Arial", bold: true, color: C.grayDark, fill: { color: i % 2 === 0 ? C.ice : C.white } } },
  { text: r[1], options: { fontSize: 10, fontFace: "Arial", align: "center", color: C.gray, fill: { color: i % 2 === 0 ? C.ice : C.white } } },
  { text: r[2], options: { fontSize: 10, fontFace: "Arial", align: "center", bold: true, color: C.green, fill: { color: i % 2 === 0 ? C.ice : C.white } } },
  { text: r[3], options: { fontSize: 10, fontFace: "Arial", align: "center", color: C.gray, fill: { color: i % 2 === 0 ? C.ice : C.white } } },
]);

const kpiRowH = [0.35, ...KPI_TARGET_ROWS.map(() => 0.35)];
s9.addTable([kpiHeader, ...kpiDataPptx], {
  x: 0.5, y: 1.2, w: 9, h: 0.1,
  colW: [2.8, 2.0, 2.2, 2.0],
  border: { pt: 0.5, color: "CBD5E1" },
  rowH: kpiRowH,
});

const kpiTableBottom = 1.2 + kpiRowH.reduce((a, b) => a + b, 0) + 0.1;
s9.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: Math.min(kpiTableBottom, 4.25), w: 9, h: 0.7, fill: { color: C.white }, shadow: shadow() });
s9.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: Math.min(kpiTableBottom, 4.25), w: 0.06, h: 0.7, fill: { color: C.blue } });
s9.addText([
  { text: "モニタリング方針：", options: { bold: true, color: C.navy } },
  { text: "毎月のサイト別スコア集計と口コミ内容の定性分析を実施。四半期ごとに改善施策の効果を検証し、Phase 2・3の優先順位を見直します。", options: { color: C.gray } },
], { x: 0.7, y: Math.min(kpiTableBottom, 4.25), w: 8.6, h: 0.7, fontSize: 10, fontFace: "Arial", valign: "middle", margin: 0 });

addFooter(s9, 10);

// ============================================================
// Save both files
// ============================================================
const docxPath = path.join(OUTPUT_DIR, `${OUTPUT_PREFIX}_口コミ分析改善レポート.docx`);
const pptxPath = path.join(OUTPUT_DIR, `${OUTPUT_PREFIX}_口コミ分析レポート.pptx`);

async function saveAll() {
  try {
    // Save DOCX
    const buffer = await Packer.toBuffer(doc);
    fs.writeFileSync(docxPath, buffer);
    console.log("DOCX created: " + docxPath);
    console.log("DOCX size: " + (buffer.length / 1024).toFixed(1) + " KB");

    // Save PPTX
    await pres.writeFile({ fileName: pptxPath });
    console.log("PPTX created: " + pptxPath);
    const pptxStats = fs.statSync(pptxPath);
    console.log("PPTX size: " + (pptxStats.size / 1024).toFixed(1) + " KB");

    console.log("\nDone! Both reports generated successfully for: " + HOTEL_NAME);
  } catch (err) {
    console.error("Error generating reports:", err);
    process.exit(1);
  }
}

saveAll();
