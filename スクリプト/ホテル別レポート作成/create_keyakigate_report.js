const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, PageNumber, PageBreak, LevelFormat
} = require("docx");

// ============================================================
// ホテル固有の設定
// ============================================================
const HOTEL_NAME = "ホテルケヤキゲート東京府中";
const ANALYSIS_PERIOD = "2026年2月〜3月";
const REPORT_DATE = "2026年3月7日";
const OUTPUT_DIR = "/Users/mitsugutakahashi/ホテル口コミ";

const REVIEW_COUNT = "66件（重複除外後）";
const TARGET_SITES = "対象サイト：Booking.com / Trip.com / じゃらん / 楽天トラベル / Agoda / Google";

const KPI_CARDS = [
  { label: "全体平均(10pt換算)", value: "8.74", color: "27AE60", bgColor: "E8F5E9" },
  { label: "高評価率(8-10点)", value: "84.8%", color: "27AE60", bgColor: "E8F5E9" },
  { label: "低評価率(1-4点)", value: "4.5%", color: "E67E22", bgColor: "FFF3E0" },
  { label: "レビュー総数", value: "66件", color: "1B3A5C", bgColor: "D5E8F0" },
];

const EXECUTIVE_SUMMARY_INTRO = "2026年2月〜3月に各OTAサイト・口コミサイトに投稿された66件のレビューを包括的に分析しました。以下が主要な発見事項です。";
const EXECUTIVE_SUMMARY_EVALUATION_1 = "ホテルケヤキゲート東京府中は、全体平均8.74点（10点換算）と高水準の評価を維持しています。特にTrip.comでは満点10.0点を獲得し、楽天トラベル・Agoda・じゃらん・Booking.comでも8.7点以上の良好な評価を得ています。高評価率（8-10点）は84.8%に達しており、大多数のゲストが満足していることが分かります。";
const EXECUTIVE_SUMMARY_EVALUATION_2 = "一方、低評価率は4.5%（3件）と少数ながら、清掃品質に関する厳しい指摘が含まれており、改善が急務です。Googleの評価が6.0点と他サイトを大きく下回っている点も注視が必要です。";

const KEY_FINDING_STRENGTH = "駅直結の圧倒的な立地利便性（35件以上で言及）、朝食の高品質（20件以上）、新しく清潔な施設（15件以上）が三大強みとして確認されました。";
const KEY_FINDING_WEAKNESS = "清掃品質のばらつき（シーツの汚れ・毛髪4件）、スタッフ間の情報連携不足（3件）、空調・乾燥問題（2件）が主要な改善課題です。";
const KEY_FINDING_OPPORTUNITY = "リピート意向が非常に高く（「また利用したい」30件以上）、清掃品質の安定化とWi-Fi環境の改善で高評価率90%以上を目指せるポテンシャルがあります。";

// サイト別レビューデータ（10pt換算降順）
const SITE_DATA = [
  ["Trip.com",     "4",  "10.00", "/10", "10.00", "10.0", "優秀"],
  ["楽天トラベル",  "21", "4.43",  "/5",  "8.86",  "10.0", "良好"],
  ["Agoda",        "12", "8.83",  "/10", "8.83",  "9.5",  "良好"],
  ["じゃらん",     "16", "4.38",  "/5",  "8.75",  "9.0",  "良好"],
  ["Booking.com",  "10", "8.70",  "/10", "8.70",  "9.5",  "良好"],
  ["Google",       "3",  "3.00",  "/5",  "6.00",  "6.0",  "要改善"],
];

// 評価分布データ
const DISTRIBUTION_DATA = [
  [10, 38, "57.6%"],
  [9,  3,  "4.5%"],
  [8,  15, "22.7%"],
  [6,  7,  "10.6%"],
  [4,  1,  "1.5%"],
  [2,  2,  "3.0%"],
];
const DISTRIBUTION_MAX_COUNT = 38;

const HIGH_RATING_SUMMARY = "56件（84.8%）";
const MID_RATING_SUMMARY = "7件（10.6%）";
const LOW_RATING_SUMMARY = "3件（4.5%）";
const LOW_RATING_COLOR = "C0392B";

const DATA_OVERVIEW_SITE_TEXT = "各予約サイトの評価基準が異なる点に留意が必要です。海外OTA（Booking.com/Trip.com/Agoda）は10点満点、国内サイト（じゃらん/楽天トラベル/Google）は5点満点のため、統一比較のため10点換算（5点満点×2）を行いました。";
const DATA_OVERVIEW_DIST_TEXT = "10点（満点）評価が38件（57.6%）と最多を占め、8点以上の高評価が全体の84.8%を占めています。一方、4点以下の低評価は3件（4.5%）にとどまりますが、いずれも清掃品質やスタッフ対応に関する具体的な不満が含まれています。";

// 強みテーマデータ
const STRENGTH_THEMES = [
  ["立地・アクセス", "35件以上", "「駅直結で雨に濡れない」「府中駅から徒歩1分」「周辺に飲食店・コンビニ充実」"],
  ["朝食の質", "20件以上", "「一品一品が美味しい」「高級ホテルを凌ぐクオリティ」「お茶漬けが最高」"],
  ["清潔感・新しさ", "15件以上", "「新しくて清潔」「施設が綺麗」「部屋も清潔で安心」"],
  ["スタッフ対応", "8件", "「フロントの方も感じ良い対応」「きめ細やかなサービス」「子供用セットの心遣い」"],
  ["部屋・設備の快適さ", "10件", "「部屋が広々」「ベッドサイズに余裕」「お風呂で足が伸ばせる」"],
  ["コスパ・リピート意向", "30件以上", "「また利用したい」「リーズナブル」「この周辺で1番おすすめ」"],
];

const STRENGTH_SUB_1_TITLE = "3.1 最大の強み：「駅直結の圧倒的立地」";
const STRENGTH_SUB_1_TEXT = "京王線府中駅と直結している点は、当ホテル最大の差別化要因です。特に雨天時でも濡れずにアクセスできる点が高く評価されています。周辺にはショッピングモール、飲食店、コンビニ、ドラッグストアが充実しており、滞在中の利便性も極めて高いと評価されています。";
const STRENGTH_SUB_1_BULLETS = [
  "府中駅から屋根付きデッキで直結（徒歩約1分）",
  "特急停車駅のため、都心からのアクセスも良好",
  "周辺施設（飲食店・コンビニ・ドラッグストア）の充実度が高い",
  "サッカー観戦（味の素スタジアム）やピューロランド訪問の拠点としても好評",
];

const STRENGTH_SUB_2_TITLE = "3.2 朝食の高品質が差別化ポイント";
const STRENGTH_SUB_2_TEXT = "朝食ビュッフェの品質は多くのゲストから絶賛されており、「高級ホテルを凌ぐクオリティ」との評価もあります。品数は多くないものの一品一品の質が高く、お茶漬けやカレー、だし汁かけご飯など和食メニューが特に好評です。季節限定メニュー（正月仕様等）の心遣いも高く評価されています。";

// 弱み分析データ
const WEAKNESS_PRIORITY_DATA = [
  ["S", "清掃品質のばらつき", "シーツに毛やゴミが残っている（複数件報告）、風呂場に毛髪、置き時計の遅れ、部屋着の匂いなど清掃チェック漏れが散見される", "直接的な低評価要因・4件"],
  ["A", "スタッフ間の情報連携不足", "連泊時の清掃不要の申し送りが伝わらない、アメニティ依頼の漏れ、フロントで同じことを何度も聞かれる", "サービス品質への不信感・3件"],
  ["A", "フロント対応の温度差", "スタッフにより対応の質にばらつきがある（愛想がない、マニュアル的対応）。チェックアウト時のスタッフは笑顔で好印象との対比あり", "顧客満足度への影響・2件"],
  ["B", "空調・室内環境", "エアコンで室内が十分暖まらない、冬季の乾燥（加湿器未設置）", "快適性への影響・2件"],
  ["B", "Wi-Fi接続品質", "2.4GHz帯のみの無線LAN設計により、夕方〜夜に頻繁な切断が発生。動画視聴に支障あり", "ビジネス・レジャー利用への支障・2件"],
  ["B", "眺望・客室情報の事前説明", "すりガラス窓で眺望なしの部屋について、予約時に条件として明示されていない", "期待値とのギャップ・1件"],
  ["C", "駐車場の不在", "提携駐車場がなく、ゲスト自身でコインパーキングを探す必要がある", "車利用ゲストの不便・2件"],
  ["C", "共用設備の混雑", "合宿利用者による洗濯機・乾燥機の占有、ロビーでの居眠り", "他ゲストへの影響・2件"],
];

// 改善施策 Phase 1
const PHASE1_DESCRIPTION = "投資不要・オペレーション改善で対応可能な施策";
const PHASE1_ITEMS = [
  {
    title: "(1) 清掃品質チェック体制の強化",
    bullets: [
      "チェックリストの導入：シーツ・枕カバーの毛髪確認、バスルーム清掃確認、備品動作確認（時計等）を必須項目化",
      "ダブルチェック制度：清掃担当者とは別のスタッフによる抜き打ち検査を1日3室以上実施",
      "部屋着の洗濯・品質管理基準の見直し：匂い残りがないか嗅覚チェックを追加",
      "清掃品質のフィードバックループ：口コミ指摘を即日清掃チームに共有する仕組みの構築",
    ],
  },
  {
    title: "(2) スタッフ間情報連携の改善",
    bullets: [
      "連泊ゲストの申し送り事項をPMS（宿泊管理システム）に必ず記録し、全スタッフが参照できる体制に",
      "シフト交代時の引き継ぎチェックリスト導入（特記事項・リクエスト内容の確認）",
      "アメニティ依頼の管理：依頼受付から準備・配達のステータス管理を明確化",
    ],
  },
  {
    title: "(3) フロント接客品質の標準化",
    bullets: [
      "ゲスト到着時の第一声・笑顔の統一基準を設定",
      "早着ゲストへの対応マニュアル整備（待機中のお声がけ、お待たせ時の謝辞）",
      "月1回のロールプレイング研修の実施（特にチェックイン・クレーム対応）",
    ],
  },
  {
    title: "(4) 眺望なし客室の事前情報開示",
    bullets: [
      "すりガラス窓の客室は予約サイトの客室説明に「窓からの眺望はございません」と明記",
      "写真掲載時にも当該客室の窓の状態が分かる画像を追加",
      "予約確認メールで客室タイプごとの特徴を案内",
    ],
  },
];

// 改善施策 Phase 2
const PHASE2_DESCRIPTION = "一定の投資を伴うが、比較的早期に実行可能な施策";
const PHASE2_ITEMS = [
  {
    title: "(1) Wi-Fi環境の抜本改善",
    bullets: [
      "5GHz帯対応のアクセスポイント増設とチャネル設計の最適化",
      "キャスティング用ネットワークと一般Wi-Fiの帯域分離",
      "各フロアの接続テストを定期実施（特に夕方〜夜間帯）",
    ],
  },
  {
    title: "(2) 客室環境の改善",
    bullets: [
      "各室への加湿器の常設配備（冬季は特にニーズ高）",
      "空調設定の見直し・メンテナンス実施",
    ],
  },
  {
    title: "(3) Google口コミ対策の強化",
    bullets: [
      "Googleの評価が6.0点と低水準のため、チェックアウト時にQRコードで投稿を依頼",
      "Google口コミへの返信を100%・48時間以内に実施",
    ],
  },
];

// 改善施策 Phase 3
const PHASE3_DESCRIPTION = "設備投資を伴う抜本的な改善施策";
const PHASE3_ITEMS = [
  {
    title: "(1) 駐車場ソリューションの構築",
    bullets: [
      "近隣コインパーキングとの提携契約締結（割引コード発行）",
      "予約サイト・公式サイトでの駐車場情報の充実化",
      "駐車場予約代行サービスの導入検討",
    ],
  },
  {
    title: "(2) 朝食メニューのさらなる充実",
    bullets: [
      "季節限定メニューのローテーション強化",
      "ベーコンなど洋食メニューの品目追加検討",
      "朝食の品数バリエーション拡大",
    ],
  },
  {
    title: "(3) 共用スペースの利用ルール整備",
    bullets: [
      "ランドリールームの利用時間制限・予約制の導入",
      "ロビーエリアの利用マナー掲示と巡回強化",
    ],
  },
];

// KPI目標データ
const KPI_TARGET_DATA = [
  ["全体平均(10pt換算)", "8.74点", "9.0点以上", "2026年9月"],
  ["高評価率（8-10点）", "84.8%", "90%以上", "2026年9月"],
  ["低評価率（1-4点）", "4.5%", "2%以下", "2026年9月"],
  ["清掃関連クレーム", "4件/2ヶ月", "0件/2ヶ月", "2026年6月"],
  ["Google平均評価", "3.0/5点", "4.0/5点以上", "2026年12月"],
  ["口コミ返信率", "未計測", "100%（48h以内）", "2026年6月"],
  ["楽天トラベル平均", "4.43/5点", "4.6/5点以上", "2026年9月"],
  ["リピート意向率", "約50%（推定）", "60%以上", "2026年12月"],
];

// 総括テキスト
const CONCLUSION_PARAGRAPHS = [
  "ホテルケヤキゲート東京府中は、京王線府中駅直結という圧倒的な立地優位性と、高品質な朝食、清潔で新しい施設を武器に、全体平均8.74点（10点換算）という高い評価を獲得しています。",
  "最優先課題は清掃品質の安定化です。シーツの毛髪やバスルームの清掃漏れは、他の優れた要素を一瞬で台無しにしかねません。チェックリストの徹底とダブルチェック体制の構築を即座に実行すべきです。",
  "次に、スタッフ間の情報連携とフロント接客品質の標準化に取り組むことで、「人」に起因する不満を解消できます。Wi-Fi環境の改善は、ビジネス・レジャー双方の顧客満足度を大きく向上させます。",
  "これらの改善を着実に実行することで、高評価率90%以上、全体平均9.0点以上の達成は十分に現実的な目標です。",
];
const CONCLUSION_FINAL_PARAGRAPH = "当ホテルの「立地」「朝食」「清潔感」という三大強みをさらに磨き上げ、弱みを一つずつ解消していくことで、府中エリアNo.1のホテルとしてのポジションを確固たるものにしていきましょう。";

// ============================================================
// Color scheme
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
// Helper functions
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
function bulletItem(text, opts = {}) {
  return new Paragraph({
    numbering: { reference: "bullets", level: opts.level || 0 },
    spacing: { after: 80, line: 300 },
    children: [new TextRun({ text, size: 21, font: "Arial", color: "333333" })],
  });
}
function headerCell(text, width) {
  return new TableCell({
    borders, width: { size: width, type: WidthType.DXA },
    shading: { fill: NAVY, type: ShadingType.CLEAR }, margins: cellMargins, verticalAlign: "center",
    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text, bold: true, size: 20, font: "Arial", color: WHITE })] })],
  });
}
function dataCell(text, width, opts = {}) {
  return new TableCell({
    borders, width: { size: width, type: WidthType.DXA },
    shading: opts.fill ? { fill: opts.fill, type: ShadingType.CLEAR } : undefined,
    margins: cellMargins, verticalAlign: "center",
    children: [new Paragraph({ alignment: opts.alignment || AlignmentType.LEFT, children: [new TextRun({ text: String(text), size: 20, font: "Arial", color: opts.color || "333333", bold: opts.bold || false })] })],
  });
}
function spacer(height = 100) { return new Paragraph({ spacing: { after: height }, children: [] }); }
function divider() { return new Paragraph({ spacing: { before: 200, after: 200 }, border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: ACCENT, space: 1 } }, children: [] }); }

function kpiRow(items) {
  const colWidth = Math.floor(9360 / items.length);
  return new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: items.map(() => colWidth),
    rows: [new TableRow({ children: items.map(item => new TableCell({
      borders: { ...noBorders, right: { style: BorderStyle.SINGLE, size: 1, color: "DDDDDD" } },
      width: { size: colWidth, type: WidthType.DXA }, shading: { fill: item.bgColor || LIGHT_BLUE, type: ShadingType.CLEAR },
      margins: { top: 160, bottom: 160, left: 200, right: 200 },
      children: [
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [new TextRun({ text: item.label, size: 18, font: "Arial", color: "666666" })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: item.value, bold: true, size: 36, font: "Arial", color: item.color || NAVY })] }),
      ],
    })) })] });
}

function priorityRow(priority, category, detail, impact) {
  const colors = { "S": RED_ACCENT, "A": ORANGE_ACCENT, "B": ACCENT, "C": "888888" };
  return new TableRow({ children: [
    new TableCell({ borders, width: { size: 800, type: WidthType.DXA }, margins: cellMargins, shading: { fill: LIGHT_GRAY, type: ShadingType.CLEAR }, verticalAlign: "center",
      children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: priority, bold: true, size: 22, font: "Arial", color: colors[priority] || NAVY })] })] }),
    dataCell(category, 2200, { bold: true }), dataCell(detail, 4360), dataCell(impact, 2000, { alignment: AlignmentType.CENTER }),
  ] });
}

function buildPhaseItems(items) {
  const result = [];
  items.forEach((item, index) => { result.push(heading3(item.title)); item.bullets.forEach(b => result.push(bulletItem(b))); if (index < items.length - 1) result.push(spacer(80)); });
  return result;
}
function buildSiteTableRows(data) {
  return data.map((row, index) => {
    const [siteName, count, nativeAvg, scale, tenPt, median, verdict] = row;
    const fill = index % 2 === 1 ? LIGHT_GRAY : undefined;
    const verdictColor = verdict === "優秀" ? GREEN_ACCENT : (verdict === "良好" ? GREEN_ACCENT : (verdict === "概ね良好" ? ORANGE_ACCENT : RED_ACCENT));
    return new TableRow({ children: [
      dataCell(siteName, 1800, { bold: true, fill }), dataCell(count, 900, { alignment: AlignmentType.CENTER, fill }),
      dataCell(nativeAvg, 1300, { alignment: AlignmentType.CENTER, bold: true, fill }), dataCell(scale, 900, { alignment: AlignmentType.CENTER, fill }),
      dataCell(tenPt, 1300, { alignment: AlignmentType.CENTER, color: verdictColor, bold: true, fill }),
      dataCell(median, 1526, { alignment: AlignmentType.CENTER, fill }), dataCell(verdict, 1300, { alignment: AlignmentType.CENTER, color: verdictColor, bold: true, fill }),
    ]});
  });
}
function buildDistributionRows(data, maxCount) {
  return data.map(([rating, count, pct]) => {
    const barWidth = Math.round(count / maxCount * 100);
    const barColor = rating >= 8 ? GREEN_ACCENT : (rating >= 5 ? ORANGE_ACCENT : RED_ACCENT);
    const fill = [10, 8, 6, 4, 2].includes(rating) ? LIGHT_GRAY : WHITE;
    return new TableRow({ children: [
      dataCell(String(rating), 1200, { alignment: AlignmentType.CENTER, bold: true, fill }), dataCell(String(count), 1200, { alignment: AlignmentType.CENTER, fill }),
      dataCell(pct, 1500, { alignment: AlignmentType.CENTER, fill }),
      new TableCell({ borders, width: { size: 5126, type: WidthType.DXA }, margins: cellMargins,
        shading: fill !== WHITE ? { fill, type: ShadingType.CLEAR } : undefined,
        children: [new Paragraph({ children: [
          new TextRun({ text: "\u2588".repeat(Math.max(1, Math.round(barWidth / 5))), size: 20, font: "Arial", color: barColor }),
          new TextRun({ text: ` ${count}件`, size: 18, font: "Arial", color: "666666" }),
        ]})] }),
    ]});
  });
}
function buildStrengthRows(data) {
  return data.map((row, index) => {
    const [theme, mentionCount, comments] = row;
    const fill = index % 2 === 0 ? "E8F5E9" : undefined;
    return new TableRow({ children: [
      dataCell(theme, 2600, { bold: true, fill }), dataCell(mentionCount, 1200, { alignment: AlignmentType.CENTER, bold: true, color: GREEN_ACCENT, fill }),
      dataCell(comments, 5226, { fill }),
    ]});
  });
}
function buildKPITargetRows(data) {
  return data.map((row, index) => {
    const [item, current, target, deadline] = row;
    const fill = index % 2 === 1 ? LIGHT_GRAY : undefined;
    return new TableRow({ children: [
      dataCell(item, 2800, { bold: true, fill }), dataCell(current, 2200, { alignment: AlignmentType.CENTER, fill }),
      dataCell(target, 2200, { alignment: AlignmentType.CENTER, color: GREEN_ACCENT, bold: true, fill }),
      dataCell(deadline, 1826, { alignment: AlignmentType.CENTER, fill }),
    ]});
  });
}

// ============================================================
// Build the document
// ============================================================
const doc = new Document({
  styles: {
    default: { document: { run: { font: "Arial", size: 21 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 32, bold: true, font: "Arial", color: NAVY }, paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 26, bold: true, font: "Arial", color: ACCENT }, paragraph: { spacing: { before: 280, after: 160 }, outlineLevel: 1 } },
    ],
  },
  numbering: { config: [
    { reference: "bullets", levels: [
      { level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
      { level: 1, format: LevelFormat.BULLET, text: "\u25E6", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 1440, hanging: 360 } } } },
    ]},
    { reference: "numbers", levels: [
      { level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
    ]},
  ]},
  sections: [
    // ===== COVER PAGE =====
    { properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
      children: [
        spacer(2000),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: ACCENT, space: 8 } }, children: [new TextRun({ text: "口コミ分析", size: 56, font: "Arial", color: ACCENT })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 }, children: [new TextRun({ text: "改善レポート", size: 56, font: "Arial", bold: true, color: NAVY })] }),
        spacer(400),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 120 }, children: [new TextRun({ text: HOTEL_NAME, size: 32, font: "Arial", color: NAVY, bold: true })] }),
        spacer(200),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [new TextRun({ text: `分析対象期間：${ANALYSIS_PERIOD}`, size: 22, font: "Arial", color: "666666" })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [new TextRun({ text: `レビュー総数：${REVIEW_COUNT}`, size: 22, font: "Arial", color: "666666" })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [new TextRun({ text: TARGET_SITES, size: 20, font: "Arial", color: "666666" })] }),
        spacer(1600),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: `作成日：${REPORT_DATE}`, size: 20, font: "Arial", color: "999999" })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Confidential - For Internal Use Only", size: 18, font: "Arial", color: "AAAAAA", italics: true })] }),
      ],
    },
    // ===== MAIN CONTENT =====
    { properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
      headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, border: { bottom: { style: BorderStyle.SINGLE, size: 2, color: ACCENT, space: 4 } }, children: [new TextRun({ text: `${HOTEL_NAME}｜口コミ分析改善レポート`, size: 16, font: "Arial", color: "999999" })] })] }) },
      footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Page ", size: 16, font: "Arial", color: "999999" }), new TextRun({ children: [PageNumber.CURRENT], size: 16, font: "Arial", color: "999999" })] })] }) },
      children: [
        heading1("1. エグゼクティブサマリー"), para(EXECUTIVE_SUMMARY_INTRO), spacer(100),
        kpiRow(KPI_CARDS), spacer(200),
        heading2("総合評価"), para(EXECUTIVE_SUMMARY_EVALUATION_1), para(EXECUTIVE_SUMMARY_EVALUATION_2), spacer(100),
        new Table({ width: { size: 9026, type: WidthType.DXA }, columnWidths: [9026],
          rows: [new TableRow({ children: [new TableCell({
            borders: { top: { style: BorderStyle.SINGLE, size: 4, color: ACCENT }, bottom: border, left: { style: BorderStyle.SINGLE, size: 4, color: ACCENT }, right: border },
            width: { size: 9026, type: WidthType.DXA }, shading: { fill: "F0F7FC", type: ShadingType.CLEAR }, margins: { top: 200, bottom: 200, left: 300, right: 300 },
            children: [
              new Paragraph({ spacing: { after: 120 }, children: [new TextRun({ text: "KEY FINDINGS", bold: true, size: 22, font: "Arial", color: ACCENT })] }),
              new Paragraph({ spacing: { after: 80 }, children: [new TextRun({ text: "Strength：", bold: true, size: 20, font: "Arial", color: GREEN_ACCENT }), new TextRun({ text: KEY_FINDING_STRENGTH, size: 20, font: "Arial", color: "333333" })] }),
              new Paragraph({ spacing: { after: 80 }, children: [new TextRun({ text: "Weakness：", bold: true, size: 20, font: "Arial", color: RED_ACCENT }), new TextRun({ text: KEY_FINDING_WEAKNESS, size: 20, font: "Arial", color: "333333" })] }),
              new Paragraph({ children: [new TextRun({ text: "Opportunity：", bold: true, size: 20, font: "Arial", color: ORANGE_ACCENT }), new TextRun({ text: KEY_FINDING_OPPORTUNITY, size: 20, font: "Arial", color: "333333" })] }),
            ] })] })] }),
        new Paragraph({ children: [new PageBreak()] }),

        heading1("2. データ概要"),
        heading2("2.1 サイト別レビュー件数・評価"), para(DATA_OVERVIEW_SITE_TEXT), spacer(80),
        new Table({ width: { size: 9026, type: WidthType.DXA }, columnWidths: [1800, 900, 1300, 900, 1300, 1526, 1300],
          rows: [new TableRow({ children: [headerCell("サイト名", 1800), headerCell("件数", 900), headerCell("ネイティブ平均", 1300), headerCell("尺度", 900), headerCell("10pt換算", 1300), headerCell("中央値(10pt)", 1526), headerCell("判定", 1300)] }), ...buildSiteTableRows(SITE_DATA)] }),
        spacer(200),
        heading2("2.2 評価分布（10点換算）"), para(DATA_OVERVIEW_DIST_TEXT), spacer(80),
        new Table({ width: { size: 9026, type: WidthType.DXA }, columnWidths: [1200, 1200, 1500, 5126],
          rows: [new TableRow({ children: [headerCell("評価", 1200), headerCell("件数", 1200), headerCell("割合", 1500), headerCell("分布", 5126)] }), ...buildDistributionRows(DISTRIBUTION_DATA, DISTRIBUTION_MAX_COUNT)] }),
        spacer(100),
        new Table({ width: { size: 9026, type: WidthType.DXA }, columnWidths: [3008, 3009, 3009],
          rows: [new TableRow({ children: [
            new TableCell({ borders, width: { size: 3008, type: WidthType.DXA }, shading: { fill: "E8F5E9", type: ShadingType.CLEAR }, margins: { top: 120, bottom: 120, left: 200, right: 200 },
              children: [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [new TextRun({ text: "高評価（8-10点）", bold: true, size: 20, font: "Arial", color: GREEN_ACCENT })] }), new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: HIGH_RATING_SUMMARY, bold: true, size: 24, font: "Arial", color: GREEN_ACCENT })] })] }),
            new TableCell({ borders, width: { size: 3009, type: WidthType.DXA }, shading: { fill: "FFF3E0", type: ShadingType.CLEAR }, margins: { top: 120, bottom: 120, left: 200, right: 200 },
              children: [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [new TextRun({ text: "中評価（5-7点）", bold: true, size: 20, font: "Arial", color: ORANGE_ACCENT })] }), new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: MID_RATING_SUMMARY, bold: true, size: 24, font: "Arial", color: ORANGE_ACCENT })] })] }),
            new TableCell({ borders, width: { size: 3009, type: WidthType.DXA }, shading: { fill: "FDEDEC", type: ShadingType.CLEAR }, margins: { top: 120, bottom: 120, left: 200, right: 200 },
              children: [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [new TextRun({ text: "低評価（1-4点）", bold: true, size: 20, font: "Arial", color: LOW_RATING_COLOR })] }), new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: LOW_RATING_SUMMARY, bold: true, size: 24, font: "Arial", color: LOW_RATING_COLOR })] })] }),
          ] })] }),
        new Paragraph({ children: [new PageBreak()] }),

        heading1("3. 強み分析（ポジティブ要因）"), para("口コミのテキストマイニングにより、以下のポジティブテーマが特定されました。"), spacer(100),
        new Table({ width: { size: 9026, type: WidthType.DXA }, columnWidths: [2600, 1200, 5226],
          rows: [new TableRow({ children: [headerCell("ポジティブテーマ", 2600), headerCell("言及数", 1200), headerCell("代表的なコメント", 5226)] }), ...buildStrengthRows(STRENGTH_THEMES)] }),
        spacer(200),
        heading2(STRENGTH_SUB_1_TITLE), para(STRENGTH_SUB_1_TEXT), ...STRENGTH_SUB_1_BULLETS.map(b => bulletItem(b)), spacer(100),
        heading2(STRENGTH_SUB_2_TITLE), para(STRENGTH_SUB_2_TEXT),
        new Paragraph({ children: [new PageBreak()] }),

        heading1("4. 弱み分析（改善課題）"), para("ネガティブコメントの分析から、以下の改善課題が抽出されました。影響度と頻度に基づく優先度をS〜Cで設定しています。"), spacer(100),
        new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [800, 2200, 4360, 2000],
          rows: [new TableRow({ children: [headerCell("優先度", 800), headerCell("課題カテゴリ", 2200), headerCell("具体的内容", 4360), headerCell("影響度", 2000)] }), ...WEAKNESS_PRIORITY_DATA.map(([p, c, d, i]) => priorityRow(p, c, d, i))] }),
        new Paragraph({ children: [new PageBreak()] }),

        heading1("5. 改善施策提案"), para("分析結果に基づき、以下の改善施策を「即座対応」「短期」「中期」の3フェーズに分けて提案いたします。"), spacer(100),
        heading2("Phase 1：即座対応（今週〜1ヶ月以内）"), para(PHASE1_DESCRIPTION), spacer(80), ...buildPhaseItems(PHASE1_ITEMS),
        new Paragraph({ children: [new PageBreak()] }),
        heading2("Phase 2：短期施策（1〜3ヶ月）"), para(PHASE2_DESCRIPTION), spacer(80), ...buildPhaseItems(PHASE2_ITEMS),
        spacer(200),
        heading2("Phase 3：中期施策（3〜6ヶ月）"), para(PHASE3_DESCRIPTION), spacer(80), ...buildPhaseItems(PHASE3_ITEMS),
        new Paragraph({ children: [new PageBreak()] }),

        heading1("6. KPI目標設定"), para("以下のKPIを設定し、四半期ごとにモニタリングすることを推奨します。"), spacer(100),
        new Table({ width: { size: 9026, type: WidthType.DXA }, columnWidths: [2800, 2200, 2200, 1826],
          rows: [new TableRow({ children: [headerCell("KPI項目", 2800), headerCell("現状値", 2200), headerCell("目標値（6ヶ月後）", 2200), headerCell("期限", 1826)] }), ...buildKPITargetRows(KPI_TARGET_DATA)] }),
        new Paragraph({ children: [new PageBreak()] }),

        heading1("7. 総括と今後のアクション"), spacer(100),
        new Table({ width: { size: 9026, type: WidthType.DXA }, columnWidths: [9026],
          rows: [new TableRow({ children: [new TableCell({
            borders: { top: { style: BorderStyle.SINGLE, size: 4, color: NAVY }, bottom: { style: BorderStyle.SINGLE, size: 4, color: NAVY }, left: { style: BorderStyle.SINGLE, size: 4, color: NAVY }, right: { style: BorderStyle.SINGLE, size: 4, color: NAVY } },
            width: { size: 9026, type: WidthType.DXA }, shading: { fill: "F8F9FA", type: ShadingType.CLEAR }, margins: { top: 300, bottom: 300, left: 400, right: 400 },
            children: [
              ...CONCLUSION_PARAGRAPHS.map(text => new Paragraph({ spacing: { after: 160 }, children: [new TextRun({ text, size: 21, font: "Arial", color: "333333" })] })),
              new Paragraph({ children: [new TextRun({ text: CONCLUSION_FINAL_PARAGRAPH, size: 21, font: "Arial", color: NAVY, bold: true })] }),
            ] })] })] }),
        spacer(300), divider(), spacer(100),
        para("本レポートに関するご質問・ご相談がございましたら、お気軽にお問い合わせください。", { alignment: AlignmentType.CENTER, color: "999999" }),
      ],
    },
  ],
});

Packer.toBuffer(doc).then(buffer => {
  const outputPath = `${OUTPUT_DIR}/${HOTEL_NAME}_口コミ分析改善レポート.docx`;
  fs.writeFileSync(outputPath, buffer);
  console.log("Report created successfully!");
  console.log("Output: " + outputPath);
  console.log("File size: " + (buffer.length / 1024).toFixed(1) + " KB");
});
