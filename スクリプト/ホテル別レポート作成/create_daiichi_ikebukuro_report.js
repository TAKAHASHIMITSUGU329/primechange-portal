const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, PageNumber, PageBreak, LevelFormat
} = require("docx");

// ============================================================
// ホテル固有の設定
// ============================================================
const HOTEL_NAME = "第一イン池袋";
const ANALYSIS_PERIOD = "2026年2月〜3月";
const REPORT_DATE = "2026年3月8日";
const OUTPUT_DIR = "/Users/mitsugutakahashi/ホテル口コミ";

const REVIEW_COUNT = "74件（重複除外後）";
const TARGET_SITES = "対象サイト：Booking.com / Trip.com / じゃらん / 楽天トラベル / Agoda / Google";

// KPIカードの値
const KPI_CARDS = [
  { label: "全体平均(10pt換算)", value: "9.03", color: "27AE60", bgColor: "E8F5E9" },
  { label: "高評価率(8-10点)", value: "91.9%", color: "27AE60", bgColor: "E8F5E9" },
  { label: "低評価率(1-4点)", value: "1.4%", color: "27AE60", bgColor: "E8F5E9" },
  { label: "レビュー総数", value: "74件", color: "1B3A5C", bgColor: "D5E8F0" },
];

// エグゼクティブサマリーのテキスト
const EXECUTIVE_SUMMARY_INTRO = "2026年2月〜3月に各OTAサイト・口コミサイトに投稿された74件のレビューを包括的に分析しました。以下が主要な発見事項です。";
const EXECUTIVE_SUMMARY_EVALUATION_1 = "全体平均スコア9.03点（10点換算）は非常に高い水準であり、91.9%が高評価（8-10点）に分類されます。低評価（1-4点）はわずか1件（1.4%）で、ゲスト満足度は極めて良好です。Trip.com（9.44点）とGoogle（9.2点）で特に高い評価を獲得しており、海外・国内の両方のゲストから信頼されているホテルと言えます。";
const EXECUTIVE_SUMMARY_EVALUATION_2 = "一方で、Agoda（6.0点）では相対的に低いスコアとなっており、設備の老朽化やアメニティ関連の不満がスコアに影響しています。全体としては「立地の圧倒的な利便性」と「スタッフの丁寧な対応」がホテルの最大の強みとして際立っています。";

// KEY FINDINGS
const KEY_FINDING_STRENGTH = "駅近の立地（50件以上で言及）、1Fコンビニの利便性（20件以上）、スタッフの丁寧な対応（15件以上）";
const KEY_FINDING_WEAKNESS = "防音性能（4件）、設備の老朽化（5件）、浴室・水回りの不便さ（4件）、朝食の改善要望（3件）";
const KEY_FINDING_OPPORTUNITY = "リピート意向のコメントが多数あり、設備面の小改善で高評価率をさらに安定させるポテンシャルが大きい";

// サイト別レビューデータ
const SITE_DATA = [
  ["Trip.com", "36", "9.44", "/10", "9.44", "10.0", "優秀"],
  ["Google", "20", "4.6", "/5 (×2)", "9.2", "10.0", "優秀"],
  ["Booking.com", "11", "8.55", "/10", "8.55", "9.0", "良好"],
  ["じゃらん", "1", "4.0", "/5 (×2)", "8.0", "8.0", "良好"],
  ["楽天トラベル", "3", "4.0", "/5 (×2)", "8.0", "8.0", "良好"],
  ["Agoda", "3", "6.0", "/10", "6.0", "5.0", "要改善"],
];

// 評価分布データ
const DISTRIBUTION_DATA = [
  [10, 41, "55.4%"],
  [9,  12, "16.2%"],
  [8,  15, "20.3%"],
  [7,  1,  "1.4%"],
  [6,  1,  "1.4%"],
  [5,  3,  "4.1%"],
  [2,  1,  "1.4%"],
];
const DISTRIBUTION_MAX_COUNT = 41;

// 評価分布サマリー
const HIGH_RATING_SUMMARY = "68件（91.9%）";
const MID_RATING_SUMMARY = "5件（6.8%）";
const LOW_RATING_SUMMARY = "1件（1.4%）";
const LOW_RATING_COLOR = "27AE60"; // 低評価1件なので緑

// データ概要テキスト
const DATA_OVERVIEW_SITE_TEXT = "各予約サイトの評価基準が異なる点に留意が必要です。海外OTA（Booking.com/Trip.com/Agoda）は10点満点、国内サイト（じゃらん/楽天トラベル）およびGoogleは5点満点のため、統一比較のため10点換算（5点満点×2）を行いました。";
const DATA_OVERVIEW_DIST_TEXT = "評価分布は10点（55.4%）に大きく偏っており、ゲストの過半数が最高評価を付与しています。8点以上の高評価が91.9%を占め、極めて高い顧客満足度を示しています。5点以下の低スコアは計4件（5.5%）で、主にAgodaでの投稿に集中しています。";

// 強みテーマデータ
const STRENGTH_THEMES = [
  ["駅近・立地の利便性", "50件以上", "「池袋駅東口から徒歩2分」「駅にとても近く、旅行に便利」"],
  ["1Fコンビニ（ファミマ）", "20件以上", "「ロビーにコンビニがあるので便利」「階下にファミリーマートもあります」"],
  ["スタッフの対応・ホスピタリティ", "15件以上", "「フロントの対応がすごく良かった」「全従業員にもてなす気持ちがある」"],
  ["清潔さ・毎日清掃", "15件以上", "「毎日清掃も行き届いている」「とても清潔でした」"],
  ["部屋の広さ（ビジネスホテル比）", "10件以上", "「ビジネスホテルに比べると部屋はかなり広い」「部屋もかなり広い」"],
  ["コストパフォーマンス", "8件以上", "「コスパが最高です」「コストパフォーマンスに優れています」"],
];

// 強みサブセクション
const STRENGTH_SUB_1_TITLE = "3.1 最大の強み：「駅近・圧倒的な立地の利便性」";
const STRENGTH_SUB_1_TEXT = "口コミの大多数で「駅から近い」「立地が最高」と言及されており、池袋駅東口から徒歩2〜5分というアクセスの良さが最大の差別化要因です。さらに、1階にファミリーマートと一蘭ラーメンが直結しており、ゲストの利便性を大幅に高めています。";
const STRENGTH_SUB_1_BULLETS = [
  "池袋駅東口（北口）から徒歩2〜5分の抜群のアクセス",
  "1F直結のファミリーマートで24時間の買い物が可能",
  "周辺にドン・キホーテ、ビックカメラ、西武百貨店、飲食店街が集積",
  "JR山手線で新宿・渋谷へ10分圏内、東京観光の拠点として最適",
];

const STRENGTH_SUB_2_TITLE = "3.2 高いホスピタリティとサービス品質";
const STRENGTH_SUB_2_TEXT = "スタッフの親切さ、丁寧な対応、多言語対応（韓国語・英語）が多くのゲストから高く評価されています。レイトチェックアウトへの柔軟な対応、タバコ臭除去の要望への継続的な清掃対応、蛇口交換時のメモによる通知など、きめ細やかなサービスが印象に残っているゲストが多数います。セルフサービスの荷物預かりシステムも好評です。";

// 弱み分析データ（優先度マトリクス）
const WEAKNESS_PRIORITY_DATA = [
  ["S", "設備の老朽化", "建物・内装の古さが複数のゲストから指摘。清潔感はあるが「古いホテルの感じ」との声", "印象低下・5件"],
  ["A", "防音性能", "電車の音が聞こえる、隣室の音が漏れるとの指摘。防音レベルの改善が望まれる", "快適性低下・4件"],
  ["A", "浴室・水回り", "ユニットバスの狭さ、排水の悪さ、シャワー使用時の異音、浴室の寒さ", "快適性低下・4件"],
  ["B", "朝食の充実度", "メニューが少ない、器が小さくすぐなくなる、やや高価との声", "満足度低下・3件"],
  ["B", "枕の品質", "柔らかすぎて支えがない、1つでは低すぎ2つでは高すぎる", "睡眠品質・2件"],
  ["C", "アメニティ配置", "一部アメニティがロビーまで取りに行く必要あり。エアコン設定の不備", "利便性低下・2件"],
];

// 改善施策 Phase 1（即座対応）
const PHASE1_DESCRIPTION = "投資不要・オペレーション改善で対応可能な施策";
const PHASE1_ITEMS = [
  {
    title: "(1) 枕の選択肢拡充",
    bullets: [
      "硬めの枕を各フロアのリネン室に常備し、フロントで案内",
      "チェックイン時に「枕の硬さの好みはございますか？」と声掛け",
      "枕メニューカードを客室内に設置し、内線で交換依頼可能に",
    ],
  },
  {
    title: "(2) アメニティの客室内充実",
    bullets: [
      "現在ロビー設置のアメニティを客室内にも基本セットとして配置",
      "客室内アメニティリスト・案内カードの設置",
      "不足分はフロント内線で追加提供可能であることを明示",
    ],
  },
  {
    title: "(3) 朝食サービスの改善",
    bullets: [
      "メニューの補充頻度を上げ、品切れを防止",
      "器のサイズ見直し、適量を維持できる運用に変更",
      "宿泊客数に応じた朝食準備量の最適化",
    ],
  },
  {
    title: "(4) エアコン・設備の定期点検強化",
    bullets: [
      "冬季の暖房設定を全室デフォルトで確認する運用フロー導入",
      "冷蔵庫の冷却性能を定期チェックリストに追加",
      "シャワー異音の原因特定と該当室の優先修繕",
    ],
  },
];

// 改善施策 Phase 2（短期）
const PHASE2_DESCRIPTION = "一定の投資を伴うが、比較的早期に実行可能な施策";
const PHASE2_ITEMS = [
  {
    title: "(1) 浴室の快適性向上",
    bullets: [
      "排水口の清掃・修繕を全室で実施し、排水速度を改善",
      "シャワーヘッドの交換（異音対策・水圧改善）",
      "浴室暖房の導入検討（小型ヒーター設置）",
    ],
  },
  {
    title: "(2) 防音対策の強化",
    bullets: [
      "窓際への防音カーテン導入（電車騒音対策）",
      "ドア下部への隙間テープ設置（廊下・隣室からの音漏れ軽減）",
      "低層階ゲストへの高層階優先案内の仕組み化",
    ],
  },
  {
    title: "(3) 口コミ返信体制の構築",
    bullets: [
      "全サイトの口コミに48時間以内の返信を目標に設定",
      "特にAgodaの低評価レビューへのフォローアップ強化",
    ],
  },
];

// 改善施策 Phase 3（中期）
const PHASE3_DESCRIPTION = "設備投資を伴う抜本的な改善施策";
const PHASE3_ITEMS = [
  {
    title: "(1) 客室リニューアル計画",
    bullets: [
      "壁紙・カーペットの張替えによる「古さ」印象の払拭",
      "照明のLED化・調光機能の追加で客室の雰囲気向上",
      "USB充電ポート付きコンセントの設置",
    ],
  },
  {
    title: "(2) 浴室の抜本的改修",
    bullets: [
      "ユニットバスの配管更新による排水・異音問題の根本解決",
      "換気設備の強化による浴室内の温度・湿度改善",
    ],
  },
  {
    title: "(3) 防音工事",
    bullets: [
      "線路側客室の窓を二重サッシに交換",
      "隣室との壁面に遮音材を追加する改修工事の検討",
    ],
  },
];

// KPI目標データ
const KPI_TARGET_DATA = [
  ["全体平均(10pt換算)", "9.03点", "9.2点以上", "2026年9月"],
  ["高評価率（8-10点）", "91.9%", "93%以上", "2026年9月"],
  ["低評価率（1-4点）", "1.4%", "0%維持", "2026年9月"],
  ["Agoda平均評価", "6.0点", "7.5点以上", "2026年9月"],
  ["口コミ返信率", "未計測", "100%（48h以内）", "2026年6月"],
  ["設備クレーム件数", "約5件/2ヶ月", "2件以下/2ヶ月", "2026年9月"],
];

// 総括テキスト
const CONCLUSION_PARAGRAPHS = [
  "第一イン池袋は、全体平均9.03点（10点換算）、高評価率91.9%という非常に高い顧客満足度を維持しているホテルです。特に「池袋駅から徒歩2〜5分」という圧倒的な立地の利便性、1階直結のファミリーマートと一蘭ラーメン、そしてスタッフの丁寧で心のこもったサービスが、多くのゲストに深い印象を残しています。",
  "改善課題としては、建物・設備の老朽化感、防音性能、浴室の快適性が主要なテーマとして浮上しています。ただし、これらの課題を指摘するゲストの多くも「清潔さ」と「立地」を高く評価しており、ネガティブ要素が全体評価を大きく押し下げているわけではありません。",
  "Phase 1（即座対応）の枕の選択肢拡充やアメニティ配置改善は、低コストで即座に実行可能であり、ゲスト体験の底上げに直結します。Phase 2（短期）の防音対策や浴室改善は、中評価ゲストを高評価に引き上げる効果が期待できます。Phase 3（中期）の客室リニューアルは、「古さ」の印象を払拭し、ブランド価値の向上に寄与します。",
  "Agodaでの低スコア（6.0点）は、口コミ返信の強化とサービス品質の一貫性確保で改善余地があります。全サイトでの口コミ返信率100%を目指すことで、ホテルの姿勢をゲストに伝え、リピート率の向上につなげることが重要です。",
];
const CONCLUSION_FINAL_PARAGRAPH = "立地という最大の強みを活かしつつ、設備面の段階的な改善を実行することで、第一イン池袋は池袋エリアにおけるコストパフォーマンス最高のホテルとしての地位をさらに強固にすることができます。";

// ============================================================
// Color scheme (DO NOT MODIFY)
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
// Helper functions (DO NOT MODIFY)
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

// KPI Box helper
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

// Priority matrix helper
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

// ============================================================
// Build improvement phase sections
// ============================================================
function buildPhaseItems(items) {
  const result = [];
  items.forEach((item, index) => {
    result.push(heading3(item.title));
    item.bullets.forEach(b => result.push(bulletItem(b)));
    if (index < items.length - 1) result.push(spacer(80));
  });
  return result;
}

// ============================================================
// Build site table rows from data array
// ============================================================
function buildSiteTableRows(data) {
  return data.map((row, index) => {
    const [siteName, count, nativeAvg, scale, tenPt, median, verdict] = row;
    const fill = index % 2 === 1 ? LIGHT_GRAY : undefined;
    const verdictColor = verdict === "優秀" ? GREEN_ACCENT : (verdict === "良好" ? GREEN_ACCENT : ORANGE_ACCENT);
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

// ============================================================
// Build distribution table rows from data array
// ============================================================
function buildDistributionRows(data, maxCount) {
  return data.map(([rating, count, pct]) => {
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

// ============================================================
// Build strength theme rows from data array
// ============================================================
function buildStrengthRows(data) {
  return data.map((row, index) => {
    const [theme, mentionCount, comments] = row;
    const fill = index % 2 === 0 ? "E8F5E9" : undefined;
    return new TableRow({ children: [
      dataCell(theme, 2600, { bold: true, fill: fill || undefined }),
      dataCell(mentionCount, 1200, { alignment: AlignmentType.CENTER, bold: true, color: GREEN_ACCENT, fill: fill || undefined }),
      dataCell(comments, 5226, { fill: fill || undefined }),
    ]});
  });
}

// ============================================================
// Build KPI target rows from data array
// ============================================================
function buildKPITargetRows(data) {
  return data.map((row, index) => {
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
// Build the document
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
        para(EXECUTIVE_SUMMARY_INTRO),
        spacer(100),

        kpiRow(KPI_CARDS),
        spacer(200),

        heading2("総合評価"),
        para(EXECUTIVE_SUMMARY_EVALUATION_1),
        para(EXECUTIVE_SUMMARY_EVALUATION_2),
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

        heading2(STRENGTH_SUB_1_TITLE),
        para(STRENGTH_SUB_1_TEXT),
        ...STRENGTH_SUB_1_BULLETS.map(b => bulletItem(b)),
        spacer(100),

        heading2(STRENGTH_SUB_2_TITLE),
        para(STRENGTH_SUB_2_TEXT),

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
        para(PHASE1_DESCRIPTION),
        spacer(80),
        ...buildPhaseItems(PHASE1_ITEMS),

        new Paragraph({ children: [new PageBreak()] }),

        // Phase 2
        heading2("Phase 2：短期施策（1〜3ヶ月）"),
        para(PHASE2_DESCRIPTION),
        spacer(80),
        ...buildPhaseItems(PHASE2_ITEMS),

        spacer(200),

        // Phase 3
        heading2("Phase 3：中期施策（3〜6ヶ月）"),
        para(PHASE3_DESCRIPTION),
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

        // ===== 7. CONCLUSION =====
        heading1("7. 総括と今後のアクション"),
        spacer(100),

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
                new Paragraph({ children: [new TextRun({ text: CONCLUSION_FINAL_PARAGRAPH, size: 21, font: "Arial", color: NAVY, bold: true })] }),
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

Packer.toBuffer(doc).then(buffer => {
  const outputPath = `${OUTPUT_DIR}/${HOTEL_NAME}_口コミ分析改善レポート.docx`;
  fs.writeFileSync(outputPath, buffer);
  console.log("Report created successfully!");
  console.log("Output: " + outputPath);
  console.log("File size: " + (buffer.length / 1024).toFixed(1) + " KB");
});
