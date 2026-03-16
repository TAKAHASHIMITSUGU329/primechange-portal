const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, PageNumber, PageBreak, LevelFormat
} = require("docx");

// ============================================================
// ホテル固有の設定
// ============================================================
const HOTEL_NAME = "ハートンホテル東品川";
const ANALYSIS_PERIOD = "2026年2月〜3月";
const REPORT_DATE = "2026年3月7日";
const OUTPUT_DIR = "/Users/mitsugutakahashi/ホテル口コミ";

const REVIEW_COUNT = "66件（重複除外後）";
const TARGET_SITES = "対象サイト：Booking.com / Trip.com / じゃらん / 楽天トラベル / Agoda / Google";

const KPI_CARDS = [
  { label: "全体平均(10pt換算)", value: "8.74", color: "27AE60", bgColor: "E8F5E9" },
  { label: "高評価率(8-10点)", value: "84.8%", color: "27AE60", bgColor: "E8F5E9" },
  { label: "低評価率(1-4点)", value: "4.5%", color: "E67E22", bgColor: "D5E8F0" },
  { label: "レビュー総数", value: "66件", color: "1B3A5C", bgColor: "D5E8F0" },
];

const EXECUTIVE_SUMMARY_INTRO = "2026年2月〜3月に各OTAサイト・口コミサイトに投稿された66件のレビューを包括的に分析しました。以下が主要な発見事項です。";
const EXECUTIVE_SUMMARY_EVALUATION_1 = "全体平均は10点換算で8.74点と「良好」水準にあり、84.8%のゲストが8点以上の高評価を付けています。特にTrip.comでは満点10.0点を記録し、楽天トラベル・Agoda・じゃらん・Booking.comでもいずれも8点台後半と安定した高評価を得ています。";
const EXECUTIVE_SUMMARY_EVALUATION_2 = "一方、Google評価が6.0点と他サイトに比べ大きく下回っており、スタッフ対応の一貫性や清掃品質に関する改善余地が見られます。低評価（1-4点）は4.5%（3件）発生しており、清掃不備やスタッフ対応、眺望に関する不満が確認されました。";

const KEY_FINDING_STRENGTH = "駅直結の立地（40件以上で言及）、朝食の高品質（25件以上）、清潔で新しい施設（20件以上）が最大の強み";
const KEY_FINDING_WEAKNESS = "清掃品質のばらつき（4件）、スタッフ対応の一貫性不足（3件）、眺望なし客室の事前告知不足（1件）";
const KEY_FINDING_OPPORTUNITY = "リピート意向が非常に高く（20件以上で「また利用したい」と明記）、朝食体験の訴求強化と清掃品質の安定化で高評価率90%超えを目指せるポテンシャル大";

const SITE_DATA = [
  ["Trip.com", "4", "10.00", "/10", "10.00", "10.0", "優秀"],
  ["楽天トラベル", "21", "4.43", "/5", "8.86", "10.0", "良好"],
  ["Agoda", "12", "8.83", "/10", "8.83", "9.5", "良好"],
  ["じゃらん", "16", "4.38", "/5", "8.75", "9.0", "良好"],
  ["Booking.com", "10", "8.70", "/10", "8.70", "9.5", "良好"],
  ["Google", "3", "3.00", "/5", "6.00", "6.0", "要改善"],
];

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
const LOW_RATING_COLOR = "E67E22";

const DATA_OVERVIEW_SITE_TEXT = "各予約サイトの評価基準が異なる点に留意が必要です。海外OTA（Booking.com/Trip.com/Agoda）は10点満点、国内サイト（じゃらん/楽天トラベル/Google）は5点満点のため、統一比較のため10点換算（5点満点×2）を行いました。";
const DATA_OVERVIEW_DIST_TEXT = "10点（満点）が全体の57.6%を占め、8点以上の高評価が全体の84.8%に達しています。一方、6点帯が10.6%、4点以下が4.5%存在し、特にGoogle評価の低さが全体平均を押し下げる要因となっています。";

const STRENGTH_THEMES = [
  ["立地・アクセス", "42件", "「駅直結で雨にも濡れず便利」「コンビニ・飲食店が近くて快適」"],
  ["朝食・食事", "28件", "「朝食の質が高級ホテル並み」「カレーやお茶漬けが特に美味しい」"],
  ["清潔感・新しさ", "22件", "「新しくて清潔」「施設がスタイリッシュで綺麗」"],
  ["スタッフ対応", "18件", "「フロントの対応が丁寧」「きめ細やかなサービスが良い」"],
  ["部屋・設備", "15件", "「部屋が広めでゆっくりできる」「ベッドが快適」「アメニティが充実」"],
  ["コスパ・リピート", "20件", "「また利用したい」「値段の割に質が高い」「コスパが良い」"],
];

const STRENGTH_SUB_1_TITLE = "3.1 最大の強み：「駅直結の圧倒的な立地利便性」";
const STRENGTH_SUB_1_TEXT = "口コミの約64%で立地の利便性が言及されており、駅から直結・徒歩1分というアクセスの良さが最大の差別化要因となっています。雨天時でも濡れずにアクセスできる点、周辺の飲食店やコンビニ、ショッピング施設の充実度も高く評価されています。";
const STRENGTH_SUB_1_BULLETS = [
  "駅直結のペデストリアンデッキにより天候を問わない快適なアクセスが可能",
  "周辺にコンビニ・ドラッグストア・飲食店・スーパーが充実し生活利便性が高い",
  "ビジネス利用（出張・展示会）と観光利用の双方に対応可能な立地",
  "リピーター獲得の最大要因として機能し、「次回も利用したい」理由の第1位",
];

const STRENGTH_SUB_2_TITLE = "3.2 高品質な朝食が満足度を大きく押し上げ";
const STRENGTH_SUB_2_TEXT = "朝食に関する言及が28件に上り、品数は多くないものの一品一品の質の高さが宿泊体験全体の評価を大きく引き上げています。特にカレー、お茶漬け、和洋ビュッフェのバランスが好評で、「高級ホテルを凌ぐクオリティ」という声も複数見られました。";

const WEAKNESS_PRIORITY_DATA = [
  ["S", "清掃品質のばらつき", "シーツの毛・ゴミ残り、冷蔵庫内の髪の毛、風呂場の毛髪、置き時計の時刻ずれなど基本的な清掃不備が複数報告", "直接的な低評価要因・4件"],
  ["A", "スタッフ対応の一貫性", "一部スタッフの無愛想な対応、連泊時の情報伝達不足（清掃スタッフへの延泊連絡漏れ）、アメニティ補充漏れ", "サービス品質への信頼低下・3件"],
  ["A", "眺望なし客室の告知不足", "すりガラス窓の客室について事前の情報提示がなく、入室時のガッカリ感が低評価の直接原因に", "期待値コントロール・1件"],
  ["B", "朝食の混雑・時間帯", "7時オープンでは遅い、混雑時のストレスという指摘。6時30分開始を希望する声あり", "改善でさらなる満足度向上・3件"],
  ["B", "空調・室温管理", "冬季にエアコンが効きにくく室内が暖まらない、シャワー中も寒いという報告", "季節的要因だが快適性に影響・2件"],
  ["C", "駐車場の不在", "ホテル専用駐車場がなく、提携駐車場の確約もないため車利用者には不便", "車利用ゲスト限定の課題・2件"],
  ["C", "大浴場の不在", "大浴場がないことへの要望。系列ホテルには大浴場があるため比較される", "設備面の期待ギャップ・1件"],
];

const PHASE1_DESCRIPTION = "投資不要・オペレーション改善で対応可能な施策";
const PHASE1_ITEMS = [
  {
    title: "(1) 清掃品質チェック体制の強化",
    bullets: [
      "チェックリスト式の客室清掃完了検査シートを導入（シーツ・浴室・冷蔵庫・備品を重点確認）",
      "清掃担当者と検査担当者の分離（ダブルチェック体制）",
      "月次で清掃関連クレームを集計し、傾向分析を実施",
    ],
  },
  {
    title: "(2) スタッフ間の情報共有改善",
    bullets: [
      "連泊・延泊情報をPMS（ホテル管理システム）から清掃チームに自動通知する仕組みを整備",
      "アメニティ補充リクエストの記録・完了確認フローを標準化",
      "チェックイン時の「お出迎え」品質向上のためのロールプレイング研修を月1回実施",
    ],
  },
  {
    title: "(3) 眺望なし客室の事前告知",
    bullets: [
      "予約サイトの客室説明に「一部客室は眺望がございません」と明記",
      "該当客室の予約時にメールで事前説明を送付",
      "チェックイン時に口頭でも説明し、期待値とのギャップを解消",
    ],
  },
  {
    title: "(4) 朝食オペレーションの見直し",
    bullets: [
      "繁忙期の開始時間を6:30に前倒しする試験運用を開始",
      "混雑予測に基づくスタッフ増員体制の構築",
      "朝食メニューのバリエーション追加（ベーコンなど要望の多い品目）",
    ],
  },
];

const PHASE2_DESCRIPTION = "一定の投資を伴うが、比較的早期に実行可能な施策";
const PHASE2_ITEMS = [
  {
    title: "(1) Google口コミ評価の改善プログラム",
    bullets: [
      "チェックアウト時にQRコードでGoogle口コミ投稿を依頼（満足度の高いゲストに絞る）",
      "Google口コミへの24時間以内返信体制を構築",
      "低評価口コミに対する丁寧なフォローアップ回答のテンプレート作成",
    ],
  },
  {
    title: "(2) 接客品質の標準化・研修強化",
    bullets: [
      "「ホスピタリティ基準書」を策定し、全スタッフに共有",
      "外部講師によるホスピタリティ研修を四半期ごとに実施",
      "ゲストからの評価を個人フィードバックとして活用する仕組みの導入",
    ],
  },
  {
    title: "(3) 朝食の訴求強化",
    bullets: [
      "朝食の高評価コメントをOTA掲載写真・説明文に反映",
      "季節限定メニューの導入でリピーターの再訪動機を創出",
    ],
  },
];

const PHASE3_DESCRIPTION = "設備投資を伴う抜本的な改善施策";
const PHASE3_ITEMS = [
  {
    title: "(1) 空調設備の点検・改修",
    bullets: [
      "冬季の暖房効率が低い客室の特定と個別対策（断熱・設備更新）",
      "全客室の空調設備の定期メンテナンススケジュール策定",
      "補助暖房器具（ファンヒーター等）の貸出サービスの検討",
    ],
  },
  {
    title: "(2) 空気清浄機のフィルター更新プログラム",
    bullets: [
      "全客室の空気清浄機フィルター交換スケジュールの策定（3ヶ月ごと）",
      "匂いの出やすい機器は優先的にリプレースを実施",
    ],
  },
  {
    title: "(3) 近隣駐車場との連携強化",
    bullets: [
      "近隣コインパーキング2〜3箇所との提携契約を締結（宿泊者割引適用）",
      "予約時に駐車場情報を自動送信するシステム対応",
      "ホテル公式サイトに駐車場マップ・料金情報を掲載",
    ],
  },
];

const KPI_TARGET_DATA = [
  ["全体平均(10pt換算)", "8.74点", "9.0点以上", "2026年9月"],
  ["高評価率（8-10点）", "84.8%", "90%以上", "2026年9月"],
  ["低評価率（1-4点）", "4.5%", "2%以下", "2026年9月"],
  ["Google平均評価", "3.0/5点", "4.0/5点以上", "2026年9月"],
  ["じゃらん平均評価", "4.38/5点", "4.5/5点以上", "2026年9月"],
  ["口コミ返信率", "未計測", "100%（48h以内）", "2026年6月"],
  ["清掃クレーム件数", "4件/2ヶ月", "0件/2ヶ月", "2026年9月"],
];

const CONCLUSION_PARAGRAPHS = [
  "ハートンホテル東品川は、駅直結の優れた立地、高品質な朝食、清潔で新しい施設という3つの確かな強みを持ち、全体平均8.74点・高評価率84.8%と安定した評価基盤を構築しています。",
  "一方で、清掃品質のばらつき（4件のクレーム）、スタッフ対応の一貫性不足、Google評価の低迷（6.0点）という明確な改善ポイントも存在します。これらは比較的少ない投資で改善可能な運用上の課題であり、迅速な対応により大きな改善効果が期待できます。",
  "特に清掃品質の安定化とスタッフ研修の強化は、短期間で成果が見える施策です。また、リピート意向の高さ（20件以上で「また利用したい」と明記）は、既存の強みが確実に顧客ロイヤルティに繋がっていることを示しており、この良い循環をさらに強化することが重要です。",
  "提案した3フェーズの改善施策を段階的に実施することで、6ヶ月後には全体平均9.0点以上、高評価率90%超えという目標の達成が十分に可能と考えます。",
];
const CONCLUSION_FINAL_PARAGRAPH = "口コミ分析に基づく継続的な改善サイクル（PDCA）を回し、「選ばれ続けるホテル」としてのブランド価値をさらに高めてまいりましょう。";

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

function buildSiteTableRows(data) {
  return data.map((row, index) => {
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
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [new TextRun({ text: `分析対象期間：${ANALYSIS_PERIOD}`, size: 22, font: "Arial", color: "666666" })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [new TextRun({ text: `レビュー総数：${REVIEW_COUNT}`, size: 22, font: "Arial", color: "666666" })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [new TextRun({ text: TARGET_SITES, size: 20, font: "Arial", color: "666666" })] }),
        spacer(1600),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: `作成日：${REPORT_DATE}`, size: 20, font: "Arial", color: "999999" })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Confidential - For Internal Use Only", size: 18, font: "Arial", color: "AAAAAA", italics: true })] }),
      ],
    },
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
            children: [new TextRun({ text: "Page ", size: 16, font: "Arial", color: "999999" }), new TextRun({ children: [PageNumber.CURRENT], size: 16, font: "Arial", color: "999999" })],
          })],
        }),
      },
      children: [
        heading1("1. エグゼクティブサマリー"),
        para(EXECUTIVE_SUMMARY_INTRO),
        spacer(100),
        kpiRow(KPI_CARDS),
        spacer(200),
        heading2("総合評価"),
        para(EXECUTIVE_SUMMARY_EVALUATION_1),
        para(EXECUTIVE_SUMMARY_EVALUATION_2),
        spacer(100),
        new Table({ width: { size: 9026, type: WidthType.DXA }, columnWidths: [9026], rows: [new TableRow({ children: [new TableCell({ borders: { top: { style: BorderStyle.SINGLE, size: 4, color: ACCENT }, bottom: border, left: { style: BorderStyle.SINGLE, size: 4, color: ACCENT }, right: border }, width: { size: 9026, type: WidthType.DXA }, shading: { fill: "F0F7FC", type: ShadingType.CLEAR }, margins: { top: 200, bottom: 200, left: 300, right: 300 }, children: [
          new Paragraph({ spacing: { after: 120 }, children: [new TextRun({ text: "KEY FINDINGS", bold: true, size: 22, font: "Arial", color: ACCENT })] }),
          new Paragraph({ spacing: { after: 80 }, children: [new TextRun({ text: "Strength：", bold: true, size: 20, font: "Arial", color: GREEN_ACCENT }), new TextRun({ text: KEY_FINDING_STRENGTH, size: 20, font: "Arial", color: "333333" })] }),
          new Paragraph({ spacing: { after: 80 }, children: [new TextRun({ text: "Weakness：", bold: true, size: 20, font: "Arial", color: RED_ACCENT }), new TextRun({ text: KEY_FINDING_WEAKNESS, size: 20, font: "Arial", color: "333333" })] }),
          new Paragraph({ children: [new TextRun({ text: "Opportunity：", bold: true, size: 20, font: "Arial", color: ORANGE_ACCENT }), new TextRun({ text: KEY_FINDING_OPPORTUNITY, size: 20, font: "Arial", color: "333333" })] }),
        ] })] })] }),
        new Paragraph({ children: [new PageBreak()] }),
        heading1("2. データ概要"),
        heading2("2.1 サイト別レビュー件数・評価"),
        para(DATA_OVERVIEW_SITE_TEXT),
        spacer(80),
        new Table({ width: { size: 9026, type: WidthType.DXA }, columnWidths: [1800, 900, 1300, 900, 1300, 1526, 1300], rows: [
          new TableRow({ children: [headerCell("サイト名", 1800), headerCell("件数", 900), headerCell("ネイティブ平均", 1300), headerCell("尺度", 900), headerCell("10pt換算", 1300), headerCell("中央値(10pt)", 1526), headerCell("判定", 1300)] }),
          ...buildSiteTableRows(SITE_DATA),
        ] }),
        spacer(200),
        heading2("2.2 評価分布（10点換算）"),
        para(DATA_OVERVIEW_DIST_TEXT),
        spacer(80),
        new Table({ width: { size: 9026, type: WidthType.DXA }, columnWidths: [1200, 1200, 1500, 5126], rows: [
          new TableRow({ children: [headerCell("評価", 1200), headerCell("件数", 1200), headerCell("割合", 1500), headerCell("分布", 5126)] }),
          ...buildDistributionRows(DISTRIBUTION_DATA, DISTRIBUTION_MAX_COUNT),
        ] }),
        spacer(100),
        new Table({ width: { size: 9026, type: WidthType.DXA }, columnWidths: [3008, 3009, 3009], rows: [new TableRow({ children: [
          new TableCell({ borders, width: { size: 3008, type: WidthType.DXA }, shading: { fill: "E8F5E9", type: ShadingType.CLEAR }, margins: { top: 120, bottom: 120, left: 200, right: 200 }, children: [
            new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [new TextRun({ text: "高評価（8-10点）", bold: true, size: 20, font: "Arial", color: GREEN_ACCENT })] }),
            new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: HIGH_RATING_SUMMARY, bold: true, size: 24, font: "Arial", color: GREEN_ACCENT })] }),
          ] }),
          new TableCell({ borders, width: { size: 3009, type: WidthType.DXA }, shading: { fill: "FFF3E0", type: ShadingType.CLEAR }, margins: { top: 120, bottom: 120, left: 200, right: 200 }, children: [
            new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [new TextRun({ text: "中評価（5-7点）", bold: true, size: 20, font: "Arial", color: ORANGE_ACCENT })] }),
            new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: MID_RATING_SUMMARY, bold: true, size: 24, font: "Arial", color: ORANGE_ACCENT })] }),
          ] }),
          new TableCell({ borders, width: { size: 3009, type: WidthType.DXA }, shading: { fill: "FDEDEC", type: ShadingType.CLEAR }, margins: { top: 120, bottom: 120, left: 200, right: 200 }, children: [
            new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [new TextRun({ text: "低評価（1-4点）", bold: true, size: 20, font: "Arial", color: LOW_RATING_COLOR })] }),
            new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: LOW_RATING_SUMMARY, bold: true, size: 24, font: "Arial", color: LOW_RATING_COLOR })] }),
          ] }),
        ] })] }),
        new Paragraph({ children: [new PageBreak()] }),
        heading1("3. 強み分析（ポジティブ要因）"),
        para("口コミのテキストマイニングにより、以下のポジティブテーマが特定されました。"),
        spacer(100),
        new Table({ width: { size: 9026, type: WidthType.DXA }, columnWidths: [2600, 1200, 5226], rows: [
          new TableRow({ children: [headerCell("ポジティブテーマ", 2600), headerCell("言及数", 1200), headerCell("代表的なコメント", 5226)] }),
          ...buildStrengthRows(STRENGTH_THEMES),
        ] }),
        spacer(200),
        heading2(STRENGTH_SUB_1_TITLE),
        para(STRENGTH_SUB_1_TEXT),
        ...STRENGTH_SUB_1_BULLETS.map(b => bulletItem(b)),
        spacer(100),
        heading2(STRENGTH_SUB_2_TITLE),
        para(STRENGTH_SUB_2_TEXT),
        new Paragraph({ children: [new PageBreak()] }),
        heading1("4. 弱み分析（改善課題）"),
        para("ネガティブコメントの分析から、以下の改善課題が抽出されました。影響度と頻度に基づく優先度をS〜Cで設定しています。"),
        spacer(100),
        new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [800, 2200, 4360, 2000], rows: [
          new TableRow({ children: [headerCell("優先度", 800), headerCell("課題カテゴリ", 2200), headerCell("具体的内容", 4360), headerCell("影響度", 2000)] }),
          ...WEAKNESS_PRIORITY_DATA.map(([p, c, d, i]) => priorityRow(p, c, d, i)),
        ] }),
        new Paragraph({ children: [new PageBreak()] }),
        heading1("5. 改善施策提案"),
        para("分析結果に基づき、以下の改善施策を「即座対応」「短期」「中期」の3フェーズに分けて提案いたします。"),
        spacer(100),
        heading2("Phase 1：即座対応（今週〜1ヶ月以内）"),
        para(PHASE1_DESCRIPTION),
        spacer(80),
        ...buildPhaseItems(PHASE1_ITEMS),
        new Paragraph({ children: [new PageBreak()] }),
        heading2("Phase 2：短期施策（1〜3ヶ月）"),
        para(PHASE2_DESCRIPTION),
        spacer(80),
        ...buildPhaseItems(PHASE2_ITEMS),
        spacer(200),
        heading2("Phase 3：中期施策（3〜6ヶ月）"),
        para(PHASE3_DESCRIPTION),
        spacer(80),
        ...buildPhaseItems(PHASE3_ITEMS),
        new Paragraph({ children: [new PageBreak()] }),
        heading1("6. KPI目標設定"),
        para("以下のKPIを設定し、四半期ごとにモニタリングすることを推奨します。"),
        spacer(100),
        new Table({ width: { size: 9026, type: WidthType.DXA }, columnWidths: [2800, 2200, 2200, 1826], rows: [
          new TableRow({ children: [headerCell("KPI項目", 2800), headerCell("現状値", 2200), headerCell("目標値（6ヶ月後）", 2200), headerCell("期限", 1826)] }),
          ...buildKPITargetRows(KPI_TARGET_DATA),
        ] }),
        new Paragraph({ children: [new PageBreak()] }),
        heading1("7. 総括と今後のアクション"),
        spacer(100),
        new Table({ width: { size: 9026, type: WidthType.DXA }, columnWidths: [9026], rows: [new TableRow({ children: [new TableCell({
          borders: { top: { style: BorderStyle.SINGLE, size: 4, color: NAVY }, bottom: { style: BorderStyle.SINGLE, size: 4, color: NAVY }, left: { style: BorderStyle.SINGLE, size: 4, color: NAVY }, right: { style: BorderStyle.SINGLE, size: 4, color: NAVY } },
          width: { size: 9026, type: WidthType.DXA },
          shading: { fill: "F8F9FA", type: ShadingType.CLEAR },
          margins: { top: 300, bottom: 300, left: 400, right: 400 },
          children: [
            ...CONCLUSION_PARAGRAPHS.map(text => new Paragraph({ spacing: { after: 160 }, children: [new TextRun({ text, size: 21, font: "Arial", color: "333333" })] })),
            new Paragraph({ children: [new TextRun({ text: CONCLUSION_FINAL_PARAGRAPH, size: 21, font: "Arial", color: NAVY, bold: true })] }),
          ],
        })] })] }),
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
