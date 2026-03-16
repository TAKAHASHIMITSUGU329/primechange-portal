const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, PageNumber, PageBreak, LevelFormat
} = require("docx");

// Color scheme
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
function divider() {
  return new Paragraph({ spacing: { before: 200, after: 200 }, border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: ACCENT, space: 1 } }, children: [] });
}
function kpiRow(items) {
  const colWidth = Math.floor(9360 / items.length);
  return new Table({
    width: { size: 9360, type: WidthType.DXA }, columnWidths: items.map(() => colWidth),
    rows: [new TableRow({
      children: items.map(item => new TableCell({
        borders: { ...noBorders, right: { style: BorderStyle.SINGLE, size: 1, color: "DDDDDD" } },
        width: { size: colWidth, type: WidthType.DXA },
        shading: { fill: item.bgColor || LIGHT_BLUE, type: ShadingType.CLEAR },
        margins: { top: 160, bottom: 160, left: 200, right: 200 },
        children: [
          new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [new TextRun({ text: item.label, size: 18, font: "Arial", color: "666666" })] }),
          new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: item.value, bold: true, size: 36, font: "Arial", color: item.color || NAVY })] }),
        ],
      })),
    })],
  });
}
function priorityRow(priority, category, detail, impact) {
  const colors = { "S": RED_ACCENT, "A": ORANGE_ACCENT, "B": ACCENT, "C": "888888" };
  return new TableRow({
    children: [
      new TableCell({ borders, width: { size: 800, type: WidthType.DXA }, margins: cellMargins, shading: { fill: LIGHT_GRAY, type: ShadingType.CLEAR }, verticalAlign: "center",
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: priority, bold: true, size: 22, font: "Arial", color: colors[priority] || NAVY })] })] }),
      dataCell(category, 2200, { bold: true }), dataCell(detail, 4360), dataCell(impact, 2000, { alignment: AlignmentType.CENTER }),
    ],
  });
}

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
  ]},
  sections: [
    // ===== COVER PAGE =====
    {
      properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
      children: [
        spacer(2000),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: ACCENT, space: 8 } },
          children: [new TextRun({ text: "口コミ分析", size: 56, font: "Arial", color: ACCENT })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 },
          children: [new TextRun({ text: "改善レポート", size: 56, font: "Arial", bold: true, color: NAVY })] }),
        spacer(400),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 120 },
          children: [new TextRun({ text: "チサンホテル浜松町", size: 32, font: "Arial", color: NAVY, bold: true })] }),
        spacer(200),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 },
          children: [new TextRun({ text: "分析対象期間：2026年2月〜3月", size: 22, font: "Arial", color: "666666" })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 },
          children: [new TextRun({ text: "レビュー総数：47件（重複除外後）", size: 22, font: "Arial", color: "666666" })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 },
          children: [new TextRun({ text: "対象サイト：楽天トラベル / Google / Agoda / じゃらん / Booking.com", size: 20, font: "Arial", color: "666666" })] }),
        spacer(1600),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 },
          children: [new TextRun({ text: "作成日：2026年3月7日", size: 20, font: "Arial", color: "999999" })] }),
        new Paragraph({ alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "Confidential - For Internal Use Only", size: 18, font: "Arial", color: "AAAAAA", italics: true })] }),
      ],
    },

    // ===== MAIN CONTENT =====
    {
      properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
      headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, border: { bottom: { style: BorderStyle.SINGLE, size: 2, color: ACCENT, space: 4 } },
        children: [new TextRun({ text: "チサンホテル浜松町｜口コミ分析改善レポート", size: 16, font: "Arial", color: "999999" })] })] }) },
      footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "Page ", size: 16, font: "Arial", color: "999999" }), new TextRun({ children: [PageNumber.CURRENT], size: 16, font: "Arial", color: "999999" })] })] }) },
      children: [
        // ===== 1. EXECUTIVE SUMMARY =====
        heading1("1. エグゼクティブサマリー"),
        para("2026年2月〜3月に各OTAサイト・口コミサイトに投稿された47件のレビューを包括的に分析しました。当ホテルは全体平均7.15点（10点換算）で中程度の水準にあり、強みを活かした戦略的改善が求められます。"),
        spacer(100),

        kpiRow([
          { label: "全体平均評価(10pt)", value: "7.15", color: ACCENT, bgColor: LIGHT_BLUE },
          { label: "高評価率(8-10点)", value: "59.6%", color: GREEN_ACCENT, bgColor: "E8F5E9" },
          { label: "低評価率(1-4点)", value: "21.3%", color: ORANGE_ACCENT, bgColor: "FFF3E0" },
          { label: "レビュー総数", value: "47件", color: NAVY, bgColor: LIGHT_BLUE },
        ]),
        spacer(200),

        heading2("総合評価"),
        para("当ホテルは10点換算で全体平均7.15点と、概ね中程度の評価水準にあります。高評価率59.6%は過半数の宿泊客が満足していることを示しており、基本的な宿泊体験は一定の水準を維持しています。"),
        para("海外OTA（Agoda 8.67点、Booking.com 8.00点）では高い評価を得ている一方、国内サイトはGoogle（3.08/5点＝6.17/10点換算）が最も低く、じゃらん（3.43/5点＝6.86/10点換算）、楽天トラベル（3.53/5点＝7.06/10点換算）と続きます。国内サイトの評価底上げが全体平均向上の鍵です。"),
        para("なお、海外OTA（Agoda/Booking.com）は10点満点、国内サイト（楽天トラベル/じゃらん/Google）は5点満点という異なる評価尺度を使用しています。本レポートでは5点満点の評価を2倍して10点満点に換算し、統一的な比較を行っています。"),
        spacer(100),

        new Table({
          width: { size: 9026, type: WidthType.DXA }, columnWidths: [9026],
          rows: [new TableRow({ children: [new TableCell({
            borders: { top: { style: BorderStyle.SINGLE, size: 4, color: ACCENT }, bottom: border, left: { style: BorderStyle.SINGLE, size: 4, color: ACCENT }, right: border },
            width: { size: 9026, type: WidthType.DXA }, shading: { fill: "F0F7FC", type: ShadingType.CLEAR },
            margins: { top: 200, bottom: 200, left: 300, right: 300 },
            children: [
              new Paragraph({ spacing: { after: 120 }, children: [new TextRun({ text: "KEY FINDINGS", bold: true, size: 22, font: "Arial", color: ACCENT })] }),
              new Paragraph({ spacing: { after: 80 }, children: [
                new TextRun({ text: "Strength：", bold: true, size: 20, font: "Arial", color: GREEN_ACCENT }),
                new TextRun({ text: "立地の利便性（16件）、コスパ（12件）、清潔感（10件）。基盤は健在", size: 20, font: "Arial", color: "333333" }),
              ]}),
              new Paragraph({ spacing: { after: 80 }, children: [
                new TextRun({ text: "Weakness：", bold: true, size: 20, font: "Arial", color: RED_ACCENT }),
                new TextRun({ text: "設備老朽化（10件）、騒音・防音（9件）、水回り（6件）、清掃（5件）", size: 20, font: "Arial", color: "333333" }),
              ]}),
              new Paragraph({ children: [
                new TextRun({ text: "Opportunity：", bold: true, size: 20, font: "Arial", color: ORANGE_ACCENT }),
                new TextRun({ text: "朝食の好評（7件）・立地の強みを活かし、Google評価改善とOTA最適化で全体水準の底上げが可能", size: 20, font: "Arial", color: "333333" }),
              ]}),
            ],
          })] })],
        }),

        new Paragraph({ children: [new PageBreak()] }),

        // ===== 2. DATA OVERVIEW =====
        heading1("2. データ概要"),
        heading2("2.1 サイト別レビュー件数・評価"),
        para("各予約サイトの評価基準が異なります。海外OTA（Booking.com/Agoda）は10点満点、国内サイト（楽天トラベル/じゃらん/Google）は5点満点です。公平な比較のため10点換算値を併記します。"),
        spacer(80),

        new Table({
          width: { size: 9026, type: WidthType.DXA }, columnWidths: [2000, 1000, 1500, 1000, 1526, 1000],
          rows: [
            new TableRow({ children: [headerCell("サイト名", 2000), headerCell("件数", 1000), headerCell("平均評価", 1500), headerCell("尺度", 1000), headerCell("10点換算", 1526), headerCell("判定", 1000)] }),
            new TableRow({ children: [
              dataCell("Agoda", 2000, { bold: true }), dataCell("9", 1000, { alignment: AlignmentType.CENTER }),
              dataCell("8.67", 1500, { alignment: AlignmentType.CENTER, color: GREEN_ACCENT, bold: true }),
              dataCell("/10", 1000, { alignment: AlignmentType.CENTER }),
              dataCell("8.67", 1526, { alignment: AlignmentType.CENTER, color: GREEN_ACCENT, bold: true }),
              dataCell("良好", 1000, { alignment: AlignmentType.CENTER, color: GREEN_ACCENT, bold: true }),
            ]}),
            new TableRow({ children: [
              dataCell("Booking.com", 2000, { bold: true, fill: LIGHT_GRAY }), dataCell("2", 1000, { alignment: AlignmentType.CENTER, fill: LIGHT_GRAY }),
              dataCell("8.00", 1500, { alignment: AlignmentType.CENTER, color: GREEN_ACCENT, bold: true, fill: LIGHT_GRAY }),
              dataCell("/10", 1000, { alignment: AlignmentType.CENTER, fill: LIGHT_GRAY }),
              dataCell("8.00", 1526, { alignment: AlignmentType.CENTER, color: GREEN_ACCENT, bold: true, fill: LIGHT_GRAY }),
              dataCell("良好", 1000, { alignment: AlignmentType.CENTER, color: GREEN_ACCENT, bold: true, fill: LIGHT_GRAY }),
            ]}),
            new TableRow({ children: [
              dataCell("楽天トラベル", 2000, { bold: true }), dataCell("17", 1000, { alignment: AlignmentType.CENTER }),
              dataCell("3.53", 1500, { alignment: AlignmentType.CENTER, color: ORANGE_ACCENT, bold: true }),
              dataCell("/5", 1000, { alignment: AlignmentType.CENTER }),
              dataCell("7.06", 1526, { alignment: AlignmentType.CENTER, color: ORANGE_ACCENT, bold: true }),
              dataCell("平均的", 1000, { alignment: AlignmentType.CENTER, color: ORANGE_ACCENT, bold: true }),
            ]}),
            new TableRow({ children: [
              dataCell("じゃらん", 2000, { bold: true, fill: LIGHT_GRAY }), dataCell("7", 1000, { alignment: AlignmentType.CENTER, fill: LIGHT_GRAY }),
              dataCell("3.43", 1500, { alignment: AlignmentType.CENTER, color: ORANGE_ACCENT, bold: true, fill: LIGHT_GRAY }),
              dataCell("/5", 1000, { alignment: AlignmentType.CENTER, fill: LIGHT_GRAY }),
              dataCell("6.86", 1526, { alignment: AlignmentType.CENTER, color: ORANGE_ACCENT, bold: true, fill: LIGHT_GRAY }),
              dataCell("要改善", 1000, { alignment: AlignmentType.CENTER, color: ORANGE_ACCENT, bold: true, fill: LIGHT_GRAY }),
            ]}),
            new TableRow({ children: [
              dataCell("Google", 2000, { bold: true }), dataCell("12", 1000, { alignment: AlignmentType.CENTER }),
              dataCell("3.08", 1500, { alignment: AlignmentType.CENTER, color: RED_ACCENT, bold: true }),
              dataCell("/5", 1000, { alignment: AlignmentType.CENTER }),
              dataCell("6.17", 1526, { alignment: AlignmentType.CENTER, color: RED_ACCENT, bold: true }),
              dataCell("要改善", 1000, { alignment: AlignmentType.CENTER, color: RED_ACCENT, bold: true }),
            ]}),
          ],
        }),
        spacer(200),

        heading2("2.2 評価分布（10点満点換算）"),
        para("10点満点に換算した評価分布です。国内5点満点サイトの評価は×2で変換しているため、偶数値（10,8,6,4,2）に集中します。8点（15件・31.9%）が最多で、高評価層が厚い一方、4点以下も10件（21.3%）あり改善余地が認められます。"),
        spacer(80),

        new Table({
          width: { size: 9026, type: WidthType.DXA }, columnWidths: [1200, 1200, 1500, 5126],
          rows: [
            new TableRow({ children: [headerCell("評価(10pt)", 1200), headerCell("件数", 1200), headerCell("割合", 1500), headerCell("分布", 5126)] }),
            ...[[10,11,"23.4%"], [9,2,"4.3%"], [8,15,"31.9%"], [6,9,"19.1%"], [4,7,"14.9%"], [2,3,"6.4%"]].map(([rating, count, pct]) => {
              const barWidth = Math.round(count / 15 * 100);
              const isLow = rating <= 4;
              const barColor = isLow ? RED_ACCENT : (rating <= 6 ? ORANGE_ACCENT : GREEN_ACCENT);
              const fill = [10,8,4].includes(rating) ? LIGHT_GRAY : WHITE;
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
            }),
          ],
        }),
        spacer(100),
        para("※5点満点サイト（楽天・じゃらん・Google）の評価は×2で10点換算。奇数値（7,5,3,1点）は海外OTAの評価のみで構成されます。", { color: "888888" }),
        spacer(100),

        new Table({
          width: { size: 9026, type: WidthType.DXA }, columnWidths: [3008, 3009, 3009],
          rows: [new TableRow({ children: [
            new TableCell({ borders, width: { size: 3008, type: WidthType.DXA }, shading: { fill: "E8F5E9", type: ShadingType.CLEAR }, margins: { top: 120, bottom: 120, left: 200, right: 200 },
              children: [
                new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [new TextRun({ text: "高評価（8-10点）", bold: true, size: 20, font: "Arial", color: GREEN_ACCENT })] }),
                new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "28件（59.6%）", bold: true, size: 24, font: "Arial", color: GREEN_ACCENT })] }),
              ] }),
            new TableCell({ borders, width: { size: 3009, type: WidthType.DXA }, shading: { fill: "FFF3E0", type: ShadingType.CLEAR }, margins: { top: 120, bottom: 120, left: 200, right: 200 },
              children: [
                new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [new TextRun({ text: "中評価（5-7点）", bold: true, size: 20, font: "Arial", color: ORANGE_ACCENT })] }),
                new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "9件（19.1%）", bold: true, size: 24, font: "Arial", color: ORANGE_ACCENT })] }),
              ] }),
            new TableCell({ borders, width: { size: 3009, type: WidthType.DXA }, shading: { fill: "FDEDEC", type: ShadingType.CLEAR }, margins: { top: 120, bottom: 120, left: 200, right: 200 },
              children: [
                new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [new TextRun({ text: "低評価（1-4点）", bold: true, size: 20, font: "Arial", color: RED_ACCENT })] }),
                new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "10件（21.3%）", bold: true, size: 24, font: "Arial", color: RED_ACCENT })] }),
              ] }),
          ] })],
        }),

        new Paragraph({ children: [new PageBreak()] }),

        // ===== 3. STRENGTH ANALYSIS =====
        heading1("3. 強み分析（ポジティブ要因）"),
        para("口コミのテキストマイニングにより、以下6つのポジティブテーマが特定されました。"),
        spacer(100),

        new Table({
          width: { size: 9026, type: WidthType.DXA }, columnWidths: [2600, 1200, 5226],
          rows: [
            new TableRow({ children: [headerCell("ポジティブテーマ", 2600), headerCell("言及数", 1200), headerCell("代表的な評価ポイント", 5226)] }),
            new TableRow({ children: [dataCell("立地・アクセス", 2600, { bold: true, fill: "E8F5E9" }), dataCell("16件", 1200, { alignment: AlignmentType.CENTER, bold: true, color: GREEN_ACCENT, fill: "E8F5E9" }),
              dataCell("浜松町駅・大門駅から徒歩圏内、羽田空港へのモノレールアクセス良好、東京タワー近く", 5226, { fill: "E8F5E9" })] }),
            new TableRow({ children: [dataCell("コストパフォーマンス", 2600, { bold: true }), dataCell("12件", 1200, { alignment: AlignmentType.CENTER, bold: true, color: GREEN_ACCENT }),
              dataCell("料金が安い、価格の割に立地が良い、ビジネス利用にちょうど良い価格帯", 5226)] }),
            new TableRow({ children: [dataCell("清潔感", 2600, { bold: true, fill: LIGHT_GRAY }), dataCell("10件", 1200, { alignment: AlignmentType.CENTER, bold: true, color: GREEN_ACCENT, fill: LIGHT_GRAY }),
              dataCell("清掃が行き届いている、部屋がきれい、気持ちよく過ごせた", 5226, { fill: LIGHT_GRAY })] }),
            new TableRow({ children: [dataCell("設備・部屋", 2600, { bold: true }), dataCell("8件", 1200, { alignment: AlignmentType.CENTER, bold: true, color: GREEN_ACCENT }),
              dataCell("部屋が思ったより広い、設備が整っている、使い勝手が良い", 5226)] }),
            new TableRow({ children: [dataCell("朝食", 2600, { bold: true, fill: LIGHT_GRAY }), dataCell("7件", 1200, { alignment: AlignmentType.CENTER, bold: true, color: GREEN_ACCENT, fill: LIGHT_GRAY }),
              dataCell("朝食が美味しい、品数が豊富、和食メニューが充実", 5226, { fill: LIGHT_GRAY })] }),
            new TableRow({ children: [dataCell("スタッフ対応", 2600, { bold: true }), dataCell("6件", 1200, { alignment: AlignmentType.CENTER, bold: true, color: GREEN_ACCENT }),
              dataCell("スタッフが親切、丁寧な対応、フロントの笑顔が良い", 5226)] }),
          ],
        }),
        spacer(200),

        heading2("3.1 最大の強み：「立地・アクセス」"),
        para("全コメントの34%が立地の良さに言及しており、当ホテルの最大の競争優位性です。"),
        bulletItem("浜松町駅・大門駅から徒歩圏内の交通アクセス"),
        bulletItem("羽田空港へのモノレール直結（出張・インバウンド需要）"),
        bulletItem("東京タワー・芝公園への近さによる観光利便性"),
        spacer(100),

        heading2("3.2 コストパフォーマンスと朝食"),
        para("コスパの良さ（12件）と朝食の好評（7件）は当ホテルの差別化要素です。特に和食メニューの充実が評価されており、OTAでのアピールポイントとして更に活用すべき強みです。"),

        new Paragraph({ children: [new PageBreak()] }),

        // ===== 4. WEAKNESS ANALYSIS =====
        heading1("4. 弱み分析（改善課題）"),
        para("ネガティブコメントの分析から、以下の改善課題が抽出されました。影響度と頻度に基づく優先度をS〜Cで設定しています。"),
        spacer(100),

        new Table({
          width: { size: 9360, type: WidthType.DXA }, columnWidths: [800, 2200, 4360, 2000],
          rows: [
            new TableRow({ children: [headerCell("優先度", 800), headerCell("課題カテゴリ", 2200), headerCell("具体的内容", 4360), headerCell("影響度", 2000)] }),
            priorityRow("S", "設備老朽化", "壁紙の劣化、カーペットの汚れ、家具の傷み、ドアの建て付け不良", "評価直結・10件"),
            priorityRow("S", "騒音・防音", "隣室の音、電車の音、エレベーター振動、廊下の足音", "快適性低下・9件"),
            priorityRow("A", "水回り", "水圧弱い、排水遅い、バスルーム狭い、シャワーヘッド劣化", "基本品質・6件"),
            priorityRow("A", "清掃品質", "髪の毛、ほこり、ベッド周りの清掃不足、換気不十分", "衛生認知・5件"),
            priorityRow("B", "コスパ割高感", "設備の古さに対して料金が高いと感じる声", "価格妥当性・4件"),
            priorityRow("B", "立地（ネガティブ）", "駅からの道順が分かりにくい、夜道が暗い", "アクセス不安・3件"),
            priorityRow("C", "タバコ臭", "禁煙室でもタバコの匂いが残存", "品質認知・2件"),
            priorityRow("C", "空調", "エアコンの効きが悪い、温度調整困難", "快適性低下・2件"),
            priorityRow("C", "加湿器", "加湿器が古い、または設置されていない", "季節的課題・1件"),
          ],
        }),

        new Paragraph({ children: [new PageBreak()] }),

        // ===== 5. IMPROVEMENT PLAN =====
        heading1("5. 改善施策提案"),
        para("分析結果に基づき、以下の改善施策を「即座対応」「短期」「中期」の3フェーズに分けて提案いたします。"),
        spacer(100),

        heading2("Phase 1：即座対応（今週〜1ヶ月以内）"),
        para("投資を最小限に抑え、オペレーション改善で効果を目指す施策"),
        spacer(80),
        heading3("(1) 清掃品質の強化"),
        bulletItem("清掃チェックリストの見直し：髪の毛、ほこり、水垢の重点チェック項目を追加"),
        bulletItem("ダブルチェック体制の導入（清掃スタッフ→チームリーダーの2段階確認）"),
        bulletItem("清掃後の換気時間を最低30分確保するオペレーション変更"),
        spacer(80),
        heading3("(2) 騒音対策（即効性のあるもの）"),
        bulletItem("簡易防音テープの設置（ドア下部・窓際）"),
        bulletItem("耳栓のアメニティ提供を検討"),
        bulletItem("騒音の多い客室を特定し、静かな客室を優先配分"),
        spacer(80),
        heading3("(3) 水回りメンテナンス"),
        bulletItem("シャワーヘッドの目詰まりチェック・交換を月次で実施"),
        bulletItem("排水口クリーニングの頻度を月2回に"),
        spacer(80),
        heading3("(4) 口コミ返信の開始・強化"),
        bulletItem("全サイトの口コミに48時間以内の返信を開始"),
        bulletItem("低評価口コミには謝罪と具体的改善策を記載"),
        bulletItem("Google口コミの改善を最優先（12件・平均3.08/5で最も低い）"),

        new Paragraph({ children: [new PageBreak()] }),

        heading2("Phase 2：短期施策（1〜3ヶ月）"),
        para("一定の投資を伴うが、比較的早期に実行可能な施策"),
        spacer(80),
        heading3("(1) 客室リフレッシュ（部分改修）"),
        bulletItem("壁紙・カーペットの交換を優先フロアから段階的に実施"),
        bulletItem("照明のLED化（明るく清潔感のある印象を演出）"),
        bulletItem("小物（ゴミ箱、ハンガー、リモコン等）の新品交換"),
        spacer(80),
        heading3("(2) OTA対策の強化"),
        bulletItem("楽天トラベル・じゃらんの施設写真を刷新"),
        bulletItem("Googleビジネスプロフィールの写真・情報を最新化"),
        bulletItem("朝食の魅力をOTA写真・説明で積極的にアピール"),
        spacer(80),
        heading3("(3) アクセス案内の改善"),
        bulletItem("駅からの写真付きアクセスガイドを作成し、予約確認メールに添付"),
        bulletItem("QRコード経路案内のフロントカード設置"),

        spacer(200),

        heading2("Phase 3：中期施策（3〜6ヶ月）"),
        para("設備投資を伴う抜本的な改善施策"),
        spacer(80),
        heading3("(1) 防音工事"),
        bulletItem("低評価集中フロアの防音調査と改修計画策定"),
        bulletItem("窓の二重サッシ化・壁面の遮音材追加の検討"),
        spacer(80),
        heading3("(2) 水回りリニューアル"),
        bulletItem("水圧改善のための給水設備の更新"),
        bulletItem("バスルーム内装の改修"),
        spacer(80),
        heading3("(3) 段階的客室リノベーション"),
        bulletItem("最も老朽化が進んでいるフロアからの段階的改修"),
        bulletItem("USB充電ポート、高速WiFi等のデジタルインフラ整備"),

        new Paragraph({ children: [new PageBreak()] }),

        // ===== 6. KPI & TARGETS =====
        heading1("6. KPI目標設定"),
        para("以下のKPIを設定し、四半期ごとにモニタリングすることを推奨します。"),
        spacer(100),

        new Table({
          width: { size: 9026, type: WidthType.DXA }, columnWidths: [2800, 2200, 2200, 1826],
          rows: [
            new TableRow({ children: [headerCell("KPI項目", 2800), headerCell("現状値", 2200), headerCell("目標値（6ヶ月後）", 2200), headerCell("期限", 1826)] }),
            new TableRow({ children: [
              dataCell("全体平均評価(10pt)", 2800, { bold: true }), dataCell("7.15点", 2200, { alignment: AlignmentType.CENTER }),
              dataCell("7.8点以上", 2200, { alignment: AlignmentType.CENTER, color: GREEN_ACCENT, bold: true }), dataCell("2026年9月", 1826, { alignment: AlignmentType.CENTER }),
            ]}),
            new TableRow({ children: [
              dataCell("高評価率（8-10点）", 2800, { bold: true, fill: LIGHT_GRAY }), dataCell("59.6%", 2200, { alignment: AlignmentType.CENTER, fill: LIGHT_GRAY }),
              dataCell("65%以上", 2200, { alignment: AlignmentType.CENTER, color: GREEN_ACCENT, bold: true, fill: LIGHT_GRAY }), dataCell("2026年9月", 1826, { alignment: AlignmentType.CENTER, fill: LIGHT_GRAY }),
            ]}),
            new TableRow({ children: [
              dataCell("低評価率（1-4点）", 2800, { bold: true }), dataCell("21.3%", 2200, { alignment: AlignmentType.CENTER }),
              dataCell("15%以下", 2200, { alignment: AlignmentType.CENTER, color: GREEN_ACCENT, bold: true }), dataCell("2026年9月", 1826, { alignment: AlignmentType.CENTER }),
            ]}),
            new TableRow({ children: [
              dataCell("楽天トラベル平均", 2800, { bold: true, fill: LIGHT_GRAY }), dataCell("3.53/5点", 2200, { alignment: AlignmentType.CENTER, fill: LIGHT_GRAY }),
              dataCell("3.8/5点以上", 2200, { alignment: AlignmentType.CENTER, color: GREEN_ACCENT, bold: true, fill: LIGHT_GRAY }), dataCell("2026年9月", 1826, { alignment: AlignmentType.CENTER, fill: LIGHT_GRAY }),
            ]}),
            new TableRow({ children: [
              dataCell("Google評価", 2800, { bold: true }), dataCell("3.08/5点", 2200, { alignment: AlignmentType.CENTER, color: RED_ACCENT }),
              dataCell("3.5/5点以上", 2200, { alignment: AlignmentType.CENTER, color: GREEN_ACCENT, bold: true }), dataCell("2026年9月", 1826, { alignment: AlignmentType.CENTER }),
            ]}),
            new TableRow({ children: [
              dataCell("口コミ返信率", 2800, { bold: true, fill: LIGHT_GRAY }), dataCell("未計測", 2200, { alignment: AlignmentType.CENTER, fill: LIGHT_GRAY }),
              dataCell("100%（48h以内）", 2200, { alignment: AlignmentType.CENTER, color: GREEN_ACCENT, bold: true, fill: LIGHT_GRAY }), dataCell("2026年6月", 1826, { alignment: AlignmentType.CENTER, fill: LIGHT_GRAY }),
            ]}),
          ],
        }),

        new Paragraph({ children: [new PageBreak()] }),

        // ===== 7. CONCLUSION =====
        heading1("7. 総括と今後のアクション"),
        spacer(100),

        new Table({
          width: { size: 9026, type: WidthType.DXA }, columnWidths: [9026],
          rows: [new TableRow({ children: [new TableCell({
            borders: { top: { style: BorderStyle.SINGLE, size: 4, color: NAVY }, bottom: { style: BorderStyle.SINGLE, size: 4, color: NAVY }, left: { style: BorderStyle.SINGLE, size: 4, color: NAVY }, right: { style: BorderStyle.SINGLE, size: 4, color: NAVY } },
            width: { size: 9026, type: WidthType.DXA }, shading: { fill: "F8F9FA", type: ShadingType.CLEAR },
            margins: { top: 300, bottom: 300, left: 400, right: 400 },
            children: [
              new Paragraph({ spacing: { after: 160 }, children: [new TextRun({ text: "チサンホテル浜松町は全体平均7.15点（10点換算）、高評価率59.6%と、中程度の評価水準を維持しています。「立地」「コスパ」「朝食」の3つの強みは健在であり、海外OTAでの高評価も安定しています。", size: 21, font: "Arial", color: "333333" })] }),
              new Paragraph({ spacing: { after: 160 }, children: [new TextRun({ text: "一方で、Google（3.08/5点）とじゃらん（3.43/5点）の国内サイト評価に改善余地があり、「設備老朽化」「騒音・防音」「水回り」の3課題が低評価の主因となっています。", size: 21, font: "Arial", color: "333333" })] }),
              new Paragraph({ spacing: { after: 160 }, children: [new TextRun({ text: "Phase 1のオペレーション改善（清掃強化・騒音対策・口コミ返信）を着実に実行し、月次モニタリングで効果を検証しながら、Phase 2・3の施策を計画的に展開していくことを推奨いたします。", size: 21, font: "Arial", color: "333333" })] }),
              new Paragraph({ children: [new TextRun({ text: "特にGoogle口コミの改善は集客に直結するため、返信対応の開始と写真・情報の更新を優先的に進めることが重要です。", size: 21, font: "Arial", color: NAVY, bold: true })] }),
            ],
          })] })],
        }),

        spacer(300), divider(), spacer(100),
        para("本レポートに関するご質問・ご相談がございましたら、お気軽にお問い合わせください。", { alignment: AlignmentType.CENTER, color: "999999" }),
      ],
    },
  ],
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("/Users/mitsugutakahashi/ホテル口コミ/チサンホテル浜松町_口コミ分析改善レポート.docx", buffer);
  console.log("DOCX created successfully!");
  console.log("File size: " + (buffer.length / 1024).toFixed(1) + " KB");
});
