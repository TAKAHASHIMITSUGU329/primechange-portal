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

// Helper functions
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

// Build the document
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
          children: [new TextRun({ text: "ダイワロイネットホテル東京大崎", size: 32, font: "Arial", color: NAVY, bold: true })],
        }),
        spacer(200),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 80 },
          children: [new TextRun({ text: "分析対象期間：2026年1月〜2月", size: 22, font: "Arial", color: "666666" })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 80 },
          children: [new TextRun({ text: "レビュー総数：71件（重複除外後）", size: 22, font: "Arial", color: "666666" })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 80 },
          children: [new TextRun({ text: "対象サイト：Booking.com / Trip.com / じゃらん / 楽天トラベル / Agoda / Google", size: 20, font: "Arial", color: "666666" })],
        }),
        spacer(1600),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 60 },
          children: [new TextRun({ text: "作成日：2026年3月7日", size: 20, font: "Arial", color: "999999" })],
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
            children: [new TextRun({ text: "ダイワロイネットホテル東京大崎｜口コミ分析改善レポート", size: 16, font: "Arial", color: "999999" })],
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
        para("2026年1月〜2月に各OTAサイト・口コミサイトに投稿された71件のレビューを包括的に分析しました。以下が主要な発見事項です。"),
        spacer(100),

        kpiRow([
          { label: "全体平均(10pt換算)", value: "8.77", color: GREEN_ACCENT, bgColor: "E8F5E9" },
          { label: "高評価率(8-10点)", value: "84.5%", color: GREEN_ACCENT, bgColor: "E8F5E9" },
          { label: "低評価率(1-4点)", value: "0.0%", color: GREEN_ACCENT, bgColor: "E8F5E9" },
          { label: "レビュー総数", value: "71件", color: NAVY, bgColor: LIGHT_BLUE },
        ]),
        spacer(200),

        heading2("総合評価"),
        para("当ホテルは10点換算で全体平均8.77点と高水準の評価を獲得しています。海外OTA（Booking.com平均8.62点、Trip.com平均9.58点）はもとより、国内サイトも10pt換算で全て8点以上（じゃらん8.36、楽天トラベル8.89、Google 9.00）と安定した高評価を維持しています。"),
        para("「立地」「清潔感」「スタッフ対応」の三大基本要素で非常に高い評価を得ており、高評価率84.5%、低評価率0.0%は業界でも優秀な水準です。更なる向上を目指し、「水回り・バスルーム」「部屋の狭さへの対策」「エレベーター混雑」「周辺案内の充実」の4領域を重点改善課題として提案します。"),
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
                  new TextRun({ text: "立地の利便性（32件で言及）、清潔感（17件）、スタッフの親切さ（17件）", size: 20, font: "Arial", color: "333333" }),
                ]}),
                new Paragraph({ spacing: { after: 80 }, children: [
                  new TextRun({ text: "Weakness：", bold: true, size: 20, font: "Arial", color: RED_ACCENT }),
                  new TextRun({ text: "水回り不備（7件）、部屋狭小感（6件）、エレベーター待ち（3件）", size: 20, font: "Arial", color: "333333" }),
                ]}),
                new Paragraph({ children: [
                  new TextRun({ text: "Opportunity：", bold: true, size: 20, font: "Arial", color: ORANGE_ACCENT }),
                  new TextRun({ text: "リピート意向が高く（9件が再訪希望）、体験価値の向上で高評価率をさらに伸ばせるポテンシャル大", size: 20, font: "Arial", color: "333333" }),
                ]}),
              ],
            })],
          })],
        }),

        new Paragraph({ children: [new PageBreak()] }),

        // ===== 2. DATA OVERVIEW =====
        heading1("2. データ概要"),

        heading2("2.1 サイト別レビュー件数・評価"),
        para("各予約サイトの評価基準が異なる点に留意が必要です。海外OTA（Booking.com/Trip.com/Agoda）は10点満点、国内サイト（じゃらん/楽天トラベル/Google）は5点満点のため、統一比較のため10点換算（5点満点×2）を行いました。"),
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
            new TableRow({ children: [
              dataCell("Trip.com", 1800, { bold: true }), dataCell("12", 900, { alignment: AlignmentType.CENTER }),
              dataCell("9.58", 1300, { alignment: AlignmentType.CENTER, bold: true }),
              dataCell("/10", 900, { alignment: AlignmentType.CENTER }),
              dataCell("9.58", 1300, { alignment: AlignmentType.CENTER, color: GREEN_ACCENT, bold: true }),
              dataCell("10.0", 1526, { alignment: AlignmentType.CENTER }),
              dataCell("優秀", 1300, { alignment: AlignmentType.CENTER, color: GREEN_ACCENT, bold: true }),
            ]}),
            new TableRow({ children: [
              dataCell("Google", 1800, { bold: true, fill: LIGHT_GRAY }), dataCell("2", 900, { alignment: AlignmentType.CENTER, fill: LIGHT_GRAY }),
              dataCell("4.50", 1300, { alignment: AlignmentType.CENTER, bold: true, fill: LIGHT_GRAY }),
              dataCell("/5", 900, { alignment: AlignmentType.CENTER, fill: LIGHT_GRAY }),
              dataCell("9.00", 1300, { alignment: AlignmentType.CENTER, color: GREEN_ACCENT, bold: true, fill: LIGHT_GRAY }),
              dataCell("9.0", 1526, { alignment: AlignmentType.CENTER, fill: LIGHT_GRAY }),
              dataCell("優秀", 1300, { alignment: AlignmentType.CENTER, color: GREEN_ACCENT, bold: true, fill: LIGHT_GRAY }),
            ]}),
            new TableRow({ children: [
              dataCell("楽天トラベル", 1800, { bold: true }), dataCell("9", 900, { alignment: AlignmentType.CENTER }),
              dataCell("4.44", 1300, { alignment: AlignmentType.CENTER, bold: true }),
              dataCell("/5", 900, { alignment: AlignmentType.CENTER }),
              dataCell("8.89", 1300, { alignment: AlignmentType.CENTER, color: GREEN_ACCENT, bold: true }),
              dataCell("8.0", 1526, { alignment: AlignmentType.CENTER }),
              dataCell("良好", 1300, { alignment: AlignmentType.CENTER, color: GREEN_ACCENT, bold: true }),
            ]}),
            new TableRow({ children: [
              dataCell("Booking.com", 1800, { bold: true, fill: LIGHT_GRAY }), dataCell("32", 900, { alignment: AlignmentType.CENTER, fill: LIGHT_GRAY }),
              dataCell("8.62", 1300, { alignment: AlignmentType.CENTER, bold: true, fill: LIGHT_GRAY }),
              dataCell("/10", 900, { alignment: AlignmentType.CENTER, fill: LIGHT_GRAY }),
              dataCell("8.62", 1300, { alignment: AlignmentType.CENTER, color: GREEN_ACCENT, bold: true, fill: LIGHT_GRAY }),
              dataCell("9.0", 1526, { alignment: AlignmentType.CENTER, fill: LIGHT_GRAY }),
              dataCell("良好", 1300, { alignment: AlignmentType.CENTER, color: GREEN_ACCENT, bold: true, fill: LIGHT_GRAY }),
            ]}),
            new TableRow({ children: [
              dataCell("Agoda", 1800, { bold: true }), dataCell("5", 900, { alignment: AlignmentType.CENTER }),
              dataCell("8.40", 1300, { alignment: AlignmentType.CENTER, bold: true }),
              dataCell("/10", 900, { alignment: AlignmentType.CENTER }),
              dataCell("8.40", 1300, { alignment: AlignmentType.CENTER, color: GREEN_ACCENT, bold: true }),
              dataCell("8.0", 1526, { alignment: AlignmentType.CENTER }),
              dataCell("良好", 1300, { alignment: AlignmentType.CENTER, color: GREEN_ACCENT, bold: true }),
            ]}),
            new TableRow({ children: [
              dataCell("じゃらん", 1800, { bold: true, fill: LIGHT_GRAY }), dataCell("11", 900, { alignment: AlignmentType.CENTER, fill: LIGHT_GRAY }),
              dataCell("4.18", 1300, { alignment: AlignmentType.CENTER, bold: true, fill: LIGHT_GRAY }),
              dataCell("/5", 900, { alignment: AlignmentType.CENTER, fill: LIGHT_GRAY }),
              dataCell("8.36", 1300, { alignment: AlignmentType.CENTER, color: GREEN_ACCENT, bold: true, fill: LIGHT_GRAY }),
              dataCell("10.0", 1526, { alignment: AlignmentType.CENTER, fill: LIGHT_GRAY }),
              dataCell("良好", 1300, { alignment: AlignmentType.CENTER, color: GREEN_ACCENT, bold: true, fill: LIGHT_GRAY }),
            ]}),
          ],
        }),
        spacer(200),

        heading2("2.2 評価分布（10点換算）"),
        para("国内サイト（5点満点）は×2で10点換算しています。全レビューの84.5%が8点以上の高評価であり、低評価（1-4点）は0件です。10点が最多（32件/45.1%）で、満点評価を得やすいホテルと言えます。"),
        spacer(80),

        new Table({
          width: { size: 9026, type: WidthType.DXA },
          columnWidths: [1200, 1200, 1500, 5126],
          rows: [
            new TableRow({ children: [headerCell("評価", 1200), headerCell("件数", 1200), headerCell("割合", 1500), headerCell("分布", 5126)] }),
            ...[[10,32,"45.1%"], [9,9,"12.7%"], [8,19,"26.8%"], [7,5,"7.0%"], [6,5,"7.0%"], [5,1,"1.4%"]].map(([rating, count, pct]) => {
              const barWidth = Math.round(count / 32 * 100);
              const barColor = rating >= 8 ? GREEN_ACCENT : (rating >= 5 ? ORANGE_ACCENT : RED_ACCENT);
              const fill = [10,8,6].includes(rating) ? LIGHT_GRAY : WHITE;
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
                  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "60件（84.5%）", bold: true, size: 24, font: "Arial", color: GREEN_ACCENT })] }),
                ],
              }),
              new TableCell({
                borders, width: { size: 3009, type: WidthType.DXA },
                shading: { fill: "FFF3E0", type: ShadingType.CLEAR },
                margins: { top: 120, bottom: 120, left: 200, right: 200 },
                children: [
                  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [new TextRun({ text: "中評価（5-7点）", bold: true, size: 20, font: "Arial", color: ORANGE_ACCENT })] }),
                  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "11件（15.5%）", bold: true, size: 24, font: "Arial", color: ORANGE_ACCENT })] }),
                ],
              }),
              new TableCell({
                borders, width: { size: 3009, type: WidthType.DXA },
                shading: { fill: "FDEDEC", type: ShadingType.CLEAR },
                margins: { top: 120, bottom: 120, left: 200, right: 200 },
                children: [
                  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [new TextRun({ text: "低評価（1-4点）", bold: true, size: 20, font: "Arial", color: GREEN_ACCENT })] }),
                  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "0件（0.0%）", bold: true, size: 24, font: "Arial", color: GREEN_ACCENT })] }),
                ],
              }),
            ],
          })],
        }),

        new Paragraph({ children: [new PageBreak()] }),

        // ===== 3. STRENGTH ANALYSIS =====
        heading1("3. 強み分析（ポジティブ要因）"),
        para("口コミのテキストマイニングにより、以下6つのポジティブテーマが特定されました。"),
        spacer(100),

        new Table({
          width: { size: 9026, type: WidthType.DXA },
          columnWidths: [2600, 1200, 5226],
          rows: [
            new TableRow({ children: [headerCell("ポジティブテーマ", 2600), headerCell("言及数", 1200), headerCell("代表的なコメント", 5226)] }),
            new TableRow({ children: [
              dataCell("立地・アクセス", 2600, { bold: true, fill: "E8F5E9" }),
              dataCell("32件", 1200, { alignment: AlignmentType.CENTER, bold: true, color: GREEN_ACCENT, fill: "E8F5E9" }),
              dataCell("「駅直結で出張に最適」「大崎駅南口から徒歩わずか3分」", 5226, { fill: "E8F5E9" }),
            ]}),
            new TableRow({ children: [
              dataCell("部屋・設備", 2600, { bold: true }),
              dataCell("22件", 1200, { alignment: AlignmentType.CENTER, bold: true, color: GREEN_ACCENT }),
              dataCell("「部屋は快適で居心地が良い」「アメニティも充実」「マッサージチェアがありがたい」", 5226),
            ]}),
            new TableRow({ children: [
              dataCell("清潔感", 2600, { bold: true, fill: LIGHT_GRAY }),
              dataCell("17件", 1200, { alignment: AlignmentType.CENTER, bold: true, color: GREEN_ACCENT, fill: LIGHT_GRAY }),
              dataCell("「ホテルは清潔」「何より綺麗」「清潔感があって気持ち良い」", 5226, { fill: LIGHT_GRAY }),
            ]}),
            new TableRow({ children: [
              dataCell("スタッフ対応", 2600, { bold: true }),
              dataCell("17件", 1200, { alignment: AlignmentType.CENTER, bold: true, color: GREEN_ACCENT }),
              dataCell("「スタッフは親切で丁寧」「笑顔で対応」「フレンドリーで温かく迎えてくれた」", 5226),
            ]}),
            new TableRow({ children: [
              dataCell("朝食", 2600, { bold: true, fill: LIGHT_GRAY }),
              dataCell("9件", 1200, { alignment: AlignmentType.CENTER, bold: true, color: GREEN_ACCENT, fill: LIGHT_GRAY }),
              dataCell("「朝食も美味しかった」「食事が美味しい」「ビュッフェを利用、十分でした」", 5226, { fill: LIGHT_GRAY }),
            ]}),
            new TableRow({ children: [
              dataCell("リピート意向", 2600, { bold: true }),
              dataCell("9件", 1200, { alignment: AlignmentType.CENTER, bold: true, color: GREEN_ACCENT }),
              dataCell("「5年ほど前から年に1〜2回利用」「また泊まりたいホテル」「次も利用したい」", 5226),
            ]}),
          ],
        }),
        spacer(200),

        heading2("3.1 最大の強み：「立地・アクセス」"),
        para("全コメントの45%が立地の良さに言及しており、これは当ホテルの最大の競争優位性です。特に以下のポイントが評価されています。"),
        bulletItem("大崎駅直結・徒歩圏内という抜群のアクセス（JR山手線・りんかい線）"),
        bulletItem("ディズニーリゾート、品川、東京へのアクセスの良さ"),
        bulletItem("屋根付き通路による全天候型の動線"),
        bulletItem("周辺の商業施設（ゲートシティ、ニューシティ）の充実"),
        spacer(100),

        heading2("3.2 高い清潔感とスタッフの質"),
        para("「清潔感」と「スタッフ対応」はそれぞれ17件で同数の言及があり、ホテルの基本品質が高水準であることを示しています。特にインバウンド旅行者からの評価が高く、Trip.comの平均9.58点は業界トップクラスです。"),

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
            priorityRow("S", "水回り・バスルーム", "風呂の匂い、湯船の清掃不備、トイレ位置の使いにくさ、ウォシュレットの配置問題", "評価直結・7件"),
            priorityRow("A", "部屋の狭さ", "スーツケースを広げるスペース不足、2名利用時の圧迫感。特に海外ゲストから指摘多い", "満足度低下・6件"),
            priorityRow("A", "エレベーター混雑", "混雑時の待ち時間が長い。チェックイン時に特に不満が集中", "第一印象悪化・3件"),
            priorityRow("B", "周辺環境の案内不足", "深夜のコンビニ、レストラン情報が不十分。夜間到着客の不安感", "利便性低下・5件"),
            priorityRow("B", "アクセス・案内", "ホテルへのアクセスが分かりにくいとの声。初訪問者が迷う", "機会損失・2件"),
            priorityRow("B", "チェックイン設備", "タッチパネルの反応不良、操作方法の不明瞭さ", "第一印象悪化・2件"),
            priorityRow("B", "スタッフ対応（情報提供）", "大浴場の案内誤り（隣接スポーツセンター施設との混同）", "信頼低下・2件"),
            priorityRow("C", "騒音・防音", "部屋内での物音が気になるとの声", "快適性低下・2件"),
            priorityRow("C", "客室清掃", "湯船の清掃が不十分との指摘", "品質認知低下・2件"),
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

        heading3("(1) 水回り清掃の徹底強化"),
        bulletItem("清掃チェックリストに「湯船の隅・排水口・換気扇」を追加し、ダブルチェック体制を導入"),
        bulletItem("清掃スタッフへの再教育（特にバスルーム清掃基準の見直し）"),
        bulletItem("水回りの匂い対策として、清掃後の換気時間を延長（最低30分）"),
        bulletItem("定期的な排水管クリーニングのスケジュール化（月1回→月2回）"),
        spacer(80),

        heading3("(2) エレベーター混雑の緩和"),
        bulletItem("チェックイン時間帯（15:00〜17:00）のエレベーター運行を最適化"),
        bulletItem("チェックイン手続きの分散化：モバイルチェックインの積極的な案内"),
        bulletItem("荷物先預かりサービスの案内を強化し、身軽な状態でのチェックインを促進"),
        bulletItem("エレベーター前に「次のエレベーターまでの推定待ち時間」表示の検討"),
        spacer(80),

        heading3("(3) 周辺情報・アクセス案内の充実"),
        bulletItem("チェックイン時に「深夜営業コンビニMAP」「周辺飲食店ガイド（24h対応含む）」を配布"),
        bulletItem("多言語対応の館内・周辺案内パンフレットの充実（英語・中国語・韓国語）"),
        bulletItem("ホテルHPおよびOTAページに大崎駅からの写真付きアクセスガイドを掲載"),
        bulletItem("QRコードでGoogle Map経路案内を表示するカードをフロントに設置"),
        spacer(80),

        heading3("(4) スタッフ情報共有の改善"),
        bulletItem("隣接スポーツセンターの大浴場について正確な案内文を作成・全スタッフに周知"),
        bulletItem("外国人スタッフ向けに施設案内のFAQマニュアルを多言語で整備"),
        bulletItem("電話予約時の案内品質をチェックするモニタリング体制の構築"),
        spacer(80),

        heading3("(5) チェックイン機器のメンテナンス"),
        bulletItem("タッチパネルの反応テストを毎朝のルーチンチェックに追加"),
        bulletItem("タッチパネル横に操作ガイド（図解）を常設"),
        bulletItem("機器不具合時の即座対応マニュアル整備（スタッフによる手動チェックインへの切替）"),

        new Paragraph({ children: [new PageBreak()] }),

        // Phase 2
        heading2("Phase 2：短期施策（1〜3ヶ月）"),
        para("一定の投資を伴うが、比較的早期に実行可能な施策"),
        spacer(80),

        heading3("(1) 客室の「広さ感」向上"),
        bulletItem("バゲージラック・スーツケース置き場の設置（折り畳み式）を検討"),
        bulletItem("ベッド下収納スペースの活用促進（案内POPの設置）"),
        bulletItem("壁掛けテレビへの変更による床面積の有効活用（窓前テレビ設置の見直し）"),
        bulletItem("ミラーの戦略的配置による視覚的な広がりの演出"),
        spacer(80),

        heading3("(2) OTA写真・説明文の改善"),
        bulletItem("楽天トラベルのバスルーム写真を刷新（広い湯船のサイズ感が伝わる撮影角度に変更）"),
        bulletItem("じゃらん・楽天の口コミ返信対応を強化（全件に48時間以内の返信を目標）"),
        bulletItem("各OTAの施設説明文に「駅直結」「大崎駅南口から徒歩3分」を明記"),
        bulletItem("海外OTAでの高評価を活かし、英語レビューの返信を充実させ更なる評価向上へ"),
        spacer(80),

        heading3("(3) 防音対策の強化"),
        bulletItem("低評価が集中する特定の客室・フロアの防音状況を調査"),
        bulletItem("必要に応じてドア下部の防音テープ、窓の気密性確認を実施"),
        bulletItem("耳栓のアメニティ提供を検討"),

        spacer(200),

        // Phase 3
        heading2("Phase 3：中期施策（3〜6ヶ月）"),
        para("設備投資を伴う抜本的な改善施策"),
        spacer(80),

        heading3("(1) バスルームリニューアル計画"),
        bulletItem("トイレとウォシュレットの位置関係の見直し（壁との距離改善）"),
        bulletItem("換気システムの更新による匂い対策の根本解決"),
        bulletItem("排水設備の大規模メンテナンス計画の策定"),
        spacer(80),

        heading3("(2) エレベーター効率化"),
        bulletItem("エレベーター制御システムの更新検討（AI制御による効率化）"),
        bulletItem("ピーク時間帯の運行パターン最適化"),
        spacer(80),

        heading3("(3) デジタル施策の強化"),
        bulletItem("自社アプリによるモバイルキー・モバイルチェックインの導入"),
        bulletItem("客室タブレットでの周辺情報提供・多言語対応"),
        bulletItem("自動翻訳対応のチャットボット導入による外国人ゲスト対応の向上"),

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
            new TableRow({ children: [
              dataCell("全体平均(10pt換算)", 2800, { bold: true }),
              dataCell("8.77点", 2200, { alignment: AlignmentType.CENTER }),
              dataCell("9.0点以上", 2200, { alignment: AlignmentType.CENTER, color: GREEN_ACCENT, bold: true }),
              dataCell("2026年9月", 1826, { alignment: AlignmentType.CENTER }),
            ]}),
            new TableRow({ children: [
              dataCell("高評価率（8-10点）", 2800, { bold: true, fill: LIGHT_GRAY }),
              dataCell("84.5%", 2200, { alignment: AlignmentType.CENTER, fill: LIGHT_GRAY }),
              dataCell("90%以上", 2200, { alignment: AlignmentType.CENTER, color: GREEN_ACCENT, bold: true, fill: LIGHT_GRAY }),
              dataCell("2026年9月", 1826, { alignment: AlignmentType.CENTER, fill: LIGHT_GRAY }),
            ]}),
            new TableRow({ children: [
              dataCell("低評価率（1-4点）", 2800, { bold: true }),
              dataCell("0.0%", 2200, { alignment: AlignmentType.CENTER }),
              dataCell("0%維持", 2200, { alignment: AlignmentType.CENTER, color: GREEN_ACCENT, bold: true }),
              dataCell("2026年9月", 1826, { alignment: AlignmentType.CENTER }),
            ]}),
            new TableRow({ children: [
              dataCell("じゃらん平均評価", 2800, { bold: true, fill: LIGHT_GRAY }),
              dataCell("4.18/5点", 2200, { alignment: AlignmentType.CENTER, fill: LIGHT_GRAY }),
              dataCell("4.5/5点以上", 2200, { alignment: AlignmentType.CENTER, color: GREEN_ACCENT, bold: true, fill: LIGHT_GRAY }),
              dataCell("2026年9月", 1826, { alignment: AlignmentType.CENTER, fill: LIGHT_GRAY }),
            ]}),
            new TableRow({ children: [
              dataCell("楽天トラベル平均評価", 2800, { bold: true }),
              dataCell("4.44/5点", 2200, { alignment: AlignmentType.CENTER }),
              dataCell("4.6/5点以上", 2200, { alignment: AlignmentType.CENTER, color: GREEN_ACCENT, bold: true }),
              dataCell("2026年9月", 1826, { alignment: AlignmentType.CENTER }),
            ]}),
            new TableRow({ children: [
              dataCell("口コミ返信率", 2800, { bold: true, fill: LIGHT_GRAY }),
              dataCell("未計測", 2200, { alignment: AlignmentType.CENTER, fill: LIGHT_GRAY }),
              dataCell("100%（48h以内）", 2200, { alignment: AlignmentType.CENTER, color: GREEN_ACCENT, bold: true, fill: LIGHT_GRAY }),
              dataCell("2026年6月", 1826, { alignment: AlignmentType.CENTER, fill: LIGHT_GRAY }),
            ]}),
            new TableRow({ children: [
              dataCell("水回り関連クレーム", 2800, { bold: true }),
              dataCell("7件/2ヶ月", 2200, { alignment: AlignmentType.CENTER }),
              dataCell("2件以下/2ヶ月", 2200, { alignment: AlignmentType.CENTER, color: GREEN_ACCENT, bold: true }),
              dataCell("2026年7月", 1826, { alignment: AlignmentType.CENTER }),
            ]}),
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
                new Paragraph({ spacing: { after: 160 }, children: [new TextRun({ text: "ダイワロイネットホテル東京大崎は、10点換算で全体平均8.77点、高評価率84.5%、低評価率0.0%と非常に高い水準の評価を獲得しています。「立地」「清潔感」「スタッフの質」の三大基本要素が確固たる競争基盤を形成しています。", size: 21, font: "Arial", color: "333333" })] }),
                new Paragraph({ spacing: { after: 160 }, children: [new TextRun({ text: "全6サイトにおいて10pt換算で8点以上を達成しており、特にTrip.com（9.58）やGoogle（9.00）での評価は業界トップクラスです。インバウンド需要拡大を見据えた戦略的ポジショニングも極めて有利です。", size: 21, font: "Arial", color: "333333" })] }),
                new Paragraph({ spacing: { after: 160 }, children: [new TextRun({ text: "現在の高水準を更に向上させるため、「水回りの品質管理」「部屋の狭さへの対応」「エレベーター混雑」の3領域を重点改善課題として設定しました。これらは高評価をさらに伸ばすための施策であり、既存の強みを維持しつつ取り組むべきテーマです。", size: 21, font: "Arial", color: "333333" })] }),
                new Paragraph({ spacing: { after: 160 }, children: [new TextRun({ text: "提案したPhase 1の施策は投資を最小限に抑えつつ、オペレーション改善で着実な効果が期待できます。特に「水回り清掃の強化」と「周辺案内の充実」は、顧客満足度をさらに高めるレバレッジの高い施策です。", size: 21, font: "Arial", color: "333333" })] }),
                new Paragraph({ children: [new TextRun({ text: "Phase 1を着実に推進し、月次で口コミモニタリングを行いながら、全体平均9.0点以上・高評価率90%以上を目指してPhase 2・3を計画的に展開していくことを推奨いたします。", size: 21, font: "Arial", color: NAVY, bold: true })] }),
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
  fs.writeFileSync("/Users/mitsugutakahashi/ホテル口コミ/ダイワロイネットホテル東京大崎_口コミ分析改善レポート.docx", buffer);
  console.log("Report created successfully!");
  console.log("File size: " + (buffer.length / 1024).toFixed(1) + " KB");
});
