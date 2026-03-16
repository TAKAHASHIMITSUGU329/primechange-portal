#!/usr/bin/env node
"use strict";

const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, PageNumber, PageBreak
} = require("docx");

const OUTPUT = path.resolve(__dirname, "../../納品レポート/PRIMECHANGE戦略レポート/PRIMECHANGE_打ち合わせアジェンダ_20260313.docx");

// ============================================================
// Styles
// ============================================================
const C = { NAVY: "1B3A5C", ACCENT: "2E75B6", WHITE: "FFFFFF", LIGHT_BG: "F5F7FA", TEXT: "333333", SUBTEXT: "666666", GREEN: "27AE60", ORANGE: "FF9800", RED: "E74C3C", BLUE: "2196F3" };

const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };
const noBorders = { top: { style: BorderStyle.NONE, size: 0 }, bottom: { style: BorderStyle.NONE, size: 0 }, left: { style: BorderStyle.NONE, size: 0 }, right: { style: BorderStyle.NONE, size: 0 } };
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
    children: Array.isArray(o.children) ? o.children : [new Paragraph({ alignment: o.a || AlignmentType.CENTER, children: [new TextRun({ text: String(t ?? "-"), size: o.sz || 16, font: "Arial", color: o.c || C.TEXT, bold: !!o.b })] })],
  });
}
const PB = () => new Paragraph({ children: [new PageBreak()] });

// ============================================================
// Meeting Info
// ============================================================
const MEETING = {
  title: "口コミ分析プロジェクト 報告・戦略打ち合わせ",
  date: "2026年3月13日（金）",
  time: "9:30 〜 10:30（60分）",
  location: "（オンラインまたは対面）",
  attendees: [
    { name: "春野 様", role: "株式会社PRIMECHANGE 代表取締役" },
    { name: "小林 様", role: "株式会社PRIMECHANGE" },
    { name: "高橋", role: "外部コンサルタント（分析・報告）" },
  ],
  purpose: "口コミ分析結果の報告および今後の改善アクション・進め方について合意を得る",
};

// ============================================================
// Agenda Items
// ============================================================
const AGENDA = [
  {
    no: 1, time: "9:30", duration: "5分", title: "ご挨拶・本日の目的",
    points: [
      "分析プロジェクトの概要確認（対象: 19ホテル、口コミ1,901件）",
      "本日のゴール: 分析報告 ＋ 今後の進め方の合意",
    ],
    materials: [],
  },
  {
    no: 2, time: "9:35", duration: "10分", title: "口コミ分析結果の全体像",
    points: [
      "19ホテル・1,901件の分析概要とスコア分布",
      "品質 × 売上 4象限マトリクス（ポートフォリオ分析結果）",
      "全体平均スコア: 8.04 / カテゴリ別傾向（清掃・接客・設備・コスパ）",
      "競合比較ポジショニング",
    ],
    materials: [
      "CS向上戦略提案書",
      "品質売上総合レポート",
    ],
  },
  {
    no: 3, time: "9:45", duration: "10分", title: "重点課題: URGENT 4ホテルの分析",
    points: [
      "コンフォートイン博多（7.75）— 清掃クレーム率最多、髪の毛・水回り",
      "コンフォートイン浜松町（7.00）— 全カテゴリ低評価、設備老朽化",
      "コンフォートイン蒲田（7.34）— 清掃ムラ、スタッフ対応ばらつき",
      "コンフォートイン新横浜（7.03）— 朝食・設備低評価、口コミ量少",
      "各ホテルの緊急対応案の提示",
    ],
    materials: [
      "ブレスト議事録（3/10実施分）",
      "ホテル別個別レポート（該当4ホテル分）",
    ],
  },
  {
    no: 4, time: "9:55", duration: "10分", title: "清掃品質改善戦略のご提案",
    points: [
      "クレーム分類分析: 髪の毛（38%）・水回り（25%）・ほこり/ゴミ（18%）等",
      "50項目清掃チェックリストの概要と運用案",
      "QC（品質管理）体制の構築提案",
      "3フェーズ改善ロードマップ（即時対応 → 短期 → 中期）",
    ],
    materials: [
      "清掃戦略レポート",
    ],
  },
  {
    no: 5, time: "10:05", duration: "5分", title: "投資対効果（ROI）",
    points: [
      "3シナリオ別投資回収試算",
      "  - 保守的: 年間約12M円改善",
      "  - 標準: 年間約36M円改善",
      "  - 積極的: 年間約60M円改善",
      "KPI目標設定（2026年9月時点）",
      "  - 全ホテル平均スコア8.5以上",
      "  - URGENT 4ホテル: 8.0以上",
    ],
    materials: [
      "品質売上総合レポート",
      "分析6（品質売上弾力性）",
    ],
  },
  {
    no: 6, time: "10:10", duration: "15分", title: "ディスカッション",
    points: [
      "優先施策の選定（どのホテルから着手するか）",
      "予算感のすり合わせ",
      "社内体制・推進担当の確認",
      "スケジュール感の共有（フェーズ1開始時期）",
    ],
    materials: [],
  },
  {
    no: 7, time: "10:25", duration: "5分", title: "まとめ・ネクストアクション",
    points: [
      "本日の決定事項の確認",
      "役割分担と次回アクション整理",
      "次回MTG日程の調整",
    ],
    materials: [],
  },
];

// ============================================================
// Build: Cover / Header Section
// ============================================================
function buildHeader() {
  return [
    new Paragraph({ spacing: { before: 600 } }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 120 }, children: [new TextRun({ text: "PRIMECHANGE", size: 48, font: "Arial", bold: true, color: C.NAVY })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: MEETING.title, size: 36, font: "Arial", bold: true, color: C.ACCENT })] }),

    // Meeting info table
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        ["日時", `${MEETING.date}  ${MEETING.time}`],
        ["場所", MEETING.location],
        ["目的", MEETING.purpose],
      ].map(([label, value]) =>
        new TableRow({
          children: [
            cell(label, { w: 1800, bg: C.NAVY, c: C.WHITE, b: true, sz: 18 }),
            cell(value, { a: AlignmentType.LEFT, sz: 18 }),
          ],
        })
      ),
    }),

    p("", { sa: 120 }),

    // Attendees
    h2("出席者"),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            cell("氏名", { bg: C.NAVY, c: C.WHITE, b: true, sz: 18, w: 2400 }),
            cell("所属・役割", { bg: C.NAVY, c: C.WHITE, b: true, sz: 18 }),
          ],
        }),
        ...MEETING.attendees.map((a, i) =>
          new TableRow({
            children: [
              cell(a.name, { b: true, sz: 18, bg: i % 2 === 1 ? C.LIGHT_BG : undefined }),
              cell(a.role, { a: AlignmentType.LEFT, sz: 18, bg: i % 2 === 1 ? C.LIGHT_BG : undefined }),
            ],
          })
        ),
      ],
    }),

    PB(),
  ];
}

// ============================================================
// Build: Agenda Table
// ============================================================
function buildAgendaTable() {
  const rows = [
    new TableRow({
      children: [
        cell("No.", { bg: C.NAVY, c: C.WHITE, b: true, sz: 16, w: 600 }),
        cell("時間", { bg: C.NAVY, c: C.WHITE, b: true, sz: 16, w: 1000 }),
        cell("議題", { bg: C.NAVY, c: C.WHITE, b: true, sz: 16, w: 3600 }),
        cell("所要", { bg: C.NAVY, c: C.WHITE, b: true, sz: 16, w: 900 }),
      ],
    }),
  ];

  AGENDA.forEach((item, i) => {
    const bgColor = i % 2 === 1 ? C.LIGHT_BG : undefined;
    rows.push(
      new TableRow({
        children: [
          cell(String(item.no), { sz: 16, bg: bgColor }),
          cell(item.time, { sz: 16, bg: bgColor }),
          cell(item.title, { a: AlignmentType.LEFT, sz: 16, b: true, bg: bgColor }),
          cell(item.duration, { sz: 16, bg: bgColor }),
        ],
      })
    );
  });

  return [
    h1("アジェンダ"),
    p("以下の流れで進行いたします。", { sa: 160 }),
    new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows }),
    p("", { sa: 200 }),
  ];
}

// ============================================================
// Build: Detail Sections
// ============================================================
function buildDetails() {
  const sections = [];
  sections.push(h1("各議題の詳細・論点"));

  AGENDA.forEach((item) => {
    sections.push(h2(`${item.no}. ${item.title}（${item.time}〜 / ${item.duration}）`));

    // Discussion points
    sections.push(h3("論点・内容"));
    item.points.forEach((pt) => {
      if (pt.startsWith("  - ")) {
        sections.push(bp(pt.trim().replace(/^- /, ""), { level: 1 }));
      } else {
        sections.push(bp(pt));
      }
    });

    // Materials
    if (item.materials.length > 0) {
      sections.push(h3("配布資料"));
      item.materials.forEach((m) => {
        sections.push(bp(m, { c: C.ACCENT }));
      });
    }

    sections.push(p("", { sa: 120 }));
  });

  return sections;
}

// ============================================================
// Build: Materials Summary
// ============================================================
function buildMaterialsSummary() {
  const allMaterials = new Map();
  AGENDA.forEach((item) => {
    item.materials.forEach((m) => {
      if (!allMaterials.has(m)) allMaterials.set(m, []);
      allMaterials.get(m).push(`議題${item.no}`);
    });
  });

  return [
    PB(),
    h1("配布資料一覧"),
    p("本打ち合わせで参照する資料の一覧です。", { sa: 160 }),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            cell("No.", { bg: C.NAVY, c: C.WHITE, b: true, sz: 16, w: 600 }),
            cell("資料名", { bg: C.NAVY, c: C.WHITE, b: true, sz: 16 }),
            cell("該当議題", { bg: C.NAVY, c: C.WHITE, b: true, sz: 16, w: 2000 }),
          ],
        }),
        ...[...allMaterials.entries()].map(([name, agendas], i) =>
          new TableRow({
            children: [
              cell(String(i + 1), { sz: 16, bg: i % 2 === 1 ? C.LIGHT_BG : undefined }),
              cell(name, { a: AlignmentType.LEFT, sz: 16, b: true, bg: i % 2 === 1 ? C.LIGHT_BG : undefined }),
              cell(agendas.join("、"), { sz: 16, bg: i % 2 === 1 ? C.LIGHT_BG : undefined }),
            ],
          })
        ),
      ],
    }),
    p(""),
    bp("ダッシュボードHP（デモ用）— 全体を通して適宜参照", { c: C.ACCENT }),
    p("", { sa: 200 }),
  ];
}

// ============================================================
// Build: Notes Section
// ============================================================
function buildNotes() {
  return [
    PB(),
    h1("メモ欄"),
    p("打ち合わせ中のメモにご利用ください。", { sa: 200, c: C.SUBTEXT, i: true }),
    // Empty lined area using a bordered table
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: Array.from({ length: 12 }, () =>
        new TableRow({
          height: { value: 500, rule: "atLeast" },
          children: [
            new TableCell({
              borders: {
                top: { style: BorderStyle.NONE, size: 0 },
                bottom: { style: BorderStyle.SINGLE, size: 1, color: "DDDDDD" },
                left: { style: BorderStyle.NONE, size: 0 },
                right: { style: BorderStyle.NONE, size: 0 },
              },
              children: [new Paragraph({ children: [new TextRun({ text: "", size: 20 })] })],
            }),
          ],
        })
      ),
    }),
  ];
}

// ============================================================
// Assemble Document
// ============================================================
async function main() {
  const doc = new Document({
    styles: {
      default: { document: { run: { font: "Arial", size: 21 } } },
      paragraphStyles: [
        { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", run: { size: 32, bold: true, font: "Arial", color: C.NAVY }, paragraph: { spacing: { before: 400, after: 200 } } },
        { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", run: { size: 26, bold: true, font: "Arial", color: C.ACCENT }, paragraph: { spacing: { before: 300, after: 160 } } },
        { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", run: { size: 22, bold: true, font: "Arial", color: C.NAVY }, paragraph: { spacing: { before: 200, after: 120 } } },
      ],
    },
    numbering: {
      config: [{ reference: "bullet-list", levels: [{ level: 0, format: "bullet", text: "\u2022", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }, { level: 1, format: "bullet", text: "\u25E6", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 1440, hanging: 360 } } } }] }],
    },
    sections: [
      {
        properties: {
          page: {
            size: { width: 11906, height: 16838 },
            margin: { top: 1134, bottom: 1134, left: 1418, right: 1418 },
          },
        },
        headers: {
          default: new Header({
            children: [new Paragraph({
              alignment: AlignmentType.RIGHT,
              border: { bottom: { style: BorderStyle.SINGLE, size: 2, color: C.ACCENT, space: 4 } },
              children: [new TextRun({ text: "PRIMECHANGE｜打ち合わせアジェンダ 2026.03.13", size: 16, font: "Arial", color: "999999" })],
            })],
          }),
        },
        footers: {
          default: new Footer({
            children: [new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [
                new TextRun({ text: "Confidential | Page ", size: 16, font: "Arial", color: "999999" }),
                new TextRun({ children: [PageNumber.CURRENT], size: 16, font: "Arial", color: "999999" }),
              ],
            })],
          }),
        },
        children: [
          ...buildHeader(),
          ...buildAgendaTable(),
          ...buildDetails(),
          ...buildMaterialsSummary(),
          ...buildNotes(),
        ],
      },
    ],
  });

  const buf = await Packer.toBuffer(doc);
  fs.writeFileSync(OUTPUT, buf);
  console.log(`✓ アジェンダ生成完了: ${OUTPUT}`);
}

main().catch((e) => { console.error(e); process.exit(1); });
