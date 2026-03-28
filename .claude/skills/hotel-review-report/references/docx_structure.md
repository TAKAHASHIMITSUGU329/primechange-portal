# Hotel Review DOCX Report Structure Reference

## 1. Dependencies and Imports

```js
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, PageNumber, PageBreak, LevelFormat } = require("docx");
```

## 2. Color Palette

| Constant       | Hex      | Usage                                      |
|----------------|----------|---------------------------------------------|
| NAVY           | `1B3A5C` | Heading1, heading3, KPI values, cover title |
| ACCENT         | `2E75B6` | Heading2, divider, header underline, KEY FINDINGS title |
| LIGHT_BLUE     | `D5E8F0` | Default KPI card background                |
| LIGHT_GRAY     | `F2F2F2` | Alternating table row shading              |
| WHITE          | `FFFFFF` | Header cell text, default backgrounds      |
| RED_ACCENT     | `C0392B` | Priority S, Weakness label                 |
| GREEN_ACCENT   | `27AE60` | Positive values, Strength label, high-rating cells |
| ORANGE_ACCENT  | `E67E22` | Priority A, Opportunity label, mid-rating cells |

Additional inline colors: `333333` (body text), `666666` (secondary text), `999999` (muted/footer), `AAAAAA` (confidential note), `888888` (priority C).

## 3. Border and Margin Constants

```js
const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };
const noBorders = { top/bottom/left/right: { style: BorderStyle.NONE, size: 0 } };
const cellMargins = { top: 80, bottom: 80, left: 120, right: 120 };
```

## 4. Helper Functions

### heading1(text)
- HeadingLevel.HEADING_1, bold, size 32, font Arial, color NAVY
- Spacing: before 360, after 200

### heading2(text)
- HeadingLevel.HEADING_2, bold, size 26, font Arial, color ACCENT
- Spacing: before 280, after 160

### heading3(text)
- No HeadingLevel (plain Paragraph), bold, size 22, font Arial, color NAVY
- Spacing: before 200, after 120

### para(text, opts={})
- Size 21, font Arial, color `333333` (default)
- Spacing: after `opts.afterSpacing || 120`, line 320
- Accepts: `alignment`, `color`, `bold`, any TextRun property via spread

### multiRunPara(runs[], opts={})
- Same spacing as para. Each run element is spread into TextRun defaults (size 21, Arial, `333333`).

### bulletItem(text, opts={})
- Uses numbering reference `"bullets"`, level `opts.level || 0`
- Size 21, font Arial, color `333333`
- Spacing: after 80, line 300

### headerCell(text, width)
- NAVY background, white bold text, size 20, centered
- Standard borders, cellMargins, verticalAlign center
- Width in DXA

### dataCell(text, width, opts={})
- Size 20, font Arial, color `opts.color || "333333"`, bold `opts.bold || false`
- Optional: `opts.fill` (shading), `opts.alignment` (default LEFT)
- Standard borders, cellMargins

### kpiRow(items[])
- Creates a single-row Table, total width 9360 DXA, columns split evenly
- Each item: `{ label, value, color?, bgColor? }`
- Label: size 18, color `666666`, centered
- Value: bold, size 36, centered, color defaults to NAVY
- Cell: no visible borders except right divider (`DDDDDD`), padding 160/200
- Default bgColor: LIGHT_BLUE

### priorityRow(priority, category, detail, impact)
- Returns a TableRow for the priority matrix (4 columns: 800 + 2200 + 4360 + 2000 DXA)
- Priority cell: centered, bold, size 22, color mapped from `{ S: RED, A: ORANGE, B: ACCENT, C: "888888" }`, LIGHT_GRAY bg
- Other cells: standard dataCell

### spacer(height=100)
- Empty paragraph with `spacing.after = height`

### divider()
- Empty paragraph with bottom border: ACCENT, size 6
- Spacing: before 200, after 200

## 5. Document Configuration

### Styles
```js
styles: {
  default: { document: { run: { font: "Arial", size: 21 } } },
  paragraphStyles: [
    { id: "Heading1", size: 32, bold, color: NAVY, spacing: before 360 / after 200, outlineLevel: 0 },
    { id: "Heading2", size: 26, bold, color: ACCENT, spacing: before 280 / after 160, outlineLevel: 1 },
  ]
}
```

### Numbering
```js
numbering.config: [
  { reference: "bullets", levels: [
    { level: 0, format: BULLET, text: "\u2022", indent: left 720, hanging 360 },
    { level: 1, format: BULLET, text: "\u25E6", indent: left 1440, hanging 360 },
  ]},
  { reference: "numbers", levels: [
    { level: 0, format: DECIMAL, text: "%1.", indent: left 720, hanging 360 },
  ]},
]
```

## 6. Document Sections

The document has exactly **2 sections**: Cover Page and Main Content.

### Page Properties (both sections)
- Size: width 11906, height 16838 (A4 in DXA twips)
- Margins: top/right/bottom/left all 1440 (1 inch)

---

## 7. Section 1: Cover Page

No header/footer. Structure top to bottom:

1. `spacer(2000)` -- push content down
2. Subtitle line: centered, size 56, color ACCENT, bottom border (size 8, ACCENT)
   - Text: "口コミ分析"
3. Title line: centered, size 56, bold, color NAVY
   - Text: "改善レポート"
4. `spacer(400)`
5. Hotel name: centered, size 32, bold, color NAVY
6. `spacer(200)`
7. Analysis period: centered, size 22, color `666666`
8. Review count: centered, size 22, color `666666`
9. Target sites: centered, size 20, color `666666`
10. `spacer(1600)`
11. Creation date: centered, size 20, color `999999`
12. Confidential notice: centered, size 18, color `AAAAAA`, italics

---

## 8. Section 2: Main Content

### Header
- Right-aligned, bottom border (ACCENT, size 2, space 4)
- Text: "{Hotel Name}｜口コミ分析改善レポート", size 16, color `999999`

### Footer
- Centered, "Page " + PageNumber.CURRENT, size 16, color `999999`

---

## 9. Chapter Details

### Ch1: Executive Summary (heading1 "1. エグゼクティブサマリー")

1. Introductory paragraph (para)
2. **kpiRow** with 4 items:
   - Overall avg (10pt), High-rating %, Low-rating %, Total reviews
   - Green bg (`E8F5E9`) for good metrics, LIGHT_BLUE for neutral
3. heading2 "総合評価" + 1-2 summary paragraphs
4. **KEY FINDINGS box**: single-cell Table (width 9026)
   - Left+top thick border (ACCENT, size 4), light bg `F0F7FC`
   - Title: "KEY FINDINGS" (bold, ACCENT)
   - Three lines: Strength (GREEN label), Weakness (RED label), Opportunity (ORANGE label)
   - Each line: bold colored label + normal text
5. PageBreak

### Ch2: Data Overview (heading1 "2. データ概要")

**2.1 Site-by-site table** (heading2):
- 7-column Table (width 9026): サイト名(1800), 件数(900), ネイティブ平均(1300), 尺度(900), 10pt換算(1300), 中央値(1526), 判定(1300)
- Header row with headerCell
- Data rows alternate LIGHT_GRAY fill
- Rating values use GREEN_ACCENT color + bold

**2.2 Distribution table** (heading2):
- 4-column Table (width 9026): 評価(1200), 件数(1200), 割合(1500), 分布(5126)
- Bar column uses unicode block chars (`\u2588`) with color-coded bars:
  - 8+ = GREEN, 5-7 = ORANGE, <5 = RED
  - Bar width = `Math.round(count / maxCount * 100)`, repeat = `Math.round(barWidth / 5)`
- Alternating LIGHT_GRAY fill on even ratings

**High/Mid/Low summary**: 3-column Table (width 9026, each ~3008)
- Green bg `E8F5E9` for High (8-10), Orange bg `FFF3E0` for Mid (5-7), Red bg `FDEDEC` for Low (1-4)
- Each cell: label (bold, size 20) + value (bold, size 24)
- PageBreak

### Ch3: Strength Analysis (heading1 "3. 強み分析（ポジティブ要因）")

**Theme table**: 3-column Table (width 9026): テーマ(2600), 言及数(1200), 代表コメント(5226)
- 6 theme rows, alternating fill (`E8F5E9` for top rows, LIGHT_GRAY for alternating)
- Count values: bold, GREEN_ACCENT, centered

**Subsections**: heading2 for each detail section (e.g., "3.1 最大の強み：...")
- Paragraphs + bulletItems for supporting details
- PageBreak

### Ch4: Weakness Analysis (heading1 "4. 弱み分析（改善課題）")

**Priority matrix table**: 4-column Table (width 9360): 優先度(800), 課題カテゴリ(2200), 具体的内容(4360), 影響度(2000)
- Uses `priorityRow(priority, category, detail, impact)` for each row
- Priority levels: S (red), A (orange), B (blue), C (gray)
- PageBreak

### Ch5: Improvement Plan (heading1 "5. 改善施策提案")

Three phases, each with heading2:
- **Phase 1** "即座対応（今週〜1ヶ月以内）" -- no/low investment, operations
- **Phase 2** "短期施策（1〜3ヶ月）" -- moderate investment
- **Phase 3** "中期施策（3〜6ヶ月）" -- capital investment

Each phase contains multiple initiatives:
- heading3 for initiative title: "(1) Initiative Name"
- bulletItem list for action items (3-4 items each)
- spacer(80) between initiatives
- PageBreak between Phase 1->Phase 2 boundary and after Phase 3

### Ch6: KPI Targets (heading1 "6. KPI目標設定")

**KPI table**: 4-column Table (width 9026): KPI項目(2800), 現状値(2200), 目標値(2200), 期限(1826)
- Alternating LIGHT_GRAY fill
- Target values: GREEN_ACCENT, bold
- Typical rows: overall avg, high-rating %, low-rating %, per-site averages, response rate, complaint counts
- PageBreak

### Ch7: Conclusion (heading1 "7. 総括と今後のアクション")

**Bordered summary box**: single-cell Table (width 9026)
- All 4 borders: NAVY, size 4
- Background: `F8F9FA`
- Margins: top/bottom 300, left/right 400
- Contains 4-5 Paragraphs (size 21, spacing after 160)
- Final paragraph: bold, color NAVY (call to action)

**Closing**:
- spacer(300)
- divider()
- spacer(100)
- Centered muted paragraph (contact note, color `999999`)

## 10. Output

```js
Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("output_path.docx", buffer);
});
```

## 11. Quick Reference: Table Widths

| Context              | Total Width | Column Widths                              |
|----------------------|-------------|---------------------------------------------|
| kpiRow               | 9360        | Evenly split by item count                  |
| Site table (Ch2)     | 9026        | 1800, 900, 1300, 900, 1300, 1526, 1300     |
| Distribution (Ch2)   | 9026        | 1200, 1200, 1500, 5126                     |
| High/Mid/Low (Ch2)   | 9026        | 3008, 3009, 3009                           |
| Strength (Ch3)       | 9026        | 2600, 1200, 5226                           |
| Priority matrix (Ch4)| 9360        | 800, 2200, 4360, 2000                      |
| KPI targets (Ch6)    | 9026        | 2800, 2200, 2200, 1826                     |
| KEY FINDINGS (Ch1)   | 9026        | 9026 (single column)                       |
| Conclusion (Ch7)     | 9026        | 9026 (single column)                       |

## 12. Adaptation Checklist

When creating a new hotel report, replace:
1. Hotel name (cover page, header, conclusion)
2. Analysis period and review count
3. Target site list
4. All KPI values in kpiRow (Ch1)
5. KEY FINDINGS strengths/weaknesses/opportunities text
6. Site table data rows (Ch2) -- adjust row count per hotel
7. Distribution data array and maxCount for bar scaling
8. High/Mid/Low summary counts and percentages
9. Strength theme table rows and subsection content (Ch3)
10. Priority matrix rows -- adjust count and S/A/B/C assignments (Ch4)
11. Phase 1/2/3 initiative headings and bullet items (Ch5)
12. KPI table rows with hotel-specific current/target values (Ch6)
13. Conclusion box paragraphs (Ch7)
14. Output filename
15. Creation date
