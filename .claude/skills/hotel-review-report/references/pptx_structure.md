# PPTX Presentation Structure Reference
## Hotel Review Analysis Report (10-Slide Template)

### 1. Presentation Setup

```js
pres.layout = "LAYOUT_16x9";   // 10" x 5.625"
pres.author = "Hotel Consulting";
pres.title  = "<Hotel Name> 口コミ分析改善レポート";
```

Uses `pptxgenjs` (CommonJS: `require("pptxgenjs")`).

---

### 2. Color Palette -- Midnight Executive

```js
const C = {
  navy: "1A2744",      navyLight: "243556",
  blue: "3B7DD8",      blueLight: "5A9BE6",
  ice: "E8EFF8",       white: "FFFFFF",      offWhite: "F5F7FA",
  gray: "64748B",      grayLight: "94A3B8",  grayDark: "334155",
  green: "16A34A",     greenBg: "DCFCE7",
  red: "DC2626",       redBg: "FEE2E2",
  orange: "EA580C",    orangeBg: "FFF7ED",
  gold: "D4A843",
};
```

---

### 3. Helper Functions

**shadow()** -- Standard drop shadow for cards/panels:
```js
{ type: "outer", blur: 4, offset: 2, angle: 135, color: "000000", opacity: 0.10 }
```

**addFooter(slide, pageNum)** -- Navy bar at bottom:
- Rectangle: x:0, y:5.25, w:10, h:0.375, fill navy
- Left text "Confidential": x:0.5, fontSize:8, color grayLight
- Right text page number: x:9, w:0.5, align right, fontSize:8

**addContentHeader(slide, title, subtitle)** -- Top header band:
- Rectangle: x:0, y:0, w:10, h:0.85, fill navy
- Title: x:0.6, y:0.08, fontSize:22, white, bold
- Subtitle (optional): x:0.6, y:0.5, fontSize:10, blueLight

**kpiCard(slide, x, y, w, h, label, value, color, bgColor)** -- Metric card:
- Background rectangle with shadow
- Thin accent bar: same x/y, full width, h:0.05
- Label: y+0.15, fontSize:10, gray, center
- Value: y+0.4, h:0.55, fontSize:32, bold, center

---

### 4. Slide-by-Slide Specification

#### Slide 1: Title
- **Background:** navy solid
- **Decorative overlay:** navyLight rectangle, full slide, transparency:40
- **Gold accent bar:** x:0, y:2.0, w:0.08, h:1.8, fill gold
- **Subtitle text** "口コミ分析": x:0.8, y:1.5, fontSize:20, blueLight, charSpacing:6
- **Main title** "改善レポート": x:0.8, y:2.1, fontSize:42, white, bold
- **Gold line:** x:0.8, y:3.1, w:2, LINE shape, gold, width:2
- **Hotel name:** x:0.8, y:3.3, fontSize:16, grayLight
- **Meta info** (period, count, date): x:0.8, y:4.1, fontSize:10, grayLight, paraSpaceAfter:4
- **Right decorative panel:** x:7.5, w:2.5, full height, blue, transparency:85
- **No footer on this slide**

#### Slide 2: Executive Summary
- **Background:** offWhite | **Header:** "エグゼクティブサマリー" / "Executive Summary"
- **4 KPI cards** (y:1.15, h:1.05):
  - x:0.5,  w:2.05 -- "全体平均(10pt換算)" / value / blue / white
  - x:2.75, w:2.05 -- "高評価率(8-10点)" / value / green / white
  - x:5.0,  w:2.05 -- "低評価率(1-4点)" / value / green / white
  - x:7.25, w:2.25 -- "レビュー総数" / value / navy / white
- **Strengths panel:** x:0.5, y:2.5, w:4.3, h:2.5, white card + green accent bar (h:0.05)
  - Title "Strengths": x:0.7, y:2.6, fontSize:14, green, bold
  - Content: x:0.7, y:3.0, w:3.8, h:1.8, mixed bold headers + gray detail, fontSize:11/9
- **Weaknesses panel:** x:5.2, y:2.5, w:4.3, h:2.5, white card + red accent bar
  - Title "Weaknesses": x:5.4, fontSize:14, red, bold
  - Content: x:5.4, y:3.0, w:3.8, h:1.8, same format as strengths
- Footer: page 2

#### Slide 3: Site-by-Site Ratings
- **Header:** "サイト別評価分析" / "Rating Analysis by Platform"
- **BAR chart** (horizontal): x:0.5, y:1.1, w:5.5, h:3.5
  - barDir:"bar", chartColors:[blue], showValue:true, dataLabelPosition:"outEnd"
  - valAxisMaxVal:10, valGridLine:{color:"E2E8F0",size:0.5}, catGridLine:{style:"none"}
  - chartArea: {fill:{color:white}, roundedCorners:true}
- **Insight box:** x:6.3, y:1.1, w:3.3, h:3.5, white card + gold accent bar
  - Title "Insight": fontSize:13, gold, bold
  - Content at x:6.5, y:1.55, w:2.9, mixed formatting
- **Data table** below chart: x:0.5, y:4.7, w:9
  - 5 columns: colW:[2.0, 1.0, 2.0, 1.0, 3.0]
  - Header row: navy fill, white text, fontSize:9, bold, center
  - Data rows: alternating ice/white fill, fontSize:9, center
  - Border: pt:0.5, color:"CBD5E1"
  - rowH: header 0.25, data 0.22 each
- Footer: page 3

#### Slide 4: Rating Distribution
- **Header:** "評価分布分析（10点換算）" / "Rating Distribution"
- **DOUGHNUT chart:** x:0.3, y:1.2, w:4.0, h:3.2
  - chartColors:[green, gold], showPercent:true, dataLabelColor:white
  - showLegend:true, legendPos:"b", legendFontSize:9, legendColor:gray
- **Column BAR chart:** x:4.5, y:1.2, w:5.2, h:3.2
  - barDir:"col", per-bar colors via chartColors array
  - chartArea white + roundedCorners, showValue:true, outEnd labels
- **3 summary cards** at y:4.55, h:0.55, w:2.9 each, spaced x = 0.5 + i*3.1
  - Each: colored bg rectangle, left label (bold, fontSize:10), right value (fontSize:11)
- Footer: page 4

#### Slide 5: Strengths
- **Header:** "強み分析" / "Strength Analysis"
- **3x2 card grid** (6 items): each card w:2.9, h:1.75
  - Position: x = 0.5 + col*3.1, y = 1.15 + row*2.0 (col 0-2, row 0-1)
  - White card with shadow + green left accent bar (w:0.06, full height)
  - **Count badge:** x+2.0, y+0.08, w:0.72, h:0.25, greenBg fill, green text, fontSize:9
  - **Theme title:** x+0.15, y+0.08, fontSize:13, navy, bold
  - **Description:** x+0.15, y+0.4, fontSize:9, gray
  - **Quote:** x+0.15, y+0.85, h:0.7, fontSize:9, blue, italic
- Footer: page 5

#### Slide 6: Weakness / Priority Matrix
- **Header:** "弱み分析・優先度マトリクス" / "Weakness Analysis & Priority Matrix"
- **Table:** x:0.5, y:1.15, w:9, colW:[0.8, 2.0, 5.0, 1.2]
  - Header: navy fill, white text, bold, fontSize:9
  - Priority column: fontSize:12, bold, color-coded (S=red, A=orange, B=gold, C=grayLight)
  - Data: alternating ice/white, fontSize:9
  - Border: pt:0.5, color:"CBD5E1", rowH: header 0.3, data 0.35 each
- **Legend** at y:4.8: inline text with colored priority letters + gray labels
- Footer: page 6

#### Slide 7: Phase 1 Improvements
- **Header:** "改善施策 Phase 1：即座対応" / "Immediate Actions (Today ~ 1 Month)"
- **Subheadline:** x:0.6, y:0.95, fontSize:10, blue, italic
- **2x2 card grid** (4 items): each w:4.35, h:1.75
  - Position: x = 0.5 + col*4.6, y = 1.35 + row*1.95 (col 0-1, row 0-1)
  - White card with shadow + blue left accent bar (w:0.06)
  - **Title:** x+0.15, y+0.05, fontSize:12, navy, bold
  - **Bullet list:** x+0.15, y+0.4, h:1.3, fontSize:9, grayDark, bullet:true, paraSpaceAfter:3
- Footer: page 7

#### Slide 8: Phase 2 & 3
- **Header:** "改善施策 Phase 2・3" / "Short-term & Mid-term Actions"
- **Left panel (Phase 2):** x:0.5, y:1.1, w:4.35, h:3.8
  - White card + blue header band (h:0.4, blue fill, white text, fontSize:12, bold)
  - Content items start at y:1.6, each has bold title (fontSize:10, navy) + bullet list (fontSize:9, gray)
  - Vertical spacing: y += 0.28 + items.length*0.22 + 0.2
- **Right panel (Phase 3):** x:5.15, y:1.1, w:4.35, h:3.8
  - Same structure as Phase 2 but header band is navy fill
  - Content starts at x:5.35
- Footer: page 8

#### Slide 9: KPI Targets
- **Header:** "KPI目標設定" / "Key Performance Indicators"
- **Table:** x:0.5, y:1.2, w:9, colW:[2.8, 2.0, 2.2, 2.0]
  - Header: navy fill, white text, bold, fontSize:10
  - Col 1: bold, grayDark (KPI name)
  - Col 2: center, gray (current value)
  - Col 3: center, bold, green (target value)
  - Col 4: center, gray (deadline)
  - Alternating ice/white, border pt:0.5 "CBD5E1", rowH:0.35 all
- **Note box:** x:0.5, y:4.25, w:9, h:0.7
  - White card + blue left accent bar (w:0.06)
  - Mixed text: bold navy label + gray description, fontSize:10, valign:middle
- Footer: page 9

#### Slide 10: Closing
- **Background:** navy solid
- **Decorative overlay:** navyLight, full slide, transparency:40
- **Section label** "総括": x:0.8, y:0.8, fontSize:14, blueLight, charSpacing:6
- **Transparent content box:** x:0.8, y:1.5, w:8.4, h:2.8, white fill, transparency:90
- **Summary text:** x:1.0, y:1.6, w:8, h:2.5, fontSize:12, white, paraSpaceAfter:4
  - Multiple paragraphs with empty-line spacers (fontSize:6 blank lines)
  - Final sentence: bold, gold color
- **Gold line:** x:0.8, y:4.5, w:2, LINE, gold, width:2
- **Closing text:** x:0.8, y:4.6, fontSize:14, grayLight
- **No footer on this slide**

---

### 5. Table Formatting Conventions

| Property | Value |
|----------|-------|
| Header fill | navy |
| Header text | white, bold, fontSize 9-10, Arial |
| Data rows | alternating ice (even index) / white (odd index) |
| Border | pt: 0.5, color: "CBD5E1" |
| Font | Arial, fontSize 9-10 |
| Alignment | "center" for numeric/short columns, left for text |
| Row heights | Header: 0.25-0.35, Data: 0.22-0.35 |

Cell formatting uses per-cell options objects:
```js
{ text: "value", options: { fontSize:9, fontFace:"Arial", align:"center",
  fill:{color: i%2===0 ? C.ice : C.white }, color: C.grayDark } }
```

---

### 6. Chart Configurations

**Horizontal BAR chart** (Slide 3):
```js
{ barDir:"bar", chartColors:[C.blue], showValue:true, dataLabelPosition:"outEnd",
  dataLabelColor:C.grayDark, showLegend:false, valAxisMaxVal:10,
  chartArea:{fill:{color:C.white}, roundedCorners:true},
  catAxisLabelColor:C.grayDark, catAxisLabelFontSize:10,
  valAxisLabelColor:C.gray, valAxisLabelFontSize:9,
  valGridLine:{color:"E2E8F0",size:0.5}, catGridLine:{style:"none"} }
```

**DOUGHNUT chart** (Slide 4):
```js
{ chartColors:[C.green, C.gold], showPercent:true, dataLabelColor:C.white,
  dataLabelFontSize:11, showTitle:false, showLegend:true, legendPos:"b",
  legendFontSize:9, legendColor:C.gray }
```

**Column BAR chart** (Slide 4):
```js
{ barDir:"col", chartColors:[array of per-bar colors],
  chartArea:{fill:{color:C.white}, roundedCorners:true},
  showValue:true, dataLabelPosition:"outEnd",
  dataLabelColor:C.grayDark, dataLabelFontSize:9, showLegend:false,
  catAxisLabelColor:C.grayDark, catAxisLabelFontSize:9,
  valAxisLabelColor:C.gray, valAxisLabelFontSize:8,
  valGridLine:{color:"E2E8F0",size:0.5}, catGridLine:{style:"none"} }
```

---

### 7. Key Patterns for New Hotels

1. **All coordinates are in inches** (10" x 5.625" canvas).
2. **Font:** Always Arial. Japanese text renders correctly.
3. **Accent bars:** Thin rectangles (w:0.05-0.06 or h:0.05) at card edges for color coding.
4. **Card pattern:** White fill + shadow() + colored accent bar at top or left edge.
5. **Text arrays:** Use `{ text, options: { breakLine, bold, fontSize, color } }` for mixed formatting.
6. **Empty-line spacers:** `{ text: "", options: { fontSize: 6, breakLine: true } }` between paragraphs.
7. **Grid layout formula:** `x = startX + col * spacing`, `y = startY + row * spacing`.
8. **Save:** `pres.writeFile({ fileName: path })` returns a Promise.
