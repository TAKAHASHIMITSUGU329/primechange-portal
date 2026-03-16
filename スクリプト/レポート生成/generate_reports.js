const fs = require('fs');
const path = require('path');
const docx = require('docx');
const PptxGenJS = require('pptxgenjs');

// ============================================================
// 1. CSV PARSING
// ============================================================

function parseCSV(filePath) {
  const content = fs.readFileSync(filePath, 'utf-8');
  const lines = content.split('\n');

  const reviews = [];
  const seen = new Set();

  // Data starts from line index 6 (row 7 in 1-indexed)
  for (let i = 6; i < lines.length; i++) {
    const line = lines[i].trim();
    if (!line) continue;

    // Parse CSV respecting quoted fields
    const fields = parseCSVLine(line);

    // First set of columns (indices 0-11): check, site, rating, date, overall, good, improvement, translated_overall, translated_good, translated_improvement, reviewer
    // Second set repeats from index 12 onward (for the other month)

    // Process first set
    const sets = [
      { site: fields[2], rating: fields[3], date: fields[4], overall: fields[5], good: fields[6], improvement: fields[7], translatedOverall: fields[8], translatedGood: fields[9], translatedImprovement: fields[10], reviewer: fields[11] },
    ];

    // Process second set if exists
    if (fields.length > 14 && fields[14]) {
      sets.push({
        site: fields[14], rating: fields[15], date: fields[16], overall: fields[17], good: fields[18], improvement: fields[19], translatedOverall: fields[20], translatedGood: fields[21], translatedImprovement: fields[22], reviewer: fields[23]
      });
    }

    for (const s of sets) {
      if (!s.site || !s.date || !s.rating) continue;

      // Filter for 2026-02 and 2026-03
      if (!s.date.startsWith('2026-02') && !s.date.startsWith('2026-03')) continue;

      // Deduplicate using site+date+reviewer+rating
      const key = `${s.site}|${s.date}|${s.reviewer}|${s.rating}`;
      if (seen.has(key)) continue;
      seen.add(key);

      // Get comment text - prefer translated, fallback to original
      const comment = s.translatedOverall || s.overall || '';
      const goodComment = s.translatedGood || s.good || '';
      const improvementComment = s.translatedImprovement || s.improvement || '';

      reviews.push({
        site: s.site.trim(),
        rating: parseFloat(s.rating),
        date: s.date.trim(),
        comment: comment.trim(),
        goodComment: goodComment.trim(),
        improvementComment: improvementComment.trim(),
        reviewer: (s.reviewer || '').trim(),
      });
    }
  }

  return reviews;
}

function parseCSVLine(line) {
  const fields = [];
  let current = '';
  let inQuotes = false;

  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (inQuotes) {
      if (ch === '"') {
        if (i + 1 < line.length && line[i + 1] === '"') {
          current += '"';
          i++;
        } else {
          inQuotes = false;
        }
      } else {
        current += ch;
      }
    } else {
      if (ch === '"') {
        inQuotes = true;
      } else if (ch === ',') {
        fields.push(current);
        current = '';
      } else {
        current += ch;
      }
    }
  }
  fields.push(current);
  return fields;
}

// ============================================================
// 2. ANALYSIS
// ============================================================

function analyzeReviews(reviews) {
  // Rating scale normalization: Booking.com, Trip.com, Agoda use /10; Google, jalan, rakuten use /5
  const tenPointSites = ['Booking.com', 'Trip.com', 'Agoda'];
  const fivePointSites = ['Google', 'じゃらん', '楽天トラベル'];

  // Normalize all to /5 scale
  const normalized = reviews.map(r => {
    let normalizedRating = r.rating;
    if (tenPointSites.includes(r.site)) {
      normalizedRating = r.rating / 2;
    }
    return { ...r, normalizedRating };
  });

  // Overall stats
  const totalReviews = normalized.length;
  const avgRating = normalized.reduce((sum, r) => sum + r.normalizedRating, 0) / totalReviews;

  // By site
  const siteStats = {};
  for (const r of normalized) {
    if (!siteStats[r.site]) {
      siteStats[r.site] = { count: 0, totalRating: 0, totalOriginal: 0, reviews: [] };
    }
    siteStats[r.site].count++;
    siteStats[r.site].totalRating += r.normalizedRating;
    siteStats[r.site].totalOriginal += r.rating;
    siteStats[r.site].reviews.push(r);
  }

  for (const site in siteStats) {
    siteStats[site].avgRating = siteStats[site].totalRating / siteStats[site].count;
    siteStats[site].avgOriginal = siteStats[site].totalOriginal / siteStats[site].count;
    siteStats[site].scale = tenPointSites.includes(site) ? 10 : 5;
  }

  // By month
  const monthStats = {};
  for (const r of normalized) {
    const month = r.date.substring(0, 7);
    if (!monthStats[month]) {
      monthStats[month] = { count: 0, totalRating: 0, reviews: [] };
    }
    monthStats[month].count++;
    monthStats[month].totalRating += r.normalizedRating;
    monthStats[month].reviews.push(r);
  }
  for (const m in monthStats) {
    monthStats[m].avgRating = monthStats[m].totalRating / monthStats[m].count;
  }

  // Rating distribution (normalized to 5-point scale)
  const ratingDist = { '1': 0, '2': 0, '3': 0, '4': 0, '5': 0 };
  for (const r of normalized) {
    const rounded = Math.round(r.normalizedRating);
    const key = Math.max(1, Math.min(5, rounded)).toString();
    ratingDist[key]++;
  }

  // Text analysis - extract themes from comments
  const allComments = normalized.filter(r => r.comment || r.goodComment || r.improvementComment);

  // Keywords for positive themes
  const positiveThemes = {
    '立地・アクセス': ['駅', '立地', 'アクセス', '近く', '便利', '目の前', '徒歩', 'location', 'station', 'convenient'],
    '清潔感': ['清潔', '綺麗', 'きれい', 'クリーン', 'clean', '清掃'],
    'スタッフ対応': ['スタッフ', 'フロント', '対応', '親切', '丁寧', 'staff', 'courteous', 'helpful', 'サービス'],
    '朝食': ['朝食', '朝ごはん', 'breakfast', 'バイキング', 'breaky', '美味しい', '美味かった'],
    '部屋・設備': ['部屋', '広', 'ベッド', '快適', 'comfortable', 'room', '設備'],
    'コストパフォーマンス': ['コスパ', '安い', 'リーズナブル', '価格', '料金', '割安'],
    '眺望': ['眺望', '眺め', '景色', '夜景', 'view', '朝日'],
  };

  const negativeThemes = {
    '清掃不備': ['髪の毛', 'ゴミ', '不快', '汚', '清掃', '掃除'],
    '設備老朽化': ['古い', '蛇口', 'ツーハンドル', 'サーモ', '椅子', 'ギーギー', '水圧'],
    '朝食課題': ['冷めて', 'パサパサ', 'トースター', '7時', '6時半', '混雑', '料理名'],
    '部屋の狭さ': ['狭い', '狭く', '手狭', 'カーテン', 'ベッドが大き'],
    '周辺環境': ['荒', 'レストラン', '食事する場所', 'コンビニがない', 'どぶくさい', '臭い'],
    '騒音': ['音', '煩', '走行音', '高速道路'],
    'バスルーム問題': ['換気扇', '湿気', 'シャワーカーテン', '忽冷忽熱', '水温'],
  };

  function countTheme(themes, reviews) {
    const result = {};
    for (const [theme, keywords] of Object.entries(themes)) {
      let count = 0;
      const examples = [];
      for (const r of reviews) {
        const text = `${r.comment} ${r.goodComment} ${r.improvementComment}`;
        if (keywords.some(kw => text.includes(kw))) {
          count++;
          if (examples.length < 3 && text.trim().length > 5) {
            examples.push({ text: text.trim().substring(0, 100), reviewer: r.reviewer, site: r.site, date: r.date });
          }
        }
      }
      result[theme] = { count, examples };
    }
    return result;
  }

  const positiveAnalysis = countTheme(positiveThemes, allComments);
  const negativeAnalysis = countTheme(negativeThemes, allComments);

  // Low rating reviews (normalized <= 2.5)
  const lowRating = normalized.filter(r => r.normalizedRating <= 2.5 && (r.comment || r.improvementComment));

  // High rating reviews (normalized >= 4.5)
  const highRating = normalized.filter(r => r.normalizedRating >= 4.5 && (r.comment || r.goodComment));

  return {
    totalReviews,
    avgRating,
    siteStats,
    monthStats,
    ratingDist,
    positiveAnalysis,
    negativeAnalysis,
    lowRating,
    highRating,
    normalized,
  };
}

// ============================================================
// 3. GENERATE DOCX
// ============================================================

async function generateDocx(analysis, outputPath) {
  const {
    Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
    HeadingLevel, AlignmentType, WidthType, BorderStyle, ShadingType,
    Header, Footer, PageNumber, NumberFormat
  } = docx;

  const BLUE = '1F4E79';
  const LIGHT_BLUE = 'D6E4F0';
  const DARK_GRAY = '333333';
  const WHITE = 'FFFFFF';
  const GREEN = '2E7D32';
  const RED = 'C62828';
  const ORANGE = 'E65100';

  function createHeading(text, level = HeadingLevel.HEADING_1) {
    return new Paragraph({
      text: text,
      heading: level,
      spacing: { before: 300, after: 150 },
      run: { color: BLUE, bold: true },
    });
  }

  function createSubHeading(text) {
    return new Paragraph({
      text: text,
      heading: HeadingLevel.HEADING_2,
      spacing: { before: 250, after: 100 },
      run: { color: BLUE },
    });
  }

  function createBody(text) {
    return new Paragraph({
      children: [new TextRun({ text, size: 22, color: DARK_GRAY, font: 'Yu Gothic' })],
      spacing: { after: 100 },
    });
  }

  function createBullet(text, level = 0) {
    return new Paragraph({
      children: [new TextRun({ text, size: 22, color: DARK_GRAY, font: 'Yu Gothic' })],
      bullet: { level },
      spacing: { after: 50 },
    });
  }

  function createTableCell(text, opts = {}) {
    const { bold, color, bgColor, width, align } = opts;
    return new TableCell({
      children: [
        new Paragraph({
          children: [new TextRun({ text: String(text), size: 20, bold: !!bold, color: color || DARK_GRAY, font: 'Yu Gothic' })],
          alignment: align || AlignmentType.LEFT,
        }),
      ],
      width: width ? { size: width, type: WidthType.DXA } : undefined,
      shading: bgColor ? { fill: bgColor, type: ShadingType.CLEAR } : undefined,
      verticalAlign: 'center',
    });
  }

  // Build site stats table
  const siteTableRows = [];
  // Header row
  siteTableRows.push(new TableRow({
    children: [
      createTableCell('サイト名', { bold: true, color: WHITE, bgColor: BLUE }),
      createTableCell('件数', { bold: true, color: WHITE, bgColor: BLUE, align: AlignmentType.CENTER }),
      createTableCell('評価スケール', { bold: true, color: WHITE, bgColor: BLUE, align: AlignmentType.CENTER }),
      createTableCell('平均評価(原点)', { bold: true, color: WHITE, bgColor: BLUE, align: AlignmentType.CENTER }),
      createTableCell('平均評価(/5換算)', { bold: true, color: WHITE, bgColor: BLUE, align: AlignmentType.CENTER }),
    ],
  }));

  const sortedSites = Object.entries(analysis.siteStats).sort((a, b) => b[1].count - a[1].count);
  for (const [site, stats] of sortedSites) {
    siteTableRows.push(new TableRow({
      children: [
        createTableCell(site),
        createTableCell(stats.count.toString(), { align: AlignmentType.CENTER }),
        createTableCell(`${stats.scale}点満点`, { align: AlignmentType.CENTER }),
        createTableCell(stats.avgOriginal.toFixed(2), { align: AlignmentType.CENTER }),
        createTableCell(stats.avgRating.toFixed(2), { align: AlignmentType.CENTER }),
      ],
    }));
  }

  const siteTable = new Table({
    rows: siteTableRows,
    width: { size: 100, type: WidthType.PERCENTAGE },
  });

  // Build month stats table
  const monthTableRows = [];
  monthTableRows.push(new TableRow({
    children: [
      createTableCell('月', { bold: true, color: WHITE, bgColor: BLUE }),
      createTableCell('件数', { bold: true, color: WHITE, bgColor: BLUE, align: AlignmentType.CENTER }),
      createTableCell('平均評価(/5換算)', { bold: true, color: WHITE, bgColor: BLUE, align: AlignmentType.CENTER }),
    ],
  }));
  for (const [month, stats] of Object.entries(analysis.monthStats).sort()) {
    monthTableRows.push(new TableRow({
      children: [
        createTableCell(month),
        createTableCell(stats.count.toString(), { align: AlignmentType.CENTER }),
        createTableCell(stats.avgRating.toFixed(2), { align: AlignmentType.CENTER }),
      ],
    }));
  }
  const monthTable = new Table({
    rows: monthTableRows,
    width: { size: 100, type: WidthType.PERCENTAGE },
  });

  // Build rating distribution table
  const distTableRows = [];
  distTableRows.push(new TableRow({
    children: [
      createTableCell('評価(/5換算)', { bold: true, color: WHITE, bgColor: BLUE }),
      createTableCell('件数', { bold: true, color: WHITE, bgColor: BLUE, align: AlignmentType.CENTER }),
      createTableCell('割合', { bold: true, color: WHITE, bgColor: BLUE, align: AlignmentType.CENTER }),
    ],
  }));
  for (let i = 5; i >= 1; i--) {
    const count = analysis.ratingDist[i.toString()];
    const pct = ((count / analysis.totalReviews) * 100).toFixed(1);
    const stars = '\u2605'.repeat(i) + '\u2606'.repeat(5 - i);
    distTableRows.push(new TableRow({
      children: [
        createTableCell(`${stars}  (${i})`, { align: AlignmentType.LEFT }),
        createTableCell(count.toString(), { align: AlignmentType.CENTER }),
        createTableCell(`${pct}%`, { align: AlignmentType.CENTER }),
      ],
    }));
  }
  const distTable = new Table({
    rows: distTableRows,
    width: { size: 100, type: WidthType.PERCENTAGE },
  });

  // Positive themes table
  const posThemeRows = [];
  posThemeRows.push(new TableRow({
    children: [
      createTableCell('ポジティブテーマ', { bold: true, color: WHITE, bgColor: GREEN }),
      createTableCell('言及数', { bold: true, color: WHITE, bgColor: GREEN, align: AlignmentType.CENTER }),
    ],
  }));
  const sortedPositive = Object.entries(analysis.positiveAnalysis).sort((a, b) => b[1].count - a[1].count);
  for (const [theme, data] of sortedPositive) {
    posThemeRows.push(new TableRow({
      children: [
        createTableCell(theme),
        createTableCell(data.count.toString(), { align: AlignmentType.CENTER }),
      ],
    }));
  }
  const posThemeTable = new Table({
    rows: posThemeRows,
    width: { size: 100, type: WidthType.PERCENTAGE },
  });

  // Negative themes table
  const negThemeRows = [];
  negThemeRows.push(new TableRow({
    children: [
      createTableCell('ネガティブテーマ', { bold: true, color: WHITE, bgColor: RED }),
      createTableCell('言及数', { bold: true, color: WHITE, bgColor: RED, align: AlignmentType.CENTER }),
    ],
  }));
  const sortedNegative = Object.entries(analysis.negativeAnalysis).sort((a, b) => b[1].count - a[1].count);
  for (const [theme, data] of sortedNegative) {
    negThemeRows.push(new TableRow({
      children: [
        createTableCell(theme),
        createTableCell(data.count.toString(), { align: AlignmentType.CENTER }),
      ],
    }));
  }
  const negThemeTable = new Table({
    rows: negThemeRows,
    width: { size: 100, type: WidthType.PERCENTAGE },
  });

  // Improvement recommendations
  const improvements = [
    {
      title: '1. 清掃品質の強化',
      details: [
        '館内着やリネン類の毛髪チェックを清掃チェックリストに追加し、ダブルチェック体制を導入する。',
        '冷蔵庫内部、机下などの死角部分の清掃も標準手順に組み込む。',
        '清掃スタッフへの再教育プログラムを実施し、品質意識を向上させる。',
      ],
    },
    {
      title: '2. 朝食サービスの改善',
      details: [
        '朝食開始時間を6:30に前倒しし、ビジネス利用客の早朝ニーズに対応する。',
        'トースターの設置を検討し、パンの品質向上を図る。',
        '料理名の表示を全品に徹底し、外国人ゲストにも分かりやすいよう多言語表記を導入する。',
        '料理の温度管理を強化し、温かい料理は保温、冷たい料理は冷蔵を徹底する。',
      ],
    },
    {
      title: '3. 設備の更新・メンテナンス',
      details: [
        '浴室・洗面台の蛇口をサーモスタット式に順次交換し、温度調整の不便を解消する。',
        '空気清浄機のフィルター交換を定期的に実施し、臭いの問題を防止する。',
        '客室の椅子の整備（異音対策）を実施する。',
        'シャワーカーテンのサイズを見直し、浴室外への水漏れを防止する。',
        'トイレの水洗ノブなど経年劣化した部品を速やかに交換する。',
      ],
    },
    {
      title: '4. クレーム対応力の向上',
      details: [
        '電話でのクレーム対応マニュアルを整備し、まず謝罪と共感を示す対応を徹底する。',
        '交換品の提供時にも丁寧な説明と謝罪を心がけるよう指導する。',
        '低評価レビューへのフォローアップ体制を構築する。',
      ],
    },
    {
      title: '5. 周辺情報の充実',
      details: [
        '周辺の飲食店マップを作成し、チェックイン時に配布する。',
        'コンビニ・スーパーの場所と営業時間を案内に明記する。',
        '羽田空港へのアクセス方法（リムジンバス停留所の案内）をより分かりやすく提供する。',
      ],
    },
    {
      title: '6. 喫煙室の臭い対策',
      details: [
        '喫煙室の換気・消臭をより強化する。',
        '電子タバコ利用者向けの配慮として、消臭スプレーや空気清浄機の強化を検討する。',
      ],
    },
  ];

  // Build sections
  const sections = [];

  // --- Title Page Content ---
  const titlePage = [
    new Paragraph({ spacing: { before: 2000 } }),
    new Paragraph({
      children: [new TextRun({ text: 'ハートンホテル東品川', size: 56, bold: true, color: BLUE, font: 'Yu Gothic' })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 200 },
    }),
    new Paragraph({
      children: [new TextRun({ text: '口コミ分析レポート', size: 48, bold: true, color: BLUE, font: 'Yu Gothic' })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 400 },
    }),
    new Paragraph({
      children: [new TextRun({ text: '分析対象期間：2026年2月〜3月', size: 28, color: DARK_GRAY, font: 'Yu Gothic' })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 100 },
    }),
    new Paragraph({
      children: [new TextRun({ text: `作成日：${new Date().toLocaleDateString('ja-JP')}`, size: 24, color: DARK_GRAY, font: 'Yu Gothic' })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 100 },
    }),
    new Paragraph({
      children: [new TextRun({ text: `総口コミ数：${analysis.totalReviews}件 | 総合平均評価：${analysis.avgRating.toFixed(2)}/5.00`, size: 24, color: DARK_GRAY, font: 'Yu Gothic' })],
      alignment: AlignmentType.CENTER,
    }),
  ];

  // --- Executive Summary ---
  const execSummary = [
    createHeading('1. エグゼクティブサマリー'),
    createBody(`本レポートは、2026年2月から3月にかけてハートンホテル東品川に投稿された口コミ（全${analysis.totalReviews}件）の分析結果をまとめたものです。6つの予約・口コミサイト（Booking.com、楽天トラベル、じゃらん、Trip.com、Agoda、Google）から収集した口コミデータを基に、顧客満足度の現状と改善すべき課題を特定しました。`),
    createBody(`全サイトの評価を5点満点に統一換算した結果、総合平均評価は${analysis.avgRating.toFixed(2)}点でした。立地の良さ、スタッフの対応、清潔感が高く評価されている一方、清掃の不備（特に館内着への毛髪付着）、朝食の開始時間、設備の老朽化に関する改善要望が見られました。`),
  ];

  // --- Overall Analysis ---
  const overallSection = [
    createHeading('2. 全体分析'),
    createSubHeading('2.1 サイト別評価'),
    createBody('各口コミサイトの評価スケールは異なるため、全て5点満点に換算して比較しています。Booking.com、Trip.com、Agodaは10点満点、じゃらん、楽天トラベル、Googleは5点満点です。'),
    siteTable,
    new Paragraph({ spacing: { after: 200 } }),
    createSubHeading('2.2 月別推移'),
    monthTable,
    new Paragraph({ spacing: { after: 200 } }),
    createSubHeading('2.3 評価分布'),
    distTable,
  ];

  // --- Positive Analysis ---
  const positiveSection = [
    createHeading('3. ポジティブ分析（強み）'),
    createBody('口コミ内容をテーマ別に分析した結果、以下の項目が頻繁にポジティブな言及を受けています。'),
    posThemeTable,
    new Paragraph({ spacing: { after: 200 } }),
  ];

  // Add examples for top positive themes
  for (const [theme, data] of sortedPositive.slice(0, 4)) {
    if (data.examples.length > 0) {
      positiveSection.push(createBody(`【${theme}】の口コミ例：`));
      for (const ex of data.examples.slice(0, 2)) {
        positiveSection.push(createBullet(`「${ex.text}...」(${ex.site}, ${ex.date})`));
      }
      positiveSection.push(new Paragraph({ spacing: { after: 100 } }));
    }
  }

  // --- Negative Analysis ---
  const negativeSection = [
    createHeading('4. ネガティブ分析（課題）'),
    createBody('口コミ内容から抽出された改善課題のテーマと言及数は以下の通りです。'),
    negThemeTable,
    new Paragraph({ spacing: { after: 200 } }),
  ];

  for (const [theme, data] of sortedNegative.slice(0, 4)) {
    if (data.examples.length > 0) {
      negativeSection.push(createBody(`【${theme}】の口コミ例：`));
      for (const ex of data.examples.slice(0, 2)) {
        negativeSection.push(createBullet(`「${ex.text}...」(${ex.site}, ${ex.date})`));
      }
      negativeSection.push(new Paragraph({ spacing: { after: 100 } }));
    }
  }

  // --- Low Rating Details ---
  const lowRatingSection = [
    createHeading('5. 低評価レビュー詳細'),
    createBody(`5点満点換算で2.5点以下のレビューは${analysis.lowRating.length}件でした。以下に主要な低評価レビューの内容を記載します。`),
  ];

  for (const r of analysis.lowRating.slice(0, 5)) {
    const commentText = r.comment || r.improvementComment || '（コメントなし）';
    lowRatingSection.push(createBody(`[${r.site}] ${r.date} - 評価: ${r.rating}点 (投稿者: ${r.reviewer || '不明'})`));
    lowRatingSection.push(createBullet(commentText.substring(0, 200)));
    lowRatingSection.push(new Paragraph({ spacing: { after: 100 } }));
  }

  // --- Improvements ---
  const improvementSection = [
    createHeading('6. 改善提案'),
    createBody('口コミ分析の結果に基づき、以下の改善施策を提案します。'),
  ];

  for (const imp of improvements) {
    improvementSection.push(createSubHeading(imp.title));
    for (const detail of imp.details) {
      improvementSection.push(createBullet(detail));
    }
  }

  // --- Conclusion ---
  const conclusionSection = [
    createHeading('7. まとめ'),
    createBody(`ハートンホテル東品川は、品川シーサイド駅直結の抜群の立地、リーズナブルな価格設定、親切なスタッフ対応が多くのゲストから高く評価されています。総合平均評価${analysis.avgRating.toFixed(2)}/5.00は良好な水準と言えます。`),
    createBody('特に、駅からの近さ、イオンやコンビニなど周辺施設の充実度、羽田空港やビッグサイトへのアクセスの良さが繰り返し言及されており、ビジネス利用・イベント利用の両面で高い需要があります。'),
    createBody('一方、清掃品質の一貫性、朝食サービスの改善、設備の更新が主な課題として浮上しています。特に清掃に関する問題は低評価に直結しやすく、最優先で取り組むべき事項です。'),
    createBody('上記の改善提案を段階的に実施することで、顧客満足度の更なる向上とリピーター獲得が期待できます。'),
  ];

  const doc = new Document({
    sections: [{
      properties: {
        page: {
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
        },
      },
      headers: {
        default: new Header({
          children: [new Paragraph({
            children: [new TextRun({ text: 'ハートンホテル東品川 口コミ分析レポート 2026年2月-3月', size: 16, color: '999999', font: 'Yu Gothic' })],
            alignment: AlignmentType.RIGHT,
          })],
        }),
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({ children: [PageNumber.CURRENT], size: 18, color: '999999' }),
              new TextRun({ text: ' / ', size: 18, color: '999999' }),
              new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 18, color: '999999' }),
            ],
          })],
        }),
      },
      children: [
        ...titlePage,
        ...execSummary,
        ...overallSection,
        ...positiveSection,
        ...negativeSection,
        ...lowRatingSection,
        ...improvementSection,
        ...conclusionSection,
      ],
    }],
  });

  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(outputPath, buffer);
  console.log(`DOCX saved to: ${outputPath}`);
}

// ============================================================
// 4. GENERATE PPTX
// ============================================================

async function generatePptx(analysis, outputPath) {
  const pptx = new PptxGenJS();

  pptx.title = 'ハートンホテル東品川 口コミ分析レポート';
  pptx.subject = '2026年2月-3月 口コミ分析';
  pptx.author = '口コミ分析チーム';

  pptx.defineSlideMaster({
    title: 'MASTER',
    background: { color: 'FFFFFF' },
    objects: [
      { rect: { x: 0, y: 0, w: '100%', h: 0.6, fill: { color: '1F4E79' } } },
      { text: { text: 'ハートンホテル東品川 口コミ分析', options: { x: 0.5, y: 0.1, w: 8, h: 0.4, fontSize: 12, color: 'FFFFFF', fontFace: 'Yu Gothic' } } },
      { rect: { x: 0, y: '95%', w: '100%', h: '5%', fill: { color: 'D6E4F0' } } },
    ],
  });

  // ---- Slide 1: Title ----
  const slide1 = pptx.addSlide();
  slide1.background = { color: '1F4E79' };
  slide1.addText('ハートンホテル東品川', { x: 0.5, y: 1.5, w: 9, h: 1, fontSize: 36, bold: true, color: 'FFFFFF', fontFace: 'Yu Gothic', align: 'center' });
  slide1.addText('口コミ分析レポート', { x: 0.5, y: 2.5, w: 9, h: 0.8, fontSize: 28, bold: true, color: 'D6E4F0', fontFace: 'Yu Gothic', align: 'center' });
  slide1.addText('分析対象期間：2026年2月〜3月', { x: 0.5, y: 3.5, w: 9, h: 0.5, fontSize: 18, color: 'FFFFFF', fontFace: 'Yu Gothic', align: 'center' });
  slide1.addText(`総口コミ数：${analysis.totalReviews}件  |  総合平均評価：${analysis.avgRating.toFixed(2)} / 5.00`, { x: 0.5, y: 4.2, w: 9, h: 0.5, fontSize: 16, color: 'D6E4F0', fontFace: 'Yu Gothic', align: 'center' });

  // ---- Slide 2: Executive Summary ----
  const slide2 = pptx.addSlide({ masterName: 'MASTER' });
  slide2.addText('エグゼクティブサマリー', { x: 0.5, y: 0.8, w: 9, h: 0.5, fontSize: 24, bold: true, color: '1F4E79', fontFace: 'Yu Gothic' });

  const summaryBullets = [
    { text: `分析対象：6サイトから${analysis.totalReviews}件の口コミ（2026年2月-3月）`, options: { fontSize: 14, color: '333333', fontFace: 'Yu Gothic', bullet: true, breakLine: true } },
    { text: `総合平均評価：${analysis.avgRating.toFixed(2)} / 5.00（5点満点換算）`, options: { fontSize: 14, color: '333333', fontFace: 'Yu Gothic', bullet: true, breakLine: true } },
    { text: '強み：立地（駅直結）、清潔感、スタッフ対応、コストパフォーマンス', options: { fontSize: 14, color: '2E7D32', fontFace: 'Yu Gothic', bullet: true, breakLine: true } },
    { text: '課題：清掃不備、朝食改善、設備老朽化、周辺環境', options: { fontSize: 14, color: 'C62828', fontFace: 'Yu Gothic', bullet: true, breakLine: true } },
    { text: '特記事項：清掃問題は低評価レビューに直結。優先的な改善が必要。', options: { fontSize: 14, color: 'E65100', fontFace: 'Yu Gothic', bullet: true, breakLine: true } },
  ];
  slide2.addText(summaryBullets, { x: 0.5, y: 1.5, w: 9, h: 4, valign: 'top' });

  // ---- Slide 3: Site Stats ----
  const slide3 = pptx.addSlide({ masterName: 'MASTER' });
  slide3.addText('サイト別評価一覧', { x: 0.5, y: 0.8, w: 9, h: 0.5, fontSize: 24, bold: true, color: '1F4E79', fontFace: 'Yu Gothic' });
  slide3.addText('※ Booking.com, Trip.com, Agodaは10点満点 / じゃらん, 楽天トラベル, Googleは5点満点', { x: 0.5, y: 1.3, w: 9, h: 0.3, fontSize: 10, color: '666666', fontFace: 'Yu Gothic' });

  const siteTableData = [
    [
      { text: 'サイト名', options: { bold: true, color: 'FFFFFF', fill: { color: '1F4E79' }, fontSize: 12, fontFace: 'Yu Gothic', align: 'center' } },
      { text: '件数', options: { bold: true, color: 'FFFFFF', fill: { color: '1F4E79' }, fontSize: 12, fontFace: 'Yu Gothic', align: 'center' } },
      { text: 'スケール', options: { bold: true, color: 'FFFFFF', fill: { color: '1F4E79' }, fontSize: 12, fontFace: 'Yu Gothic', align: 'center' } },
      { text: '平均(原点)', options: { bold: true, color: 'FFFFFF', fill: { color: '1F4E79' }, fontSize: 12, fontFace: 'Yu Gothic', align: 'center' } },
      { text: '平均(/5換算)', options: { bold: true, color: 'FFFFFF', fill: { color: '1F4E79' }, fontSize: 12, fontFace: 'Yu Gothic', align: 'center' } },
    ],
  ];

  const sortedSites2 = Object.entries(analysis.siteStats).sort((a, b) => b[1].count - a[1].count);
  for (const [site, stats] of sortedSites2) {
    const ratingColor = stats.avgRating >= 4.0 ? '2E7D32' : stats.avgRating >= 3.5 ? 'E65100' : 'C62828';
    siteTableData.push([
      { text: site, options: { fontSize: 11, fontFace: 'Yu Gothic', align: 'center' } },
      { text: stats.count.toString(), options: { fontSize: 11, fontFace: 'Yu Gothic', align: 'center' } },
      { text: `/${stats.scale}`, options: { fontSize: 11, fontFace: 'Yu Gothic', align: 'center' } },
      { text: stats.avgOriginal.toFixed(2), options: { fontSize: 11, fontFace: 'Yu Gothic', align: 'center' } },
      { text: stats.avgRating.toFixed(2), options: { fontSize: 11, bold: true, color: ratingColor, fontFace: 'Yu Gothic', align: 'center' } },
    ]);
  }

  slide3.addTable(siteTableData, {
    x: 0.5, y: 1.7, w: 9,
    border: { type: 'solid', pt: 0.5, color: 'CCCCCC' },
    rowH: [0.4, ...Array(sortedSites2.length).fill(0.35)],
    colW: [2.2, 1.2, 1.4, 1.8, 2.4],
    autoPage: false,
  });

  // ---- Slide 4: Monthly Trends & Distribution ----
  const slide4 = pptx.addSlide({ masterName: 'MASTER' });
  slide4.addText('月別評価推移と評価分布', { x: 0.5, y: 0.8, w: 9, h: 0.5, fontSize: 24, bold: true, color: '1F4E79', fontFace: 'Yu Gothic' });

  const monthTableData = [
    [
      { text: '月', options: { bold: true, color: 'FFFFFF', fill: { color: '1F4E79' }, fontSize: 12, fontFace: 'Yu Gothic', align: 'center' } },
      { text: '件数', options: { bold: true, color: 'FFFFFF', fill: { color: '1F4E79' }, fontSize: 12, fontFace: 'Yu Gothic', align: 'center' } },
      { text: '平均評価(/5換算)', options: { bold: true, color: 'FFFFFF', fill: { color: '1F4E79' }, fontSize: 12, fontFace: 'Yu Gothic', align: 'center' } },
    ],
  ];

  for (const [month, stats] of Object.entries(analysis.monthStats).sort()) {
    const ratingColor = stats.avgRating >= 4.0 ? '2E7D32' : stats.avgRating >= 3.5 ? 'E65100' : 'C62828';
    monthTableData.push([
      { text: month, options: { fontSize: 12, fontFace: 'Yu Gothic', align: 'center' } },
      { text: stats.count.toString(), options: { fontSize: 12, fontFace: 'Yu Gothic', align: 'center' } },
      { text: stats.avgRating.toFixed(2), options: { fontSize: 12, bold: true, color: ratingColor, fontFace: 'Yu Gothic', align: 'center' } },
    ]);
  }

  slide4.addTable(monthTableData, {
    x: 0.5, y: 1.5, w: 4,
    border: { type: 'solid', pt: 0.5, color: 'CCCCCC' },
    rowH: [0.4, 0.35, 0.35],
    colW: [1.5, 1, 1.5],
    autoPage: false,
  });

  // Rating distribution on right side
  slide4.addText('評価分布（5点満点換算）', { x: 5, y: 1.3, w: 5, h: 0.3, fontSize: 14, bold: true, color: '1F4E79', fontFace: 'Yu Gothic' });

  const distTableData2 = [
    [
      { text: '評価', options: { bold: true, color: 'FFFFFF', fill: { color: '1F4E79' }, fontSize: 10, fontFace: 'Yu Gothic', align: 'center' } },
      { text: '件数', options: { bold: true, color: 'FFFFFF', fill: { color: '1F4E79' }, fontSize: 10, fontFace: 'Yu Gothic', align: 'center' } },
      { text: '割合', options: { bold: true, color: 'FFFFFF', fill: { color: '1F4E79' }, fontSize: 10, fontFace: 'Yu Gothic', align: 'center' } },
    ],
  ];

  for (let i = 5; i >= 1; i--) {
    const count = analysis.ratingDist[i.toString()];
    const pct = ((count / analysis.totalReviews) * 100).toFixed(1);
    const colors = ['C62828', 'D84315', 'E65100', '558B2F', '2E7D32'];
    distTableData2.push([
      { text: `${i}点`, options: { fontSize: 11, fontFace: 'Yu Gothic', align: 'center', bold: true } },
      { text: count.toString(), options: { fontSize: 11, fontFace: 'Yu Gothic', align: 'center', color: colors[i - 1] } },
      { text: `${pct}%`, options: { fontSize: 11, fontFace: 'Yu Gothic', align: 'center' } },
    ]);
  }

  slide4.addTable(distTableData2, {
    x: 5, y: 1.7, w: 4.5,
    border: { type: 'solid', pt: 0.5, color: 'CCCCCC' },
    rowH: [0.35, 0.3, 0.3, 0.3, 0.3, 0.3],
    colW: [1.5, 1.5, 1.5],
    autoPage: false,
  });

  // ---- Slide 5: Positive Themes ----
  const slide5 = pptx.addSlide({ masterName: 'MASTER' });
  slide5.addText('ポジティブ分析（強み）', { x: 0.5, y: 0.8, w: 9, h: 0.5, fontSize: 24, bold: true, color: '2E7D32', fontFace: 'Yu Gothic' });

  const sortedPos = Object.entries(analysis.positiveAnalysis).sort((a, b) => b[1].count - a[1].count);
  const posTableData = [
    [
      { text: 'テーマ', options: { bold: true, color: 'FFFFFF', fill: { color: '2E7D32' }, fontSize: 12, fontFace: 'Yu Gothic', align: 'center' } },
      { text: '言及数', options: { bold: true, color: 'FFFFFF', fill: { color: '2E7D32' }, fontSize: 12, fontFace: 'Yu Gothic', align: 'center' } },
      { text: '代表的なコメント', options: { bold: true, color: 'FFFFFF', fill: { color: '2E7D32' }, fontSize: 12, fontFace: 'Yu Gothic', align: 'center' } },
    ],
  ];

  for (const [theme, data] of sortedPos) {
    const example = data.examples.length > 0 ? data.examples[0].text.substring(0, 50) + '...' : '-';
    posTableData.push([
      { text: theme, options: { fontSize: 11, fontFace: 'Yu Gothic' } },
      { text: data.count.toString(), options: { fontSize: 11, fontFace: 'Yu Gothic', align: 'center', bold: true } },
      { text: example, options: { fontSize: 9, fontFace: 'Yu Gothic' } },
    ]);
  }

  slide5.addTable(posTableData, {
    x: 0.3, y: 1.5, w: 9.4,
    border: { type: 'solid', pt: 0.5, color: 'CCCCCC' },
    rowH: [0.35, ...Array(sortedPos.length).fill(0.45)],
    colW: [2.2, 1.2, 6],
    autoPage: false,
  });

  // ---- Slide 6: Negative Themes ----
  const slide6 = pptx.addSlide({ masterName: 'MASTER' });
  slide6.addText('ネガティブ分析（課題）', { x: 0.5, y: 0.8, w: 9, h: 0.5, fontSize: 24, bold: true, color: 'C62828', fontFace: 'Yu Gothic' });

  const sortedNeg = Object.entries(analysis.negativeAnalysis).sort((a, b) => b[1].count - a[1].count);
  const negTableData = [
    [
      { text: 'テーマ', options: { bold: true, color: 'FFFFFF', fill: { color: 'C62828' }, fontSize: 12, fontFace: 'Yu Gothic', align: 'center' } },
      { text: '言及数', options: { bold: true, color: 'FFFFFF', fill: { color: 'C62828' }, fontSize: 12, fontFace: 'Yu Gothic', align: 'center' } },
      { text: '代表的なコメント', options: { bold: true, color: 'FFFFFF', fill: { color: 'C62828' }, fontSize: 12, fontFace: 'Yu Gothic', align: 'center' } },
    ],
  ];

  for (const [theme, data] of sortedNeg) {
    const example = data.examples.length > 0 ? data.examples[0].text.substring(0, 50) + '...' : '-';
    negTableData.push([
      { text: theme, options: { fontSize: 11, fontFace: 'Yu Gothic' } },
      { text: data.count.toString(), options: { fontSize: 11, fontFace: 'Yu Gothic', align: 'center', bold: true } },
      { text: example, options: { fontSize: 9, fontFace: 'Yu Gothic' } },
    ]);
  }

  slide6.addTable(negTableData, {
    x: 0.3, y: 1.5, w: 9.4,
    border: { type: 'solid', pt: 0.5, color: 'CCCCCC' },
    rowH: [0.35, ...Array(sortedNeg.length).fill(0.45)],
    colW: [2.2, 1.2, 6],
    autoPage: false,
  });

  // ---- Slide 7: Improvement Recommendations (1) ----
  const slide7 = pptx.addSlide({ masterName: 'MASTER' });
  slide7.addText('改善提案（1/2）', { x: 0.5, y: 0.8, w: 9, h: 0.5, fontSize: 24, bold: true, color: '1F4E79', fontFace: 'Yu Gothic' });

  const imp1 = [
    { text: '1. 清掃品質の強化', options: { fontSize: 14, bold: true, color: 'C62828', fontFace: 'Yu Gothic', breakLine: true } },
    { text: '  - 館内着・リネン類の毛髪チェックをダブルチェック体制に', options: { fontSize: 11, color: '333333', fontFace: 'Yu Gothic', breakLine: true } },
    { text: '  - 冷蔵庫内部・机下など死角部分の清掃を標準手順に追加', options: { fontSize: 11, color: '333333', fontFace: 'Yu Gothic', breakLine: true } },
    { text: '', options: { fontSize: 8, breakLine: true } },
    { text: '2. 朝食サービスの改善', options: { fontSize: 14, bold: true, color: 'E65100', fontFace: 'Yu Gothic', breakLine: true } },
    { text: '  - 開始時間を6:30に前倒し（混雑緩和・ビジネス客対応）', options: { fontSize: 11, color: '333333', fontFace: 'Yu Gothic', breakLine: true } },
    { text: '  - トースター設置、料理名の多言語表示、温度管理の強化', options: { fontSize: 11, color: '333333', fontFace: 'Yu Gothic', breakLine: true } },
    { text: '', options: { fontSize: 8, breakLine: true } },
    { text: '3. 設備の更新・メンテナンス', options: { fontSize: 14, bold: true, color: 'E65100', fontFace: 'Yu Gothic', breakLine: true } },
    { text: '  - 蛇口をサーモスタット式に順次交換', options: { fontSize: 11, color: '333333', fontFace: 'Yu Gothic', breakLine: true } },
    { text: '  - 空気清浄機フィルター定期交換、椅子の異音対策', options: { fontSize: 11, color: '333333', fontFace: 'Yu Gothic', breakLine: true } },
    { text: '  - シャワーカーテンのサイズ見直し', options: { fontSize: 11, color: '333333', fontFace: 'Yu Gothic', breakLine: true } },
  ];
  slide7.addText(imp1, { x: 0.5, y: 1.5, w: 9, h: 4, valign: 'top' });

  // ---- Slide 8: Improvement Recommendations (2) ----
  const slide8 = pptx.addSlide({ masterName: 'MASTER' });
  slide8.addText('改善提案（2/2）', { x: 0.5, y: 0.8, w: 9, h: 0.5, fontSize: 24, bold: true, color: '1F4E79', fontFace: 'Yu Gothic' });

  const imp2 = [
    { text: '4. クレーム対応力の向上', options: { fontSize: 14, bold: true, color: 'E65100', fontFace: 'Yu Gothic', breakLine: true } },
    { text: '  - 電話クレーム対応マニュアルの整備（謝罪と共感を最優先）', options: { fontSize: 11, color: '333333', fontFace: 'Yu Gothic', breakLine: true } },
    { text: '  - 低評価レビューへのフォローアップ体制の構築', options: { fontSize: 11, color: '333333', fontFace: 'Yu Gothic', breakLine: true } },
    { text: '', options: { fontSize: 8, breakLine: true } },
    { text: '5. 周辺情報の充実', options: { fontSize: 14, bold: true, color: '1F4E79', fontFace: 'Yu Gothic', breakLine: true } },
    { text: '  - 飲食店マップの作成・配布', options: { fontSize: 11, color: '333333', fontFace: 'Yu Gothic', breakLine: true } },
    { text: '  - 羽田空港リムジンバス停留所の案内を強化', options: { fontSize: 11, color: '333333', fontFace: 'Yu Gothic', breakLine: true } },
    { text: '', options: { fontSize: 8, breakLine: true } },
    { text: '6. 喫煙室の臭い対策', options: { fontSize: 14, bold: true, color: '1F4E79', fontFace: 'Yu Gothic', breakLine: true } },
    { text: '  - 換気・消臭の強化、空気清浄機の性能向上', options: { fontSize: 11, color: '333333', fontFace: 'Yu Gothic', breakLine: true } },
  ];
  slide8.addText(imp2, { x: 0.5, y: 1.5, w: 9, h: 4, valign: 'top' });

  // ---- Slide 9: Key Low-Rating Reviews ----
  const slide9 = pptx.addSlide({ masterName: 'MASTER' });
  slide9.addText('注目すべき低評価レビュー', { x: 0.5, y: 0.8, w: 9, h: 0.5, fontSize: 24, bold: true, color: 'C62828', fontFace: 'Yu Gothic' });

  const lowReviewTexts = [];
  for (const r of analysis.lowRating.slice(0, 3)) {
    const commentText = (r.comment || r.improvementComment || '').substring(0, 80);
    lowReviewTexts.push(
      { text: `[${r.site}] ${r.date} - 評価: ${r.rating}点`, options: { fontSize: 12, bold: true, color: 'C62828', fontFace: 'Yu Gothic', breakLine: true } },
      { text: `${commentText}...`, options: { fontSize: 10, color: '333333', fontFace: 'Yu Gothic', breakLine: true } },
      { text: '', options: { fontSize: 6, breakLine: true } },
    );
  }
  slide9.addText(lowReviewTexts, { x: 0.5, y: 1.5, w: 9, h: 4, valign: 'top' });

  // ---- Slide 10: Summary ----
  const slide10 = pptx.addSlide({ masterName: 'MASTER' });
  slide10.addText('まとめ', { x: 0.5, y: 0.8, w: 9, h: 0.5, fontSize: 24, bold: true, color: '1F4E79', fontFace: 'Yu Gothic' });

  // Left box - Strengths
  slide10.addShape(pptx.ShapeType.rect, { x: 0.3, y: 1.5, w: 4.4, h: 3.5, fill: { color: 'E8F5E9' }, line: { color: '2E7D32', width: 1 } });
  slide10.addText('強み', { x: 0.5, y: 1.6, w: 4, h: 0.4, fontSize: 16, bold: true, color: '2E7D32', fontFace: 'Yu Gothic' });

  const strengths = [
    { text: '品川シーサイド駅直結の好立地', options: { fontSize: 11, color: '333333', fontFace: 'Yu Gothic', bullet: true, breakLine: true } },
    { text: 'リーズナブルな価格設定', options: { fontSize: 11, color: '333333', fontFace: 'Yu Gothic', bullet: true, breakLine: true } },
    { text: '親切で丁寧なスタッフ対応', options: { fontSize: 11, color: '333333', fontFace: 'Yu Gothic', bullet: true, breakLine: true } },
    { text: '清潔感のある客室', options: { fontSize: 11, color: '333333', fontFace: 'Yu Gothic', bullet: true, breakLine: true } },
    { text: '12時チェックアウト', options: { fontSize: 11, color: '333333', fontFace: 'Yu Gothic', bullet: true, breakLine: true } },
    { text: '各階のウォーターサーバー', options: { fontSize: 11, color: '333333', fontFace: 'Yu Gothic', bullet: true, breakLine: true } },
    { text: '羽田空港へのアクセス良好', options: { fontSize: 11, color: '333333', fontFace: 'Yu Gothic', bullet: true, breakLine: true } },
  ];
  slide10.addText(strengths, { x: 0.5, y: 2.1, w: 4, h: 2.8, valign: 'top' });

  // Right box - Weaknesses
  slide10.addShape(pptx.ShapeType.rect, { x: 5.3, y: 1.5, w: 4.4, h: 3.5, fill: { color: 'FFEBEE' }, line: { color: 'C62828', width: 1 } });
  slide10.addText('改善課題', { x: 5.5, y: 1.6, w: 4, h: 0.4, fontSize: 16, bold: true, color: 'C62828', fontFace: 'Yu Gothic' });

  const weaknesses = [
    { text: '清掃品質の一貫性', options: { fontSize: 11, color: '333333', fontFace: 'Yu Gothic', bullet: true, breakLine: true } },
    { text: '朝食開始時間の早期化', options: { fontSize: 11, color: '333333', fontFace: 'Yu Gothic', bullet: true, breakLine: true } },
    { text: '水回り設備の更新', options: { fontSize: 11, color: '333333', fontFace: 'Yu Gothic', bullet: true, breakLine: true } },
    { text: 'クレーム対応力の向上', options: { fontSize: 11, color: '333333', fontFace: 'Yu Gothic', bullet: true, breakLine: true } },
    { text: '周辺飲食情報の充実', options: { fontSize: 11, color: '333333', fontFace: 'Yu Gothic', bullet: true, breakLine: true } },
    { text: '喫煙室の臭い対策', options: { fontSize: 11, color: '333333', fontFace: 'Yu Gothic', bullet: true, breakLine: true } },
  ];
  slide10.addText(weaknesses, { x: 5.5, y: 2.1, w: 4, h: 2.8, valign: 'top' });

  await pptx.writeFile({ fileName: outputPath });
  console.log(`PPTX saved to: ${outputPath}`);
}

// ============================================================
// 5. MAIN
// ============================================================

async function main() {
  const csvPath = '/Users/mitsugutakahashi/ホテル口コミ/hearton_data.csv';
  const outputDir = '/Users/mitsugutakahashi/ホテル口コミ/hotel-review-report-workspace/iteration-1/eval-hearton-report/without_skill/outputs';

  // Ensure output directory exists
  fs.mkdirSync(outputDir, { recursive: true });

  console.log('Parsing CSV data...');
  const reviews = parseCSV(csvPath);
  console.log(`Parsed ${reviews.length} unique reviews`);

  console.log('Analyzing reviews...');
  const analysis = analyzeReviews(reviews);

  console.log(`Total reviews: ${analysis.totalReviews}`);
  console.log(`Average rating (5-pt scale): ${analysis.avgRating.toFixed(2)}`);
  console.log('Site stats:');
  for (const [site, stats] of Object.entries(analysis.siteStats)) {
    console.log(`  ${site}: ${stats.count} reviews, avg ${stats.avgRating.toFixed(2)}/5`);
  }
  console.log('Month stats:');
  for (const [month, stats] of Object.entries(analysis.monthStats)) {
    console.log(`  ${month}: ${stats.count} reviews, avg ${stats.avgRating.toFixed(2)}/5`);
  }
  console.log('Rating distribution:', analysis.ratingDist);

  const docxPath = path.join(outputDir, 'ハートンホテル東品川_口コミ分析レポート.docx');
  const pptxPath = path.join(outputDir, 'ハートンホテル東品川_口コミ分析レポート.pptx');

  console.log('\nGenerating DOCX report...');
  await generateDocx(analysis, docxPath);

  console.log('Generating PPTX presentation...');
  await generatePptx(analysis, pptxPath);

  console.log('\nDone! Files generated:');
  console.log(`  DOCX: ${docxPath}`);
  console.log(`  PPTX: ${pptxPath}`);
}

main().catch(err => {
  console.error('Error:', err);
  process.exit(1);
});
