const fs = require('fs');
const path = require('path');
const { esc, nav, footer, pageHead, pageFoot, writeCommonCSS, copyAssets } = require('./lib/common');
const { renderA1, renderA2, renderA3, renderA4, renderA5, renderA6, renderA7 } = require('./lib/deep-analysis-renderers');

const DATA_DIR = path.join(__dirname, '..', '..', 'データ', '分析結果JSON');
const OUTPUT_DIR = path.join(__dirname, '..', '..', 'ホームページ');
if (!fs.existsSync(OUTPUT_DIR)) fs.mkdirSync(OUTPUT_DIR, { recursive: true });

// Create output subdirectories
['styles', 'scripts', 'data'].forEach(function(d) {
  var dir = path.join(OUTPUT_DIR, d);
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
});

// Load all data
const portfolio = JSON.parse(fs.readFileSync(path.join(DATA_DIR, 'primechange_portfolio_analysis.json'), 'utf8'));
const meta = portfolio.report_metadata;
const pov = portfolio.portfolio_overview;
const cleanDive = portfolio.cleaning_deep_dive;
const priMatrix = portfolio.priority_matrix;
const actionPlans = portfolio.action_plans;
const crossRec = portfolio.cross_cutting_recommendations;
const kpi = portfolio.kpi_framework;
const roi = portfolio.roi_estimation;

// Load analysis data
function loadAnalysis(n) {
  try { return JSON.parse(fs.readFileSync(path.join(DATA_DIR, 'analysis_' + n + '_data.json'), 'utf8')); }
  catch(e) { return null; }
}
const a1 = loadAnalysis(1);
const a2 = loadAnalysis(2);
const a3 = loadAnalysis(3);
const a4 = loadAnalysis(4);
const a5 = loadAnalysis(5);
const a6 = loadAnalysis(6);
const a7 = loadAnalysis(7);

// Load hotel details
const hotelFiles = fs.readdirSync(DATA_DIR).filter(function(f) { return f.endsWith('_analysis.json') && !f.includes('portfolio'); });
var hotelDetails = {};
var allReviewsCompact = {}; // For hotel-reviews-all.json
var allDates = [];
hotelFiles.forEach(function(file) {
  var key = file.replace('_analysis.json', '');
  var data = JSON.parse(fs.readFileSync(path.join(DATA_DIR, file), 'utf8'));
  var allComments = (data.comments || []).map(function(c) {
    var d = c.date || '';
    if (d) allDates.push(d);
    return { site: c.site, rating_10pt: c.rating_10pt, date: d,
      comment: (c.translated || c.comment || '').slice(0, 300),
      good: (c.translated_good || c.good || '').slice(0, 200),
      bad: (c.translated_bad || c.bad || '').slice(0, 200) };
  });
  hotelDetails[key] = {
    total_reviews: data.total_reviews, overall_avg_10pt: data.overall_avg_10pt,
    high_count: data.high_count, high_rate: data.high_rate,
    mid_count: data.mid_count, mid_rate: data.mid_rate,
    low_count: data.low_count, low_rate: data.low_rate,
    site_stats: data.site_stats, distribution: data.distribution,
    comments: allComments.slice(0, 30)
  };
  // Compact keys for all reviews: s=site, r=rating, d=date, c=comment, g=good, b=bad
  allReviewsCompact[key] = allComments.map(function(c) {
    return { s: c.site, r: c.rating_10pt, d: c.date, c: c.comment, g: c.good, b: c.bad };
  });
});
allDates.sort();
var dateMin = allDates.length > 0 ? allDates[0] : '';
var dateMax = allDates.length > 0 ? allDates[allDates.length - 1] : '';
var totalAllReviews = allDates.length;

var tierColor = { '優秀': '#10B981', '良好': '#3B82F6', '概ね良好': '#F59E0B', '要改善': '#EF4444' };
var priorityLabel = { 'MAINTENANCE': '維持', 'STANDARD': '標準', 'HIGH': '要注意', 'URGENT': '緊急' };
var priorityColor = { 'MAINTENANCE': '#10B981', 'STANDARD': '#6B7280', 'HIGH': '#F59E0B', 'URGENT': '#EF4444' };
var severityColor = { 'CRITICAL': '#EF4444', 'HIGH': '#F59E0B', 'MEDIUM': '#3B82F6', 'LOW': '#6B7280' };

function writePage(filename, content) {
  fs.writeFileSync(path.join(OUTPUT_DIR, filename), content, 'utf8');
  console.log('  ' + filename);
}

function writeJSON(filename, data) {
  fs.writeFileSync(path.join(OUTPUT_DIR, 'data', filename), JSON.stringify(data), 'utf8');
  console.log('  data/' + filename);
}

function badgeFor(tier) {
  var cls = tier === '優秀' ? 'badge-green' : tier === '良好' ? 'badge-blue' : tier === '概ね良好' ? 'badge-orange' : 'badge-red';
  return '<span class="badge ' + cls + '">' + esc(tier) + '</span>';
}
function priBadge(pri) {
  var cls = pri === 'URGENT' ? 'badge-red' : pri === 'HIGH' ? 'badge-orange' : pri === 'MAINTENANCE' ? 'badge-green' : 'badge-gray';
  return '<span class="badge ' + cls + '">' + (priorityLabel[pri] || pri) + '</span>';
}

// ============================================================
// 1. INDEX (Portal)
// ============================================================
function buildIndex() {
  var reportCards = [
    { icon: '&#128200;', title: 'ホテル別口コミダッシュボード', desc: '19ホテルの口コミ分析をカード形式で一覧。サイト別評価・スコア分布・口コミ詳細をモーダルで閲覧。', link: 'hotel-dashboard.html', stat: pov.hotels_ranked.length + 'ホテル / ' + meta.total_reviews + '件' },
    { icon: '&#128167;', title: '清掃戦略レポート', desc: '清掃品質の課題分析、カテゴリ別ヒートマップ、優先度マトリクス、横断的改善施策を一覧。', link: 'cleaning-strategy.html', stat: cleanDive.total_cleaning_mentions + '件の清掃指摘' },
    { icon: '&#128270;', title: '7つの深掘り分析', desc: 'クレーム類型・スタッフ・人員配置・完了時間・安全・品質売上・ベストプラクティスの7分析。', link: 'deep-analysis.html', stat: '7つの専門分析' },
    { icon: '&#128176;', title: '品質×売上・ROI分析', desc: '品質スコアと売上の相関分析、弾力性係数、3つのROIシナリオ比較。', link: 'revenue-impact.html', stat: '月間約' + Math.round(a6 && a6.portfolio_summary ? a6.portfolio_summary.total_monthly_revenue / 10000 : 0).toLocaleString() + '万円' },
    { icon: '&#9989;', title: 'アクションプラン', desc: '19ホテル別の3フェーズ改善計画（即時/短期/中期）とKPI目標管理。', link: 'action-plans.html', stat: actionPlans ? actionPlans.length + 'ホテルの改善計画' : '' },
  ];

  // Urgent hotels highlight
  var urgentHotels = (priMatrix && priMatrix.urgent) || [];

  var html = [
    pageHead('PRIMECHANGE ホテル品質管理ポータル'),
    nav('index'),
    '<div class="container">',
    '  <h1 class="page-title">ホテル品質管理ポータル</h1>',
    '  <p class="page-subtitle">PRIMECHANGE ' + meta.total_hotels + 'ホテル・' + meta.total_reviews.toLocaleString() + '件の口コミデータに基づく分析ダッシュボード &mdash; ' + esc(meta.date) + '</p>',
    '',
    '  <div class="kpi-grid" id="indexKpiGrid">',
    '    <div class="kpi-card" style="border-left-color:var(--accent);"><div class="kpi-label">管理ホテル数</div><div class="kpi-value" data-kpi="total_hotels">' + meta.total_hotels + '</div><div class="kpi-sub">ホテル</div></div>',
    '    <div class="kpi-card" style="border-left-color:var(--green);"><div class="kpi-label">総口コミ数</div><div class="kpi-value" data-kpi="total_reviews">' + meta.total_reviews.toLocaleString() + '</div><div class="kpi-sub">件</div></div>',
    '    <div class="kpi-card" style="border-left-color:' + (pov.avg_score >= 8 ? 'var(--green)' : 'var(--orange)') + ';"><div class="kpi-label">平均スコア</div><div class="kpi-value" data-kpi="avg_score">' + pov.avg_score + '</div><div class="kpi-sub">/ 10 点</div></div>',
    '    <div class="kpi-card" style="border-left-color:var(--green);"><div class="kpi-label">高評価率</div><div class="kpi-value" data-kpi="high_rate">' + pov.portfolio_high_rate + '%</div><div class="kpi-sub">8点以上</div></div>',
    '    <div class="kpi-card" style="border-left-color:var(--red);"><div class="kpi-label">清掃課題率</div><div class="kpi-value" data-kpi="cleaning_issue_rate">' + cleanDive.portfolio_cleaning_issue_rate + '%</div><div class="kpi-sub" data-kpi="cleaning_issue_count">' + cleanDive.total_cleaning_mentions + '件</div></div>',
    '  </div>',
    '',
    '  <div class="grid-cards">',
  ];

  reportCards.forEach(function(c) {
    html.push(
      '    <a href="' + c.link + '" class="link-card">',
      '    <div class="card">',
      '      <div class="link-card-icon">' + c.icon + '</div>',
      '      <div class="link-card-title">' + c.title + '</div>',
      '      <div class="link-card-desc">' + c.desc + '</div>',
      '      <div class="link-card-stat">' + c.stat + ' &rarr;</div>',
      '    </div></a>'
    );
  });

  html.push('  </div>');

  // Urgent alert
  if (urgentHotels.length > 0) {
    html.push(
      '  <div class="card alert-card">',
      '    <div class="card-title" style="color:var(--red);">&#9888; 緊急対応が必要なホテル (' + urgentHotels.length + '件)</div>',
      '    <div class="alert-grid">'
    );
    urgentHotels.forEach(function(h) {
      html.push(
        '      <div class="alert-item">',
        '        <div class="alert-item-title">' + esc(h.hotel) + '</div>',
        '        <div class="alert-item-detail">スコア: <strong style="color:var(--red);">' + h.avg + '</strong> / 清掃課題率: <strong>' + h.cleaning_rate + '%</strong></div>',
        '        <div class="alert-item-sub">主要課題: ' + (h.key_problems || []).join(', ') + '</div>',
        '      </div>'
      );
    });
    html.push('    </div>', '  </div>');
  }

  // Portfolio trend chart
  html.push(
    '  <div class="card">',
    '    <div class="card-title">&#128200; 口コミトレンド（日別）</div>',
    '    <div id="portfolioTrend" style="width:100%;"></div>',
    '  </div>'
  );

  // KPI targets with SVG gauges
  if (kpi && kpi.portfolio_targets) {
    var targets = kpi.portfolio_targets;
    html.push(
      '  <div class="card">',
      '    <div class="card-title">&#127919; KPI目標 (2026年9月)</div>',
      '    <div class="gauge-row">'
    );
    var gauges = [
      { label: '平均スコア', current: pov.avg_score, target: targets.target_avg_score || 8.89, unit: '', color: 'var(--accent)', lower: false },
      { label: '清掃課題率', current: cleanDive.portfolio_cleaning_issue_rate, target: targets.target_cleaning_rate || 1.8, unit: '%', color: 'var(--red)', lower: true },
      { label: '高評価率', current: pov.portfolio_high_rate, target: targets.target_high_rate || 83.1, unit: '%', color: 'var(--green)', lower: false },
      { label: '低評価率', current: pov.portfolio_low_rate, target: targets.target_low_rate || 2.4, unit: '%', color: 'var(--orange)', lower: true },
    ];
    gauges.forEach(function(g) {
      html.push(
        '      <div class="gauge-item">',
        '        <div class="gauge-label">' + g.label + '</div>',
        '        <div class="svg-gauge" data-value="' + g.current + '" data-target="' + g.target + '" data-unit="' + g.unit + '" data-label="' + g.label + '" data-color="' + g.color + '" data-lower="' + g.lower + '"></div>',
        '      </div>'
      );
    });
    html.push('    </div>', '  </div>');
  }

  html.push('</div>', footer(), pageFoot());
  writePage('index.html', html.join('\n'));
}

// ============================================================
// 2. HOTEL DASHBOARD
// ============================================================
function buildHotelDashboard() {
  var extraCSS = [
    '.hotel-card { background: var(--card); border-radius: var(--radius); box-shadow: var(--shadow); overflow: hidden; cursor: pointer; transition: transform 0.2s, box-shadow 0.2s; }',
    '.hotel-card:hover { transform: translateY(-3px); box-shadow: var(--shadow-lg); }',
    '.filter-bar { background: var(--card); border-radius: var(--radius); padding: 1rem 1.5rem; box-shadow: var(--shadow); margin-bottom: 1.5rem; display: flex; gap: 0.75rem; flex-wrap: wrap; align-items: center; }',
    '.filter-btn { padding: 0.4rem 1rem; border-radius: 20px; border: 1.5px solid var(--border); background: white; cursor: pointer; font-size: 0.8rem; font-weight: 500; transition: all 0.2s; }',
    '.filter-btn:hover { border-color: var(--accent); color: var(--accent); }',
    '.filter-btn.active { background: var(--accent); color: white; border-color: var(--accent); }',
    '.modal-overlay { display:none; position:fixed; top:0; left:0; right:0; bottom:0; background:rgba(0,0,0,0.5); z-index:200; justify-content:center; align-items:flex-start; padding:2rem; overflow-y:auto; }',
    '.modal-overlay.active { display:flex; }',
    '.modal { background:var(--card); border-radius:var(--radius); max-width:900px; width:100%; box-shadow:var(--shadow-lg); animation:slideUp 0.3s ease; }',
    '@keyframes slideUp { from { opacity:0; transform:translateY(20px); } to { opacity:1; transform:translateY(0); } }',
    '.modal-header { padding:1.5rem 2rem; border-bottom:1px solid var(--border); display:flex; justify-content:space-between; align-items:center; }',
    '.modal-close { width:36px; height:36px; border-radius:50%; border:none; background:#F1F5F9; cursor:pointer; font-size:1.2rem; display:flex; align-items:center; justify-content:center; transition: background 0.2s; }',
    '.modal-close:hover { background:#E2E8F0; }',
    '.modal-body { padding:2rem; }',
    '.review-list { display:flex; flex-direction:column; gap:0.75rem; max-height:500px; overflow-y:auto; }',
    '.review-card { padding:1rem; border-radius:8px; background:#F8FAFC; border-left:3px solid var(--border); transition: background 0.2s; }',
    '.review-card:hover { background:#EFF6FF; }',
    '.review-card.high { border-left-color:var(--green); }',
    '.review-card.mid { border-left-color:var(--orange); }',
    '.review-card.low { border-left-color:var(--red); }',
    '@media (max-width:768px) { .filter-bar { flex-direction:column; align-items:stretch; } }',
  ].join('\n');

  function buildCards() {
    return pov.hotels_ranked.map(function(h) {
      var tc = tierColor[h.tier];
      var pc = priorityColor[h.priority];
      var pl = priorityLabel[h.priority];
      var barPct = (h.avg / 10 * 100).toFixed(0);
      var rankBg = h.rank <= 3 ? '#F59E0B' : h.rank <= 10 ? '#3B82F6' : '#94A3B8';
      var detail = hotelDetails[h.key] || {};
      var sites = (detail.site_stats || []).map(function(s) {
        return '<span style="font-size:0.65rem;padding:0.15rem 0.5rem;border-radius:10px;background:#F1F5F9;color:var(--text-light);white-space:nowrap;">' + esc(s.site) + ' ' + s.avg_10pt + '</span>';
      }).join(' ');
      var cleanColor = h.cleaning_issue_rate > 5 ? 'var(--red)' : h.cleaning_issue_rate > 0 ? 'var(--orange)' : 'var(--green)';
      return [
        '<div class="hotel-card" data-tier="' + esc(h.tier) + '" data-key="' + esc(h.key) + '" data-name="' + esc(h.name) + '" data-rank="' + h.rank + '" data-reviews="' + h.total_reviews + '" data-high-rate="' + h.high_rate + '" data-low-rate="' + h.low_rate + '" data-cleaning="' + h.cleaning_issue_rate + '" onclick="openModal(\'' + h.key + '\')">',
        '  <div style="padding:1.25rem 1.5rem 0.75rem;display:flex;justify-content:space-between;align-items:flex-start;">',
        '    <h3 style="font-size:0.95rem;font-weight:700;line-height:1.4;flex:1;">' + esc(h.name) + '</h3>',
        '    <div style="width:32px;height:32px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:0.75rem;font-weight:800;color:white;flex-shrink:0;margin-left:0.75rem;background:' + rankBg + ';">' + h.rank + '</div>',
        '  </div>',
        '  <div style="padding:0 1.5rem 1.25rem;">',
        '    <div style="display:flex;align-items:center;gap:0.75rem;margin-bottom:0.75rem;">',
        '      <div style="font-size:2rem;font-weight:800;color:' + tc + ';">' + h.avg + '</div>',
        '      <div style="flex:1;"><div style="height:8px;background:#E2E8F0;border-radius:4px;overflow:hidden;"><div style="height:100%;border-radius:4px;width:' + barPct + '%;background:' + tc + ';"></div></div>',
        '        <div style="font-size:0.7rem;color:var(--text-light);margin-top:2px;display:flex;justify-content:space-between;"><span>' + badgeFor(h.tier) + ' ' + priBadge(h.priority) + '</span><span>' + h.total_reviews + '件</span></div>',
        '      </div>',
        '    </div>',
        '    <div style="display:flex;gap:0.5rem;margin-top:0.75rem;">',
        '      <div style="flex:1;text-align:center;padding:0.4rem;border-radius:6px;background:#F1F5F9;"><div style="font-size:0.95rem;font-weight:700;color:var(--green);">' + h.high_rate + '%</div><div style="font-size:0.6rem;color:var(--text-light);">高評価</div></div>',
        '      <div style="flex:1;text-align:center;padding:0.4rem;border-radius:6px;background:#F1F5F9;"><div style="font-size:0.95rem;font-weight:700;color:var(--red);">' + h.low_rate + '%</div><div style="font-size:0.6rem;color:var(--text-light);">低評価</div></div>',
        '      <div style="flex:1;text-align:center;padding:0.4rem;border-radius:6px;background:#F1F5F9;"><div style="font-size:0.95rem;font-weight:700;color:' + cleanColor + ';">' + h.cleaning_issue_rate + '%</div><div style="font-size:0.6rem;color:var(--text-light);">清掃課題</div></div>',
        '    </div>',
        '    <div style="display:flex;gap:0.25rem;margin-top:0.75rem;flex-wrap:wrap;">' + sites + '</div>',
        '  </div>',
        '</div>'
      ].join('\n');
    }).join('\n');
  }

  var html = [
    pageHead('ホテル別口コミダッシュボード - PRIMECHANGE', { scripts: ['hotel-dashboard.js'], extraCSS: extraCSS }),
    nav('hotel-dashboard'),
    '<div class="container">',
    '  <h1 class="page-title">ホテル別口コミダッシュボード</h1>',
    '  <p class="page-subtitle">' + meta.total_hotels + 'ホテル・' + meta.total_reviews.toLocaleString() + '件の口コミを分析 &mdash; ' + esc(meta.analysis_period) + '</p>',
    '',
    '  <div class="kpi-grid" id="dashKpiGrid">',
    '    <div class="kpi-card" style="border-left-color:var(--accent);"><div class="kpi-label">管理ホテル数</div><div class="kpi-value" data-kpi="total_hotels">' + meta.total_hotels + '</div></div>',
    '    <div class="kpi-card" style="border-left-color:var(--green);"><div class="kpi-label">総口コミ数</div><div class="kpi-value" data-kpi="total_reviews">' + meta.total_reviews.toLocaleString() + '</div></div>',
    '    <div class="kpi-card" style="border-left-color:var(--green);"><div class="kpi-label">平均スコア</div><div class="kpi-value" data-kpi="avg_score">' + pov.avg_score + '</div></div>',
    '    <div class="kpi-card" style="border-left-color:var(--green);"><div class="kpi-label">高評価率</div><div class="kpi-value" data-kpi="high_rate">' + pov.portfolio_high_rate + '%</div></div>',
    '    <div class="kpi-card" style="border-left-color:var(--red);"><div class="kpi-label">低評価率</div><div class="kpi-value" data-kpi="low_rate">' + pov.portfolio_low_rate + '%</div></div>',
    '  </div>',
    '',
    '  <div class="filter-bar">',
    '    <label style="font-size:0.8rem;font-weight:600;color:var(--text-light);">絞り込み:</label>',
    '    <button class="filter-btn active" data-filter="all">すべて (' + pov.hotels_ranked.length + ')</button>',
    '    <button class="filter-btn" data-filter="優秀">優秀</button>',
    '    <button class="filter-btn" data-filter="良好">良好</button>',
    '    <button class="filter-btn" data-filter="概ね良好">概ね良好</button>',
    '    <button class="filter-btn" data-filter="要改善">要改善</button>',
    '    <div style="flex:1;"></div>',
    '    <select id="sortSelect" style="padding:0.4rem 1rem;border-radius:20px;border:1.5px solid var(--border);font-size:0.8rem;background:white;">',
    '      <option value="rank">ランキング順</option><option value="reviews">口コミ数順</option><option value="high_rate">高評価率順</option><option value="cleaning">清掃課題率順</option>',
    '    </select>',
    '    <input type="text" id="searchInput" placeholder="ホテル名で検索..." style="padding:0.4rem 1rem;border-radius:20px;border:1.5px solid var(--border);font-size:0.85rem;width:220px;outline:none;">',
    '  </div>',
    '',
    '  <div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(340px,1fr));gap:1.25rem;margin-bottom:2rem;" id="hotelGrid">',
    buildCards(),
    '  </div>',
    '</div>',
    '',
    '<div class="modal-overlay" id="modalOverlay" onclick="if(event.target===this)closeModal();">',
    '  <div class="modal"><div class="modal-header"><h2 id="modalTitle" style="font-size:1.25rem;font-weight:700;"></h2><button class="modal-close" onclick="closeModal();">&#10005;</button></div><div class="modal-body" id="modalBody"></div><div style="padding:0 2rem 1.5rem;"><div style="font-size:0.9rem;font-weight:700;color:var(--navy);margin-bottom:0.75rem;padding-bottom:0.5rem;border-bottom:2px solid var(--accent);">&#128200; 口コミトレンド</div><div id="modalTrend" data-hotel-key=""></div></div></div>',
    '</div>',
    footer(),
    pageFoot()
  ];

  // Write external data files
  writeJSON('hotel-details.json', hotelDetails);
  writeJSON('hotel-ranked.json', pov.hotels_ranked);
  writeJSON('tier-color.json', tierColor);

  writePage('hotel-dashboard.html', html.join('\n'));
}

// ============================================================
// 3. CLEANING STRATEGY
// ============================================================
function buildCleaningStrategy() {
  var cats = cleanDive.category_summary || [];
  var matrix = cleanDive.hotel_cleaning_matrix || [];
  var allCats = cats.map(function(c) { return c.category; });

  // Category bars
  var maxMention = Math.max.apply(null, cats.map(function(c){return c.total_mentions;}));
  var catBars = cats.map(function(c) {
    var pct = (c.total_mentions / maxMention * 100).toFixed(0);
    var col = severityColor[c.severity] || '#6B7280';
    return '<div class="h-bar"><div class="h-bar-label">' + esc(c.category) + '</div><div class="h-bar-track"><div class="h-bar-fill" style="width:' + pct + '%;background:' + col + ';"><span class="h-bar-val">' + c.total_mentions + '件 (' + c.hotels_affected + 'ホテル)</span></div></div></div>';
  }).join('\n');

  // Heatmap
  var heatmapHead = '<tr><th class="row-header">ホテル名</th><th>スコア</th><th>課題率</th>' + allCats.map(function(c){ return '<th>' + esc(c) + '</th>'; }).join('') + '<th>合計</th></tr>';
  var heatmapRows = matrix.map(function(h) {
    var cells = allCats.map(function(cat) {
      var val = (h.categories && h.categories[cat]) || 0;
      var bg = val === 0 ? '#FFFFFF' : val <= 1 ? '#FEF3C7' : val <= 3 ? '#FBBF24' : val <= 5 ? '#F97316' : '#EF4444';
      var fc = val >= 4 ? 'white' : 'var(--text)';
      return '<td style="background:' + bg + ';color:' + fc + ';">' + (val || '') + '</td>';
    }).join('');
    var pc = priorityColor[h.priority] || '#6B7280';
    return '<tr><td class="row-header">' + esc(h.name) + '</td><td>' + h.avg + '</td><td style="color:' + pc + ';font-weight:700;">' + h.cleaning_issue_rate + '%</td>' + cells + '<td style="font-weight:700;">' + h.cleaning_issue_count + '</td></tr>';
  }).join('\n');

  // Priority matrix
  function priCards(level, items) {
    if (!items || items.length === 0) return '';
    var cls = level.toLowerCase();
    return items.map(function(h) {
      return '<div class="priority-card ' + cls + '"><div class="priority-title">' + esc(h.hotel) + ' <span style="font-size:0.75rem;color:var(--text-light);">(' + h.avg + '点)</span></div><div class="priority-hotels">課題: ' + h.cleaning_issues + '件 (' + h.cleaning_rate + '%) / ' + (h.key_problems || []).join(', ') + '</div></div>';
    }).join('\n');
  }

  // Cross-cutting recommendations
  var recHtml = '';
  if (crossRec && crossRec.length) {
    recHtml = crossRec.map(function(r, i) {
      return '<div class="card"><div class="card-title">施策' + (i+1) + ': ' + esc(r.title || r.theme || '') + '</div><p style="font-size:0.82rem;color:var(--text-light);line-height:1.7;">' + esc(r.description || r.detail || JSON.stringify(r).slice(0,300)) + '</p></div>';
    }).join('\n');
  } else if (crossRec && typeof crossRec === 'object' && !Array.isArray(crossRec)) {
    recHtml = Object.keys(crossRec).map(function(k, i) {
      var v = crossRec[k];
      var title = v.title || k;
      var desc = v.description || v.detail || (typeof v === 'string' ? v : JSON.stringify(v).slice(0, 500));
      return '<div class="card"><div class="card-title">施策' + (i+1) + ': ' + esc(title) + '</div><p style="font-size:0.82rem;color:var(--text-light);line-height:1.7;">' + esc(desc) + '</p></div>';
    }).join('\n');
  }

  var html = [
    pageHead('清掃戦略レポート - PRIMECHANGE'),
    nav('cleaning-strategy'),
    '<div class="container">',
    '  <h1 class="page-title">清掃戦略レポート</h1>',
    '  <p class="page-subtitle">清掃品質の課題分析と優先度別改善戦略</p>',
    '',
    '  <div class="kpi-grid" id="cleaningKpiGrid">',
    '    <div class="kpi-card" style="border-left-color:var(--red);"><div class="kpi-label">清掃課題率</div><div class="kpi-value" data-kpi="cleaning_issue_rate">' + cleanDive.portfolio_cleaning_issue_rate + '%</div></div>',
    '    <div class="kpi-card" style="border-left-color:var(--orange);"><div class="kpi-label">清掃指摘件数</div><div class="kpi-value" data-kpi="cleaning_issue_count">' + cleanDive.total_cleaning_mentions + '</div></div>',
    '    <div class="kpi-card" style="border-left-color:var(--red);"><div class="kpi-label">課題カテゴリ数</div><div class="kpi-value">' + cats.length + '</div></div>',
    '    <div class="kpi-card" style="border-left-color:var(--red);"><div class="kpi-label">緊急対応ホテル</div><div class="kpi-value">' + (priMatrix.urgent || []).length + '</div></div>',
    '  </div>',
    '',
    '  <div class="card">',
    '    <div class="card-title">&#128200; 清掃課題発生トレンド</div>',
    '    <div id="cleaningTrend" style="width:100%;"></div>',
    '  </div>',
    '',
    '  <div class="card">',
    '    <div class="card-title">&#128202; 課題カテゴリ別件数</div>',
    catBars,
    '    <div style="margin-top:0.75rem;font-size:0.7rem;color:var(--text-light);">色: <span style="color:#EF4444;">&#9632; CRITICAL</span> <span style="color:#F59E0B;">&#9632; HIGH</span> <span style="color:#3B82F6;">&#9632; MEDIUM</span> <span style="color:#6B7280;">&#9632; LOW</span></div>',
    '  </div>',
    '',
    '  <div class="card">',
    '    <div class="card-title">&#128293; ホテル×課題カテゴリ ヒートマップ</div>',
    '    <div class="heatmap-wrap"><table class="heatmap">' + heatmapHead + heatmapRows + '</table></div>',
    '  </div>',
    '',
    '  <div class="card">',
    '    <div class="card-title">&#9888; 優先度マトリクス</div>',
    '    <h3 style="font-size:0.85rem;color:var(--red);margin-bottom:0.75rem;">URGENT（緊急）</h3>',
    '    <div class="priority-grid">' + priCards('URGENT', priMatrix.urgent) + '</div>',
    '    <h3 style="font-size:0.85rem;color:var(--orange);margin:1rem 0 0.75rem;">HIGH（要注意）</h3>',
    '    <div class="priority-grid">' + priCards('HIGH', priMatrix.high) + '</div>',
    '    <h3 style="font-size:0.85rem;color:var(--blue);margin:1rem 0 0.75rem;">STANDARD（標準）</h3>',
    '    <div class="priority-grid">' + priCards('STANDARD', priMatrix.standard) + '</div>',
    '    <h3 style="font-size:0.85rem;color:var(--green);margin:1rem 0 0.75rem;">MAINTENANCE（維持）</h3>',
    '    <div class="priority-grid">' + priCards('MAINTENANCE', priMatrix.maintenance) + '</div>',
    '  </div>',
    '',
    '  <div class="card">',
    '    <div class="card-title">&#128161; 横断的改善施策</div>',
    recHtml || '<p style="font-size:0.82rem;color:var(--text-light);">データなし</p>',
    '  </div>',
    '</div>',
    footer(),
    pageFoot()
  ];
  writePage('cleaning-strategy.html', html.join('\n'));
}

// ============================================================
// 4. DEEP ANALYSIS (7 tabs)
// ============================================================
function buildDeepAnalysis() {
  var analyses = [
    { id: '1', data: a1, title: 'クレーム類型×スコア連動', icon: '&#128196;', render: renderA1 },
    { id: '2', data: a2, title: 'スタッフパフォーマンス', icon: '&#128100;', render: renderA2 },
    { id: '3', data: a3, title: '人員配置×品質相関', icon: '&#128101;', render: renderA3 },
    { id: '4', data: a4, title: '清掃完了時間×品質', icon: '&#9200;', render: renderA4 },
    { id: '5', data: a5, title: '安全チェック×予兆検出', icon: '&#128737;', render: renderA5 },
    { id: '6', data: a6, title: '品質→売上弾力性', icon: '&#128176;', render: renderA6 },
    { id: '7', data: a7, title: 'ベストプラクティス横展開', icon: '&#127942;', render: renderA7 },
  ];

  var tabBtns = analyses.map(function(a, i) {
    return '<button class="tab-btn' + (i === 0 ? ' active' : '') + '" onclick="showTab(\'' + a.id + '\')">' + a.icon + ' 分析' + a.id + '</button>';
  }).join('\n');

  var tabPanels = analyses.map(function(a, i) {
    var content = a.render(a.data);
    return '<div class="tab-panel' + (i === 0 ? ' active' : '') + '" id="tab-' + a.id + '">' + content + '</div>';
  }).join('\n');

  // Save HTML fragment for snapshot switching
  var daContentTabs = analyses.map(function(a) {
    return { id: a.id, title: a.title, icon: a.icon, html: a.render(a.data) };
  });
  writeJSON('deep-analysis-content.json', { tabs: daContentTabs });

  var html = [
    pageHead('7つの深掘り分析 - PRIMECHANGE'),
    nav('deep-analysis'),
    '<div class="container">',
    '  <h1 class="page-title">7つの深掘り分析</h1>',
    '  <p class="page-subtitle">清掃品質・スタッフ・安全・売上の多角的分析</p>',
    '  <div class="fulldata-banner"><span>&#9432;</span> このページの分析（スタッフ・人員配置・売上等）は全期間データに基づいています。日付フィルターは口コミ由来の分析に部分適用されます。</div>',
    '  <div id="da-content">',
    '  <div class="tabs">' + tabBtns + '</div>',
    tabPanels,
    '  </div>',
    '</div>',
    footer(),
    pageFoot()
  ];
  writePage('deep-analysis.html', html.join('\n'));
}

// ============================================================
// 5. REVENUE IMPACT
// ============================================================
function buildRevenueImpact() {
  // Build content portion separately for both page and snapshot fragment
  var content = [];

  // Portfolio summary
  if (a6 && a6.portfolio_summary) {
    var ps = a6.portfolio_summary;
    content.push(
      '  <div class="kpi-grid">',
      '    <div class="kpi-card" style="border-left-color:var(--accent);"><div class="kpi-label">月間売上合計</div><div class="kpi-value">&#165;' + Math.round(ps.total_monthly_revenue / 10000).toLocaleString() + '万</div></div>',
      '    <div class="kpi-card" style="border-left-color:var(--green);"><div class="kpi-label">平均稼働率</div><div class="kpi-value">' + ps.avg_occupancy + '</div></div>',
      '    <div class="kpi-card" style="border-left-color:var(--blue);"><div class="kpi-label">平均ADR</div><div class="kpi-value">&#165;' + Math.round(ps.avg_adr).toLocaleString() + '</div></div>',
      '    <div class="kpi-card" style="border-left-color:var(--accent);"><div class="kpi-label">平均RevPAR</div><div class="kpi-value">&#165;' + Math.round(ps.avg_revpar).toLocaleString() + '</div></div>',
      '  </div>'
    );
  }

  // Regression results
  if (a6 && a6.regression_results) {
    var reg = a6.regression_results;
    content.push('<div class="card"><div class="card-title">&#128200; 回帰分析結果</div>');
    content.push('<div style="overflow-x:auto;"><table class="data-table"><thead><tr><th>指標</th><th>傾き</th><th>相関係数 r</th><th>R&sup2;</th><th>解釈</th></tr></thead><tbody>');
    Object.keys(reg).forEach(function(k) {
      var r = reg[k];
      var rColor = Math.abs(r.r) >= 0.5 ? 'var(--green)' : Math.abs(r.r) >= 0.3 ? 'var(--orange)' : 'var(--red)';
      content.push('<tr><td><strong>' + esc(r.y_label || k) + '</strong></td><td>' + (typeof r.slope === 'number' ? r.slope.toFixed(2) : r.slope) + '</td><td style="color:' + rColor + ';font-weight:700;">' + r.r.toFixed(4) + '</td><td>' + r.r_squared.toFixed(4) + '</td><td style="font-size:0.75rem;">' + esc(r.interpretation || '') + '</td></tr>');
    });
    content.push('</tbody></table></div></div>');
  }

  // Threshold analysis
  if (a6 && a6.threshold_analysis && a6.threshold_analysis.groups) {
    var groups = a6.threshold_analysis.groups;
    content.push('<div class="card"><div class="card-title">&#128201; スコア帯別パフォーマンス</div>');
    content.push('<div style="overflow-x:auto;"><table class="data-table"><thead><tr><th>スコア帯</th><th>ホテル数</th><th>平均稼働率</th><th>平均ADR</th><th>平均RevPAR</th><th>平均月間売上</th></tr></thead><tbody>');
    groups.forEach(function(g) {
      content.push('<tr><td><strong>' + esc(g.range) + '</strong></td><td>' + g.count + '</td><td>' + g.avg_occupancy_pct + '</td><td>&#165;' + Math.round(g.avg_adr).toLocaleString() + '</td><td>&#165;' + Math.round(g.avg_revpar).toLocaleString() + '</td><td>&#165;' + Math.round(g.avg_revenue).toLocaleString() + '</td></tr>');
    });
    content.push('</tbody></table></div></div>');
  }

  // Revenue impact scenarios
  if (a6 && a6.revenue_impact_scenarios) {
    var scenarios = a6.revenue_impact_scenarios;
    content.push('<div class="card"><div class="card-title">&#128176; 品質改善による売上インパクト</div>');
    scenarios.forEach(function(sc) {
      content.push('<h3 style="font-size:0.9rem;font-weight:700;color:var(--accent);margin:1rem 0 0.5rem;">スコア +' + sc.score_improvement + '点改善シナリオ</h3>');
      content.push('<div class="kpi-grid" style="margin-bottom:1rem;">');
      content.push('<div class="kpi-card" style="border-left-color:var(--green);"><div class="kpi-label">RevPAR変動</div><div class="kpi-value">+&#165;' + Math.round(sc.revpar_change).toLocaleString() + '</div><div class="kpi-sub">' + sc.revpar_pct_change + '</div></div>');
      content.push('<div class="kpi-card" style="border-left-color:var(--green);"><div class="kpi-label">月間売上変動</div><div class="kpi-value">+&#165;' + Math.round(sc.total_monthly_revenue_change).toLocaleString() + '</div></div>');
      content.push('<div class="kpi-card" style="border-left-color:var(--accent);"><div class="kpi-label">年間売上変動</div><div class="kpi-value">+&#165;' + Math.round(sc.annual_revenue_change).toLocaleString() + '</div></div>');
      content.push('</div>');

      // Per hotel
      if (sc.per_hotel_impacts) {
        content.push('<details style="margin-bottom:1rem;"><summary style="cursor:pointer;font-size:0.8rem;font-weight:600;color:var(--accent);">ホテル別詳細 (' + sc.per_hotel_impacts.length + '件)</summary>');
        content.push('<table class="data-table" style="margin-top:0.5rem;"><thead><tr><th>ホテル名</th><th>現スコア</th><th>現月間売上</th><th>売上変動額</th></tr></thead><tbody>');
        sc.per_hotel_impacts.forEach(function(h) {
          content.push('<tr><td>' + esc(h.name) + '</td><td>' + h.current_score + '</td><td>&#165;' + h.current_revenue.toLocaleString() + '</td><td style="color:var(--green);font-weight:700;">+&#165;' + h.estimated_revenue_change.toLocaleString() + '</td></tr>');
        });
        content.push('</tbody></table></details>');
      }
    });
    content.push('</div>');
  }

  // ROI estimation
  if (roi) {
    content.push('<div class="card"><div class="card-title">&#128178; ROI推定シナリオ</div>');
    if (roi.methodology) content.push('<p style="font-size:0.8rem;color:var(--text-light);margin-bottom:1rem;">' + esc(roi.methodology) + '</p>');
    var roiScenarios = roi.scenarios || [roi.scenario_a, roi.scenario_b, roi.scenario_c].filter(Boolean);
    if (roiScenarios.length) {
      content.push('<div style="overflow-x:auto;"><table class="data-table"><thead><tr><th>シナリオ</th><th>対象</th><th>投資額</th><th>期待効果</th><th>ROI回収期間</th></tr></thead><tbody>');
      roiScenarios.forEach(function(s, i) {
        content.push('<tr><td><strong>' + esc(s.name || s.label || 'シナリオ' + (i+1)) + '</strong></td><td>' + esc(s.scope || s.target || '') + '</td><td>' + esc(s.cost || s.investment || '') + '</td><td>' + esc(s.effect || s.expected_gain || '') + '</td><td>' + esc(s.payback || s.roi_period || '') + '</td></tr>');
      });
      content.push('</tbody></table></div>');
    } else {
      content.push('<pre style="font-size:0.75rem;background:#F1F5F9;padding:1rem;border-radius:8px;overflow-x:auto;">' + esc(JSON.stringify(roi, null, 2).slice(0, 2000)) + '</pre>');
    }
    content.push('</div>');
  }

  // Save HTML fragment for snapshot switching
  writeJSON('revenue-impact-content.json', { html: content.join('\n') });

  var html = [
    pageHead('品質×売上・ROI分析 - PRIMECHANGE'),
    nav('revenue-impact'),
    '<div class="container">',
    '  <h1 class="page-title">品質×売上・ROI分析</h1>',
    '  <p class="page-subtitle">口コミスコアと売上の相関分析・投資対効果シナリオ</p>',
    '  <div class="fulldata-banner"><span>&#9432;</span> このページは全期間データに基づく回帰分析・ROI推定を表示しています。日付フィルターは適用されません。</div>',
    '  <div id="ri-content">',
    content.join('\n'),
    '  </div>',
    '</div>',
    footer(),
    pageFoot()
  ];
  writePage('revenue-impact.html', html.join('\n'));
}

// ============================================================
// 6. ACTION PLANS
// ============================================================
function buildActionPlans() {
  var plans = actionPlans || [];

  // Group by priority
  var groups = { URGENT: [], HIGH: [], STANDARD: [], MAINTENANCE: [] };
  plans.forEach(function(p) { if (groups[p.priority_level]) groups[p.priority_level].push(p); else groups.STANDARD.push(p); });

  function renderPhase(phase, num, color) {
    if (!phase) return '';
    var actions = phase.actions || [];
    var items = actions.map(function(a) {
      return '<li>' + esc(a.action) + (a.category ? ' <span class="action-cat">' + esc(a.category) + '</span>' : '') + '</li>';
    }).join('\n');
    return [
      '<div class="phase">',
      '  <div class="phase-header">',
      '    <div class="phase-num" style="background:' + color + ';">' + num + '</div>',
      '    <div class="phase-title">フェーズ' + num + '</div>',
      '    <div class="phase-timeline">' + esc(phase.timeline || '') + '</div>',
      '  </div>',
      '  <ul class="action-list">' + items + '</ul>',
      '</div>'
    ].join('\n');
  }

  function renderPlan(p) {
    return [
      '<div class="accordion-item">',
      '  <div class="accordion-header">',
      '    <div>' + priBadge(p.priority_level) + ' ' + esc(p.hotel) + ' <span style="font-size:0.75rem;color:var(--text-light);">(' + p.current_avg + ' &rarr; ' + p.target_avg + ')</span></div>',
      '    <span class="accordion-arrow">&#9660;</span>',
      '  </div>',
      '  <div class="accordion-body">',
      renderPhase(p.phase1_immediate, 1, 'var(--red)'),
      renderPhase(p.phase2_short_term, 2, 'var(--orange)'),
      renderPhase(p.phase3_medium_term, 3, 'var(--blue)'),
      '  </div>',
      '</div>'
    ].join('\n');
  }

  // Build content portion separately for snapshot fragment
  var content = [];

  content.push(
    '  <div class="kpi-grid">',
    '    <div class="kpi-card" style="border-left-color:var(--red);"><div class="kpi-label">緊急</div><div class="kpi-value">' + groups.URGENT.length + '</div></div>',
    '    <div class="kpi-card" style="border-left-color:var(--orange);"><div class="kpi-label">要注意</div><div class="kpi-value">' + groups.HIGH.length + '</div></div>',
    '    <div class="kpi-card" style="border-left-color:var(--blue);"><div class="kpi-label">標準</div><div class="kpi-value">' + groups.STANDARD.length + '</div></div>',
    '    <div class="kpi-card" style="border-left-color:var(--green);"><div class="kpi-label">維持</div><div class="kpi-value">' + groups.MAINTENANCE.length + '</div></div>',
    '  </div>'
  );

  // KPI targets
  if (kpi && kpi.portfolio_targets) {
    content.push('<div class="card"><div class="card-title">&#127919; ポートフォリオ KPI目標</div>');
    content.push('<div style="overflow-x:auto;"><table class="data-table"><thead><tr><th>指標</th><th>現在値</th><th>目標値</th><th>期限</th></tr></thead><tbody>');
    var pt = kpi.portfolio_targets;
    var kpiRows = [
      ['平均スコア', pov.avg_score, pt.target_avg_score],
      ['清掃課題率', cleanDive.portfolio_cleaning_issue_rate + '%', pt.target_cleaning_rate + '%'],
      ['高評価率', pov.portfolio_high_rate + '%', pt.target_high_rate + '%'],
      ['低評価率', pov.portfolio_low_rate + '%', pt.target_low_rate + '%'],
    ];
    kpiRows.forEach(function(r) {
      content.push('<tr><td><strong>' + r[0] + '</strong></td><td>' + r[1] + '</td><td style="color:var(--accent);font-weight:700;">' + (r[2] || '-') + '</td><td>' + (pt.deadline || '2026年9月') + '</td></tr>');
    });
    content.push('</tbody></table></div></div>');
  }

  // Plans by group
  [['URGENT', '緊急対応', 'var(--red)'], ['HIGH', '要注意', 'var(--orange)'], ['STANDARD', '標準', 'var(--blue)'], ['MAINTENANCE', '維持', 'var(--green)']].forEach(function(g) {
    var level = g[0], label = g[1], color = g[2];
    if (groups[level].length === 0) return;
    content.push('<h2 style="font-size:1rem;font-weight:700;color:' + color + ';margin:1.5rem 0 0.75rem;">' + label + ' (' + groups[level].length + 'ホテル)</h2>');
    groups[level].forEach(function(p) { content.push(renderPlan(p)); });
  });

  // Save HTML fragment for snapshot switching
  writeJSON('action-plans-content.json', { html: content.join('\n') });

  var html = [
    pageHead('アクションプラン - PRIMECHANGE'),
    nav('action-plans'),
    '<div class="container">',
    '  <h1 class="page-title">アクションプラン</h1>',
    '  <p class="page-subtitle">19ホテル別 3フェーズ改善計画</p>',
    '  <div class="fulldata-banner"><span>&#9432;</span> アクションプランは全期間の分析結果に基づいています。日付フィルターは適用されません。</div>',
    '  <div id="ap-content">',
    content.join('\n'),
    '  </div>',
    '</div>',
    footer(),
    pageFoot()
  ];
  writePage('action-plans.html', html.join('\n'));
}

// ============================================================
// BUILD ALL
// ============================================================
console.log('Building PRIMECHANGE Portal...');

// Write shared assets
writeCommonCSS(OUTPUT_DIR);
copyAssets(OUTPUT_DIR);

// Write all reviews data for client-side filtering
writeJSON('hotel-reviews-all.json', allReviewsCompact);

// Calculate portfolio summary KPIs for snapshot
var CLEANING_KEYWORDS = ['清掃', '汚れ', 'ゴミ', '髪の毛', 'シミ', 'カビ', 'ほこり', '埃', '汚い', '不潔', '臭い', 'におい', '匂い', 'ホコリ', 'しみ', 'かび', 'ごみ'];
function hasCleaningIssue(text) {
  for (var i = 0; i < CLEANING_KEYWORDS.length; i++) {
    if (text.indexOf(CLEANING_KEYWORDS[i]) !== -1) return true;
  }
  return false;
}

var portfolioTotalReviews = 0, portfolioScoreSum = 0, portfolioHighCount = 0, portfolioCleanCount = 0;
Object.keys(allReviewsCompact).forEach(function(key) {
  allReviewsCompact[key].forEach(function(r) {
    portfolioTotalReviews++;
    var score = parseFloat(r.r) || 0;
    portfolioScoreSum += score;
    if (score >= 8) portfolioHighCount++;
    var text = (r.c || '') + (r.g || '') + (r.b || '');
    if (hasCleaningIssue(text)) portfolioCleanCount++;
  });
});

var portfolioAvgScore = portfolioTotalReviews > 0 ? Math.round(portfolioScoreSum / portfolioTotalReviews * 100) / 100 : 0;
var portfolioHighRate = portfolioTotalReviews > 0 ? Math.round(portfolioHighCount / portfolioTotalReviews * 1000) / 10 : 0;
var portfolioCleanRate = portfolioTotalReviews > 0 ? Math.round(portfolioCleanCount / portfolioTotalReviews * 1000) / 10 : 0;

var portfolioSummary = {
  total_hotels: Object.keys(allReviewsCompact).length,
  total_reviews: portfolioTotalReviews,
  avg_score: portfolioAvgScore,
  high_rate: portfolioHighRate,
  cleaning_issue_rate: portfolioCleanRate,
  cleaning_issue_count: portfolioCleanCount
};

writeJSON('portfolio-summary.json', portfolioSummary);

// Write build metadata
var buildDate = new Date().toISOString().slice(0, 10);
var buildMetaData = {
  build_date: buildDate,
  data_range: { min: dateMin, max: dateMax },
  total_reviews: totalAllReviews,
  snapshot_id: buildDate
};
writeJSON('build-meta.json', buildMetaData);

// Save snapshot
var snapshotDir = path.join(OUTPUT_DIR, 'data', 'snapshots', buildDate);
if (!fs.existsSync(snapshotDir)) fs.mkdirSync(snapshotDir, { recursive: true });
fs.copyFileSync(path.join(OUTPUT_DIR, 'data', 'hotel-reviews-all.json'), path.join(snapshotDir, 'hotel-reviews-all.json'));
fs.copyFileSync(path.join(OUTPUT_DIR, 'data', 'hotel-details.json'), path.join(snapshotDir, 'hotel-details.json'));
fs.copyFileSync(path.join(OUTPUT_DIR, 'data', 'build-meta.json'), path.join(snapshotDir, 'build-meta.json'));
fs.copyFileSync(path.join(OUTPUT_DIR, 'data', 'portfolio-summary.json'), path.join(snapshotDir, 'portfolio-summary.json'));
console.log('  data/snapshots/' + buildDate + '/');

// Update snapshot index (with KPI summary for comparison without fetching)
var snapshotIndexPath = path.join(OUTPUT_DIR, 'data', 'snapshot-index.json');
var snapshotIndex = [];
try { snapshotIndex = JSON.parse(fs.readFileSync(snapshotIndexPath, 'utf8')); } catch(e) {}
var existingIdx = snapshotIndex.findIndex(function(s) { return s.id === buildDate; });
var snapshotEntry = {
  id: buildDate,
  date: buildDate,
  total_reviews: portfolioTotalReviews,
  avg_score: portfolioAvgScore,
  high_rate: portfolioHighRate,
  cleaning_issue_rate: portfolioCleanRate,
  data_range: { min: dateMin, max: dateMax },
  content_files: ['deep-analysis', 'revenue-impact', 'action-plans']
};
if (existingIdx >= 0) {
  snapshotIndex[existingIdx] = snapshotEntry;
} else {
  snapshotIndex.push(snapshotEntry);
}
snapshotIndex.sort(function(a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : 0; });

// Warn if too many snapshots
if (snapshotIndex.length > 100) {
  console.warn('  WARNING: ' + snapshotIndex.length + ' snapshots stored. Consider running cleanup.');
}

writeJSON('snapshot-index.json', snapshotIndex);

// Build all pages
buildIndex();
buildHotelDashboard();
buildCleaningStrategy();
buildDeepAnalysis();
buildRevenueImpact();
buildActionPlans();

// Copy content HTML fragments to snapshot (after page builds generate them)
['deep-analysis-content.json', 'revenue-impact-content.json', 'action-plans-content.json'].forEach(function(f) {
  var src = path.join(OUTPUT_DIR, 'data', f);
  if (fs.existsSync(src)) fs.copyFileSync(src, path.join(snapshotDir, f));
});

console.log('Done! All pages generated in: ' + OUTPUT_DIR);
