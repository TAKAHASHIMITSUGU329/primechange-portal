// V2 Cleaning Strategy page builder
// Generates cleaning-strategy.html with revenue impact on category bars

var { esc, nav, footer, pageHead, pageFoot } = require('./common-v2');
var { formatYen } = require('./revenue-calc');

var severityColor = {
  'CRITICAL': '#EF4444',
  'HIGH': '#F59E0B',
  'MEDIUM': '#3B82F6',
  'LOW': '#6B7280'
};

var priorityColor = {
  'URGENT': '#EF4444',
  'HIGH': '#F59E0B',
  'STANDARD': '#3B82F6',
  'MAINTENANCE': '#10B981'
};

function buildCleaningStrategy(data, revenueOps) {
  var cleanDive = data.cleanDive || {};
  var priMatrix = data.priMatrix || {};
  var crossRec = data.crossRec || [];

  var cats = cleanDive.category_summary || [];
  var matrix = cleanDive.hotel_cleaning_matrix || [];
  var allCats = cats.map(function(c) { return c.category; });
  var urgentCount = (priMatrix.urgent || []).length;

  // Calculate total cleaning issues for revenue proportioning
  var totalCleaningIssues = 0;
  cats.forEach(function(c) { totalCleaningIssues += (c.total_mentions || 0); });

  // Estimate average monthly revenue loss across portfolio from revenueOps
  var totalPortfolioLoss = 0;
  if (revenueOps) {
    var keys = Object.keys(revenueOps);
    keys.forEach(function(k) {
      totalPortfolioLoss += (revenueOps[k].monthlyLoss || 0);
    });
  }

  // Category bars with V2 revenue badges
  var maxMention = cats.length > 0
    ? Math.max.apply(null, cats.map(function(c) { return c.total_mentions || 0; }))
    : 1;

  var catBars = cats.map(function(c) {
    var mentions = c.total_mentions || 0;
    var pct = (mentions / maxMention * 100).toFixed(0);
    var col = severityColor[c.severity] || '#6B7280';
    var hotelsAffected = c.hotels_affected || 0;

    // V2: proportional revenue loss for this category
    var catShare = totalCleaningIssues > 0 ? mentions / totalCleaningIssues : 0;
    var catLoss = Math.round(totalPortfolioLoss * catShare);

    var barHtml = '<div class="h-bar">'
      + '<div class="h-bar-label">' + esc(c.category) + '</div>'
      + '<div class="h-bar-track">'
      + '<div class="h-bar-fill" style="width:' + pct + '%;background:' + col + ';">'
      + '<span class="h-bar-val">' + mentions + '件 (' + hotelsAffected + 'ホテル)</span>'
      + '</div></div></div>';

    // V2 revenue badge
    if (catLoss > 0) {
      barHtml += '<div style="margin-left:126px;margin-top:-0.2rem;margin-bottom:0.5rem;">'
        + '<span class="revenue-badge loss">推定損失: &yen;' + esc(formatYen(catLoss)) + '/月</span>'
        + '</div>';
    }

    return barHtml;
  }).join('\n');

  // Heatmap
  var heatmapHtml = '';
  if (matrix.length > 0 && allCats.length > 0) {
    var heatmapHead = '<tr><th class="row-header">ホテル名</th><th>スコア</th><th>課題率</th>'
      + allCats.map(function(c) { return '<th>' + esc(c) + '</th>'; }).join('')
      + '<th>合計</th></tr>';

    var heatmapRows = matrix.map(function(h) {
      var cells = allCats.map(function(cat) {
        var val = (h.categories && h.categories[cat]) || 0;
        var bg = val === 0 ? '#FFFFFF' : val <= 1 ? '#FEF3C7' : val <= 3 ? '#FBBF24' : val <= 5 ? '#F97316' : '#EF4444';
        var fc = val >= 4 ? 'white' : 'var(--text)';
        return '<td style="background:' + bg + ';color:' + fc + ';">' + (val || '') + '</td>';
      }).join('');
      var pc = priorityColor[h.priority] || '#6B7280';
      return '<tr>'
        + '<td class="row-header">' + esc(h.name) + '</td>'
        + '<td>' + esc(String(h.avg || '')) + '</td>'
        + '<td style="color:' + pc + ';font-weight:700;">' + esc(String(h.cleaning_issue_rate || 0)) + '%</td>'
        + cells
        + '<td style="font-weight:700;">' + (h.cleaning_issue_count || 0) + '</td>'
        + '</tr>';
    }).join('\n');

    heatmapHtml = [
      '  <div class="card">',
      '    <div class="card-title">&#128293; ホテル&times;課題カテゴリ ヒートマップ</div>',
      '    <div class="heatmap-wrap"><table class="heatmap">' + heatmapHead + heatmapRows + '</table></div>',
      '  </div>'
    ].join('\n');
  }

  // Priority matrix cards
  function priCards(level, items) {
    if (!items || items.length === 0) return '<p style="font-size:0.8rem;color:var(--text-light);">該当なし</p>';
    var cls = level.toLowerCase();
    return items.map(function(h) {
      var hotelName = h.hotel || h.name || '';
      var score = h.avg || h.current_score || '';
      var issues = h.cleaning_issues || h.cleaning_issue_count || 0;
      var rate = h.cleaning_rate || h.cleaning_issue_rate || 0;
      var problems = h.key_problems || [];
      var problemText = problems.length > 0 ? problems.map(esc).join(', ') : '';

      return '<div class="priority-card ' + cls + '">'
        + '<div class="priority-title">' + esc(hotelName)
        + ' <span style="font-size:0.75rem;color:var(--text-light);">(' + esc(String(score)) + '点)</span></div>'
        + '<div class="priority-hotels">課題: ' + issues + '件 (' + esc(String(rate)) + '%)'
        + (problemText ? ' / ' + problemText : '')
        + '</div></div>';
    }).join('\n');
  }

  // Cross-cutting recommendations
  var recHtml = '';
  if (crossRec && Array.isArray(crossRec) && crossRec.length) {
    recHtml = crossRec.map(function(r, i) {
      var title = r.title || r.theme || '';
      var desc = r.description || r.detail || '';
      var actionsHtml = '';
      if (r.actions && r.actions.length) {
        actionsHtml = '<ul style="font-size:0.8rem;padding-left:1.2rem;margin-top:0.5rem;">'
          + r.actions.map(function(a) { return '<li style="margin-bottom:0.3rem;">' + esc(a) + '</li>'; }).join('\n')
          + '</ul>';
      }
      return '<div class="card">'
        + '<div class="card-title">施策' + (i + 1) + ': ' + esc(title) + '</div>'
        + '<p style="font-size:0.82rem;color:var(--text-light);line-height:1.7;">' + esc(desc) + '</p>'
        + actionsHtml
        + '</div>';
    }).join('\n');
  } else if (crossRec && typeof crossRec === 'object' && !Array.isArray(crossRec)) {
    recHtml = Object.keys(crossRec).map(function(k, i) {
      var v = crossRec[k];
      var title = v.title || k;
      var desc = v.description || v.detail || (typeof v === 'string' ? v : JSON.stringify(v).slice(0, 500));
      var actionsHtml = '';
      if (v.actions && v.actions.length) {
        actionsHtml = '<ul style="font-size:0.8rem;padding-left:1.2rem;margin-top:0.5rem;">'
          + v.actions.map(function(a) { return '<li style="margin-bottom:0.3rem;">' + esc(a) + '</li>'; }).join('\n')
          + '</ul>';
      }
      return '<div class="card">'
        + '<div class="card-title">施策' + (i + 1) + ': ' + esc(title) + '</div>'
        + '<p style="font-size:0.82rem;color:var(--text-light);line-height:1.7;">' + esc(desc) + '</p>'
        + actionsHtml
        + '</div>';
    }).join('\n');
  }

  // Extra CSS for this page
  var extraCSS = [
    '.severity-legend { margin-top:0.75rem; font-size:0.7rem; color:var(--text-light); }',
    '.severity-legend span { margin-right:0.75rem; }',
  ].join('\n');

  var html = [
    pageHead('清掃品質改善戦略 - PRIMECHANGE V2', { extraCSS: extraCSS }),
    nav('cleaning-strategy'),
    '<div class="container">',
    '  <h1 class="page-title">清掃品質改善戦略</h1>',
    '  <p class="page-subtitle">カテゴリ別課題分析・収益インパクト・優先度マトリクスによる戦略的清掃改善</p>',
    '',
    '  <div class="fulldata-banner">',
    '    <span>&#9888;&#65039;</span>',
    '    <div>この分析は全期間データがベースです。日付フィルターは口コミ件数にのみ適用されます。</div>',
    '  </div>',
    '',
    '  <div class="kpi-grid" id="cleaningKpiGrid">',
    '    <div class="kpi-card" style="border-left-color:var(--red);"><div class="kpi-label">清掃クレーム率</div><div class="kpi-value" data-kpi="cleaning_issue_rate">' + esc(String(cleanDive.portfolio_cleaning_issue_rate || cleanDive.overall_cleaning_issue_rate || 0)) + '%</div></div>',
    '    <div class="kpi-card" style="border-left-color:var(--orange);"><div class="kpi-label">清掃クレーム件数</div><div class="kpi-value" data-kpi="cleaning_issue_count">' + (cleanDive.total_cleaning_mentions || cleanDive.total_cleaning_issues || 0) + '</div></div>',
    '    <div class="kpi-card" style="border-left-color:var(--purple);"><div class="kpi-label">カテゴリ数</div><div class="kpi-value">' + cats.length + '</div></div>',
    '    <div class="kpi-card" style="border-left-color:var(--red);"><div class="kpi-label">緊急対応ホテル</div><div class="kpi-value">' + urgentCount + '</div></div>',
    '  </div>',
    '',
    '  <div class="card">',
    '    <div class="card-title">&#128200; 清掃課題発生トレンド</div>',
    '    <div id="cleaningTrend" style="width:100%;"></div>',
    '  </div>',
    '',
    '  <div class="card">',
    '    <div class="card-title">&#128202; 課題カテゴリ別件数 &amp; 推定収益損失</div>',
    catBars || '    <p style="font-size:0.82rem;color:var(--text-light);">カテゴリデータがありません</p>',
    '    <div class="severity-legend">',
    '      色: <span style="color:#EF4444;">&#9632; CRITICAL</span>',
    '      <span style="color:#F59E0B;">&#9632; HIGH</span>',
    '      <span style="color:#3B82F6;">&#9632; MEDIUM</span>',
    '      <span style="color:#6B7280;">&#9632; LOW</span>',
    '    </div>',
    totalPortfolioLoss > 0
      ? '    <div style="margin-top:0.75rem;padding:0.75rem;background:#FEF2F2;border-radius:8px;font-size:0.82rem;text-align:center;">'
        + '<strong>清掃課題による推定総損失:</strong> &yen;' + esc(formatYen(totalPortfolioLoss)) + '/月'
        + '</div>'
      : '',
    '  </div>',
    '',
    heatmapHtml,
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
    recHtml || '    <p style="font-size:0.82rem;color:var(--text-light);">データなし</p>',
    '  </div>',
    '</div>',
    footer(),
    pageFoot()
  ];

  return html.join('\n');
}

module.exports = { buildCleaningStrategy };
