// V2 Hotel Dashboard Page Generator
// Generates hotel-dashboard.html with revenue opportunity badges and target scores
'use strict';

const { esc, nav, footer, pageHead, pageFoot, deltaBadge, deltaBadgeCompact, deltaSummaryBanner } = require('./common-v2');
const { formatYen } = require('./revenue-calc');

var tierColor = { '優秀': '#10B981', '良好': '#3B82F6', '概ね良好': '#F59E0B', '要改善': '#EF4444' };

function buildHotelDashboard(data, revenueOps, deltas) {
  var pov = data.pov;
  var meta = data.meta || {};
  var hotelsRanked = pov.hotels_ranked || [];
  var perHotelTargets = data.perHotelTargets || {};
  var kpiTargets = data.kpiTargets || {};

  // --- KPI values ---
  var totalHotels = hotelsRanked.length;
  var totalReviews = hotelsRanked.reduce(function(sum, h) { return sum + (h.total_reviews || 0); }, 0);
  var avgScore = pov.avg_score || 0;
  var highRate = pov.portfolio_high_rate || 0;
  var lowRate = pov.portfolio_low_rate || 0;

  // KPI target helpers
  function kpiTarget(kpiName) {
    var t = kpiTargets[kpiName];
    if (!t) return '';
    var target = t.target;
    var gap = typeof t.gap !== 'undefined' ? t.gap : null;
    var gapStr = gap !== null ? (gap >= 0 ? '+' + gap : '' + gap) : '';
    return '<div class="kpi-target">目標: ' + target + (gapStr ? ' <span class="achievement ' + (gap >= 0 ? 'good' : gap >= -0.5 ? 'ok' : 'bad') + '">(差: ' + gapStr + ')</span>' : '') + '</div>';
  }

  // --- Extra CSS specific to this page ---
  var extraCSS = [
    '.hotel-card-header { display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 0.75rem; }',
    '.hotel-card-header h3 { font-size: 0.92rem; font-weight: 700; line-height: 1.4; flex: 1; margin: 0; }',
    '.rank-badge { width: 30px; height: 30px; border-radius: 50%; display: flex; align-items: center; justify-content: center; font-size: 0.72rem; font-weight: 800; color: white; flex-shrink: 0; margin-left: 0.5rem; }',
    '.score-row { display: flex; align-items: center; gap: 0.75rem; margin-bottom: 0.75rem; }',
    '.score-big { font-size: 2rem; font-weight: 800; line-height: 1; }',
    '.score-bar-wrap { flex: 1; }',
    '.score-bar-bg { height: 8px; background: #E2E8F0; border-radius: 4px; overflow: hidden; }',
    '.score-bar { height: 100%; border-radius: 4px; transition: width 0.6s ease; }',
    '.score-label { font-size: 0.7rem; color: var(--text-light); margin-top: 3px; display: flex; justify-content: space-between; align-items: center; }',
    '.tier-badge { display: inline-block; padding: 0.15rem 0.5rem; border-radius: 4px; font-size: 0.68rem; font-weight: 700; color: white; }',
    '.stats-row { display: flex; gap: 0.5rem; margin-top: 0.75rem; }',
    '.stat-chip { flex: 1; text-align: center; padding: 0.4rem; border-radius: 6px; background: #F1F5F9; }',
    '.stat-chip .val { font-size: 0.95rem; font-weight: 700; }',
    '.stat-chip .lbl { font-size: 0.6rem; color: var(--text-light); }',
    '.site-dots { display: flex; gap: 0.25rem; margin-top: 0.75rem; flex-wrap: wrap; }',
    '.site-dot { font-size: 0.65rem; padding: 0.15rem 0.5rem; border-radius: 10px; background: #F1F5F9; color: var(--text-light); white-space: nowrap; }',
    '.target-display { font-size: 0.72rem; color: var(--text-light); margin-top: 0.5rem; }',
    '.target-display .gap-positive { color: var(--green); font-weight: 600; }',
    '.target-display .gap-negative { color: var(--red); font-weight: 600; }',
    '.revenue-row { margin-top: 0.5rem; display: flex; align-items: center; justify-content: space-between; }',
    '.dist-chart { display: flex; align-items: flex-end; gap: 4px; height: 120px; padding: 0.5rem 0; }',
    '.dist-bar-wrap { flex: 1; display: flex; flex-direction: column; align-items: center; height: 100%; justify-content: flex-end; }',
    '.dist-bar { width: 100%; min-width: 20px; border-radius: 4px 4px 0 0; }',
    '.dist-label { font-size: 0.65rem; color: var(--text-light); margin-top: 4px; }',
    '.dist-pct { font-size: 0.6rem; font-weight: 600; margin-bottom: 2px; }',
    '.section-title { font-size: 0.9rem; font-weight: 700; color: var(--navy); margin: 1.5rem 0 0.75rem; padding-bottom: 0.5rem; border-bottom: 2px solid var(--accent); }',
    '.section-title:first-child { margin-top: 0; }',
    '.site-table { width: 100%; border-collapse: collapse; font-size: 0.85rem; }',
    '.site-table th { background: #F8FAFC; padding: 0.6rem 0.75rem; text-align: left; font-weight: 600; color: var(--text-light); font-size: 0.75rem; }',
    '.site-table td { padding: 0.6rem 0.75rem; border-top: 1px solid var(--border); }',
    '.site-table tr:hover td { background: #F8FAFC; }',
    '.rate-bars { display: flex; gap: 1rem; flex-direction: column; }',
    '.rate-bar-item { }',
    '.rate-bar-header { display: flex; justify-content: space-between; margin-bottom: 4px; }',
    '.rate-bar-label { font-size: 0.75rem; font-weight: 600; }',
    '.rate-bar-val { font-size: 0.75rem; font-weight: 700; }',
    '.rate-bar-track { height: 10px; background: #E2E8F0; border-radius: 5px; overflow: hidden; }',
    '.rate-bar-fill { height: 100%; border-radius: 5px; }',
    '.review-list { display: flex; flex-direction: column; gap: 0.75rem; max-height: 500px; overflow-y: auto; }',
    '.review-meta { display: flex; gap: 0.75rem; align-items: center; margin-bottom: 0.5rem; flex-wrap: wrap; }',
    '.review-site { font-size: 0.7rem; padding: 0.15rem 0.5rem; border-radius: 4px; background: var(--navy); color: white; font-weight: 600; }',
    '.review-score { font-size: 0.8rem; font-weight: 700; }',
    '.review-date { font-size: 0.7rem; color: var(--text-light); }',
    '.review-text { font-size: 0.8rem; line-height: 1.7; color: var(--text); }',
    '.review-good { color: var(--green); }',
    '.review-bad { color: var(--red); }',
    '#modalTrend { margin-top: 1rem; }',
  ].join('\n');

  // --- Build hotel cards ---
  var cardsHtml = hotelsRanked.map(function(h) {
    var tc = tierColor[h.tier] || '#6B7280';
    var barPct = ((h.avg || 0) / 10 * 100).toFixed(0);
    var rankBg = h.rank <= 3 ? '#F59E0B' : h.rank <= 10 ? '#3B82F6' : '#94A3B8';
    var cleanRate = h.cleaning_issue_rate || 0;
    var cleanColor = cleanRate > 5 ? 'var(--red)' : cleanRate > 0 ? 'var(--orange)' : 'var(--green)';

    // Delta badge for hotel overall score
    var hotelDelta = deltas && deltas.hotels && deltas.hotels[h.key] && deltas.hotels[h.key].overall_avg_10pt;
    var hotelDeltaHtml = deltaBadgeCompact(hotelDelta || null, 'higher');

    // Site dots
    var siteStats = h.site_stats || [];
    var sites = siteStats.map(function(s) {
      var siteDelta = deltas && deltas.hotels && deltas.hotels[h.key] && deltas.hotels[h.key].sites && deltas.hotels[h.key].sites[s.site] && deltas.hotels[h.key].sites[s.site].avg_10pt;
      return '<span class="site-dot">' + esc(s.site) + ' ' + (s.avg_10pt || s.avg || '') + deltaBadgeCompact(siteDelta || null, 'higher') + '</span>';
    }).join('');

    // Revenue opportunity (V2)
    var revOp = revenueOps && revenueOps[h.key] ? revenueOps[h.key] : null;
    var revBadgeHtml = '';
    if (revOp && revOp.monthlyLoss > 0) {
      revBadgeHtml = '<span class="revenue-badge loss">+¥' + formatYen(revOp.monthlyLoss) + '/月</span>';
    }

    // Target score + gap (V2)
    var targetHtml = '';
    if (revOp && revOp.targetScore) {
      var gap = revOp.gap || 0;
      var gapSign = gap > 0 ? '-' : '+';
      var gapAbs = Math.abs(gap);
      var gapClass = gap > 0 ? 'gap-negative' : 'gap-positive';
      targetHtml = '<div class="target-display">目標: ' + revOp.targetScore + ' <span class="' + gapClass + '">(差: ' + gapSign + gapAbs + ')</span></div>';
    } else {
      // Fallback to perHotelTargets
      var pht = perHotelTargets[h.name];
      if (pht && pht.target_avg) {
        var gapVal = Math.round((pht.target_avg - (h.avg || 0)) * 100) / 100;
        var gClass = gapVal > 0 ? 'gap-negative' : 'gap-positive';
        var gSign = gapVal > 0 ? '-' : '+';
        targetHtml = '<div class="target-display">目標: ' + pht.target_avg + ' <span class="' + gClass + '">(差: ' + gSign + Math.abs(gapVal) + ')</span></div>';
      }
    }

    return [
      '<div class="hotel-card" data-key="' + esc(h.key) + '" data-tier="' + esc(h.tier) + '" data-name="' + esc(h.name) + '" data-rank="' + h.rank + '" data-reviews="' + (h.total_reviews || 0) + '" data-high-rate="' + (h.high_rate_pct || h.high_rate || 0) + '" data-low-rate="' + (h.low_rate_pct || h.low_rate || 0) + '" data-cleaning="' + cleanRate + '" onclick="openModal(\'' + esc(h.key) + '\')">',
      '  <div class="hotel-card-header">',
      '    <h3>' + esc(h.name) + '</h3>',
      '    <div class="rank-badge" style="background:' + rankBg + ';">' + h.rank + '</div>',
      '  </div>',
      '  <div class="score-row">',
      '    <div class="score-big" style="color:' + tc + ';">' + (h.avg || 0) + hotelDeltaHtml + '</div>',
      '    <div class="score-bar-wrap">',
      '      <div class="score-bar-bg"><div class="score-bar" style="width:' + barPct + '%;background:' + tc + ';"></div></div>',
      '      <div class="score-label">',
      '        <span><span class="tier-badge" style="background:' + tc + ';">' + esc(h.tier) + '</span></span>',
      '        <span>' + (h.total_reviews || 0) + '件' + deltaBadgeCompact(deltas && deltas.hotels && deltas.hotels[h.key] && deltas.hotels[h.key].total_reviews || null, 'higher') + '</span>',
      '      </div>',
      '    </div>',
      '  </div>',
      '  <div class="stats-row">',
      '    <div class="stat-chip"><div class="val" style="color:var(--green);">' + (h.high_rate_pct || h.high_rate || 0) + '%' + deltaBadgeCompact(deltas && deltas.hotels && deltas.hotels[h.key] && deltas.hotels[h.key].high_rate || null, 'higher') + '</div><div class="lbl">高評価</div></div>',
      '    <div class="stat-chip"><div class="val" style="color:var(--red);">' + (h.low_rate_pct || h.low_rate || 0) + '%' + deltaBadgeCompact(deltas && deltas.hotels && deltas.hotels[h.key] && deltas.hotels[h.key].low_rate || null, 'lower') + '</div><div class="lbl">低評価</div></div>',
      '    <div class="stat-chip"><div class="val" style="color:' + cleanColor + ';">' + cleanRate + '%</div><div class="lbl">清掃課題</div></div>',
      '  </div>',
      '  <div class="site-dots">' + sites + '</div>',
      revBadgeHtml || targetHtml ? '  <div class="revenue-row">' + revBadgeHtml + '</div>' : '',
      targetHtml ? '  ' + targetHtml : '',
      '</div>',
    ].join('\n');
  }).join('\n');

  // --- Tier counts for filter buttons ---
  var tierCounts = {};
  hotelsRanked.forEach(function(h) {
    tierCounts[h.tier] = (tierCounts[h.tier] || 0) + 1;
  });

  // --- KPI target lookups ---
  var avgScoreTarget = kpiTarget('ポートフォリオ平均スコア');
  var highRateTarget = kpiTarget('高評価率');
  var lowRateTarget = kpiTarget('低評価率');
  var totalHotelsTarget = kpiTarget('管理ホテル数');
  var totalReviewsTarget = kpiTarget('総口コミ数');

  // --- Assemble HTML ---
  var lines = [];

  // Head
  lines.push(pageHead('ホテル別口コミダッシュボード - PRIMECHANGE V2', {
    scripts: ['hotel-dashboard-v2.js'],
    extraCSS: extraCSS
  }));

  // Nav
  lines.push(nav('hotel-dashboard'));

  // Container start
  lines.push('<div class="container">');

  // Delta Summary Banner
  lines.push(deltaSummaryBanner(deltas));

  // Page title
  lines.push('  <h1 class="page-title">ホテル別口コミダッシュボード</h1>');
  lines.push('  <p class="page-subtitle">全' + totalHotels + 'ホテルの口コミ評価を一覧比較。カードをクリックして詳細を確認できます。</p>');

  // KPI Grid
  lines.push('  <div class="kpi-grid" id="dashKpiGrid">');

  // KPI 1: Total Hotels
  lines.push('    <div class="kpi-card" data-kpi="total_hotels" style="border-left-color: var(--accent);">');
  lines.push('      <div class="kpi-label">管理ホテル数</div>');
  lines.push('      <div class="kpi-value">' + totalHotels + '</div>');
  lines.push('      <div class="kpi-sub">ホテル</div>');
  lines.push('      ' + deltaBadge(deltas && deltas.hasDeltas && deltas.metrics && deltas.metrics.total_hotels || null, 'higher'));
  lines.push('      ' + totalHotelsTarget);
  lines.push('    </div>');

  // KPI 2: Total Reviews
  lines.push('    <div class="kpi-card" data-kpi="total_reviews" style="border-left-color: var(--green);">');
  lines.push('      <div class="kpi-label">総口コミ数</div>');
  lines.push('      <div class="kpi-value">' + totalReviews.toLocaleString() + '</div>');
  lines.push('      <div class="kpi-sub">件の口コミを分析</div>');
  lines.push('      ' + deltaBadge(deltas && deltas.hasDeltas && deltas.metrics && deltas.metrics.total_reviews || null, 'higher'));
  lines.push('      ' + totalReviewsTarget);
  lines.push('    </div>');

  // KPI 3: Avg Score
  var avgColor = avgScore >= 8 ? 'var(--green)' : avgScore >= 6 ? 'var(--orange)' : 'var(--red)';
  lines.push('    <div class="kpi-card" data-kpi="avg_score" style="border-left-color: ' + avgColor + ';">');
  lines.push('      <div class="kpi-label">ポートフォリオ平均</div>');
  lines.push('      <div class="kpi-value">' + avgScore + '</div>');
  lines.push('      <div class="kpi-sub">/ 10 点</div>');
  lines.push('      ' + deltaBadge(deltas && deltas.hasDeltas && deltas.metrics && deltas.metrics.avg_score || null, 'higher'));
  lines.push('      ' + avgScoreTarget);
  lines.push('    </div>');

  // KPI 4: High Rate
  lines.push('    <div class="kpi-card" data-kpi="high_rate" style="border-left-color: var(--green);">');
  lines.push('      <div class="kpi-label">高評価率</div>');
  lines.push('      <div class="kpi-value">' + highRate + '%</div>');
  lines.push('      <div class="kpi-sub">8点以上の割合</div>');
  lines.push('      ' + deltaBadge(deltas && deltas.hasDeltas && deltas.metrics && deltas.metrics.high_rate || null, 'higher'));
  lines.push('      ' + highRateTarget);
  lines.push('    </div>');

  // KPI 5: Low Rate
  lines.push('    <div class="kpi-card" data-kpi="low_rate" style="border-left-color: var(--red);">');
  lines.push('      <div class="kpi-label">低評価率</div>');
  lines.push('      <div class="kpi-value">' + lowRate + '%</div>');
  lines.push('      <div class="kpi-sub">4点以下の割合</div>');
  lines.push('      ' + deltaBadge(deltas && deltas.hasDeltas && deltas.metrics && deltas.metrics.low_rate || null, 'lower'));
  lines.push('      ' + lowRateTarget);
  lines.push('    </div>');

  lines.push('  </div>'); // kpi-grid

  // Filter bar
  lines.push('  <div class="filter-bar">');
  lines.push('    <button class="filter-btn active" data-filter="all">全て (' + hotelsRanked.length + ')</button>');
  lines.push('    <button class="filter-btn" data-filter="優秀" style="border-color:#10B981;">優秀 (' + (tierCounts['優秀'] || 0) + ')</button>');
  lines.push('    <button class="filter-btn" data-filter="良好" style="border-color:#3B82F6;">良好 (' + (tierCounts['良好'] || 0) + ')</button>');
  lines.push('    <button class="filter-btn" data-filter="概ね良好" style="border-color:#F59E0B;">概ね良好 (' + (tierCounts['概ね良好'] || 0) + ')</button>');
  lines.push('    <button class="filter-btn" data-filter="要改善" style="border-color:#EF4444;">要改善 (' + (tierCounts['要改善'] || 0) + ')</button>');
  lines.push('    <div style="flex:1;"></div>');
  lines.push('    <input type="text" class="search-input" placeholder="ホテル名で検索..." id="searchInput">');
  lines.push('    <select class="sort-select" id="sortSelect">');
  lines.push('      <option value="rank">ランキング順</option>');
  lines.push('      <option value="reviews">口コミ数順</option>');
  lines.push('      <option value="high_rate">高評価率順</option>');
  lines.push('      <option value="low_rate">低評価率順</option>');
  lines.push('      <option value="cleaning">清掃課題率順</option>');
  lines.push('    </select>');
  lines.push('  </div>');

  // Hotel grid
  lines.push('  <div class="hotel-grid" id="hotelGrid">');
  lines.push(cardsHtml);
  lines.push('  </div>');

  // Container end
  lines.push('</div>');

  // Modal overlay
  lines.push('<div class="modal-overlay" id="modalOverlay" onclick="if(event.target===this)closeModal();">');
  lines.push('  <div class="modal-box">');
  lines.push('    <div class="modal-header">');
  lines.push('      <h2 id="modalTitle"></h2>');
  lines.push('      <button class="modal-close" onclick="closeModal();">&#10005;</button>');
  lines.push('    </div>');
  lines.push('    <div class="modal-body" id="modalBody"></div>');
  lines.push('    <div id="modalTrend"></div>');
  lines.push('  </div>');
  lines.push('</div>');

  // Embedded data for client-side JS
  lines.push('<script>');
  lines.push('var hotelDetails = ' + JSON.stringify(data.hotelDetails || {}) + ';');
  lines.push('var hotelRanked = ' + JSON.stringify(hotelsRanked) + ';');
  lines.push('var tierColor = ' + JSON.stringify(tierColor) + ';');
  lines.push('var revenueOps = ' + JSON.stringify(revenueOps || {}) + ';');
  lines.push('</script>');

  // Footer
  lines.push(footer());

  // Page foot
  lines.push(pageFoot());

  var html = lines.join('\n');
  return html;
}

module.exports = { buildHotelDashboard };
