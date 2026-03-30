// V2 Action Plans page builder
// Generates action-plans.html with action execution management (改善4)

var { esc, nav, footer, pageHead, pageFoot, deltaBadge, deltaBadgeCompact } = require('./common-v2');
var { formatYen } = require('./revenue-calc');

var phaseColors = ['var(--red)', 'var(--orange)', 'var(--blue)'];
var phaseLabels = ['フェーズ1', 'フェーズ2', 'フェーズ3'];
var phaseKeys = ['phase1_immediate', 'phase2_short_term', 'phase3_medium_term'];
var phaseROISplit = [0.5, 0.3, 0.2]; // proportional ROI split across phases

var priorityBadge = {
  URGENT:      '<span class="badge badge-red">URGENT</span>',
  HIGH:        '<span class="badge badge-orange">HIGH</span>',
  STANDARD:    '<span class="badge badge-blue">STANDARD</span>',
  MAINTENANCE: '<span class="badge badge-green">MAINTENANCE</span>'
};

function findHotelKey(hotelsRanked, hotelName) {
  if (!hotelsRanked || !hotelName) return '';
  for (var i = 0; i < hotelsRanked.length; i++) {
    if (hotelsRanked[i].name === hotelName) return hotelsRanked[i].key || '';
  }
  // Fallback: sanitise name to key
  return hotelName.replace(/[^a-zA-Z0-9]/g, '_').toLowerCase();
}

function buildActionPlans(data, revenueOps, deltas) {
  var plans = data.actionPlans || [];
  var priMatrix = data.priMatrix || {};
  var kpi = data.kpi || {};
  var hotelsRanked = (data.pov && data.pov.hotels_ranked) ? data.pov.hotels_ranked : [];

  // Group by priority
  var groups = { URGENT: [], HIGH: [], STANDARD: [], MAINTENANCE: [] };
  plans.forEach(function(p) {
    if (groups[p.priority_level]) groups[p.priority_level].push(p);
    else groups.STANDARD.push(p);
  });

  // Count priorities from priMatrix
  var urgentCount = (priMatrix.urgent || []).length;
  var highCount = (priMatrix.high || []).length;
  var standardCount = (priMatrix.standard || []).length;
  var maintenanceCount = (priMatrix.maintenance || []).length;

  // Build action status object and collect hotel names for filter
  var actionStatus = {};
  var hotelNames = [];
  var totalActionCount = 0;

  plans.forEach(function(p) {
    var hotelKey = findHotelKey(hotelsRanked, p.hotel);
    if (hotelNames.indexOf(p.hotel) === -1) hotelNames.push(p.hotel);

    phaseKeys.forEach(function(pk, pi) {
      var phase = p[pk];
      if (!phase || !phase.actions) return;
      phase.actions.forEach(function(a, ai) {
        var actionId = hotelKey + '_p' + (pi + 1) + '_a' + ai;
        actionStatus[actionId] = { status: 'not-started', assignee: '', deadline: '' };
        totalActionCount++;
      });
    });
  });

  // Calculate total portfolio revenue loss
  var totalPortfolioLoss = 0;
  if (revenueOps) {
    Object.keys(revenueOps).forEach(function(k) {
      totalPortfolioLoss += (revenueOps[k].monthlyLoss || 0);
    });
  }

  // --- Render helpers ---

  function renderPhase(phase, phaseNum, hotelKey, revOp) {
    if (!phase) return '';
    var actions = phase.actions || [];
    if (actions.length === 0) return '';

    // Expected ROI for this phase
    var phaseROI = 0;
    if (revOp && revOp.monthlyLoss > 0) {
      phaseROI = Math.round(revOp.monthlyLoss * phaseROISplit[phaseNum - 1]);
    }

    var items = actions.map(function(a, ai) {
      var actionId = hotelKey + '_p' + phaseNum + '_a' + ai;
      return '<li data-action-id="' + esc(actionId) + '">'
        + '<span class="action-text">' + esc(a.action)
        + (a.category ? ' <span class="action-cat">' + esc(a.category) + '</span>' : '')
        + '</span>'
        + '<span class="status-badge not-started" data-status="not-started" data-action-id="' + esc(actionId) + '">'
        + '\u672A\u7740\u624B</span>'
        + '</li>';
    }).join('\n');

    var roiBadge = phaseROI > 0
      ? ' <span class="revenue-badge" style="margin-left:0.5rem;">\u671F\u5F85ROI: &yen;' + esc(formatYen(phaseROI)) + '/\u6708</span>'
      : '';

    return [
      '<div class="phase">',
      '  <div class="phase-header">',
      '    <div class="phase-num" style="background:' + phaseColors[phaseNum - 1] + ';">' + phaseNum + '</div>',
      '    <div class="phase-title">' + phaseLabels[phaseNum - 1] + roiBadge + '</div>',
      '    <div class="phase-timeline">' + esc(phase.timeline || '') + '</div>',
      '  </div>',
      '  <ul class="action-list">' + items + '</ul>',
      '</div>'
    ].join('\n');
  }

  function renderPlan(p) {
    var hotelKey = findHotelKey(hotelsRanked, p.hotel);
    var revOp = revenueOps ? revenueOps[hotelKey] : null;
    var revBadge = '';
    if (revOp && revOp.monthlyLoss > 0) {
      revBadge = ' <span class="revenue-badge loss">\u640D\u5931: &yen;' + esc(formatYen(revOp.monthlyLoss)) + '/\u6708</span>';
    }

    return [
      '<div class="accordion-item" data-hotel="' + esc(p.hotel) + '" data-priority="' + esc(p.priority_level || 'STANDARD') + '">',
      '  <div class="accordion-header">',
      '    <div>' + (priorityBadge[p.priority_level] || '') + ' ' + esc(p.hotel)
        + ' <span style="font-size:0.75rem;color:var(--text-light);">(' + esc(String(p.current_avg)) + deltaBadgeCompact(deltas && deltas.hotels && deltas.hotels[hotelKey] && deltas.hotels[hotelKey].overall_avg_10pt || null, 'higher') + ' &rarr; ' + esc(String(p.target_avg)) + ')</span>'
        + revBadge + '</div>',
      '    <span class="accordion-arrow">&#9660;</span>',
      '  </div>',
      '  <div class="accordion-body">',
      renderPhase(p.phase1_immediate, 1, hotelKey, revOp),
      renderPhase(p.phase2_short_term, 2, hotelKey, revOp),
      renderPhase(p.phase3_medium_term, 3, hotelKey, revOp),
      '  </div>',
      '</div>'
    ].join('\n');
  }

  // --- Build content ---
  var content = [];

  // Progress summary bar (V2)
  content.push(
    '<div class="card">',
    '  <div class="progress-summary"><span>\u30A2\u30AF\u30B7\u30E7\u30F3\u9032\u6357</span><span class="progress-pct" id="progressPct">0%</span></div>',
    '  <div class="progress-bar-wrap"><div class="progress-bar-fill" id="progressBar" style="width:0%;"></div></div>',
    '  <div style="font-size:0.8rem;color:var(--text-light);margin-top:0.3rem;">',
    '    \u5B8C\u4E86: <span id="completedCount">0</span> / \u5168 <span id="totalCount">' + totalActionCount + '</span> \u30A2\u30AF\u30B7\u30E7\u30F3',
    '  </div>',
    '</div>'
  );

  // Filter bar (V2)
  var hotelOptions = hotelNames.map(function(n) {
    return '<option value="' + esc(n) + '">' + esc(n) + '</option>';
  }).join('');

  content.push(
    '<div class="action-filter-bar">',
    '  <select id="filterStatus">',
    '    <option value="all">\u5168\u3066</option>',
    '    <option value="not-started">\u672A\u7740\u624B</option>',
    '    <option value="in-progress">\u9032\u884C\u4E2D</option>',
    '    <option value="completed">\u5B8C\u4E86</option>',
    '  </select>',
    '  <select id="filterHotel">',
    '    <option value="all">\u5168\u30DB\u30C6\u30EB</option>',
    hotelOptions,
    '  </select>',
    '  <select id="filterPhase">',
    '    <option value="all">\u5168\u30D5\u30A7\u30FC\u30BA</option>',
    '    <option value="1">\u30D5\u30A7\u30FC\u30BA1</option>',
    '    <option value="2">\u30D5\u30A7\u30FC\u30BA2</option>',
    '    <option value="3">\u30D5\u30A7\u30FC\u30BA3</option>',
    '  </select>',
    '  <button class="filter-btn" id="exportBtn" style="margin-left:auto;">\u30A8\u30AF\u30B9\u30DD\u30FC\u30C8</button>',
    '  <label class="filter-btn" style="cursor:pointer;">\u30A4\u30F3\u30DD\u30FC\u30C8<input type="file" id="importBtn" accept=".json" style="display:none;"></label>',
    '</div>'
  );

  // KPI Grid
  content.push(
    '<div class="kpi-grid">',
    '  <div class="kpi-card" style="border-left-color:var(--red);"><div class="kpi-label">\u7DCA\u6025 (URGENT)</div><div class="kpi-value">' + urgentCount + '</div></div>',
    '  <div class="kpi-card" style="border-left-color:var(--orange);"><div class="kpi-label">\u8981\u6CE8\u610F (HIGH)</div><div class="kpi-value">' + highCount + '</div></div>',
    '  <div class="kpi-card" style="border-left-color:var(--blue);"><div class="kpi-label">\u6A19\u6E96 (STANDARD)</div><div class="kpi-value">' + standardCount + '</div></div>',
    '  <div class="kpi-card" style="border-left-color:var(--green);"><div class="kpi-label">\u7DAD\u6301 (MAINTENANCE)</div><div class="kpi-value">' + maintenanceCount + '</div></div>',
    '</div>'
  );

  // Portfolio KPI targets table
  var portfolioTargets = (kpi && kpi.portfolio_targets) ? kpi.portfolio_targets : [];
  if (Array.isArray(portfolioTargets) && portfolioTargets.length > 0) {
    var kpiRows = portfolioTargets.map(function(t) {
      return '<tr>'
        + '<td><strong>' + esc(t.kpi || '') + '</strong></td>'
        + '<td>' + esc(String(t.current || '')) + '</td>'
        + '<td style="color:var(--accent);font-weight:700;">' + esc(String(t.target || '')) + '</td>'
        + '<td>' + esc(t.deadline || '') + '</td>'
        + '</tr>';
    }).join('\n');

    content.push(
      '<div class="card">',
      '  <div class="card-title">&#127919; \u30DD\u30FC\u30C8\u30D5\u30A9\u30EA\u30AA KPI\u76EE\u6A19</div>',
      '  <div style="overflow-x:auto;">',
      '    <table class="data-table">',
      '      <thead><tr><th>KPI</th><th>\u73FE\u72B6</th><th>\u76EE\u6A19</th><th>\u671F\u9650</th></tr></thead>',
      '      <tbody>' + kpiRows + '</tbody>',
      '    </table>',
      '  </div>',
      '</div>'
    );
  }

  // Hotel accordions by priority group
  var groupDefs = [
    ['URGENT', '\u7DCA\u6025\u5BFE\u5FDC', 'var(--red)'],
    ['HIGH', '\u8981\u6CE8\u610F', 'var(--orange)'],
    ['STANDARD', '\u6A19\u6E96', 'var(--blue)'],
    ['MAINTENANCE', '\u7DAD\u6301', 'var(--green)']
  ];

  groupDefs.forEach(function(g) {
    var level = g[0], label = g[1], color = g[2];
    if (groups[level].length === 0) return;
    content.push(
      '<h2 style="font-size:1rem;font-weight:700;color:' + color + ';margin:1.5rem 0 0.75rem;">'
        + label + ' (' + groups[level].length + '\u30DB\u30C6\u30EB)</h2>'
    );
    groups[level].forEach(function(p) { content.push(renderPlan(p)); });
  });

  var contentHtml = content.join('\n');

  // Snapshot content object
  var snapshotContent = { html: contentHtml };

  // Extra CSS for this page
  var extraCSS = [
    '.action-text { flex: 1; }',
    '.export-import-group { display: flex; gap: 0.4rem; }',
  ].join('\n');

  // Full page HTML
  var html = [
    pageHead('\u30A2\u30AF\u30B7\u30E7\u30F3\u30D7\u30E9\u30F3\u7BA1\u7406 - PRIMECHANGE V2', { scripts: ['action-tracker.js'], extraCSS: extraCSS }),
    nav('action-plans'),
    '<div class="container">',
    '  <h1 class="page-title">\u30A2\u30AF\u30B7\u30E7\u30F3\u30D7\u30E9\u30F3\u7BA1\u7406</h1>',
    '  <p class="page-subtitle">19\u30DB\u30C6\u30EB\u5225 3\u30D5\u30A7\u30FC\u30BA\u6539\u5584\u8A08\u753B\u30FB\u5B9F\u884C\u7BA1\u7406</p>',
    '',
    '  <div class="fulldata-banner">',
    '    <span>&#9432;</span>',
    '    <div>\u30A2\u30AF\u30B7\u30E7\u30F3\u30D7\u30E9\u30F3\u306F\u5168\u671F\u9593\u306E\u5206\u6790\u7D50\u679C\u306B\u57FA\u3065\u3044\u3066\u3044\u307E\u3059\u3002\u65E5\u4ED8\u30D5\u30A3\u30EB\u30BF\u30FC\u306F\u9069\u7528\u3055\u308C\u307E\u305B\u3093\u3002</div>',
    '  </div>',
    '',
    '  <div id="ap-content">',
    contentHtml,
    '  </div>',
    '</div>',
    footer(),
    pageFoot()
  ].join('\n');

  return {
    html: html,
    snapshotContent: snapshotContent,
    actionStatus: actionStatus
  };
}

module.exports = { buildActionPlans };
