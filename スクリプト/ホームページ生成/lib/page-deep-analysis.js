// V2 Deep Analysis page generator - produces deep-analysis.html
const { esc, nav, footer, pageHead, pageFoot } = require('./common-v2');
const { renderA1, renderA2, renderA3, renderA4, renderA5, renderA6, renderA7 } = require('./deep-analysis-renderers');

var TAB_DEFS = [
  { n: 1, icon: '&#128195;', label: '分析1' },
  { n: 2, icon: '&#128100;', label: '分析2' },
  { n: 3, icon: '&#128101;', label: '分析3' },
  { n: 4, icon: '&#9200;',   label: '分析4' },
  { n: 5, icon: '&#128737;', label: '分析5' },
  { n: 6, icon: '&#128176;', label: '分析6' },
  { n: 7, icon: '&#127942;', label: '分析7' },
];

var RENDERERS = [renderA1, renderA2, renderA3, renderA4, renderA5, renderA6, renderA7];

var CS_AXES = ['接客態度', '立地', '朝食', '設備', '清掃', 'コスパ'];

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function axisCellClass(axisScore) {
  if (!axisScore) return 'cs-cell-neutral';
  var pos = axisScore.positive || 0;
  var neg = axisScore.negative || 0;
  if (pos > neg) return 'cs-cell-strong';
  if (neg > pos) return 'cs-cell-weak';
  return 'cs-cell-neutral';
}

function axisCellContent(axisScore) {
  if (!axisScore) return '-';
  var pos = axisScore.positive || 0;
  var neg = axisScore.negative || 0;
  if (pos === 0 && neg === 0) return '-';
  return '+' + pos + '/-' + neg;
}

function npsColor(nps) {
  if (nps >= 30) return 'var(--green)';
  if (nps >= 0) return 'var(--blue)';
  if (nps >= -30) return 'var(--orange)';
  return 'var(--red)';
}

// ---------------------------------------------------------------------------
// Main builder
// ---------------------------------------------------------------------------

function buildDeepAnalysis(data, csResults, keywordFreq) {
  var analyses = data.analyses || {};
  var html = [];
  var snapshotContent = { tabs: [] };

  // ---- Head ----
  html.push(pageHead('深掘り分析 | PRIMECHANGE V2', {
    scripts: ['deep-analysis-v2.js'],
    extraCSS: [
      '.da-tabs { display: flex; gap: 0.25rem; border-bottom: 2px solid var(--border); margin-bottom: 1.5rem; flex-wrap: wrap; }',
      '.da-tab-btn { padding: 0.6rem 1rem; border: none; background: none; cursor: pointer; font-size: 0.8rem; font-weight: 600; color: var(--text-light); border-bottom: 2px solid transparent; margin-bottom: -2px; transition: all 0.2s; white-space: nowrap; }',
      '.da-tab-btn:hover { color: var(--accent); }',
      '.da-tab-btn.active { color: var(--accent); border-bottom-color: var(--accent); }',
      '.da-tab-panel { display: none; }',
      '.da-tab-panel.active { display: block; }',
      '.cs-section { margin-top: 2.5rem; }',
      '.kw-bar { display: flex; align-items: center; gap: 0.5rem; margin-bottom: 0.4rem; }',
      '.kw-bar-label { font-size: 0.78rem; width: 140px; text-align: right; flex-shrink: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }',
      '.kw-bar-track { flex: 1; height: 22px; background: #E2E8F0; border-radius: 4px; overflow: hidden; }',
      '.kw-bar-fill { height: 100%; border-radius: 4px; display: flex; align-items: center; padding-left: 6px; transition: width 0.6s ease; background: var(--accent); }',
      '.kw-bar-val { font-size: 0.65rem; font-weight: 700; color: white; white-space: nowrap; }',
      '.kw-axis-badge { font-size: 0.6rem; padding: 0.1rem 0.4rem; border-radius: 4px; background: #F1F5F9; color: var(--text-light); font-weight: 600; white-space: nowrap; }',
    ].join('\n'),
  }));

  // ---- Nav ----
  html.push(nav('deep-analysis'));

  // ---- Container ----
  html.push('<div class="container">');

  // ---- Title ----
  html.push('<h1 class="page-title">深掘り分析</h1>');
  html.push('<p class="page-subtitle">7つの視点から清掃品質を多角的に深掘り分析します。</p>');

  // ---- Full-data banner ----
  html.push('<div class="fulldata-banner"><span>&#9888;&#65039;</span><div>このページは全期間の口コミ・清掃データに基づいた分析結果を表示しています。日付フィルターは適用されません。</div></div>');

  // ---- Tab interface ----
  html.push('<div id="da-content">');

  // Tab buttons
  html.push('<div class="da-tabs">');
  TAB_DEFS.forEach(function(tab, idx) {
    var cls = idx === 0 ? ' class="da-tab-btn active"' : ' class="da-tab-btn"';
    html.push('<button' + cls + ' data-tab="da-panel-' + tab.n + '" onclick="switchDaTab(this)">' + tab.icon + ' ' + tab.label + '</button>');
  });
  html.push('</div>');

  // Tab panels
  TAB_DEFS.forEach(function(tab, idx) {
    var renderer = RENDERERS[idx];
    var analysisData = analyses[tab.n] || analyses[String(tab.n)] || null;
    var panelHtml = renderer(analysisData);

    var cls = idx === 0 ? ' class="da-tab-panel active"' : ' class="da-tab-panel"';
    html.push('<div id="da-panel-' + tab.n + '"' + cls + '>');
    html.push(panelHtml);
    html.push('</div>');

    // Store for snapshot
    snapshotContent.tabs.push({
      id: 'da-panel-' + tab.n,
      icon: tab.icon,
      html: panelHtml,
    });
  });

  html.push('</div>'); // #da-content

  // ====================================================================
  // CS分析セクション
  // ====================================================================
  html.push('<div class="cs-section">');
  html.push('<div class="card">');
  html.push('<div class="card-title">CS 6軸分析 ＆ NPS推定</div>');

  // --- 5a. NPS overview KPIs ---
  if (csResults) {
    var hotelKeys = Object.keys(csResults);
    var totalNps = 0;
    var totalPromoters = 0;
    var totalDetractors = 0;
    var totalPassives = 0;
    var npsCount = 0;

    hotelKeys.forEach(function(key) {
      var r = csResults[key];
      if (r && r.nps != null) {
        totalNps += r.nps;
        npsCount++;
      }
      if (r && r.npsBreakdown) {
        totalPromoters += r.npsBreakdown.promoters || 0;
        totalDetractors += r.npsBreakdown.detractors || 0;
        totalPassives += r.npsBreakdown.passives || 0;
      } else if (r) {
        // Fallback: count from nps classification if breakdown not available
        if (r.nps != null && r.nps > 0) totalPromoters++;
        else if (r.nps != null && r.nps < 0) totalDetractors++;
        else totalPassives++;
      }
    });

    var avgNps = npsCount > 0 ? (totalNps / npsCount) : 0;

    html.push('<div class="kpi-grid">');
    html.push('<div class="kpi-card" style="border-left-color:' + npsColor(avgNps) + ';"><div class="kpi-label">ポートフォリオNPS</div><div class="kpi-value" style="color:' + npsColor(avgNps) + ';">' + avgNps.toFixed(1) + '</div><div class="kpi-sub">' + npsCount + 'ホテル平均</div></div>');
    html.push('<div class="kpi-card" style="border-left-color:var(--green);"><div class="kpi-label">推奨者数</div><div class="kpi-value">' + totalPromoters + '</div><div class="kpi-sub">Promoters</div></div>');
    html.push('<div class="kpi-card" style="border-left-color:var(--red);"><div class="kpi-label">批判者数</div><div class="kpi-value">' + totalDetractors + '</div><div class="kpi-sub">Detractors</div></div>');
    html.push('<div class="kpi-card" style="border-left-color:var(--text-light);"><div class="kpi-label">中立者数</div><div class="kpi-value">' + totalPassives + '</div><div class="kpi-sub">Passives</div></div>');
    html.push('</div>');

    // --- 5b. CS 6軸 × 19ホテル マトリクス ---
    html.push('<h3 style="font-size:0.95rem;font-weight:700;color:var(--navy);margin:1.5rem 0 0.75rem;">CS 6軸 × 19ホテル マトリクス</h3>');
    html.push('<div class="cs-matrix"><table>');

    // Header
    html.push('<thead><tr><th style="text-align:left;">ホテル</th>');
    CS_AXES.forEach(function(axis) {
      html.push('<th>' + esc(axis) + '</th>');
    });
    html.push('<th>NPS</th></tr></thead>');

    // Body
    html.push('<tbody>');
    hotelKeys.forEach(function(key) {
      var r = csResults[key];
      if (!r) return;
      var hotelName = r.hotelName || r.name || key;
      var axisScores = r.axisScores || {};

      html.push('<tr>');
      html.push('<td style="text-align:left;font-weight:600;font-size:0.75rem;white-space:nowrap;">' + esc(hotelName) + '</td>');

      CS_AXES.forEach(function(axis) {
        var score = axisScores[axis] || null;
        var cellCls = axisCellClass(score);
        var cellContent = axisCellContent(score);
        html.push('<td class="' + cellCls + '">' + cellContent + '</td>');
      });

      // NPS column
      var npsVal = r.nps != null ? r.nps.toFixed(1) : '-';
      var npsClr = r.nps != null ? npsColor(r.nps) : 'var(--text-light)';
      html.push('<td style="font-weight:700;color:' + npsClr + ';">' + npsVal + '</td>');

      html.push('</tr>');
    });
    html.push('</tbody></table></div>');
  } else {
    html.push('<p style="color:var(--text-light);font-size:0.85rem;">CS分析データがありません。</p>');
  }

  // --- 5c. キーワード頻度 Top20 ---
  if (keywordFreq && keywordFreq.length > 0) {
    html.push('<h3 style="font-size:0.95rem;font-weight:700;color:var(--navy);margin:1.5rem 0 0.75rem;">キーワード頻度 Top20</h3>');

    var top20 = keywordFreq.slice(0, 20);
    var maxCount = top20.length > 0 ? top20[0].count : 1;

    top20.forEach(function(kw) {
      var w = maxCount > 0 ? Math.min(kw.count / maxCount * 100, 100) : 0;
      var axisBadge = kw.axis ? '<span class="kw-axis-badge">' + esc(kw.axis) + '</span>' : '';

      html.push('<div class="kw-bar">');
      html.push('<div class="kw-bar-label" title="' + esc(kw.keyword || kw.word || '') + '">' + esc(kw.keyword || kw.word || '') + '</div>');
      html.push('<div class="kw-bar-track"><div class="kw-bar-fill" style="width:' + w.toFixed(1) + '%;"><span class="kw-bar-val">' + kw.count + '</span></div></div>');
      html.push(axisBadge);
      html.push('</div>');
    });
  }

  html.push('</div>'); // .card
  html.push('</div>'); // .cs-section

  // ---- Tab-switching script (inline) ----
  html.push('<script>');
  html.push('function switchDaTab(btn) {');
  html.push('  var tabs = document.querySelectorAll(".da-tab-btn");');
  html.push('  var panels = document.querySelectorAll(".da-tab-panel");');
  html.push('  tabs.forEach(function(t) { t.classList.remove("active"); });');
  html.push('  panels.forEach(function(p) { p.classList.remove("active"); });');
  html.push('  btn.classList.add("active");');
  html.push('  var targetId = btn.getAttribute("data-tab");');
  html.push('  var target = document.getElementById(targetId);');
  html.push('  if (target) target.classList.add("active");');
  html.push('}');
  html.push('</script>');

  // ---- Footer ----
  html.push('</div>'); // .container
  html.push(footer());
  html.push(pageFoot());

  return { html: html.join('\n'), snapshotContent: snapshotContent };
}

module.exports = { buildDeepAnalysis };
