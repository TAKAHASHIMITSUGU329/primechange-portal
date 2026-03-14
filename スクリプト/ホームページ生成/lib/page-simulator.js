// V3 Revenue Simulator Page Generator
'use strict';
var common = require('./common-v2');
var esc = common.esc, nav = common.nav, footer = common.footer, pageHead = common.pageHead, pageFoot = common.pageFoot;
var formatYen = require('./revenue-calc').formatYen;

function buildSimulator(data, revenueOps) {
  var analysis6 = data.analyses ? data.analyses[6] || {} : {};
  var regression = analysis6.regression_results || {};
  var thresholdGroups = (analysis6.threshold_analysis && analysis6.threshold_analysis.groups) || [];
  var revenueData = data.revenueData || {};
  var hotelsRanked = (data.pov && data.pov.hotels_ranked) || [];
  var revparSlope = (regression.score_vs_revpar && regression.score_vs_revpar.slope) || 228.945845;

  // --- Build sorted opportunity list ---
  var opsList = [];
  Object.keys(revenueOps || {}).forEach(function(k) {
    var op = revenueOps[k];
    opsList.push({
      key: k,
      name: op.name || k,
      currentScore: op.currentScore || 0,
      targetScore: op.targetScore || 8.89,
      gap: op.scoreDelta || (op.targetScore - op.currentScore) || 0,
      roomCount: op.roomCount || 0,
      monthlyLoss: op.monthlyLoss || 0,
      revparGain: op.revparGain || 0
    });
  });
  opsList.sort(function(a, b) { return b.monthlyLoss - a.monthlyLoss; });

  // --- Calculate portfolio averages ---
  var totalScore = 0, hotelCount = 0;
  hotelsRanked.forEach(function(h) {
    totalScore += (h.avg || 0);
    hotelCount++;
  });
  var avgScore = hotelCount > 0 ? totalScore / hotelCount : 8.36;

  // --- What-If Scenario calculations ---
  // Scenario 1: All hotels +0.5
  var scenario1Gain = 0;
  opsList.forEach(function(op) {
    scenario1Gain += Math.round(0.5 * revparSlope * op.roomCount * 30);
  });

  // Scenario 2: Bottom 5 hotels to portfolio average
  var sortedByScore = opsList.slice().sort(function(a, b) { return a.currentScore - b.currentScore; });
  var bottom5 = sortedByScore.slice(0, 5);
  var scenario2Gain = 0;
  var scenario2Names = [];
  bottom5.forEach(function(op) {
    var gap = avgScore - op.currentScore;
    if (gap > 0) {
      scenario2Gain += Math.round(gap * revparSlope * op.roomCount * 30);
      scenario2Names.push(op.name);
    }
  });

  // Scenario 3: 30% of total opportunity recovered (cleaning issues resolved)
  var totalOpportunity = 0;
  opsList.forEach(function(op) { totalOpportunity += op.monthlyLoss; });
  var scenario3Gain = Math.round(totalOpportunity * 0.3);

  // --- Room counts array for client-side slider ---
  var roomCountsJS = [];
  opsList.forEach(function(op) {
    roomCountsJS.push(op.roomCount);
  });

  // --- Extra CSS ---
  var extraCSS = [
    '.revenue-overview { display: grid; grid-template-columns: repeat(3, 1fr); gap: 1rem; margin-bottom: 2rem; }',
    '.revenue-card { background: white; border-radius: 12px; padding: 1.5rem; box-shadow: 0 1px 3px rgba(0,0,0,0.08); text-align: center; }',
    '.revenue-card .big-num { font-size: 1.8rem; font-weight: 800; color: #1A1A2E; }',
    '.revenue-card .sub-label { font-size: 0.75rem; color: #64748B; margin-top: 0.25rem; }',
    '.revenue-card .slope-value { font-size: 2rem; font-weight: 800; color: #C23B3A; margin: 0.5rem 0; }',
    '.revenue-card .interp { font-size: 0.78rem; color: #64748B; line-height: 1.4; }',
    '.revenue-card .metric-label { font-size: 0.82rem; font-weight: 700; color: #1A1A2E; margin-bottom: 0.5rem; }',
    '.revenue-card .r-squared { font-size: 0.72rem; color: #94A3B8; margin-top: 0.25rem; }',
    '.data-table { width: 100%; border-collapse: collapse; font-size: 0.82rem; }',
    '.data-table th { background: #1A1A2E; color: white; padding: 0.6rem 0.75rem; text-align: left; font-weight: 600; font-size: 0.75rem; }',
    '.data-table td { padding: 0.55rem 0.75rem; border-bottom: 1px solid #F1F5F9; }',
    '.data-table tr:hover { background: #F8FAFC; }',
    '.data-table .highlight-row { background: #FFF5F5; }',
    '.data-table .rank-badge { display: inline-flex; align-items: center; justify-content: center; width: 24px; height: 24px; border-radius: 50%; font-weight: 800; font-size: 0.75rem; color: white; }',
    '.rank-1 { background: #C23B3A; }',
    '.rank-2 { background: #E85D5C; }',
    '.rank-3 { background: #F09090; }',
    '.scenario-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 1rem; margin-bottom: 2rem; }',
    '.scenario-card { background: white; border-radius: 12px; padding: 1.5rem; box-shadow: 0 1px 3px rgba(0,0,0,0.08); border-top: 4px solid #C23B3A; }',
    '.scenario-card:nth-child(2) { border-top-color: #F59E0B; }',
    '.scenario-card:nth-child(3) { border-top-color: #3B82F6; }',
    '.scenario-card .scenario-title { font-size: 0.88rem; font-weight: 700; color: #1A1A2E; margin-bottom: 0.75rem; }',
    '.scenario-card .scenario-detail { font-size: 0.75rem; color: #64748B; margin-bottom: 0.25rem; }',
    '.scenario-card .scenario-gain { font-size: 1.5rem; font-weight: 800; color: #C23B3A; margin-top: 0.75rem; }',
    '.scenario-card .scenario-gain-label { font-size: 0.72rem; color: #64748B; }',
    '.slider-section { background: white; border-radius: 12px; padding: 2rem; box-shadow: 0 1px 3px rgba(0,0,0,0.08); text-align: center; border: 2px solid #C23B3A; }',
    '.slider-section .slider-title { font-size: 1.1rem; font-weight: 800; color: #1A1A2E; margin-bottom: 1rem; }',
    '.slider-section input[type="range"] { width: 80%; max-width: 500px; height: 8px; -webkit-appearance: none; appearance: none; background: linear-gradient(to right, #E2E8F0, #C23B3A); border-radius: 4px; outline: none; margin: 1rem 0; }',
    '.slider-section input[type="range"]::-webkit-slider-thumb { -webkit-appearance: none; appearance: none; width: 24px; height: 24px; border-radius: 50%; background: #C23B3A; cursor: pointer; box-shadow: 0 2px 6px rgba(194,59,58,0.4); }',
    '.slider-section .slider-value { font-size: 1.2rem; font-weight: 700; color: #1A1A2E; margin-bottom: 0.5rem; }',
    '.slider-section .slider-result { font-size: 2.5rem; font-weight: 800; color: #C23B3A; margin-top: 1rem; }',
    '.slider-section .slider-result-label { font-size: 0.85rem; color: #64748B; margin-top: 0.25rem; }',
    '.text-right { text-align: right; }',
    '.text-center { text-align: center; }',
    '.loss-amount { color: #C23B3A; font-weight: 700; }',
    '@media (max-width: 768px) { .revenue-overview, .scenario-grid { grid-template-columns: 1fr; } .slider-section input[type="range"] { width: 95%; } }'
  ].join('\n');

  var lines = [];
  lines.push(pageHead('REVENUE SIMULATOR - PRIME CHANGE', { extraCSS: extraCSS }));
  lines.push(nav('simulator'));
  lines.push('<div class="container">');

  // === Section Heading ===
  lines.push('<div class="section-heading"><span class="heading-en">REVENUE SIMULATOR</span><span class="heading-ja">売上シミュレーター &mdash; 口コミスコア改善による売上インパクト推計</span></div>');

  // === 1. Regression Summary (3 cards) ===
  var regRevpar = regression.score_vs_revpar || {};
  var regAdr = regression.score_vs_adr || {};
  var regOcc = regression.score_vs_occupancy || {};

  lines.push('<div class="revenue-overview">');

  // RevPAR card
  lines.push('<div class="revenue-card">');
  lines.push('  <div class="metric-label">Score &rarr; RevPAR</div>');
  lines.push('  <div class="slope-value">+' + (regRevpar.slope || 228.9).toFixed(1) + '円</div>');
  lines.push('  <div class="interp">' + esc(regRevpar.interpretation || 'スコア1点上昇でRevPARが229円変動') + '</div>');
  lines.push('  <div class="r-squared">R&sup2; = ' + ((regRevpar.r_squared || 0) * 100).toFixed(1) + '%</div>');
  lines.push('</div>');

  // ADR card
  lines.push('<div class="revenue-card">');
  lines.push('  <div class="metric-label">Score &rarr; ADR</div>');
  lines.push('  <div class="slope-value">+' + (regAdr.slope || 254.5).toFixed(1) + '円</div>');
  lines.push('  <div class="interp">' + esc(regAdr.interpretation || 'スコア1点上昇でADRが255円変動') + '</div>');
  lines.push('  <div class="r-squared">R&sup2; = ' + ((regAdr.r_squared || 0) * 100).toFixed(1) + '%</div>');
  lines.push('</div>');

  // Occupancy card
  lines.push('<div class="revenue-card">');
  lines.push('  <div class="metric-label">Score &rarr; Occupancy</div>');
  lines.push('  <div class="slope-value">+' + ((regOcc.slope || 0.0375) * 100).toFixed(1) + '%</div>');
  lines.push('  <div class="interp">' + esc(regOcc.interpretation || 'スコア1点上昇で稼働率が3.8%ポイント変動') + '</div>');
  lines.push('  <div class="r-squared">R&sup2; = ' + ((regOcc.r_squared || 0) * 100).toFixed(1) + '%</div>');
  lines.push('</div>');

  lines.push('</div>');

  // === 2. Threshold Analysis Table ===
  lines.push('<div class="card"><div class="card-title">スコア帯別パフォーマンス（閾値分析）</div>');
  lines.push('<table class="data-table">');
  lines.push('<thead><tr><th>スコア帯</th><th class="text-center">ホテル数</th><th class="text-right">平均稼働率</th><th class="text-right">平均ADR</th><th class="text-right">平均RevPAR</th></tr></thead>');
  lines.push('<tbody>');
  thresholdGroups.forEach(function(g) {
    lines.push('<tr>');
    lines.push('  <td><strong>' + esc(g.range) + '</strong></td>');
    lines.push('  <td class="text-center">' + g.count + '</td>');
    lines.push('  <td class="text-right">' + esc(g.avg_occupancy_pct || ((g.avg_occupancy * 100).toFixed(1) + '%')) + '</td>');
    lines.push('  <td class="text-right">&yen;' + (g.avg_adr || 0).toLocaleString() + '</td>');
    lines.push('  <td class="text-right">&yen;' + (g.avg_revpar || 0).toLocaleString() + '</td>');
    lines.push('</tr>');
  });
  lines.push('</tbody></table>');
  lines.push('</div>');

  // === 3. Hotel Opportunity Table ===
  lines.push('<div class="card"><div class="card-title">ホテル別 売上改善機会ランキング</div>');
  lines.push('<table class="data-table">');
  lines.push('<thead><tr><th class="text-center">#</th><th>ホテル名</th><th class="text-right">現在スコア</th><th class="text-right">目標スコア</th><th class="text-right">Gap</th><th class="text-right">客室数</th><th class="text-right">月間損失推定</th></tr></thead>');
  lines.push('<tbody>');
  opsList.forEach(function(op, i) {
    var rank = i + 1;
    var rowClass = rank <= 3 ? ' class="highlight-row"' : '';
    var rankBadge = rank <= 3
      ? '<span class="rank-badge rank-' + rank + '">' + rank + '</span>'
      : String(rank);
    lines.push('<tr' + rowClass + '>');
    lines.push('  <td class="text-center">' + rankBadge + '</td>');
    lines.push('  <td>' + esc(op.name) + '</td>');
    lines.push('  <td class="text-right">' + op.currentScore.toFixed(2) + '</td>');
    lines.push('  <td class="text-right">8.89</td>');
    lines.push('  <td class="text-right">' + (op.gap > 0 ? '+' : '') + op.gap.toFixed(2) + '</td>');
    lines.push('  <td class="text-right">' + op.roomCount + '</td>');
    lines.push('  <td class="text-right loss-amount">&yen;' + formatYen(op.monthlyLoss) + '/月</td>');
    lines.push('</tr>');
  });
  lines.push('</tbody></table>');
  lines.push('</div>');

  // === 4. What-If Scenarios (3 cards) ===
  lines.push('<div class="card"><div class="card-title">What-If シナリオ分析</div>');
  lines.push('<div class="scenario-grid">');

  // Scenario 1
  lines.push('<div class="scenario-card">');
  lines.push('  <div class="scenario-title">全ホテル +0.5点改善</div>');
  lines.push('  <div class="scenario-detail">対象: 全 ' + opsList.length + ' ホテル</div>');
  lines.push('  <div class="scenario-detail">各ホテルのスコアを一律0.5点改善した場合</div>');
  lines.push('  <div class="scenario-gain">&yen;' + formatYen(scenario1Gain) + '</div>');
  lines.push('  <div class="scenario-gain-label">月間売上増加（推定）</div>');
  lines.push('</div>');

  // Scenario 2
  lines.push('<div class="scenario-card">');
  lines.push('  <div class="scenario-title">下位5ホテルを平均まで引き上げ</div>');
  lines.push('  <div class="scenario-detail">対象: ' + scenario2Names.join('、') + '</div>');
  lines.push('  <div class="scenario-detail">ポートフォリオ平均 ' + avgScore.toFixed(2) + ' まで改善</div>');
  lines.push('  <div class="scenario-gain">&yen;' + formatYen(scenario2Gain) + '</div>');
  lines.push('  <div class="scenario-gain-label">月間売上増加（推定）</div>');
  lines.push('</div>');

  // Scenario 3
  lines.push('<div class="scenario-card">');
  lines.push('  <div class="scenario-title">清掃問題ゼロ達成</div>');
  lines.push('  <div class="scenario-detail">全体の改善余地の30%を回収</div>');
  lines.push('  <div class="scenario-detail">清掃関連の口コミ課題を解消した場合</div>');
  lines.push('  <div class="scenario-gain">&yen;' + formatYen(scenario3Gain) + '</div>');
  lines.push('  <div class="scenario-gain-label">月間売上増加（推定）</div>');
  lines.push('</div>');

  lines.push('</div></div>');

  // === 5. Interactive Slider Section ===
  lines.push('<div class="card"><div class="card-title">インタラクティブ・シミュレーター</div>');
  lines.push('<div class="slider-section">');
  lines.push('  <div class="slider-title">スコア改善幅を選択してポートフォリオ全体の売上インパクトを確認</div>');
  lines.push('  <div class="slider-value">スコア改善: <strong id="sim-slider-val">+0.0</strong> 点</div>');
  lines.push('  <input type="range" id="sim-slider" min="0" max="20" step="1" value="0">');
  lines.push('  <div class="slider-result" id="sim-result">&yen;0</div>');
  lines.push('  <div class="slider-result-label">月間ポートフォリオ売上増加（推定）</div>');
  lines.push('</div></div>');

  // Embed client-side JS
  lines.push('<script>');
  lines.push('(function() {');
  lines.push('  var REVPAR_SLOPE = ' + revparSlope + ';');
  lines.push('  var ROOM_COUNTS = ' + JSON.stringify(roomCountsJS) + ';');
  lines.push('  var slider = document.getElementById("sim-slider");');
  lines.push('  var valEl = document.getElementById("sim-slider-val");');
  lines.push('  var resultEl = document.getElementById("sim-result");');
  lines.push('  function formatYenJS(n) {');
  lines.push('    if (Math.abs(n) >= 10000) return Math.round(n / 10000) + "万";');
  lines.push('    return n.toLocaleString();');
  lines.push('  }');
  lines.push('  function update() {');
  lines.push('    var imp = parseInt(slider.value, 10) / 10;');
  lines.push('    valEl.textContent = "+" + imp.toFixed(1);');
  lines.push('    var total = 0;');
  lines.push('    for (var i = 0; i < ROOM_COUNTS.length; i++) {');
  lines.push('      total += Math.round(imp * REVPAR_SLOPE * ROOM_COUNTS[i] * 30);');
  lines.push('    }');
  lines.push('    resultEl.innerHTML = "\\u00A5" + formatYenJS(total);');
  lines.push('  }');
  lines.push('  slider.addEventListener("input", update);');
  lines.push('  update();');
  lines.push('})();');
  lines.push('</script>');

  lines.push('</div>');
  lines.push(footer());
  lines.push(pageFoot());
  return lines.join('\n');
}

module.exports = { buildSimulator };
