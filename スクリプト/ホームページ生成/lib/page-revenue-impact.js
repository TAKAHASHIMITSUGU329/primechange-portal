// V2 Revenue Impact page generator - produces revenue-impact.html
// Improvement 2: 売上インパクト横串 with Hotel Revenue Opportunity Table
const { esc, nav, footer, pageHead, pageFoot, deltaBadge } = require('./common-v2');
const { formatYen } = require('./revenue-calc');

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function corrBadge(r) {
  if (r == null) return '<span class="badge badge-gray">N/A</span>';
  var abs = Math.abs(r);
  var cls = abs >= 0.5 ? 'badge-green' : abs >= 0.3 ? 'badge-orange' : 'badge-red';
  var label = abs >= 0.5 ? '強い' : abs >= 0.3 ? '中程度' : '弱い';
  return '<span class="badge ' + cls + '">' + label + ' r=' + r.toFixed(3) + '</span>';
}

function num(v, digits) {
  if (v == null) return '-';
  if (typeof v !== 'number') return esc(String(v));
  if (digits !== undefined) return v.toFixed(digits);
  if (Math.abs(v) < 0.01) return v.toFixed(4);
  if (Math.abs(v) < 1) return v.toFixed(2);
  return Number(v.toFixed(2)).toLocaleString();
}

function pct(v) {
  if (v == null) return '-';
  return num(v, 1) + '%';
}

function priBadge(pri) {
  var cls = 'badge-gray';
  if (pri === 'URGENT') cls = 'badge-red';
  else if (pri === 'HIGH') cls = 'badge-orange';
  else if (pri === 'STANDARD') cls = 'badge-blue';
  else if (pri === 'MAINTENANCE') cls = 'badge-green';
  return '<span class="badge ' + cls + '">' + esc(pri) + '</span>';
}

// ---------------------------------------------------------------------------
// Main builder
// ---------------------------------------------------------------------------

function buildRevenueImpact(data, revenueOps, deltas) {
  var a6 = (data.analyses && data.analyses[6]) || (data.analyses && data.analyses['6']) || {};
  var html = [];
  var content = [];

  // ---- Extra CSS for this page ----
  var extraCSS = [
    '.scenario-section { margin-bottom: 1.5rem; padding-bottom: 1rem; border-bottom: 1px solid var(--border); }',
    '.scenario-section:last-child { border-bottom: none; }',
    '.opp-table-wrap { overflow-x: auto; }',
    '.opp-total-row td { background: #F1F5F9; font-weight: 700; border-top: 2px solid var(--navy); }',
    '.revenue-highlight { color: var(--green); font-weight: 700; }',
    '.benchmark-box { padding: 1.25rem; background: #F0F9FF; border-radius: var(--radius); border-left: 4px solid var(--accent); }',
  ].join('\n');

  // ---- Head ----
  html.push(pageHead('品質×売上 インパクト分析 | PRIMECHANGE V2', {
    extraCSS: extraCSS,
  }));

  // ---- Nav ----
  html.push(nav('revenue-impact'));

  // ---- Container ----
  html.push('<div class="container">');

  // ---- Title ----
  html.push('<h1 class="page-title">品質×売上 インパクト分析</h1>');
  html.push('<p class="page-subtitle">口コミスコアと売上の相関分析・改善インパクトシミュレーション・ホテル別改善機会の横串比較</p>');

  // ---- Full-data banner ----
  html.push('<div class="fulldata-banner"><span>&#9432;</span><div>このページは全期間データに基づく回帰分析・売上インパクト推定を表示しています。日付フィルターは適用されません。</div></div>');

  // ---- Snapshot-switchable content ----
  html.push('<div id="ri-content">');

  // ==================================================================
  // 5. Portfolio Summary KPIs
  // ==================================================================
  if (a6.portfolio_summary) {
    var ps = a6.portfolio_summary;

    // Calculate total improvement potential from revenueOps
    var totalPotential = 0;
    if (revenueOps) {
      var opsKeys = Object.keys(revenueOps);
      for (var oi = 0; oi < opsKeys.length; oi++) {
        var op = revenueOps[opsKeys[oi]];
        if (op && op.monthlyLoss > 0) {
          totalPotential += op.monthlyLoss;
        }
      }
    }

    content.push('<div class="kpi-grid">');
    content.push('<div class="kpi-card" style="border-left-color:var(--accent);"><div class="kpi-label">月間売上合計</div><div class="kpi-value">&#165;' + Math.round(ps.total_monthly_revenue / 10000).toLocaleString() + '万</div></div>');
    content.push('<div class="kpi-card" style="border-left-color:var(--green);"><div class="kpi-label">平均稼働率</div><div class="kpi-value">' + esc(String(ps.avg_occupancy)) + '</div></div>');
    content.push('<div class="kpi-card" style="border-left-color:var(--blue);"><div class="kpi-label">平均ADR</div><div class="kpi-value">&#165;' + Math.round(ps.avg_adr).toLocaleString() + '</div></div>');
    content.push('<div class="kpi-card"><div class="kpi-label">平均RevPAR</div><div class="kpi-value">&#165;' + Math.round(ps.avg_revpar).toLocaleString() + '</div></div>');
    if (totalPotential > 0) {
      content.push('<div class="kpi-card" style="border-left-color:var(--orange);"><div class="kpi-label">総改善ポテンシャル</div><div class="kpi-value" style="color:var(--green);">+&#165;' + formatYen(totalPotential) + '</div><div class="kpi-sub">月間改善機会</div></div>');
    }
    // Portfolio score delta context
    var scoreDelta = deltas && deltas.hasDeltas && deltas.metrics && deltas.metrics.avg_score;
    if (scoreDelta) {
      content.push('<div class="kpi-card" style="border-left-color:var(--navy);"><div class="kpi-label">ポートフォリオ平均スコア</div><div class="kpi-value">' + scoreDelta.current + '</div><div class="kpi-sub">/ 10 点</div>' + deltaBadge(scoreDelta, 'higher') + '</div>');
    }
    content.push('</div>');
  }

  // ==================================================================
  // 6. Regression Results Table
  // ==================================================================
  if (a6.regression_results) {
    var reg = a6.regression_results;
    content.push('<div class="card"><div class="card-title">&#128200; 回帰分析結果</div>');
    content.push('<div style="overflow-x:auto;"><table class="data-table"><thead><tr><th>指標</th><th>傾き (slope)</th><th>相関係数 r</th><th>R&sup2;</th><th>解釈</th></tr></thead><tbody>');
    var regKeys = Object.keys(reg);
    for (var ri = 0; ri < regKeys.length; ri++) {
      var rk = regKeys[ri];
      var r = reg[rk];
      content.push('<tr><td style="font-weight:600;">' + esc(r.y_label || rk) + '</td><td style="text-align:right;">' + num(r.slope) + '</td><td>' + corrBadge(r.r) + '</td><td style="text-align:right;">' + num(r.r_squared, 4) + '</td><td style="font-size:0.75rem;">' + esc(r.interpretation || '') + '</td></tr>');
    }
    content.push('</tbody></table></div></div>');
  }

  // ==================================================================
  // 7. Threshold Analysis
  // ==================================================================
  if (a6.threshold_analysis && a6.threshold_analysis.groups) {
    var groups = a6.threshold_analysis.groups;
    content.push('<div class="card"><div class="card-title">&#128201; スコア帯別パフォーマンス</div>');
    content.push('<div style="overflow-x:auto;"><table class="data-table"><thead><tr><th>スコア帯</th><th>ホテル数</th><th>平均稼働率</th><th>平均ADR</th><th>平均RevPAR</th><th>平均月間売上</th></tr></thead><tbody>');
    for (var gi = 0; gi < groups.length; gi++) {
      var g = groups[gi];
      content.push('<tr><td style="font-weight:600;">' + esc(g.range) + '</td><td style="text-align:right;">' + g.count + '</td><td style="text-align:right;">' + esc(String(g.avg_occupancy_pct)) + '</td><td style="text-align:right;">&#165;' + Math.round(g.avg_adr).toLocaleString() + '</td><td style="text-align:right;">&#165;' + Math.round(g.avg_revpar).toLocaleString() + '</td><td style="text-align:right;">&#165;' + Math.round(g.avg_revenue).toLocaleString() + '</td></tr>');
    }
    content.push('</tbody></table></div>');
    if (a6.threshold_analysis.threshold_effect) {
      content.push('<div style="margin-top:0.75rem;padding:0.75rem;background:#FFFBEB;border-radius:8px;font-size:0.82rem;"><strong>閾値効果:</strong> ' + esc(a6.threshold_analysis.threshold_effect.description || '') + '</div>');
    }
    content.push('</div>');
  }

  // ==================================================================
  // 8. Revenue Impact Scenarios
  // ==================================================================
  if (a6.revenue_impact_scenarios) {
    var scenarios = a6.revenue_impact_scenarios;
    content.push('<div class="card"><div class="card-title">&#128176; 品質改善による売上インパクト</div>');
    for (var si = 0; si < scenarios.length; si++) {
      var sc = scenarios[si];
      content.push('<div class="scenario-section">');
      content.push('<h3 style="font-size:0.9rem;font-weight:700;color:var(--accent);margin-bottom:0.75rem;">スコア +' + sc.score_improvement + '点改善シナリオ</h3>');
      content.push('<div class="kpi-grid" style="margin-bottom:0.75rem;">');
      content.push('<div class="kpi-card" style="border-left-color:var(--green);"><div class="kpi-label">RevPAR変動</div><div class="kpi-value" style="color:var(--green);">+&#165;' + Math.round(sc.revpar_change).toLocaleString() + '</div><div class="kpi-sub">' + esc(String(sc.revpar_pct_change)) + '</div></div>');
      content.push('<div class="kpi-card" style="border-left-color:var(--green);"><div class="kpi-label">月間売上変動</div><div class="kpi-value" style="color:var(--green);">+&#165;' + Math.round(sc.total_monthly_revenue_change).toLocaleString() + '</div></div>');
      content.push('<div class="kpi-card" style="border-left-color:var(--accent);"><div class="kpi-label">年間売上変動</div><div class="kpi-value" style="color:var(--accent);">+&#165;' + Math.round(sc.annual_revenue_change).toLocaleString() + '</div></div>');
      content.push('</div>');

      // Expandable per-hotel details
      if (sc.per_hotel_impacts) {
        content.push('<details style="margin-bottom:0.5rem;"><summary style="cursor:pointer;font-size:0.78rem;font-weight:600;color:var(--accent);">ホテル別詳細 (' + sc.per_hotel_impacts.length + '件)</summary>');
        content.push('<table class="data-table" style="margin-top:0.5rem;"><thead><tr><th>ホテル名</th><th>現スコア</th><th>現月間売上</th><th>売上変動額</th></tr></thead><tbody>');
        for (var phi = 0; phi < sc.per_hotel_impacts.length; phi++) {
          var h = sc.per_hotel_impacts[phi];
          content.push('<tr><td>' + esc(h.name) + '</td><td style="text-align:right;">' + num(h.current_score) + '</td><td style="text-align:right;">&#165;' + h.current_revenue.toLocaleString() + '</td><td style="text-align:right;color:var(--green);font-weight:700;">+&#165;' + h.estimated_revenue_change.toLocaleString() + '</td></tr>');
        }
        content.push('</tbody></table></details>');
      }
      content.push('</div>');
    }
    content.push('</div>');
  }

  // ==================================================================
  // 9. NEW V2: Hotel Revenue Opportunity Table
  // ==================================================================
  if (revenueOps) {
    var opsArr = [];
    var opsKeys2 = Object.keys(revenueOps);
    for (var ok = 0; ok < opsKeys2.length; ok++) {
      var opItem = revenueOps[opsKeys2[ok]];
      if (opItem) opsArr.push(opItem);
    }

    // Sort by monthlyLoss descending (highest opportunity first)
    opsArr.sort(function(a, b) { return (b.monthlyLoss || 0) - (a.monthlyLoss || 0); });

    var grandTotalLoss = 0;
    for (var gt = 0; gt < opsArr.length; gt++) {
      grandTotalLoss += (opsArr[gt].monthlyLoss || 0);
    }

    content.push('<div class="card"><div class="card-title">&#127919; ホテル別改善機会（売上インパクト横串）</div>');
    content.push('<p style="font-size:0.8rem;color:var(--text-light);margin-bottom:1rem;">口コミスコアを目標値まで改善した場合の月間売上改善ポテンシャルをホテル横断で比較します。</p>');
    content.push('<div class="opp-table-wrap"><table class="data-table"><thead><tr><th>ホテル名</th><th>現スコア</th><th>目標スコア</th><th>Gap</th><th>客室数</th><th>月間売上</th><th>月間改善機会</th></tr></thead><tbody>');

    for (var hi = 0; hi < opsArr.length; hi++) {
      var ho = opsArr[hi];
      var gapDisplay = ho.gap > 0 ? '+' + num(ho.gap, 2) : num(ho.gap, 2);
      var revBadge = ho.monthlyLoss > 0
        ? '<span class="revenue-badge">+&#165;' + formatYen(ho.monthlyLoss) + '/月</span>'
        : '<span class="badge badge-green">目標達成</span>';

      content.push('<tr>');
      content.push('<td style="font-weight:600;">' + esc(ho.name) + '</td>');
      content.push('<td style="text-align:right;">' + num(ho.currentScore, 2) + '</td>');
      content.push('<td style="text-align:right;">' + num(ho.targetScore, 2) + '</td>');
      content.push('<td style="text-align:right;">' + gapDisplay + '</td>');
      content.push('<td style="text-align:right;">' + (ho.roomCount || '-') + '</td>');
      content.push('<td style="text-align:right;">&#165;' + (ho.actualRevenue ? ho.actualRevenue.toLocaleString() : '-') + '</td>');
      content.push('<td style="text-align:right;">' + revBadge + '</td>');
      content.push('</tr>');
    }

    // Total row
    content.push('<tr class="opp-total-row">');
    content.push('<td colspan="6" style="text-align:right;">合計 (' + opsArr.length + 'ホテル)</td>');
    content.push('<td style="text-align:right;"><span class="revenue-badge">+&#165;' + formatYen(grandTotalLoss) + '/月</span></td>');
    content.push('</tr>');

    content.push('</tbody></table></div></div>');
  }

  // ==================================================================
  // 10. Benchmark Comparison
  // ==================================================================
  if (a6.benchmark_comparison) {
    var bm = a6.benchmark_comparison;
    content.push('<div class="card"><div class="card-title">&#128209; 業界ベンチマーク比較</div>');
    content.push('<div class="benchmark-box">');
    content.push('<div style="font-size:0.85rem;"><strong>業界ベンチマーク:</strong> ' + esc(bm.industry_benchmark || '') + '</div>');
    content.push('<div style="font-size:0.85rem;margin-top:0.3rem;"><strong>自社データ:</strong> スコア0.1pt改善あたりRevPAR ' + esc(String(bm.our_data_revpar_pct_per_01 || '')) + ' (' + esc(String(bm.our_revpar_change_per_01 || '')) + ')</div>');
    content.push('<div style="font-size:0.82rem;color:var(--text-light);margin-top:0.5rem;">' + esc(bm.interpretation || '') + '</div>');
    content.push('</div></div>');
  }

  // ==================================================================
  // 11. Hotel Improvement Potentials
  // ==================================================================
  if (a6.hotel_improvement_potentials) {
    content.push('<div class="card"><div class="card-title">&#128202; ホテル別改善ポテンシャル</div>');
    content.push('<div style="overflow-x:auto;"><table class="data-table"><thead><tr><th>ホテル</th><th>現スコア</th><th>目標</th><th>改善幅</th><th>優先度</th><th>月間売上変動</th></tr></thead><tbody>');
    for (var ip = 0; ip < a6.hotel_improvement_potentials.length; ip++) {
      var hp = a6.hotel_improvement_potentials[ip];
      content.push('<tr><td>' + esc(hp.name) + '</td><td style="text-align:right;">' + num(hp.current_score) + '</td><td style="text-align:right;">' + num(hp.target_score) + '</td><td style="text-align:right;">+' + num(hp.improvement) + '</td><td>' + priBadge(hp.priority) + '</td><td style="text-align:right;color:var(--green);font-weight:700;">+&#165;' + (hp.estimated_monthly_impact || 0).toLocaleString() + '</td></tr>');
    }
    content.push('</tbody></table></div>');
    if (a6.total_improvement_potential) {
      content.push('<div style="margin-top:0.75rem;padding:0.75rem;background:#ECFDF5;border-radius:8px;font-size:0.85rem;text-align:center;"><strong>総改善ポテンシャル:</strong> 月間 +&#165;' + Math.round(a6.total_improvement_potential.monthly).toLocaleString() + ' / 年間 +&#165;' + Math.round(a6.total_improvement_potential.annual).toLocaleString() + '</div>');
    }
    content.push('</div>');
  }

  // ==================================================================
  // 12. Recommendations Accordion
  // ==================================================================
  if (a6.recommendations && a6.recommendations.length) {
    content.push('<div class="card"><div class="card-title">&#128161; 改善提言</div>');
    for (var rc = 0; rc < a6.recommendations.length; rc++) {
      var rec = a6.recommendations[rc];
      var priCls = (rec.priority === 'HIGH' || rec.priority === '高') ? 'badge-red' : (rec.priority === 'MEDIUM' || rec.priority === '中') ? 'badge-orange' : 'badge-blue';
      var isOpen = rc === 0 ? ' open' : '';
      content.push('<div class="accordion-item' + isOpen + '">');
      content.push('<div class="accordion-header" onclick="this.parentElement.classList.toggle(\'open\')">' + esc(rec.title) + ' <span class="badge ' + priCls + '">' + esc(rec.priority || '') + '</span><span class="accordion-arrow">&#9660;</span></div>');
      content.push('<div class="accordion-body">');
      if (rec.rationale) {
        content.push('<p style="font-size:0.8rem;color:var(--text-light);margin-bottom:0.5rem;">' + esc(rec.rationale) + '</p>');
      }
      if (rec.actions && rec.actions.length) {
        content.push('<ul style="font-size:0.8rem;padding-left:1.2rem;">');
        for (var ai = 0; ai < rec.actions.length; ai++) {
          content.push('<li style="margin-bottom:0.3rem;">' + esc(rec.actions[ai]) + '</li>');
        }
        content.push('</ul>');
      }
      content.push('</div></div>');
    }
    content.push('</div>');
  }

  // ---- Assemble content into page ----
  var contentHtml = content.join('\n');
  html.push(contentHtml);
  html.push('</div>'); // #ri-content

  // ---- Footer ----
  html.push('</div>'); // .container
  html.push(footer());
  html.push(pageFoot());

  return {
    html: html.join('\n'),
    snapshotContent: { html: contentHtml },
  };
}

module.exports = { buildRevenueImpact };
