// V3 Executive Summary Page Generator
'use strict';
const { esc, nav, footer, pageHead, pageFoot, deltaBadge, deltaBadgeCompact, deltaSummaryBanner } = require('./common-v2');
const { formatYen } = require('./revenue-calc');

function buildExecutive(data, deltas, revenueOps, csResults) {
  var pov = data.pov || {};
  var meta = data.meta || {};
  var kpiTargets = data.kpi ? data.kpi.portfolio_targets || [] : [];
  var priMatrix = data.priMatrix || {};
  var roi = data.roi || {};
  var actionPlans = data.actionPlans || [];
  var revenueData = data.revenueData || {};
  var hotelsRanked = pov.hotels_ranked || [];

  // --- Calculate revenue totals (Feb / Mar / Apr) ---
  var febRevenue = 0, marRevenue = 0, aprRevenue = 0;
  var febOccupancy = 0, marOccupancy = 0, aprOccupancy = 0;
  var totalOpportunity = 0, hotelCount = 0;
  Object.keys(revenueData).forEach(function(k) {
    var rd = revenueData[k];
    febRevenue += rd.actual_revenue || 0;
    marRevenue += rd.march_revenue || 0;
    aprRevenue += rd.april_revenue || 0;
    febOccupancy += rd.occupancy_rate || 0;
    marOccupancy += rd.march_occupancy || 0;
    aprOccupancy += rd.april_occupancy || 0;
    hotelCount++;
  });
  var totalRevenue = febRevenue;
  var totalOccupancy = febOccupancy;
  var avgOccupancy = hotelCount > 0 ? (totalOccupancy / hotelCount * 100).toFixed(1) : 0;
  var avgFebOcc = hotelCount > 0 ? (febOccupancy / hotelCount * 100).toFixed(1) : 0;
  var avgMarOcc = hotelCount > 0 ? (marOccupancy / hotelCount * 100).toFixed(1) : 0;
  var avgAprOcc = hotelCount > 0 ? (aprOccupancy / hotelCount * 100).toFixed(1) : 0;
  Object.keys(revenueOps || {}).forEach(function(k) { totalOpportunity += (revenueOps[k].monthlyLoss || 0); });

  // --- Calculate portfolio NPS ---
  var totalProm = 0, totalDet = 0, totalPass = 0, totalRev = 0;
  Object.keys(csResults || {}).forEach(function(k) {
    var cs = csResults[k];
    totalProm += cs.promoters || 0;
    totalDet += cs.detractors || 0;
    totalPass += cs.passives || 0;
    totalRev += cs.totalReviews || 0;
  });
  var nps = totalRev > 0 ? Math.round((totalProm / totalRev - totalDet / totalRev) * 100) : 0;
  var npsColor = nps > 50 ? '#10B981' : nps > 0 ? '#F59E0B' : '#EF4444';

  // --- KPI progress calculation ---
  function calcProgress(kpi) {
    var current = parseFloat(String(kpi.current).replace(/[%以下以上]/g, '')) || 0;
    var target = parseFloat(String(kpi.target).replace(/[%以下以上]/g, '')) || 0;
    var isLowerBetter = String(kpi.target).indexOf('以下') !== -1 || kpi.kpi.indexOf('クレーム') !== -1 || kpi.kpi.indexOf('低評価') !== -1;
    var pct;
    if (isLowerBetter) {
      if (current <= target) pct = 100;
      else pct = Math.max(0, Math.round((1 - (current - target) / current) * 100));
    } else {
      pct = target > 0 ? Math.min(100, Math.round(current / target * 100)) : 0;
    }
    var color = pct >= 80 ? '#10B981' : pct >= 50 ? '#F59E0B' : '#EF4444';
    return { current: current, target: target, pct: pct, color: color, isLowerBetter: isLowerBetter };
  }

  var extraCSS = [
    '.kpi-progress-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 1rem; margin-bottom: 2rem; }',
    '.kpi-progress-card { background: white; border-radius: 12px; padding: 1.25rem; box-shadow: 0 1px 3px rgba(0,0,0,0.08); }',
    '.kpi-progress-label { font-size: 0.75rem; font-weight: 600; color: #64748B; margin-bottom: 0.5rem; }',
    '.kpi-progress-values { display: flex; align-items: center; gap: 0.5rem; margin-bottom: 0.75rem; }',
    '.kpi-current { font-size: 1.5rem; font-weight: 800; color: #1A1A2E; }',
    '.kpi-arrow { color: #94A3B8; }',
    '.kpi-ptarget { font-size: 1rem; font-weight: 600; color: #C23B3A; }',
    '.kpi-progress-footer { font-size: 0.7rem; color: #64748B; margin-top: 0.5rem; }',
    '.revenue-overview { display: grid; grid-template-columns: repeat(3, 1fr); gap: 1rem; margin-bottom: 2rem; }',
    '.revenue-card { background: white; border-radius: 12px; padding: 1.5rem; box-shadow: 0 1px 3px rgba(0,0,0,0.08); text-align: center; }',
    '.revenue-card .big-num { font-size: 1.8rem; font-weight: 800; color: #1A1A2E; }',
    '.revenue-card .sub-label { font-size: 0.75rem; color: #64748B; margin-top: 0.25rem; }',
    '.risk-card { background: #FFF5F5; border-left: 4px solid #C23B3A; border-radius: 8px; padding: 1rem 1.25rem; margin-bottom: 0.75rem; display: flex; justify-content: space-between; align-items: center; }',
    '.risk-info { flex: 1; }',
    '.risk-hotel-name { font-size: 0.95rem; font-weight: 700; color: #1A1A2E; }',
    '.risk-detail { font-size: 0.78rem; color: #64748B; margin-top: 0.2rem; }',
    '.risk-problems { font-size: 0.72rem; color: #C23B3A; margin-top: 0.2rem; }',
    '.roi-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 1rem; margin-bottom: 2rem; }',
    '.roi-card { background: white; border-radius: 12px; padding: 1.25rem; box-shadow: 0 1px 3px rgba(0,0,0,0.08); border-top: 3px solid #C23B3A; }',
    '.roi-card:nth-child(2) { border-top-color: #F59E0B; }',
    '.roi-card:nth-child(3) { border-top-color: #3B82F6; }',
    '.roi-card-title { font-size: 0.82rem; font-weight: 700; margin-bottom: 0.75rem; color: #1A1A2E; }',
    '.roi-item { font-size: 0.75rem; color: #64748B; padding: 0.3rem 0; border-bottom: 1px solid #F1F5F9; }',
    '.roi-item strong { color: #1A1A2E; }',
    '.nps-display { text-align: center; padding: 1.5rem; }',
    '.nps-number { font-size: 3.5rem; font-weight: 800; }',
    '.nps-label { font-size: 0.85rem; color: #64748B; }',
    '.nps-breakdown { display: flex; justify-content: center; gap: 2rem; margin-top: 1rem; font-size: 0.8rem; }',
    '.action-row { display: flex; justify-content: space-between; align-items: center; padding: 0.75rem 1rem; border-bottom: 1px solid #F1F5F9; }',
    '.action-hotel { font-weight: 700; font-size: 0.85rem; color: #1A1A2E; }',
    '.action-detail { font-size: 0.72rem; color: #64748B; margin-top: 0.15rem; }',
    '@media (max-width: 768px) { .revenue-overview, .roi-grid { grid-template-columns: 1fr; } .kpi-progress-grid { grid-template-columns: 1fr 1fr; } }',
  ].join('\n');

  var lines = [];
  lines.push(pageHead('EXECUTIVE SUMMARY - PRIME CHANGE', { extraCSS: extraCSS }));
  lines.push(nav('executive'));
  lines.push('<div class="container">');

  // --- Delta Summary Banner ---
  lines.push(deltaSummaryBanner(deltas));

  // --- Header ---
  lines.push('<div class="section-heading"><span class="heading-en">EXECUTIVE SUMMARY</span><span class="heading-ja">エグゼクティブサマリー &mdash; 経営会議用ダッシュボード</span></div>');

  // --- Alert banners ---
  if (deltas && deltas.hasDeltas && deltas.alerts && deltas.alerts.length > 0) {
    deltas.alerts.forEach(function(a) {
      var cls = a.severity === 'red' ? 'danger' : a.severity === 'green' ? 'improvement' : 'info';
      lines.push('<div class="alert-banner ' + cls + '"><div class="alert-banner-icon">' + (a.icon || '') + '</div><div class="alert-banner-content"><div class="alert-banner-title">' + esc(a.title) + '</div><div class="alert-banner-msg">' + esc(a.message) + '</div></div></div>');
    });
  }

  // --- KPI Progress ---
  lines.push('<div class="card"><div class="card-title">KPI目標進捗（2026年9月期限）</div>');
  lines.push('<div class="kpi-progress-grid">');
  kpiTargets.forEach(function(kpi) {
    var p = calcProgress(kpi);
    lines.push('<div class="kpi-progress-card">');
    lines.push('  <div class="kpi-progress-label">' + esc(kpi.kpi) + '</div>');
    var kpiDeltaKey = kpi.kpi.indexOf('平均スコア') !== -1 ? 'avg_score' : kpi.kpi.indexOf('高評価') !== -1 ? 'high_rate' : kpi.kpi.indexOf('クレーム') !== -1 ? 'cleaning_issue_rate' : kpi.kpi.indexOf('低評価') !== -1 ? 'low_rate' : null;
    var kpiDeltaObj = deltas && deltas.hasDeltas && kpiDeltaKey && deltas.metrics && deltas.metrics[kpiDeltaKey] ? deltas.metrics[kpiDeltaKey] : null;
    var kpiPolarity = p.isLowerBetter ? 'lower' : 'higher';
    lines.push('  <div class="kpi-progress-values"><span class="kpi-current">' + esc(kpi.current) + '</span>' + deltaBadgeCompact(kpiDeltaObj, kpiPolarity) + '<span class="kpi-arrow">&rarr;</span><span class="kpi-ptarget">' + esc(kpi.target) + '</span></div>');
    lines.push('  <div class="progress-bar-wrap"><div class="progress-bar-fill" style="width:' + p.pct + '%;background:' + p.color + ';"></div></div>');
    lines.push('  <div class="kpi-progress-footer">達成率 <strong style="color:' + p.color + ';">' + p.pct + '%</strong> &middot; 期限: ' + esc(kpi.deadline || '') + '</div>');
    lines.push('</div>');
  });
  lines.push('</div></div>');

  // --- Revenue Overview (Monthly Breakdown) ---
  lines.push('<div class="revenue-overview">');
  lines.push('<div class="revenue-card"><div class="sub-label">2月 売上 / 稼働率</div><div class="big-num">&yen;' + formatYen(febRevenue) + '</div><div class="sub-label" style="margin-top:0.3rem;font-size:0.85rem;">稼働率 ' + avgFebOcc + '%</div></div>');
  lines.push('<div class="revenue-card"><div class="sub-label">3月 売上 / 稼働率</div><div class="big-num">&yen;' + formatYen(marRevenue) + '</div><div class="sub-label" style="margin-top:0.3rem;font-size:0.85rem;">稼働率 ' + avgMarOcc + '%</div></div>');
  lines.push('<div class="revenue-card"><div class="sub-label">4月 売上 / 稼働率<span style="font-size:0.7rem;color:#94A3B8;margin-left:0.3rem;">途中</span></div><div class="big-num">&yen;' + formatYen(aprRevenue) + '</div><div class="sub-label" style="margin-top:0.3rem;font-size:0.85rem;">稼働率 ' + avgAprOcc + '%</div></div>');
  lines.push('<div class="revenue-card"><div class="sub-label">月間改善余地（推定）</div><div class="big-num" style="color:#C23B3A;">&yen;' + formatYen(totalOpportunity) + '/月</div></div>');
  lines.push('</div>');

  // --- Risk Alert TOP3 ---
  var urgentHotels = (priMatrix.urgent || []).concat(priMatrix.high || []).slice(0, 5);
  if (urgentHotels.length > 0) {
    lines.push('<div class="card"><div class="card-title">&#9888; リスクアラート TOP' + Math.min(urgentHotels.length, 5) + '</div>');
    urgentHotels.slice(0, 5).forEach(function(h) {
      var revKey = '';
      hotelsRanked.forEach(function(hr) { if (hr.name === h.hotel) revKey = hr.key; });
      var loss = revenueOps && revenueOps[revKey] ? revenueOps[revKey].monthlyLoss : 0;
      lines.push('<div class="risk-card"><div class="risk-info">');
      lines.push('  <div class="risk-hotel-name">' + esc(h.hotel) + '</div>');
      var riskKey = revKey || '';
      var riskDelta = deltas && deltas.hotels && deltas.hotels[riskKey] && deltas.hotels[riskKey].overall_avg_10pt;
      var riskHighDelta = deltas && deltas.hotels && deltas.hotels[riskKey] && deltas.hotels[riskKey].high_rate;
      var riskLowDelta = deltas && deltas.hotels && deltas.hotels[riskKey] && deltas.hotels[riskKey].low_rate;
      lines.push('  <div class="risk-detail">スコア: <strong style="color:#EF4444;">' + (h.avg || 0) + '</strong>' + deltaBadgeCompact(riskDelta || null, 'higher') + ' / 高評価: <strong>' + (h.high_rate || 0) + '%</strong>' + deltaBadgeCompact(riskHighDelta || null, 'higher') + ' / 低評価: <strong>' + (h.low_rate || 0) + '%</strong>' + deltaBadgeCompact(riskLowDelta || null, 'lower') + ' / 清掃課題率: <strong>' + (h.cleaning_rate || 0) + '%</strong></div>');
      lines.push('  <div class="risk-problems">' + esc((h.key_problems || []).join('、')) + '</div>');
      lines.push('</div>');
      if (loss > 0) lines.push('<span class="revenue-badge loss">&yen;' + formatYen(loss) + '/月</span>');
      lines.push('</div>');
    });
    lines.push('</div>');
  }

  // --- ROI Scenarios ---
  var scenarios = (roi.scenarios || []);
  if (scenarios.length > 0) {
    lines.push('<div class="card"><div class="card-title">ROI シナリオ分析</div>');
    lines.push('<div class="roi-grid">');
    scenarios.forEach(function(s) {
      lines.push('<div class="roi-card">');
      lines.push('  <div class="roi-card-title">' + esc(s.scenario) + '</div>');
      lines.push('  <div class="roi-item">対象: <strong>' + (s.target_hotels || '?') + 'ホテル</strong></div>');
      lines.push('  <div class="roi-item">投資額: <strong>' + esc(s.estimated_cost || '') + '</strong></div>');
      lines.push('  <div class="roi-item">改善見込: <strong>' + esc(s.expected_improvement || '') + '</strong></div>');
      lines.push('  <div class="roi-item">売上効果: <strong style="color:#C23B3A;">' + esc(s.revenue_impact || '') + '</strong></div>');
      lines.push('  <div class="roi-item">回収期間: <strong>' + esc(s.roi_period || '') + '</strong></div>');
      lines.push('</div>');
    });
    lines.push('</div></div>');
  }

  // --- Priority Actions ---
  var urgentActions = actionPlans.filter(function(a) {
    return a.priority_level === 'URGENT' || a.priority_level === 'HIGH';
  }).slice(0, 5);
  if (urgentActions.length > 0) {
    lines.push('<div class="card"><div class="card-title">今月の優先アクション</div>');
    urgentActions.forEach(function(ap) {
      var phase1 = ap.phase1_immediate || {};
      var actions = (phase1.actions || []).slice(0, 2);
      var revKey = '';
      hotelsRanked.forEach(function(hr) { if (hr.name === ap.hotel) revKey = hr.key; });
      var loss = revenueOps && revenueOps[revKey] ? revenueOps[revKey].monthlyLoss : 0;
      lines.push('<div class="action-row">');
      lines.push('  <div><div class="action-hotel">' + esc(ap.hotel) + ' <span class="badge badge-' + (ap.priority_level === 'URGENT' ? 'red' : 'orange') + '">' + esc(ap.priority_level) + '</span></div>');
      lines.push('  <div class="action-detail">' + actions.map(function(a) { return esc(a.action); }).join(' / ') + '</div></div>');
      if (loss > 0) lines.push('  <span class="revenue-badge loss">&yen;' + formatYen(loss) + '/月</span>');
      lines.push('</div>');
    });
    lines.push('</div>');
  }

  // --- NPS ---
  lines.push('<div class="card"><div class="card-title">ポートフォリオ NPS (Net Promoter Score)</div>');
  lines.push('<div class="nps-display">');
  lines.push('  <div class="nps-number" style="color:' + npsColor + ';">' + nps + '</div>');
  lines.push('  <div class="nps-label">NPS スコア（推定）</div>');
  lines.push('  <div class="nps-breakdown">');
  lines.push('    <span style="color:#10B981;">&#128077; 推奨者: ' + totalProm + '名 (' + (totalRev > 0 ? Math.round(totalProm / totalRev * 100) : 0) + '%)</span>');
  lines.push('    <span style="color:#64748B;">&#128528; 中立者: ' + totalPass + '名</span>');
  lines.push('    <span style="color:#EF4444;">&#128078; 批判者: ' + totalDet + '名 (' + (totalRev > 0 ? Math.round(totalDet / totalRev * 100) : 0) + '%)</span>');
  lines.push('  </div>');
  lines.push('</div></div>');

  lines.push('</div>');
  lines.push(footer());
  lines.push(pageFoot());
  return lines.join('\n');
}

module.exports = { buildExecutive };
