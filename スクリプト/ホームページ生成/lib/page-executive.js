// V3 Executive Summary Page Generator
'use strict';
const { esc, nav, footer, pageHead, pageFoot, deltaBadge, deltaBadgeCompact, deltaSummaryBanner } = require('./common-v2');
const { formatYen } = require('./revenue-calc');

var REVENUE_KEY_MAP = {
  keisei_kinshicho: 'keisei_richmond',
  comfort_yokohama_kannai: 'comfort_yokohama'
};

function resolveRevenueKey(key, revenueData) {
  if (revenueData && revenueData[key]) return key;
  if (REVENUE_KEY_MAP[key] && revenueData && revenueData[REVENUE_KEY_MAP[key]]) return REVENUE_KEY_MAP[key];
  return key;
}

function pct(n, digits) {
  if (n == null || isNaN(n)) return '-';
  return Number(n).toFixed(digits == null ? 1 : digits) + '%';
}

function buildExecutive(data, deltas, revenueOps, csResults) {
  var pov = data.pov || {};
  var meta = data.meta || {};
  var kpiTargets = data.kpi ? data.kpi.portfolio_targets || [] : [];
  var priMatrix = data.priMatrix || {};
  var roi = data.roi || {};
  var actionPlans = data.actionPlans || [];
  var revenueData = data.revenueData || {};
  var hotelsRanked = pov.hotels_ranked || [];
  var cleaningActuals = data.cleaningActuals || {};
  var cleaningSummary = cleaningActuals.portfolio_summary || {};
  var cleaningHotels = cleaningActuals.hotels || {};
  var cleaningMeta = cleaningActuals.metadata || {};

  // --- Calculate revenue totals (Feb / Mar / Apr / May) ---
  var febRevenue = 0, marRevenue = 0, aprRevenue = 0, mayRevenue = 0;
  var febOccupancy = 0, marOccupancy = 0, aprOccupancy = 0, mayOccupancy = 0;
  var totalOpportunity = 0, hotelCount = 0;
  Object.keys(revenueData).forEach(function(k) {
    var rd = revenueData[k];
    febRevenue += rd.actual_revenue || 0;
    marRevenue += rd.march_revenue || 0;
    aprRevenue += rd.april_revenue || 0;
    mayRevenue += rd.may_revenue || 0;
    febOccupancy += rd.occupancy_rate || 0;
    marOccupancy += rd.march_occupancy || 0;
    aprOccupancy += rd.april_occupancy || 0;
    mayOccupancy += rd.may_occupancy || 0;
    hotelCount++;
  });
  var totalRevenue = febRevenue;
  var totalOccupancy = febOccupancy;
  var avgOccupancy = hotelCount > 0 ? (totalOccupancy / hotelCount * 100).toFixed(1) : 0;
  var avgFebOcc = hotelCount > 0 ? (febOccupancy / hotelCount * 100).toFixed(1) : 0;
  var avgMarOcc = hotelCount > 0 ? (marOccupancy / hotelCount * 100).toFixed(1) : 0;
  var avgAprOcc = hotelCount > 0 ? (aprOccupancy / hotelCount * 100).toFixed(1) : 0;
  var avgMayOcc = hotelCount > 0 ? (mayOccupancy / hotelCount * 100).toFixed(1) : 0;
  Object.keys(revenueOps || {}).forEach(function(k) { totalOpportunity += (revenueOps[k].monthlyLoss || 0); });
  var mayRevenueHotels = Object.keys(revenueData).filter(function(k) { return (revenueData[k].may_revenue || 0) > 0; }).length;
  var mayVsApr = aprRevenue > 0 ? ((mayRevenue / aprRevenue - 1) * 100) : null;
  var cleaningClaimsPer1000 = cleaningSummary.total_cleaned_rooms > 0
    ? cleaningSummary.total_claims / cleaningSummary.total_cleaned_rooms * 1000
    : null;
  var cleaningClaimsPer1000Text = cleaningClaimsPer1000 == null ? '-' : cleaningClaimsPer1000.toFixed(2) + '件';

  var managementRisks = hotelsRanked.map(function(h) {
    var revKey = resolveRevenueKey(h.key, revenueData);
    var rd = revenueData[revKey] || {};
    var cl = cleaningHotels[h.key] || cleaningHotels[revKey] || {};
    var clSummary = cl.summary || {};
    var mayRev = rd.may_revenue || 0;
    var revenueBase = mayRev || rd.april_revenue || rd.march_revenue || rd.actual_revenue || 0;
    var scoreGap = Math.max(0, 8.89 - (h.avg || 0));
    var cleaningRate = h.cleaning_issue_rate != null ? h.cleaning_issue_rate : (h.cleaning_rate || 0);
    var lowRate = h.low_rate || 0;
    var claimRate = clSummary.total_cleaned_rooms > 0 ? clSummary.total_claims / clSummary.total_cleaned_rooms * 1000 : null;
    var dataPenalty = mayRev > 0 && clSummary.total_cleaned_rooms > 0 ? 1 : 1.25;
    var riskScore = (revenueBase / 1000000) * (1 + scoreGap) * (1 + cleaningRate / 8) * (1 + lowRate / 10) * dataPenalty;
    return {
      key: h.key,
      revenueKey: revKey,
      name: h.name,
      score: h.avg || 0,
      scoreGap: scoreGap,
      highRate: h.high_rate || 0,
      lowRate: lowRate,
      cleaningRate: cleaningRate,
      mayRevenue: mayRev,
      revenueBase: revenueBase,
      monthlyLoss: revenueOps && revenueOps[h.key] ? revenueOps[h.key].monthlyLoss || 0 : 0,
      cleanedRooms: clSummary.total_cleaned_rooms || 0,
      claims: clSummary.total_claims || 0,
      claimRate: claimRate,
      dataMissing: !(mayRev > 0) || !(clSummary.total_cleaned_rooms > 0),
      riskScore: riskScore,
      priority: h.priority || ''
    };
  }).sort(function(a, b) { return b.riskScore - a.riskScore; });

  var topManagementRisks = managementRisks.slice(0, 5);
  var qualityRevenueRisks = managementRisks.filter(function(h) {
    return h.mayRevenue > 0 && (h.score < 8.1 || h.cleaningRate >= 6 || h.lowRate >= 8);
  }).slice(0, 4);
  var dataGaps = [];
  (cleaningMeta.hotels_without_data || []).forEach(function(k) {
    var h = hotelsRanked.find(function(x) { return x.key === k; });
    dataGaps.push((h ? h.name : k) + 'の5月清掃実績が未取得');
  });
  Object.keys(revenueData).forEach(function(k) {
    if (!(revenueData[k].may_revenue > 0)) {
      var h = hotelsRanked.find(function(x) { return resolveRevenueKey(x.key, revenueData) === k; });
      dataGaps.push((h ? h.name : revenueData[k].hotel_name || k) + 'の5月売上が未取得');
    }
  });
  if (cleaningSummary.time_data_points === 0) {
    dataGaps.push('清掃完了時刻はGAS側の時刻取得が未反映');
  }
  dataGaps = dataGaps.filter(function(v, i, arr) { return arr.indexOf(v) === i; }).slice(0, 5);

  var executiveCalls = [
    '5月売上は' + mayRevenueHotels + 'ホテルで&yen;' + formatYen(mayRevenue) + '、稼働率' + avgMayOcc + '%。4月比は' + (mayVsApr == null ? '-' : (mayVsApr >= 0 ? '+' : '') + mayVsApr.toFixed(1) + '%') + 'で、月中データとして進捗監視が必要です。',
    '経営リスクは「売上規模 × 品質ギャップ × 清掃課題 × 低評価率」で見ると、' + (topManagementRisks[0] ? esc(topManagementRisks[0].name) : '該当なし') + 'が最優先です。',
    '清掃実績は' + (cleaningSummary.total_cleaned_rooms || 0).toLocaleString() + '室・クレーム' + (cleaningSummary.total_claims || 0).toLocaleString() + '件、1,000室あたり' + cleaningClaimsPer1000Text + 'です。口コミと日報を同じ会議体で管理できます。'
  ];

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
    '.exec-decision { background: #FFFFFF; border: 1px solid #E8E0E0; border-left: 5px solid #C23B3A; border-radius: 10px; padding: 1.25rem 1.5rem; margin-bottom: 1.25rem; box-shadow: 0 1px 3px rgba(0,0,0,0.08); }',
    '.exec-decision-title { font-size: 0.9rem; font-weight: 800; color: #1A1A2E; margin-bottom: 0.75rem; }',
    '.exec-decision-list { display: grid; gap: 0.55rem; }',
    '.exec-decision-line { font-size: 0.86rem; color: #334155; display: flex; gap: 0.65rem; align-items: flex-start; }',
    '.exec-decision-num { width: 1.35rem; height: 1.35rem; border-radius: 50%; background: #C23B3A; color: white; display: inline-flex; align-items: center; justify-content: center; font-size: 0.72rem; font-weight: 800; flex-shrink: 0; }',
    '.exec-split { display: grid; grid-template-columns: minmax(0, 1.4fr) minmax(320px, 0.8fr); gap: 1.25rem; align-items: start; }',
    '.mgmt-risk-row { display: grid; grid-template-columns: 42px minmax(180px, 1fr) repeat(5, minmax(82px, auto)); gap: 0.75rem; align-items: center; padding: 0.8rem 0; border-bottom: 1px solid #F1F5F9; font-size: 0.78rem; }',
    '.mgmt-risk-row.header { color: #64748B; font-weight: 700; font-size: 0.68rem; text-transform: uppercase; padding-top: 0; }',
    '.mgmt-risk-row:last-child { border-bottom: none; }',
    '.mgmt-rank { width: 32px; height: 32px; border-radius: 8px; background: #FEF2F2; color: #C23B3A; display: inline-flex; align-items: center; justify-content: center; font-weight: 800; }',
    '.mgmt-hotel { font-weight: 800; color: #1A1A2E; }',
    '.mgmt-sub { font-size: 0.7rem; color: #64748B; margin-top: 0.1rem; }',
    '.metric-strong { font-weight: 800; color: #1A1A2E; }',
    '.metric-danger { font-weight: 800; color: #C23B3A; }',
    '.mini-grid { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 0.75rem; }',
    '.mini-panel { background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 0.9rem; }',
    '.mini-label { font-size: 0.7rem; color: #64748B; font-weight: 700; margin-bottom: 0.25rem; }',
    '.mini-value { font-size: 1.2rem; font-weight: 800; color: #1A1A2E; }',
    '.mini-note { font-size: 0.7rem; color: #64748B; margin-top: 0.15rem; }',
    '.gap-list { display: grid; gap: 0.55rem; }',
    '.gap-item { background: #FFFBEB; border: 1px solid #FDE68A; border-radius: 8px; padding: 0.65rem 0.75rem; font-size: 0.76rem; color: #92400E; }',
    '.quadrant-grid { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 0.85rem; }',
    '.quadrant-card { border: 1px solid #E2E8F0; border-radius: 8px; padding: 0.85rem; background: #FFFFFF; }',
    '.quadrant-title { font-size: 0.76rem; font-weight: 800; color: #1A1A2E; margin-bottom: 0.5rem; }',
    '.quadrant-hotel { display: flex; justify-content: space-between; gap: 0.5rem; border-top: 1px solid #F1F5F9; padding: 0.45rem 0; font-size: 0.73rem; }',
    '.decision-table-wrap { overflow-x: auto; }',
    '.kpi-progress-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 1rem; margin-bottom: 2rem; }',
    '.kpi-progress-card { background: white; border-radius: 12px; padding: 1.25rem; box-shadow: 0 1px 3px rgba(0,0,0,0.08); }',
    '.kpi-progress-label { font-size: 0.75rem; font-weight: 600; color: #64748B; margin-bottom: 0.5rem; }',
    '.kpi-progress-values { display: flex; align-items: center; gap: 0.5rem; margin-bottom: 0.75rem; }',
    '.kpi-current { font-size: 1.5rem; font-weight: 800; color: #1A1A2E; }',
    '.kpi-arrow { color: #94A3B8; }',
    '.kpi-ptarget { font-size: 1rem; font-weight: 600; color: #C23B3A; }',
    '.kpi-progress-footer { font-size: 0.7rem; color: #64748B; margin-top: 0.5rem; }',
    '.revenue-overview { display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 1rem; margin-bottom: 2rem; }',
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
    '@media (max-width: 980px) { .exec-split, .quadrant-grid { grid-template-columns: 1fr; } .mgmt-risk-row { grid-template-columns: 36px minmax(180px, 1fr) repeat(2, minmax(80px, auto)); } .mgmt-risk-row .hide-mobile { display: none; } }',
    '@media (max-width: 768px) { .revenue-overview, .roi-grid, .mini-grid { grid-template-columns: 1fr; } .kpi-progress-grid { grid-template-columns: 1fr; } }',
  ].join('\n');

  var lines = [];
  lines.push(pageHead('EXECUTIVE SUMMARY - PRIME CHANGE', { extraCSS: extraCSS }));
  lines.push(nav('executive'));
  lines.push('<div class="container">');

  // --- Delta Summary Banner ---
  lines.push(deltaSummaryBanner(deltas));

  // --- Header ---
  lines.push('<div class="section-heading"><span class="heading-en">EXECUTIVE SUMMARY</span><span class="heading-ja">エグゼクティブサマリー &mdash; 経営会議用ダッシュボード</span></div>');

  // --- Executive Decision Summary ---
  lines.push('<div class="exec-decision">');
  lines.push('<div class="exec-decision-title">本日の経営判断サマリー</div>');
  lines.push('<div class="exec-decision-list">');
  executiveCalls.forEach(function(call, idx) {
    lines.push('<div class="exec-decision-line"><span class="exec-decision-num">' + (idx + 1) + '</span><span>' + call + '</span></div>');
  });
  lines.push('</div></div>');

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
  lines.push('<div class="revenue-card"><div class="sub-label">5月 売上 / 稼働率<span style="font-size:0.7rem;color:#94A3B8;margin-left:0.3rem;">途中</span></div><div class="big-num">&yen;' + formatYen(mayRevenue) + '</div><div class="sub-label" style="margin-top:0.3rem;font-size:0.85rem;">稼働率 ' + avgMayOcc + '%</div></div>');
  lines.push('<div class="revenue-card"><div class="sub-label">月間改善余地（推定）</div><div class="big-num" style="color:#C23B3A;">&yen;' + formatYen(totalOpportunity) + '/月</div></div>');
  lines.push('</div>');

  // --- Management Risk and Data Quality ---
  lines.push('<div class="exec-split">');
  lines.push('<div class="card"><div class="card-title">経営リスク優先順位 TOP5</div>');
  lines.push('<div class="decision-table-wrap">');
  lines.push('<div class="mgmt-risk-row header"><div></div><div>ホテル</div><div>5月売上</div><div>品質</div><div>清掃</div><div class="hide-mobile">低評価</div><div class="hide-mobile">推定効果</div></div>');
  topManagementRisks.forEach(function(h, idx) {
    lines.push('<div class="mgmt-risk-row">');
    lines.push('<div><span class="mgmt-rank">' + (idx + 1) + '</span></div>');
    lines.push('<div><div class="mgmt-hotel">' + esc(h.name) + '</div><div class="mgmt-sub">' + (h.dataMissing ? 'データ欠損あり' : '5月データ取得済み') + ' / リスク指数 ' + h.riskScore.toFixed(1) + '</div></div>');
    lines.push('<div class="metric-strong">' + (h.mayRevenue > 0 ? '&yen;' + formatYen(h.mayRevenue) : '-') + '</div>');
    lines.push('<div class="' + (h.score < 8 ? 'metric-danger' : 'metric-strong') + '">' + h.score.toFixed(2) + '</div>');
    lines.push('<div class="' + (h.cleaningRate >= 6 ? 'metric-danger' : 'metric-strong') + '">' + pct(h.cleaningRate, 1) + '</div>');
    lines.push('<div class="hide-mobile ' + (h.lowRate >= 8 ? 'metric-danger' : 'metric-strong') + '">' + pct(h.lowRate, 1) + '</div>');
    lines.push('<div class="hide-mobile">' + (h.monthlyLoss > 0 ? '&yen;' + formatYen(h.monthlyLoss) + '/月' : '-') + '</div>');
    lines.push('</div>');
  });
  lines.push('</div></div>');
  lines.push('<div class="card"><div class="card-title">データ信頼性・運営実績</div>');
  lines.push('<div class="mini-grid">');
  lines.push('<div class="mini-panel"><div class="mini-label">5月清掃室数</div><div class="mini-value">' + (cleaningSummary.total_cleaned_rooms || 0).toLocaleString() + '</div><div class="mini-note">' + (cleaningMeta.target_period || '') + '</div></div>');
  lines.push('<div class="mini-panel"><div class="mini-label">清掃クレーム</div><div class="mini-value">' + (cleaningSummary.total_claims || 0).toLocaleString() + '</div><div class="mini-note">1,000室あたり ' + cleaningClaimsPer1000Text + '</div></div>');
  lines.push('<div class="mini-panel"><div class="mini-label">5月売上取得</div><div class="mini-value">' + mayRevenueHotels + '/' + Object.keys(revenueData).length + '</div><div class="mini-note">未取得は会議アラート対象</div></div>');
  lines.push('<div class="mini-panel"><div class="mini-label">口コミ母数</div><div class="mini-value">' + ((meta.total_reviews || totalRev || 0).toLocaleString()) + '</div><div class="mini-note">平均スコア ' + (pov.avg_score || 0) + '</div></div>');
  lines.push('</div>');
  if (dataGaps.length > 0) {
    lines.push('<div style="margin-top:1rem;" class="gap-list">');
    dataGaps.forEach(function(g) { lines.push('<div class="gap-item">' + esc(g) + '</div>'); });
    lines.push('</div>');
  }
  lines.push('</div>');
  lines.push('</div>');

  // --- Revenue x Quality View ---
  lines.push('<div class="card"><div class="card-title">売上 × 品質 ポートフォリオ判断</div>');
  lines.push('<div class="quadrant-grid">');
  [
    { title: '守るべき主力（高売上・品質リスク）', items: qualityRevenueRisks },
    { title: '早期改善（品質ギャップ大）', items: managementRisks.filter(function(h) { return h.score < 8 || h.lowRate >= 8; }).slice(0, 4) },
    { title: '清掃運営の重点監視', items: managementRisks.filter(function(h) { return h.cleaningRate >= 6 || (h.claimRate != null && h.claimRate >= 1); }).slice(0, 4) },
    { title: '成長余地（改善効果見込み）', items: managementRisks.filter(function(h) { return h.monthlyLoss > 0; }).sort(function(a, b) { return b.monthlyLoss - a.monthlyLoss; }).slice(0, 4) }
  ].forEach(function(group) {
    lines.push('<div class="quadrant-card"><div class="quadrant-title">' + esc(group.title) + '</div>');
    if (group.items.length === 0) {
      lines.push('<div class="mini-note">該当なし</div>');
    }
    group.items.forEach(function(h) {
      lines.push('<div class="quadrant-hotel"><span>' + esc(h.name) + '</span><strong>' + (h.mayRevenue > 0 ? '&yen;' + formatYen(h.mayRevenue) : '品質' + h.score.toFixed(2)) + '</strong></div>');
    });
    lines.push('</div>');
  });
  lines.push('</div></div>');

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
