// V2 Index Page Builder - generates index.html for the portfolio portal
const { esc, nav, footer, pageHead, pageFoot } = require('./common-v2');
const { formatYen } = require('./revenue-calc');

// Achievement calculation helpers
function calcAchievement(current, target, lowerIsBetter) {
  if (current == null || target == null || target === 0) return 0;
  if (lowerIsBetter) {
    // For metrics where lower is better (e.g. cleaning rate, low rate)
    if (current <= target) return 100;
    return Math.round(target / current * 1000) / 10;
  }
  // For metrics where higher is better
  return Math.round(current / target * 1000) / 10;
}

function achievementClass(pct) {
  if (pct >= 90) return 'good';
  if (pct >= 70) return 'ok';
  return 'bad';
}

function buildIndex(data, deltas, revenueOps) {
  var meta = data.meta;
  var pov = data.pov;
  var priMatrix = data.priMatrix;
  var kpiTargets = data.kpiTargets || {};
  var cleanDive = data.cleanDive;

  // Resolve KPI target values
  var targetAvgScore = (kpiTargets['ポートフォリオ平均スコア'] && kpiTargets['ポートフォリオ平均スコア'].target) || 8.89;
  var targetHighRate = (kpiTargets['高評価率'] && kpiTargets['高評価率'].target) || 83.4;
  var targetCleanRate = (kpiTargets['清掃クレーム率'] && kpiTargets['清掃クレーム率'].target) || 2.0;
  var targetLowRate = (kpiTargets['低評価率'] && kpiTargets['低評価率'].target) || 2.2;

  // Current values
  var curAvgScore = pov.avg_score || 0;
  var curHighRate = pov.high_rate_pct || pov.portfolio_high_rate || 0;
  var curLowRate = pov.low_rate_pct || pov.portfolio_low_rate || 0;
  var curCleanRate = (cleanDive && cleanDive.overall_cleaning_issue_rate != null)
    ? cleanDive.overall_cleaning_issue_rate
    : (cleanDive && cleanDive.portfolio_cleaning_issue_rate != null ? cleanDive.portfolio_cleaning_issue_rate : 0);
  var totalHotels = meta.total_hotels || (pov.hotels_ranked ? pov.hotels_ranked.length : 0);
  var totalReviews = meta.total_reviews || 0;

  var html = [
    pageHead('PRIMECHANGE ポートフォリオ品質管理ポータル V2'),
    nav('index'),
    '<div class="container">'
  ];

  // ── Alert Banners ──
  if (deltas && deltas.hasDeltas && deltas.alerts && deltas.alerts.length > 0) {
    deltas.alerts.forEach(function(a) {
      html.push(
        '<div class="alert-banner ' + esc(a.type) + '">',
        '  <div class="alert-banner-icon">' + a.icon + '</div>',
        '  <div class="alert-banner-content">',
        '    <div class="alert-banner-title">' + esc(a.title) + '</div>',
        '    <div class="alert-banner-msg">' + esc(a.message) + '</div>',
        '  </div>',
        '</div>'
      );
    });
  }

  // ── Page Title + Subtitle ──
  html.push(
    '<h1 class="page-title">PRIMECHANGE ポートフォリオ品質管理ポータル V2</h1>',
    '<p class="page-subtitle">' + esc(meta.date) + ' 更新 &mdash; ' + totalHotels + 'ホテル・' + totalReviews.toLocaleString() + '件の口コミデータに基づく統合ダッシュボード</p>',
    ''
  );

  // ── KPI Grid ──
  var kpiDefs = [
    { key: 'total_hotels', label: 'ホテル数', value: totalHotels, unit: 'ホテル', color: 'var(--accent)', target: null, lower: false, deltaKey: null },
    { key: 'total_reviews', label: '口コミ数', value: totalReviews.toLocaleString(), unit: '件', color: 'var(--green)', target: null, lower: false, deltaKey: 'total_reviews' },
    { key: 'avg_score', label: '平均スコア', value: curAvgScore, unit: '/ 10 点', color: curAvgScore >= 8 ? 'var(--green)' : 'var(--orange)', target: targetAvgScore, lower: false, deltaKey: 'avg_score' },
    { key: 'high_rate', label: '高評価率', value: curHighRate + '%', unit: '8点以上', color: 'var(--green)', target: targetHighRate, lower: false, deltaKey: 'high_rate', rawVal: curHighRate },
    { key: 'cleaning_issue_rate', label: '清掃クレーム率', value: curCleanRate + '%', unit: (cleanDive && cleanDive.total_cleaning_mentions ? cleanDive.total_cleaning_mentions + '件' : ''), color: 'var(--red)', target: targetCleanRate, lower: true, deltaKey: 'cleaning_issue_rate', rawVal: curCleanRate }
  ];

  html.push('<div class="kpi-grid" id="indexKpiGrid">');

  kpiDefs.forEach(function(k) {
    var deltaHtml = '';
    if (deltas && deltas.hasDeltas && k.deltaKey && deltas.metrics[k.deltaKey]) {
      var d = deltas.metrics[k.deltaKey];
      var deltaVal = d.delta;
      var isImproved = k.lower ? deltaVal < 0 : deltaVal > 0;
      var arrow = deltaVal > 0 ? '&#9650;' : deltaVal < 0 ? '&#9660;' : '&#9654;';
      var dCls = isImproved ? 'up' : 'down';
      deltaHtml = '<div class="kpi-delta ' + dCls + '">' + arrow + ' ' + (deltaVal > 0 ? '+' : '') + deltaVal + '</div>';
    }

    var targetHtml = '';
    if (k.target != null) {
      var rawVal = k.rawVal != null ? k.rawVal : (typeof k.value === 'number' ? k.value : parseFloat(String(k.value)));
      var achPct = calcAchievement(rawVal, k.target, k.lower);
      var achCls = achievementClass(achPct);
      targetHtml = '<div class="kpi-target">目標' + k.target + ' 達成率<span class="achievement ' + achCls + '">' + achPct + '%</span></div>';
    }

    html.push(
      '<div class="kpi-card" style="border-left-color:' + k.color + ';">',
      '  <div class="kpi-label">' + k.label + '</div>',
      '  <div class="kpi-value" data-kpi="' + k.key + '">' + k.value + '</div>',
      '  <div class="kpi-sub">' + k.unit + '</div>',
      deltaHtml,
      targetHtml,
      '</div>'
    );
  });

  html.push('</div>', '');

  // ── 7 Link Cards ──
  var a6 = data.analyses && data.analyses[6];
  var monthlyRev = (a6 && a6.portfolio_summary && a6.portfolio_summary.total_monthly_revenue)
    ? Math.round(a6.portfolio_summary.total_monthly_revenue / 10000).toLocaleString()
    : '---';

  var linkCards = [
    { icon: '&#128200;', title: 'ホテル別口コミ', desc: 'ホテルの口コミ分析をカード形式で一覧。サイト別評価・スコア分布・口コミ詳細をモーダルで閲覧。', link: 'hotel-dashboard.html', stat: totalHotels + 'ホテル / ' + totalReviews.toLocaleString() + '件' },
    { icon: '&#128167;', title: '清掃戦略', desc: '清掃品質の課題分析、カテゴリ別ヒートマップ、優先度マトリクス、横断的改善施策。', link: 'cleaning-strategy.html', stat: (cleanDive && cleanDive.total_cleaning_mentions ? cleanDive.total_cleaning_mentions + '件の清掃指摘' : '') },
    { icon: '&#128270;', title: '深掘り分析', desc: 'クレーム類型・スタッフ・人員配置・完了時間・安全・品質売上・ベストプラクティスの7分析。', link: 'deep-analysis.html', stat: '7つの専門分析' },
    { icon: '&#128176;', title: '品質×売上', desc: '品質スコアと売上の相関分析、弾力性係数、3つのROIシナリオ比較。', link: 'revenue-impact.html', stat: '月間約' + monthlyRev + '万円' },
    { icon: '&#9989;', title: 'アクションプラン', desc: 'ホテル別の3フェーズ改善計画（即時/短期/中期）とKPI目標管理。', link: 'action-plans.html', stat: (data.actionPlans ? data.actionPlans.length + 'ホテルの改善計画' : '') },
    { icon: '&#128101;', title: 'ES管理', desc: '従業員満足度と品質スコアの相関分析、負荷管理、離職リスク予測。', link: 'es-dashboard.html', stat: 'ES×品質統合管理' }
  ];

  html.push('<div class="grid-cards">');
  linkCards.forEach(function(c) {
    html.push(
      '<a href="' + c.link + '" class="link-card">',
      '<div class="card">',
      '  <div class="link-card-icon">' + c.icon + '</div>',
      '  <div class="link-card-title">' + c.title + '</div>',
      '  <div class="link-card-desc">' + c.desc + '</div>',
      '  <div class="link-card-stat">' + c.stat + ' &rarr;</div>',
      '</div></a>'
    );
  });
  html.push('</div>', '');

  // ── 今週の優先アクション TOP3 ──
  var urgentHotels = (priMatrix && priMatrix.urgent) || [];
  if (urgentHotels.length > 0) {
    var top3 = urgentHotels.slice(0, 3);
    html.push(
      '<div class="card">',
      '  <div class="card-title">&#128293; 今週の優先アクション TOP3</div>'
    );

    top3.forEach(function(h, idx) {
      var hotelName = h.hotel || h.name || '';
      var hotelKey = h.key || '';
      var score = h.avg || h.avg_score || 0;

      // Target score from per-hotel or portfolio
      var targetScore = targetAvgScore;
      if (data.perHotelTargets && data.perHotelTargets[hotelName]) {
        targetScore = data.perHotelTargets[hotelName].target_avg || targetAvgScore;
      }

      // Revenue loss
      var monthlyLoss = 0;
      if (revenueOps && revenueOps[hotelKey]) {
        monthlyLoss = revenueOps[hotelKey].monthlyLoss || 0;
      }

      // First immediate action
      var firstAction = '';
      if (data.actionPlans) {
        var plan = data.actionPlans.find(function(p) { return p.hotel === hotelName; });
        if (plan && plan.phases && plan.phases[0] && plan.phases[0].actions && plan.phases[0].actions[0]) {
          firstAction = plan.phases[0].actions[0].action || plan.phases[0].actions[0];
        }
      }

      html.push(
        '  <div style="padding:1rem;border-left:4px solid var(--red);background:#FEF2F2;border-radius:8px;margin-bottom:0.75rem;">',
        '    <div style="display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:0.5rem;">',
        '      <div>',
        '        <span style="font-weight:800;font-size:1.1rem;color:var(--red);">#' + (idx + 1) + '</span>',
        '        <span style="font-weight:700;font-size:0.95rem;margin-left:0.5rem;">' + esc(hotelName) + '</span>',
        '        <span style="font-size:0.8rem;color:var(--text-light);margin-left:0.5rem;">スコア ' + score + ' &rarr; 目標 ' + targetScore + '</span>',
        '      </div>',
        (monthlyLoss > 0 ? '      <span class="revenue-badge loss">-&yen;' + formatYen(monthlyLoss) + '/月</span>' : ''),
        '    </div>',
        (firstAction ? '    <div style="font-size:0.8rem;color:var(--text);margin-top:0.5rem;">&#9989; ' + esc(String(firstAction)) + '</div>' : ''),
        '  </div>'
      );
    });

    html.push('</div>', '');
  }

  // ── 緊急対応ホテル Section ──
  var highHotels = (priMatrix && priMatrix.high) || [];
  var alertHotels = urgentHotels.concat(highHotels);

  if (alertHotels.length > 0) {
    html.push(
      '<div class="card alert-card">',
      '  <div class="card-title" style="color:var(--red);">&#9888; 緊急対応が必要なホテル (' + alertHotels.length + '件)</div>',
      '  <div class="alert-grid">'
    );

    alertHotels.forEach(function(h) {
      var hotelName = h.hotel || h.name || '';
      var hotelKey = h.key || '';
      var score = h.avg || h.avg_score || 0;
      var cleanRate = h.cleaning_rate || h.cleaning_issue_rate || 0;

      // Revenue loss badge
      var lossHtml = '';
      if (revenueOps && revenueOps[hotelKey]) {
        var loss = revenueOps[hotelKey].monthlyLoss || 0;
        if (loss > 0) {
          lossHtml = '<div style="margin-top:0.3rem;"><span class="revenue-badge loss">-&yen;' + formatYen(loss) + '/月</span></div>';
        }
      }

      html.push(
        '    <div class="alert-item">',
        '      <div class="alert-item-title">' + esc(hotelName) + '</div>',
        '      <div class="alert-item-detail">スコア: <strong style="color:var(--red);">' + score + '</strong> / 清掃課題率: <strong>' + cleanRate + '%</strong></div>',
        '      <div class="alert-item-sub">主要課題: ' + esc((h.key_problems || []).join(', ')) + '</div>',
        lossHtml,
        '    </div>'
      );
    });

    html.push('  </div>', '</div>', '');
  }

  // ── Portfolio Trend Chart ──
  html.push(
    '<div class="card">',
    '  <div class="card-title">&#128200; 口コミトレンド（日別）</div>',
    '  <div id="portfolioTrend" style="margin-top:1rem;"></div>',
    '</div>',
    ''
  );

  // ── 4 SVG Gauges ──
  var gauges = [
    { label: '平均スコア', value: curAvgScore, target: targetAvgScore, unit: '', lower: false },
    { label: '清掃クレーム率', value: curCleanRate, target: targetCleanRate, unit: '%', lower: true },
    { label: '高評価率', value: curHighRate, target: targetHighRate, unit: '%', lower: false },
    { label: '低評価率', value: curLowRate, target: targetLowRate, unit: '%', lower: true }
  ];

  html.push(
    '<div class="card">',
    '  <div class="card-title">&#127919; KPI目標 達成状況</div>',
    '  <div class="gauge-row">'
  );

  gauges.forEach(function(g) {
    var achPct = calcAchievement(g.value, g.target, g.lower);
    var achCls = achievementClass(achPct);

    html.push(
      '    <div class="gauge-item">',
      '      <div class="gauge-label">' + g.label + '</div>',
      '      <div class="svg-gauge" data-value="' + g.value + '" data-target="' + g.target + '" data-unit="' + g.unit + '" data-label="' + g.label + '" data-lower="' + g.lower + '"></div>',
      '      <div style="font-size:0.75rem;color:var(--text-light);margin-top:0.3rem;">達成率: <span class="achievement ' + achCls + '">' + achPct + '%</span></div>',
      '    </div>'
    );
  });

  html.push('  </div>', '</div>');

  // ── Close container + footer ──
  html.push('</div>', footer(), pageFoot());

  return html.join('\n');
}

module.exports = { buildIndex };
