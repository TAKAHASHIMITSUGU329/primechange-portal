// V2 ES Dashboard page builder
// Generates es-dashboard.html - Employee Satisfaction / Staff workload analysis

var { esc, nav, footer, pageHead, pageFoot } = require('./common-v2');

// ====== Inline helpers ======

function num(v, d) {
  if (v == null) return '-';
  if (typeof v === 'number') {
    if (d != null) return v.toFixed(d);
    return Number(v.toFixed(2)).toLocaleString();
  }
  return esc(String(v));
}

function pct(v) {
  return v == null ? '-' : num(v, 1) + '%';
}

function corrBadge(r) {
  if (r == null) return '<span class="badge badge-gray">N/A</span>';
  var abs = Math.abs(r);
  var cls = abs >= 0.5 ? 'badge-green' : abs >= 0.3 ? 'badge-orange' : 'badge-red';
  var label = abs >= 0.5 ? '強い' : abs >= 0.3 ? '中程度' : '弱い';
  return '<span class="badge ' + cls + '">' + label + ' r=' + r.toFixed(3) + '</span>';
}

// ====== Main builder ======

function buildESDashboard(data) {
  var a2 = (data.analyses && data.analyses[2]) || {};
  var a3 = (data.analyses && data.analyses[3]) || {};

  var meta = a2.analysis_metadata || {};
  var maidProd = a2.maid_productivity || {};
  var maidClaims = a2.maid_claims_summary || {};
  var checkerClaims = a2.checker_claims_summary || {};
  var attendance = a2.attendance_analysis || {};
  var hotelSummaries = a2.hotel_summaries || [];

  var staffingAnalysis = a3.staffing_analysis || {};
  var optimalStaffing = staffingAnalysis.optimal_staffing || {};
  var correlations = staffingAnalysis.correlations || {};
  var hotelSummaryA3 = a3.hotel_summary || [];

  // ── KPI Grid ──
  var kpiHtml = [
    '<div class="kpi-grid" id="esKpiGrid">',
    '  <div class="kpi-card"><div class="kpi-label">総スタッフ数</div><div class="kpi-value">' + num(meta.total_staff_analyzed) + '</div><div class="kpi-sub">名</div></div>',
    '  <div class="kpi-card"><div class="kpi-label">平均出勤日数</div><div class="kpi-value">' + num(attendance.avg_days) + '</div><div class="kpi-sub">日</div></div>',
    '  <div class="kpi-card" style="border-left-color:var(--green);"><div class="kpi-label">平均清掃室数/日</div><div class="kpi-value">' + num(maidProd.avg_rooms_per_day) + '</div><div class="kpi-sub">室</div></div>',
    '  <div class="kpi-card" style="border-left-color:var(--orange);"><div class="kpi-label">クレーム有りメイド数</div><div class="kpi-value">' + num(maidClaims.total_maids_with_claims) + '</div><div class="kpi-sub">名</div></div>',
    '  <div class="kpi-card"><div class="kpi-label">最大清掃室数/日</div><div class="kpi-value">' + num(maidProd.max_rooms_per_day) + '</div><div class="kpi-sub">室</div></div>',
    '</div>',
  ].join('\n');

  // ── スタッフ負荷分析 (Horizontal bars) ──
  var loadBarsHtml = '';
  if (hotelSummaries.length > 0) {
    // Sort by avg_rooms_per_day descending
    var sorted = hotelSummaries.slice().sort(function(a, b) {
      return (b.avg_rooms_per_day || 0) - (a.avg_rooms_per_day || 0);
    });
    var maxRooms = sorted[0].avg_rooms_per_day || 1;

    var bars = sorted.map(function(h) {
      var val = h.avg_rooms_per_day || 0;
      var w = maxRooms > 0 ? Math.min(val / maxRooms * 100, 100).toFixed(0) : 0;
      var color = val > 15 ? 'var(--red)' : val >= 10 ? 'var(--orange)' : 'var(--green)';
      return '<div class="load-bar">'
        + '<div class="load-bar-label">' + esc(h.name) + '</div>'
        + '<div class="load-bar-track">'
        + '<div class="load-bar-fill" style="width:' + w + '%;background:' + color + ';"></div>'
        + '</div>'
        + '<div class="load-bar-val">' + num(val, 1) + '室/日</div>'
        + '</div>';
    }).join('\n');

    loadBarsHtml = [
      '<div class="card">',
      '  <div class="card-title">&#128200; スタッフ負荷分析</div>',
      '  <p style="font-size:0.78rem;color:var(--text-light);margin-bottom:1rem;">ホテル別 平均清掃室数/人/日（&#128994; &lt;10 &#128992; 10-15 &#128308; &gt;15）</p>',
      bars,
      '</div>',
    ].join('\n');
  }

  // ── 出勤率分析 ──
  var attendanceHtml = '';
  if (attendance.avg_days != null) {
    var highAttendance = attendance.high_attendance_count || 0;
    var totalStaff = meta.total_staff_analyzed || 0;
    var highAttendancePct = totalStaff > 0 ? (highAttendance / totalStaff * 100) : 0;

    attendanceHtml = [
      '<div class="card">',
      '  <div class="card-title">&#128197; 出勤率分析</div>',
      '  <div class="grid-3">',
      '    <div style="padding:1.25rem;background:var(--bg);border-radius:8px;border-left:4px solid var(--accent);">',
      '      <div style="font-size:0.75rem;color:var(--text-light);text-transform:uppercase;font-weight:600;">平均出勤日数</div>',
      '      <div style="font-size:1.8rem;font-weight:800;color:var(--navy);">' + num(attendance.avg_days, 1) + '<span style="font-size:0.9rem;font-weight:400;color:var(--text-light);"> 日</span></div>',
      '    </div>',
      '    <div style="padding:1.25rem;background:var(--bg);border-radius:8px;border-left:4px solid var(--green);">',
      '      <div style="font-size:0.75rem;color:var(--text-light);text-transform:uppercase;font-weight:600;">高出勤スタッフ（20日以上）</div>',
      '      <div style="font-size:1.8rem;font-weight:800;color:var(--navy);">' + highAttendance + '<span style="font-size:0.9rem;font-weight:400;color:var(--text-light);"> 名 (' + num(highAttendancePct, 1) + '%)</span></div>',
      '    </div>',
      '    <div style="padding:1.25rem;background:var(--bg);border-radius:8px;border-left:4px solid var(--blue);">',
      '      <div style="font-size:0.75rem;color:var(--text-light);text-transform:uppercase;font-weight:600;">総スタッフ数</div>',
      '      <div style="font-size:1.8rem;font-weight:800;color:var(--navy);">' + totalStaff + '<span style="font-size:0.9rem;font-weight:400;color:var(--text-light);"> 名</span></div>',
      '    </div>',
      '  </div>',
      '</div>',
    ].join('\n');
  }

  // ── クレーム集中リスク (Maid claims top 15) ──
  var claimRiskHtml = '';
  var maidTableHtml = '';
  var checkerTableHtml = '';

  if (maidClaims.top_claim_maids && maidClaims.top_claim_maids.length > 0) {
    var maids = maidClaims.top_claim_maids.slice(0, 15);
    maidTableHtml = '<h3 style="font-size:0.9rem;font-weight:700;color:var(--red);margin-bottom:0.75rem;">&#9888;&#65039; メイド別クレームTop15</h3>';
    maidTableHtml += '<div style="overflow-x:auto;"><table class="data-table"><thead><tr><th>#</th><th>スタッフ名</th><th>クレーム件数</th><th>所属ホテル</th></tr></thead><tbody>';
    maids.forEach(function(m, i) {
      var bg = i < 3 ? ' style="background:#FEF2F2;"' : '';
      maidTableHtml += '<tr' + bg + '><td>' + (i + 1) + '</td><td style="font-weight:600;">' + esc(m.name) + '</td><td style="text-align:center;"><span class="badge badge-red">' + m.claims + '</span></td><td style="font-size:0.8rem;">' + esc(m.hotel) + '</td></tr>';
    });
    maidTableHtml += '</tbody></table></div>';
  }

  if (checkerClaims.top_claim_checkers && checkerClaims.top_claim_checkers.length > 0) {
    var checkers = checkerClaims.top_claim_checkers.slice(0, 15);
    checkerTableHtml = '<h3 style="font-size:0.9rem;font-weight:700;color:var(--orange);margin:1.5rem 0 0.75rem;">&#128269; チェッカー別クレームTop15</h3>';
    checkerTableHtml += '<div style="overflow-x:auto;"><table class="data-table"><thead><tr><th>#</th><th>チェッカー名</th><th>クレーム件数</th><th>所属ホテル</th></tr></thead><tbody>';
    checkers.forEach(function(c, i) {
      var bg = i < 3 ? ' style="background:#FEF2F2;"' : '';
      checkerTableHtml += '<tr' + bg + '><td>' + (i + 1) + '</td><td style="font-weight:600;">' + esc(c.name) + '</td><td style="text-align:center;"><span class="badge badge-orange">' + c.claims + '</span></td><td style="font-size:0.8rem;">' + esc(c.hotel) + '</td></tr>';
    });
    checkerTableHtml += '</tbody></table></div>';
  }

  if (maidTableHtml || checkerTableHtml) {
    claimRiskHtml = [
      '<div class="card">',
      '  <div class="card-title">&#128680; クレーム集中リスク</div>',
      maidTableHtml,
      checkerTableHtml,
      '</div>',
    ].join('\n');
  }

  // ── 人員充足度 (Optimal staffing top5/bottom5 + hotel_summary table) ──
  var staffingHtml = '';
  var optHtml = '';

  // Top5 vs Bottom5
  if (optimalStaffing.top5_hotels || optimalStaffing.bottom5_hotels) {
    optHtml += '<div class="grid-2" style="margin-bottom:1.5rem;">';
    [['top5_hotels', 'Top5（高品質・適正配置）', 'var(--green)', '#ECFDF5'], ['bottom5_hotels', 'Bottom5（改善余地あり）', 'var(--red)', '#FEF2F2']].forEach(function(pair) {
      var list = optimalStaffing[pair[0]];
      if (!list || !list.length) return;
      optHtml += '<div><h4 style="font-size:0.85rem;color:' + pair[2] + ';margin-bottom:0.5rem;">' + pair[1] + '</h4>';
      optHtml += '<table class="data-table"><thead><tr><th>ホテル</th><th>メイド</th><th>チェッカー</th><th>比率</th><th>スコア</th></tr></thead><tbody>';
      list.forEach(function(h) {
        optHtml += '<tr style="background:' + pair[3] + ';"><td>' + esc(h.name) + '</td><td style="text-align:right;">' + num(h.avg_maids) + '</td><td style="text-align:right;">' + num(h.avg_checkers) + '</td><td style="text-align:right;">' + num(h.ratio) + '</td><td style="text-align:right;font-weight:700;">' + num(h.score) + '</td></tr>';
      });
      optHtml += '</tbody></table></div>';
    });
    optHtml += '</div>';
  }

  // Hotel summary table from analysis 3
  var hotelSummaryTableHtml = '';
  if (hotelSummaryA3.length > 0) {
    // Determine top5 and bottom5 names for color coding
    var top5Names = {};
    var bottom5Names = {};
    if (optimalStaffing.top5_hotels) {
      optimalStaffing.top5_hotels.forEach(function(h) { top5Names[h.name] = true; });
    }
    if (optimalStaffing.bottom5_hotels) {
      optimalStaffing.bottom5_hotels.forEach(function(h) { bottom5Names[h.name] = true; });
    }

    hotelSummaryTableHtml += '<h4 style="font-size:0.85rem;font-weight:700;margin-bottom:0.5rem;">全ホテル配置データ</h4>';
    hotelSummaryTableHtml += '<div style="overflow-x:auto;"><table class="data-table"><thead><tr><th>ホテル</th><th>日数</th><th>平均メイド</th><th>平均チェッカー</th><th>M/C比</th><th>完了時間</th><th>スコア</th><th>クレーム率</th></tr></thead><tbody>';
    hotelSummaryA3.forEach(function(h) {
      var bg = '';
      if (top5Names[h.name]) bg = ' style="background:#ECFDF5;"';
      else if (bottom5Names[h.name]) bg = ' style="background:#FEF2F2;"';
      hotelSummaryTableHtml += '<tr' + bg + '><td>' + esc(h.name) + '</td><td style="text-align:right;">' + (h.days || '-') + '</td><td style="text-align:right;">' + num(h.avg_maids) + '</td><td style="text-align:right;">' + num(h.avg_checkers) + '</td><td style="text-align:right;">' + num(h.maid_checker_ratio) + '</td><td style="text-align:right;">' + num(h.avg_completion_time) + '</td><td style="text-align:right;font-weight:700;">' + num(h.score) + '</td><td style="text-align:right;">' + pct(h.claim_rate) + '</td></tr>';
    });
    hotelSummaryTableHtml += '</tbody></table></div>';
  }

  if (optHtml || hotelSummaryTableHtml) {
    staffingHtml = [
      '<div class="card">',
      '  <div class="card-title">&#128101; 人員充足度</div>',
      optHtml,
      hotelSummaryTableHtml,
      '</div>',
    ].join('\n');
  }

  // ── 配置相関分析 ──
  var corrHtml = '';
  if (Object.keys(correlations).length > 0) {
    var corrLabels = {
      'staff_vs_score': '総スタッフ数 vs スコア',
      'maids_vs_score': 'メイド数 vs スコア',
      'checkers_vs_score': 'チェッカー数 vs スコア',
      'ratio_vs_score': 'メイド/チェッカー比 vs スコア',
      'staff_vs_claims': 'スタッフ数 vs クレーム率',
    };

    var corrCards = '<div class="kpi-grid">';
    Object.keys(correlations).forEach(function(k) {
      var c = correlations[k];
      corrCards += '<div class="kpi-card"><div class="kpi-label">' + esc(corrLabels[k] || k) + '</div><div class="kpi-value" style="font-size:1.2rem;">' + corrBadge(c.r) + '</div><div class="kpi-sub">R&sup2;=' + num(c.r_squared, 4) + ' / n=' + (c.n || '-') + '</div></div>';
    });
    corrCards += '</div>';

    corrHtml = [
      '<div class="card">',
      '  <div class="card-title">&#128200; 配置相関分析</div>',
      '  <p style="font-size:0.78rem;color:var(--text-light);margin-bottom:1rem;">人員配置と品質指標の相関関係</p>',
      corrCards,
      '</div>',
    ].join('\n');
  }

  // ── ES改善提言 (Accordion) ──
  var recsHtml = '';
  var allRecs = [];

  // Gather recommendations from analysis 2
  if (a2.recommendations && a2.recommendations.length) {
    a2.recommendations.forEach(function(r) { allRecs.push(r); });
  }
  // Gather recommendations from analysis 3
  if (a3.recommendations && a3.recommendations.length) {
    a3.recommendations.forEach(function(r) { allRecs.push(r); });
  }

  if (allRecs.length > 0) {
    var accItems = allRecs.map(function(r, i) {
      var priCls = 'badge-blue';
      var pri = (r.priority || '').toUpperCase();
      if (pri === 'HIGH' || pri === '高') priCls = 'badge-red';
      else if (pri === 'MEDIUM' || pri === '中') priCls = 'badge-orange';
      else if (pri === 'LOW' || pri === '低') priCls = 'badge-blue';

      var body = '';
      if (r.rationale) {
        body += '<p style="font-size:0.8rem;color:var(--text-light);margin-bottom:0.5rem;">' + esc(r.rationale) + '</p>';
      }
      if (r.description) {
        body += '<p style="font-size:0.8rem;color:var(--text-light);margin-bottom:0.5rem;">' + esc(r.description) + '</p>';
      }
      if (r.actions && r.actions.length) {
        body += '<ul style="font-size:0.8rem;padding-left:1.2rem;">';
        r.actions.forEach(function(a) { body += '<li style="margin-bottom:0.3rem;">' + esc(a) + '</li>'; });
        body += '</ul>';
      }

      return '<div class="accordion-item' + (i === 0 ? ' open' : '') + '">'
        + '<div class="accordion-header" onclick="this.parentElement.classList.toggle(\'open\')">'
        + esc(r.title || '提言 ' + (i + 1))
        + ' <span class="badge ' + priCls + '">' + esc(r.priority || '') + '</span>'
        + '<span class="accordion-arrow">&#9660;</span>'
        + '</div>'
        + '<div class="accordion-body">' + body + '</div>'
        + '</div>';
    }).join('\n');

    recsHtml = [
      '<div class="card">',
      '  <div class="card-title">&#128161; ES改善提言</div>',
      accItems,
      '</div>',
    ].join('\n');
  }

  // ── Extra CSS ──
  var extraCSS = [
    '.corr-card { text-align: center; }',
  ].join('\n');

  // ── Assemble page ──
  var html = [
    pageHead('ES（従業員満足度）ダッシュボード - PRIMECHANGE V2', { scripts: ['es-dashboard.js'], extraCSS: extraCSS }),
    nav('es-dashboard'),
    '<div class="container">',
    '  <h1 class="page-title">ES（従業員満足度）ダッシュボード</h1>',
    '  <p class="page-subtitle">スタッフ負荷・配置・クレーム集中リスクの分析</p>',
    '',
    '  <div class="fulldata-banner">',
    '    <span>&#9888;&#65039;</span>',
    '    <div>この分析は全期間データがベースです。日付フィルターは口コミ件数にのみ適用されます。</div>',
    '  </div>',
    '',
    kpiHtml,
    '',
    loadBarsHtml,
    '',
    attendanceHtml,
    '',
    claimRiskHtml,
    '',
    staffingHtml,
    '',
    corrHtml,
    '',
    recsHtml,
    '</div>',
    footer(),
    pageFoot()
  ];

  return html.join('\n');
}

module.exports = { buildESDashboard };
