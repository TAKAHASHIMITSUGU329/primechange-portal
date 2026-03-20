// V2 Talent Management Dashboard page builder
// Generates talent-management.html - Individual staff evaluation & staffing optimization
// Data source: analysis_2 (staff productivity/claims) + analysis_3 (staffing/correlations)
// NOTE: Hotel-level staffing analysis, correlations, recommendations are in ES Dashboard

var { esc, nav, footer, pageHead, pageFoot } = require('./common-v2');

// ====== Helpers ======

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

// ====== Main builder ======

function buildTalentManagement(data) {
  var a2 = (data.analyses && data.analyses[2]) || {};
  var a3 = (data.analyses && data.analyses[3]) || {};

  var meta = a2.analysis_metadata || {};
  var maidProd = a2.maid_productivity || {};
  var maidClaims = a2.maid_claims_summary || {};
  var checkerClaims = a2.checker_claims_summary || {};
  var attendance = a2.attendance_analysis || {};
  var hotelSummaries = a2.hotel_summaries || [];
  var topPerformers = maidProd.top_performers || [];

  // Aggregate stats
  var totalStaff = meta.total_staff_analyzed || 0;
  var totalMaids = 0, totalCheckers = 0, totalClaims = 0;
  hotelSummaries.forEach(function(h) {
    totalMaids += h.maid_count || 0;
    totalCheckers += h.checker_count || 0;
    totalClaims += h.total_claims || 0;
  });
  var maidsWithClaims = maidClaims.total_maids_with_claims || 0;
  var kaikinCount = topPerformers.filter(function(p) { return (p.absence_days || 0) === 0; }).length;
  var kaikinRate = topPerformers.length > 0 ? Math.round(kaikinCount / topPerformers.length * 1000) / 10 : 0;

  // ── Extra CSS ──
  var extraCSS = [
    '.perf-highlight { background: linear-gradient(90deg, #ECFDF5 0%, transparent 100%); }',
    '.rank-medal { display: inline-block; width: 22px; height: 22px; border-radius: 50%; text-align: center; line-height: 22px; font-size: 0.65rem; font-weight: 800; color: white; margin-right: 0.3rem; }',
    '.rank-1 { background: #F59E0B; }',
    '.rank-2 { background: #9CA3AF; }',
    '.rank-3 { background: #B45309; }',
    '.stacked-bar { display: flex; height: 22px; border-radius: 4px; overflow: hidden; }',
    '.stacked-seg { display: flex; align-items: center; justify-content: center; font-size: 0.6rem; font-weight: 700; color: white; white-space: nowrap; overflow: hidden; }',
    '.hist-bar { display: flex; align-items: flex-end; gap: 2px; height: 80px; }',
    '.hist-col { flex: 1; border-radius: 3px 3px 0 0; display: flex; align-items: flex-end; justify-content: center; font-size: 0.6rem; font-weight: 600; color: white; min-width: 18px; }',
  ].join('\n');

  var html = [
    pageHead('タレントマネジメント - PRIMECHANGE V2', { extraCSS: extraCSS }),
    nav('talent-management'),
    '<div class="container">',
    '  <h1 class="page-title">タレントマネジメント</h1>',
    '  <p class="page-subtitle">19ホテル・' + totalStaff + '名のスタッフデータに基づく個人評価・配置最適化分析</p>',
    '',
    '  <div class="fulldata-banner">',
    '    <span>&#9432;</span>',
    '    <div>R8期間の🏆皆勤アワード・🔵クレーム・②シフト・③日報データを統合。ホテル単位の配置分析・相関・提言は<a href="es-dashboard.html" style="color:#92400E;font-weight:700;">ES管理ページ</a>をご覧ください。</div>',
    '  </div>',
  ];

  // ============================================================
  // Section 1: 人材KPI概要
  // ============================================================
  html.push(
    '<div class="section-heading"><span class="heading-en">TALENT KPI</span><span class="heading-ja">人材KPI概要</span></div>',
    '<div class="kpi-grid">'
  );

  var kpis = [
    { label: '総スタッフ数', value: totalStaff, unit: '名', color: 'var(--accent)' },
    { label: 'メイド数', value: totalMaids, unit: '名', color: 'var(--blue)' },
    { label: 'チェッカー数', value: totalCheckers, unit: '名', color: 'var(--purple)' },
    { label: '平均出勤日数', value: num(attendance.avg_days, 1), unit: '日', color: 'var(--green)' },
    { label: '平均清掃室数/日', value: num(maidProd.avg_rooms_per_day, 1), unit: '室', color: 'var(--blue)' },
    { label: '皆勤該当率', value: kaikinRate + '%', unit: kaikinCount + '/' + topPerformers.length + '名', color: 'var(--green)' },
  ];

  kpis.forEach(function(k) {
    html.push(
      '<div class="kpi-card" style="border-left-color:' + k.color + ';">',
      '  <div class="kpi-label">' + k.label + '</div>',
      '  <div class="kpi-value">' + k.value + '</div>',
      '  <div class="kpi-sub">' + k.unit + '</div>',
      '</div>'
    );
  });

  html.push('</div>', '');

  // ============================================================
  // Section 2: スタッフ生産性ランキング Top20
  // ============================================================
  html.push(
    '<div class="section-heading"><span class="heading-en">STAFF PRODUCTIVITY</span><span class="heading-ja">スタッフ生産性ランキング</span></div>'
  );

  if (topPerformers.length > 0) {
    var top20 = topPerformers.slice(0, 20);

    html.push(
      '<div class="card">',
      '  <div class="card-title">&#127942; 清掃実績 Top20</div>',
      '  <div style="overflow-x:auto;"><table class="data-table">',
      '    <thead><tr><th>#</th><th>氏名</th><th>ホテル</th><th>ポジション</th><th>出勤日数</th><th>労働時間</th><th>清掃部屋数</th><th>室/日</th><th>皆勤</th></tr></thead>',
      '    <tbody>'
    );

    top20.forEach(function(p, i) {
      var medalHtml = '';
      if (i < 3) {
        medalHtml = '<span class="rank-medal rank-' + (i + 1) + '">' + (i + 1) + '</span>';
      }
      var bg = i < 3 ? ' class="perf-highlight"' : '';
      var kaikin = (p.absence_days || 0) === 0 ? '<span class="badge badge-green">皆勤</span>' : (p.absence_days || 0) + '日欠';

      html.push(
        '    <tr' + bg + '>',
        '      <td>' + medalHtml + (i + 1) + '</td>',
        '      <td style="font-weight:600;">' + esc(p.name) + '</td>',
        '      <td style="font-size:0.78rem;">' + esc(p.hotel) + '</td>',
        '      <td>' + esc(p.position || '-') + '</td>',
        '      <td style="text-align:right;">' + (p.total_days || '-') + '</td>',
        '      <td style="text-align:right;">' + num(p.total_hours, 1) + '</td>',
        '      <td style="text-align:right;font-weight:700;">' + (p.rooms_cleaned || '-') + '</td>',
        '      <td style="text-align:right;font-weight:700;color:var(--accent);">' + num(p.rooms_per_day, 1) + '</td>',
        '      <td>' + kaikin + '</td>',
        '    </tr>'
      );
    });

    html.push('    </tbody></table></div></div>');
  }

  html.push('');

  // ============================================================
  // Section 3: ホテル別スタッフ構成
  // ============================================================
  html.push(
    '<div class="section-heading"><span class="heading-en">STAFF COMPOSITION</span><span class="heading-ja">ホテル別スタッフ構成</span></div>'
  );

  if (hotelSummaries.length > 0) {
    // Sort by roster size descending
    var sortedBySize = hotelSummaries.slice().sort(function(a, b) {
      return (b.roster_size || 0) - (a.roster_size || 0);
    });
    var maxRoster = sortedBySize[0].roster_size || 1;

    html.push(
      '<div class="card">',
      '  <div class="card-title">&#128101; メイド / チェッカー構成比</div>',
      '  <p style="font-size:0.78rem;color:var(--text-light);margin-bottom:1rem;">&#128309; メイド &#128995; チェッカー &#9898; その他</p>'
    );

    sortedBySize.forEach(function(h) {
      var maid = h.maid_count || 0;
      var checker = h.checker_count || 0;
      var other = Math.max((h.roster_size || 0) - maid - checker, 0);
      var total = maid + checker + other;
      if (total === 0) return;

      var maidPct = (maid / total * 100).toFixed(0);
      var checkerPct = (checker / total * 100).toFixed(0);
      var otherPct = Math.max(100 - parseInt(maidPct) - parseInt(checkerPct), 0);

      html.push(
        '<div class="load-bar" style="margin-bottom:0.6rem;">',
        '  <div class="load-bar-label">' + esc(h.name) + '</div>',
        '  <div class="load-bar-track" style="background:transparent;">',
        '    <div class="stacked-bar" style="width:100%;">',
        '      <div class="stacked-seg" style="width:' + maidPct + '%;background:var(--blue);">' + (maid > 0 ? maid : '') + '</div>',
        '      <div class="stacked-seg" style="width:' + checkerPct + '%;background:var(--purple);">' + (checker > 0 ? checker : '') + '</div>',
        (otherPct > 0 ? '      <div class="stacked-seg" style="width:' + otherPct + '%;background:#CBD5E1;">' + (other > 0 ? other : '') + '</div>' : ''),
        '    </div>',
        '  </div>',
        '  <div class="load-bar-val">' + total + '名</div>',
        '</div>'
      );
    });

    html.push('</div>');

    // Per-staff efficiency comparison table
    html.push(
      '<div class="card">',
      '  <div class="card-title">&#128200; 一人あたり効率 × クレーム率 比較</div>',
      '  <div style="overflow-x:auto;"><table class="data-table">',
      '    <thead><tr><th>ホテル</th><th>総人数</th><th>メイド</th><th>チェッカー</th><th>室/人/日</th><th>クレーム/メイド</th><th>クレーム率</th><th>評価</th></tr></thead>',
      '    <tbody>'
    );

    // Sort by rooms_per_day for this view
    var sortedByEfficiency = hotelSummaries.slice().sort(function(a, b) {
      return (b.avg_rooms_per_day || 0) - (a.avg_rooms_per_day || 0);
    });

    sortedByEfficiency.forEach(function(h) {
      var rpd = h.avg_rooms_per_day || 0;
      var cpm = h.claims_per_maid || 0;
      // Evaluation: high efficiency + low claims = good
      var evalLabel, evalCls;
      if (rpd >= 12 && cpm <= 0.5) { evalLabel = '優秀'; evalCls = 'badge-green'; }
      else if (rpd >= 8 && cpm <= 1.0) { evalLabel = '良好'; evalCls = 'badge-blue'; }
      else if (cpm > 1.5) { evalLabel = '要改善'; evalCls = 'badge-red'; }
      else { evalLabel = '標準'; evalCls = 'badge-orange'; }

      html.push(
        '    <tr>',
        '      <td style="font-weight:600;">' + esc(h.name) + '</td>',
        '      <td style="text-align:right;">' + (h.roster_size || '-') + '</td>',
        '      <td style="text-align:right;">' + (h.maid_count || '-') + '</td>',
        '      <td style="text-align:right;">' + (h.checker_count || '-') + '</td>',
        '      <td style="text-align:right;font-weight:700;">' + num(rpd, 1) + '</td>',
        '      <td style="text-align:right;">' + num(cpm, 2) + '</td>',
        '      <td style="text-align:right;">' + pct(h.claim_rate || 0) + '</td>',
        '      <td><span class="badge ' + evalCls + '">' + evalLabel + '</span></td>',
        '    </tr>'
      );
    });

    html.push('    </tbody></table></div></div>');
  }

  html.push('');

  // ============================================================
  // Section 4: 個人パフォーマンス分析
  // ============================================================
  html.push(
    '<div class="section-heading"><span class="heading-en">INDIVIDUAL PERFORMANCE</span><span class="heading-ja">個人パフォーマンス分析</span></div>'
  );

  // Full staff list with performance data
  if (topPerformers.length > 0) {
    html.push(
      '<div class="card">',
      '  <div class="card-title">&#128203; 全スタッフ実績一覧</div>',
      '  <p style="font-size:0.78rem;color:var(--text-light);margin-bottom:1rem;">清掃部屋数データのあるスタッフ ' + topPerformers.length + '名</p>',
      '  <div style="overflow-x:auto;"><table class="data-table">',
      '    <thead><tr><th>#</th><th>氏名</th><th>ホテル</th><th>ポジション</th><th>給与形態</th><th>出勤</th><th>欠勤</th><th>労働h</th><th>清掃室数</th><th>室/日</th><th>皆勤</th></tr></thead>',
      '    <tbody>'
    );

    topPerformers.forEach(function(p, i) {
      var kaikin = (p.absence_days || 0) === 0;
      var rpd = p.rooms_per_day || 0;
      var rpdColor = rpd >= 18 ? 'color:var(--green);' : rpd >= 12 ? '' : 'color:var(--red);';

      html.push(
        '    <tr>',
        '      <td>' + (i + 1) + '</td>',
        '      <td style="font-weight:600;">' + esc(p.name) + '</td>',
        '      <td style="font-size:0.75rem;">' + esc(p.hotel) + '</td>',
        '      <td style="font-size:0.75rem;">' + esc(p.position || '-') + '</td>',
        '      <td style="font-size:0.75rem;">' + esc(p.pay_type || '-') + '</td>',
        '      <td style="text-align:right;">' + (p.total_days || '-') + '</td>',
        '      <td style="text-align:right;">' + (p.absence_days || 0) + '</td>',
        '      <td style="text-align:right;">' + num(p.total_hours, 1) + '</td>',
        '      <td style="text-align:right;font-weight:700;">' + (p.rooms_cleaned || '-') + '</td>',
        '      <td style="text-align:right;font-weight:700;' + rpdColor + '">' + num(rpd, 1) + '</td>',
        '      <td>' + (kaikin ? '<span class="badge badge-green">&#10003;</span>' : '') + '</td>',
        '    </tr>'
      );
    });

    html.push('    </tbody></table></div></div>');
  }

  // Claims detail with staff position context
  var allClaimStaff = [];
  if (maidClaims.top_claim_maids) {
    maidClaims.top_claim_maids.forEach(function(m) {
      allClaimStaff.push({ name: m.name, claims: m.claims, hotel: m.hotel, role: 'メイド' });
    });
  }
  if (checkerClaims.top_claim_checkers) {
    checkerClaims.top_claim_checkers.forEach(function(c) {
      allClaimStaff.push({ name: c.name, claims: c.claims, hotel: c.hotel, role: 'チェッカー' });
    });
  }
  allClaimStaff.sort(function(a, b) { return b.claims - a.claims; });

  if (allClaimStaff.length > 0) {
    html.push(
      '<div class="card">',
      '  <div class="card-title">&#128680; クレーム発生スタッフ（メイド＋チェッカー統合）</div>',
      '  <div style="overflow-x:auto;"><table class="data-table">',
      '    <thead><tr><th>#</th><th>氏名</th><th>ポジション</th><th>所属ホテル</th><th>クレーム件数</th><th>リスク</th></tr></thead>',
      '    <tbody>'
    );

    allClaimStaff.forEach(function(s, i) {
      var riskCls = s.claims >= 3 ? 'badge-red' : s.claims >= 2 ? 'badge-orange' : 'badge-blue';
      var riskLabel = s.claims >= 3 ? '高' : s.claims >= 2 ? '中' : '低';
      var bg = i < 3 ? ' style="background:#FEF2F2;"' : '';
      html.push(
        '    <tr' + bg + '>',
        '      <td>' + (i + 1) + '</td>',
        '      <td style="font-weight:600;">' + esc(s.name) + '</td>',
        '      <td><span class="badge ' + (s.role === 'メイド' ? 'badge-blue' : 'badge-purple') + '">' + s.role + '</span></td>',
        '      <td style="font-size:0.78rem;">' + esc(s.hotel) + '</td>',
        '      <td style="text-align:center;font-weight:700;">' + s.claims + '</td>',
        '      <td><span class="badge ' + riskCls + '">' + riskLabel + '</span></td>',
        '    </tr>'
      );
    });

    html.push('    </tbody></table></div></div>');
  }

  html.push('');

  // ============================================================
  // Section 5: 皆勤・勤怠サマリー
  // ============================================================
  html.push(
    '<div class="section-heading"><span class="heading-en">ATTENDANCE SUMMARY</span><span class="heading-ja">皆勤・勤怠サマリー</span></div>'
  );

  // Kakin rate by hotel
  if (hotelSummaries.length > 0 && topPerformers.length > 0) {
    // Group performers by hotel
    var hotelKaikin = {};
    topPerformers.forEach(function(p) {
      var key = p.hotel_key || p.hotel;
      if (!hotelKaikin[key]) hotelKaikin[key] = { name: p.hotel, total: 0, kaikin: 0, totalDays: 0, totalAbsence: 0 };
      hotelKaikin[key].total++;
      hotelKaikin[key].totalDays += p.total_days || 0;
      hotelKaikin[key].totalAbsence += p.absence_days || 0;
      if ((p.absence_days || 0) === 0) hotelKaikin[key].kaikin++;
    });

    var hotelKaikinList = Object.values(hotelKaikin).sort(function(a, b) {
      var rateA = a.total > 0 ? a.kaikin / a.total : 0;
      var rateB = b.total > 0 ? b.kaikin / b.total : 0;
      return rateB - rateA;
    });

    html.push(
      '<div class="card">',
      '  <div class="card-title">&#127942; ホテル別 皆勤該当率</div>',
      '  <div style="overflow-x:auto;"><table class="data-table">',
      '    <thead><tr><th>ホテル</th><th>スタッフ数</th><th>皆勤者</th><th>皆勤率</th><th>平均出勤日数</th><th>達成状況</th></tr></thead>',
      '    <tbody>'
    );

    hotelKaikinList.forEach(function(h) {
      var rate = h.total > 0 ? Math.round(h.kaikin / h.total * 1000) / 10 : 0;
      var avgDays = h.total > 0 ? (h.totalDays / h.total).toFixed(1) : '-';
      var badgeCls = rate >= 80 ? 'badge-green' : rate >= 50 ? 'badge-blue' : rate >= 30 ? 'badge-orange' : 'badge-red';
      var label = rate >= 80 ? '優秀' : rate >= 50 ? '良好' : rate >= 30 ? '標準' : '要改善';

      html.push(
        '    <tr>',
        '      <td style="font-weight:600;">' + esc(h.name) + '</td>',
        '      <td style="text-align:right;">' + h.total + '</td>',
        '      <td style="text-align:right;font-weight:700;">' + h.kaikin + '</td>',
        '      <td style="text-align:right;font-weight:700;">' + rate + '%</td>',
        '      <td style="text-align:right;">' + avgDays + '</td>',
        '      <td><span class="badge ' + badgeCls + '">' + label + '</span></td>',
        '    </tr>'
      );
    });

    html.push('    </tbody></table></div></div>');
  }

  // Attendance distribution histogram
  if (topPerformers.length > 0) {
    var bins = [0, 0, 0, 0, 0, 0]; // 0-9, 10-14, 15-19, 20-24, 25-29, 30+
    var binLabels = ['~9', '10-14', '15-19', '20-24', '25-29', '30+'];
    topPerformers.forEach(function(p) {
      var d = p.total_days || 0;
      if (d < 10) bins[0]++;
      else if (d < 15) bins[1]++;
      else if (d < 20) bins[2]++;
      else if (d < 25) bins[3]++;
      else if (d < 30) bins[4]++;
      else bins[5]++;
    });

    var maxBin = Math.max.apply(null, bins) || 1;

    html.push(
      '<div class="card">',
      '  <div class="card-title">&#128202; 出勤日数分布</div>',
      '  <p style="font-size:0.78rem;color:var(--text-light);margin-bottom:1rem;">全 ' + topPerformers.length + '名の出勤日数ヒストグラム</p>',
      '  <div class="hist-bar">'
    );

    var barColors = ['var(--red)', 'var(--orange)', 'var(--orange)', 'var(--blue)', 'var(--green)', 'var(--green)'];
    bins.forEach(function(count, idx) {
      var h = Math.max(Math.round(count / maxBin * 100), count > 0 ? 8 : 0);
      html.push(
        '    <div style="flex:1;text-align:center;">',
        '      <div style="font-size:0.7rem;font-weight:600;margin-bottom:0.2rem;">' + count + '</div>',
        '      <div style="height:' + h + 'px;background:' + barColors[idx] + ';border-radius:4px 4px 0 0;margin:0 2px;"></div>',
        '      <div style="font-size:0.6rem;color:var(--text-light);margin-top:0.3rem;border-top:1px solid var(--border);padding-top:0.2rem;">' + binLabels[idx] + '日</div>',
        '    </div>'
      );
    });

    html.push(
      '  </div>',
      '</div>'
    );
  }

  // ── Close ──
  html.push('</div>', footer(), pageFoot());

  return html.join('\n');
}

module.exports = { buildTalentManagement };
