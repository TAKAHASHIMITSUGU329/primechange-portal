// Deep Analysis custom renderers for 7 analysis tabs
var { esc } = require('./common');

// ====== Shared helpers ======

function metaBox(m) {
  var html = [];
  html.push('<div style="background:#F1F5F9;padding:1rem;border-radius:8px;margin-bottom:1.5rem;">');
  html.push('<div style="font-weight:700;font-size:0.95rem;">' + esc(m.title || '') + '</div>');
  if (m.subtitle) html.push('<div style="font-size:0.75rem;color:var(--text-light);">' + esc(m.subtitle) + '</div>');
  var stats = [];
  if (m.total_hotels) stats.push(m.total_hotels + 'ホテル');
  if (m.total_claims) stats.push(m.total_claims + 'クレーム');
  if (m.total_rooms_cleaned) stats.push(Number(m.total_rooms_cleaned).toLocaleString() + '室清掃');
  if (m.total_staff_analyzed) stats.push(m.total_staff_analyzed + '名分析');
  if (m.data_period) stats.push(m.data_period);
  if (stats.length) html.push('<div style="font-size:0.75rem;color:var(--text-light);margin-top:0.3rem;">' + stats.join(' / ') + '</div>');
  html.push('</div>');
  return html.join('\n');
}

function section(title, content) {
  return '<div class="card"><div class="card-title">' + title + '</div>' + content + '</div>';
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

function corrBadge(r) {
  if (r == null) return '<span class="badge badge-gray">N/A</span>';
  var abs = Math.abs(r);
  var cls = abs >= 0.5 ? 'badge-green' : abs >= 0.3 ? 'badge-orange' : 'badge-red';
  var label = abs >= 0.5 ? '強い' : abs >= 0.3 ? '中程度' : '弱い';
  return '<span class="badge ' + cls + '">' + label + ' r=' + r.toFixed(3) + '</span>';
}

function tierBadge(tier) {
  var cls = 'badge-gray';
  if (tier === '優秀') cls = 'badge-green';
  else if (tier === '良好') cls = 'badge-blue';
  else if (tier === '概ね良好') cls = 'badge-orange';
  else if (tier === '要改善' || tier === '要緊急対応') cls = 'badge-red';
  return '<span class="badge ' + cls + '">' + esc(tier) + '</span>';
}

function priBadge(pri) {
  var cls = 'badge-gray';
  if (pri === 'URGENT') cls = 'badge-red';
  else if (pri === 'HIGH') cls = 'badge-orange';
  else if (pri === 'STANDARD') cls = 'badge-blue';
  else if (pri === 'MAINTENANCE') cls = 'badge-green';
  return '<span class="badge ' + cls + '">' + esc(pri) + '</span>';
}

function hbar(label, value, max, color) {
  var w = max > 0 ? Math.min(value / max * 100, 100) : 0;
  return '<div class="h-bar"><div class="h-bar-label">' + esc(label) + '</div><div class="h-bar-track"><div class="h-bar-fill" style="width:' + w + '%;background:' + color + ';"><span class="h-bar-val">' + value + '</span></div></div></div>';
}

function renderRecs(recs) {
  if (!recs || !recs.length) return '';
  var html = [];
  recs.forEach(function(r, i) {
    var priCls = r.priority === 'HIGH' || r.priority === '高' ? 'badge-red' : r.priority === 'MEDIUM' || r.priority === '中' ? 'badge-orange' : 'badge-blue';
    html.push('<div class="accordion-item' + (i === 0 ? ' open' : '') + '">');
    html.push('<div class="accordion-header" onclick="this.parentElement.classList.toggle(\'open\')">' + esc(r.title) + ' <span class="badge ' + priCls + '">' + esc(r.priority || '') + '</span><span class="accordion-arrow">&#9660;</span></div>');
    html.push('<div class="accordion-body">');
    if (r.rationale) html.push('<p style="font-size:0.8rem;color:var(--text-light);margin-bottom:0.5rem;">' + esc(r.rationale) + '</p>');
    if (r.actions && r.actions.length) {
      html.push('<ul style="font-size:0.8rem;padding-left:1.2rem;">');
      r.actions.forEach(function(a) { html.push('<li style="margin-bottom:0.3rem;">' + esc(a) + '</li>'); });
      html.push('</ul>');
    }
    html.push('</div></div>');
  });
  return section('&#128161; 改善提言', html.join('\n'));
}

// ====== Analysis 1: Claims × Score ======
function renderA1(d) {
  if (!d) return '<p>データがありません</p>';
  var html = [];
  if (d.analysis_metadata) html.push(metaBox(d.analysis_metadata));

  // Summary stats
  if (d.summary_stats) {
    var s = d.summary_stats;
    html.push('<div class="kpi-grid">');
    html.push('<div class="kpi-card"><div class="kpi-label">総クレーム数</div><div class="kpi-value">' + (d.analysis_metadata.total_claims || 0) + '</div><div class="kpi-sub">件</div></div>');
    html.push('<div class="kpi-card"><div class="kpi-label">クレーム発生率</div><div class="kpi-value">' + pct(s.avg_claim_rate) + '</div><div class="kpi-sub">平均</div></div>');
    html.push('<div class="kpi-card" style="border-left-color:var(--red);"><div class="kpi-label">最高発生率</div><div class="kpi-value">' + pct(s.max_claim_rate) + '</div><div class="kpi-sub">' + esc(s.max_claim_rate_hotel || '') + '</div></div>');
    html.push('<div class="kpi-card" style="border-left-color:var(--orange);"><div class="kpi-label">最多クレーム類型</div><div class="kpi-value">' + esc(s.top_claim_type || '') + '</div><div class="kpi-sub">' + (s.top_claim_count || 0) + '件</div></div>');
    html.push('</div>');
  }

  // Type frequency ranking - horizontal bars
  if (d.type_frequency_ranking) {
    var maxCount = Math.max.apply(null, d.type_frequency_ranking.map(function(t) { return t.total_count; }));
    var bars = d.type_frequency_ranking.map(function(t) {
      return hbar(t.type, t.total_count, maxCount, 'var(--accent)') +
        '<div style="font-size:0.65rem;color:var(--text-light);margin-left:126px;margin-top:-0.3rem;margin-bottom:0.3rem;">' + pct(t.share_pct) + ' / ' + t.hotels_affected + 'ホテル</div>';
    }).join('');
    html.push(section('&#128202; クレーム類型ランキング', bars));
  }

  // Category breakdown
  if (d.category_breakdown) {
    var catHtml = '<div class="grid-2">';
    var colors = { '客室準備系': 'var(--blue)', '清潔性系': 'var(--orange)', '安全・設備系': 'var(--red)', 'その他': 'var(--text-light)' };
    Object.keys(d.category_breakdown).forEach(function(cat) {
      var c = d.category_breakdown[cat];
      catHtml += '<div style="background:var(--bg);border-radius:8px;padding:1rem;border-left:4px solid ' + (colors[cat] || 'var(--accent)') + ';">';
      catHtml += '<div style="font-weight:700;font-size:0.85rem;">' + esc(cat) + '</div>';
      catHtml += '<div style="font-size:1.3rem;font-weight:800;margin:0.3rem 0;">' + c.count + '件 <span style="font-size:0.8rem;color:var(--text-light);">(' + pct(c.share_pct) + ')</span></div>';
      if (c.types && c.types.length) {
        catHtml += '<div style="font-size:0.75rem;color:var(--text-light);">' + c.types.map(esc).join('、') + '</div>';
      }
      catHtml += '</div>';
    });
    catHtml += '</div>';
    html.push(section('&#128203; カテゴリ別集計', catHtml));
  }

  // Correlation
  if (d.correlation_analysis) {
    var ca = d.correlation_analysis;
    var corrHtml = '';
    if (ca.overall_correlation) {
      corrHtml += '<div style="margin-bottom:1rem;padding:1rem;background:var(--bg);border-radius:8px;">';
      corrHtml += '<div style="font-weight:700;margin-bottom:0.5rem;">全体相関</div>';
      corrHtml += '<div>' + corrBadge(ca.overall_correlation.r) + ' <span style="font-size:0.8rem;margin-left:0.5rem;">' + esc(ca.overall_correlation.interpretation || '') + '</span></div>';
      corrHtml += '</div>';
    }
    if (ca.type_correlations && ca.type_correlations.length) {
      corrHtml += '<table class="data-table"><thead><tr><th>クレーム類型</th><th>相関係数</th><th>方向</th><th>解釈</th></tr></thead><tbody>';
      ca.type_correlations.forEach(function(tc) {
        corrHtml += '<tr><td>' + esc(tc.type) + '</td><td>' + corrBadge(tc.correlation_r) + '</td><td>' + esc(tc.direction || '') + '</td><td style="font-size:0.75rem;">' + esc(tc.interpretation || '') + '</td></tr>';
      });
      corrHtml += '</tbody></table>';
    }
    html.push(section('&#128200; 相関分析', corrHtml));
  }

  // Hotel profiles table
  if (d.hotel_profiles) {
    var tbl = '<div style="overflow-x:auto;"><table class="data-table"><thead><tr><th>ホテル</th><th>クレーム数</th><th>発生率</th><th>清掃室数</th><th>主なクレーム類型</th></tr></thead><tbody>';
    d.hotel_profiles.forEach(function(h) {
      var topTypes = (h.top_types || []).slice(0, 3).map(function(t) { return esc(t.type) + '(' + t.count + ')'; }).join(', ');
      tbl += '<tr><td>' + esc(h.name) + '</td><td style="text-align:right;">' + h.total_claims + '</td><td style="text-align:right;">' + pct(h.claim_rate) + '</td><td style="text-align:right;">' + num(h.rooms_cleaned) + '</td><td style="font-size:0.75rem;">' + topTypes + '</td></tr>';
    });
    tbl += '</tbody></table></div>';
    html.push(section('&#127976; ホテル別クレーム一覧', tbl));
  }

  // Improvement priorities
  if (d.improvement_priorities) {
    var priTbl = '<div style="overflow-x:auto;"><table class="data-table"><thead><tr><th>#</th><th>クレーム類型</th><th>件数</th><th>影響ホテル数</th><th>相関r</th></tr></thead><tbody>';
    d.improvement_priorities.forEach(function(p, i) {
      priTbl += '<tr><td>' + (i + 1) + '</td><td style="font-weight:600;">' + esc(p.type) + '</td><td style="text-align:right;">' + p.total_count + '</td><td style="text-align:right;">' + p.hotels_affected + '</td><td>' + corrBadge(p.correlation_r) + '</td></tr>';
    });
    priTbl += '</tbody></table></div>';
    html.push(section('&#127919; 改善優先度ランキング', priTbl));
  }

  html.push(renderRecs(d.recommendations));
  return html.join('\n');
}

// ====== Analysis 2: Staff Performance ======
function renderA2(d) {
  if (!d) return '<p>データがありません</p>';
  var html = [];
  if (d.analysis_metadata) html.push(metaBox(d.analysis_metadata));

  // KPIs
  html.push('<div class="kpi-grid">');
  html.push('<div class="kpi-card"><div class="kpi-label">分析スタッフ数</div><div class="kpi-value">' + (d.analysis_metadata.total_staff_analyzed || 0) + '</div><div class="kpi-sub">名</div></div>');
  if (d.maid_claims_summary) html.push('<div class="kpi-card" style="border-left-color:var(--orange);"><div class="kpi-label">クレーム有りメイド</div><div class="kpi-value">' + d.maid_claims_summary.total_maids_with_claims + '</div><div class="kpi-sub">名</div></div>');
  if (d.maid_productivity) {
    html.push('<div class="kpi-card" style="border-left-color:var(--green);"><div class="kpi-label">平均清掃室数/日</div><div class="kpi-value">' + num(d.maid_productivity.avg_rooms_per_day) + '</div><div class="kpi-sub">室</div></div>');
    html.push('<div class="kpi-card"><div class="kpi-label">最大清掃室数/日</div><div class="kpi-value">' + num(d.maid_productivity.max_rooms_per_day) + '</div><div class="kpi-sub">室</div></div>');
  }
  if (d.attendance_analysis) {
    html.push('<div class="kpi-card"><div class="kpi-label">平均出勤日数</div><div class="kpi-value">' + num(d.attendance_analysis.avg_days) + '</div><div class="kpi-sub">日</div></div>');
  }
  html.push('</div>');

  // Maid claims top table
  if (d.maid_claims_summary && d.maid_claims_summary.top_claim_maids) {
    var maids = d.maid_claims_summary.top_claim_maids;
    var tbl = '<div style="overflow-x:auto;"><table class="data-table"><thead><tr><th>#</th><th>スタッフ名</th><th>クレーム件数</th><th>所属ホテル</th></tr></thead><tbody>';
    maids.forEach(function(m, i) {
      var bg = i < 3 ? ' style="background:#FEF2F2;"' : '';
      tbl += '<tr' + bg + '><td>' + (i + 1) + '</td><td style="font-weight:600;">' + esc(m.name) + '</td><td style="text-align:center;"><span class="badge badge-red">' + m.claims + '</span></td><td style="font-size:0.8rem;">' + esc(m.hotel) + '</td></tr>';
    });
    tbl += '</tbody></table></div>';
    html.push(section('&#9888;&#65039; メイド別クレームTop15', tbl));
  }

  // Checker claims top table
  if (d.checker_claims_summary && d.checker_claims_summary.top_claim_checkers) {
    var checkers = d.checker_claims_summary.top_claim_checkers;
    var tbl2 = '<div style="overflow-x:auto;"><table class="data-table"><thead><tr><th>#</th><th>チェッカー名</th><th>クレーム件数</th><th>所属ホテル</th></tr></thead><tbody>';
    checkers.forEach(function(c, i) {
      var bg = i < 3 ? ' style="background:#FEF2F2;"' : '';
      tbl2 += '<tr' + bg + '><td>' + (i + 1) + '</td><td style="font-weight:600;">' + esc(c.name) + '</td><td style="text-align:center;"><span class="badge badge-orange">' + c.claims + '</span></td><td style="font-size:0.8rem;">' + esc(c.hotel) + '</td></tr>';
    });
    tbl2 += '</tbody></table></div>';
    html.push(section('&#128269; チェッカー別クレームTop15', tbl2));
  }

  // Maid productivity top performers
  if (d.maid_productivity && d.maid_productivity.top_performers) {
    var perf = d.maid_productivity.top_performers;
    var tbl3 = '<div style="overflow-x:auto;"><table class="data-table"><thead><tr><th>#</th><th>スタッフ名</th><th>職種</th><th>雇用</th><th>出勤日数</th><th>総清掃室数</th><th>2月</th><th>3月</th><th>室/日</th></tr></thead><tbody>';
    perf.forEach(function(p, i) {
      tbl3 += '<tr><td>' + (i + 1) + '</td><td style="font-weight:600;">' + esc(p.name) + '</td><td>' + esc(p.position || '') + '</td><td>' + esc(p.pay_type || '') + '</td><td style="text-align:right;">' + num(p.total_days) + '</td><td style="text-align:right;font-weight:700;">' + num(p.rooms_cleaned) + '</td><td style="text-align:right;">' + num(p.feb_rooms) + '</td><td style="text-align:right;">' + num(p.mar_rooms) + '</td><td style="text-align:right;color:var(--accent);font-weight:700;">' + num(p.rooms_per_day) + '</td></tr>';
    });
    tbl3 += '</tbody></table></div>';
    html.push(section('&#127941; 清掃生産性Top10', tbl3));
  }

  // Hotel summaries
  if (d.hotel_summaries) {
    var tbl4 = '<div style="overflow-x:auto;"><table class="data-table"><thead><tr><th>ホテル</th><th>人員</th><th>メイド</th><th>チェッカー</th><th>クレーム</th><th>クレーム/メイド</th><th>平均清掃室/日</th></tr></thead><tbody>';
    d.hotel_summaries.forEach(function(h) {
      tbl4 += '<tr><td>' + esc(h.name) + '</td><td style="text-align:right;">' + h.roster_size + '</td><td style="text-align:right;">' + h.maid_count + '</td><td style="text-align:right;">' + h.checker_count + '</td><td style="text-align:right;">' + h.total_claims + '</td><td style="text-align:right;">' + num(h.claims_per_maid) + '</td><td style="text-align:right;">' + num(h.avg_rooms_per_day) + '</td></tr>';
    });
    tbl4 += '</tbody></table></div>';
    html.push(section('&#127976; ホテル別スタッフ概要', tbl4));
  }

  html.push(renderRecs(d.recommendations));
  return html.join('\n');
}

// ====== Analysis 3: Staffing × Quality ======
function renderA3(d) {
  if (!d) return '<p>データがありません</p>';
  var html = [];
  if (d.analysis_metadata) html.push(metaBox(d.analysis_metadata));

  // Correlations
  if (d.staffing_analysis && d.staffing_analysis.correlations) {
    var corrs = d.staffing_analysis.correlations;
    var corrLabels = {
      'staff_vs_score': '総スタッフ数 vs スコア',
      'maids_vs_score': 'メイド数 vs スコア',
      'checkers_vs_score': 'チェッカー数 vs スコア',
      'ratio_vs_score': 'メイド/チェッカー比 vs スコア',
      'staff_vs_claims': 'スタッフ数 vs クレーム率',
    };
    var corrCards = '<div class="kpi-grid">';
    Object.keys(corrs).forEach(function(k) {
      var c = corrs[k];
      corrCards += '<div class="kpi-card"><div class="kpi-label">' + esc(corrLabels[k] || k) + '</div><div class="kpi-value" style="font-size:1.2rem;">' + corrBadge(c.r) + '</div><div class="kpi-sub">R&sup2;=' + num(c.r_squared, 4) + ' / n=' + c.n + '</div></div>';
    });
    corrCards += '</div>';
    html.push(section('&#128200; 相関分析結果', corrCards));
  }

  // Staffing tiers
  if (d.staffing_analysis && d.staffing_analysis.staffing_tiers) {
    var tiers = d.staffing_analysis.staffing_tiers;
    var tierHtml = '<div class="grid-2">';
    [['low_staff', '少人数配置グループ', 'var(--orange)'], ['high_staff', '多人数配置グループ', 'var(--green)']].forEach(function(t) {
      var tier = tiers[t[0]];
      if (!tier) return;
      tierHtml += '<div style="padding:1.25rem;background:var(--bg);border-radius:8px;border-left:4px solid ' + t[2] + ';">';
      tierHtml += '<div style="font-weight:700;margin-bottom:0.5rem;">' + t[1] + '</div>';
      tierHtml += '<div style="font-size:0.85rem;">平均メイド数: <strong>' + num(tier.avg_maids) + '名</strong></div>';
      tierHtml += '<div style="font-size:0.85rem;">平均スコア: <strong>' + num(tier.avg_score) + '</strong></div>';
      tierHtml += '<div style="font-size:0.85rem;">平均クレーム率: <strong>' + pct(tier.avg_claim_rate) + '</strong></div>';
      tierHtml += '</div>';
    });
    tierHtml += '</div>';
    html.push(section('&#128202; 配置グループ比較', tierHtml));
  }

  // Optimal staffing top5/bottom5
  if (d.staffing_analysis && d.staffing_analysis.optimal_staffing) {
    var opt = d.staffing_analysis.optimal_staffing;
    var optHtml = '<div class="grid-2">';
    [['top5_hotels', 'Top5（高品質・適正配置）', 'var(--green)'], ['bottom5_hotels', 'Bottom5（改善余地あり）', 'var(--red)']].forEach(function(pair) {
      var list = opt[pair[0]];
      if (!list) return;
      optHtml += '<div><h4 style="font-size:0.85rem;color:' + pair[2] + ';margin-bottom:0.5rem;">' + pair[1] + '</h4>';
      optHtml += '<table class="data-table"><thead><tr><th>ホテル</th><th>メイド</th><th>チェッカー</th><th>比率</th><th>スコア</th></tr></thead><tbody>';
      list.forEach(function(h) {
        optHtml += '<tr><td>' + esc(h.name) + '</td><td style="text-align:right;">' + num(h.avg_maids) + '</td><td style="text-align:right;">' + num(h.avg_checkers) + '</td><td style="text-align:right;">' + num(h.ratio) + '</td><td style="text-align:right;font-weight:700;">' + num(h.score) + '</td></tr>';
      });
      optHtml += '</tbody></table></div>';
    });
    optHtml += '</div>';
    html.push(section('&#127919; 配置最適化ランキング', optHtml));
  }

  // Hotel summary table
  if (d.hotel_summary) {
    var tbl = '<div style="overflow-x:auto;"><table class="data-table"><thead><tr><th>ホテル</th><th>日数</th><th>平均メイド</th><th>平均チェッカー</th><th>M/C比</th><th>完了時間</th><th>スコア</th><th>クレーム率</th></tr></thead><tbody>';
    d.hotel_summary.forEach(function(h) {
      tbl += '<tr><td>' + esc(h.name) + '</td><td style="text-align:right;">' + h.days + '</td><td style="text-align:right;">' + num(h.avg_maids) + '</td><td style="text-align:right;">' + num(h.avg_checkers) + '</td><td style="text-align:right;">' + num(h.maid_checker_ratio) + '</td><td style="text-align:right;">' + num(h.avg_completion_time) + '</td><td style="text-align:right;font-weight:700;">' + num(h.score) + '</td><td style="text-align:right;">' + pct(h.claim_rate) + '</td></tr>';
    });
    tbl += '</tbody></table></div>';
    html.push(section('&#127976; ホテル別配置データ', tbl));
  }

  html.push(renderRecs(d.recommendations));
  return html.join('\n');
}

// ====== Analysis 4: Completion Time × Quality ======
function renderA4(d) {
  if (!d) return '<p>データがありません</p>';
  var html = [];
  if (d.analysis_metadata) html.push(metaBox(d.analysis_metadata));

  // KPIs
  if (d.time_analysis) {
    var ta = d.time_analysis;
    html.push('<div class="kpi-grid">');
    html.push('<div class="kpi-card"><div class="kpi-label">分析ホテル数</div><div class="kpi-value">' + ta.total_hotels_analyzed + '</div></div>');
    html.push('<div class="kpi-card"><div class="kpi-label">全体平均完了時間</div><div class="kpi-value">' + num(ta.overall_avg_time) + '時</div></div>');
    if (ta.overall_time_range) html.push('<div class="kpi-card"><div class="kpi-label">時間帯レンジ</div><div class="kpi-value" style="font-size:1rem;">' + esc(ta.overall_time_range) + '</div></div>');
    html.push('</div>');
  }

  // Correlations
  if (d.time_analysis && d.time_analysis.correlations) {
    var corrs = d.time_analysis.correlations;
    var corrLabels = { 'time_vs_score': '完了時間 vs スコア', 'time_vs_claims': '完了時間 vs クレーム率', 'time_variability_vs_score': '時間ばらつき vs スコア' };
    var corrCards = '<div class="kpi-grid">';
    Object.keys(corrs).forEach(function(k) {
      var c = corrs[k];
      corrCards += '<div class="kpi-card"><div class="kpi-label">' + esc(corrLabels[k] || k) + '</div><div class="kpi-value" style="font-size:1.2rem;">' + corrBadge(c.r) + '</div><div class="kpi-sub">R&sup2;=' + num(c.r_squared, 4) + '</div></div>';
    });
    corrCards += '</div>';
    html.push(section('&#128200; 相関分析', corrCards));
  }

  // Time tiers
  if (d.time_analysis && d.time_analysis.time_tiers) {
    var tiers = d.time_analysis.time_tiers;
    var tierLabels = { 'early_finish': ['早期完了', 'var(--green)'], 'mid_finish': ['中間', 'var(--blue)'], 'late_finish': ['遅延完了', 'var(--red)'] };
    var tierHtml = '<div class="grid-3">';
    ['early_finish', 'mid_finish', 'late_finish'].forEach(function(k) {
      var t = tiers[k];
      if (!t) return;
      var lbl = tierLabels[k];
      tierHtml += '<div style="padding:1.25rem;background:var(--bg);border-radius:8px;border-top:4px solid ' + lbl[1] + ';">';
      tierHtml += '<div style="font-weight:700;color:' + lbl[1] + ';margin-bottom:0.5rem;">' + lbl[0] + ' (' + esc(t.label || '') + ')</div>';
      tierHtml += '<div style="font-size:0.85rem;">' + t.count + 'ホテル / 平均 ' + num(t.avg_time) + '時</div>';
      tierHtml += '<div style="font-size:0.85rem;">平均スコア: <strong>' + num(t.avg_score) + '</strong></div>';
      tierHtml += '<div style="font-size:0.85rem;">平均クレーム率: <strong>' + pct(t.avg_claim_rate) + '</strong></div>';
      if (t.hotels && t.hotels.length) {
        tierHtml += '<div style="font-size:0.72rem;color:var(--text-light);margin-top:0.5rem;">' + t.hotels.map(esc).join('、') + '</div>';
      }
      tierHtml += '</div>';
    });
    tierHtml += '</div>';
    html.push(section('&#9200; 完了時間帯別比較', tierHtml));
  }

  // Benchmark fastest / slowest
  if (d.time_analysis && d.time_analysis.benchmark) {
    var bm = d.time_analysis.benchmark;
    var bmHtml = '<div class="grid-2">';
    [['fastest_5', '最速Top5', 'var(--green)'], ['slowest_5', '最遅Bottom5', 'var(--red)']].forEach(function(pair) {
      var list = bm[pair[0]];
      if (!list) return;
      bmHtml += '<div><h4 style="font-size:0.85rem;color:' + pair[2] + ';margin-bottom:0.5rem;">' + pair[1] + '</h4>';
      bmHtml += '<table class="data-table"><thead><tr><th>ホテル</th><th>完了時間</th><th>スコア</th></tr></thead><tbody>';
      list.forEach(function(h) {
        bmHtml += '<tr><td>' + esc(h.name) + '</td><td style="text-align:right;">' + num(h.time) + '時</td><td style="text-align:right;font-weight:700;">' + num(h.score) + '</td></tr>';
      });
      bmHtml += '</tbody></table></div>';
    });
    bmHtml += '</div>';
    if (bm.score_difference != null) {
      bmHtml += '<div style="font-size:0.82rem;color:var(--text-light);margin-top:0.75rem;text-align:center;">最速5 平均スコア: <strong>' + num(bm.fastest_avg_score) + '</strong> / 最遅5 平均スコア: <strong>' + num(bm.slowest_avg_score) + '</strong> (差: ' + num(bm.score_difference) + ')</div>';
    }
    html.push(section('&#127942; ベンチマーク比較', bmHtml));
  }

  // Data points table
  if (d.time_analysis && d.time_analysis.data_points) {
    var dp = d.time_analysis.data_points;
    var tbl = '<div style="overflow-x:auto;"><table class="data-table"><thead><tr><th>ホテル</th><th>平均完了時間</th><th>最早</th><th>最遅</th><th>標準偏差</th><th>データ数</th><th>スコア</th><th>クレーム率</th></tr></thead><tbody>';
    dp.forEach(function(h) {
      tbl += '<tr><td>' + esc(h.name) + '</td><td style="text-align:right;font-weight:700;">' + num(h.avg_completion_time) + '</td><td style="text-align:right;">' + num(h.min_time) + '</td><td style="text-align:right;">' + num(h.max_time) + '</td><td style="text-align:right;">' + num(h.std_time) + '</td><td style="text-align:right;">' + h.time_data_points + '</td><td style="text-align:right;">' + num(h.score) + '</td><td style="text-align:right;">' + pct(h.cleaning_issue_rate) + '</td></tr>';
    });
    tbl += '</tbody></table></div>';
    html.push(section('&#127976; ホテル別時間データ', tbl));
  }

  html.push(renderRecs(d.recommendations));
  return html.join('\n');
}

// ====== Analysis 5: Safety × Early Warning ======
function renderA5(d) {
  if (!d) return '<p>データがありません</p>';
  var html = [];
  if (d.analysis_metadata) html.push(metaBox(d.analysis_metadata));

  // KPIs
  html.push('<div class="kpi-grid">');
  html.push('<div class="kpi-card"><div class="kpi-label">総問題項目</div><div class="kpi-value">' + (d.total_problem_count || 0) + '</div><div class="kpi-sub">件</div></div>');
  if (d.correlations && d.correlations.safety_vs_review) {
    html.push('<div class="kpi-card"><div class="kpi-label">安全スコア vs 口コミ</div><div class="kpi-value" style="font-size:1.2rem;">' + corrBadge(d.correlations.safety_vs_review.r) + '</div></div>');
  }
  if (d.correlations && d.correlations.safety_vs_claims) {
    html.push('<div class="kpi-card"><div class="kpi-label">安全スコア vs クレーム率</div><div class="kpi-value" style="font-size:1.2rem;">' + corrBadge(d.correlations.safety_vs_claims.r) + '</div></div>');
  }
  html.push('</div>');

  // Hotel ranking
  if (d.hotel_ranking) {
    var tbl = '<div style="overflow-x:auto;"><table class="data-table"><thead><tr><th>#</th><th>ホテル</th><th>安全スコア</th><th>口コミスコア</th><th>クレーム率</th><th>点検回数</th></tr></thead><tbody>';
    d.hotel_ranking.forEach(function(h, i) {
      var scoreColor = h.safety_score >= 4.0 ? 'var(--green)' : h.safety_score >= 3.0 ? 'var(--blue)' : h.safety_score >= 2.0 ? 'var(--orange)' : 'var(--red)';
      tbl += '<tr><td>' + (i + 1) + '</td><td>' + esc(h.name) + '</td><td style="text-align:right;color:' + scoreColor + ';font-weight:700;">' + num(h.safety_score) + '</td><td style="text-align:right;">' + num(h.review_score) + '</td><td style="text-align:right;">' + pct(h.claim_rate) + '</td><td style="text-align:right;">' + h.inspections + '</td></tr>';
    });
    tbl += '</tbody></table></div>';
    html.push(section('&#128737;&#65039; 安全スコアランキング', tbl));
  }

  // Hotel details with category scores
  if (d.hotel_details) {
    var detHtml = '';
    var keys = Object.keys(d.hotel_details);
    keys.forEach(function(k, idx) {
      var h = d.hotel_details[k];
      detHtml += '<div class="accordion-item' + (idx < 3 ? ' open' : '') + '">';
      detHtml += '<div class="accordion-header" onclick="this.parentElement.classList.toggle(\'open\')">';
      var scoreColor = h.avg_score >= 4.0 ? 'var(--green)' : h.avg_score >= 3.0 ? 'var(--blue)' : h.avg_score >= 2.0 ? 'var(--orange)' : 'var(--red)';
      detHtml += '<span>' + esc(h.name) + ' <span style="color:' + scoreColor + ';font-weight:800;">' + num(h.avg_score) + '</span> (' + h.inspections_count + '回点検)</span>';
      detHtml += '<span class="accordion-arrow">&#9660;</span></div>';
      detHtml += '<div class="accordion-body">';
      if (h.category_scores) {
        detHtml += '<div class="kpi-grid">';
        Object.keys(h.category_scores).forEach(function(cat) {
          var score = h.category_scores[cat];
          var catColor = score >= 4.0 ? 'var(--green)' : score >= 3.0 ? 'var(--blue)' : score >= 2.0 ? 'var(--orange)' : 'var(--red)';
          detHtml += '<div class="kpi-card" style="border-left-color:' + catColor + ';"><div class="kpi-label">' + esc(cat) + '</div><div class="kpi-value" style="color:' + catColor + ';">' + num(score) + '</div><div class="kpi-sub">/5.0</div></div>';
        });
        detHtml += '</div>';
      }
      detHtml += '</div></div>';
    });
    html.push(section('&#128196; ホテル別安全カテゴリスコア', detHtml));
  }

  // Problem items table
  if (d.problem_items && d.problem_items.length) {
    var probTbl = '<div style="overflow-x:auto;max-height:500px;overflow-y:auto;"><table class="data-table"><thead><tr><th>ホテル</th><th>カテゴリ</th><th>項目</th><th>評価</th><th>スコア</th></tr></thead><tbody>';
    d.problem_items.forEach(function(p) {
      var valColor = p.value === '✖' ? 'var(--red)' : 'var(--orange)';
      probTbl += '<tr><td style="font-size:0.75rem;">' + esc(p.hotel) + '</td><td style="font-size:0.75rem;">' + esc(p.category) + '</td><td style="font-size:0.75rem;">' + esc(p.item) + '</td><td style="text-align:center;color:' + valColor + ';font-weight:700;">' + esc(p.value) + '</td><td style="text-align:center;">' + p.score + '</td></tr>';
    });
    probTbl += '</tbody></table></div>';
    html.push(section('&#9888;&#65039; 問題項目一覧 (' + d.problem_items.length + '件)', probTbl));
  }

  html.push(renderRecs(d.recommendations));
  return html.join('\n');
}

// ====== Analysis 6: Quality → Revenue ======
function renderA6(d) {
  if (!d) return '<p>データがありません</p>';
  var html = [];
  if (d.analysis_metadata) html.push(metaBox(d.analysis_metadata));

  // Portfolio summary
  if (d.portfolio_summary) {
    var ps = d.portfolio_summary;
    html.push('<div class="kpi-grid">');
    html.push('<div class="kpi-card"><div class="kpi-label">月間売上合計</div><div class="kpi-value">&#165;' + Math.round(ps.total_monthly_revenue / 10000).toLocaleString() + '万</div></div>');
    html.push('<div class="kpi-card" style="border-left-color:var(--green);"><div class="kpi-label">平均稼働率</div><div class="kpi-value">' + esc(String(ps.avg_occupancy)) + '</div></div>');
    html.push('<div class="kpi-card" style="border-left-color:var(--blue);"><div class="kpi-label">平均ADR</div><div class="kpi-value">&#165;' + Math.round(ps.avg_adr).toLocaleString() + '</div></div>');
    html.push('<div class="kpi-card"><div class="kpi-label">平均RevPAR</div><div class="kpi-value">&#165;' + Math.round(ps.avg_revpar).toLocaleString() + '</div></div>');
    html.push('</div>');
  }

  // Regression results
  if (d.regression_results) {
    var reg = d.regression_results;
    var regHtml = '<div style="overflow-x:auto;"><table class="data-table"><thead><tr><th>指標</th><th>傾き</th><th>相関係数 r</th><th>R&sup2;</th><th>解釈</th></tr></thead><tbody>';
    Object.keys(reg).forEach(function(k) {
      var r = reg[k];
      regHtml += '<tr><td style="font-weight:600;">' + esc(r.y_label || k) + '</td><td style="text-align:right;">' + num(r.slope) + '</td><td>' + corrBadge(r.r) + '</td><td style="text-align:right;">' + num(r.r_squared, 4) + '</td><td style="font-size:0.75rem;">' + esc(r.interpretation || '') + '</td></tr>';
    });
    regHtml += '</tbody></table></div>';
    html.push(section('&#128200; 回帰分析結果', regHtml));
  }

  // Threshold analysis
  if (d.threshold_analysis && d.threshold_analysis.groups) {
    var groups = d.threshold_analysis.groups;
    var thHtml = '<div style="overflow-x:auto;"><table class="data-table"><thead><tr><th>スコア帯</th><th>ホテル数</th><th>平均稼働率</th><th>平均ADR</th><th>平均RevPAR</th><th>平均月間売上</th></tr></thead><tbody>';
    groups.forEach(function(g) {
      thHtml += '<tr><td style="font-weight:600;">' + esc(g.range) + '</td><td style="text-align:right;">' + g.count + '</td><td style="text-align:right;">' + esc(String(g.avg_occupancy_pct)) + '</td><td style="text-align:right;">&#165;' + Math.round(g.avg_adr).toLocaleString() + '</td><td style="text-align:right;">&#165;' + Math.round(g.avg_revpar).toLocaleString() + '</td><td style="text-align:right;">&#165;' + Math.round(g.avg_revenue).toLocaleString() + '</td></tr>';
    });
    thHtml += '</tbody></table></div>';
    if (d.threshold_analysis.threshold_effect) {
      thHtml += '<div style="margin-top:0.75rem;padding:0.75rem;background:#FFFBEB;border-radius:8px;font-size:0.82rem;"><strong>閾値効果:</strong> ' + esc(d.threshold_analysis.threshold_effect.description || '') + '</div>';
    }
    html.push(section('&#128201; スコア帯別パフォーマンス', thHtml));
  }

  // Revenue scenarios
  if (d.revenue_impact_scenarios) {
    var scHtml = '';
    d.revenue_impact_scenarios.forEach(function(sc) {
      scHtml += '<h3 style="font-size:0.9rem;font-weight:700;color:var(--accent);margin:1.25rem 0 0.5rem;">スコア +' + sc.score_improvement + '点改善シナリオ</h3>';
      scHtml += '<div class="kpi-grid" style="margin-bottom:0.75rem;">';
      scHtml += '<div class="kpi-card" style="border-left-color:var(--green);"><div class="kpi-label">RevPAR変動</div><div class="kpi-value" style="color:var(--green);">+&#165;' + Math.round(sc.revpar_change).toLocaleString() + '</div><div class="kpi-sub">' + esc(String(sc.revpar_pct_change)) + '</div></div>';
      scHtml += '<div class="kpi-card" style="border-left-color:var(--green);"><div class="kpi-label">月間売上変動</div><div class="kpi-value" style="color:var(--green);">+&#165;' + Math.round(sc.total_monthly_revenue_change).toLocaleString() + '</div></div>';
      scHtml += '<div class="kpi-card" style="border-left-color:var(--accent);"><div class="kpi-label">年間売上変動</div><div class="kpi-value" style="color:var(--accent);">+&#165;' + Math.round(sc.annual_revenue_change).toLocaleString() + '</div></div>';
      scHtml += '</div>';
      if (sc.per_hotel_impacts) {
        scHtml += '<details style="margin-bottom:0.5rem;"><summary style="cursor:pointer;font-size:0.78rem;font-weight:600;color:var(--accent);">ホテル別詳細 (' + sc.per_hotel_impacts.length + '件)</summary>';
        scHtml += '<table class="data-table" style="margin-top:0.5rem;"><thead><tr><th>ホテル名</th><th>現スコア</th><th>現月間売上</th><th>売上変動額</th></tr></thead><tbody>';
        sc.per_hotel_impacts.forEach(function(h) {
          scHtml += '<tr><td>' + esc(h.name) + '</td><td>' + num(h.current_score) + '</td><td style="text-align:right;">&#165;' + h.current_revenue.toLocaleString() + '</td><td style="text-align:right;color:var(--green);font-weight:700;">+&#165;' + h.estimated_revenue_change.toLocaleString() + '</td></tr>';
        });
        scHtml += '</tbody></table></details>';
      }
    });
    html.push(section('&#128176; 品質改善による売上インパクト', scHtml));
  }

  // Benchmark
  if (d.benchmark_comparison) {
    var bm = d.benchmark_comparison;
    var bmHtml = '<div style="padding:1rem;background:#F0F9FF;border-radius:8px;">';
    bmHtml += '<div style="font-size:0.85rem;"><strong>業界ベンチマーク:</strong> ' + esc(bm.industry_benchmark || '') + '</div>';
    bmHtml += '<div style="font-size:0.85rem;margin-top:0.3rem;"><strong>自社データ:</strong> スコア0.1pt改善あたりRevPAR ' + esc(String(bm.our_data_revpar_pct_per_01)) + ' (' + esc(String(bm.our_revpar_change_per_01)) + ')</div>';
    bmHtml += '<div style="font-size:0.82rem;color:var(--text-light);margin-top:0.5rem;">' + esc(bm.interpretation || '') + '</div>';
    bmHtml += '</div>';
    html.push(section('&#128209; 業界ベンチマーク比較', bmHtml));
  }

  // Hotel improvement potentials
  if (d.hotel_improvement_potentials) {
    var impTbl = '<div style="overflow-x:auto;"><table class="data-table"><thead><tr><th>ホテル</th><th>現スコア</th><th>目標</th><th>改善幅</th><th>優先度</th><th>月間売上変動</th></tr></thead><tbody>';
    d.hotel_improvement_potentials.forEach(function(h) {
      impTbl += '<tr><td>' + esc(h.name) + '</td><td style="text-align:right;">' + num(h.current_score) + '</td><td style="text-align:right;">' + num(h.target_score) + '</td><td style="text-align:right;">+' + num(h.improvement) + '</td><td>' + priBadge(h.priority) + '</td><td style="text-align:right;color:var(--green);font-weight:700;">+&#165;' + (h.estimated_monthly_impact || 0).toLocaleString() + '</td></tr>';
    });
    impTbl += '</tbody></table></div>';
    if (d.total_improvement_potential) {
      impTbl += '<div style="margin-top:0.75rem;padding:0.75rem;background:#ECFDF5;border-radius:8px;font-size:0.85rem;text-align:center;"><strong>総改善ポテンシャル:</strong> 月間 +&#165;' + Math.round(d.total_improvement_potential.monthly).toLocaleString() + ' / 年間 +&#165;' + Math.round(d.total_improvement_potential.annual).toLocaleString() + '</div>';
    }
    html.push(section('&#127919; ホテル別改善ポテンシャル', impTbl));
  }

  html.push(renderRecs(d.recommendations));
  return html.join('\n');
}

// ====== Analysis 7: Best Practices ======
function renderA7(d) {
  if (!d) return '<p>データがありません</p>';
  var html = [];
  if (d.analysis_metadata) html.push(metaBox(d.analysis_metadata));

  // Hotel scorecards ranked table
  if (d.best_practices && d.best_practices.hotels_ranked) {
    var ranked = d.best_practices.hotels_ranked;
    var tbl = '<div style="overflow-x:auto;"><table class="data-table"><thead><tr><th>#</th><th>ホテル</th><th>総合</th><th>口コミ</th><th>クレーム</th><th>完了時間</th><th>安全</th></tr></thead><tbody>';
    ranked.forEach(function(h, i) {
      var bg = i < 5 ? ' style="background:#ECFDF5;"' : (i >= ranked.length - 5 ? ' style="background:#FEF2F2;"' : '');
      tbl += '<tr' + bg + '><td>' + (i + 1) + '</td><td style="font-weight:600;">' + esc(h.name) + '</td><td style="text-align:right;font-weight:800;color:var(--accent);">' + num(h.composite) + '</td><td style="text-align:right;">' + num(h.review) + '</td><td style="text-align:right;">' + num(h.claims) + '</td><td style="text-align:right;">' + num(h.time) + '</td><td style="text-align:right;">' + num(h.safety) + '</td></tr>';
    });
    tbl += '</tbody></table></div>';
    if (d.best_practices.top5_avg_score != null) {
      tbl += '<div style="font-size:0.78rem;color:var(--text-light);margin-top:0.5rem;">Top5平均: ' + num(d.best_practices.top5_avg_score) + ' / Bottom5平均: ' + num(d.best_practices.bottom5_avg_score) + '</div>';
    }
    html.push(section('&#127942; 総合ランキング', tbl));
  }

  // Top 5 / Bottom 5
  if (d.best_practices) {
    var bp = d.best_practices;
    var listHtml = '<div class="grid-2">';
    if (bp.top5_best_practice) {
      listHtml += '<div style="padding:1rem;background:#ECFDF5;border-radius:8px;"><h4 style="font-size:0.85rem;color:var(--green);margin-bottom:0.5rem;">&#127775; ベストプラクティスTop5</h4><ol style="font-size:0.82rem;padding-left:1.5rem;">';
      bp.top5_best_practice.forEach(function(h) { listHtml += '<li style="margin-bottom:0.3rem;">' + esc(h) + '</li>'; });
      listHtml += '</ol></div>';
    }
    if (bp.bottom5_improvement) {
      listHtml += '<div style="padding:1rem;background:#FEF2F2;border-radius:8px;"><h4 style="font-size:0.85rem;color:var(--red);margin-bottom:0.5rem;">&#128680; 改善重点Bottom5</h4><ol style="font-size:0.82rem;padding-left:1.5rem;">';
      bp.bottom5_improvement.forEach(function(h) { listHtml += '<li style="margin-bottom:0.3rem;">' + esc(h) + '</li>'; });
      listHtml += '</ol></div>';
    }
    listHtml += '</div>';
    html.push(section('&#127919; Top5 vs Bottom5', listHtml));
  }

  // Differentiating factors
  if (d.best_practices && d.best_practices.differentiating_factors) {
    var factors = d.best_practices.differentiating_factors;
    var facHtml = '<div class="priority-grid">';
    var facColors = ['var(--green)', 'var(--blue)', 'var(--accent)', 'var(--orange)'];
    factors.forEach(function(f, i) {
      facHtml += '<div class="priority-card" style="background:var(--bg);border-color:' + facColors[i % facColors.length] + ';">';
      facHtml += '<div class="priority-title" style="color:' + facColors[i % facColors.length] + ';">' + esc(f.factor) + '</div>';
      facHtml += '<div style="font-size:0.78rem;color:var(--text-light);margin-bottom:0.5rem;">' + esc(f.description || '') + '</div>';
      if (f.hotels && f.hotels.length) {
        facHtml += '<div style="font-size:0.75rem;margin-bottom:0.5rem;"><strong>該当ホテル:</strong> ' + f.hotels.map(esc).join('、') + '</div>';
      }
      if (f.transferable_actions && f.transferable_actions.length) {
        facHtml += '<div style="font-size:0.75rem;"><strong>横展開アクション:</strong></div><ul style="font-size:0.75rem;padding-left:1rem;margin-top:0.25rem;">';
        f.transferable_actions.forEach(function(a) { facHtml += '<li>' + esc(a) + '</li>'; });
        facHtml += '</ul>';
      }
      facHtml += '</div>';
    });
    facHtml += '</div>';
    html.push(section('&#128161; 差別化要因', facHtml));
  }

  // Hotel scorecards
  if (d.hotel_scorecards) {
    var scTbl = '<div style="overflow-x:auto;"><table class="data-table"><thead><tr><th>ホテル</th><th>クレーム率</th><th>クレーム数</th><th>メイド</th><th>チェッカー</th><th>M/C比</th><th>ティア</th><th>総合スコア</th></tr></thead><tbody>';
    d.hotel_scorecards.forEach(function(h) {
      scTbl += '<tr><td>' + esc(h.name) + '</td><td style="text-align:right;">' + pct(h.claim_rate) + '</td><td style="text-align:right;">' + h.total_claims + '</td><td style="text-align:right;">' + num(h.avg_maids) + '</td><td style="text-align:right;">' + num(h.avg_checkers) + '</td><td style="text-align:right;">' + num(h.maid_checker_ratio) + '</td><td>' + tierBadge(h.tier || '') + '</td><td style="text-align:right;font-weight:700;">' + num(h.composite_score) + '</td></tr>';
    });
    scTbl += '</tbody></table></div>';
    html.push(section('&#128203; ホテル別スコアカード', scTbl));
  }

  // Implementation roadmap
  if (d.implementation_roadmap && d.implementation_roadmap.phases) {
    var phases = d.implementation_roadmap.phases;
    var phaseColors = ['var(--red)', 'var(--orange)', 'var(--blue)', 'var(--green)'];
    var roadHtml = '';
    phases.forEach(function(p, i) {
      roadHtml += '<div class="phase">';
      roadHtml += '<div class="phase-header"><div class="phase-num" style="background:' + phaseColors[i] + ';">' + (i + 1) + '</div>';
      roadHtml += '<div><div class="phase-title">' + esc(p.title || p.phase) + '</div></div></div>';
      if (p.actions && p.actions.length) {
        roadHtml += '<ul class="action-list">';
        p.actions.forEach(function(a) { roadHtml += '<li>' + esc(a) + '</li>'; });
        roadHtml += '</ul>';
      }
      if (p.expected_impact) {
        roadHtml += '<div style="font-size:0.75rem;color:var(--text-light);padding-left:2.5rem;margin-top:0.3rem;">期待効果: ' + esc(p.expected_impact) + '</div>';
      }
      roadHtml += '</div>';
    });
    html.push(section('&#128640; 実装ロードマップ', roadHtml));
  }

  // Cross-analysis insights
  if (d.cross_analysis_insights) {
    var insHtml = '';
    d.cross_analysis_insights.forEach(function(ins) {
      insHtml += '<div style="padding:1rem;background:var(--bg);border-radius:8px;margin-bottom:0.75rem;border-left:4px solid var(--accent);">';
      insHtml += '<div style="font-weight:700;font-size:0.85rem;margin-bottom:0.3rem;">' + esc(ins.title) + '</div>';
      insHtml += '<div style="font-size:0.82rem;margin-bottom:0.3rem;"><strong>発見:</strong> ' + esc(ins.finding) + '</div>';
      insHtml += '<div style="font-size:0.82rem;color:var(--text-light);"><strong>示唆:</strong> ' + esc(ins.implication) + '</div>';
      insHtml += '</div>';
    });
    html.push(section('&#128270; クロス分析インサイト', insHtml));
  }

  html.push(renderRecs(d.recommendations));
  return html.join('\n');
}

module.exports = { renderA1, renderA2, renderA3, renderA4, renderA5, renderA6, renderA7 };
