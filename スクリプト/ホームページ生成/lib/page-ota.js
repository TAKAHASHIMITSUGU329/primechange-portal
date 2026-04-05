// V3 OTA Strategy Page Generator
'use strict';
var common = require('./common-v2');
var esc = common.esc;
var nav = common.nav;
var footer = common.footer;
var pageHead = common.pageHead;
var pageFoot = common.pageFoot;
var deltaBadgeCompact = common.deltaBadgeCompact;

var TIER_COLORS = {
  '優秀': '#10B981',
  '良好': '#3B82F6',
  '概ね良好': '#F59E0B',
  '要改善': '#EF4444'
};

function cellColor(score) {
  if (score == null) return '#F1F5F9';
  if (score >= 8) return '#D1FAE5';
  if (score >= 5) return '#FEF3C7';
  return '#FEE2E2';
}

function cellTextColor(score) {
  if (score == null) return '#94A3B8';
  if (score >= 8) return '#065F46';
  if (score >= 5) return '#92400E';
  return '#991B1B';
}

function scoreColor(score) {
  if (score >= 8) return '#10B981';
  if (score >= 6.5) return '#3B82F6';
  if (score >= 5) return '#F59E0B';
  return '#EF4444';
}

function buildOTA(data, deltas) {
  var pov = data.pov || {};
  var hotelsRanked = pov.hotels_ranked || [];
  var hotelDetails = data.hotelDetails || {};

  // --- Collect all site stats across all hotels ---
  var siteAgg = {};   // site -> { totalScore, count, hotelCount, reviews }
  var hotelSites = []; // { name, key, avg, sites: { siteName: statObj } }
  var allSiteNames = {};

  // Build per-hotel site map
  hotelsRanked.forEach(function(hr) {
    var detail = hotelDetails[hr.key];
    if (!detail) return;
    var stats = detail.site_stats || [];
    var siteMap = {};
    stats.forEach(function(s) {
      allSiteNames[s.site] = true;
      siteMap[s.site] = s;

      if (!siteAgg[s.site]) {
        siteAgg[s.site] = { totalScore: 0, count: 0, hotelCount: 0, reviews: 0 };
      }
      siteAgg[s.site].totalScore += s.avg_10pt;
      siteAgg[s.site].count++;
      siteAgg[s.site].hotelCount++;
      siteAgg[s.site].reviews += s.count;
    });
    hotelSites.push({
      name: hr.name,
      key: hr.key,
      avg: hr.avg || (detail.overall_avg_10pt || 0),
      tier: hr.tier || '',
      sites: siteMap
    });
  });

  // Sort hotels by avg descending
  hotelSites.sort(function(a, b) { return b.avg - a.avg; });

  // Sorted site names list
  var siteNames = Object.keys(allSiteNames).sort(function(a, b) {
    var agg_a = siteAgg[a], agg_b = siteAgg[b];
    return (agg_b ? agg_b.reviews : 0) - (agg_a ? agg_a.reviews : 0);
  });

  // --- CSS ---
  var extraCSS = [
    '.ota-summary-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 1rem; margin-bottom: 2rem; }',
    '.ota-site-card { background: white; border-radius: 12px; padding: 1.25rem; box-shadow: 0 1px 3px rgba(0,0,0,0.08); border-top: 3px solid #C23B3A; text-align: center; }',
    '.ota-site-card .site-name { font-size: 0.95rem; font-weight: 700; color: #1A1A2E; margin-bottom: 0.5rem; }',
    '.ota-site-card .site-score { font-size: 2rem; font-weight: 800; margin: 0.25rem 0; }',
    '.ota-site-card .site-meta { font-size: 0.72rem; color: #64748B; }',
    '.ota-site-card .site-meta span { display: inline-block; margin: 0 0.3rem; }',

    '.heatmap-wrap { overflow-x: auto; margin-bottom: 2rem; }',
    '.heatmap-table { width: 100%; border-collapse: collapse; font-size: 0.8rem; }',
    '.heatmap-table th { background: #1A1A2E; color: #fff; padding: 0.6rem 0.75rem; font-weight: 600; text-align: center; white-space: nowrap; }',
    '.heatmap-table th:first-child { text-align: left; min-width: 180px; }',
    '.heatmap-table td { padding: 0.5rem 0.75rem; text-align: center; border: 1px solid #E8E0E0; font-weight: 700; font-size: 0.85rem; }',
    '.heatmap-table td:first-child { text-align: left; font-weight: 600; background: #FAFAFA; white-space: nowrap; }',
    '.heatmap-table tr:hover td { opacity: 0.9; }',
    '.heatmap-no-data { color: #94A3B8; font-weight: 400; font-size: 0.75rem; }',

    '.site-ranking-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(280px, 1fr)); gap: 1rem; margin-bottom: 2rem; }',
    '.site-ranking-card { background: white; border-radius: 12px; padding: 1.25rem; box-shadow: 0 1px 3px rgba(0,0,0,0.08); }',
    '.site-ranking-card .ranking-site-title { font-size: 0.9rem; font-weight: 700; color: #1A1A2E; margin-bottom: 0.75rem; padding-bottom: 0.5rem; border-bottom: 2px solid #C23B3A; }',
    '.ranking-list { list-style: none; padding: 0; margin: 0; }',
    '.ranking-list li { display: flex; justify-content: space-between; align-items: center; padding: 0.35rem 0; border-bottom: 1px solid #F1F5F9; font-size: 0.8rem; }',
    '.ranking-list li:last-child { border-bottom: none; }',
    '.ranking-label { font-size: 0.65rem; font-weight: 700; color: #64748B; margin-top: 0.75rem; margin-bottom: 0.25rem; }',

    '.gap-table { width: 100%; border-collapse: collapse; font-size: 0.8rem; margin-bottom: 2rem; }',
    '.gap-table th { background: #1A1A2E; color: #fff; padding: 0.6rem 0.75rem; font-weight: 600; text-align: left; }',
    '.gap-table td { padding: 0.5rem 0.75rem; border-bottom: 1px solid #E8E0E0; }',
    '.gap-table tr:nth-child(even) { background: #FAFAFA; }',
    '.gap-highlight { background: #FEF3C7 !important; }',
    '.gap-value { font-weight: 800; font-size: 0.95rem; }',

    '@media (max-width: 768px) { .ota-summary-grid { grid-template-columns: 1fr 1fr; } .site-ranking-grid { grid-template-columns: 1fr; } .heatmap-table { font-size: 0.7rem; } }',
  ].join('\n');

  var lines = [];
  lines.push(pageHead('OTA STRATEGY - PRIME CHANGE', { extraCSS: extraCSS }));
  lines.push(nav('ota'));
  lines.push('<div class="container">');

  // ========== Section Heading ==========
  lines.push('<div class="section-heading"><span class="heading-en">OTA STRATEGY</span><span class="heading-ja">OTA戦略分析</span></div>');

  // ========== 1. Site Summary Cards ==========
  lines.push('<div class="card"><div class="card-title">OTAサイト別サマリー</div>');
  lines.push('<div class="ota-summary-grid">');
  siteNames.forEach(function(site) {
    var agg = siteAgg[site];
    var avgScore = agg.count > 0 ? (agg.totalScore / agg.count) : 0;
    var color = scoreColor(avgScore);
    // Aggregate site-level delta across all hotels
    var siteScoreSum = 0, sitePrevSum = 0, siteScoreCount = 0;
    var siteCountSum = 0, sitePrevCountSum = 0, siteCountHas = false;
    if (deltas && deltas.hasDeltas && deltas.hotels) {
      Object.keys(deltas.hotels).forEach(function(hk) {
        var hd = deltas.hotels[hk];
        if (hd.sites && hd.sites[site]) {
          if (hd.sites[site].avg_10pt) {
            siteScoreSum += hd.sites[site].avg_10pt.current;
            sitePrevSum += hd.sites[site].avg_10pt.previous;
            siteScoreCount++;
          }
          if (hd.sites[site].count) {
            siteCountSum += hd.sites[site].count.current;
            sitePrevCountSum += hd.sites[site].count.previous;
            siteCountHas = true;
          }
        }
      });
    }
    var siteAvgDelta = siteScoreCount > 0 ? { current: Math.round(siteScoreSum / siteScoreCount * 100) / 100, previous: Math.round(sitePrevSum / siteScoreCount * 100) / 100, delta: Math.round((siteScoreSum / siteScoreCount - sitePrevSum / siteScoreCount) * 100) / 100 } : null;
    var siteCountDelta = siteCountHas ? { current: siteCountSum, previous: sitePrevCountSum, delta: siteCountSum - sitePrevCountSum } : null;
    lines.push('<div class="ota-site-card" style="border-top-color:' + color + ';">');
    lines.push('  <div class="site-name">' + esc(site) + '</div>');
    lines.push('  <div class="site-score" style="color:' + color + ';">' + avgScore.toFixed(2) + deltaBadgeCompact(siteAvgDelta, 'higher') + '</div>');
    lines.push('  <div class="site-meta"><span>' + agg.reviews + '件' + deltaBadgeCompact(siteCountDelta, 'higher') + '</span><span>|</span><span>' + agg.hotelCount + 'ホテル</span></div>');
    lines.push('</div>');
  });
  lines.push('</div></div>');

  // ========== 2. Site x Hotel Cross Matrix (Heatmap) ==========
  lines.push('<div class="card"><div class="card-title">OTAサイト &times; ホテル クロスマトリクス</div>');
  lines.push('<div class="heatmap-wrap"><table class="heatmap-table">');
  // Header
  lines.push('<thead><tr><th>ホテル名</th>');
  siteNames.forEach(function(site) {
    lines.push('<th>' + esc(site) + '</th>');
  });
  lines.push('<th>全体平均</th></tr></thead>');
  // Body
  lines.push('<tbody>');
  hotelSites.forEach(function(h) {
    lines.push('<tr>');
    lines.push('<td>' + esc(h.name) + '</td>');
    siteNames.forEach(function(site) {
      var stat = h.sites[site];
      if (stat) {
        var bg = cellColor(stat.avg_10pt);
        var tc = cellTextColor(stat.avg_10pt);
        var cellDelta = deltas && deltas.hasDeltas && deltas.hotels && deltas.hotels[h.key] && deltas.hotels[h.key].sites && deltas.hotels[h.key].sites[site] && deltas.hotels[h.key].sites[site].avg_10pt;
        lines.push('<td style="background:' + bg + ';color:' + tc + ';">' + stat.avg_10pt.toFixed(1) + deltaBadgeCompact(cellDelta || null, 'higher') + '</td>');
      } else {
        lines.push('<td style="background:#F1F5F9;"><span class="heatmap-no-data">&mdash;</span></td>');
      }
    });
    // Overall avg
    var avgColor = scoreColor(h.avg);
    lines.push('<td style="background:#1A1A2E;color:#fff;font-weight:800;">' + (typeof h.avg === 'number' ? h.avg.toFixed(1) : h.avg) + '</td>');
    lines.push('</tr>');
  });
  lines.push('</tbody></table></div></div>');

  // ========== 3. Site-Level Rankings ==========
  lines.push('<div class="card"><div class="card-title">OTAサイト別ランキング（TOP3 / BOTTOM3）</div>');
  lines.push('<div class="site-ranking-grid">');
  siteNames.forEach(function(site) {
    // Collect hotels with this site
    var entries = [];
    hotelSites.forEach(function(h) {
      if (h.sites[site]) {
        entries.push({ name: h.name, score: h.sites[site].avg_10pt, judgment: h.sites[site].judgment });
      }
    });
    entries.sort(function(a, b) { return b.score - a.score; });

    var top3 = entries.slice(0, 3);
    var bottom3 = entries.length > 3 ? entries.slice(-3).reverse() : [];

    lines.push('<div class="site-ranking-card">');
    lines.push('  <div class="ranking-site-title">' + esc(site) + '（' + entries.length + 'ホテル）</div>');

    // Top 3
    lines.push('  <div class="ranking-label">&#9650; TOP 3</div>');
    lines.push('  <ul class="ranking-list">');
    top3.forEach(function(e, i) {
      var jColor = TIER_COLORS[e.judgment] || '#64748B';
      lines.push('    <li><span>' + (i + 1) + '. ' + esc(e.name) + '</span><span style="color:' + jColor + ';font-weight:700;">' + e.score.toFixed(1) + '</span></li>');
    });
    lines.push('  </ul>');

    // Bottom 3
    if (bottom3.length > 0) {
      lines.push('  <div class="ranking-label">&#9660; BOTTOM 3</div>');
      lines.push('  <ul class="ranking-list">');
      bottom3.forEach(function(e, i) {
        var jColor = TIER_COLORS[e.judgment] || '#64748B';
        lines.push('    <li><span>' + (entries.length - 2 + i) + '. ' + esc(e.name) + '</span><span style="color:' + jColor + ';font-weight:700;">' + e.score.toFixed(1) + '</span></li>');
      });
      lines.push('  </ul>');
    }

    lines.push('</div>');
  });
  lines.push('</div></div>');

  // ========== 4. Strength/Weakness Gap Analysis ==========
  lines.push('<div class="card"><div class="card-title">OTAサイト間ギャップ分析（一貫性チェック）</div>');
  lines.push('<p style="font-size:0.78rem;color:#64748B;margin-bottom:1rem;">各ホテルのOTAサイト間での最高スコアと最低スコアの差を分析。ギャップが大きいほど、サイト間で評価にばらつきがあります。</p>');

  // Build gap data
  var gapData = [];
  hotelSites.forEach(function(h) {
    var siteKeys = Object.keys(h.sites);
    if (siteKeys.length < 2) return;
    var best = { site: '', score: -Infinity };
    var worst = { site: '', score: Infinity };
    siteKeys.forEach(function(site) {
      var s = h.sites[site].avg_10pt;
      if (s > best.score) { best = { site: site, score: s }; }
      if (s < worst.score) { worst = { site: site, score: s }; }
    });
    var gap = best.score - worst.score;
    gapData.push({ name: h.name, best: best, worst: worst, gap: gap, siteCount: siteKeys.length });
  });
  gapData.sort(function(a, b) { return b.gap - a.gap; });

  lines.push('<table class="gap-table">');
  lines.push('<thead><tr><th>ホテル名</th><th>最高サイト</th><th>最低サイト</th><th>ギャップ</th><th>サイト数</th></tr></thead>');
  lines.push('<tbody>');
  gapData.forEach(function(g) {
    var rowCls = g.gap > 1.0 ? ' class="gap-highlight"' : '';
    var gapColor = g.gap > 2.0 ? '#EF4444' : g.gap > 1.0 ? '#F59E0B' : '#10B981';
    lines.push('<tr' + rowCls + '>');
    lines.push('  <td style="font-weight:600;">' + esc(g.name) + '</td>');
    lines.push('  <td><span style="color:#10B981;font-weight:700;">' + g.best.score.toFixed(1) + '</span> <span style="font-size:0.72rem;color:#64748B;">(' + esc(g.best.site) + ')</span></td>');
    lines.push('  <td><span style="color:#EF4444;font-weight:700;">' + g.worst.score.toFixed(1) + '</span> <span style="font-size:0.72rem;color:#64748B;">(' + esc(g.worst.site) + ')</span></td>');
    lines.push('  <td><span class="gap-value" style="color:' + gapColor + ';">' + g.gap.toFixed(2) + '</span></td>');
    lines.push('  <td style="text-align:center;">' + g.siteCount + '</td>');
    lines.push('</tr>');
  });
  lines.push('</tbody></table></div>');

  lines.push('</div>');
  lines.push(footer());
  lines.push(pageFoot());
  return lines.join('\n');
}

module.exports = { buildOTA };
