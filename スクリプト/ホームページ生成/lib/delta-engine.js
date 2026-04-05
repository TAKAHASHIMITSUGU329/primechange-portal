// V2 Delta Engine - snapshot diff calculation
const fs = require('fs');
const path = require('path');

function calcDeltas(outputDir, currentSummary, hotelDetailsProcessed, cleanDive) {
  var deltas = { hasDeltas: false, alerts: [], metrics: {}, hotels: {}, cleaning: { categories: {} } };

  // Load snapshot index to find previous snapshot
  var snapshotIndexPath = path.join(outputDir, 'data', 'snapshot-index.json');
  if (!fs.existsSync(snapshotIndexPath)) return deltas;

  var snapshots;
  try { snapshots = JSON.parse(fs.readFileSync(snapshotIndexPath, 'utf8')); }
  catch(e) { return deltas; }

  if (!snapshots || snapshots.length < 1) return deltas;

  // Previous snapshot is the last one in the index
  var prev = snapshots[snapshots.length - 1];
  if (!prev) return deltas;

  // Load previous portfolio summary
  var prevSummaryPath = path.join(outputDir, 'data', 'snapshots', prev.id, 'portfolio-summary.json');
  var prevSummary;
  try { prevSummary = JSON.parse(fs.readFileSync(prevSummaryPath, 'utf8')); }
  catch(e) { return deltas; }

  // Load previous hotel-details.json (may not exist in old snapshots)
  var prevHotelDetailsPath = path.join(outputDir, 'data', 'snapshots', prev.id, 'hotel-details.json');
  var prevHotelDetails = null;
  try { prevHotelDetails = JSON.parse(fs.readFileSync(prevHotelDetailsPath, 'utf8')); }
  catch(e) { /* old snapshot without hotel-details — handled gracefully */ }

  // Load previous cleaning-summary.json (may not exist in old snapshots)
  var prevCleaningPath = path.join(outputDir, 'data', 'snapshots', prev.id, 'cleaning-summary.json');
  var prevCleaning = null;
  try { prevCleaning = JSON.parse(fs.readFileSync(prevCleaningPath, 'utf8')); }
  catch(e) { /* old snapshot without cleaning-summary — handled gracefully */ }

  deltas.hasDeltas = true;
  deltas.previousDate = prev.id;

  // --- Portfolio-level metric deltas ---
  var metrics = {};
  var fields = ['avg_score', 'high_rate', 'low_rate', 'cleaning_issue_rate', 'total_reviews', 'total_hotels', 'cleaning_issue_count'];
  fields.forEach(function(f) {
    var curr = currentSummary[f];
    var prevVal = prevSummary[f];
    if (curr != null && prevVal != null) {
      metrics[f] = {
        current: curr,
        previous: prevVal,
        delta: Math.round((curr - prevVal) * 100) / 100
      };
    }
  });
  deltas.metrics = metrics;

  // --- Per-hotel deltas ---
  if (hotelDetailsProcessed && prevHotelDetails && typeof hotelDetailsProcessed === 'object' && typeof prevHotelDetails === 'object') {
    // Both are objects keyed by hotel key (e.g. "daiwa_osaki")
    Object.keys(hotelDetailsProcessed).forEach(function(key) {
      var hotel = hotelDetailsProcessed[key];
      var prevH = prevHotelDetails[key];
      if (!hotel || !prevH) return;
      var id = key;

      var hotelDelta = {};

      // overall_avg_10pt
      if (hotel.overall_avg_10pt != null && prevH.overall_avg_10pt != null) {
        hotelDelta.overall_avg_10pt = {
          current: hotel.overall_avg_10pt,
          previous: prevH.overall_avg_10pt,
          delta: Math.round((hotel.overall_avg_10pt - prevH.overall_avg_10pt) * 100) / 100
        };
      }

      // high_rate
      if (hotel.high_rate != null && prevH.high_rate != null) {
        hotelDelta.high_rate = {
          current: hotel.high_rate,
          previous: prevH.high_rate,
          delta: Math.round((hotel.high_rate - prevH.high_rate) * 100) / 100
        };
      }

      // total_reviews
      if (hotel.total_reviews != null && prevH.total_reviews != null) {
        hotelDelta.total_reviews = {
          current: hotel.total_reviews,
          previous: prevH.total_reviews,
          delta: hotel.total_reviews - prevH.total_reviews
        };
      }

      // low_rate
      if (hotel.low_rate != null && prevH.low_rate != null) {
        hotelDelta.low_rate = {
          current: hotel.low_rate,
          previous: prevH.low_rate,
          delta: Math.round((hotel.low_rate - prevH.low_rate) * 100) / 100
        };
      }

      // low_count
      if (hotel.low_count != null && prevH.low_count != null) {
        hotelDelta.low_count = {
          current: hotel.low_count,
          previous: prevH.low_count,
          delta: hotel.low_count - prevH.low_count
        };
      }

      // high_count
      if (hotel.high_count != null && prevH.high_count != null) {
        hotelDelta.high_count = {
          current: hotel.high_count,
          previous: prevH.high_count,
          delta: hotel.high_count - prevH.high_count
        };
      }

      // Site-by-site avg_10pt diffs
      var sites = {};
      var currentSites = hotel.site_stats || hotel.sites || hotel.site_details || [];
      var prevSites = prevH.site_stats || prevH.sites || prevH.site_details || [];

      // Normalize arrays with 'site' or 'site_name' key
      var currSiteMap = {};
      var prevSiteMap = {};

      if (Array.isArray(currentSites)) {
        currentSites.forEach(function(s) { if (s) currSiteMap[s.site || s.site_name || ''] = s; });
      } else if (typeof currentSites === 'object') {
        currSiteMap = currentSites;
      }

      if (Array.isArray(prevSites)) {
        prevSites.forEach(function(s) { if (s) prevSiteMap[s.site || s.site_name || ''] = s; });
      } else if (typeof prevSites === 'object') {
        prevSiteMap = prevSites;
      }

      Object.keys(currSiteMap).forEach(function(siteName) {
        var currSite = currSiteMap[siteName];
        var prevSite = prevSiteMap[siteName];
        if (!currSite || !prevSite) return;

        var currAvg = currSite.avg_10pt != null ? currSite.avg_10pt : currSite.avg_score;
        var prevAvg = prevSite.avg_10pt != null ? prevSite.avg_10pt : prevSite.avg_score;
        if (currAvg != null && prevAvg != null) {
          sites[siteName] = {
            avg_10pt: {
              current: currAvg,
              previous: prevAvg,
              delta: Math.round((currAvg - prevAvg) * 100) / 100
            }
          };
        }

        var currCount = currSite.count;
        var prevCount = prevSite.count;
        if (currCount != null && prevCount != null) {
          if (!sites[siteName]) sites[siteName] = {};
          sites[siteName].count = {
            current: currCount,
            previous: prevCount,
            delta: currCount - prevCount
          };
        }
      });

      if (Object.keys(sites).length > 0) {
        hotelDelta.sites = sites;
      }

      if (Object.keys(hotelDelta).length > 0) {
        deltas.hotels[id] = hotelDelta;
      }
    });
  }

  // --- Cleaning category deltas ---
  if (cleanDive && prevCleaning) {
    var currentCategories = cleanDive.category_summary || cleanDive.categories || cleanDive.category_details || {};
    var prevCategories = prevCleaning.category_summary || prevCleaning.categories || prevCleaning.category_details || {};

    // Normalize: can be array or object
    var currCatMap = {};
    var prevCatMap = {};

    function extractCount(v) {
      if (typeof v === 'number') return v;
      if (v && v.total_mentions != null) return v.total_mentions;
      if (v && v.count != null) return v.count;
      if (v && v.mention_count != null) return v.mention_count;
      return 0;
    }

    if (Array.isArray(currentCategories)) {
      currentCategories.forEach(function(c) {
        if (c && c.category) currCatMap[c.category] = extractCount(c);
      });
    } else {
      Object.keys(currentCategories).forEach(function(k) {
        currCatMap[k] = extractCount(currentCategories[k]);
      });
    }

    if (Array.isArray(prevCategories)) {
      prevCategories.forEach(function(c) {
        if (c && c.category) prevCatMap[c.category] = extractCount(c);
      });
    } else {
      Object.keys(prevCategories).forEach(function(k) {
        prevCatMap[k] = extractCount(prevCategories[k]);
      });
    }

    var cleaningCategories = {};
    Object.keys(currCatMap).forEach(function(cat) {
      var curr = currCatMap[cat];
      var prevVal = prevCatMap[cat];
      if (curr != null && prevVal != null) {
        cleaningCategories[cat] = {
          current: curr,
          previous: prevVal,
          delta: curr - prevVal
        };
      } else if (curr != null) {
        // New category not in previous snapshot
        cleaningCategories[cat] = {
          current: curr,
          previous: 0,
          delta: curr
        };
      }
    });

    if (Object.keys(cleaningCategories).length > 0) {
      deltas.cleaning.categories = cleaningCategories;
    }
  }

  // --- Generate alerts based on thresholds ---
  var alerts = [];
  if (metrics.avg_score) {
    if (metrics.avg_score.delta <= -0.1) {
      alerts.push({
        type: 'danger',
        icon: '&#9888;&#65039;',
        title: 'スコア低下警告',
        message: 'ポートフォリオ平均スコアが ' + metrics.avg_score.delta.toFixed(2) + 'pt 低下（' + metrics.avg_score.previous + ' → ' + metrics.avg_score.current + '）',
        severity: 'red'
      });
    } else if (metrics.avg_score.delta >= 0.05) {
      alerts.push({
        type: 'improvement',
        icon: '&#128994;',
        title: 'スコア改善',
        message: 'ポートフォリオ平均スコアが +' + metrics.avg_score.delta.toFixed(2) + 'pt 改善（' + metrics.avg_score.previous + ' → ' + metrics.avg_score.current + '）',
        severity: 'green'
      });
    }
  }

  if (metrics.cleaning_issue_rate) {
    if (metrics.cleaning_issue_rate.delta >= 2.0) {
      alerts.push({
        type: 'danger',
        icon: '&#128680;',
        title: '清掃クレーム率上昇',
        message: '清掃クレーム率が +' + metrics.cleaning_issue_rate.delta.toFixed(1) + '%上昇（' + metrics.cleaning_issue_rate.previous + '% → ' + metrics.cleaning_issue_rate.current + '%）',
        severity: 'red'
      });
    }
  }

  if (metrics.low_rate) {
    if (metrics.low_rate.delta >= 1.0) {
      alerts.push({
        type: 'danger',
        icon: '&#9888;&#65039;',
        title: '低評価率上昇',
        message: '低評価率が +' + metrics.low_rate.delta.toFixed(1) + '%上昇（' + metrics.low_rate.previous + '% → ' + metrics.low_rate.current + '%）',
        severity: 'red'
      });
    } else if (metrics.low_rate.delta <= -1.0) {
      alerts.push({
        type: 'improvement',
        icon: '&#128994;',
        title: '低評価率改善',
        message: '低評価率が ' + metrics.low_rate.delta.toFixed(1) + '%改善（' + metrics.low_rate.previous + '% → ' + metrics.low_rate.current + '%）',
        severity: 'green'
      });
    }
  }

  if (metrics.total_reviews) {
    var reviewDelta = metrics.total_reviews.delta;
    if (reviewDelta > 0) {
      alerts.push({
        type: 'info',
        icon: '&#128172;',
        title: '新規口コミ',
        message: '+' + reviewDelta + '件の新規口コミを取得',
        severity: 'blue'
      });
    }
  }

  deltas.alerts = alerts;
  return deltas;
}

module.exports = { calcDeltas };
