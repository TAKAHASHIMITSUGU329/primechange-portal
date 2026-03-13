// Date Filter Engine - PRIMECHANGE Portal
// Loads all reviews, filters by date range, recalculates KPIs
(function() {
  'use strict';

  var allReviews = null;
  var buildMeta = null;
  var snapshotIndex = null;
  var currentRange = null; // {start, end} or null for all

  var CLEANING_KEYWORDS = ['清掃', '汚れ', 'ゴミ', '髪の毛', 'シミ', 'カビ', 'ほこり', '埃', '汚い', '不潔', '臭い', 'におい', '匂い', 'ホコリ', 'しみ', 'かび', 'ごみ'];

  var TIER_COLOR = { '優秀': '#10B981', '良好': '#3B82F6', '概ね良好': '#F59E0B', '要改善': '#EF4444' };

  function getTier(avg) {
    if (avg >= 9.0) return '優秀';
    if (avg >= 8.0) return '良好';
    if (avg >= 7.0) return '概ね良好';
    return '要改善';
  }

  function hasCleaningIssue(review) {
    var text = (review.c || '') + (review.g || '') + (review.b || '');
    for (var i = 0; i < CLEANING_KEYWORDS.length; i++) {
      if (text.indexOf(CLEANING_KEYWORDS[i]) !== -1) return true;
    }
    return false;
  }

  function filterByDate(reviews, start, end) {
    if (!start && !end) return reviews;
    return reviews.filter(function(r) {
      if (!r.d) return false;
      if (start && r.d < start) return false;
      if (end && r.d > end) return false;
      return true;
    });
  }

  function calcHotelKPI(reviews) {
    var total = reviews.length;
    if (total === 0) return {
      total_reviews: 0, overall_avg_10pt: 0,
      high_count: 0, high_rate: 0, mid_count: 0, mid_rate: 0,
      low_count: 0, low_rate: 0,
      cleaning_issue_count: 0, cleaning_issue_rate: 0,
      site_stats: [], distribution: [], tier: '要改善'
    };

    var sum = 0, high = 0, mid = 0, low = 0, cleanCount = 0;
    var siteMap = {};
    var distMap = {};

    reviews.forEach(function(r) {
      var score = parseFloat(r.r) || 0;
      sum += score;
      if (score >= 8) high++;
      else if (score >= 5) mid++;
      else low++;

      if (hasCleaningIssue(r)) cleanCount++;

      // Site stats
      var site = r.s || 'unknown';
      if (!siteMap[site]) siteMap[site] = { site: site, sum: 0, count: 0 };
      siteMap[site].sum += score;
      siteMap[site].count++;

      // Distribution
      var rounded = Math.round(score);
      if (!distMap[rounded]) distMap[rounded] = 0;
      distMap[rounded]++;
    });

    var avg = Math.round(sum / total * 100) / 100;
    var tier = getTier(avg);

    var siteStats = Object.keys(siteMap).map(function(k) {
      var s = siteMap[k];
      var sAvg = Math.round(s.sum / s.count * 100) / 100;
      return {
        site: s.site, count: s.count, avg_10pt: sAvg,
        judgment: getTier(sAvg)
      };
    }).sort(function(a, b) { return b.count - a.count; });

    var distribution = [];
    for (var sc = 10; sc >= 1; sc--) {
      distribution.push({ score: sc, count: distMap[sc] || 0 });
    }

    return {
      total_reviews: total,
      overall_avg_10pt: avg,
      high_count: high, high_rate: Math.round(high / total * 1000) / 10,
      mid_count: mid, mid_rate: Math.round(mid / total * 1000) / 10,
      low_count: low, low_rate: Math.round(low / total * 1000) / 10,
      cleaning_issue_count: cleanCount,
      cleaning_issue_rate: Math.round(cleanCount / total * 1000) / 10,
      site_stats: siteStats,
      distribution: distribution,
      tier: tier
    };
  }

  function calcPortfolioKPI(allFiltered) {
    var hotelKeys = Object.keys(allFiltered);
    var totalReviews = 0, totalSum = 0, totalHigh = 0, totalMid = 0, totalLow = 0, totalClean = 0;
    var hotelKPIs = {};

    hotelKeys.forEach(function(key) {
      var reviews = allFiltered[key];
      var kpi = calcHotelKPI(reviews);
      hotelKPIs[key] = kpi;
      totalReviews += kpi.total_reviews;
      totalSum += kpi.overall_avg_10pt * kpi.total_reviews;
      totalHigh += kpi.high_count;
      totalMid += kpi.mid_count;
      totalLow += kpi.low_count;
      totalClean += kpi.cleaning_issue_count;
    });

    var avgScore = totalReviews > 0 ? Math.round(totalSum / totalReviews * 100) / 100 : 0;

    return {
      total_hotels: hotelKeys.length,
      total_reviews: totalReviews,
      avg_score: avgScore,
      high_rate: totalReviews > 0 ? Math.round(totalHigh / totalReviews * 1000) / 10 : 0,
      mid_rate: totalReviews > 0 ? Math.round(totalMid / totalReviews * 1000) / 10 : 0,
      low_rate: totalReviews > 0 ? Math.round(totalLow / totalReviews * 1000) / 10 : 0,
      cleaning_issue_count: totalClean,
      cleaning_issue_rate: totalReviews > 0 ? Math.round(totalClean / totalReviews * 1000) / 10 : 0,
      hotels: hotelKPIs
    };
  }

  function getDailyStats(reviews) {
    var dayMap = {};
    reviews.forEach(function(r) {
      if (!r.d) return;
      if (!dayMap[r.d]) dayMap[r.d] = { count: 0, sum: 0 };
      dayMap[r.d].count++;
      dayMap[r.d].sum += (parseFloat(r.r) || 0);
    });

    var days = Object.keys(dayMap).sort();
    return days.map(function(d) {
      return {
        date: d,
        count: dayMap[d].count,
        avg: dayMap[d].count > 0 ? Math.round(dayMap[d].sum / dayMap[d].count * 100) / 100 : 0
      };
    });
  }

  // Moving average
  function movingAvg(dailyStats, window) {
    window = window || 7;
    return dailyStats.map(function(d, i) {
      var start = Math.max(0, i - window + 1);
      var slice = dailyStats.slice(start, i + 1);
      var sum = 0, count = 0;
      slice.forEach(function(s) { sum += s.avg * s.count; count += s.count; });
      return {
        date: d.date,
        count: d.count,
        avg: d.avg,
        ma: count > 0 ? Math.round(sum / count * 100) / 100 : 0
      };
    });
  }

  function applyFilter(start, end) {
    if (!allReviews) return;
    currentRange = (start || end) ? { start: start, end: end } : null;

    // Group reviews by hotel, filter by date
    var filtered = {};
    var allFiltered = [];
    Object.keys(allReviews).forEach(function(key) {
      var reviews = filterByDate(allReviews[key], start, end);
      filtered[key] = reviews;
      allFiltered = allFiltered.concat(reviews);
    });

    var portfolio = calcPortfolioKPI(filtered);
    var daily = getDailyStats(allFiltered);
    var dailyMA = movingAvg(daily, 7);

    // Per-hotel daily stats
    var hotelDaily = {};
    Object.keys(filtered).forEach(function(key) {
      hotelDaily[key] = movingAvg(getDailyStats(filtered[key]), 7);
    });

    // Dispatch event with recalculated data
    var event = new CustomEvent('dateFilterChanged', {
      detail: {
        range: currentRange,
        filteredCount: allFiltered.length,
        portfolio: portfolio,
        dailyStats: dailyMA,
        hotelDaily: hotelDaily,
        filteredReviews: filtered,
        cleaningKeywords: CLEANING_KEYWORDS,
        tierColor: TIER_COLOR
      }
    });
    document.dispatchEvent(event);
  }

  function loadAllReviews() {
    return fetch('data/hotel-reviews-all.json')
      .then(function(r) { return r.json(); })
      .then(function(data) {
        allReviews = data;
        return data;
      })
      .catch(function(e) {
        console.warn('hotel-reviews-all.json の読み込みに失敗:', e);
        allReviews = {};
        return {};
      });
  }

  function loadBuildMeta() {
    return fetch('data/build-meta.json')
      .then(function(r) { return r.json(); })
      .then(function(data) { buildMeta = data; return data; })
      .catch(function() { return null; });
  }

  function loadSnapshotIndex() {
    return fetch('data/snapshot-index.json')
      .then(function(r) { return r.json(); })
      .then(function(data) { snapshotIndex = data; return data; })
      .catch(function() { snapshotIndex = []; return []; });
  }

  function init() {
    Promise.all([loadAllReviews(), loadBuildMeta(), loadSnapshotIndex()])
      .then(function() {
        // Fire initial event with all data (no date filter)
        applyFilter(null, null);

        // Notify that filter engine is ready
        document.dispatchEvent(new CustomEvent('dateFilterReady', {
          detail: {
            buildMeta: buildMeta,
            snapshots: snapshotIndex,
            dataRange: buildMeta ? buildMeta.data_range : null
          }
        }));
      });
  }

  document.addEventListener('DOMContentLoaded', init);

  // Expose API
  window.DateFilter = {
    apply: applyFilter,
    getReviews: function() { return allReviews; },
    getMeta: function() { return buildMeta; },
    getSnapshots: function() { return snapshotIndex; },
    getRange: function() { return currentRange; },
    calcHotelKPI: calcHotelKPI,
    getDailyStats: getDailyStats,
    movingAvg: movingAvg,
    hasCleaningIssue: hasCleaningIssue,
    CLEANING_KEYWORDS: CLEANING_KEYWORDS,
    TIER_COLOR: TIER_COLOR,
    getTier: getTier
  };
})();
