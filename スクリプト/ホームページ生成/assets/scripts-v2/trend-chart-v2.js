// Trend Chart V2 - PRIMECHANGE Portal
// Added: target line (dashed) at 8.89 score
(function() {
  'use strict';

  var COLORS = {
    bar: '#3B82F6',
    barHover: '#2563EB',
    line: '#EF4444',
    lineMA: '#10B981',
    grid: '#E2E8F0',
    text: '#64748B',
    bg: '#F8FAFC',
    snapshotMarker: '#F59E0B',
    targetLine: '#8B5CF6'
  };

  var snapshotDates = {};
  var TARGET_SCORE = 8.89; // Portfolio target

  function createTrendChart(container, dailyStats, options) {
    options = options || {};
    var width = options.width || 800;
    var height = options.height || 280;
    var padTop = 30, padRight = 50, padBottom = 50, padLeft = 50;
    var chartW = width - padLeft - padRight;
    var chartH = height - padTop - padBottom;
    var targetScore = options.targetScore || TARGET_SCORE;

    if (!dailyStats || dailyStats.length === 0) {
      if (typeof container === 'string') container = document.getElementById(container);
      if (container) container.innerHTML = '<div style="text-align:center;padding:2rem;color:#64748B;font-size:0.85rem;">データがありません</div>';
      return;
    }

    var maxCount = Math.max.apply(null, dailyStats.map(function(d) { return d.count; }));
    maxCount = Math.max(maxCount, 1);
    var maxScore = 10;
    var n = dailyStats.length;

    var dateToIdx = {};
    dailyStats.forEach(function(d, idx) { dateToIdx[d.date] = idx; });

    var barW = Math.max(2, Math.min(20, (chartW / n) * 0.7));
    var gap = chartW / n;

    var svg = [];
    svg.push('<svg viewBox="0 0 ' + width + ' ' + height + '" class="trend-chart-svg" role="img" aria-label="日別トレンドグラフ">');
    svg.push('<rect x="0" y="0" width="' + width + '" height="' + height + '" fill="white" rx="8"/>');

    // Grid lines
    var gridLines = 5;
    for (var i = 0; i <= gridLines; i++) {
      var y = padTop + (chartH / gridLines) * i;
      var val = Math.round(maxCount * (1 - i / gridLines));
      svg.push('<line x1="' + padLeft + '" y1="' + y + '" x2="' + (width - padRight) + '" y2="' + y + '" stroke="' + COLORS.grid + '" stroke-dasharray="3,3"/>');
      svg.push('<text x="' + (padLeft - 8) + '" y="' + (y + 4) + '" text-anchor="end" font-size="10" fill="' + COLORS.text + '">' + val + '</text>');
    }

    // Y axis right (score)
    for (var j = 0; j <= gridLines; j++) {
      var yR = padTop + (chartH / gridLines) * j;
      var scoreVal = Math.round(maxScore * (1 - j / gridLines) * 10) / 10;
      svg.push('<text x="' + (width - padRight + 8) + '" y="' + (yR + 4) + '" text-anchor="start" font-size="10" fill="' + COLORS.lineMA + '">' + scoreVal + '</text>');
    }

    // Target line (dashed purple)
    var targetY = padTop + chartH - (targetScore / maxScore) * chartH;
    svg.push('<line x1="' + padLeft + '" y1="' + targetY + '" x2="' + (width - padRight) + '" y2="' + targetY + '" stroke="' + COLORS.targetLine + '" stroke-width="1.5" stroke-dasharray="6,3" opacity="0.7"/>');
    svg.push('<text x="' + (width - padRight + 4) + '" y="' + (targetY - 4) + '" text-anchor="start" font-size="8" fill="' + COLORS.targetLine + '" font-weight="600">目標 ' + targetScore + '</text>');

    // Snapshot markers
    if (options.showSnapshotMarkers !== false) {
      Object.keys(snapshotDates).forEach(function(date) {
        var idx = dateToIdx[date];
        if (idx === undefined) return;
        var snap = snapshotDates[date];
        var x = padLeft + gap * idx + gap / 2;
        svg.push('<line x1="' + x + '" y1="' + padTop + '" x2="' + x + '" y2="' + (padTop + chartH) + '" stroke="' + COLORS.snapshotMarker + '" stroke-width="1.5" stroke-dasharray="4,3" opacity="0.6"/>');
        svg.push('<circle cx="' + x + '" cy="' + (padTop - 2) + '" r="4" fill="' + COLORS.snapshotMarker + '">');
        svg.push('<title>スナップショット: ' + date + '</title>');
        svg.push('</circle>');
      });
    }

    // Bars
    dailyStats.forEach(function(d, idx) {
      var x = padLeft + gap * idx + (gap - barW) / 2;
      var barH = (d.count / maxCount) * chartH;
      var barY = padTop + chartH - barH;
      svg.push('<rect x="' + x + '" y="' + barY + '" width="' + barW + '" height="' + barH + '" fill="' + COLORS.bar + '" opacity="0.7" rx="1">');
      svg.push('<title>' + d.date + '\n件数: ' + d.count + '\n平均: ' + d.avg + '</title>');
      svg.push('</rect>');
    });

    // MA line
    var linePoints = [];
    dailyStats.forEach(function(d, idx) {
      var x = padLeft + gap * idx + gap / 2;
      var lineY = padTop + chartH - (d.ma / maxScore) * chartH;
      linePoints.push(x + ',' + lineY);
    });

    if (linePoints.length > 1) {
      svg.push('<polyline points="' + linePoints.join(' ') + '" fill="none" stroke="' + COLORS.lineMA + '" stroke-width="2" stroke-linejoin="round" stroke-linecap="round"/>');
      dailyStats.forEach(function(d, idx) {
        var x = padLeft + gap * idx + gap / 2;
        var lineY = padTop + chartH - (d.ma / maxScore) * chartH;
        svg.push('<circle cx="' + x + '" cy="' + lineY + '" r="3" fill="' + COLORS.lineMA + '" stroke="white" stroke-width="1">');
        svg.push('<title>' + d.date + '\n7日移動平均: ' + d.ma + '</title>');
        svg.push('</circle>');
      });
    }

    // X axis labels
    var labelInterval = Math.max(1, Math.ceil(n / 12));
    dailyStats.forEach(function(d, idx) {
      if (idx % labelInterval !== 0 && idx !== n - 1) return;
      var x = padLeft + gap * idx + gap / 2;
      var dateLabel = d.date.slice(5);
      svg.push('<text x="' + x + '" y="' + (height - padBottom + 18) + '" text-anchor="middle" font-size="9" fill="' + COLORS.text + '" transform="rotate(-30,' + x + ',' + (height - padBottom + 18) + ')">' + dateLabel + '</text>');
    });

    // Axis labels
    svg.push('<text x="' + (padLeft - 5) + '" y="' + (padTop - 10) + '" text-anchor="start" font-size="10" font-weight="600" fill="' + COLORS.bar + '">件数</text>');
    svg.push('<text x="' + (width - padRight + 5) + '" y="' + (padTop - 10) + '" text-anchor="start" font-size="10" font-weight="600" fill="' + COLORS.lineMA + '">平均点(MA)</text>');

    // Legend
    var legendX = padLeft + 10;
    var legendY = padTop + 10;
    svg.push('<rect x="' + legendX + '" y="' + legendY + '" width="10" height="10" fill="' + COLORS.bar + '" opacity="0.7" rx="1"/>');
    svg.push('<text x="' + (legendX + 14) + '" y="' + (legendY + 9) + '" font-size="9" fill="' + COLORS.text + '">口コミ件数</text>');
    svg.push('<line x1="' + (legendX + 80) + '" y1="' + (legendY + 5) + '" x2="' + (legendX + 95) + '" y2="' + (legendY + 5) + '" stroke="' + COLORS.lineMA + '" stroke-width="2"/>');
    svg.push('<text x="' + (legendX + 99) + '" y="' + (legendY + 9) + '" font-size="9" fill="' + COLORS.text + '">7日移動平均</text>');
    svg.push('<line x1="' + (legendX + 170) + '" y1="' + (legendY + 5) + '" x2="' + (legendX + 185) + '" y2="' + (legendY + 5) + '" stroke="' + COLORS.targetLine + '" stroke-width="1.5" stroke-dasharray="4,2"/>');
    svg.push('<text x="' + (legendX + 189) + '" y="' + (legendY + 9) + '" font-size="9" fill="' + COLORS.text + '">目標ライン</text>');

    if (Object.keys(snapshotDates).length > 0) {
      var snapLegendX = legendX + 260;
      svg.push('<line x1="' + snapLegendX + '" y1="' + legendY + '" x2="' + snapLegendX + '" y2="' + (legendY + 10) + '" stroke="' + COLORS.snapshotMarker + '" stroke-width="1.5" stroke-dasharray="2,2"/>');
      svg.push('<circle cx="' + snapLegendX + '" cy="' + legendY + '" r="3" fill="' + COLORS.snapshotMarker + '"/>');
      svg.push('<text x="' + (snapLegendX + 6) + '" y="' + (legendY + 9) + '" font-size="9" fill="' + COLORS.text + '">スナップショット</text>');
    }

    svg.push('</svg>');

    if (typeof container === 'string') container = document.getElementById(container);
    if (container) container.innerHTML = svg.join('\n');
  }

  // Auto-init
  document.addEventListener('dateFilterChanged', function(e) {
    var detail = e.detail;

    if (detail.snapshots) {
      snapshotDates = {};
      detail.snapshots.forEach(function(snap) {
        var markerDate = (snap.data_range && snap.data_range.max) || snap.date;
        snapshotDates[markerDate] = snap;
      });
    }

    var portfolioTrend = document.getElementById('portfolioTrend');
    if (portfolioTrend) {
      createTrendChart(portfolioTrend, detail.dailyStats);
    }

    var modalTrend = document.getElementById('modalTrend');
    if (modalTrend && modalTrend.dataset.hotelKey) {
      var hotelKey = modalTrend.dataset.hotelKey;
      var hotelDaily = detail.hotelDaily[hotelKey] || [];
      createTrendChart(modalTrend, hotelDaily);
    }

    var cleaningTrend = document.getElementById('cleaningTrend');
    if (cleaningTrend && detail.filteredReviews) {
      var cleaningDaily = calcCleaningDaily(detail.filteredReviews, detail.cleaningKeywords);
      createTrendChart(cleaningTrend, cleaningDaily);
    }
  });

  function calcCleaningDaily(filteredReviews, keywords) {
    var dayMap = {};
    Object.keys(filteredReviews).forEach(function(key) {
      filteredReviews[key].forEach(function(r) {
        if (!r.d) return;
        if (!dayMap[r.d]) dayMap[r.d] = { count: 0, total: 0 };
        dayMap[r.d].total++;
        var text = (r.c || '') + (r.g || '') + (r.b || '');
        for (var i = 0; i < keywords.length; i++) {
          if (text.indexOf(keywords[i]) !== -1) {
            dayMap[r.d].count++;
            break;
          }
        }
      });
    });

    var days = Object.keys(dayMap).sort();
    return days.map(function(d) {
      return {
        date: d,
        count: dayMap[d].count,
        avg: dayMap[d].total > 0 ? Math.round(dayMap[d].count / dayMap[d].total * 100 * 10) / 10 : 0,
        ma: 0
      };
    });
  }

  window.TrendChart = {
    create: createTrendChart,
    calcCleaningDaily: calcCleaningDaily
  };
})();
