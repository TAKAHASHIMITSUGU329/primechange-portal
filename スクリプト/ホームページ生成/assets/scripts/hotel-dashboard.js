// Hotel Dashboard - PRIMECHANGE Portal
(function() {
  'use strict';

  var hotelDetails = null;
  var hotelRanked = null;
  var tierColor = null;
  var lastFilterDetail = null;

  function init() {
    // Setup filter buttons
    document.querySelectorAll('.filter-btn').forEach(function(btn) {
      btn.addEventListener('click', function() {
        document.querySelectorAll('.filter-btn').forEach(function(b) { b.classList.remove('active'); });
        btn.classList.add('active');
        applyFilters();
      });
    });

    var searchInput = document.getElementById('searchInput');
    if (searchInput) searchInput.addEventListener('input', applyFilters);

    var sortSelect = document.getElementById('sortSelect');
    if (sortSelect) sortSelect.addEventListener('change', function() {
      sortCards();
      applyFilters();
    });
  }

  function applyFilters() {
    var activeBtn = document.querySelector('.filter-btn.active');
    var f = activeBtn ? activeBtn.dataset.filter : 'all';
    var searchInput = document.getElementById('searchInput');
    var s = searchInput ? searchInput.value.toLowerCase() : '';
    document.querySelectorAll('.hotel-card').forEach(function(c) {
      var show = (f === 'all' || c.dataset.tier === f) && (!s || c.dataset.name.toLowerCase().includes(s));
      c.style.display = show ? '' : 'none';
    });
  }

  function sortCards() {
    var sortSelect = document.getElementById('sortSelect');
    var s = sortSelect ? sortSelect.value : 'rank';
    var g = document.getElementById('hotelGrid');
    if (!g) return;
    var c = Array.from(g.children);
    c.sort(function(a, b) {
      if (s === 'rank') return +a.dataset.rank - +b.dataset.rank;
      if (s === 'reviews') return +b.dataset.reviews - +a.dataset.reviews;
      if (s === 'high_rate') return +b.dataset.highRate - +a.dataset.highRate;
      if (s === 'cleaning') return +b.dataset.cleaning - +a.dataset.cleaning;
      return 0;
    });
    c.forEach(function(el) { g.appendChild(el); });
  }

  function openModal(key) {
    var detail, info, tc;

    // Use filtered data if available, otherwise fall back to static data
    if (lastFilterDetail && lastFilterDetail.portfolio && lastFilterDetail.portfolio.hotels[key]) {
      var kpi = lastFilterDetail.portfolio.hotels[key];
      detail = kpi;
      tc = lastFilterDetail.tierColor[kpi.tier] || '#6B7280';
      // Get name from hotelRanked
      info = hotelRanked ? hotelRanked.find(function(h) { return h.key === key; }) : null;
      if (!info) info = { name: key, tier: kpi.tier, key: key };
    } else {
      if (!hotelDetails || !hotelRanked) return;
      info = hotelRanked.find(function(h) { return h.key === key; });
      detail = hotelDetails[key];
      if (!info || !detail) return;
      tc = tierColor[info.tier];
    }

    document.getElementById('modalTitle').textContent = info.name;

    var siteRows = (detail.site_stats || []).map(function(s) {
      var jc = (tierColor || lastFilterDetail.tierColor || {})[s.judgment] || '#6B7280';
      return '<tr><td><strong>' + escHtml(s.site) + '</strong></td><td>' + s.count + '件</td><td><strong>' + s.avg_10pt + '</strong></td><td><span style="color:' + jc + ';font-weight:700;">' + escHtml(s.judgment) + '</span></td></tr>';
    }).join('');

    var allScores = [10, 9, 8, 7, 6, 5, 4, 3, 2, 1];
    var distMap = {};
    (detail.distribution || []).forEach(function(d) { distMap[d.score] = d; });
    var mx = Math.max.apply(null, (detail.distribution || []).map(function(d) { return d.count; }).concat([1]));

    var distBars = allScores.map(function(sc) {
      var d = distMap[sc] || { count: 0 };
      var h = (d.count / mx * 100).toFixed(0);
      var col = sc >= 8 ? 'var(--green)' : sc >= 5 ? 'var(--orange)' : 'var(--red)';
      return '<div style="flex:1;display:flex;flex-direction:column;align-items:center;height:100%;justify-content:flex-end;"><div style="font-size:0.6rem;font-weight:600;margin-bottom:2px;">' + (d.count > 0 ? d.count : '') + '</div><div style="width:100%;min-width:20px;border-radius:4px 4px 0 0;height:' + Math.max(h, 2) + '%;background:' + col + ';"></div><div style="font-size:0.65rem;color:var(--text-light);margin-top:4px;">' + sc + '</div></div>';
    }).join('');

    // Get reviews from filtered data or static
    var comments = [];
    if (lastFilterDetail && lastFilterDetail.filteredReviews && lastFilterDetail.filteredReviews[key]) {
      comments = lastFilterDetail.filteredReviews[key].slice(0, 30).map(function(r) {
        return { site: r.s, rating_10pt: parseFloat(r.r), date: r.d, comment: r.c, good: r.g, bad: r.b };
      });
    } else if (detail.comments) {
      comments = detail.comments;
    }

    var reviewCards = comments.map(function(c) {
      var cls = c.rating_10pt >= 8 ? 'high' : c.rating_10pt >= 5 ? 'mid' : 'low';
      var sc = c.rating_10pt >= 8 ? 'var(--green)' : c.rating_10pt >= 5 ? 'var(--orange)' : 'var(--red)';
      var text = '';
      if (c.good || c.bad) {
        if (c.good) text += '<span style="color:var(--green);">&#128077; ' + escHtml(c.good) + '</span><br>';
        if (c.bad) text += '<span style="color:var(--red);">&#128078; ' + escHtml(c.bad) + '</span>';
      } else {
        text = escHtml(c.comment || '');
      }
      return '<div class="review-card ' + cls + '"><div style="display:flex;gap:0.75rem;align-items:center;margin-bottom:0.5rem;flex-wrap:wrap;"><span style="font-size:0.7rem;padding:0.15rem 0.5rem;border-radius:4px;background:var(--navy);color:white;font-weight:600;">' + escHtml(c.site) + '</span><span style="font-size:0.8rem;font-weight:700;color:' + sc + ';">' + c.rating_10pt + '点</span><span style="font-size:0.7rem;color:var(--text-light);">' + escHtml(c.date || '') + '</span></div><div style="font-size:0.8rem;line-height:1.7;">' + text + '</div></div>';
    }).join('');

    document.getElementById('modalBody').innerHTML =
      '<div style="display:flex;gap:1.5rem;flex-wrap:wrap;margin-bottom:1rem;">' +
      '<div style="text-align:center;"><div style="font-size:3rem;font-weight:800;color:' + tc + ';">' + detail.overall_avg_10pt + '</div><div style="font-size:0.8rem;color:var(--text-light);">/ 10 点</div><span class="badge" style="background:' + tc + ';margin-top:0.5rem;">' + escHtml(detail.tier || info.tier) + '</span></div>' +
      '<div style="flex:1;min-width:250px;display:flex;flex-direction:column;gap:0.75rem;">' +
      '<div><div style="display:flex;justify-content:space-between;font-size:0.75rem;margin-bottom:4px;"><span style="font-weight:600;color:var(--green);">高評価 (8-10)</span><span style="font-weight:700;">' + detail.high_rate + '% (' + detail.high_count + '件)</span></div><div style="height:10px;background:#E2E8F0;border-radius:5px;overflow:hidden;"><div style="height:100%;border-radius:5px;width:' + detail.high_rate + '%;background:var(--green);"></div></div></div>' +
      '<div><div style="display:flex;justify-content:space-between;font-size:0.75rem;margin-bottom:4px;"><span style="font-weight:600;color:var(--orange);">中評価 (5-7)</span><span style="font-weight:700;">' + detail.mid_rate + '% (' + detail.mid_count + '件)</span></div><div style="height:10px;background:#E2E8F0;border-radius:5px;overflow:hidden;"><div style="height:100%;border-radius:5px;width:' + detail.mid_rate + '%;background:var(--orange);"></div></div></div>' +
      '<div><div style="display:flex;justify-content:space-between;font-size:0.75rem;margin-bottom:4px;"><span style="font-weight:600;color:var(--red);">低評価 (1-4)</span><span style="font-weight:700;">' + detail.low_rate + '% (' + detail.low_count + '件)</span></div><div style="height:10px;background:#E2E8F0;border-radius:5px;overflow:hidden;"><div style="height:100%;border-radius:5px;width:' + detail.low_rate + '%;background:var(--red);"></div></div></div>' +
      '</div></div>' +
      '<div style="font-size:0.9rem;font-weight:700;color:var(--navy);margin:1.5rem 0 0.75rem;padding-bottom:0.5rem;border-bottom:2px solid var(--accent);">&#128202; スコア分布</div>' +
      '<div style="display:flex;align-items:flex-end;gap:4px;height:120px;padding:0.5rem 0;">' + distBars + '</div>' +
      '<div style="font-size:0.9rem;font-weight:700;color:var(--navy);margin:1.5rem 0 0.75rem;padding-bottom:0.5rem;border-bottom:2px solid var(--accent);">&#127760; サイト別評価</div>' +
      '<table class="data-table"><thead><tr><th>サイト</th><th>件数</th><th>10点換算</th><th>判定</th></tr></thead><tbody>' + siteRows + '</tbody></table>' +
      '<div style="font-size:0.9rem;font-weight:700;color:var(--navy);margin:1.5rem 0 0.75rem;padding-bottom:0.5rem;border-bottom:2px solid var(--accent);">&#128172; 口コミ一覧 (' + comments.length + '件)</div>' +
      '<div class="review-list">' + reviewCards + '</div>';

    // Set hotel key on modal trend container
    var modalTrend = document.getElementById('modalTrend');
    if (modalTrend) {
      modalTrend.dataset.hotelKey = key;
      // Render trend for this hotel
      if (lastFilterDetail && lastFilterDetail.hotelDaily && lastFilterDetail.hotelDaily[key]) {
        window.TrendChart.create(modalTrend, lastFilterDetail.hotelDaily[key]);
      }
    }

    document.getElementById('modalOverlay').classList.add('active');
    document.body.style.overflow = 'hidden';
  }

  // Update cards with filtered KPI data
  function updateCardsFromFilter(detail) {
    lastFilterDetail = detail;
    var portfolio = detail.portfolio;

    // Update dashboard KPI cards
    updateKPI('dashKpiGrid', 'total_hotels', portfolio.total_hotels);
    updateKPI('dashKpiGrid', 'total_reviews', portfolio.total_reviews.toLocaleString());
    updateKPI('dashKpiGrid', 'avg_score', portfolio.avg_score);
    updateKPI('dashKpiGrid', 'high_rate', portfolio.high_rate + '%');
    updateKPI('dashKpiGrid', 'low_rate', portfolio.low_rate + '%');

    // Update hotel cards
    var grid = document.getElementById('hotelGrid');
    if (!grid) return;

    var cards = grid.querySelectorAll('.hotel-card');
    cards.forEach(function(card) {
      var key = card.dataset.key;
      var kpi = portfolio.hotels[key];
      if (!kpi) return;

      var tc = detail.tierColor[kpi.tier] || '#6B7280';
      var cleanColor = kpi.cleaning_issue_rate > 5 ? 'var(--red)' : kpi.cleaning_issue_rate > 0 ? 'var(--orange)' : 'var(--green)';

      // Update data attributes for filtering/sorting
      card.dataset.tier = kpi.tier;
      card.dataset.reviews = kpi.total_reviews;
      card.dataset.highRate = kpi.high_rate;
      card.dataset.cleaning = kpi.cleaning_issue_rate;

      // Update score display
      var scoreEl = card.querySelector('[style*="font-size:2rem"]');
      if (scoreEl) {
        scoreEl.textContent = kpi.overall_avg_10pt;
        scoreEl.style.color = tc;
      }

      // Update score bar
      var barFill = card.querySelector('[style*="height:100%;border-radius:4px;width:"]');
      if (barFill) {
        barFill.style.width = (kpi.overall_avg_10pt / 10 * 100).toFixed(0) + '%';
        barFill.style.background = tc;
      }

      // Update rate badges
      var rateEls = card.querySelectorAll('[style*="flex:1;text-align:center"]');
      if (rateEls.length >= 3) {
        rateEls[0].querySelector('[style*="font-weight:700"]').textContent = kpi.high_rate + '%';
        rateEls[1].querySelector('[style*="font-weight:700"]').textContent = kpi.low_rate + '%';
        rateEls[2].querySelector('[style*="font-weight:700"]').textContent = kpi.cleaning_issue_rate + '%';
        rateEls[2].querySelector('[style*="font-weight:700"]').style.color = cleanColor;
      }

      // Update reviews count
      var countSpan = card.querySelector('.kpi-sub, [style*="justify-content:space-between"] span:last-child');
      // Find the reviews count span
      var spans = card.querySelectorAll('span');
      spans.forEach(function(span) {
        if (span.textContent.match(/^\d+件$/)) {
          span.textContent = kpi.total_reviews + '件';
        }
      });
    });

    applyFilters();
  }

  function updateKPI(gridId, kpiName, value) {
    var grid = document.getElementById(gridId);
    if (!grid) return;
    var el = grid.querySelector('[data-kpi="' + kpiName + '"]');
    if (el) el.textContent = value;
  }

  // Listen for date filter changes
  document.addEventListener('dateFilterChanged', function(e) {
    updateCardsFromFilter(e.detail);
  });

  // Load data and initialize
  function loadData() {
    var loaded = 0;
    var total = 3;

    function checkReady() {
      loaded++;
      if (loaded >= total) init();
    }

    fetch('data/hotel-details.json')
      .then(function(r) { return r.json(); })
      .then(function(data) { hotelDetails = data; checkReady(); })
      .catch(function() { checkReady(); });

    fetch('data/hotel-ranked.json')
      .then(function(r) { return r.json(); })
      .then(function(data) { hotelRanked = data; checkReady(); })
      .catch(function() { checkReady(); });

    fetch('data/tier-color.json')
      .then(function(r) { return r.json(); })
      .then(function(data) { tierColor = data; checkReady(); })
      .catch(function() { checkReady(); });
  }

  document.addEventListener('DOMContentLoaded', loadData);

  // Expose globals
  window.openModal = openModal;
  window.applyFilters = applyFilters;
  window.sortCards = sortCards;
})();
