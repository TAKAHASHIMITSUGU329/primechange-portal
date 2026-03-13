// Common UI utilities - PRIMECHANGE Portal
(function() {
  'use strict';

  // HTML escape
  function escHtml(s) {
    var d = document.createElement('div');
    d.textContent = s;
    return d.innerHTML;
  }

  // Tab switching
  function showTab(id) {
    document.querySelectorAll('.tab-btn').forEach(function(b) {
      b.classList.toggle('active', b.textContent.includes('分析' + id));
    });
    document.querySelectorAll('.tab-panel').forEach(function(p) {
      p.classList.toggle('active', p.id === 'tab-' + id);
    });
  }

  // Accordion toggle
  document.addEventListener('click', function(e) {
    var header = e.target.closest('.accordion-header');
    if (header) {
      header.parentElement.classList.toggle('open');
    }
  });

  // Mobile nav toggle
  document.addEventListener('click', function(e) {
    if (e.target.closest('.nav-toggle')) {
      var links = document.querySelector('.nav-links');
      if (links) links.classList.toggle('open');
    }
  });

  // Navbar auto-hide on scroll
  var lastScroll = 0;
  var nav = document.querySelector('.main-nav');
  if (nav) {
    window.addEventListener('scroll', function() {
      var currentScroll = window.scrollY;
      if (currentScroll > lastScroll && currentScroll > 100) {
        nav.style.transform = 'translateY(-100%)';
      } else {
        nav.style.transform = 'translateY(0)';
      }
      lastScroll = currentScroll;
    }, { passive: true });
  }

  // Animate bars on scroll into view
  var barObserver = new IntersectionObserver(function(entries) {
    entries.forEach(function(entry) {
      if (entry.isIntersecting) {
        var fills = entry.target.querySelectorAll('.h-bar-fill');
        fills.forEach(function(fill) {
          var width = fill.style.width;
          fill.style.width = '0';
          fill.getBoundingClientRect();
          fill.style.width = width;
        });
        barObserver.unobserve(entry.target);
      }
    });
  }, { threshold: 0.1 });

  document.addEventListener('DOMContentLoaded', function() {
    document.querySelectorAll('.h-bar').forEach(function(bar) {
      barObserver.observe(bar);
    });
  });

  // ESC key to close modal
  document.addEventListener('keydown', function(e) {
    if (e.key === 'Escape') {
      var overlay = document.getElementById('modalOverlay');
      if (overlay && overlay.classList.contains('active')) {
        closeModal();
      }
    }
  });

  // Close modal function
  function closeModal() {
    var overlay = document.getElementById('modalOverlay');
    if (overlay) {
      overlay.classList.remove('active');
      document.body.style.overflow = '';
    }
  }

  // Generic KPI update helper
  function updateKPI(kpiName, value) {
    document.querySelectorAll('[data-kpi="' + kpiName + '"]').forEach(function(el) {
      el.textContent = value;
    });
  }

  // Diff badge helper: creates a small badge showing change from previous snapshot
  // direction: 'higher_better' (green=positive) or 'lower_better' (green=negative)
  function createDiffBadge(current, previous, direction) {
    if (previous === null || previous === undefined) return '';
    var diff = Math.round((current - previous) * 100) / 100;
    if (diff === 0) return '<span class="kpi-diff neutral">\u2015 0</span>';

    var isImprovement = direction === 'lower_better' ? diff < 0 : diff > 0;
    var cls = isImprovement ? 'positive' : 'negative';
    var arrow = diff > 0 ? '\u25B2' : '\u25BC';
    var label = arrow + (diff > 0 ? '+' : '') + diff;

    return '<span class="kpi-diff ' + cls + '">' + label + '</span>';
  }

  // Find previous snapshot entry from index
  function findPreviousSnapshot(snapshots, currentId) {
    if (!snapshots || snapshots.length < 2) return null;
    // snapshots are sorted by date ascending
    for (var i = 1; i < snapshots.length; i++) {
      if (snapshots[i].id === currentId) return snapshots[i - 1];
    }
    // If currentId is null (latest), return second-to-last
    if (!currentId) return snapshots[snapshots.length - 2];
    return null;
  }

  // Index page: update KPIs on date filter change
  document.addEventListener('dateFilterChanged', function(e) {
    var p = e.detail.portfolio;
    if (!p) return;

    var snapshots = e.detail.snapshots;
    var snapshotId = e.detail.snapshotId;
    var prev = findPreviousSnapshot(snapshots, snapshotId);

    // Index page KPIs
    var indexGrid = document.getElementById('indexKpiGrid');
    if (indexGrid) {
      updateKPI('total_hotels', p.total_hotels);
      updateKPI('total_reviews', p.total_reviews.toLocaleString());
      updateKPI('avg_score', p.avg_score);
      updateKPI('high_rate', p.high_rate + '%');
      updateKPI('cleaning_issue_rate', p.cleaning_issue_rate + '%');
      updateKPI('cleaning_issue_count', p.cleaning_issue_count + '件');

      // Add diff badges if previous snapshot data available (only when viewing full period)
      if (prev && !e.detail.range) {
        appendDiffBadge(indexGrid, 'total_reviews', p.total_reviews, prev.total_reviews, 'higher_better');
        appendDiffBadge(indexGrid, 'avg_score', p.avg_score, prev.avg_score, 'higher_better');
        appendDiffBadge(indexGrid, 'high_rate', p.high_rate, prev.high_rate, 'higher_better');
        appendDiffBadge(indexGrid, 'cleaning_issue_rate', p.cleaning_issue_rate, prev.cleaning_issue_rate, 'lower_better');
      } else {
        clearDiffBadges(indexGrid);
      }
    }

    // Dashboard KPIs
    var dashGrid = document.getElementById('dashKpiGrid');
    if (dashGrid) {
      updateKPI('total_hotels', p.total_hotels);
      updateKPI('total_reviews', p.total_reviews.toLocaleString());
      updateKPI('avg_score', p.avg_score);
      updateKPI('high_rate', p.high_rate + '%');
      updateKPI('low_rate', p.low_rate + '%');
    }

    // Cleaning strategy KPIs
    var cleanGrid = document.getElementById('cleaningKpiGrid');
    if (cleanGrid) {
      var cleanRate = cleanGrid.querySelector('[data-kpi="cleaning_issue_rate"]');
      var cleanCount = cleanGrid.querySelector('[data-kpi="cleaning_issue_count"]');
      if (cleanRate) cleanRate.textContent = p.cleaning_issue_rate + '%';
      if (cleanCount) cleanCount.textContent = p.cleaning_issue_count;

      if (prev && !e.detail.range) {
        appendDiffBadge(cleanGrid, 'cleaning_issue_rate', p.cleaning_issue_rate, prev.cleaning_issue_rate, 'lower_better');
      } else {
        clearDiffBadges(cleanGrid);
      }
    }
  });

  function appendDiffBadge(container, kpiName, current, previous, direction) {
    var el = container.querySelector('[data-kpi="' + kpiName + '"]');
    if (!el) return;
    // Remove existing badge
    var existing = el.parentElement.querySelector('.kpi-diff');
    if (existing) existing.remove();
    var badge = createDiffBadge(current, previous, direction);
    if (badge) el.insertAdjacentHTML('afterend', badge);
  }

  function clearDiffBadges(container) {
    container.querySelectorAll('.kpi-diff').forEach(function(b) { b.remove(); });
  }

  // Expose globals
  window.escHtml = escHtml;
  window.showTab = showTab;
  window.closeModal = closeModal;
})();
