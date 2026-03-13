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

  // Index page: update KPIs on date filter change
  document.addEventListener('dateFilterChanged', function(e) {
    var p = e.detail.portfolio;
    if (!p) return;

    // Index page KPIs
    var indexGrid = document.getElementById('indexKpiGrid');
    if (indexGrid) {
      updateKPI('total_hotels', p.total_hotels);
      updateKPI('total_reviews', p.total_reviews.toLocaleString());
      updateKPI('avg_score', p.avg_score);
      updateKPI('high_rate', p.high_rate + '%');
      updateKPI('cleaning_issue_rate', p.cleaning_issue_rate + '%');
      updateKPI('cleaning_issue_count', p.cleaning_issue_count + '件');
    }

    // Cleaning strategy KPIs
    var cleanGrid = document.getElementById('cleaningKpiGrid');
    if (cleanGrid) {
      var cleanRate = cleanGrid.querySelector('[data-kpi="cleaning_issue_rate"]');
      var cleanCount = cleanGrid.querySelector('[data-kpi="cleaning_issue_count"]');
      if (cleanRate) cleanRate.textContent = p.cleaning_issue_rate + '%';
      if (cleanCount) cleanCount.textContent = p.cleaning_issue_count;
    }
  });

  // Expose globals
  window.escHtml = escHtml;
  window.showTab = showTab;
  window.closeModal = closeModal;
})();
