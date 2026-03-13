// Page Snapshot Switcher - PRIMECHANGE Portal
// Handles snapshot switching for deep-analysis, revenue-impact, action-plans pages
// by fetching pre-rendered HTML fragments and replacing container innerHTML.
(function() {
  'use strict';

  // Map container IDs to their content JSON filenames
  var PAGE_MAP = {
    'da-content': 'deep-analysis-content.json',
    'ri-content': 'revenue-impact-content.json',
    'ap-content': 'action-plans-content.json'
  };

  var containerId = null;
  var contentFile = null;
  var latestHTML = null; // cached initial (latest) content for restore

  function detectPage() {
    var ids = Object.keys(PAGE_MAP);
    for (var i = 0; i < ids.length; i++) {
      if (document.getElementById(ids[i])) {
        containerId = ids[i];
        contentFile = PAGE_MAP[ids[i]];
        return true;
      }
    }
    return false;
  }

  function cacheLatest() {
    var el = document.getElementById(containerId);
    if (el) latestHTML = el.innerHTML;
  }

  function applyContent(html) {
    var el = document.getElementById(containerId);
    if (!el) return;
    el.innerHTML = html;

    // Re-initialize tabs if on deep-analysis page
    if (containerId === 'da-content') {
      initTabs();
    }

    // Re-initialize accordions if on action-plans page
    if (containerId === 'ap-content') {
      initAccordions();
    }
  }

  function initTabs() {
    // Re-bind tab button click handlers
    var tabBtns = document.querySelectorAll('#da-content .tab-btn');
    tabBtns.forEach(function(btn) {
      btn.addEventListener('click', function() {
        var tabId = btn.getAttribute('onclick');
        if (tabId) {
          // Extract tab ID from onclick="showTab('1')"
          var match = tabId.match(/showTab\('(\d+)'\)/);
          if (match && typeof window.showTab === 'function') {
            window.showTab(match[1]);
          }
        }
      });
    });
  }

  function initAccordions() {
    var headers = document.querySelectorAll('#ap-content .accordion-header');
    headers.forEach(function(header) {
      header.addEventListener('click', function() {
        var item = header.parentElement;
        if (item) item.classList.toggle('open');
      });
    });
  }

  function loadSnapshotContent(snapshotId) {
    if (!containerId || !contentFile) return;

    if (!snapshotId) {
      // Restore latest content without reload
      if (latestHTML !== null) {
        applyContent(latestHTML);
      }
      return;
    }

    var url = 'data/snapshots/' + snapshotId + '/' + contentFile;
    fetch(url).then(function(res) {
      if (!res.ok) throw new Error('HTTP ' + res.status);
      return res.json();
    }).then(function(data) {
      if (containerId === 'da-content' && data.tabs) {
        // Deep analysis: rebuild tabs + panels from tab array
        var tabBtns = data.tabs.map(function(t, i) {
          return '<button class="tab-btn' + (i === 0 ? ' active' : '') + '" onclick="showTab(\'' + t.id + '\')">' + t.icon + ' 分析' + t.id + '</button>';
        }).join('\n');
        var tabPanels = data.tabs.map(function(t, i) {
          return '<div class="tab-panel' + (i === 0 ? ' active' : '') + '" id="tab-' + t.id + '">' + t.html + '</div>';
        }).join('\n');
        applyContent('<div class="tabs">' + tabBtns + '</div>\n' + tabPanels);
      } else if (data.html) {
        applyContent(data.html);
      }
    }).catch(function(err) {
      // Graceful degradation: content file not available for old snapshots
      console.warn('Snapshot content not available for ' + snapshotId + ':', err.message);
    });
  }

  document.addEventListener('DOMContentLoaded', function() {
    if (!detectPage()) return;
    cacheLatest();
  });

  document.addEventListener('snapshotChanged', function(e) {
    if (!containerId) return;
    var snapshotId = e.detail ? e.detail.snapshotId : null;
    loadSnapshotContent(snapshotId);
  });
})();
