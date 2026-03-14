// Page Snapshot Switcher V2 - PRIMECHANGE Portal
(function() {
  'use strict';

  var PAGE_MAP = {
    'da-content': 'deep-analysis-content.json',
    'ri-content': 'revenue-impact-content.json',
    'ap-content': 'action-plans-content.json'
  };

  var containerId = null;
  var contentFile = null;
  var latestHTML = null;

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
    if (containerId === 'da-content') initTabs();
    if (containerId === 'ap-content') initAccordions();
  }

  function initTabs() {
    document.querySelectorAll('#da-content .tab-btn').forEach(function(btn) {
      btn.addEventListener('click', function() {
        var tabId = btn.dataset.tab;
        if (tabId && typeof window.showTab === 'function') {
          window.showTab(tabId);
        }
      });
    });
  }

  function initAccordions() {
    document.querySelectorAll('#ap-content .accordion-header').forEach(function(header) {
      header.addEventListener('click', function() {
        header.parentElement.classList.toggle('open');
      });
    });
  }

  function loadSnapshotContent(snapshotId) {
    if (!containerId || !contentFile) return;
    if (!snapshotId) {
      if (latestHTML !== null) applyContent(latestHTML);
      return;
    }

    var url = 'data/snapshots/' + snapshotId + '/' + contentFile;
    fetch(url).then(function(res) {
      if (!res.ok) throw new Error('HTTP ' + res.status);
      return res.json();
    }).then(function(data) {
      if (containerId === 'da-content' && data.tabs) {
        var tabBtns = data.tabs.map(function(t, i) {
          return '<button class="tab-btn' + (i === 0 ? ' active' : '') + '" data-tab="' + t.id + '" onclick="showTab(\'' + t.id + '\')">' + t.icon + ' 分析' + t.id + '</button>';
        }).join('\n');
        var tabPanels = data.tabs.map(function(t, i) {
          return '<div class="tab-panel' + (i === 0 ? ' active' : '') + '" id="tab-' + t.id + '">' + t.html + '</div>';
        }).join('\n');
        applyContent('<div class="tabs">' + tabBtns + '</div>\n' + tabPanels);
      } else if (data.html) {
        applyContent(data.html);
      }
    }).catch(function(err) {
      console.warn('Snapshot content not available for ' + snapshotId + ':', err.message);
    });
  }

  document.addEventListener('DOMContentLoaded', function() {
    if (!detectPage()) return;
    cacheLatest();
  });

  document.addEventListener('snapshotChanged', function(e) {
    if (!containerId) return;
    loadSnapshotContent(e.detail ? e.detail.snapshotId : null);
  });
})();
