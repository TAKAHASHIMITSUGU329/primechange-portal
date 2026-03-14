// Action Tracker V2 - Status management with localStorage
(function() {
  'use strict';

  var STORAGE_KEY = 'primechange_action_status_v2';
  var actionStatus = {};

  function loadStatus() {
    // Try localStorage first, then fall back to server JSON
    var saved = localStorage.getItem(STORAGE_KEY);
    if (saved) {
      try { actionStatus = JSON.parse(saved); return; }
      catch(e) { /* fall through */ }
    }

    // Load initial state from server
    fetch('data/action-status.json')
      .then(function(r) { return r.json(); })
      .then(function(data) {
        actionStatus = data || {};
        localStorage.setItem(STORAGE_KEY, JSON.stringify(actionStatus));
      })
      .catch(function() { actionStatus = {}; });
  }

  function saveStatus() {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(actionStatus));
    updateProgressBar();
  }

  function getStatusLabel(status) {
    switch(status) {
      case 'in-progress': return '進行中';
      case 'completed': return '完了';
      default: return '未着手';
    }
  }

  function getNextStatus(current) {
    switch(current) {
      case 'not-started': return 'in-progress';
      case 'in-progress': return 'completed';
      case 'completed': return 'not-started';
      default: return 'in-progress';
    }
  }

  function cycleStatus(actionId) {
    var current = (actionStatus[actionId] && actionStatus[actionId].status) || 'not-started';
    var next = getNextStatus(current);

    if (!actionStatus[actionId]) {
      actionStatus[actionId] = { status: 'not-started', assignee: '', deadline: '' };
    }
    actionStatus[actionId].status = next;
    saveStatus();

    // Update badge
    var badge = document.querySelector('[data-action-id="' + actionId + '"]');
    if (badge) {
      badge.className = 'status-badge ' + next;
      badge.textContent = getStatusLabel(next);
    }
  }

  function updateProgressBar() {
    var total = 0, completed = 0;
    Object.keys(actionStatus).forEach(function(id) {
      total++;
      if (actionStatus[id].status === 'completed') completed++;
    });

    var pct = total > 0 ? Math.round(completed / total * 100) : 0;

    var pctEl = document.getElementById('progressPct');
    var barEl = document.getElementById('progressBar');
    var completedEl = document.getElementById('completedCount');
    var totalEl = document.getElementById('totalCount');

    if (pctEl) pctEl.textContent = pct + '%';
    if (barEl) barEl.style.width = pct + '%';
    if (completedEl) completedEl.textContent = completed;
    if (totalEl) totalEl.textContent = total;
  }

  function applyStatusFilters() {
    var statusFilter = document.getElementById('statusFilter');
    var hotelFilter = document.getElementById('hotelFilter');
    var phaseFilter = document.getElementById('phaseFilter');

    var sf = statusFilter ? statusFilter.value : 'all';
    var hf = hotelFilter ? hotelFilter.value : 'all';
    var pf = phaseFilter ? phaseFilter.value : 'all';

    document.querySelectorAll('.accordion-item[data-hotel-key]').forEach(function(item) {
      var hotelKey = item.dataset.hotelKey;
      var showHotel = hf === 'all' || hotelKey === hf;

      if (!showHotel) {
        item.style.display = 'none';
        return;
      }

      // Check if any actions in this hotel match status/phase filters
      var hasVisibleAction = false;
      item.querySelectorAll('[data-action-id]').forEach(function(badge) {
        var actionId = badge.dataset.actionId;
        var actionStatus_ = (actionStatus[actionId] && actionStatus[actionId].status) || 'not-started';
        var actionPhase = badge.dataset.phase || '';
        var li = badge.closest('li');

        var matchStatus = sf === 'all' || actionStatus_ === sf;
        var matchPhase = pf === 'all' || actionPhase === pf;

        if (li) {
          li.style.display = matchStatus && matchPhase ? '' : 'none';
        }
        if (matchStatus && matchPhase) hasVisibleAction = true;
      });

      item.style.display = hasVisibleAction ? '' : 'none';
    });
  }

  function exportStatus() {
    var blob = new Blob([JSON.stringify(actionStatus, null, 2)], { type: 'application/json' });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = 'action-status-export.json';
    a.click();
    URL.revokeObjectURL(url);
  }

  function importStatus() {
    var input = document.createElement('input');
    input.type = 'file';
    input.accept = '.json';
    input.onchange = function(e) {
      var file = e.target.files[0];
      if (!file) return;
      var reader = new FileReader();
      reader.onload = function(ev) {
        try {
          var data = JSON.parse(ev.target.result);
          actionStatus = data;
          saveStatus();
          // Re-render all badges
          Object.keys(actionStatus).forEach(function(id) {
            var badge = document.querySelector('[data-action-id="' + id + '"]');
            if (badge) {
              var st = actionStatus[id].status || 'not-started';
              badge.className = 'status-badge ' + st;
              badge.textContent = getStatusLabel(st);
            }
          });
        } catch(err) {
          alert('JSONファイルの読み込みに失敗しました');
        }
      };
      reader.readAsText(file);
    };
    input.click();
  }

  // Initialize
  document.addEventListener('DOMContentLoaded', function() {
    loadStatus();

    // Wait a bit for status to load, then apply
    setTimeout(function() {
      // Apply saved statuses to badges
      document.querySelectorAll('[data-action-id]').forEach(function(badge) {
        var id = badge.dataset.actionId;
        var st = (actionStatus[id] && actionStatus[id].status) || 'not-started';
        badge.className = 'status-badge ' + st;
        badge.textContent = getStatusLabel(st);
      });
      updateProgressBar();
    }, 500);

    // Click handler for status badges
    document.addEventListener('click', function(e) {
      var badge = e.target.closest('.status-badge[data-action-id]');
      if (badge) {
        cycleStatus(badge.dataset.actionId);
      }
    });

    // Filter handlers
    var statusFilter = document.getElementById('statusFilter');
    var hotelFilter = document.getElementById('hotelFilter');
    var phaseFilter = document.getElementById('phaseFilter');
    if (statusFilter) statusFilter.addEventListener('change', applyStatusFilters);
    if (hotelFilter) hotelFilter.addEventListener('change', applyStatusFilters);
    if (phaseFilter) phaseFilter.addEventListener('change', applyStatusFilters);

    // Export/Import buttons
    var exportBtn = document.getElementById('exportBtn');
    var importBtn = document.getElementById('importBtn');
    if (exportBtn) exportBtn.addEventListener('click', exportStatus);
    if (importBtn) importBtn.addEventListener('click', importStatus);
  });

  window.ActionTracker = {
    getStatus: function() { return actionStatus; },
    exportStatus: exportStatus,
    importStatus: importStatus
  };
})();
