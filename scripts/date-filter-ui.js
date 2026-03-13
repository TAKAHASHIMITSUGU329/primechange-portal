// Date Filter UI - PRIMECHANGE Portal
// Inserts date filter bar below .page-subtitle on all pages
(function() {
  'use strict';

  var filterBar = null;
  var statusEl = null;
  var startInput = null;
  var endInput = null;
  var snapshotSelect = null;
  var dataRange = null;

  function formatDate(d) {
    return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
  }

  function createFilterBar() {
    var subtitle = document.querySelector('.page-subtitle');
    if (!subtitle) return;

    filterBar = document.createElement('div');
    filterBar.className = 'date-filter-bar';
    filterBar.innerHTML =
      '<div class="date-filter-row">' +
        '<div class="date-filter-label">期間フィルター:</div>' +
        '<div class="date-preset-group">' +
          '<button class="date-preset-btn active" data-preset="all">全期間</button>' +
          '<button class="date-preset-btn" data-preset="this-month">今月</button>' +
          '<button class="date-preset-btn" data-preset="last-month">先月</button>' +
          '<button class="date-preset-btn" data-preset="7days">直近7日</button>' +
          '<button class="date-preset-btn" data-preset="30days">直近30日</button>' +
        '</div>' +
        '<div class="date-range-inputs">' +
          '<input type="date" id="dateStart" class="date-input">' +
          '<span class="date-range-sep">〜</span>' +
          '<input type="date" id="dateEnd" class="date-input">' +
          '<button class="date-apply-btn" id="dateApplyBtn">適用</button>' +
        '</div>' +
        '<div class="date-filter-spacer"></div>' +
        '<select id="snapshotSelect" class="date-snapshot-select" style="display:none;">' +
          '<option value="">最新データ</option>' +
        '</select>' +
      '</div>' +
      '<div class="date-filter-status" id="dateFilterStatus"></div>';

    subtitle.parentNode.insertBefore(filterBar, subtitle.nextSibling);

    startInput = document.getElementById('dateStart');
    endInput = document.getElementById('dateEnd');
    statusEl = document.getElementById('dateFilterStatus');
    snapshotSelect = document.getElementById('snapshotSelect');

    // Preset buttons
    filterBar.querySelectorAll('.date-preset-btn').forEach(function(btn) {
      btn.addEventListener('click', function() {
        filterBar.querySelectorAll('.date-preset-btn').forEach(function(b) { b.classList.remove('active'); });
        btn.classList.add('active');
        applyPreset(btn.dataset.preset);
      });
    });

    // Custom date apply
    document.getElementById('dateApplyBtn').addEventListener('click', function() {
      filterBar.querySelectorAll('.date-preset-btn').forEach(function(b) { b.classList.remove('active'); });
      var start = startInput.value || null;
      var end = endInput.value || null;
      if (start || end) {
        window.DateFilter.apply(start, end);
      }
    });

    // Snapshot select
    snapshotSelect.addEventListener('change', function() {
      var val = snapshotSelect.value;
      if (!val) {
        // Reload current data
        window.location.reload();
      } else {
        // Load snapshot data
        loadSnapshot(val);
      }
    });
  }

  function applyPreset(preset) {
    var today = new Date();
    var start = null, end = null;

    switch (preset) {
      case 'all':
        break;
      case 'this-month':
        start = formatDate(new Date(today.getFullYear(), today.getMonth(), 1));
        end = formatDate(today);
        break;
      case 'last-month':
        var lastMonth = new Date(today.getFullYear(), today.getMonth() - 1, 1);
        start = formatDate(lastMonth);
        end = formatDate(new Date(today.getFullYear(), today.getMonth(), 0));
        break;
      case '7days':
        var d7 = new Date(today);
        d7.setDate(d7.getDate() - 6);
        start = formatDate(d7);
        end = formatDate(today);
        break;
      case '30days':
        var d30 = new Date(today);
        d30.setDate(d30.getDate() - 29);
        start = formatDate(d30);
        end = formatDate(today);
        break;
    }

    if (start) startInput.value = start;
    else startInput.value = '';
    if (end) endInput.value = end;
    else endInput.value = '';

    window.DateFilter.apply(start, end);
  }

  function updateStatus(detail) {
    if (!statusEl) return;

    if (!detail.range) {
      statusEl.textContent = '';
      statusEl.style.display = 'none';
    } else {
      var rangeText = (detail.range.start || '開始') + ' 〜 ' + (detail.range.end || '最新');
      statusEl.textContent = rangeText + ' (' + detail.filteredCount.toLocaleString() + '件)';
      statusEl.style.display = 'block';
    }
  }

  function populateSnapshots(snapshots) {
    if (!snapshotSelect || !snapshots || snapshots.length === 0) return;
    snapshotSelect.style.display = 'block';
    snapshots.forEach(function(snap) {
      var opt = document.createElement('option');
      opt.value = snap.id || snap.date;
      opt.textContent = snap.date + ' (' + snap.total_reviews + '件)';
      snapshotSelect.appendChild(opt);
    });
  }

  function loadSnapshot(snapshotId) {
    // For static site, snapshots are stored under data/snapshots/{id}/
    var basePath = 'data/snapshots/' + snapshotId + '/';
    fetch(basePath + 'hotel-reviews-all.json')
      .then(function(r) { return r.json(); })
      .then(function(data) {
        // Replace the current data and re-apply filter
        var reviews = window.DateFilter.getReviews();
        Object.keys(data).forEach(function(k) { reviews[k] = data[k]; });
        // Reset to all
        filterBar.querySelectorAll('.date-preset-btn').forEach(function(b) { b.classList.remove('active'); });
        filterBar.querySelector('[data-preset="all"]').classList.add('active');
        startInput.value = '';
        endInput.value = '';
        window.DateFilter.apply(null, null);
      })
      .catch(function() {
        console.warn('Snapshot not found: ' + snapshotId);
      });
  }

  // Set min/max on date inputs when data range is known
  function setDateBounds(range) {
    if (!range || !startInput || !endInput) return;
    dataRange = range;
    if (range.min) {
      startInput.min = range.min;
      endInput.min = range.min;
    }
    if (range.max) {
      startInput.max = range.max;
      endInput.max = range.max;
    }
  }

  // Listen for filter engine ready
  document.addEventListener('dateFilterReady', function(e) {
    if (e.detail.dataRange) setDateBounds(e.detail.dataRange);
    if (e.detail.snapshots) populateSnapshots(e.detail.snapshots);
  });

  // Listen for filter changes to update status
  document.addEventListener('dateFilterChanged', function(e) {
    updateStatus(e.detail);
  });

  document.addEventListener('DOMContentLoaded', createFilterBar);
})();
