// Date Filter UI V2 - PRIMECHANGE Portal
// Same as V1 - Dark compact design
(function() {
  'use strict';

  var filterBar = null;
  var statusBar = null;
  var snapshotBanner = null;
  var startInput = null;
  var endInput = null;
  var snapshotSelect = null;
  var dataRange = null;
  var cachedSnapshots = null;

  function formatDate(d) {
    return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
  }

  function shortDate(dateStr) {
    var parts = dateStr.split('-');
    return parseInt(parts[1]) + '/' + parseInt(parts[2]);
  }

  function createFilterBar() {
    var subtitle = document.querySelector('.page-subtitle');
    if (!subtitle) return;

    filterBar = document.createElement('div');
    filterBar.className = 'df-bar';
    filterBar.innerHTML =
      '<div class="df-main">' +
        '<div class="df-label">' +
          '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="18" rx="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>' +
          '期間' +
        '</div>' +
        '<div class="df-presets">' +
          '<button class="df-preset active" data-preset="all">全期間</button>' +
          '<button class="df-preset" data-preset="this-month">今月</button>' +
          '<button class="df-preset" data-preset="last-month">先月</button>' +
          '<button class="df-preset" data-preset="7days">7日</button>' +
          '<button class="df-preset" data-preset="30days">30日</button>' +
        '</div>' +
        '<div class="df-range">' +
          '<input type="date" id="dateStart" class="df-date">' +
          '<span class="df-sep">&ndash;</span>' +
          '<input type="date" id="dateEnd" class="df-date">' +
          '<button class="df-apply" id="dateApplyBtn">適用</button>' +
        '</div>' +
        '<select id="snapshotSelect" class="df-snapshot" style="display:none;">' +
          '<option value="">最新データ</option>' +
        '</select>' +
      '</div>' +
      '<div class="df-snapshot-banner" id="snapshotBanner" style="display:none;">' +
        '<span class="df-snapshot-banner-text" id="snapshotBannerText"></span>' +
        '<button class="df-snapshot-banner-close" id="snapshotBannerClose">最新に戻す</button>' +
      '</div>' +
      '<div class="df-status" id="dateFilterStatus" style="display:none;">' +
        '<span class="df-dot"></span>' +
        '<span class="df-status-text" id="dfStatusText"></span>' +
      '</div>';

    subtitle.parentNode.insertBefore(filterBar, subtitle.nextSibling);

    startInput = document.getElementById('dateStart');
    endInput = document.getElementById('dateEnd');
    statusBar = document.getElementById('dateFilterStatus');
    snapshotSelect = document.getElementById('snapshotSelect');
    snapshotBanner = document.getElementById('snapshotBanner');

    filterBar.querySelectorAll('.df-preset').forEach(function(btn) {
      btn.addEventListener('click', function() {
        filterBar.querySelectorAll('.df-preset').forEach(function(b) { b.classList.remove('active'); });
        btn.classList.add('active');
        applyPreset(btn.dataset.preset);
      });
    });

    document.getElementById('dateApplyBtn').addEventListener('click', function() {
      filterBar.querySelectorAll('.df-preset').forEach(function(b) { b.classList.remove('active'); });
      var start = startInput.value || null;
      var end = endInput.value || null;
      if (start || end) window.DateFilter.apply(start, end);
    });

    snapshotSelect.addEventListener('change', function() {
      window.DateFilter.loadSnapshot(snapshotSelect.value || null);
    });

    document.getElementById('snapshotBannerClose').addEventListener('click', function() {
      snapshotSelect.value = '';
      window.DateFilter.loadSnapshot(null);
    });
  }

  function applyPreset(preset) {
    var today = new Date();
    var start = null, end = null;
    switch (preset) {
      case 'all': break;
      case 'this-month':
        start = formatDate(new Date(today.getFullYear(), today.getMonth(), 1));
        end = formatDate(today); break;
      case 'last-month':
        start = formatDate(new Date(today.getFullYear(), today.getMonth() - 1, 1));
        end = formatDate(new Date(today.getFullYear(), today.getMonth(), 0)); break;
      case '7days':
        var d7 = new Date(today); d7.setDate(d7.getDate() - 6);
        start = formatDate(d7); end = formatDate(today); break;
      case '30days':
        var d30 = new Date(today); d30.setDate(d30.getDate() - 29);
        start = formatDate(d30); end = formatDate(today); break;
    }
    if (start) startInput.value = start; else startInput.value = '';
    if (end) endInput.value = end; else endInput.value = '';
    window.DateFilter.apply(start, end);
  }

  function updateStatus(detail) {
    if (!statusBar) return;
    if (!detail.range && !detail.snapshotId) {
      statusBar.style.display = 'none';
    } else if (detail.range) {
      document.getElementById('dfStatusText').textContent =
        'フィルター適用中: ' + (detail.range.start || '開始') + ' 〜 ' + (detail.range.end || '最新') + '（' + detail.filteredCount.toLocaleString() + '件）';
      statusBar.style.display = 'flex';
    }
  }

  function updateSnapshotBanner(snapshotId) {
    if (!snapshotBanner) return;
    if (!snapshotId) { snapshotBanner.style.display = 'none'; return; }
    var snap = findSnapshot(snapshotId);
    var text = '\uD83D\uDCC5 ' + snapshotId + ' のデータを表示中';
    if (snap) {
      text += '（' + snap.total_reviews.toLocaleString() + '件';
      if (snap.avg_score) text += ' / ' + snap.avg_score + 'pt';
      text += '）';
    }
    document.getElementById('snapshotBannerText').textContent = text;
    snapshotBanner.style.display = 'flex';
  }

  function findSnapshot(id) {
    if (!cachedSnapshots) return null;
    for (var i = 0; i < cachedSnapshots.length; i++) {
      if (cachedSnapshots[i].id === id) return cachedSnapshots[i];
    }
    return null;
  }

  function populateSnapshots(snapshots) {
    if (!snapshotSelect || !snapshots || snapshots.length === 0) return;
    cachedSnapshots = snapshots;
    if (snapshots.length <= 1) return;
    snapshotSelect.style.display = 'block';
    var latest = snapshots[snapshots.length - 1];
    snapshots.slice().reverse().forEach(function(snap) {
      var opt = document.createElement('option');
      opt.value = snap.id || snap.date;
      var label = shortDate(snap.date) + ' (' + snap.total_reviews.toLocaleString() + '件';
      if (snap.avg_score) label += ' / ' + snap.avg_score + 'pt';
      label += ')';
      if (latest && snap.id !== latest.id) {
        var diff = snap.total_reviews - latest.total_reviews;
        label += ' [' + (diff >= 0 ? '+' : '') + diff + '件]';
      }
      opt.textContent = label;
      snapshotSelect.appendChild(opt);
    });
  }

  function setDateBounds(range) {
    if (!range || !startInput || !endInput) return;
    dataRange = range;
    if (range.min) { startInput.min = range.min; endInput.min = range.min; }
    if (range.max) { startInput.max = range.max; endInput.max = range.max; }
  }

  document.addEventListener('dateFilterReady', function(e) {
    if (e.detail.dataRange) setDateBounds(e.detail.dataRange);
    if (e.detail.snapshots) populateSnapshots(e.detail.snapshots);
  });
  document.addEventListener('dateFilterChanged', function(e) { updateStatus(e.detail); });
  document.addEventListener('snapshotChanged', function(e) {
    updateSnapshotBanner(e.detail.snapshotId);
    if (filterBar) {
      filterBar.querySelectorAll('.df-preset').forEach(function(b) { b.classList.remove('active'); });
      filterBar.querySelector('[data-preset="all"]').classList.add('active');
      if (startInput) startInput.value = '';
      if (endInput) endInput.value = '';
    }
  });

  document.addEventListener('DOMContentLoaded', createFilterBar);
})();
