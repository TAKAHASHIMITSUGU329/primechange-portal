// Alert Engine V2 - Change detection alerts
// Reads deltas.json and displays/hides alert banners
(function() {
  'use strict';

  function renderAlerts(deltas) {
        if (!deltas || !deltas.hasDeltas || !deltas.alerts || deltas.alerts.length === 0) return;

        var container = document.getElementById('alertBannerContainer');
        if (!container) return;

        deltas.alerts.forEach(function(alert) {
          var el = document.createElement('div');
          el.className = 'alert-banner ' + (alert.severity === 'red' ? 'danger' : alert.severity === 'green' ? 'improvement' : 'info');
          el.innerHTML =
            '<div class="alert-banner-icon">' + alert.icon + '</div>' +
            '<div class="alert-banner-content">' +
            '<div class="alert-banner-title">' + alert.title + '</div>' +
            '<div class="alert-banner-msg">' + alert.message + '</div>' +
            '</div>';
          container.appendChild(el);
        });

        container.style.display = 'block';
  }

  function loadAlerts() {
    // Use inline data if available (works with file:// protocol)
    if (window.__DELTAS_DATA__) {
      renderAlerts(window.__DELTAS_DATA__);
      return;
    }
    fetch('data/deltas.json')
      .then(function(r) { return r.json(); })
      .then(renderAlerts)
      .catch(function() { /* No deltas available */ });
  }

  document.addEventListener('DOMContentLoaded', loadAlerts);
})();
