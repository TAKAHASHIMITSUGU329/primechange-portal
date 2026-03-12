// SVG Gauge Chart - PRIMECHANGE Portal
(function() {
  'use strict';

  function createSVGGauge(container, value, target, options) {
    options = options || {};
    var color = options.color || '#3B82F6';
    var unit = options.unit || '';
    var label = options.label || '';
    var size = options.size || 120;
    var strokeWidth = options.strokeWidth || 8;
    var lower = options.lower || false;

    // Calculate percentage (0-100)
    var pct;
    if (lower) {
      // For metrics where lower is better (e.g., cleaning issue rate)
      pct = Math.max(0, Math.min(100, (1 - value / (target * 3)) * 100));
    } else {
      pct = Math.max(0, Math.min(100, (value / target) * 100));
    }

    var cx = size / 2;
    var cy = size / 2;
    var r = (size - strokeWidth * 2) / 2;
    var circumference = 2 * Math.PI * r;
    var arcLength = circumference * 0.75; // 270 degree arc
    var offset = arcLength * (1 - pct / 100);
    var rotation = 135; // Start from bottom-left

    // Determine status color based on achievement
    var statusColor = color;
    if (pct >= 90) statusColor = '#10B981';
    else if (pct >= 70) statusColor = color;
    else if (pct >= 50) statusColor = '#F59E0B';
    else statusColor = '#EF4444';

    if (lower) {
      // Reverse color logic for lower-is-better
      if (value <= target) statusColor = '#10B981';
      else if (value <= target * 1.5) statusColor = '#F59E0B';
      else statusColor = '#EF4444';
    }

    var svg = [
      '<svg viewBox="0 0 ' + size + ' ' + size + '" width="' + size + '" height="' + size + '" role="img" aria-label="' + label + ': ' + value + unit + ' (目標: ' + target + unit + ')">',
      '  <title>' + label + ': ' + value + unit + '</title>',
      // Background arc
      '  <circle cx="' + cx + '" cy="' + cy + '" r="' + r + '"',
      '    fill="none" stroke="#E2E8F0" stroke-width="' + strokeWidth + '"',
      '    stroke-dasharray="' + arcLength + ' ' + circumference + '"',
      '    stroke-linecap="round"',
      '    transform="rotate(' + rotation + ' ' + cx + ' ' + cy + ')" />',
      // Value arc
      '  <circle cx="' + cx + '" cy="' + cy + '" r="' + r + '"',
      '    fill="none" stroke="' + statusColor + '" stroke-width="' + strokeWidth + '"',
      '    stroke-dasharray="' + arcLength + ' ' + circumference + '"',
      '    stroke-dashoffset="' + offset + '"',
      '    stroke-linecap="round"',
      '    transform="rotate(' + rotation + ' ' + cx + ' ' + cy + ')"',
      '    class="gauge-arc" style="transition: stroke-dashoffset 1s ease-out;" />',
      // Value text
      '  <text x="' + cx + '" y="' + (cy - 2) + '" text-anchor="middle" dominant-baseline="central"',
      '    font-size="' + (size * 0.2) + '" font-weight="800" fill="' + statusColor + '">',
      '    ' + value + unit,
      '  </text>',
      // Target text
      '  <text x="' + cx + '" y="' + (cy + size * 0.18) + '" text-anchor="middle"',
      '    font-size="' + (size * 0.09) + '" fill="#94A3B8">',
      '    目標: ' + target + unit,
      '  </text>',
      '</svg>'
    ].join('\n');

    if (typeof container === 'string') {
      container = document.getElementById(container);
    }
    if (container) {
      container.innerHTML = svg;
    }
    return svg;
  }

  // Auto-initialize gauges from data attributes
  function initGauges() {
    var els = document.querySelectorAll('.svg-gauge');
    els.forEach(function(el) {
      var value = parseFloat(el.dataset.value);
      var target = parseFloat(el.dataset.target);
      var options = {
        color: el.dataset.color || '#3B82F6',
        unit: el.dataset.unit || '',
        label: el.dataset.label || '',
        lower: el.dataset.lower === 'true',
        size: parseInt(el.dataset.size) || 120,
        strokeWidth: parseInt(el.dataset.strokeWidth) || 8
      };
      createSVGGauge(el, value, target, options);
    });
  }

  // Animate on scroll into view
  function animateOnView() {
    var observer = new IntersectionObserver(function(entries) {
      entries.forEach(function(entry) {
        if (entry.isIntersecting) {
          var arcs = entry.target.querySelectorAll('.gauge-arc');
          arcs.forEach(function(arc) {
            var finalOffset = arc.style.strokeDashoffset || arc.getAttribute('stroke-dashoffset');
            var dashArray = arc.getAttribute('stroke-dasharray');
            var arcLength = parseFloat(dashArray);
            arc.setAttribute('stroke-dashoffset', arcLength);
            // Force reflow
            arc.getBoundingClientRect();
            arc.style.strokeDashoffset = finalOffset;
          });
          observer.unobserve(entry.target);
        }
      });
    }, { threshold: 0.3 });

    document.querySelectorAll('.svg-gauge').forEach(function(el) {
      observer.observe(el);
    });
  }

  document.addEventListener('DOMContentLoaded', function() {
    initGauges();
    animateOnView();
  });

  window.createSVGGauge = createSVGGauge;
})();
