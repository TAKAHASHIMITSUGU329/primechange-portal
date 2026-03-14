// CS Charts V2 - CS visualization (placeholder for future chart enhancements)
(function() {
  'use strict';

  // CS matrix highlighting on hover
  document.addEventListener('DOMContentLoaded', function() {
    var matrix = document.querySelector('.cs-matrix table');
    if (!matrix) return;

    matrix.querySelectorAll('td').forEach(function(td) {
      td.addEventListener('mouseenter', function() {
        td.style.outline = '2px solid var(--accent)';
      });
      td.addEventListener('mouseleave', function() {
        td.style.outline = '';
      });
    });
  });
})();
