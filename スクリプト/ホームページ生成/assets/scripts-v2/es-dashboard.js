// ES Dashboard V2 - client-side interactions
(function() {
  'use strict';

  // Sort table by column
  function sortTable(tableId, colIndex) {
    var table = document.getElementById(tableId);
    if (!table) return;
    var tbody = table.querySelector('tbody');
    if (!tbody) return;
    var rows = Array.from(tbody.rows);

    var ascending = table.dataset.sortCol === String(colIndex) && table.dataset.sortDir === 'asc';
    table.dataset.sortCol = colIndex;
    table.dataset.sortDir = ascending ? 'desc' : 'asc';

    rows.sort(function(a, b) {
      var aVal = a.cells[colIndex].textContent.replace(/[^0-9.\-]/g, '');
      var bVal = b.cells[colIndex].textContent.replace(/[^0-9.\-]/g, '');
      var aNum = parseFloat(aVal) || 0;
      var bNum = parseFloat(bVal) || 0;
      return ascending ? aNum - bNum : bNum - aNum;
    });

    rows.forEach(function(row) { tbody.appendChild(row); });
  }

  document.addEventListener('DOMContentLoaded', function() {
    // Make table headers clickable for sorting
    document.querySelectorAll('.sortable-table th').forEach(function(th, idx) {
      th.style.cursor = 'pointer';
      th.addEventListener('click', function() {
        var table = th.closest('table');
        if (table) sortTable(table.id, idx);
      });
    });
  });

  window.sortESTable = sortTable;
})();
