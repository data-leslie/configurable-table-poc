(function () {
  'use strict';

  // ===========================================================================
  // Configurable Table — NETWORK-ENABLED PoC
  //   Same table as the data app, but self-hosted (not sandboxed), so the two
  //   things the data-app sandbox blocked now work:
  //     • real file downloads (CSV + PDF)
  //     • per-user saved layouts (localStorage)
  //   Uses a small SAMPLE dataset so this PoC proves the *environment* with no
  //   data-access dependency. Real data wiring is the next step.
  // ===========================================================================

  var root = document.getElementById('app');

  // --- sample data (stand-in for the live Azure Lakehouse query) -------------
  var MEAS_GBP = 'Total Price GBP', MEAS_LOCAL = 'Total Price Local', CAMPAIGN = 'Campaign Name';
  // 100% FICTIONAL demo data — no real advertisers, campaigns, or figures.
  var SAMPLE = [
    ['Team North','Media Agency A','Acme Retail','Acme Spring Brand Push','UK','2024','January',12000,12000],
    ['Team South','Media Agency B','Globex','Globex Q3 Network Deal','NL','2025','July',8500,9800],
    ['Trading Desk','Media Agency A','Globex','Globex Always-On W2-3','UK','2026','January',4500,4500],
    ['Direct','Direct','Initech','Initech Full-Year Panel','SE','2023','February',3000,36000],
    ['Team North','Media Agency C','Umbrella Foods','Umbrella Winter Loop','DK','2024','December',21000,180000],
    ['Direct','Direct','Soylent Corp','Soylent 2023 Programme','SE','2023','January',6000,72000],
    ['Team South','Media Agency C','Hooli','Hooli City Centre Takeover','DK','2024','July',9000,74000],
    ['Direct','Trading Desk','Stark Industries','Stark Loop Burst','DK','2025','June',11000,98000],
    ['Team North','Media Agency D','Wayne Enterprises','Wayne EV Launch','UK','2026','March',63000,63000],
    ['Team South','Media Agency B','Wonka Co','Wonka Fusion National','SE','2026','February',110000,1500000],
    ['Team North','Media Agency E','Cyberdyne','Cyberdyne Christmas Loop','DE','2025','November',49000,55000],
    ['Team South','Media Agency D','Acme Retail','Acme Winter Sale Network','NL','2025','January',27000,31000],
    ['Direct','Direct','Local Co','Local New-Year Push','UK','2026','January',3000,3000],
    ['Team North','Media Agency A','Globex','Globex Teaser Flight','DE','2026','April',75000,85000],
    ['Team South','Media Agency B','Stark Industries','Stark Day Takeover','UK','2025','March',54000,54000],
    ['Team North','Media Agency C','Umbrella Foods','Umbrella 5G City Network','UK','2024','September',41000,41000]
  ];
  var FIELDS = ['Specialist', 'Agency', 'Advertiser Group', CAMPAIGN, 'Country', 'Year', 'Month', MEAS_GBP, MEAS_LOCAL];
  function rowsFromSample() {
    return SAMPLE.map(function (r) { var o = {}; FIELDS.forEach(function (f, i) { o[f] = r[i]; }); return o; });
  }

  var COLS = [
    { colId: 'Specialist', field: 'Specialist', headerName: 'Specialist', width: 164 },
    { colId: 'Agency', field: 'Agency', headerName: 'Agency', width: 190 },
    { colId: 'Advertiser Group', field: 'Advertiser Group', headerName: 'Advertiser Group', width: 190 },
    { colId: 'Campaign Name', field: 'Campaign Name', headerName: 'Campaign Name', width: 320 },
    { colId: 'Country', field: 'Country', headerName: 'Country', width: 100 },
    { colId: 'Year', field: 'Year', headerName: 'Year', width: 90, yearSort: true },
    { colId: 'Month', field: 'Month', headerName: 'Month', width: 110, monthSort: true },
    { colId: 'Total Price', field: null, headerName: 'Total Price', width: 150, totalPrice: true }
  ];
  var FILTERS = ['Specialist', 'Agency', 'Advertiser Group', 'Campaign Name', 'Country', 'Year', 'Month'];
  var MONTHS = { january:1,february:2,march:3,april:4,may:5,june:6,july:7,august:8,september:9,october:10,november:11,december:12 };
  var numberFmt = new Intl.NumberFormat('en-GB', { maximumFractionDigits: 0 });
  var currency = 'GBP';
  function currencyField() { return currency === 'GBP' ? MEAS_GBP : MEAS_LOCAL; }
  function fmtMoney(v) { if (v == null || v === '') return ''; var n = numberFmt.format(Math.round(v)); return currency === 'GBP' ? ('£' + n) : n; }
  function norm(v) { return v == null ? '' : String(v); }

  function el(tag, props, children) {
    var n = document.createElement(tag);
    if (props) Object.keys(props).forEach(function (k) { if (k === 'class') n.className = props[k]; else if (k === 'text') n.textContent = props[k]; else n.setAttribute(k, props[k]); });
    (children || []).forEach(function (c) { n.appendChild(c); });
    return n;
  }
  function headerFor(id) { for (var i = 0; i < COLS.length; i++) if (COLS[i].colId === id) return COLS[i].headerName; return id; }
  function sortVals(field, vals) {
    if (field === 'Year') vals.sort(function (a, b) { return (parseInt(b, 10) || 0) - (parseInt(a, 10) || 0); });
    else if (field === 'Month') vals.sort(function (a, b) { return (MONTHS[a.toLowerCase()] || 0) - (MONTHS[b.toLowerCase()] || 0); });
    else vals.sort(function (a, b) { return a.localeCompare(b); });
    return vals;
  }
  function closeAllMenus() { document.querySelectorAll('.menu, .filter-panel').forEach(function (p) { p.setAttribute('hidden', 'hidden'); }); }

  // real file download — works here because this PoC is NOT sandboxed
  function downloadBlob(filename, content, mime) {
    var blob = new Blob([content], { type: mime }), url = URL.createObjectURL(blob);
    var a = el('a', { href: url, download: filename }); document.body.appendChild(a); a.click();
    setTimeout(function () { document.body.removeChild(a); URL.revokeObjectURL(url); }, 1000);
  }

  function build(rows) {
    root.textContent = '';
    var gridApi, active = {};
    var STORE_KEY = 'cfgtbl-poc:v1:colstate';

    function valuesFor(field, respectSelf) {
      var out = {};
      rows.forEach(function (r) {
        for (var i = 0; i < FILTERS.length; i++) { var f = FILTERS[i]; if (f === field && !respectSelf) continue; var s = active[f]; if (s && s.size && !s.has(norm(r[f]))) return; }
        out[norm(r[field])] = true;
      });
      return sortVals(field, Object.keys(out));
    }

    // header
    root.appendChild(el('div', { class: 'appheader' }, [
      el('div', { class: 'brand-titles' }, [
        el('h1', { text: 'Performance — PoC (self-hosted)' }),
        el('div', { class: 'refreshed', text: 'Network-enabled extension · sample data · ' + new Date().toLocaleString() })
      ])
    ]));

    // filter bar
    var filterBar = el('div', { class: 'filterbar' });
    root.appendChild(filterBar);
    FILTERS.forEach(function (f) { filterBar.appendChild(makeFilter(f)); });
    var clearAll = el('button', { class: 'linkbtn', type: 'button', text: 'Clear all filters' });
    clearAll.addEventListener('click', function () { active = {}; FILTERS.forEach(updateFilterButton); gridApi.onFilterChanged(); });
    filterBar.appendChild(clearAll);
    var years = valuesFor('Year', true); if (years.length) active['Year'] = new Set([years[0]]);

    // toolbar
    var statusEl = el('span', { class: 'status' });
    var curSelect = el('select', { class: 'cur-select' });
    ['GBP', 'Local'].forEach(function (c) { var o = el('option', { value: c, text: c }); if (c === currency) o.setAttribute('selected', 'selected'); curSelect.appendChild(o); });
    var curWrap = el('label', { class: 'cur-wrap' }, [el('span', { text: 'Currency' }), curSelect]);

    var colBtn = el('button', { class: 'btn', text: 'Columns ▾', type: 'button' });
    var colPanel = el('div', { class: 'menu', hidden: 'hidden' });
    var colWrap = el('div', { class: 'menu-wrap' }, [colBtn, colPanel]);

    var dlBtn = el('button', { class: 'btn', text: 'Download ▾', type: 'button' });
    var dlPanel = el('div', { class: 'menu', hidden: 'hidden' });
    var dlCsv = el('button', { class: 'menu-item', type: 'button', text: 'CSV' });
    var dlXlsx = el('button', { class: 'menu-item', type: 'button', text: 'Excel (.xlsx)' });
    var dlPdf = el('button', { class: 'menu-item', type: 'button', text: 'PDF' });
    dlPanel.appendChild(dlCsv); dlPanel.appendChild(dlXlsx); dlPanel.appendChild(dlPdf);
    var dlWrap = el('div', { class: 'menu-wrap' }, [dlBtn, dlPanel]);

    var searchInput = el('input', { class: 'quick-search', type: 'search', placeholder: 'Search table…', 'aria-label': 'Search table' });

    var saveBtn = el('button', { class: 'btn primary', text: 'Save view', type: 'button' });
    var resetBtn = el('button', { class: 'btn', text: 'Reset', type: 'button' });

    root.appendChild(el('div', { class: 'toolbar' }, [
      el('div', { class: 'meta', text: rows.length + ' rows' }),
      searchInput, curWrap, colWrap, dlWrap, saveBtn, resetBtn, statusEl
    ]));

    // grid
    var gridEl = el('div', { class: 'grid ag-theme-quartz' });
    root.appendChild(gridEl);

    function filteredTotal() { var s = 0; gridApi.forEachNodeAfterFilter(function (n) { if (n.data) { var v = n.data[currencyField()]; if (typeof v === 'number') s += v; } }); return s; }
    function recomputeTotals() {
      if (!gridApi) return;
      var n = 0; gridApi.forEachNodeAfterFilter(function () { n++; });
      var vis = gridApi.getColumnState().filter(function (c) { return !c.hide; });
      var label = vis.length ? vis[0].colId : 'Specialist';
      var pinned = { 'Total Price': filteredTotal() }; if (label !== 'Total Price') pinned[label] = 'Total';
      gridApi.setGridOption('pinnedBottomRowData', [pinned]);
      statusEl.textContent = n + ' rows';
    }

    function colDefs() {
      return COLS.map(function (c) {
        var d = { colId: c.colId, headerName: c.headerName, width: c.width };
        if (c.totalPrice) { d.type = 'numericColumn'; d.filter = 'agNumberColumnFilter'; d.cellClass = 'bold-cell'; d.pinned = 'right'; d.valueGetter = function (p) { if (p.node && p.node.rowPinned) return p.data ? p.data['Total Price'] : null; return p.data ? p.data[currencyField()] : null; }; d.valueFormatter = function (p) { return fmtMoney(p.value); }; }
        else { d.field = c.field; if (c.yearSort) d.comparator = function (a, b) { return (parseInt(a, 10) || 0) - (parseInt(b, 10) || 0); }; if (c.monthSort) d.comparator = function (a, b) { return (MONTHS[String(a).toLowerCase()] || 0) - (MONTHS[String(b).toLowerCase()] || 0); }; }
        return d;
      });
    }

    gridApi = agGrid.createGrid(gridEl, {
      columnDefs: colDefs(), rowData: rows,
      defaultColDef: { sortable: true, filter: true, resizable: true },
      suppressDragLeaveHidesColumns: true, enableCellTextSelection: true,
      getRowStyle: function (p) { if (p.node.rowPinned) return { fontWeight: '700' }; },
      isExternalFilterPresent: function () { return FILTERS.some(function (f) { return active[f] && active[f].size; }); },
      doesExternalFilterPass: function (node) { for (var i = 0; i < FILTERS.length; i++) { var f = FILTERS[i], s = active[f]; if (s && s.size && !s.has(norm(node.data[f]))) return false; } return true; },
      onGridReady: function () { gridApi.onFilterChanged(); restoreSaved(); },
      onFilterChanged: recomputeTotals, onColumnVisible: recomputeTotals,
      onColumnMoved: function (e) { if (e.finished) recomputeTotals(); }
    });
    FILTERS.forEach(updateFilterButton);

    curSelect.addEventListener('change', function () { currency = curSelect.value; gridApi.refreshCells({ force: true, columns: ['Total Price'] }); recomputeTotals(); });

    // columns menu
    function rebuildCols() {
      colPanel.textContent = '';
      gridApi.getColumnState().forEach(function (c) {
        var cb = el('input', { type: 'checkbox' }); cb.checked = !c.hide;
        cb.addEventListener('change', function () { gridApi.applyColumnState({ state: [{ colId: c.colId, hide: !cb.checked }] }); recomputeTotals(); });
        colPanel.appendChild(el('label', { class: 'menu-check' }, [cb, el('span', { text: c.colId })]));
      });
    }
    colBtn.addEventListener('click', function (ev) { ev.stopPropagation(); var open = colPanel.hasAttribute('hidden'); closeAllMenus(); if (open) { rebuildCols(); colPanel.removeAttribute('hidden'); } });
    dlBtn.addEventListener('click', function (ev) { ev.stopPropagation(); var open = dlPanel.hasAttribute('hidden'); closeAllMenus(); if (open) dlPanel.removeAttribute('hidden'); });

    function currentView(raw) {
      var cols = gridApi.getColumnState().filter(function (c) { return !c.hide; });
      var headers = cols.map(function (c) { return headerFor(c.colId); });
      var numeric = cols.map(function (c) { return c.colId === 'Total Price'; });
      var body = [];
      gridApi.forEachNodeAfterFilter(function (node) {
        body.push(cols.map(function (c) { if (c.colId === 'Total Price') { var v = node.data[currencyField()]; return raw ? (v == null ? '' : v) : fmtMoney(v); } return norm(node.data[c.colId]); }));
      });
      var total = cols.map(function (c, i) { if (c.colId === 'Total Price') return raw ? filteredTotal() : fmtMoney(filteredTotal()); return i === 0 ? 'Total' : ''; });
      return { headers: headers, body: body, total: total, numeric: numeric };
    }
    dlCsv.addEventListener('click', function () {
      dlPanel.setAttribute('hidden', 'hidden');
      var v = currentView(true), esc = function (s) { s = (s == null ? '' : String(s)); return /[",\n\r]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s; };
      var lines = [v.headers.map(esc).join(',')]; v.body.forEach(function (r) { lines.push(r.map(esc).join(',')); }); lines.push(v.total.map(esc).join(','));
      downloadBlob('Performance_' + currency + '.csv', '﻿' + lines.join('\r\n'), 'text/csv;charset=utf-8;');
      statusEl.textContent = 'CSV downloaded ✓';
    });
    dlPdf.addEventListener('click', function () {
      dlPanel.setAttribute('hidden', 'hidden');
      if (!window.jspdf || !window.jspdf.jsPDF) { statusEl.textContent = 'PDF lib missing'; return; }
      var v = currentView(false), doc = new window.jspdf.jsPDF({ orientation: 'landscape', unit: 'pt', format: 'a4' });
      doc.setFontSize(13); doc.text('Performance (PoC)', 24, 26);
      doc.setFontSize(8); doc.setTextColor(120); doc.text('Currency: ' + currency + ' · ' + v.body.length + ' rows', 24, 40);
      doc.autoTable({ head: [v.headers], body: v.body, foot: [v.total], startY: 50, styles: { fontSize: 7, cellPadding: 2, overflow: 'linebreak' }, headStyles: { fillColor: [74, 74, 74], textColor: 255 }, footStyles: { fillColor: [235, 235, 235], textColor: 20, fontStyle: 'bold' }, margin: { left: 18, right: 18 } });
      doc.save('Performance_' + currency + '.pdf');
      statusEl.textContent = 'PDF downloaded ✓';
    });

    // save / restore layout via localStorage (works here — our own origin)
    saveBtn.addEventListener('click', function () {
      try { localStorage.setItem(STORE_KEY, JSON.stringify(gridApi.getColumnState())); statusEl.textContent = 'Layout saved to this browser ✓'; }
      catch (e) { statusEl.textContent = 'localStorage blocked: ' + (e.message || e); }
    });
    resetBtn.addEventListener('click', function () {
      try { localStorage.removeItem(STORE_KEY); } catch (e) {}
      gridApi.applyColumnState({ state: COLS.map(function (c) { return { colId: c.colId, hide: false, sort: null, width: c.width }; }), applyOrder: true });
      recomputeTotals(); statusEl.textContent = 'Reset to default';
    });
    function restoreSaved() {
      var s = null; try { s = localStorage.getItem(STORE_KEY); } catch (e) {}
      if (s) { try { gridApi.applyColumnState({ state: JSON.parse(s), applyOrder: true }); statusEl.textContent = 'Restored your saved layout ✓'; } catch (e) {} }
      recomputeTotals();
    }

    // Excel (.xlsx) download — real workbook, client-side (not sandboxed)
    dlXlsx.addEventListener('click', function () {
      dlPanel.setAttribute('hidden', 'hidden');
      if (!window.XLSX) { statusEl.textContent = 'Excel library missing'; return; }
      var v = currentView(true);
      var aoa = [v.headers].concat(v.body).concat([v.total]);
      var ws = XLSX.utils.aoa_to_sheet(aoa);
      var wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Performance');
      XLSX.writeFile(wb, 'Performance_' + currency + '.xlsx');
      statusEl.textContent = 'Excel downloaded ✓';
    });

    // quick search across the whole table
    searchInput.addEventListener('input', function () {
      gridApi.setGridOption('quickFilterText', searchInput.value);
      recomputeTotals();
    });

    // ---- filter dropdown ----
    function makeFilter(field) {
      var btn = el('button', { class: 'filterbtn', type: 'button' }); btn._field = field; updateBtn(btn, field, 0);
      var panel = el('div', { class: 'filter-panel', hidden: 'hidden' });
      var wrap = el('div', { class: 'filter-wrap' }, [btn, panel]);
      var pending = new Set();
      var search = el('input', { class: 'filter-search', type: 'text', placeholder: 'Search…' });
      var list = el('div', { class: 'filter-list' });
      var fa = el('button', { type: 'button', text: 'All' }), fc = el('button', { type: 'button', text: 'Clear' }), fap = el('button', { class: 'apply', type: 'button', text: 'Apply' });
      panel.appendChild(search); panel.appendChild(list); panel.appendChild(el('div', { class: 'filter-foot' }, [fa, fc, el('span', { class: 'spacer' }), fap]));
      function render() {
        list.textContent = ''; var q = search.value.trim().toLowerCase(); var vals = valuesFor(field, false);
        vals.forEach(function (val) { if (q && val.toLowerCase().indexOf(q) === -1) return; var cb = el('input', { type: 'checkbox' }); cb.checked = pending.has(val); cb.addEventListener('change', function () { if (cb.checked) pending.add(val); else pending['delete'](val); }); list.appendChild(el('label', { class: 'filter-opt' }, [cb, el('span', { text: val === '' ? '(blank)' : val })])); });
      }
      btn.addEventListener('click', function (ev) {
        ev.stopPropagation();
        if (panel.hasAttribute('hidden')) { closeAllMenus(); var cur = active[field]; pending = new Set(cur && cur.size ? Array.from(cur) : valuesFor(field, false)); search.value = ''; render(); panel.removeAttribute('hidden'); }
        else panel.setAttribute('hidden', 'hidden');
      });
      search.addEventListener('input', render);
      fa.addEventListener('click', function () { pending = new Set(valuesFor(field, false)); render(); });
      fc.addEventListener('click', function () { pending = new Set(); render(); });
      fap.addEventListener('click', function () { var all = valuesFor(field, false); if (pending.size === 0 || pending.size === all.length) delete active[field]; else active[field] = new Set(pending); updateFilterButton(field); panel.setAttribute('hidden', 'hidden'); gridApi.onFilterChanged(); });
      return wrap;
    }
    function updateFilterButton(field) { filterBar.querySelectorAll('.filterbtn').forEach(function (b) { if (b._field === field) updateBtn(b, field, active[field] ? active[field].size : 0); }); }
    function updateBtn(btn, field, count) { btn._field = field; btn.textContent = field + (count ? ' (' + count + ') ▾' : ' ▾'); btn.className = 'filterbtn' + (count ? ' on' : ''); }

    recomputeTotals();
  }

  document.addEventListener('click', function (ev) { document.querySelectorAll('.menu, .filter-panel').forEach(function (p) { var w = p.parentNode; if (w && !w.contains(ev.target)) p.setAttribute('hidden', 'hidden'); }); });

  // boot: render immediately (works standalone on GitHub Pages too), then
  // connect to Tableau in the background if we're inside it.
  function boot() {
    if (!window.agGrid) { root.textContent = 'Grid library failed to load.'; return; }
    build(rowsFromSample());
    try { if (window.tableau && tableau.extensions && tableau.extensions.initializeAsync) { tableau.extensions.initializeAsync().catch(function () {}); } } catch (e) {}
  }
  boot();
})();
