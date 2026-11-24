(function () {
  'use strict';

  // ===========================================================================
  // Configurable Table — network-enabled VIZ extension (dataset-agnostic)
  //   Reads the fields you place on THIS worksheet and builds the table:
  //     • dimensions  -> text columns + auto multi-select filters
  //     • measures    -> numeric columns + bold totals row
  //   Viewer-usable; real file downloads (CSV/Excel/PDF); per-user saved layout.
  //   Opened outside Tableau, it shows fictional sample data so the page previews.
  //   (The Configure panel — field/filter/total choices — comes next.)
  // ===========================================================================

  var root = document.getElementById('app');
  var numberFmt = new Intl.NumberFormat('en-GB', { maximumFractionDigits: 0 });
  var CURRENT = { setData: null };

  function el(tag, props, children) {
    var n = document.createElement(tag);
    if (props) Object.keys(props).forEach(function (k) { if (k === 'class') n.className = props[k]; else if (k === 'text') n.textContent = props[k]; else n.setAttribute(k, props[k]); });
    (children || []).forEach(function (c) { n.appendChild(c); });
    return n;
  }
  function norm(v) { return v == null ? '' : String(v); }
  function closeAllMenus() { document.querySelectorAll('.menu, .filter-panel').forEach(function (p) { p.setAttribute('hidden', 'hidden'); }); }
  function downloadBlob(filename, content, mime) {
    var b = new Blob([content], { type: mime }), url = URL.createObjectURL(b);
    var a = el('a', { href: url, download: filename }); document.body.appendChild(a); a.click();
    setTimeout(function () { document.body.removeChild(a); URL.revokeObjectURL(url); }, 1000);
  }
  function fmtNum(v) { if (v == null || v === '') return ''; return (typeof v === 'number') ? numberFmt.format(Math.round(v)) : String(v); }
  function renderError(msg) { root.textContent = ''; root.appendChild(el('div', { class: 'error', text: msg })); }

  // --- configuration (set by an editor; stored in the workbook; read by all) --
  function defaultConfig() { return { title: 'Configurable Table', subtitle: '', showHeader: true, bgColor: '', labels: {}, density: 'standard', zebra: true, zebraColor: '', align: {}, wrap: {}, numFmt: {} }; }
  function loadConfig() {
    var raw = null;
    try { if (window.tableau && tableau.extensions && tableau.extensions.settings) raw = tableau.extensions.settings.get('tableConfig'); } catch (e) {}
    if (raw == null) { try { raw = localStorage.getItem('cfgtbl-poc:config'); } catch (e) {} }
    var cfg = defaultConfig();
    if (raw) { try { var o = JSON.parse(raw); Object.keys(o).forEach(function (k) { cfg[k] = o[k]; }); } catch (e) {} }
    ['labels', 'align', 'wrap', 'numFmt'].forEach(function (k) { if (!cfg[k] || typeof cfg[k] !== 'object') cfg[k] = {}; });
    return cfg;
  }
  // Format a number with an Excel-style format string (via SheetJS SSF), e.g. "£"#,##0.00;-"£"#,##0.00
  function fmtMeasure(cfg, field, v) {
    if (v == null || v === '') return '';
    if (typeof v !== 'number') return String(v);
    var f = cfg.numFmt && cfg.numFmt[field];
    if (f && window.XLSX && XLSX.SSF) { try { return XLSX.SSF.format(f.replace(/\[[A-Za-z]+\]/g, ''), v); } catch (e) {} }
    return new Intl.NumberFormat('en-GB', { maximumFractionDigits: 0 }).format(Math.round(v));
  }
  function saveConfig(cfg) {
    var raw = JSON.stringify(cfg), done = Promise.resolve();
    try { if (window.tableau && tableau.extensions && tableau.extensions.settings && tableau.extensions.settings.saveAsync) { tableau.extensions.settings.set('tableConfig', raw); done = tableau.extensions.settings.saveAsync(); } } catch (e) {}
    try { localStorage.setItem('cfgtbl-poc:config', raw); } catch (e) {}
    return done;
  }
  // Configure is shown only to editors (authoring mode), or when previewed outside Tableau.
  function showConfigure() { try { if (!(window.tableau && tableau.extensions && tableau.extensions.environment)) return true; return tableau.extensions.environment.mode === 'authoring'; } catch (e) { return true; } }

  // --- map Tableau summary data -> { columns:[{field,type}], rows:[{field:val}] }
  function mapType(dt) { dt = (dt || '').toLowerCase(); if (dt.indexOf('int') >= 0 || dt.indexOf('float') >= 0 || dt === 'number' || dt === 'real') return 'number'; if (dt.indexOf('date') >= 0) return 'date'; return 'string'; }
  function fromDataTable(dt) {
    var columns = dt.columns.map(function (c) { return { field: c.fieldName, type: mapType(c.dataType), index: c.index }; });
    var rows = dt.data.map(function (r) {
      var o = {};
      columns.forEach(function (c) { var dv = r[c.index]; var v = (dv && dv.nativeValue !== undefined && dv.nativeValue !== null) ? dv.nativeValue : (dv ? dv.value : null); o[c.field] = v; });
      return o;
    });
    return { columns: columns, rows: rows };
  }
  function sampleModel() {
    var columns = [
      { field: 'Specialist', type: 'string' }, { field: 'Agency', type: 'string' }, { field: 'Advertiser Group', type: 'string' },
      { field: 'Campaign Name', type: 'string' }, { field: 'Country', type: 'string' }, { field: 'Year', type: 'string' }, { field: 'Month', type: 'string' },
      { field: 'Total Price', type: 'number' }
    ];
    var S = [
      ['Team North', 'Media Agency A', 'Acme Retail', 'Acme Spring Brand Push', 'UK', '2024', 'January', 12000],
      ['Team South', 'Media Agency B', 'Globex', 'Globex Q3 Network Deal', 'NL', '2025', 'July', 8500],
      ['Trading Desk', 'Media Agency A', 'Globex', 'Globex Always-On W2-3', 'UK', '2026', 'January', 4500],
      ['Direct', 'Direct', 'Initech', 'Initech Full-Year Panel', 'SE', '2023', 'February', 3000],
      ['Team North', 'Media Agency C', 'Umbrella Foods', 'Umbrella Winter Loop', 'DK', '2024', 'December', 21000],
      ['Direct', 'Direct', 'Soylent Corp', 'Soylent 2023 Programme', 'SE', '2023', 'January', 6000],
      ['Team South', 'Media Agency C', 'Hooli', 'Hooli City Centre Takeover', 'DK', '2024', 'July', 9000],
      ['Direct', 'Trading Desk', 'Stark Industries', 'Stark Loop Burst', 'DK', '2025', 'June', 11000],
      ['Team North', 'Media Agency D', 'Wayne Enterprises', 'Wayne EV Launch', 'UK', '2026', 'March', 63000],
      ['Team South', 'Media Agency B', 'Wonka Co', 'Wonka Fusion National', 'SE', '2026', 'February', 110000],
      ['Team North', 'Media Agency E', 'Cyberdyne', 'Cyberdyne Christmas Loop', 'DE', '2025', 'November', 49000],
      ['Team South', 'Media Agency D', 'Acme Retail', 'Acme Winter Sale Network', 'NL', '2025', 'January', 27000]
    ];
    var rows = S.map(function (r) { var o = {}; columns.forEach(function (c, i) { o[c.field] = r[i]; }); return o; });
    return { columns: columns, rows: rows };
  }

  // --- main ------------------------------------------------------------------
  function build(model) {
    root.textContent = '';
    var columns = model.columns.slice();
    var rows = model.rows.slice();
    var measures = columns.filter(function (c) { return c.type === 'number'; }).map(function (c) { return c.field; });
    var FILTERS = columns.filter(function (c) { return c.type !== 'number'; }).map(function (c) { return c.field; });
    var firstField = columns.length ? columns[0].field : '';
    var gridApi, active = {};
    var config = loadConfig();
    function labelFor(f) { return (config.labels && config.labels[f]) || f; }

    function valuesFor(field, respectSelf) {
      var out = {};
      rows.forEach(function (r) {
        for (var i = 0; i < FILTERS.length; i++) { var f = FILTERS[i]; if (f === field && !respectSelf) continue; var s = active[f]; if (s && s.size && !s.has(norm(r[f]))) return; }
        out[norm(r[field])] = true;
      });
      return Object.keys(out).sort(function (a, b) { var na = parseFloat(a), nb = parseFloat(b); if (!isNaN(na) && !isNaN(nb)) return na - nb; return a.localeCompare(b); });
    }

    // header (configurable: title, subtitle, or hidden entirely)
    function renderHeader() {
      var existing = root.querySelector('.appheader'); if (existing) existing.parentNode.removeChild(existing);
      if (!config.showHeader) return;
      var kids = [el('h1', { text: config.title || 'Table' })];
      if (config.subtitle) kids.push(el('div', { class: 'refreshed', text: config.subtitle }));
      root.insertBefore(el('div', { class: 'appheader' }, [el('div', { class: 'brand-titles' }, kids)]), root.firstChild);
    }
    renderHeader();

    // filter bar (one per dimension; wraps over rows)
    var filterBar = el('div', { class: 'filterbar' });
    root.appendChild(filterBar);
    FILTERS.forEach(function (f) { filterBar.appendChild(makeFilter(f)); });
    var clearAll = el('button', { class: 'linkbtn', type: 'button', text: 'Clear all filters' });
    clearAll.addEventListener('click', function () { active = {}; FILTERS.forEach(updateFilterButton); gridApi.onFilterChanged(); });
    filterBar.appendChild(clearAll);

    // action toolbar
    var metaEl = el('div', { class: 'meta' });
    var statusEl = el('span', { class: 'status' });
    var searchInput = el('input', { class: 'quick-search', type: 'search', placeholder: 'Search table…', 'aria-label': 'Search table' });
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
    var saveBtn = el('button', { class: 'btn primary', text: 'Save view', type: 'button' });
    var resetBtn = el('button', { class: 'btn', text: 'Reset', type: 'button' });
    var toolbarKids = [metaEl, statusEl, searchInput, colWrap, dlWrap, saveBtn, resetBtn];
    if (showConfigure()) { var configBtn = el('button', { class: 'btn', type: 'button', text: '⚙ Configure' }); configBtn.addEventListener('click', openConfigure); toolbarKids.push(configBtn); }
    root.appendChild(el('div', { class: 'toolbar' }, toolbarKids));

    function applyBackground() { var c = config.bgColor || ''; root.style.background = c; var fb = root.querySelector('.filterbar'); if (fb) fb.style.background = c; var tb = root.querySelector('.toolbar'); if (tb) tb.style.background = c; }
    function applyDensity() { if (!gridApi) return; var h = config.density === 'compact' ? 28 : config.density === 'comfortable' ? 50 : 40; gridApi.setGridOption('rowHeight', h); gridApi.setGridOption('headerHeight', h); if (gridApi.resetRowHeights) gridApi.resetRowHeights(); }
    function applyZebra() { if (!config.zebra) gridEl.style.setProperty('--ag-odd-row-background-color', 'transparent'); else if (config.zebraColor) gridEl.style.setProperty('--ag-odd-row-background-color', config.zebraColor); else gridEl.style.removeProperty('--ag-odd-row-background-color'); }
    function applyConfig() { renderHeader(); applyBackground(); applyDensity(); applyZebra(); FILTERS.forEach(updateFilterButton); if (gridApi) gridApi.setGridOption('columnDefs', colDefs()); }
    function openConfigure() {
      var backdrop = el('div', { class: 'modal-backdrop' });
      var titleIn = el('input', { type: 'text', class: 'cfg-input' }); titleIn.value = config.title || '';
      var subIn = el('input', { type: 'text', class: 'cfg-input' }); subIn.value = config.subtitle || '';
      var showHdr = el('input', { type: 'checkbox' }); if (config.showHeader) showHdr.checked = true;
      var bgIn = el('input', { type: 'text', class: 'cfg-input', placeholder: '#ffffff (blank = default)' }); bgIn.value = config.bgColor || '';
      var bgSwatch = el('span', { class: 'cfg-swatch' });
      function updSwatch() { bgSwatch.style.background = (bgIn.value.trim() || '#ffffff'); }
      bgIn.addEventListener('input', updSwatch); updSwatch();
      var densitySel = el('select', { class: 'cfg-input' }); [['standard', 'Standard'], ['compact', 'Compact'], ['comfortable', 'Comfortable']].forEach(function (o) { var op = el('option', { value: o[0], text: o[1] }); if (config.density === o[0]) op.setAttribute('selected', 'selected'); densitySel.appendChild(op); });
      var zebraChk = el('input', { type: 'checkbox' }); if (config.zebra) zebraChk.checked = true;
      var zebraIn = el('input', { type: 'text', class: 'cfg-input', placeholder: '#f6f8fb (optional)' }); zebraIn.value = config.zebraColor || '';
      var zebraSw = el('span', { class: 'cfg-swatch' }); function updZ() { zebraSw.style.background = (zebraIn.value.trim() || '#f6f8fb'); } zebraIn.addEventListener('input', updZ); updZ();

      var labelInputs = {}, alignSel = {}, wrapChk = {}, fmtIn = {};
      var colList = el('div', { class: 'cfg-cols' });
      colList.appendChild(el('div', { class: 'cfg-colrow cfg-colhead' }, [el('span', { text: 'Field' }), el('span', { text: 'Label' }), el('span', { text: 'Align' }), el('span', { text: 'Wrap' }), el('span', { text: 'Number format' })]));
      columns.forEach(function (c) {
        var isMeas = measures.indexOf(c.field) >= 0;
        var lab = el('input', { type: 'text', class: 'cfg-input', placeholder: c.field }); lab.value = config.labels[c.field] || ''; labelInputs[c.field] = lab;
        var al = el('select', { class: 'cfg-input' }); [['', 'Auto'], ['left', 'Left'], ['center', 'Centre'], ['right', 'Right']].forEach(function (o) { var op = el('option', { value: o[0], text: o[1] }); if ((config.align[c.field] || '') === o[0]) op.setAttribute('selected', 'selected'); al.appendChild(op); }); alignSel[c.field] = al;
        var wr = el('input', { type: 'checkbox' }); if (config.wrap[c.field]) wr.checked = true; wrapChk[c.field] = wr;
        var fm = el('input', { type: 'text', class: 'cfg-input', placeholder: isMeas ? '#,##0.00' : '' }); fm.value = config.numFmt[c.field] || ''; if (!isMeas) fm.disabled = true; fmtIn[c.field] = fm;
        colList.appendChild(el('div', { class: 'cfg-colrow' }, [el('span', { class: 'cfg-fieldname', text: c.field }), lab, al, wr, fm]));
      });
      var saveB = el('button', { class: 'btn primary', type: 'button', text: 'Save' });
      var cancelB = el('button', { class: 'btn', type: 'button', text: 'Cancel' });
      saveB.addEventListener('click', function () {
        config.title = titleIn.value; config.subtitle = subIn.value; config.showHeader = !!showHdr.checked; config.bgColor = bgIn.value.trim();
        config.density = densitySel.value; config.zebra = !!zebraChk.checked; config.zebraColor = zebraIn.value.trim();
        config.labels = {}; config.align = {}; config.wrap = {}; config.numFmt = {};
        columns.forEach(function (c) {
          var f = c.field;
          var lv = labelInputs[f].value.trim(); if (lv) config.labels[f] = lv;
          var av = alignSel[f].value; if (av) config.align[f] = av;
          if (wrapChk[f].checked) config.wrap[f] = true;
          var fv = fmtIn[f].value.trim(); if (fv) config.numFmt[f] = fv;
        });
        saveConfig(config); applyConfig(); statusEl.textContent = 'Configuration saved ✓'; document.body.removeChild(backdrop);
      });
      cancelB.addEventListener('click', function () { document.body.removeChild(backdrop); });
      backdrop.appendChild(el('div', { class: 'modal cfg-modal' }, [
        el('div', { class: 'modal-head' }, [el('strong', { text: 'Configure table' })]),
        el('div', { class: 'modal-body' }, [
          el('div', { class: 'cfg-row' }, [el('span', { class: 'cfg-field', text: 'Title' }), titleIn]),
          el('div', { class: 'cfg-row' }, [el('span', { class: 'cfg-field', text: 'Subtitle' }), subIn]),
          el('label', { class: 'cfg-row cfg-check' }, [showHdr, el('span', { text: 'Show header (title + subtitle)' })]),
          el('div', { class: 'cfg-row' }, [el('span', { class: 'cfg-field', text: 'Background colour' }), bgIn, bgSwatch]),
          el('div', { class: 'cfg-row' }, [el('span', { class: 'cfg-field', text: 'Row density' }), densitySel]),
          el('label', { class: 'cfg-row cfg-check' }, [zebraChk, el('span', { text: 'Zebra striping' })]),
          el('div', { class: 'cfg-row' }, [el('span', { class: 'cfg-field', text: 'Zebra colour' }), zebraIn, zebraSw]),
          el('div', { class: 'cfg-sec', text: 'Columns (format strings like "£"#,##0.00;[Red]-"£"#,##0.00)' }),
          colList
        ]),
        el('div', { class: 'modal-foot' }, [cancelB, saveB])
      ]));
      document.body.appendChild(backdrop);
    }

    // grid
    var gridEl = el('div', { class: 'grid ag-theme-quartz' });
    root.appendChild(gridEl);

    function colDefs() {
      return columns.map(function (c) {
        var field = c.field, isMeas = measures.indexOf(field) >= 0;
        var align = config.align[field] || (isMeas ? 'right' : 'left');
        var wrap = !!config.wrap[field];
        var d = { colId: field, headerName: labelFor(field), field: field, wrapText: wrap, autoHeight: wrap, headerClass: 'hdr-' + align };
        if (isMeas) { d.type = 'numericColumn'; d.filter = 'agNumberColumnFilter'; d.valueFormatter = function (p) { return fmtMeasure(config, field, p.value); }; }
        d.cellStyle = function (p) {
          var st = { textAlign: align, whiteSpace: wrap ? 'normal' : 'nowrap' };
          if (p.node && p.node.rowPinned) st.fontWeight = '700';
          if (isMeas) { var f = config.numFmt[field]; if (f && /\[red\]/i.test(f) && typeof p.value === 'number' && p.value < 0) st.color = '#b00020'; }
          return st;
        };
        return d;
      });
    }
    function recompute() {
      if (!gridApi) return;
      var sums = {}; measures.forEach(function (m) { sums[m] = 0; });
      var n = 0;
      gridApi.forEachNodeAfterFilter(function (node) { n++; measures.forEach(function (m) { var v = node.data[m]; if (typeof v === 'number') sums[m] += v; }); });
      if (measures.length) {
        var vis = gridApi.getColumnState().filter(function (c) { return !c.hide; });
        var label = vis.length ? vis[0].colId : firstField;
        var pinned = {}; if (measures.indexOf(label) < 0) pinned[label] = 'Total';
        measures.forEach(function (m) { pinned[m] = sums[m]; });
        gridApi.setGridOption('pinnedBottomRowData', [pinned]);
      }
      metaEl.textContent = (n === rows.length) ? (rows.length + ' Total rows') : ('Showing ' + n + ' of ' + rows.length + ' rows');
    }

    gridApi = agGrid.createGrid(gridEl, {
      columnDefs: colDefs(), rowData: rows,
      defaultColDef: { sortable: true, filter: true, resizable: true },
      autoSizeStrategy: { type: 'fitCellContents' },
      suppressDragLeaveHidesColumns: true, enableCellTextSelection: true,
      getRowStyle: function (p) { if (p.node.rowPinned) return { fontWeight: '700' }; },
      isExternalFilterPresent: function () { return FILTERS.some(function (f) { return active[f] && active[f].size; }); },
      doesExternalFilterPass: function (node) { for (var i = 0; i < FILTERS.length; i++) { var f = FILTERS[i], s = active[f]; if (s && s.size && !s.has(norm(node.data[f]))) return false; } return true; },
      onGridReady: function () { gridApi.onFilterChanged(); restoreSaved(); },
      onFilterChanged: recompute, onColumnVisible: recompute,
      onColumnMoved: function (e) { if (e.finished) recompute(); }
    });
    FILTERS.forEach(updateFilterButton);

    // columns menu
    function rebuildCols() {
      colPanel.textContent = '';
      gridApi.getColumnState().forEach(function (c) {
        var cb = el('input', { type: 'checkbox' }); cb.checked = !c.hide;
        cb.addEventListener('change', function () { gridApi.applyColumnState({ state: [{ colId: c.colId, hide: !cb.checked }] }); recompute(); });
        colPanel.appendChild(el('label', { class: 'menu-check' }, [cb, el('span', { text: config.labels[c.colId] || c.colId })]));
      });
    }
    colBtn.addEventListener('click', function (ev) { ev.stopPropagation(); var open = colPanel.hasAttribute('hidden'); closeAllMenus(); if (open) { rebuildCols(); colPanel.removeAttribute('hidden'); } });
    dlBtn.addEventListener('click', function (ev) { ev.stopPropagation(); var open = dlPanel.hasAttribute('hidden'); closeAllMenus(); if (open) dlPanel.removeAttribute('hidden'); });

    function currentView(raw) {
      var cols = gridApi.getColumnState().filter(function (c) { return !c.hide; });
      var headers = cols.map(function (c) { return labelFor(c.colId); });
      var numeric = cols.map(function (c) { return measures.indexOf(c.colId) >= 0; });
      var sums = {}; measures.forEach(function (m) { sums[m] = 0; });
      var body = [];
      gridApi.forEachNodeAfterFilter(function (node) {
        body.push(cols.map(function (c) { var v = node.data[c.colId]; if (measures.indexOf(c.colId) >= 0) { if (typeof v === 'number') sums[c.colId] += v; return raw ? (v == null ? '' : v) : fmtMeasure(config, c.colId, v); } return norm(v); }));
      });
      var total = cols.map(function (c, i) { if (measures.indexOf(c.colId) >= 0) return raw ? sums[c.colId] : fmtMeasure(config, c.colId, sums[c.colId]); return i === 0 ? 'Total' : ''; });
      return { headers: headers, body: body, total: total, numeric: numeric };
    }
    dlCsv.addEventListener('click', function () {
      dlPanel.setAttribute('hidden', 'hidden');
      var v = currentView(true), esc = function (s) { s = (s == null ? '' : String(s)); return /[",\n\r]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s; };
      var lines = [v.headers.map(esc).join(',')]; v.body.forEach(function (r) { lines.push(r.map(esc).join(',')); });
      if (measures.length) lines.push(v.total.map(esc).join(','));
      downloadBlob('Table.csv', '﻿' + lines.join('\r\n'), 'text/csv;charset=utf-8;'); statusEl.textContent = 'CSV downloaded ✓';
    });
    dlXlsx.addEventListener('click', function () {
      dlPanel.setAttribute('hidden', 'hidden');
      if (!window.XLSX) { statusEl.textContent = 'Excel library missing'; return; }
      var v = currentView(true), aoa = [v.headers].concat(v.body); if (measures.length) aoa.push(v.total);
      var ws = XLSX.utils.aoa_to_sheet(aoa), wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Table'); XLSX.writeFile(wb, 'Table.xlsx'); statusEl.textContent = 'Excel downloaded ✓';
    });
    dlPdf.addEventListener('click', function () {
      dlPanel.setAttribute('hidden', 'hidden');
      if (!window.jspdf || !window.jspdf.jsPDF) { statusEl.textContent = 'PDF library missing'; return; }
      var v = currentView(false), doc = new window.jspdf.jsPDF({ orientation: 'landscape', unit: 'pt', format: 'a4' });
      doc.setFontSize(13); doc.text('Table', 24, 26); doc.setFontSize(8); doc.setTextColor(120); doc.text(v.body.length + ' rows', 24, 40);
      doc.autoTable({ head: [v.headers], body: v.body, foot: measures.length ? [v.total] : undefined, startY: 50, styles: { fontSize: 7, cellPadding: 2, overflow: 'linebreak' }, headStyles: { fillColor: [74, 74, 74], textColor: 255 }, footStyles: { fillColor: [235, 235, 235], textColor: 20, fontStyle: 'bold' }, margin: { left: 18, right: 18 } });
      doc.save('Table.pdf'); statusEl.textContent = 'PDF downloaded ✓';
    });

    searchInput.addEventListener('input', function () { gridApi.setGridOption('quickFilterText', searchInput.value); recompute(); });

    var STORE_KEY = 'cfgtbl-poc:v2:' + (columns.map(function (c) { return c.field; }).join('|'));
    saveBtn.addEventListener('click', function () { try { localStorage.setItem(STORE_KEY, JSON.stringify(gridApi.getColumnState())); statusEl.textContent = 'Layout saved to this browser ✓'; } catch (e) { statusEl.textContent = 'Could not save: ' + (e.message || e); } });
    resetBtn.addEventListener('click', function () { try { localStorage.removeItem(STORE_KEY); } catch (e) {} gridApi.applyColumnState({ state: columns.map(function (c) { return { colId: c.field, hide: false, sort: null }; }), applyOrder: true }); gridApi.autoSizeAllColumns && gridApi.autoSizeAllColumns(); recompute(); statusEl.textContent = 'Reset to default'; });
    function restoreSaved() { var s = null; try { s = localStorage.getItem(STORE_KEY); } catch (e) {} if (s) { try { gridApi.applyColumnState({ state: JSON.parse(s), applyOrder: true }); statusEl.textContent = 'Restored your saved layout ✓'; } catch (e) {} } recompute(); }

    // live refresh when the worksheet's data changes (native filters etc.)
    CURRENT.setData = function (m) { rows = m.rows.slice(); gridApi.setGridOption('rowData', rows); FILTERS.forEach(updateFilterButton); recompute(); };

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
      function render() { list.textContent = ''; var q = search.value.trim().toLowerCase(); valuesFor(field, false).forEach(function (val) { if (q && val.toLowerCase().indexOf(q) === -1) return; var cb = el('input', { type: 'checkbox' }); cb.checked = pending.has(val); cb.addEventListener('change', function () { if (cb.checked) pending.add(val); else pending['delete'](val); }); list.appendChild(el('label', { class: 'filter-opt' }, [cb, el('span', { text: val === '' ? '(blank)' : val })])); }); }
      btn.addEventListener('click', function (ev) { ev.stopPropagation(); if (panel.hasAttribute('hidden')) { closeAllMenus(); var cur = active[field]; pending = new Set(cur && cur.size ? Array.from(cur) : valuesFor(field, false)); search.value = ''; render(); panel.removeAttribute('hidden'); } else panel.setAttribute('hidden', 'hidden'); });
      search.addEventListener('input', render);
      fa.addEventListener('click', function () { pending = new Set(valuesFor(field, false)); render(); });
      fc.addEventListener('click', function () { pending = new Set(); render(); });
      fap.addEventListener('click', function () { var all = valuesFor(field, false); if (pending.size === 0 || pending.size === all.length) delete active[field]; else active[field] = new Set(pending); updateFilterButton(field); panel.setAttribute('hidden', 'hidden'); gridApi.onFilterChanged(); });
      return wrap;
    }
    function updateFilterButton(field) { filterBar.querySelectorAll('.filterbtn').forEach(function (b) { if (b._field === field) updateBtn(b, field, active[field] ? active[field].size : 0); }); }
    function updateBtn(btn, field, count) { btn._field = field; btn.textContent = labelFor(field) + (count ? ' (' + count + ') ▾' : ' ▾'); btn.className = 'filterbtn' + (count ? ' on' : ''); }

    applyBackground(); applyDensity(); applyZebra();
    recompute();
  }

  document.addEventListener('click', function (ev) { document.querySelectorAll('.menu, .filter-panel').forEach(function (p) { var w = p.parentNode; if (w && !w.contains(ev.target)) p.setAttribute('hidden', 'hidden'); }); });

  // --- boot ------------------------------------------------------------------
  function readWorksheet(ws) {
    return ws.getSummaryDataReaderAsync().then(function (reader) {
      return reader.getAllPagesAsync().then(function (dt) { try { reader.releaseAsync(); } catch (e) {} return fromDataTable(dt); });
    });
  }
  function boot() {
    if (!window.agGrid) { root.textContent = 'Grid library failed to load.'; return; }
    var inTableau = !!(window.tableau && tableau.extensions && tableau.extensions.initializeAsync);
    // Outside Tableau (raw page preview) show fictional sample data.
    if (!inTableau) { build(sampleModel()); return; }
    tableau.extensions.initializeAsync().then(function () {
      var ws = tableau.extensions.worksheetContent && tableau.extensions.worksheetContent.worksheet;
      if (!ws || !ws.getSummaryDataReaderAsync) { renderError('Add this extension to a worksheet that has fields on it.'); return; }
      readWorksheet(ws).then(function (m) {
        build(m);
        try { ws.addEventListener(tableau.TableauEventType.SummaryDataChanged, function () { readWorksheet(ws).then(function (m2) { if (CURRENT.setData) CURRENT.setData(m2); }); }); } catch (e) {}
      }).catch(function () {
        // In Tableau: DON'T show fake data — show a friendly request-access message.
        renderError('You don’t currently have permission to view this table. Please request access to this dashboard and its data source from your Tableau administrator, then reload.');
      });
    }).catch(function (err) { renderError('Extension failed to initialize: ' + ((err && err.message) || err)); });
  }
  boot();
})();
