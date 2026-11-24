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
  function defaultConfig() {
    return {
      title: 'Configurable Table', subtitle: '', showHeader: true, bgColor: '',
      labels: {}, density: 'standard', zebra: true, zebraColor: '',
      align: {}, wrap: {}, numFmt: {}, order: [], hidden: {}, pin: {},
      sortField: '', sortDir: 'asc',
      totals: { enabled: true, measures: null, bold: true, position: 'bottom', label: 'Total' },
      // filters
      filterType: {}, filterSort: {}, filterDefault: {}, filterHidden: {}, cascade: true, applyButton: true,
      // interaction
      exports: { csv: true, excel: true, pdf: true }, fileName: 'Table', fileStamp: false, saveView: true, quickSearch: true
    };
  }
  function loadConfig() {
    var raw = null;
    try { if (window.tableau && tableau.extensions && tableau.extensions.settings) raw = tableau.extensions.settings.get('tableConfig'); } catch (e) {}
    if (raw == null) { try { raw = localStorage.getItem('cfgtbl-poc:config'); } catch (e) {} }
    var cfg = defaultConfig();
    if (raw) { try { var o = JSON.parse(raw); Object.keys(o).forEach(function (k) { cfg[k] = o[k]; }); } catch (e) {} }
    ['labels', 'align', 'wrap', 'numFmt', 'hidden', 'pin', 'filterType', 'filterSort', 'filterDefault', 'filterHidden'].forEach(function (k) { if (!cfg[k] || typeof cfg[k] !== 'object') cfg[k] = {}; });
    if (!Array.isArray(cfg.order)) cfg.order = [];
    if (!cfg.totals || typeof cfg.totals !== 'object') cfg.totals = defaultConfig().totals;
    if (!cfg.exports || typeof cfg.exports !== 'object') cfg.exports = defaultConfig().exports;
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

    // A filter value in `active` is one of: a Set (membership), {kind:'search',q}, {kind:'range',min,max}, {kind:'date',from,to}
    function filterActive(a) { if (!a) return false; if (a instanceof Set) return a.size > 0; if (a.kind === 'search') return !!a.q; if (a.kind === 'range') return a.min != null || a.max != null; if (a.kind === 'date') return !!a.from || !!a.to; return false; }
    function rowPasses(field, val) {
      var a = active[field]; if (!filterActive(a)) return true; var s = norm(val);
      if (a instanceof Set) return a.has(s);
      if (a.kind === 'search') return s.toLowerCase().indexOf(a.q.toLowerCase()) >= 0;
      if (a.kind === 'range') { var n = parseFloat(s); if (isNaN(n)) return false; if (a.min != null && n < a.min) return false; if (a.max != null && n > a.max) return false; return true; }
      if (a.kind === 'date') { var d = Date.parse(s); if (isNaN(d)) return false; if (a.from && d < Date.parse(a.from)) return false; if (a.to && d > Date.parse(a.to) + 86399999) return false; return true; }
      return true;
    }
    var optCounts = {};
    function valuesFor(field) {
      var counts = {};
      rows.forEach(function (r) {
        if (config.cascade !== false) { for (var i = 0; i < FILTERS.length; i++) { var f = FILTERS[i]; if (f === field) continue; if (!rowPasses(f, r[f])) return; } }
        var k = norm(r[field]); counts[k] = (counts[k] || 0) + 1;
      });
      optCounts[field] = counts;
      var keys = Object.keys(counts), sort = (config.filterSort && config.filterSort[field]) || 'az';
      keys.sort(function (a, b) {
        var na = parseFloat(a), nb = parseFloat(b), bothNum = !isNaN(na) && !isNaN(nb);
        if (sort === 'count') return (counts[b] - counts[a]) || (bothNum ? na - nb : a.localeCompare(b));
        if (sort === 'za') return bothNum ? nb - na : b.localeCompare(a);
        return bothNum ? na - nb : a.localeCompare(b);
      });
      return keys;
    }
    function rawValues(field) { var o = {}; rows.forEach(function (r) { o[norm(r[field])] = true; }); return Object.keys(o); }
    function latestOf(vals) {
      if (!vals.length) return null;
      var nums = vals.map(parseFloat); if (nums.every(function (n) { return !isNaN(n); })) { var mi = 0; for (var i = 1; i < nums.length; i++) if (nums[i] > nums[mi]) mi = i; return vals[mi]; }
      var dates = vals.map(Date.parse); if (dates.every(function (d) { return !isNaN(d); })) { var di = 0; for (var j = 1; j < dates.length; j++) if (dates[j] > dates[di]) di = j; return vals[di]; }
      return vals.slice().sort()[vals.length - 1];
    }
    function initDefaults() {
      FILTERS.forEach(function (f) {
        var d = config.filterDefault && config.filterDefault[f]; if (!d) return;
        if (d === 'latest') { var mx = latestOf(rawValues(f)); if (mx != null) active[f] = new Set([mx]); }
        else if (Array.isArray(d) && d.length) { active[f] = new Set(d.map(norm)); }
      });
    }
    initDefaults();

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
    FILTERS.forEach(function (f) { if (config.filterHidden && config.filterHidden[f]) return; filterBar.appendChild(makeFilter(f)); });
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
    var exp = config.exports || {}; var anyExp = (exp.csv !== false) || (exp.excel !== false) || (exp.pdf !== false);
    var dlBtn = el('button', { class: 'btn', text: 'Download ▾', type: 'button' });
    var dlPanel = el('div', { class: 'menu', hidden: 'hidden' });
    var dlCsv = el('button', { class: 'menu-item', type: 'button', text: 'CSV' });
    var dlXlsx = el('button', { class: 'menu-item', type: 'button', text: 'Excel (.xlsx)' });
    var dlPdf = el('button', { class: 'menu-item', type: 'button', text: 'PDF' });
    if (exp.csv !== false) dlPanel.appendChild(dlCsv);
    if (exp.excel !== false) dlPanel.appendChild(dlXlsx);
    if (exp.pdf !== false) dlPanel.appendChild(dlPdf);
    var dlWrap = el('div', { class: 'menu-wrap' }, [dlBtn, dlPanel]);
    var saveBtn = el('button', { class: 'btn primary', text: 'Save view', type: 'button' });
    var resetBtn = el('button', { class: 'btn', text: 'Reset', type: 'button' });
    var toolbarKids = [metaEl, statusEl, el('span', { class: 'tb-spacer' })];
    if (config.quickSearch !== false) toolbarKids.push(searchInput);
    toolbarKids.push(colWrap);
    if (anyExp) toolbarKids.push(dlWrap);
    if (config.saveView !== false) { toolbarKids.push(saveBtn); toolbarKids.push(resetBtn); }
    if (showConfigure()) { var configBtn = el('button', { class: 'btn', type: 'button', text: '⚙ Configure' }); configBtn.addEventListener('click', openConfigure); toolbarKids.push(configBtn); }
    root.appendChild(el('div', { class: 'toolbar' }, toolbarKids));

    function applyBackground() { var c = config.bgColor || ''; root.style.background = c; var fb = root.querySelector('.filterbar'); if (fb) fb.style.background = c; var tb = root.querySelector('.toolbar'); if (tb) tb.style.background = c; }
    function applyDensity() { if (!gridApi) return; var h = config.density === 'compact' ? 28 : config.density === 'comfortable' ? 50 : 40; gridApi.setGridOption('rowHeight', h); gridApi.setGridOption('headerHeight', h); if (gridApi.resetRowHeights) gridApi.resetRowHeights(); }
    function applyZebra() { if (!config.zebra) gridEl.style.setProperty('--ag-odd-row-background-color', 'transparent'); else if (config.zebraColor) gridEl.style.setProperty('--ag-odd-row-background-color', config.zebraColor); else gridEl.style.removeProperty('--ag-odd-row-background-color'); }
    function applyConfig() {
      renderHeader(); applyBackground(); applyDensity(); applyZebra(); FILTERS.forEach(updateFilterButton);
      if (gridApi) {
        gridApi.setGridOption('columnDefs', colDefs());
        if (config.sortField) gridApi.applyColumnState({ state: [{ colId: config.sortField, sort: (config.sortDir === 'desc' ? 'desc' : 'asc') }], defaultState: { sort: null } });
        else gridApi.applyColumnState({ defaultState: { sort: null } });
        recompute();
      }
    }
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

      // default sort
      var sortFieldSel = el('select', { class: 'cfg-input' }); sortFieldSel.appendChild(el('option', { value: '', text: '(none)' }));
      orderedColumns().forEach(function (c) { var op = el('option', { value: c.field, text: labelFor(c.field) }); if (config.sortField === c.field) op.setAttribute('selected', 'selected'); sortFieldSel.appendChild(op); });
      var sortDirSel = el('select', { class: 'cfg-input' }); [['asc', 'Ascending'], ['desc', 'Descending']].forEach(function (o) { var op = el('option', { value: o[0], text: o[1] }); if (config.sortDir === o[0]) op.setAttribute('selected', 'selected'); sortDirSel.appendChild(op); });

      // totals
      var totEnable = el('input', { type: 'checkbox' }); if (config.totals.enabled) totEnable.checked = true;
      var totBold = el('input', { type: 'checkbox' }); if (config.totals.bold) totBold.checked = true;
      var totPos = el('select', { class: 'cfg-input' }); [['bottom', 'Bottom'], ['top', 'Top']].forEach(function (o) { var op = el('option', { value: o[0], text: o[1] }); if (config.totals.position === o[0]) op.setAttribute('selected', 'selected'); totPos.appendChild(op); });
      var totLabel = el('input', { type: 'text', class: 'cfg-input', placeholder: 'Total' }); totLabel.value = config.totals.label || '';
      var totMeasChk = {}; var totMeasWrap = el('div', { class: 'cfg-measchk' });
      measures.forEach(function (m) { var cb = el('input', { type: 'checkbox' }); cb.checked = (!config.totals.measures || config.totals.measures.indexOf(m) >= 0); totMeasChk[m] = cb; totMeasWrap.appendChild(el('label', { class: 'cfg-measitem' }, [cb, el('span', { text: labelFor(m) })])); });

      // per-column: order / show / label / align / wrap / pin / format
      var orderArr = orderedColumns().map(function (c) { return c.field; });
      var labelInputs = {}, alignSel = {}, wrapChk = {}, fmtIn = {}, showChk = {}, pinSel = {};
      columns.forEach(function (c) {
        var f = c.field, isMeas = measures.indexOf(f) >= 0;
        var lab = el('input', { type: 'text', class: 'cfg-input', placeholder: f }); lab.value = config.labels[f] || ''; labelInputs[f] = lab;
        var al = el('select', { class: 'cfg-input' }); [['', 'Auto'], ['left', 'Left'], ['center', 'Centre'], ['right', 'Right']].forEach(function (o) { var op = el('option', { value: o[0], text: o[1] }); if ((config.align[f] || '') === o[0]) op.setAttribute('selected', 'selected'); al.appendChild(op); }); alignSel[f] = al;
        var wr = el('input', { type: 'checkbox' }); if (config.wrap[f]) wr.checked = true; wrapChk[f] = wr;
        var pn = el('select', { class: 'cfg-input' }); [['', 'None'], ['left', 'Left'], ['right', 'Right']].forEach(function (o) { var op = el('option', { value: o[0], text: o[1] }); if ((config.pin[f] || '') === o[0]) op.setAttribute('selected', 'selected'); pn.appendChild(op); }); pinSel[f] = pn;
        var fm = el('input', { type: 'text', class: 'cfg-input', placeholder: isMeas ? '#,##0.00' : '' }); fm.value = config.numFmt[f] || ''; if (!isMeas) fm.disabled = true; fmtIn[f] = fm;
        var sh = el('input', { type: 'checkbox' }); if (!config.hidden[f]) sh.checked = true; showChk[f] = sh;
      });
      var colList = el('div', { class: 'cfg-cols' });
      function renderColList() {
        colList.textContent = '';
        colList.appendChild(el('div', { class: 'cfg-colrow cfg-colrow8 cfg-colhead' }, [el('span', { text: '' }), el('span', { text: 'Show' }), el('span', { text: 'Field' }), el('span', { text: 'Label' }), el('span', { text: 'Align' }), el('span', { text: 'Wrap' }), el('span', { text: 'Pin' }), el('span', { text: 'Format' })]));
        orderArr.forEach(function (f, idx) {
          var up = el('button', { class: 'cfg-ord', type: 'button', text: '↑' }); up.addEventListener('click', function () { if (idx > 0) { var t = orderArr[idx - 1]; orderArr[idx - 1] = f; orderArr[idx] = t; renderColList(); } });
          var dn = el('button', { class: 'cfg-ord', type: 'button', text: '↓' }); dn.addEventListener('click', function () { if (idx < orderArr.length - 1) { var t = orderArr[idx + 1]; orderArr[idx + 1] = f; orderArr[idx] = t; renderColList(); } });
          colList.appendChild(el('div', { class: 'cfg-colrow cfg-colrow8' }, [el('span', { class: 'cfg-ordwrap' }, [up, dn]), showChk[f], el('span', { class: 'cfg-fieldname', text: f }), labelInputs[f], alignSel[f], wrapChk[f], pinSel[f], fmtIn[f]]));
        });
      }
      renderColList();

      // filters (dimensions only)
      var cascadeChk = el('input', { type: 'checkbox' }); if (config.cascade !== false) cascadeChk.checked = true;
      var applyChk = el('input', { type: 'checkbox' }); if (config.applyButton !== false) applyChk.checked = true;
      var fltShow = {}, fltType = {}, fltSort = {}, fltLatest = {};
      var fltList = el('div', { class: 'cfg-cols' });
      fltList.appendChild(el('div', { class: 'cfg-fltrow cfg-colhead' }, [el('span', { text: 'Field' }), el('span', { text: 'Show' }), el('span', { text: 'Type' }), el('span', { text: 'Sort' }), el('span', { text: 'Latest' })]));
      FILTERS.forEach(function (f) {
        var sh = el('input', { type: 'checkbox' }); if (!config.filterHidden[f]) sh.checked = true; fltShow[f] = sh;
        var ty = el('select', { class: 'cfg-input' }); [['multi', 'Multi-select'], ['single', 'Single-select'], ['search', 'Search'], ['range', 'Range'], ['date', 'Date range']].forEach(function (o) { var op = el('option', { value: o[0], text: o[1] }); if ((config.filterType[f] || 'multi') === o[0]) op.setAttribute('selected', 'selected'); ty.appendChild(op); }); fltType[f] = ty;
        var so = el('select', { class: 'cfg-input' }); [['az', 'A–Z'], ['za', 'Z–A'], ['count', 'Count']].forEach(function (o) { var op = el('option', { value: o[0], text: o[1] }); if ((config.filterSort[f] || 'az') === o[0]) op.setAttribute('selected', 'selected'); so.appendChild(op); }); fltSort[f] = so;
        var la = el('input', { type: 'checkbox' }); if (config.filterDefault[f] === 'latest') la.checked = true; fltLatest[f] = la;
        fltList.appendChild(el('div', { class: 'cfg-fltrow' }, [el('span', { class: 'cfg-fieldname', text: f }), sh, ty, so, la]));
      });

      // interaction
      var expCsv = el('input', { type: 'checkbox' }); if (!config.exports || config.exports.csv !== false) expCsv.checked = true;
      var expXlsx = el('input', { type: 'checkbox' }); if (!config.exports || config.exports.excel !== false) expXlsx.checked = true;
      var expPdf = el('input', { type: 'checkbox' }); if (!config.exports || config.exports.pdf !== false) expPdf.checked = true;
      var fnameIn = el('input', { type: 'text', class: 'cfg-input', placeholder: 'Table' }); fnameIn.value = config.fileName || '';
      var fstampChk = el('input', { type: 'checkbox' }); if (config.fileStamp) fstampChk.checked = true;
      var saveViewChk = el('input', { type: 'checkbox' }); if (config.saveView !== false) saveViewChk.checked = true;
      var qsChk = el('input', { type: 'checkbox' }); if (config.quickSearch !== false) qsChk.checked = true;

      var saveB = el('button', { class: 'btn primary', type: 'button', text: 'Save' });
      var cancelB = el('button', { class: 'btn', type: 'button', text: 'Cancel' });
      saveB.addEventListener('click', function () {
        config.title = titleIn.value; config.subtitle = subIn.value; config.showHeader = !!showHdr.checked; config.bgColor = bgIn.value.trim();
        config.density = densitySel.value; config.zebra = !!zebraChk.checked; config.zebraColor = zebraIn.value.trim();
        config.sortField = sortFieldSel.value; config.sortDir = sortDirSel.value;
        config.totals = { enabled: !!totEnable.checked, bold: !!totBold.checked, position: totPos.value, label: totLabel.value.trim(), measures: measures.filter(function (m) { return totMeasChk[m].checked; }) };
        config.order = orderArr.slice();
        config.labels = {}; config.align = {}; config.wrap = {}; config.numFmt = {}; config.hidden = {}; config.pin = {};
        orderArr.forEach(function (f) {
          var lv = labelInputs[f].value.trim(); if (lv) config.labels[f] = lv;
          var av = alignSel[f].value; if (av) config.align[f] = av;
          if (wrapChk[f].checked) config.wrap[f] = true;
          var fv = fmtIn[f].value.trim(); if (fv) config.numFmt[f] = fv;
          if (!showChk[f].checked) config.hidden[f] = true;
          var pv = pinSel[f].value; if (pv) config.pin[f] = pv;
        });
        config.cascade = !!cascadeChk.checked; config.applyButton = !!applyChk.checked;
        config.filterHidden = {}; config.filterType = {}; config.filterSort = {}; config.filterDefault = {};
        FILTERS.forEach(function (f) {
          if (!fltShow[f].checked) config.filterHidden[f] = true;
          if (fltType[f].value !== 'multi') config.filterType[f] = fltType[f].value;
          if (fltSort[f].value !== 'az') config.filterSort[f] = fltSort[f].value;
          if (fltLatest[f].checked) config.filterDefault[f] = 'latest';
        });
        config.exports = { csv: !!expCsv.checked, excel: !!expXlsx.checked, pdf: !!expPdf.checked };
        config.fileName = fnameIn.value.trim() || 'Table'; config.fileStamp = !!fstampChk.checked;
        config.saveView = !!saveViewChk.checked; config.quickSearch = !!qsChk.checked;
        saveConfig(config); document.body.removeChild(backdrop); build({ columns: columns, rows: rows });
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
          el('div', { class: 'cfg-sec', text: 'Default sort' }),
          el('div', { class: 'cfg-row' }, [el('span', { class: 'cfg-field', text: 'Sort by' }), sortFieldSel, sortDirSel]),
          el('div', { class: 'cfg-sec', text: 'Totals' }),
          el('label', { class: 'cfg-row cfg-check' }, [totEnable, el('span', { text: 'Show totals row' })]),
          el('div', { class: 'cfg-row' }, [el('span', { class: 'cfg-field', text: 'Position / style' }), totPos, el('label', { class: 'cfg-check' }, [totBold, el('span', { text: 'Bold' })])]),
          el('div', { class: 'cfg-row' }, [el('span', { class: 'cfg-field', text: 'Label' }), totLabel]),
          el('div', { class: 'cfg-row' }, [el('span', { class: 'cfg-field', text: 'Total columns' }), totMeasWrap]),
          el('div', { class: 'cfg-sec', text: 'Columns — reorder (↑↓), show, label, align, wrap, pin, format' }),
          colList,
          el('div', { class: 'cfg-sec', text: 'Filters — show, type, sort, default' }),
          el('label', { class: 'cfg-row cfg-check' }, [cascadeChk, el('span', { text: 'Cascade filters (options reflect other selections)' })]),
          el('label', { class: 'cfg-row cfg-check' }, [applyChk, el('span', { text: 'Show Apply button (off = filter as you click)' })]),
          fltList,
          el('div', { class: 'cfg-sec', text: 'Interaction' }),
          el('div', { class: 'cfg-row' }, [el('span', { class: 'cfg-field', text: 'Export buttons' }), el('label', { class: 'cfg-check' }, [expCsv, el('span', { text: 'CSV' })]), el('label', { class: 'cfg-check' }, [expXlsx, el('span', { text: 'Excel' })]), el('label', { class: 'cfg-check' }, [expPdf, el('span', { text: 'PDF' })])]),
          el('div', { class: 'cfg-row' }, [el('span', { class: 'cfg-field', text: 'Download filename' }), fnameIn]),
          el('label', { class: 'cfg-row cfg-check' }, [fstampChk, el('span', { text: 'Append date/time to filename' })]),
          el('label', { class: 'cfg-row cfg-check' }, [saveViewChk, el('span', { text: 'Show “Save view” + Reset buttons' })]),
          el('label', { class: 'cfg-row cfg-check' }, [qsChk, el('span', { text: 'Show quick-search box' })])
        ]),
        el('div', { class: 'modal-foot' }, [cancelB, saveB])
      ]));
      document.body.appendChild(backdrop);
    }

    // grid
    var gridEl = el('div', { class: 'grid ag-theme-quartz' });
    root.appendChild(gridEl);

    function orderedColumns() {
      var byField = {}; columns.forEach(function (c) { byField[c.field] = c; });
      var out = [], seen = {};
      (config.order || []).forEach(function (f) { if (byField[f] && !seen[f]) { out.push(byField[f]); seen[f] = true; } });
      columns.forEach(function (c) { if (!seen[c.field]) out.push(c); });
      return out;
    }
    function colDefs() {
      return orderedColumns().map(function (c) {
        var field = c.field, isMeas = measures.indexOf(field) >= 0;
        var align = config.align[field] || (isMeas ? 'right' : 'left');
        var wrap = !!config.wrap[field];
        var d = { colId: field, headerName: labelFor(field), field: field, wrapText: wrap, autoHeight: wrap, headerClass: 'hdr-' + align, hide: !!config.hidden[field] };
        if (config.pin[field] === 'left' || config.pin[field] === 'right') d.pinned = config.pin[field];
        if (config.sortField === field) d.sort = (config.sortDir === 'desc' ? 'desc' : 'asc');
        if (isMeas) { d.type = 'numericColumn'; d.filter = 'agNumberColumnFilter'; d.valueFormatter = function (p) { return fmtMeasure(config, field, p.value); }; }
        d.cellStyle = function (p) {
          var st = { textAlign: align, whiteSpace: wrap ? 'normal' : 'nowrap' };
          if (p.node && p.node.rowPinned && config.totals.bold) st.fontWeight = '700';
          if (isMeas) { var f = config.numFmt[field]; if (f && /\[red\]/i.test(f) && typeof p.value === 'number' && p.value < 0) st.color = '#b00020'; }
          return st;
        };
        return d;
      });
    }
    function totalMeasures() { var t = config.totals; return (t.measures && t.measures.length) ? t.measures.filter(function (m) { return measures.indexOf(m) >= 0; }) : measures; }
    function recompute() {
      if (!gridApi) return;
      var sums = {}; measures.forEach(function (m) { sums[m] = 0; });
      var n = 0;
      gridApi.forEachNodeAfterFilter(function (node) { if (node.rowPinned) return; n++; measures.forEach(function (m) { var v = node.data[m]; if (typeof v === 'number') sums[m] += v; }); });
      var tmeas = totalMeasures();
      if (config.totals.enabled && tmeas.length) {
        var vis = gridApi.getColumnState().filter(function (c) { return !c.hide; });
        var label = vis.length ? vis[0].colId : firstField;
        var pinned = {}; if (measures.indexOf(label) < 0) pinned[label] = (config.totals.label || 'Total');
        tmeas.forEach(function (m) { pinned[m] = sums[m]; });
        if (config.totals.position === 'top') { gridApi.setGridOption('pinnedTopRowData', [pinned]); gridApi.setGridOption('pinnedBottomRowData', []); }
        else { gridApi.setGridOption('pinnedBottomRowData', [pinned]); gridApi.setGridOption('pinnedTopRowData', []); }
      } else { gridApi.setGridOption('pinnedBottomRowData', []); gridApi.setGridOption('pinnedTopRowData', []); }
      metaEl.textContent = (n === rows.length) ? (rows.length + ' Total rows') : ('Showing ' + n + ' of ' + rows.length + ' rows');
    }

    if (CURRENT.gridApi) { try { CURRENT.gridApi.destroy(); } catch (e) {} CURRENT.gridApi = null; }
    gridApi = agGrid.createGrid(gridEl, {
      columnDefs: colDefs(), rowData: rows,
      defaultColDef: { sortable: true, filter: true, resizable: true },
      autoSizeStrategy: { type: 'fitCellContents' },
      suppressDragLeaveHidesColumns: true, enableCellTextSelection: true,
      getRowStyle: function (p) { if (p.node.rowPinned && config.totals.bold) return { fontWeight: '700' }; },
      isExternalFilterPresent: function () { return FILTERS.some(function (f) { return filterActive(active[f]); }); },
      doesExternalFilterPass: function (node) { for (var i = 0; i < FILTERS.length; i++) { if (!rowPasses(FILTERS[i], node.data[FILTERS[i]])) return false; } return true; },
      onGridReady: function () { gridApi.onFilterChanged(); restoreSaved(); },
      onFilterChanged: recompute, onColumnVisible: recompute,
      onColumnMoved: function (e) { if (e.finished) recompute(); }
    });
    CURRENT.gridApi = gridApi;
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

    function exportName(ext) {
      var base = (config.fileName || 'Table');
      if (config.fileStamp) { var d = new Date(), p = function (n) { return (n < 10 ? '0' : '') + n; }; base += '_' + d.getFullYear() + p(d.getMonth() + 1) + p(d.getDate()) + '_' + p(d.getHours()) + p(d.getMinutes()); }
      return base + '.' + ext;
    }
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
      downloadBlob(exportName('csv'), '﻿' + lines.join('\r\n'), 'text/csv;charset=utf-8;'); statusEl.textContent = 'CSV downloaded ✓';
    });
    dlXlsx.addEventListener('click', function () {
      dlPanel.setAttribute('hidden', 'hidden');
      if (!window.XLSX) { statusEl.textContent = 'Excel library missing'; return; }
      var v = currentView(true), aoa = [v.headers].concat(v.body); if (measures.length) aoa.push(v.total);
      var ws = XLSX.utils.aoa_to_sheet(aoa), wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Table'); XLSX.writeFile(wb, exportName('xlsx')); statusEl.textContent = 'Excel downloaded ✓';
    });
    dlPdf.addEventListener('click', function () {
      dlPanel.setAttribute('hidden', 'hidden');
      if (!window.jspdf || !window.jspdf.jsPDF) { statusEl.textContent = 'PDF library missing'; return; }
      var v = currentView(false), doc = new window.jspdf.jsPDF({ orientation: 'landscape', unit: 'pt', format: 'a4' });
      doc.setFontSize(13); doc.text(config.title || 'Table', 24, 26); doc.setFontSize(8); doc.setTextColor(120); doc.text(v.body.length + ' rows', 24, 40);
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

    // ---- filter dropdown (type-aware: multi / single / search / range / date) ----
    function makeFilter(field) {
      var type = (config.filterType && config.filterType[field]) || 'multi';
      var immediate = (config.applyButton === false);
      var btn = el('button', { class: 'filterbtn', type: 'button' }); btn._field = field; updateBtn(btn, field, 0);
      var panel = el('div', { class: 'filter-panel', hidden: 'hidden' });
      var wrap = el('div', { class: 'filter-wrap' }, [btn, panel]);

      if (type === 'search') {
        var sIn = el('input', { class: 'filter-search', type: 'text', placeholder: 'Contains…' });
        var sApply = el('button', { class: 'apply', type: 'button', text: 'Apply' }), sClear = el('button', { type: 'button', text: 'Clear' });
        var commitSearch = function () { var q = sIn.value.trim(); if (q) active[field] = { kind: 'search', q: q }; else delete active[field]; updateFilterButton(field); gridApi.onFilterChanged(); };
        var sFoot = [sClear, el('span', { class: 'spacer' })]; if (!immediate) sFoot.push(sApply);
        panel.appendChild(sIn); panel.appendChild(el('div', { class: 'filter-foot' }, sFoot));
        sApply.addEventListener('click', function () { commitSearch(); panel.setAttribute('hidden', 'hidden'); });
        sClear.addEventListener('click', function () { sIn.value = ''; commitSearch(); });
        if (immediate) sIn.addEventListener('input', commitSearch);
        btn.addEventListener('click', function (ev) { ev.stopPropagation(); if (panel.hasAttribute('hidden')) { closeAllMenus(); var a = active[field]; sIn.value = (a && a.kind === 'search') ? a.q : ''; panel.removeAttribute('hidden'); sIn.focus(); } else panel.setAttribute('hidden', 'hidden'); });
        return wrap;
      }

      if (type === 'range' || type === 'date') {
        var isDate = (type === 'date');
        var inA = el('input', { class: 'filter-range', type: isDate ? 'date' : 'number', placeholder: 'From' });
        var inB = el('input', { class: 'filter-range', type: isDate ? 'date' : 'number', placeholder: 'To' });
        var rApply = el('button', { class: 'apply', type: 'button', text: 'Apply' }), rClear = el('button', { type: 'button', text: 'Clear' });
        var commitRange = function () {
          if (isDate) { if (inA.value || inB.value) active[field] = { kind: 'date', from: inA.value, to: inB.value }; else delete active[field]; }
          else { var mn = inA.value === '' ? null : parseFloat(inA.value), mx = inB.value === '' ? null : parseFloat(inB.value); if (mn != null || mx != null) active[field] = { kind: 'range', min: mn, max: mx }; else delete active[field]; }
          updateFilterButton(field); gridApi.onFilterChanged();
        };
        var rFoot = [rClear, el('span', { class: 'spacer' })]; if (!immediate) rFoot.push(rApply);
        panel.appendChild(el('div', { class: 'filter-rangewrap' }, [inA, inB]));
        panel.appendChild(el('div', { class: 'filter-foot' }, rFoot));
        rApply.addEventListener('click', function () { commitRange(); panel.setAttribute('hidden', 'hidden'); });
        rClear.addEventListener('click', function () { inA.value = ''; inB.value = ''; commitRange(); });
        if (immediate) { inA.addEventListener('change', commitRange); inB.addEventListener('change', commitRange); }
        btn.addEventListener('click', function (ev) { ev.stopPropagation(); if (panel.hasAttribute('hidden')) { closeAllMenus(); var a = active[field]; if (a && a.kind === 'date') { inA.value = a.from || ''; inB.value = a.to || ''; } else if (a && a.kind === 'range') { inA.value = a.min == null ? '' : a.min; inB.value = a.max == null ? '' : a.max; } panel.removeAttribute('hidden'); } else panel.setAttribute('hidden', 'hidden'); });
        return wrap;
      }

      // multi / single (checkbox or radio list)
      var single = (type === 'single');
      var pending = new Set();
      var search = el('input', { class: 'filter-search', type: 'text', placeholder: 'Search…' });
      var list = el('div', { class: 'filter-list' });
      var fa = el('button', { type: 'button', text: 'All' }), fc = el('button', { type: 'button', text: 'Clear' }), fap = el('button', { class: 'apply', type: 'button', text: 'Apply' });
      var commit = function () { var all = valuesFor(field); if (pending.size === 0 || (!single && pending.size === all.length)) delete active[field]; else active[field] = new Set(pending); updateFilterButton(field); gridApi.onFilterChanged(); };
      function render() {
        list.textContent = ''; var q = search.value.trim().toLowerCase();
        valuesFor(field).forEach(function (val) {
          if (q && val.toLowerCase().indexOf(q) === -1) return;
          var inp = el('input', single ? { type: 'radio', name: 'flt-' + field } : { type: 'checkbox' }); inp.checked = pending.has(val);
          inp.addEventListener('change', function () { if (single) pending = new Set([val]); else { if (inp.checked) pending.add(val); else pending['delete'](val); } if (immediate) { commit(); if (single) panel.setAttribute('hidden', 'hidden'); } });
          var cnt = (config.filterSort && config.filterSort[field] === 'count' && optCounts[field]) ? optCounts[field][val] : null;
          list.appendChild(el('label', { class: 'filter-opt' }, [inp, el('span', { text: (val === '' ? '(blank)' : val) + (cnt ? '  (' + cnt + ')' : '') })]));
        });
      }
      panel.appendChild(search); panel.appendChild(list);
      var foot = [fa, fc, el('span', { class: 'spacer' })]; if (!immediate) foot.push(fap);
      panel.appendChild(el('div', { class: 'filter-foot' }, foot));
      btn.addEventListener('click', function (ev) { ev.stopPropagation(); if (panel.hasAttribute('hidden')) { closeAllMenus(); var cur = active[field]; pending = new Set(cur instanceof Set ? Array.from(cur) : (single ? [] : valuesFor(field))); search.value = ''; render(); panel.removeAttribute('hidden'); } else panel.setAttribute('hidden', 'hidden'); });
      search.addEventListener('input', render);
      fa.addEventListener('click', function () { pending = new Set(valuesFor(field)); render(); if (immediate) commit(); });
      fc.addEventListener('click', function () { pending = new Set(); render(); if (immediate) commit(); });
      fap.addEventListener('click', function () { commit(); panel.setAttribute('hidden', 'hidden'); });
      return wrap;
    }
    function filterCount(field) { var a = active[field]; if (!filterActive(a)) return 0; return (a instanceof Set) ? a.size : 1; }
    function updateFilterButton(field) { filterBar.querySelectorAll('.filterbtn').forEach(function (b) { if (b._field === field) updateBtn(b, field, filterCount(field)); }); }
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
    }).catch(function (err) {
      var msg = (err && err.message) || String(err);
      // Opened directly in a browser (no Tableau host iframe): show a fictional-data preview.
      if (window.self === window.top || /iframe|desktop|popup/i.test(msg)) { build(sampleModel()); return; }
      renderError('Extension failed to initialize: ' + msg);
    });
  }
  boot();
})();
