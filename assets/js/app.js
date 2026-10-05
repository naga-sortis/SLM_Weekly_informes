/* SLM Weekly Informes — UI controller */
(function () {
  'use strict';

  const C = window.SLMCore;
  const CH = window.SLMCharts;
  const $ = (sel) => document.querySelector(sel);
  const $$ = (sel) => Array.from(document.querySelectorAll(sel));

  const MAX_WEEKS = 8;
  const PAGE_SIZE = 50;

  const state = {
    fileName: '',
    sheets: [],          // [{name, rows}]
    detected: [],        // detectSheets() output
    date1904: false,
    tickets: [],
    summary: null,
    report: null,
    charts: {},
    pendingFilter: 'all',
    data: { query: '', type: '', status: '', vendor: '', year: '', sortKey: 'created', sortDir: -1, page: 0, filtered: [] },
  };

  /* ------------------------------------------------------------------ */
  /* Utilities                                                           */
  /* ------------------------------------------------------------------ */

  function el(tag, attrs, ...children) {
    const node = document.createElement(tag);
    if (attrs) {
      for (const [k, v] of Object.entries(attrs)) {
        if (v === null || v === undefined || v === false) continue;
        if (k === 'class') node.className = v;
        else if (k === 'text') node.textContent = v;
        else if (k.startsWith('on') && typeof v === 'function') node.addEventListener(k.slice(2), v);
        else node.setAttribute(k, v === true ? '' : v);
      }
    }
    for (const c of children.flat()) {
      if (c === null || c === undefined || c === false) continue;
      node.appendChild(typeof c === 'string' || typeof c === 'number' ? document.createTextNode(String(c)) : c);
    }
    return node;
  }

  class UserError extends Error {}

  let toastTimer = null;
  function toast(msg, isError) {
    const t = $('#toast');
    t.textContent = msg;
    t.classList.toggle('error', !!isError);
    t.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => { t.hidden = true; }, isError ? 7000 : 3500);
  }

  function setStatus(msg, kind) {
    const s = $('#loadStatus');
    s.replaceChildren();
    if (!msg) { s.hidden = true; return; }
    s.hidden = false;
    s.classList.toggle('error', kind === 'error');
    if (kind === 'loading') s.appendChild(el('span', { class: 'spinner', 'aria-hidden': 'true' }));
    s.appendChild(el('span', { text: msg }));
  }

  function nextFrame() {
    return new Promise((r) => requestAnimationFrame(() => setTimeout(r, 0)));
  }

  function downloadBlob(blob, name) {
    const url = URL.createObjectURL(blob);
    const a = el('a', { href: url, download: name });
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 4000);
  }

  function safeFilePart(s) {
    return String(s).normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^A-Za-z0-9_-]+/g, '-').replace(/^-+|-+$/g, '');
  }

  function weekLabel(key) {
    const w = C.weekFromKey(key);
    const dm = (ms) => C.fmtDate(ms).slice(0, 5);
    return `${w.year} · sem. ${w.week} (${dm(w.start)}–${dm(w.end - 86400000)})`;
  }

  function isDark() {
    const forced = document.documentElement.getAttribute('data-theme');
    if (forced) return forced === 'dark';
    return window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches;
  }

  /* ------------------------------------------------------------------ */
  /* Theme                                                               */
  /* ------------------------------------------------------------------ */

  function initTheme() {
    try {
      const saved = localStorage.getItem('slm-theme');
      if (saved === 'dark' || saved === 'light') document.documentElement.setAttribute('data-theme', saved);
    } catch (e) { /* storage unavailable */ }
    $('#themeToggle').addEventListener('click', () => {
      const next = isDark() ? 'light' : 'dark';
      document.documentElement.setAttribute('data-theme', next);
      try { localStorage.setItem('slm-theme', next); } catch (e) { /* ignore */ }
      if (state.report) renderCharts();
    });
    if (window.matchMedia) {
      const mq = window.matchMedia('(prefers-color-scheme: dark)');
      const onChange = () => { if (state.report && !document.documentElement.getAttribute('data-theme')) renderCharts(); };
      if (mq.addEventListener) mq.addEventListener('change', onChange);
    }
  }

  /* ------------------------------------------------------------------ */
  /* File loading                                                        */
  /* ------------------------------------------------------------------ */

  function initUpload() {
    const dz = $('#dropZone');
    const input = $('#fileInput');
    input.addEventListener('change', () => {
      if (input.files && input.files[0]) loadFile(input.files[0]);
      input.value = '';
    });
    dz.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); input.click(); }
    });
    ['dragenter', 'dragover'].forEach((ev) => dz.addEventListener(ev, (e) => {
      e.preventDefault();
      dz.classList.add('drag');
    }));
    ['dragleave', 'dragend', 'drop'].forEach((ev) => dz.addEventListener(ev, (e) => {
      e.preventDefault();
      dz.classList.remove('drag');
    }));
    dz.addEventListener('drop', (e) => {
      const f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
      if (f) loadFile(f);
    });
    // Dropping a file elsewhere on the page must not navigate away.
    window.addEventListener('dragover', (e) => e.preventDefault());
    window.addEventListener('drop', (e) => e.preventDefault());
    $('#sheetSelect').addEventListener('change', () => useSheet($('#sheetSelect').value));
  }

  async function loadFile(file) {
    if (!/\.(xlsx|xlsm|xls|csv)$/i.test(file.name)) {
      setStatus(`"${file.name}" is not an Excel/CSV file. Please choose a .xlsx, .xls or .csv file.`, 'error');
      return;
    }
    setStatus(`Reading ${file.name}…`, 'loading');
    $('#warnings').hidden = true;
    await nextFrame();
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: 'array', cellDates: false, cellFormula: false, cellHTML: false, cellStyles: false, dense: true });
      state.date1904 = !!(wb.Workbook && wb.Workbook.WBProps && wb.Workbook.WBProps.date1904);
      state.sheets = wb.SheetNames.map((name) => ({
        name,
        rows: XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: null, blankrows: true }),
      }));
      state.detected = C.detectSheets(state.sheets);
      const usable = state.detected.filter((d) => d.ok);
      if (!usable.length) {
        const partial = state.detected.find((d) => d.header);
        const why = partial ? `Missing column(s): ${partial.missing.map((f) => C.FIELD_LABELS[f]).join(', ')}.` :
          'No sheet contains a "Ticket ID" header row.';
        throw new UserError(`Could not find the ticket table in this file. ${why}`);
      }
      state.fileName = file.name;
      fillSheetSelect(usable);
      const best = usable.reduce((a, b) => (b.rows > a.rows ? b : a));
      $('#sheetSelect').value = best.name;
      useSheet(best.name);
    } catch (err) {
      if (!(err instanceof UserError)) console.error(err);
      setStatus(err && err.message ? err.message : 'The file could not be read.', 'error');
    }
  }

  function fillSheetSelect(usable) {
    const sel = $('#sheetSelect');
    sel.replaceChildren();
    for (const d of usable) sel.appendChild(el('option', { value: d.name, text: `${d.name} (${d.rows.toLocaleString('en')} rows)` }));
    if (usable.length > 1) sel.appendChild(el('option', { value: '__all__', text: 'All ticket sheets combined' }));
  }

  function useSheet(name) {
    const usable = state.detected.filter((d) => d.ok).map((d) => d.name);
    const chosen = name === '__all__' ? state.sheets.filter((s) => usable.includes(s.name)) : state.sheets.filter((s) => s.name === name);
    const ex = C.extractTickets(chosen, { date1904: state.date1904 });
    if (!ex.tickets.length) {
      setStatus('No tickets with a valid creation date were found in this sheet.', 'error');
      return;
    }
    state.tickets = ex.tickets;
    state.summary = C.summarize(ex.tickets);
    setStatus('');
    $('#fileInfo').hidden = false;
    $('#fileName').textContent = state.fileName;
    $('#fileStats').textContent = `${state.summary.total.toLocaleString('en')} tickets · ${C.fmtDate(state.summary.minDate)} → ${C.fmtDate(state.summary.maxDate)}`;
    const w = $('#warnings');
    w.replaceChildren(...ex.warnings.map((m) => el('li', { text: m })));
    w.hidden = ex.warnings.length === 0;
    $('#dropZone').classList.add('compact');
    $('#dropZone .dz-title').replaceChildren('Drop another file or ', el('u', { text: 'browse' }));
    $('#settingsCard').hidden = false;
    $('#results').hidden = false;
    $('#emptyState').hidden = true;
    initPeriodSelectors();
    initDataFilters();
    recompute();
  }

  /* ------------------------------------------------------------------ */
  /* Period & settings                                                   */
  /* ------------------------------------------------------------------ */

  function initPeriodSelectors() {
    const s = state.summary;
    const monthSel = $('#monthSelect');
    monthSel.replaceChildren(...s.monthKeys.slice().reverse().map((k) => el('option', { value: k, text: C.monthLabel(k) })));
    const firstWeek = s.weekKeys[0];
    const lastWeek = s.weekKeys[s.weekKeys.length - 1];
    const all = C.weekRange(firstWeek, lastWeek).map((w) => w.key).reverse();
    for (const id of ['#weekFrom', '#weekTo']) {
      $(id).replaceChildren(...all.map((k) => el('option', { value: k, text: weekLabel(k) })));
    }
    // Default: month of the latest data week (by its Monday)
    const defMonth = C.monthKeyOf(C.weekFromKey(lastWeek).start);
    monthSel.value = String(s.monthKeys.includes(defMonth) ? defMonth : s.monthKeys[s.monthKeys.length - 1]);
    applyMonth();
  }

  function applyMonth() {
    const mk = Number($('#monthSelect').value);
    const s = state.summary;
    const first = s.weekKeys[0];
    const last = s.weekKeys[s.weekKeys.length - 1];
    let weeks = C.weeksOfMonth(mk).map((w) => w.key).filter((k) => k >= first && k <= last);
    if (!weeks.length) {
      // month with data only in a week that starts in the previous month
      const k = C.isoWeekInfo(Date.UTC(Math.floor(mk / 100), (mk % 100) - 1, 1)).key;
      weeks = [Math.min(Math.max(k, first), last)];
    }
    $('#weekFrom').value = String(weeks[0]);
    $('#weekTo').value = String(weeks[weeks.length - 1]);
  }

  function selectedWeeks() {
    let a = Number($('#weekFrom').value);
    let b = Number($('#weekTo').value);
    if (a > b) [a, b] = [b, a];
    let weeks = C.weekRange(a, b);
    if (weeks.length > MAX_WEEKS) {
      weeks = weeks.slice(-MAX_WEEKS);
      $('#weekFrom').value = String(weeks[0].key);
      toast(`A report can cover at most ${MAX_WEEKS} weeks — the range was shortened.`);
    }
    $('#weekFrom').value = String(weeks[0].key);
    $('#weekTo').value = String(weeks[weeks.length - 1].key);
    return weeks;
  }

  function slmGroups() {
    return $('#slmGroups').value.split(/[,;\s]+/).map((s) => s.trim()).filter(Boolean);
  }

  function initSettings() {
    $('#monthSelect').addEventListener('change', () => { applyMonth(); recompute(); });
    $('#weekFrom').addEventListener('change', recompute);
    $('#weekTo').addEventListener('change', recompute);
    $('#monthsBack').addEventListener('change', recompute);
    let t;
    $('#slmGroups').addEventListener('input', () => { clearTimeout(t); t = setTimeout(recompute, 350); });
    const today = new Date();
    $('#revDate').value = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}-${String(today.getDate()).padStart(2, '0')}`;
    // remember names between sessions (convenience only)
    for (const id of ['revAuthor', 'revSupervisor', 'docSubtitle']) {
      const input = document.getElementById(id);
      try { const v = localStorage.getItem('slm-' + id); if (v) input.value = v; } catch (e) { /* ignore */ }
      input.addEventListener('change', () => { try { localStorage.setItem('slm-' + id, input.value); } catch (e) { /* ignore */ } });
    }
    $('#btnWord').addEventListener('click', generateWord);
    $('#btnExcel').addEventListener('click', exportExcel);
  }

  function recompute() {
    if (!state.tickets.length) return;
    const weeks = selectedWeeks();
    const reportMonth = Number($('#monthSelect').value) || C.monthKeyOf(weeks[weeks.length - 1].start);
    state.report = C.computeReport(state.tickets, {
      weeks,
      slmGroups: slmGroups(),
      monthsBack: Number($('#monthsBack').value) || 7,
      monthKey: reportMonth,
    });
    renderAll();
  }

  /* ------------------------------------------------------------------ */
  /* Rendering                                                           */
  /* ------------------------------------------------------------------ */

  function renderAll() {
    renderKpis();
    renderWeekly();
    renderGestor();
    renderPending();
    renderCharts();
    applyDataFilters();
  }

  function renderKpis() {
    const r = state.report;
    const sum = (arr, f) => arr.reduce((a, x) => a + f(x), 0);
    const newInc = sum(r.inc, (x) => x.nuevos);
    const newOt = sum(r.ot, (x) => x.nuevos);
    const resolved = sum(r.inc, (x) => x.resueltas) + sum(r.ot, (x) => x.resueltas);
    const escalated = r.pending.filter((t) => t.escalated).length;
    const w0 = r.weeks[0], w1 = r.weeks[r.weeks.length - 1];
    const range = r.weeks.length > 1 ? `sem. ${w0.week}–${w1.week}` : `sem. ${w0.week}`;
    const kpis = [
      ['New incidencias', newInc, range],
      ['New OTs', newOt, range],
      ['Resolved (of new)', resolved, `${newInc + newOt ? Math.round((resolved / (newInc + newOt)) * 100) : 0}% of new cases`],
      ['Unresolved at period end', r.pending.length, `as of ${C.fmtDate(r.cutoff - 86400000)}`],
      ['Escalated & unresolved', escalated, 'Ericsson · Huawei · Nokia'],
    ];
    $('#kpis').replaceChildren(...kpis.map(([l, v, s]) => el('div', { class: 'kpi' },
      el('div', { class: 'k-label', text: l }),
      el('div', { class: 'k-value', text: v.toLocaleString('en') }),
      el('div', { class: 'k-sub', text: s }))));
  }

  function dualTable(title, rows) {
    const weeks = state.report.weeks;
    const n = weeks.length;
    const thead = el('thead', null,
      el('tr', { class: 'group' },
        el('th', { class: 'blank' }),
        el('th', { class: 'sep blank' }),
        el('th', { colspan: n, text: 'Nº Incidencias' }),
        el('th', { class: 'sep' }),
        el('th', { colspan: n, text: 'Nº OTs' })),
      el('tr', { class: 'weeks' },
        el('th', { text: title }),
        el('th', { class: 'sep' }),
        weeks.map((w) => el('th', { text: `sem. ${w.week}`, title: weekLabel(w.key) })),
        el('th', { class: 'sep' }),
        weeks.map((w) => el('th', { text: `sem. ${w.week}`, title: weekLabel(w.key) }))));
    const tbody = el('tbody', null, rows.map((r) => el('tr', { class: r.total ? 'total' : null },
      el('td', { text: r.label }),
      el('td', { class: 'sep' }),
      r.inc.map((v) => el('td', { class: v === 0 ? 'zero' : null, text: v })),
      el('td', { class: 'sep' }),
      r.ot.map((v) => el('td', { class: v === 0 ? 'zero' : null, text: v })))));
    return el('table', { class: 'rep' }, thead, tbody);
  }

  function renderWeekly() {
    const r = state.report;
    const both = (fn) => ({ inc: r.inc.map(fn), ot: r.ot.map(fn) });
    $('#weeklyNote').textContent = `${C.monthLabel(r.reportMonth)} · ${r.weeks.length} week(s) from ${C.fmtDate(r.weeks[0].start)} to ${C.fmtDate(r.cutoff - 86400000)}. Data available until ${C.fmtDateTime(state.summary.maxDate)}.` +
      (r.weeks[r.weeks.length - 1].start > state.summary.maxDate ? ' ⚠ The last selected week(s) are after the last ticket in the file.' : '');
    $('#tblNew').replaceChildren(dualTable('CASOS NUEVOS SEMANA', [
      { label: 'Nuevos durante la semana', ...both((x) => x.nuevos) },
      { label: 'Resueltas de las abiertas durante la semana', ...both((x) => x.resueltas) },
      { label: 'Sin resolver (no escaladas) de las abiertas esa semana', ...both((x) => x.sinResolverNoEsc) },
      { label: 'Sin resolver (escaladas) de las abiertas esa semana', ...both((x) => x.sinResolverEsc) },
      { label: 'Devueltas de las abiertas durante la semana', ...both((x) => x.devueltas) },
      { label: 'Abiertas esa semana por el SLM', ...both((x) => x.abiertasSLM) },
    ]));
    const bl = (k) => [
      { label: 'Backlog al comenzar la semana', ...both((x) => x[k].inicio) },
      { label: 'Resueltos', ...both((x) => x[k].resueltos) },
      { label: 'Sin resolver', ...both((x) => x[k].sinResolver) },
    ];
    $('#tblBacklogNo').replaceChildren(dualTable('CASOS DEL BACKLOG NO ESCALADOS', bl('backlogNoEsc')));
    $('#tblBacklogEsc').replaceChildren(dualTable('CASOS DEL BACKLOG ESCALADOS', bl('backlogEsc')));
    $('#tblUnresolved').replaceChildren(dualTable('CASOS SIN RESOLVER', [
      { label: 'Sin resolver (no escaladas) de las abiertas esa semana', ...both((x) => x.sinResolverNoEsc) },
      { label: 'Sin resolver (escaladas) de las abiertas esa semana', ...both((x) => x.sinResolverEsc) },
      { label: 'Sin resolver (no escaladas) de semanas anteriores (backlog)', ...both((x) => x.backlogNoEsc.sinResolver) },
      { label: 'Sin resolver (escaladas) de semanas anteriores (backlog)', ...both((x) => x.backlogEsc.sinResolver) },
      { label: 'TOTAL', total: true, ...both((x) => x.sinResolverNoEsc + x.sinResolverEsc + x.backlogNoEsc.sinResolver + x.backlogEsc.sinResolver) },
    ]));
  }

  function renderGestor() {
    const r = state.report;
    const box = $('#tblGestor');
    if (!r.gestores.length) {
      box.replaceChildren(el('p', { class: 'muted', text: 'No incidencias with a gestor in "Current action" for the selected weeks.' }));
      return;
    }
    const totals = r.weeks.map((_, i) => r.gestores.reduce((a, g) => a + g.counts[i], 0));
    const table = el('table', { class: 'rep' },
      el('thead', null, el('tr', { class: 'weeks' },
        el('th', { text: 'CASOS POR GESTOR' }),
        r.weeks.map((w) => el('th', { text: `sem. ${w.week}`, title: weekLabel(w.key) })),
        el('th', { text: 'Total' }))),
      el('tbody', null,
        r.gestores.map((g) => el('tr', null,
          el('td', { text: g.gestor }),
          g.counts.map((v) => el('td', { class: v === 0 ? 'zero' : null, text: v })),
          el('td', { text: g.total, style: 'font-weight:600' }))),
        el('tr', { class: 'total' }, el('td', { text: 'TOTAL' }), totals.map((v) => el('td', { text: v })),
          el('td', { text: totals.reduce((a, b) => a + b, 0) }))));
    box.replaceChildren(table);
  }

  function renderPending() {
    const r = state.report;
    const f = state.pendingFilter;
    const list = r.pending.filter((t) => f === 'all' || t.category === f);
    $('#pendingCount').textContent = `${list.length} case(s) unresolved at ${C.fmtDate(r.cutoff - 86400000)}`;
    const box = $('#pendingList');
    if (!list.length) {
      box.replaceChildren(el('p', { class: 'muted', text: 'No unresolved cases 🎉' }));
      return;
    }
    const groupsInc = slmGroups()[0] || 'XSP00025';
    const groupsOt = slmGroups()[1] || 'XSP00027';
    box.replaceChildren(...list.map((t) => {
      const prio = t.priority.toLowerCase();
      const desc = t.description ? el('div', { class: 'desc', text: t.description }) : null;
      const toggle = t.description && t.description.length > 220 ? el('button', {
        type: 'button', class: 'linkish', text: 'Show more',
        onclick: (e) => { const ex = desc.classList.toggle('expanded'); e.currentTarget.textContent = ex ? 'Show less' : 'Show more'; },
      }) : null;
      return el('article', { class: 'pcard' },
        el('div', { class: 'pcard-head' },
          el('span', { class: 'pcard-id', text: t.id }),
          el('span', { class: 'badges' },
            el('span', { class: `badge ${t.category === 'INC' ? 'inc' : 'ot'}`, text: t.category === 'INC' ? 'Incidencia' : 'OT' }),
            el('span', { class: `badge ${prio === 'p1' ? 'p1' : prio === 'p2' ? 'p2' : ''}`, text: t.priority }),
            t.escalated ? el('span', { class: 'badge esc', text: `Escalada · ${t.vendor}` }) : null)),
        el('dl', null,
          el('dt', { text: 'Fecha de creación' }), el('dd', { text: C.fmtDateTime(t.created) }),
          el('dt', { text: 'Estado' }), el('dd', { text: `${t.open ? t.status : `Current (now: ${t.status})`} · ${t.type}` }),
          el('dt', { text: 'Persona que lo inicia' }), el('dd', { text: t.userName || '—' }),
          el('dt', { text: 'Grupo que lo inicia' }), el('dd', { text: [t.groupId, t.groupName].filter(Boolean).join(' · ') || '—' }),
          el('dt', { text: 'Acción actual' }), el('dd', { text: t.action || '—' }),
          el('dt', { text: 'Cola asociada' }), el('dd', { text: t.restorationGroupId || (t.category === 'INC' ? groupsInc : groupsOt) }),
          el('dt', { text: 'Ref. Third Party' }), el('dd', { text: t.thirdParty ? `${t.thirdParty}${t.vendor && t.vendor !== 'Otro' ? ` (${t.vendor})` : ''}` : '—' })),
        desc, toggle);
    }));
  }

  function initPendingFilter() {
    $$('#pendingSeg .seg-btn').forEach((b) => b.addEventListener('click', () => {
      $$('#pendingSeg .seg-btn').forEach((x) => x.classList.toggle('active', x === b));
      state.pendingFilter = b.dataset.v;
      if (state.report) renderPending();
    }));
  }

  /* ---------- charts ---------- */

  const CHART_DEFS = [
    ['chOpened', 'opened', false],
    ['chResolved', 'resolved', false],
    ['chEscalated', 'escalated', false],
    ['chEscGestor', 'escByGestor', false],
    ['chEscVendor', 'escByVendor', false],
    ['chStatus', 'status', true],
  ];

  function renderCharts() {
    const panelVisible = !document.querySelector('[data-panel="monthly"]').hidden;
    if (!panelVisible) { state.chartsDirty = true; return; }
    state.chartsDirty = false;
    const r = state.report;
    const labels = r.months.map(C.monthShort);
    for (const [id, key, split] of CHART_DEFS) {
      if (state.charts[id]) state.charts[id].destroy();
      const cfg = CH.barConfig(labels, r.monthly[key], { dark: isDark(), splitStacks: split });
      state.charts[id] = new Chart(document.getElementById(id), cfg);
    }
  }

  /* ---------- tabs ---------- */

  function initTabs() {
    const tabs = $$('.tab');
    const activate = (tab) => {
      tabs.forEach((t) => {
        const on = t === tab;
        t.classList.toggle('active', on);
        t.setAttribute('aria-selected', on ? 'true' : 'false');
        t.tabIndex = on ? 0 : -1;
      });
      $$('.tab-panel').forEach((p) => { p.hidden = p.dataset.panel !== tab.dataset.tab; });
      if (tab.dataset.tab === 'monthly' && state.report && (state.chartsDirty || !state.charts.chOpened)) renderCharts();
    };
    tabs.forEach((t, i) => {
      t.addEventListener('click', () => activate(t));
      t.addEventListener('keydown', (e) => {
        if (e.key !== 'ArrowRight' && e.key !== 'ArrowLeft') return;
        const next = tabs[(i + (e.key === 'ArrowRight' ? 1 : tabs.length - 1)) % tabs.length];
        next.focus();
        activate(next);
      });
    });
  }

  /* ------------------------------------------------------------------ */
  /* Extracted data table                                                */
  /* ------------------------------------------------------------------ */

  // The 11 requested fields + helpful derived columns.
  const DATA_COLUMNS = [
    { key: 'id', label: 'Ticket ID' },
    { key: 'thirdParty', label: 'Third party reference' },
    { key: 'vendor', label: 'Fabricante' },
    { key: 'created', label: 'Creation date', fmt: (v) => C.fmtDateTime(v) },
    { key: 'week', label: 'Creation week', num: true },
    { key: 'month', label: 'Creation month', num: true },
    { key: 'year', label: 'Creation year', num: true },
    { key: 'priority', label: 'Processing priority' },
    { key: 'category', label: 'Failure / OT', fmt: (v) => (v === 'INC' ? 'Failure' : 'OT') },
    { key: 'type', label: 'Ticket type' },
    { key: 'status', label: 'Status' },
    { key: 'groupId', label: 'Initiator - Group ID' },
    { key: 'groupName', label: 'Initiator - Group abbreviation name' },
    { key: 'action', label: 'Current action' },
    { key: 'gestor', label: 'Gestor' },
    { key: 'problema', label: 'Problema' },
    { key: 'tecnico', label: 'Técnico' },
  ];

  function initDataFilters() {
    const statuses = Array.from(new Set(state.tickets.map((t) => t.status))).sort();
    $('#fStatus').replaceChildren(el('option', { value: '', text: 'All statuses' }), ...statuses.map((s) => el('option', { value: s, text: s })));
    const years = Array.from(new Set(state.tickets.map((t) => new Date(t.created).getUTCFullYear()))).sort((a, b) => b - a);
    $('#fYear').replaceChildren(el('option', { value: '', text: 'All years' }), ...years.map((y) => el('option', { value: y, text: y })));
    Object.assign(state.data, { query: '', type: '', status: '', vendor: '', year: '', page: 0 });
    $('#dataSearch').value = '';
    ['#fType', '#fStatus', '#fVendor', '#fYear'].forEach((s) => { $(s).value = ''; });
  }

  function initDataTable() {
    const tr = $('#dataTable thead tr');
    tr.replaceChildren(...DATA_COLUMNS.map((c) => el('th', {
      'data-key': c.key, tabindex: 0, scope: 'col', title: 'Sort',
      onclick: () => sortBy(c.key),
      onkeydown: (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); sortBy(c.key); } },
    }, c.label, el('span', { class: 'arrow' }))));
    let t;
    $('#dataSearch').addEventListener('input', () => {
      clearTimeout(t);
      t = setTimeout(() => { state.data.query = $('#dataSearch').value; state.data.page = 0; applyDataFilters(); }, 200);
    });
    [['#fType', 'type'], ['#fStatus', 'status'], ['#fVendor', 'vendor'], ['#fYear', 'year']].forEach(([sel, k]) => {
      $(sel).addEventListener('change', () => { state.data[k] = $(sel).value; state.data.page = 0; applyDataFilters(); });
    });
    $('#pgPrev').addEventListener('click', () => { if (state.data.page > 0) { state.data.page--; renderDataPage(); } });
    $('#pgNext').addEventListener('click', () => {
      const pages = Math.ceil(state.data.filtered.length / PAGE_SIZE);
      if (state.data.page < pages - 1) { state.data.page++; renderDataPage(); }
    });
  }

  function sortBy(key) {
    const d = state.data;
    if (d.sortKey === key) d.sortDir = -d.sortDir;
    else { d.sortKey = key; d.sortDir = key === 'created' ? -1 : 1; }
    d.page = 0;
    applyDataFilters();
  }

  function applyDataFilters() {
    const d = state.data;
    const q = d.query.trim().toLowerCase();
    let list = state.tickets.filter((t) => {
      if (d.type && t.category !== d.type) return false;
      if (d.status && t.status !== d.status) return false;
      if (d.vendor === '-' && t.vendor) return false;
      if (d.vendor && d.vendor !== '-' && t.vendor !== d.vendor) return false;
      if (d.year && String(new Date(t.created).getUTCFullYear()) !== d.year) return false;
      if (q) {
        const hay = `${t.id} ${t.thirdParty} ${t.groupId} ${t.groupName} ${t.action} ${t.userName} ${t.status} ${t.type}`.toLowerCase();
        if (!hay.includes(q)) return false;
      }
      return true;
    });
    const k = d.sortKey;
    const dir = d.sortDir;
    list = list.slice().sort((a, b) => {
      const va = a[k], vb = b[k];
      if (typeof va === 'number' && typeof vb === 'number') return (va - vb) * dir;
      return String(va ?? '').localeCompare(String(vb ?? ''), 'es', { numeric: true }) * dir;
    });
    d.filtered = list;
    $$('#dataTable thead th').forEach((th) => {
      const arrow = th.querySelector('.arrow');
      const on = th.dataset.key === k;
      arrow.textContent = on ? (dir > 0 ? '▲' : '▼') : '';
      th.setAttribute('aria-sort', on ? (dir > 0 ? 'ascending' : 'descending') : 'none');
    });
    renderDataPage();
  }

  function renderDataPage() {
    const d = state.data;
    const total = d.filtered.length;
    const pages = Math.max(1, Math.ceil(total / PAGE_SIZE));
    if (d.page >= pages) d.page = pages - 1;
    const start = d.page * PAGE_SIZE;
    const rows = d.filtered.slice(start, start + PAGE_SIZE);
    const tbody = $('#dataTable tbody');
    if (!rows.length) {
      tbody.replaceChildren(el('tr', null, el('td', { colspan: DATA_COLUMNS.length, class: 'muted', text: 'No tickets match the filters.' })));
    } else {
      tbody.replaceChildren(...rows.map((t) => el('tr', null, DATA_COLUMNS.map((c) => {
        const v = c.fmt ? c.fmt(t[c.key]) : t[c.key];
        const s = v === null || v === undefined ? '' : String(v);
        return el('td', { text: s, title: s.length > 40 ? s : null });
      }))));
    }
    $('#dataCount').textContent = `${total.toLocaleString('en')} of ${state.tickets.length.toLocaleString('en')} tickets`;
    $('#pgInfo').textContent = total ? `${start + 1}–${Math.min(start + PAGE_SIZE, total)} · page ${d.page + 1} of ${pages}` : '';
    $('#pgPrev').disabled = d.page === 0;
    $('#pgNext').disabled = d.page >= pages - 1;
  }

  /* ------------------------------------------------------------------ */
  /* Exports                                                             */
  /* ------------------------------------------------------------------ */

  function setBusy(busy, msg) {
    $('#btnWord').disabled = busy;
    $('#btnExcel').disabled = busy;
    $('#actionStatus').textContent = msg || '';
  }

  async function loadLogo() {
    try {
      const res = await fetch('assets/img/logo.png');
      if (!res.ok) return null;
      const blob = await res.blob();
      const bytes = new Uint8Array(await blob.arrayBuffer());
      const bmp = await createImageBitmap(blob);
      const out = { bytes, width: bmp.width, height: bmp.height };
      if (bmp.close) bmp.close();
      return out;
    } catch (e) {
      console.warn('Logo not available', e);
      return null;
    }
  }

  async function generateWord() {
    if (!state.report) return;
    setBusy(true, 'Building the Word report…');
    await nextFrame();
    try {
      const r = state.report;
      const labels = r.months.map(C.monthShort);
      const idx = r.months.length - 1;
      const monthTxt = C.monthLabel(r.reportMonth);
      const charts = {};
      const titles = {
        opened: ['Casos abiertos por mes y prioridad', `Casos abiertos en ${monthTxt} por prioridad`],
        resolved: ['Casos resueltos por mes y prioridad', `Casos resueltos en ${monthTxt} por prioridad`],
        escalated: ['Casos escalados por mes y prioridad', `Casos escalados en ${monthTxt} por prioridad`],
        escByGestor: ['Casos escalados por mes y gestor', `Casos escalados en ${monthTxt} por gestor`],
        escByVendor: ['Casos escalados por mes y fabricante', `Casos escalados en ${monthTxt} por fabricante`],
        status: ['Casos abiertos por mes y estado', `Casos abiertos en ${monthTxt} por estado`],
      };
      for (const [, key, split] of CHART_DEFS) {
        const series = r.monthly[key];
        const bar = CH.renderPng(CH.barConfig(labels, series, { static: true, splitStacks: split, title: titles[key][0] }), 1000, 460);
        const d = CH.doughnutConfig(series, idx, titles[key][1]);
        const pie = d.empty ? null : CH.renderPng(d.config, 760, 340);
        charts[key] = { bar, pie };
        await nextFrame();
      }
      const logo = $('#includeLogo').checked ? await loadLogo() : null;
      const groups = slmGroups();
      const revDate = $('#revDate').value ? $('#revDate').value.split('-').reverse().join('.') : '';
      const meta = {
        title: $('#docTitle').value.trim() || 'Informe SLM-OSS',
        subtitle: $('#docSubtitle').value.trim(),
        revision: {
          version: $('#revVersion').value.trim(),
          date: revDate,
          author: $('#revAuthor').value.trim(),
          supervisor: $('#revSupervisor').value.trim(),
          comments: $('#revComments').value.trim(),
        },
        highlights: $('#highlights').value.split(/\r?\n/),
        otherWork: $('#otherWork').value.split(/\r?\n/),
        includeDescription: $('#includeDesc').checked,
        queueInc: groups[0] || 'XSP00025',
        queueOt: groups[1] || 'XSP00027',
        dataUntil: state.summary.maxDate,
      };
      const doc = SLMDocx.buildDocument(window.docx, C, r, meta, { logo, charts });
      const blob = await docx.Packer.toBlob(doc);
      const name = `INFORME-SLM-OSS-Sortis-${safeFilePart(monthTxt.replace(' ', '-'))}-sem${r.weeks[0].week}-${r.weeks[r.weeks.length - 1].week}.docx`;
      downloadBlob(blob, name);
      setBusy(false, '');
      toast('Word report downloaded.');
    } catch (err) {
      console.error(err);
      setBusy(false, '');
      toast(`Could not create the Word report: ${err && err.message ? err.message : err}`, true);
    }
  }

  async function exportExcel() {
    if (!state.tickets.length) return;
    setBusy(true, 'Building the Excel file…');
    await nextFrame();
    try {
      const cols = DATA_COLUMNS.filter((c) => c.key !== 'created');
      cols.splice(3, 0, { key: 'created', label: 'Creation date', date: true });
      cols.push({ key: 'escalated', label: 'Escalada', fmt: (v) => (v ? 'Sí' : 'No') });
      cols.push({ key: 'devuelto', label: 'Devuelta', fmt: (v) => (v ? 'Sí' : 'No') });
      cols.push({ key: 'restored', label: 'Restoration date', date: true });
      const aoa = [cols.map((c) => c.label)];
      // export what is currently filtered in the "Extracted data" tab (all tickets when no filter)
      const list = hasDataFilter() ? state.data.filtered : state.tickets;
      for (const t of list) {
        aoa.push(cols.map((c) => {
          const v = t[c.key];
          if (c.date) return v === null || v === undefined ? null : C.toExcelSerial(v);
          if (c.fmt) return c.fmt(v);
          return v === '' ? null : v;
        }));
      }
      const ws = XLSX.utils.aoa_to_sheet(aoa);
      // date formatting
      cols.forEach((c, ci) => {
        if (!c.date) return;
        for (let ri = 1; ri < aoa.length; ri++) {
          const ref = XLSX.utils.encode_cell({ r: ri, c: ci });
          if (ws[ref] && ws[ref].t === 'n') ws[ref].z = 'yyyy-mm-dd hh:mm';
        }
      });
      ws['!cols'] = cols.map((c) => ({ wch: Math.min(40, Math.max(10, c.label.length + 2)) }));
      ws['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: aoa.length - 1, c: cols.length - 1 } }) };
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Tickets');
      appendReportSheet(wb);
      const base = state.fileName.replace(/\.[^.]+$/, '');
      XLSX.writeFile(wb, `${safeFilePart(base) || 'tickets'}-extract.xlsx`, { compression: true });
      setBusy(false, '');
      toast(`Excel file downloaded (${list.length.toLocaleString('en')} tickets).`);
    } catch (err) {
      console.error(err);
      setBusy(false, '');
      toast(`Could not create the Excel file: ${err && err.message ? err.message : err}`, true);
    }
  }

  function hasDataFilter() {
    const d = state.data;
    return !!(d.query.trim() || d.type || d.status || d.vendor || d.year);
  }

  /** Second sheet with the weekly report tables. */
  function appendReportSheet(wb) {
    const r = state.report;
    if (!r) return;
    const wk = r.weeks.map((w) => `sem. ${w.week}`);
    const aoa = [[`Informe SLM-OSS · ${C.monthLabel(r.reportMonth)}`], []];
    const block = (title, rows) => {
      aoa.push([title, ...wk.map((w) => `INC ${w}`), '', ...wk.map((w) => `OT ${w}`)]);
      for (const [label, fn] of rows) aoa.push([label, ...r.inc.map(fn), '', ...r.ot.map(fn)]);
      aoa.push([]);
    };
    block('CASOS NUEVOS SEMANA', [
      ['Nuevos durante la semana', (x) => x.nuevos],
      ['Resueltas de las abiertas durante la semana', (x) => x.resueltas],
      ['Sin resolver (no escaladas) de las abiertas esa semana', (x) => x.sinResolverNoEsc],
      ['Sin resolver (escaladas) de las abiertas esa semana', (x) => x.sinResolverEsc],
      ['Devueltas de las abiertas durante la semana', (x) => x.devueltas],
      ['Abiertas esa semana por el SLM', (x) => x.abiertasSLM],
    ]);
    for (const [title, k] of [['CASOS DEL BACKLOG NO ESCALADOS', 'backlogNoEsc'], ['CASOS DEL BACKLOG ESCALADOS', 'backlogEsc']]) {
      block(title, [
        ['Backlog al comenzar la semana', (x) => x[k].inicio],
        ['Resueltos', (x) => x[k].resueltos],
        ['Sin resolver', (x) => x[k].sinResolver],
      ]);
    }
    aoa.push(['CASOS POR GESTOR', ...wk, 'Total']);
    for (const g of r.gestores) aoa.push([g.gestor, ...g.counts, g.total]);
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = [{ wch: 55 }, ...Array(wk.length * 2 + 1).fill({ wch: 11 })];
    XLSX.utils.book_append_sheet(wb, ws, 'Informe semanal');
  }

  /* ------------------------------------------------------------------ */
  /* Boot                                                                */
  /* ------------------------------------------------------------------ */

  function boot() {
    const missing = ['XLSX', 'Chart', 'docx', 'SLMCore', 'SLMCharts', 'SLMDocx'].filter((g) => !window[g]);
    if (missing.length) {
      setStatus(`Some components failed to load (${missing.join(', ')}). Please reload the page.`, 'error');
      return;
    }
    initTheme();
    initUpload();
    initSettings();
    initTabs();
    initPendingFilter();
    initDataTable();
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
  else boot();
})();
