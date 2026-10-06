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
    return `Semana ${w.week} de ${w.year} · ${C.fmtDate(w.start)} – ${C.fmtDate(w.end - 86400000)} · ${C.periodMonthsLabel(w.start, w.end)}`;
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
      if (state.report) { renderCharts(); renderPeriodCharts(); renderTeamCharts(); }
    });
    if (window.matchMedia) {
      const mq = window.matchMedia('(prefers-color-scheme: dark)');
      const onChange = () => { if (state.report && !document.documentElement.getAttribute('data-theme')) { renderCharts(); renderPeriodCharts(); renderTeamCharts(); } };
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
    // Every ticket sheet of the file, merged (Sortis team view + ticket details)
    const allEx = C.extractTickets(state.sheets.filter((s) => usable.includes(s.name)), { date1904: state.date1904 });
    state.allTickets = allEx.tickets;
    state.byId = new Map(state.allTickets.map((t) => [t.id, t]));
    initPeriodSelectors();
    initPeriodFilter();
    initTeamFilter();
    initDataFilters();
    recompute();
  }

  /* ------------------------------------------------------------------ */
  /* Period & settings                                                   */
  /* ------------------------------------------------------------------ */

  function initPeriodSelectors() {
    const s = state.summary;
    const first = s.weekKeys[0];
    const last = s.weekKeys[s.weekKeys.length - 1];
    const all = C.weekRange(first, last).map((w) => w.key).reverse();
    $('#reportWeek').replaceChildren(...all.map((k) => el('option', { value: k, text: weekLabel(k) })));
    // Default: the latest week that has tickets in the file
    $('#reportWeek').value = String(last);
  }

  /** The report week plus the previous weeks shown in the tables (oldest first). */
  function selectedWeeks() {
    const endKey = Number($('#reportWeek').value);
    const n = Math.max(1, Math.min(MAX_WEEKS, Number($('#weeksShown').value) || 5));
    const weeks = [C.weekFromKey(endKey)];
    while (weeks.length < n) weeks.unshift(C.weekFromKey(C.isoWeekInfo(weeks[0].start - 86400000).key));
    return weeks;
  }

  function slmGroups() {
    return $('#slmGroups').value.split(/[,;\s]+/).map((s) => s.trim()).filter(Boolean);
  }

  function initSettings() {
    $('#reportWeek').addEventListener('change', recompute);
    $('#weeksShown').addEventListener('change', recompute);
    $('#monthsBack').addEventListener('change', recompute);
    $('#inferGestor').addEventListener('change', recompute);
    $('#checksToggle').addEventListener('click', () => {
      const list = $('#checksList');
      list.hidden = !list.hidden;
      $('#checksToggle').setAttribute('aria-expanded', list.hidden ? 'false' : 'true');
    });
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
    state.report = C.computeReport(state.tickets, {
      weeks: selectedWeeks(),
      slmGroups: slmGroups(),
      monthsBack: Number($('#monthsBack').value) || 7,
      inferGestor: $('#inferGestor').checked,
      dataUntil: state.summary.dataUntil,
    });
    renderAll();
    recomputePeriod();
    fillTeamQueues();
    recomputeTeam();
  }

  /* ------------------------------------------------------------------ */
  /* Month / Year view                                                   */
  /* ------------------------------------------------------------------ */

  function initPeriodFilter() {
    state.periodsAvail = C.periodsAvailable(state.tickets);
    $('#pYear').replaceChildren(...state.periodsAvail.map((p) => el('option', { value: p.year, text: p.year })));
    const latest = state.periodsAvail[0];
    $('#pYear').value = String(latest.year);
    fillPeriodMonths(latest.months[latest.months.length - 1]);
  }

  function fillPeriodMonths(selected) {
    const y = Number($('#pYear').value);
    const entry = state.periodsAvail.find((p) => p.year === y);
    const months = entry ? entry.months : [];
    $('#pMonth').replaceChildren(el('option', { value: '', text: `Todo el año ${y}` }),
      ...months.map((m) => el('option', { value: m, text: C.MONTHS_ES[m - 1] })));
    $('#pMonth').value = selected && months.includes(selected) ? String(selected) : '';
  }

  function initPeriodControls() {
    $('#pYear').addEventListener('change', () => {
      const prev = Number($('#pMonth').value) || null;
      fillPeriodMonths(prev);
      recomputePeriod();
    });
    $('#pMonth').addEventListener('change', recomputePeriod);
    $('#pInWord').addEventListener('change', () => { $('#secPeriod').checked = $('#pInWord').checked; updateWordButton(); });
  }

  function recomputePeriod() {
    if (!state.tickets.length || !$('#pYear').value) return;
    state.period = C.computePeriod(state.tickets, {
      year: Number($('#pYear').value),
      month: Number($('#pMonth').value) || null,
      slmGroups: slmGroups(),
      inferGestor: $('#inferGestor').checked,
      dataUntil: state.summary.dataUntil,
    });
    renderPeriodView();
    updateWordButton();
  }

  function renderPeriodView() {
    const p = state.period;
    const I = p.summary.INC, O = p.summary.OT;
    $('#pNote').textContent = periodNote(p);
    const failed = p.checks.filter((c) => !c.ok);
    const box = $('#pChecks');
    box.className = `checks small-checks${failed.length ? ' fail' : ''}`;
    box.replaceChildren(el('div', { class: 'checks-head', text: failed.length
      ? `✖ ${failed.length} of ${p.checks.length} checks failed for ${p.label}: ${failed.map((c) => c.label + (c.detail ? ' — ' + c.detail : '')).join(' · ')}`
      : `✔ All ${p.checks.length} consistency checks passed for ${p.label}` }));
    const rows = [
      ['Casos abiertos en el periodo', I.abiertos, O.abiertos],
      ['  · Resueltos', I.resueltos, O.resueltos],
      ['  · Pendientes al cierre', I.pendientes, O.pendientes],
      ['  · Devueltos', I.devueltos, O.devueltos],
      ['Escalados (Ericsson / Huawei / Nokia)', I.escalados, O.escalados],
      ['Abiertos por el SLM', I.abiertosSLM, O.abiertosSLM],
      ['Resueltos durante el periodo (incl. abiertos antes)', I.resueltosEnPeriodo, O.resueltosEnPeriodo],
      ['Pendientes totales al cierre (incl. backlog)', I.pendientesTotales, O.pendientesTotales],
    ];
    $('#pSummary').replaceChildren(el('table', { class: 'rep compact' },
      el('thead', null, el('tr', { class: 'weeks' }, el('th', { text: p.label.toUpperCase() }), el('th', { text: 'Incidencias' }), el('th', { text: 'OTs' }), el('th', { text: 'Total' }))),
      el('tbody', null, rows.map(([l, a, b], i) => el('tr', { class: i === 0 ? 'total' : null },
        el('td', { text: l, style: l.startsWith('  ') ? 'padding-left:22px' : null }),
        el('td', { class: a === 0 ? 'zero' : null, text: a }), el('td', { class: b === 0 ? 'zero' : null, text: b }),
        el('td', { text: a + b, style: 'font-weight:600' }))))));
    const tot = p.gestores.reduce((a, g) => [a[0] + g.inc, a[1] + g.ot, a[2] + g.total], [0, 0, 0]);
    $('#pGestor').replaceChildren(p.gestores.length ? el('table', { class: 'rep compact' },
      el('thead', null, el('tr', { class: 'weeks' }, el('th', { text: 'GESTOR' }), el('th', { text: 'Incidencias' }), el('th', { text: 'OTs' }), el('th', { text: 'Total' }))),
      el('tbody', null,
        p.gestores.map((g) => el('tr', { class: g.unidentified ? 'unidentified' : null },
          el('td', { text: g.gestor, title: g.unidentified ? 'Show these tickets' : null,
            onclick: g.unidentified ? () => $('#pNoGestorPanel').scrollIntoView({ behavior: 'smooth', block: 'start' }) : null }), el('td', { class: g.inc ? null : 'zero', text: g.inc }), el('td', { class: g.ot ? null : 'zero', text: g.ot }),
          el('td', { text: g.total, style: 'font-weight:600' }))),
        el('tr', { class: 'total' }, el('td', { text: 'TOTAL' }), tot.map((v) => el('td', { text: v })))))
      : el('p', { class: 'muted', text: 'No cases in this period.' }));
    renderNoGestorPanel('period', 'pNoGestorPanel', p.noGestorTickets, `in ${p.label} (incidencias + OTs)`);
    renderPeriodCharts();
  }

  function periodNote(p) {
    const I = p.summary.INC, O = p.summary.OT;
    let t = `${p.label}: ${I.abiertos + O.abiertos} casos abiertos (${I.abiertos} incidencias y ${O.abiertos} OTs), del ${C.fmtDate(p.start)} al ${C.fmtDate(p.end - 86400000)}.`;
    if (p.partial) t += ` El periodo no ha terminado: datos hasta el ${C.fmtDateTime(p.asOf)}.`;
    t += p.month ? ' Las semanas del mes se recortan a los días del mes.' : '';
    return t;
  }

  function gestorCaption(p) {
    const g = p.gestorChart;
    const parts = [`Los ${g.labels.length} gestores con más casos`];
    if (g.othersCount) parts.push(`otros ${g.othersCount} gestores suman ${g.othersTotal} casos`);
    if (g.unidentifiedTotal) parts.push(`${g.unidentifiedTotal} casos sin gestor en la acción actual`);
    return parts.join('; ') + ' (detalle completo en la tabla).';
  }

  const PERIOD_CHARTS = [
    ['pchOpened', 'opened', false],
    ['pchResolved', 'resolved', false],
    ['pchStatus', 'status', true],
    ['pchVendor', 'vendor', false],
  ];

  function renderPeriodCharts() {
    if (!state.period) return;
    if (document.querySelector('[data-panel="period"]').hidden) { state.periodChartsDirty = true; return; }
    state.periodChartsDirty = false;
    const p = state.period;
    const labels = p.buckets.map((b) => b.label);
    for (const [id, key, split] of PERIOD_CHARTS) makeChart(id, labels, p.series[key], { splitStacks: split });
    $('#pGestorCaption').textContent = gestorCaption(p);
    makeChart('pchGestor', p.gestorChart.labels, p.gestorChart.series, { horizontal: true });
  }

  /* ---------- "Sin gestor identificado" panel ---------- */

  const ngState = {};

  function renderNoGestorPanel(key, containerId, tickets, scopeTxt) {
    const box = document.getElementById(containerId);
    const st = ngState[key] || (ngState[key] = { reason: '', q: '', limit: 100 });
    if (!tickets.length) {
      box.replaceChildren(el('h4', { text: 'Sin gestor identificado' }),
        el('p', { class: 'ok-note', text: `✔ Every case ${scopeTxt} has a gestor in “Current action”.` }));
      return;
    }
    const counts = new Map();
    for (const t of tickets) { const r = C.noGestorReason(t); counts.set(r.code, (counts.get(r.code) || 0) + 1); }
    if (st.reason && !counts.has(st.reason)) st.reason = '';
    const q = st.q.trim().toLowerCase();
    const list = tickets.filter((t) => (!st.reason || C.noGestorReason(t).code === st.reason) &&
      (!q || `${t.id} ${t.action} ${t.description} ${t.groupName} ${t.gestorDesc}`.toLowerCase().includes(q)));
    const failures = counts.get('unrecognised') || 0;
    const rerender = () => renderNoGestorPanel(key, containerId, tickets, scopeTxt);

    const chips = el('div', { class: 'chips', role: 'group', 'aria-label': 'Filter by reason' },
      el('button', { type: 'button', class: `chip${st.reason ? '' : ' active'}`, onclick: () => { st.reason = ''; st.limit = 100; rerender(); } },
        'All', el('b', { text: tickets.length })),
      Array.from(counts.entries()).sort((a, b) => b[1] - a[1]).map(([code, n]) => el('button', {
        type: 'button', class: `chip${st.reason === code ? ' active' : ''}`, title: C.NO_GESTOR_REASONS[code],
        onclick: () => { st.reason = code; st.limit = 100; rerender(); },
      }, C.NO_GESTOR_REASONS[code].split(' (')[0].split(' — ')[0], el('b', { text: n }))));

    const search = el('input', { type: 'search', placeholder: 'Search ticket ID, action, description…', value: st.q, 'aria-label': 'Search tickets without gestor' });
    let timer;
    search.addEventListener('input', () => { clearTimeout(timer); timer = setTimeout(() => { st.q = search.value; st.limit = 100; rerender(); const s2 = box.querySelector('input[type=search]'); if (s2) { s2.focus(); s2.setSelectionRange(s2.value.length, s2.value.length); } }, 250); });

    const shown = list.slice(0, st.limit);
    const table = el('table', { class: 'data-table' },
      el('thead', null, el('tr', null, ['Ticket ID', 'Creation date', 'Week', 'Type', 'Status', 'Current action (as in the Excel)', 'Reason', 'Gestor mentioned in description', 'Description']
        .map((h) => el('th', { text: h, style: 'cursor:default' })))),
      el('tbody', null, shown.length ? shown.map((t) => el('tr', null,
        el('td', null, tidButton(t.id)),
        el('td', { text: C.fmtDateTime(t.created) }),
        el('td', { text: `${t.weekYear}-S${String(t.week).padStart(2, '0')}` }),
        el('td', { text: t.category === 'INC' ? 'Incidencia' : `OT (${t.type})` }),
        el('td', { text: t.status }),
        t.action ? el('td', { text: t.action, title: t.action }) : el('td', null, el('span', { class: 'empty-val', text: '(empty)' })),
        el('td', { class: 'reason', text: C.noGestorReason(t).label }),
        el('td', { text: t.gestorDesc || '—' }),
        el('td', { class: 'desc', text: t.description.replace(/\s+/g, ' '), title: t.description.slice(0, 1500) })))
        : el('tr', null, el('td', { colspan: 9, class: 'muted', text: 'No tickets match the filter.' }))));

    box.replaceChildren(...[
      el('div', { class: 'ng-head' },
        el('div', null,
          el('h4', { text: `Sin gestor identificado — ${tickets.length} ticket(s) ${scopeTxt}` }),
          el('p', { class: failures ? 'small' : 'ok-note', style: failures ? 'color:var(--err-text);margin:0' : 'margin:0',
            text: failures ? `⚠ ${failures} ticket(s) have an action the tool could not read as GESTOR - PROBLEMA - TÉCNICO — please check them.`
              : '✔ None of these is a reading error: the Excel has no gestor in “Current action” for them (see the reason of each ticket).' })),
        el('button', { type: 'button', class: 'btn small', onclick: () => exportNoGestor(list, scopeTxt) }, `Export ${list.length} to Excel`)),
      chips,
      el('div', { class: 'toolbar' }, search, el('span', { class: 'muted small', text: `${list.length} shown of ${tickets.length}` })),
      el('div', { class: 'table-wrap data-wrap' }, table),
      list.length > shown.length ? el('button', { type: 'button', class: 'btn small more', onclick: () => { st.limit += 200; rerender(); } },
        `Show more (${list.length - shown.length} remaining)`) : null].filter(Boolean));
  }

  function exportNoGestor(list, scopeTxt) {
    try {
      const aoa = [['Ticket ID', 'Creation date', 'Week (ISO)', 'Ticket type', 'Status', 'Current action', 'Reason', 'Gestor mentioned in description',
        'Initiator - Group ID', 'Initiator - Group abbreviation name', 'Description']];
      for (const t of list) {
        aoa.push([t.id, C.toExcelSerial(t.created), `${t.weekYear}-S${String(t.week).padStart(2, '0')}`, t.type, t.status, t.action,
          C.noGestorReason(t).label, t.gestorDesc, t.groupId, t.groupName, t.description.slice(0, 32000)]);
      }
      const ws = XLSX.utils.aoa_to_sheet(aoa);
      for (let r = 1; r < aoa.length; r++) { const ref = XLSX.utils.encode_cell({ r, c: 1 }); if (ws[ref]) ws[ref].z = 'yyyy-mm-dd hh:mm'; }
      ws['!cols'] = [14, 17, 11, 14, 18, 40, 55, 22, 14, 30, 80].map((w) => ({ wch: w }));
      ws['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: aoa.length - 1, c: aoa[0].length - 1 } }) };
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Sin gestor');
      XLSX.writeFile(wb, `sin-gestor-${safeFilePart(scopeTxt)}.xlsx`, { compression: true });
      toast(`Excel downloaded (${list.length} tickets).`);
    } catch (err) {
      console.error(err);
      toast(`Could not create the Excel file: ${err && err.message ? err.message : err}`, true);
    }
  }

  /* ---------- Ticket detail window ---------- */

  function findTicket(id) {
    return (state.byId && state.byId.get(id)) || state.tickets.find((t) => t.id === id) || null;
  }

  function hoursTxt(h) {
    if (h === null || h === undefined) return '—';
    if (h < 1) return `${Math.round(h * 60)} min`;
    if (h < 48) return `${h.toFixed(1)} h`;
    return `${(h / 24).toFixed(1)} días`;
  }

  function openTicket(id) {
    const t = findTicket(id);
    if (!t) { toast(`Ticket ${id} not found in the file.`, true); return; }
    const groups = new Set(slmGroups().map((g) => g.toUpperCase()));
    const sortis = (g) => g && groups.has(g.toUpperCase());
    $('#tdTitle').textContent = t.id;
    const prio = t.priority.toLowerCase();
    $('#tdBadges').replaceChildren(...[
      el('span', { class: `badge ${t.category === 'INC' ? 'inc' : 'ot'}`, text: t.category === 'INC' ? 'Incidencia' : `OT · ${t.type}` }),
      el('span', { class: `badge ${prio === 'p1' ? 'p1' : prio === 'p2' ? 'p2' : ''}`, text: t.priority }),
      el('span', { class: 'badge', text: t.status }),
      t.escalated ? el('span', { class: 'badge esc', text: `Escalada · ${t.vendor}` }) : null,
      t.devuelto ? el('span', { class: 'badge p1', text: 'Devuelto' }) : null,
      sortis(t.restorationGroupId) ? el('span', { class: 'badge p2', text: `Sortis · ${t.restorationGroupId}` }) : null].filter(Boolean));

    const step = (title, date, who, grp, done, isSortis) => el('div', { class: `td-step${done ? ' done' : ''}${isSortis ? ' sortis' : ''}` },
      el('div', { class: 'st-title', text: title }),
      el('div', { class: 'st-date', text: date || '—' }),
      el('div', { class: 'st-who', text: who || '—' }),
      el('div', { class: 'st-grp', text: grp || '' }));
    const hrs = t.restored !== null && t.restored >= t.created ? (t.restored - t.created) / 3600000 : null;
    const flow = el('div', { class: 'td-flow' },
      step('1 · Opened by', C.fmtDateTime(t.created), t.userName, [t.groupId, t.groupName].filter(Boolean).join(' · '), true, sortis(t.groupId)),
      step('2 · Acknowledged', t.ackDate !== null ? C.fmtDateTime(t.ackDate) : '', '', '', t.ackDate !== null, false),
      step('3 · Restored by', t.restored !== null ? `${C.fmtDateTime(t.restored)} · ${hoursTxt(hrs)}` : '', t.restorationUser,
        [t.restorationGroupId, t.restorationGroupName].filter(Boolean).join(' · '), t.restored !== null, sortis(t.restorationGroupId)),
      step('4 · Closed by', t.closureUser || t.closureGroups.length ? t.status : '', t.closureUser, t.closureGroups.join(' / '),
        !t.open, t.closureGroups.some((g) => sortis(g.split(' ')[0]))));

    const grid = (pairs) => el('div', { class: 'td-grid' }, pairs.flatMap(([k, v, alt]) => [
      el('div', { class: 'k', text: k }),
      el('div', { class: `v${v ? '' : ' empty'}` }, v || '(empty)', alt && alt.length ? el('span', { class: 'alt', text: `Other value(s) in another row of the file: ${alt.join(' / ')}` }) : null),
    ]));
    const ng = t.gestor ? null : C.noGestorReason(t).label;
    const derived = grid([
      ['Incidencia / OT', t.category === 'INC' ? 'Incidencia (Failure)' : `OT (${t.type})`],
      ['Week (ISO) of creation', `${t.weekYear}-S${String(t.week).padStart(2, '0')}`],
      ['Gestor (from Current action)', t.gestor || `— ${ng}`],
      ['Problema', t.problema],
      ['Técnico', t.tecnico],
      ['Gestor mentioned in description', t.gestorDesc],
      ['Third party vendor', t.vendor ? `${t.vendor}${t.escalated ? ' (escalated)' : ''}` : ''],
      ['Time from creation to restoration', hrs === null ? '' : hoursTxt(hrs)],
      ['Found in sheet(s)', t.sheets.join(' / ')],
    ]);
    const raw = grid(t.raw.filter(([k]) => k !== 'Description').map(([k, v]) => [k, v, t.rawAlt[k]]));
    $('#tdBody').replaceChildren(
      flow,
      el('div', { class: 'td-sec' }, el('h4', { text: 'Interpreted by the tool' }), derived),
      el('div', { class: 'td-sec' }, el('h4', { text: 'Description' }), el('div', { class: 'td-desc', text: t.description || '(empty)' })),
      el('div', { class: 'td-sec' }, el('h4', { text: `All columns as in the Excel (${t.raw.length})` }), raw));
    const dlg = $('#ticketDialog');
    if (typeof dlg.showModal === 'function') { if (!dlg.open) dlg.showModal(); } else dlg.setAttribute('open', '');
    $('#tdBody').scrollTop = 0;
  }

  function initTicketDialog() {
    document.addEventListener('click', (e) => {
      const b = e.target.closest && e.target.closest('[data-tid]');
      if (b) { e.preventDefault(); openTicket(b.dataset.tid); }
    });
    const dlg = $('#ticketDialog');
    $('#tdClose').addEventListener('click', () => dlg.close());
    dlg.addEventListener('click', (e) => { if (e.target === dlg) dlg.close(); }); // click on the backdrop
  }

  /* ---------- Sortis team view ---------- */

  const tState = { engineer: '', type: '', q: '', limit: 200 };

  function initTeamFilter() {
    const years = Array.from(new Set(state.allTickets.flatMap((t) => [t.created, t.restored].filter((x) => x !== null))
      .map((ms) => new Date(ms).getUTCFullYear()))).sort((a, b) => b - a);
    $('#tYear').replaceChildren(el('option', { value: '', text: 'All years' }), ...years.map((y) => el('option', { value: y, text: y })));
    $('#tYear').value = years.length ? String(years[0]) : '';
    fillTeamMonths();
    fillTeamQueues();
  }

  function fillTeamMonths() {
    const y = $('#tYear').value;
    $('#tMonth').disabled = !y;
    $('#tMonth').replaceChildren(el('option', { value: '', text: y ? `Todo el año ${y}` : '—' }),
      ...(y ? C.MONTHS_ES.map((m, i) => el('option', { value: i + 1, text: m })) : []));
  }

  function fillTeamQueues() {
    const prev = $('#tQueue').value;
    const groups = slmGroups();
    $('#tQueue').replaceChildren(el('option', { value: '', text: groups.join(' + ') || '—' }), ...groups.map((g) => el('option', { value: g, text: g })));
    $('#tQueue').value = groups.includes(prev) ? prev : '';
    $('#tGroupsLabel').textContent = `(${groups.join(' / ')})`;
  }

  function initTeamControls() {
    $('#tYear').addEventListener('change', () => { fillTeamMonths(); recomputeTeam(); });
    ['#tMonth', '#tBasis', '#tQueue'].forEach((id) => $(id).addEventListener('change', recomputeTeam));
    $('#tEngineer').addEventListener('change', () => { tState.engineer = $('#tEngineer').value; tState.limit = 200; renderTeamList(); });
    $('#tType').addEventListener('change', () => { tState.type = $('#tType').value; tState.limit = 200; renderTeamList(); });
    let timer;
    $('#tSearch').addEventListener('input', () => { clearTimeout(timer); timer = setTimeout(() => { tState.q = $('#tSearch').value; tState.limit = 200; renderTeamList(); }, 250); });
    $('#tExport').addEventListener('click', exportTeam);
    $('#tInWord').addEventListener('change', () => { $('#secTeam').checked = $('#tInWord').checked; updateWordButton(); });
  }

  function recomputeTeam() {
    if (!state.allTickets || !state.allTickets.length) return;
    state.team = C.computeTeam(state.allTickets, {
      groups: slmGroups(),
      year: $('#tYear').value || null,
      month: $('#tMonth').value || null,
      basis: $('#tBasis').value,
      queue: $('#tQueue').value,
    });
    renderTeam();
    updateWordButton();
  }

  function teamNote(T) {
    const basis = T.basis === 'created' ? 'fecha de creación' : 'fecha de restauración';
    let t = `${T.label}: ${T.total} casos restaurados por el equipo Sortis (${T.queue || T.groups.join(' / ')}): ${T.inc} incidencias y ${T.ot} OTs, ` +
      `por ${T.engineers.filter((e) => !e.noUser).length} ingenieros. Periodo según la ${basis}.`;
    if (T.year && T.undated) t += ` ${T.undated} caso(s) del equipo no tienen fecha de restauración y no se pueden situar en un periodo (aparecen en “All years”).`;
    return t;
  }

  function renderTeam() {
    const T = state.team;
    const fmtH = (h) => (h === null ? '—' : hoursTxt(h));
    $('#tNote').textContent = teamNote(T);
    const failed = T.checks.filter((c) => !c.ok);
    const cb = $('#tChecks');
    cb.className = `checks small-checks${failed.length ? ' fail' : ''}`;
    cb.replaceChildren(el('div', { class: 'checks-head', text: failed.length
      ? `✖ ${failed.length} of ${T.checks.length} checks failed: ${failed.map((c) => c.label + (c.detail ? ' — ' + c.detail : '')).join(' · ')}`
      : `✔ All ${T.checks.length} consistency checks passed for the Sortis team view` }));
    $('#tKpis').replaceChildren(...[
      ['Restored by the team', T.total, T.label],
      ['Incidencias', T.inc, `${T.total ? Math.round((T.inc / T.total) * 100) : 0}%`],
      ['OTs', T.ot, `${T.total ? Math.round((T.ot / T.total) * 100) : 0}%`],
      ['Engineers', T.engineers.filter((e) => !e.noUser).length, 'Restoration user name'],
      ['Median time to restore', fmtH(T.medianHours), 'creation → restoration'],
    ].map(([l, v, sub]) => el('div', { class: 'kpi' }, el('div', { class: 'k-label', text: l }),
      el('div', { class: 'k-value', text: typeof v === 'number' ? v.toLocaleString('en') : v }), el('div', { class: 'k-sub', text: sub }))));

    $('#tGroups').replaceChildren(el('table', { class: 'rep compact' },
      el('thead', null, el('tr', { class: 'weeks' }, ['Cola (Restoration group)', 'Nombre', 'Incidencias', 'OTs', 'Total', 'Ingenieros'].map((h) => el('th', { text: h })))),
      el('tbody', null, T.byGroup.map((g) => el('tr', null, el('td', { text: g.group }), el('td', { text: g.name }), el('td', { text: g.inc }),
        el('td', { text: g.ot }), el('td', { text: g.total, style: 'font-weight:600' }), el('td', { text: g.engineers }))),
      el('tr', { class: 'total' }, el('td', { text: 'TOTAL' }), el('td'), el('td', { text: T.inc }), el('td', { text: T.ot }), el('td', { text: T.total }), el('td')))));

    const gh = T.groups.flatMap((g) => [`${g} Inc.`, `${g} OTs`]);
    $('#tEngineers').replaceChildren(T.engineers.length ? el('table', { class: 'rep compact' },
      el('thead', null, el('tr', { class: 'weeks' }, ['Ingeniero (Restoration user name)', ...gh, 'Total', 'Median time to restore', 'Also closed by them', 'First', 'Last'].map((h) => el('th', { text: h })))),
      el('tbody', null, T.engineers.map((e) => el('tr', { class: e.noUser ? 'unidentified' : null },
        el('td', null, el('button', { type: 'button', class: 'tid', style: 'font-family:inherit', text: e.user, title: 'Show this engineer’s tickets',
          onclick: () => { tState.engineer = e.user; $('#tEngineer').value = e.user; tState.limit = 200; renderTeamList(); $('#tList').scrollIntoView({ behavior: 'smooth', block: 'start' }); } })),
        ...T.groups.flatMap((g) => [e.byGroup[g].inc, e.byGroup[g].ot]).map((v) => el('td', { class: v === 0 ? 'zero' : null, text: v })),
        el('td', { text: e.total, style: 'font-weight:600' }),
        el('td', { text: fmtH(e.medianHours) }),
        el('td', { text: e.closedByThem }),
        el('td', { text: Number.isFinite(e.first) ? C.fmtDate(e.first) : '—' }),
        el('td', { text: Number.isFinite(e.last) ? C.fmtDate(e.last) : '—' }))),
      el('tr', { class: 'total' }, el('td', { text: 'TOTAL' }),
        ...T.groups.flatMap((g) => [T.engineers.reduce((a, e) => a + e.byGroup[g].inc, 0), T.engineers.reduce((a, e) => a + e.byGroup[g].ot, 0)]).map((v) => el('td', { text: v })),
        el('td', { text: T.total }), el('td', { text: fmtH(T.medianHours) }), el('td', { text: T.engineers.reduce((a, e) => a + e.closedByThem, 0) }), el('td'), el('td'))))
      : el('p', { class: 'muted', text: 'No tickets restored by the team in this period.' }));

    $('#tAllGroups').replaceChildren(el('table', { class: 'rep compact' },
      el('thead', null, el('tr', { class: 'weeks' }, ['Restoration group', 'Nombre', 'Incidencias', 'OTs', 'Total', 'Users'].map((h) => el('th', { text: h })))),
      el('tbody', null, T.restorationGroups.map((g) => el('tr', { style: g.sortis ? 'font-weight:600' : null },
        el('td', { text: g.sortis ? `${g.group} ★` : g.group }), el('td', { text: g.name }), el('td', { text: g.inc }), el('td', { text: g.ot }),
        el('td', { text: g.total }), el('td', { text: g.users }))))));

    const prev = tState.engineer;
    $('#tEngineer').replaceChildren(el('option', { value: '', text: 'All engineers' }), ...T.engineers.map((e) => el('option', { value: e.user, text: `${e.user} (${e.total})` })));
    tState.engineer = T.engineers.some((e) => e.user === prev) ? prev : '';
    $('#tEngineer').value = tState.engineer;
    renderTeamList();
    renderTeamCharts();
  }

  function teamEngineerSeries(T) {
    const named = T.engineers;
    return {
      labels: named.map((e) => e.user),
      series: T.groups.flatMap((g) => [
        { name: `${g} Incidencias`, data: named.map((e) => e.byGroup[g].inc) },
        { name: `${g} OTs`, data: named.map((e) => e.byGroup[g].ot) },
      ]).filter((x) => x.data.some((v) => v)),
    };
  }

  function renderTeamCharts() {
    if (!state.team) return;
    if (document.querySelector('[data-panel="team"]').hidden) { state.teamChartsDirty = true; return; }
    state.teamChartsDirty = false;
    const T = state.team;
    const es = teamEngineerSeries(T);
    makeChart('tchEngineers', es.labels, es.series.length ? es.series : [{ name: 'Casos', data: es.labels.map(() => 0) }], { horizontal: true });
    makeChart('tchTrend', T.buckets.map((b) => b.label), T.trend.length ? T.trend : [{ name: 'Casos', data: T.buckets.map(() => 0) }], {});
  }

  function teamListFiltered() {
    const T = state.team;
    const q = tState.q.trim().toLowerCase();
    return T.team.filter((t) => (!tState.engineer || (t.restorationUser || '(sin usuario de restauración)') === tState.engineer) &&
      (!tState.type || t.category === tState.type) &&
      (!q || `${t.id} ${t.action} ${t.description} ${t.closureUser} ${t.restorationUser} ${t.groupName} ${t.userName}`.toLowerCase().includes(q)));
  }

  function renderTeamList() {
    if (!state.team) return;
    const list = teamListFiltered();
    const shown = list.slice(0, tState.limit);
    $('#tListLabel').textContent = `(${list.length} of ${state.team.total})`;
    const head = ['Ticket ID', 'Restoration date', 'Creation date', 'Type', 'Prio', 'Status', 'Restoration group', 'Restored by', 'Time to restore', 'Closed by', 'Closure group(s)', 'Opened by (group)', 'Current action'];
    $('#tList').replaceChildren(...[
      el('table', { class: 'data-table' },
        el('thead', null, el('tr', null, head.map((h) => el('th', { text: h, style: 'cursor:default' })))),
        el('tbody', null, shown.length ? shown.map((t) => {
          const h = t.restored !== null && t.restored >= t.created ? (t.restored - t.created) / 3600000 : null;
          return el('tr', null,
            el('td', null, tidButton(t.id)),
            el('td', { text: t.restored !== null ? C.fmtDateTime(t.restored) : '—' }),
            el('td', { text: C.fmtDateTime(t.created) }),
            el('td', { text: t.category === 'INC' ? 'Incidencia' : 'OT' }),
            el('td', { text: t.priority }),
            el('td', { text: t.status }),
            el('td', { text: t.restorationGroupId }),
            el('td', { text: t.restorationUser || '—' }),
            el('td', { text: hoursTxt(h) }),
            el('td', { text: t.closureUser || '—' }),
            el('td', { text: t.closureGroups.join(' / ') || '—', title: t.closureGroups.join(' / ') }),
            el('td', { text: [t.groupId, t.groupName].filter(Boolean).join(' · '), title: t.userName }),
            el('td', { text: t.action || '—', title: t.action }));
        }) : el('tr', null, el('td', { colspan: head.length, class: 'muted', text: 'No tickets match the filters.' })))),
      list.length > shown.length ? el('button', { type: 'button', class: 'btn small more', onclick: () => { tState.limit += 400; renderTeamList(); } },
        `Show more (${list.length - shown.length} remaining)`) : null,
    ].filter(Boolean));
  }

  function exportTeam() {
    if (!state.team) return;
    try {
      const list = teamListFiltered();
      const head = ['Ticket ID', 'Restoration date', 'Creation date', 'Ticket type', 'Incidencia / OT', 'Processing priority', 'Status',
        'Restoration group ID', 'Restoration group abbreviation name', 'Restoration user name', 'Hours to restore', 'Closure user name', 'Closure group(s)',
        'Initiator - Group ID', 'Initiator - Group abbreviation name', 'Initiator - User name', 'Current action', 'Third party reference', 'Final nature', 'Short label', 'Description'];
      const aoa = [head];
      for (const t of list) {
        const h = t.restored !== null && t.restored >= t.created ? Math.round(((t.restored - t.created) / 3600000) * 10) / 10 : null;
        aoa.push([t.id, t.restored !== null ? C.toExcelSerial(t.restored) : null, C.toExcelSerial(t.created), t.type, t.category === 'INC' ? 'Incidencia' : 'OT',
          t.priority, t.status, t.restorationGroupId, t.restorationGroupName, t.restorationUser, h, t.closureUser, t.closureGroups.join(' / '),
          t.groupId, t.groupName, t.userName, t.action, t.thirdParty, t.finalNature, t.shortLabel, t.description.slice(0, 32000)]);
      }
      const ws = XLSX.utils.aoa_to_sheet(aoa);
      for (let r = 1; r < aoa.length; r++) for (const c of [1, 2]) { const ref = XLSX.utils.encode_cell({ r, c }); if (ws[ref]) ws[ref].z = 'yyyy-mm-dd hh:mm'; }
      ws['!cols'] = head.map((h) => ({ wch: Math.min(40, Math.max(12, h.length + 2)) }));
      ws['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: aoa.length - 1, c: head.length - 1 } }) };
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Sortis team');
      const T = state.team;
      const eh = ['Ingeniero', ...T.groups.flatMap((g) => [`${g} Inc.`, `${g} OTs`]), 'Total', 'Median hours to restore', 'Also closed by them'];
      const ea = [eh, ...T.engineers.map((e) => [e.user, ...T.groups.flatMap((g) => [e.byGroup[g].inc, e.byGroup[g].ot]), e.total,
        e.medianHours === null ? null : Math.round(e.medianHours * 10) / 10, e.closedByThem])];
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(ea), 'By engineer');
      XLSX.writeFile(wb, `sortis-team-${safeFilePart(T.label)}.xlsx`, { compression: true });
      toast(`Excel downloaded (${list.length} tickets).`);
    } catch (err) {
      console.error(err);
      toast(`Could not create the Excel file: ${err && err.message ? err.message : err}`, true);
    }
  }

  function updateWordButton() {
    const weeklyFail = state.report && state.report.checks.some((c) => !c.ok);
    const periodFail = $('#pInWord').checked && state.period && state.period.checks.some((c) => !c.ok);
    const teamFail = $('#secTeam').checked && state.team && state.team.checks.some((c) => !c.ok);
    $('#btnWord').disabled = !!(weeklyFail || periodFail || teamFail);
    $('#btnWord').title = weeklyFail || periodFail || teamFail ? 'Blocked: the report did not pass the consistency checks' : '';
  }

  function initSectionPicker() {
    $('#secPeriod').addEventListener('change', () => { $('#pInWord').checked = $('#secPeriod').checked; updateWordButton(); });
    $('#secTeam').addEventListener('change', () => { $('#tInWord').checked = $('#secTeam').checked; updateWordButton(); });
  }

  function wordSections() {
    const inc = {};
    $$('[data-sec]').forEach((cb) => { inc[cb.dataset.sec] = cb.checked; });
    return inc;
  }

  /* ------------------------------------------------------------------ */
  /* Rendering                                                           */
  /* ------------------------------------------------------------------ */

  function renderAll() {
    renderPeriod();
    renderChecks();
    renderKpis();
    renderWeekly();
    renderGestor();
    renderPending();
    renderCharts();
    applyDataFilters();
  }

  function renderPeriod() {
    const r = state.report;
    const w = r.reportWeek;
    const parts = [
      el('span', { class: 'pb-main', text: `Semana ${w.week} · ${C.fmtDate(w.start)} – ${C.fmtDate(w.end - 86400000)} · ${r.periodLabel}` }),
      el('span', { class: 'pb-sub', text: r.weeks.length > 1 ? `Tables show semanas ${r.weeks.map((x) => x.week).join(', ')}` : 'Tables show the report week only' }),
      el('span', { class: 'pb-sub', text: `Data in file until ${C.fmtDateTime(state.summary.dataUntil)}` }),
    ];
    if (r.partialWeek) parts.push(el('span', { class: 'pb-warn', text: `⚠ The file ends before this week is over — figures are as of ${C.fmtDateTime(r.asOf)}.` }));
    $('#periodBanner').replaceChildren(...parts);
  }

  function renderChecks() {
    const checks = state.report.checks;
    const failed = checks.filter((c) => !c.ok);
    const box = $('#checksBox');
    box.hidden = false;
    box.classList.toggle('fail', failed.length > 0);
    $('#checksToggle').textContent = failed.length
      ? `✖ ${failed.length} of ${checks.length} consistency checks failed — Word export is blocked until this is fixed (click for details)`
      : `✔ All ${checks.length} consistency checks passed (click for details)`;
    $('#checksList').replaceChildren(...checks.map((c) => el('li', { class: c.ok ? 'ok' : 'bad' },
      el('span', { class: 'ic', text: c.ok ? '✔' : '✖' }),
      el('span', null, c.label, c.detail ? el('span', { class: 'det', text: ` — ${c.detail}` }) : null))));
    if (failed.length) { $('#checksList').hidden = false; $('#checksToggle').setAttribute('aria-expanded', 'true'); }
    updateWordButton();
  }

  function renderKpis() {
    const r = state.report;
    const i = r.weeks.length - 1;
    const x = r.inc[i], o = r.ot[i];
    const wk = `semana ${r.reportWeek.week}`;
    const newAll = x.nuevos + o.nuevos;
    const resolved = x.resueltas + o.resueltas;
    const escalated = r.pending.filter((t) => t.escalated).length;
    const kpis = [
      ['New incidencias', x.nuevos, wk],
      ['New OTs', o.nuevos, wk],
      ['Resolved (of new)', resolved, `${newAll ? Math.round((resolved / newAll) * 100) : 0}% of the ${newAll} new cases`],
      ['Unresolved at end of week', r.pending.length, `as of ${C.fmtDateTime(r.asOf)}`],
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
    $('#weeklyNote').textContent = `Semanas ${r.weeks[0].week}–${r.reportWeek.week}: del ${C.fmtDate(r.weeks[0].start)} al ${C.fmtDate(r.cutoff - 86400000)}. Each week shows the situation at the end of that week.`;
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
        r.gestores.map((g) => el('tr', { class: g.unidentified ? 'unidentified' : null },
          el('td', { text: g.gestor, title: g.unidentified ? 'Show these tickets' : null,
            onclick: g.unidentified ? () => $('#noGestorPanel').scrollIntoView({ behavior: 'smooth', block: 'start' }) : null }),
          g.counts.map((v) => el('td', { class: v === 0 ? 'zero' : null, text: v })),
          el('td', { text: g.total, style: 'font-weight:600' }))),
        el('tr', { class: 'total' }, el('td', { text: 'TOTAL' }), totals.map((v) => el('td', { text: v })),
          el('td', { text: totals.reduce((a, b) => a + b, 0) }))));
    box.replaceChildren(table);
    const notes = [`Totals match “Nuevos durante la semana” (Incidencias) in 2.1.`];
    if (r.gestorUnidentified) notes.push(`Across the ${r.weeks.length} week(s) shown, ${r.gestorUnidentified} incidencia(s) have no gestor in “Current action” (e.g. “JR - Trabajando”, empty or DEVUELTO) and are shown as “Sin gestor identificado”.`);
    if (r.gestorInferred) notes.push(`${r.gestorInferred} gestor(s) were taken from the ticket description (option enabled); this is stated in the Word report.`);
    $('#gestorNote').textContent = notes.join(' ');
    renderNoGestorPanel('weekly', 'noGestorPanel', r.noGestorTickets,
      `in semanas ${r.weeks[0].week}–${r.reportWeek.week} (incidencias)`);
  }

  function renderPending() {
    const r = state.report;
    const f = state.pendingFilter;
    const list = r.pending.filter((t) => f === 'all' || t.category === f);
    $('#pendingCount').textContent = `${list.length} case(s) unresolved at the end of semana ${r.reportWeek.week} (${C.fmtDateTime(r.asOf)})`;
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
          el('span', { class: 'pcard-id' }, tidButton(t.id)),
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
    makeChart('chWeekly', weeklyLabels(r), weeklySeries(r), {});
    $('#monthlyNote').textContent = C.weekSplitText(r);
    for (const [id, key, split] of CHART_DEFS) makeChart(id, r.monthLabels, r.monthly[key], { splitStacks: split });
  }

  /** Create (or re-create) a chart and the table of its numbers under the legend. */
  function makeChart(id, labels, series, opts) {
    if (state.charts[id]) state.charts[id].destroy();
    state.charts[id] = new Chart(document.getElementById(id), CH.barConfig(labels, series, { dark: isDark(), ...opts }));
    if (!opts.horizontal) renderChartTable(id, labels, series);
  }

  function renderChartTable(canvasId, labels, series) {
    const fig = document.getElementById(canvasId).closest('.chart-card');
    let box = fig.querySelector('.chart-table');
    if (!box) { box = el('div', { class: 'chart-table' }); fig.appendChild(box); }
    const t = CH.chartTable(labels, series);
    const multi = labels.length > 1;
    box.replaceChildren(el('table', null,
      el('thead', null, el('tr', null, el('th', { text: 'Casos' }), labels.map((l) => el('th', { text: l })), multi ? el('th', { text: 'Total' }) : null)),
      el('tbody', null, t.rows.map((r, i) => el('tr', { class: r.kind === 'series' ? null : r.kind },
        el('td', null, r.kind === 'series' ? el('span', { class: 'sw', style: `background:${CH.colorFor(r.name, i)}` }) : null, r.name),
        r.data.map((v) => el('td', { class: v === 0 ? 'zero' : null, text: v })),
        multi ? el('td', { text: r.total, style: 'font-weight:600' }) : null)))));
  }

  function weeklyLabels(r) {
    return r.weeks.map((w) => `sem. ${w.week} (${C.fmtDate(w.start).slice(0, 5)})`);
  }

  function weeklySeries(r) {
    return [
      { name: 'Incidencias', data: r.inc.map((x) => x.nuevos) },
      { name: 'OTs', data: r.ot.map((x) => x.nuevos) },
    ];
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
      if (tab.dataset.tab === 'period' && state.period && (state.periodChartsDirty || !state.charts.pchOpened)) renderPeriodCharts();
      if (tab.dataset.tab === 'team' && state.team && (state.teamChartsDirty || !state.charts.tchEngineers)) renderTeamCharts();
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
    { key: 'weekYear', label: 'Year of the week (ISO)', num: true },
    { key: 'priority', label: 'Processing priority' },
    { key: 'category', label: 'Failure / OT', fmt: (v) => (v === 'INC' ? 'Failure' : 'OT') },
    { key: 'type', label: 'Ticket type' },
    { key: 'status', label: 'Status' },
    { key: 'groupId', label: 'Initiator - Group ID' },
    { key: 'groupName', label: 'Initiator - Group abbreviation name' },
    { key: 'action', label: 'Current action' },
    { key: 'gestor', label: 'Gestor' },
    { key: 'gestorDesc', label: 'Gestor (from description)' },
    { key: 'problema', label: 'Problema' },
    { key: 'tecnico', label: 'Técnico' },
  ];

  // Every other column of the Excel (hidden by default in the table; all exported to Excel).
  const listFmt = (v) => (Array.isArray(v) ? v.join(' / ') : v);
  const DETAIL_COLUMNS = [
    { key: 'userName', label: 'Initiator - User name' },
    { key: 'ackDate', label: 'Acknowledgement date', fmt: (v) => C.fmtDateTime(v), date: true },
    { key: 'identifier1', label: 'Identifier 1' },
    { key: 'identifier2', label: 'Identifier 2' },
    { key: 'identifier3', label: 'Identifier 3' },
    { key: 'identifier4', label: 'Identifier 4' },
    { key: 'restored', label: 'Restoration date', fmt: (v) => C.fmtDateTime(v), date: true },
    { key: 'duration', label: 'Calculated ticket duration', num: true },
    { key: 'shortLabel', label: 'Short label' },
    { key: 'finalNature', label: 'Final nature' },
    { key: 'restorationGroupId', label: 'Restoration group ID' },
    { key: 'restorationGroupName', label: 'Restoration group abbreviation name' },
    { key: 'restorationUser', label: 'Restoration user name' },
    { key: 'closureGroups', label: 'Closure group(s)', fmt: listFmt },
    { key: 'closureUser', label: 'Closure user name' },
    { key: 'description', label: 'Description' },
    { key: 'sheets', label: 'Found in sheet(s)', fmt: listFmt },
  ];
  const ALL_COLUMNS = [...DATA_COLUMNS, ...DETAIL_COLUMNS];
  const DEFAULT_VISIBLE = DATA_COLUMNS.map((c) => c.key).concat(['restorationGroupId', 'restorationUser', 'closureUser']);

  function visibleColumns() {
    return ALL_COLUMNS.filter((c) => state.data.cols.has(c.key));
  }

  function loadVisibleColumns() {
    let keys = DEFAULT_VISIBLE;
    try {
      const saved = JSON.parse(localStorage.getItem('slm-cols') || 'null');
      if (Array.isArray(saved) && saved.length) keys = saved.filter((k) => ALL_COLUMNS.some((c) => c.key === k));
    } catch (e) { /* ignore */ }
    if (!keys.includes('id')) keys = ['id', ...keys];
    state.data.cols = new Set(keys);
  }

  function saveVisibleColumns() {
    try { localStorage.setItem('slm-cols', JSON.stringify(Array.from(state.data.cols))); } catch (e) { /* ignore */ }
  }

  function renderColumnPicker() {
    const box = $('#colPick');
    const setAll = (keys) => { state.data.cols = new Set(['id', ...keys]); saveVisibleColumns(); renderColumnPicker(); renderDataHeader(); applyDataFilters(); };
    box.replaceChildren(
      el('div', { class: 'cp-actions' },
        el('button', { type: 'button', class: 'btn small', onclick: () => setAll(ALL_COLUMNS.map((c) => c.key)) }, 'All'),
        el('button', { type: 'button', class: 'btn small', onclick: () => setAll(DEFAULT_VISIBLE) }, 'Default')),
      ...ALL_COLUMNS.map((c) => {
        const cb = el('input', { type: 'checkbox', checked: state.data.cols.has(c.key), disabled: c.key === 'id' });
        cb.addEventListener('change', () => {
          if (cb.checked) state.data.cols.add(c.key); else state.data.cols.delete(c.key);
          saveVisibleColumns(); renderDataHeader(); applyDataFilters();
        });
        return el('label', null, cb, c.label);
      }));
  }

  function renderDataHeader() {
    $('#dataTable thead tr').replaceChildren(...visibleColumns().map((c) => el('th', {
      'data-key': c.key, tabindex: 0, scope: 'col', title: 'Sort',
      onclick: () => sortBy(c.key),
      onkeydown: (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); sortBy(c.key); } },
    }, c.label, el('span', { class: 'arrow' }))));
  }

  /** A Ticket ID that opens the detail window. */
  function tidButton(id) {
    return el('button', { type: 'button', class: 'tid', 'data-tid': id, title: 'Show every detail of this ticket', text: id });
  }

  function initDataFilters() {
    const statuses = Array.from(new Set(state.tickets.map((t) => t.status))).sort();
    $('#fStatus').replaceChildren(el('option', { value: '', text: 'All statuses' }), ...statuses.map((s) => el('option', { value: s, text: s })));
    const years = Array.from(new Set(state.tickets.map((t) => new Date(t.created).getUTCFullYear()))).sort((a, b) => b - a);
    $('#fYear').replaceChildren(el('option', { value: '', text: 'All years' }), ...years.map((y) => el('option', { value: y, text: y })));
    const rgroups = Array.from(new Set(state.tickets.map((t) => t.restorationGroupId).filter(Boolean))).sort();
    $('#fRGroup').replaceChildren(el('option', { value: '', text: 'All restoration groups' }), el('option', { value: '-', text: '(no restoration group)' }),
      ...rgroups.map((g) => el('option', { value: g, text: g })));
    const rusers = Array.from(new Set(state.tickets.map((t) => t.restorationUser).filter(Boolean))).sort((a, b) => a.localeCompare(b, 'es'));
    $('#fRUser').replaceChildren(el('option', { value: '', text: 'All restoration users' }), el('option', { value: '-', text: '(no restoration user)' }),
      ...rusers.map((u) => el('option', { value: u, text: u })));
    Object.assign(state.data, { query: '', type: '', status: '', vendor: '', year: '', rgroup: '', ruser: '', page: 0 });
    $('#dataSearch').value = '';
    ['#fType', '#fStatus', '#fVendor', '#fYear', '#fRGroup', '#fRUser'].forEach((s) => { $(s).value = ''; });
  }

  function initDataTable() {
    loadVisibleColumns();
    renderColumnPicker();
    renderDataHeader();
    let t;
    $('#dataSearch').addEventListener('input', () => {
      clearTimeout(t);
      t = setTimeout(() => { state.data.query = $('#dataSearch').value; state.data.page = 0; applyDataFilters(); }, 200);
    });
    [['#fType', 'type'], ['#fStatus', 'status'], ['#fVendor', 'vendor'], ['#fYear', 'year'], ['#fRGroup', 'rgroup'], ['#fRUser', 'ruser']].forEach(([sel, k]) => {
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
      if (d.rgroup === '-' && t.restorationGroupId) return false;
      if (d.rgroup && d.rgroup !== '-' && t.restorationGroupId !== d.rgroup) return false;
      if (d.ruser === '-' && t.restorationUser) return false;
      if (d.ruser && d.ruser !== '-' && t.restorationUser !== d.ruser) return false;
      if (q) {
        const hay = `${t.id} ${t.thirdParty} ${t.groupId} ${t.groupName} ${t.action} ${t.userName} ${t.status} ${t.type} ${t.restorationGroupId} ${t.restorationUser} ${t.closureUser} ${(t.closureGroups || []).join(' ')}`.toLowerCase();
        if (!hay.includes(q)) return false;
      }
      return true;
    });
    const k = d.sortKey;
    const dir = d.sortDir;
    list = list.slice().sort((a, b) => {
      const va = Array.isArray(a[k]) ? a[k].join(' / ') : a[k], vb = Array.isArray(b[k]) ? b[k].join(' / ') : b[k];
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
    const cols = visibleColumns();
    if (!rows.length) {
      tbody.replaceChildren(el('tr', null, el('td', { colspan: cols.length, class: 'muted', text: 'No tickets match the filters.' })));
    } else {
      tbody.replaceChildren(...rows.map((t) => el('tr', null, cols.map((c) => {
        if (c.key === 'id') return el('td', null, tidButton(t.id));
        const v = c.fmt ? c.fmt(t[c.key]) : t[c.key];
        const s = v === null || v === undefined ? '' : String(v);
        return el('td', { text: s, title: s.length > 40 ? s.slice(0, 1500) : null });
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
    if (!busy) updateWordButton();
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
    if (state.report.checks.some((c) => !c.ok)) { toast('The report did not pass the consistency checks — see the list above.', true); return; }
    const includePeriod = $('#pInWord').checked && state.period && state.period.hasData;
    if (includePeriod && state.period.checks.some((c) => !c.ok)) { toast('The Month / Year section did not pass its consistency checks.', true); return; }
    const includeTeam = $('#secTeam').checked && state.team && state.team.total > 0;
    if (includeTeam && state.team.checks.some((c) => !c.ok)) { toast('The Sortis team section did not pass its consistency checks.', true); return; }
    const sections = wordSections();
    if (!Object.values(sections).some(Boolean) && !includePeriod && !includeTeam) { toast('Select at least one section for the Word report.', true); return; }
    setBusy(true, 'Building the Word report…');
    await nextFrame();
    try {
      const r = state.report;
      const labels = r.monthLabels;
      const idx = r.months.length - 1;
      const monthTxt = C.monthLabel(r.reportMonth) + (r.partialMonth ? ` (hasta ${C.fmtDate(r.asOf)})` : '');
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
        charts[key] = { bar, pie, table: CH.chartTable(labels, series) };
        await nextFrame();
      }
      charts.weekly = {
        bar: CH.renderPng(CH.barConfig(weeklyLabels(r), weeklySeries(r), { static: true, title: `Casos nuevos por semana (semanas ${r.weeks[0].week}–${r.reportWeek.week})` }), 1000, 420),
        pie: null,
        table: CH.chartTable(weeklyLabels(r), weeklySeries(r)),
      };
      let periodCharts = null;
      let periodTables = null;
      if (includePeriod) {
        const p = state.period;
        const pl = p.buckets.map((b) => b.label);
        periodTables = {
          opened: CH.chartTable(pl, p.series.opened),
          resolved: CH.chartTable(pl, p.series.resolved),
          status: CH.chartTable(pl, p.series.status),
          vendor: CH.chartTable(pl, p.series.vendor),
        };
        periodCharts = {
          opened: CH.renderPng(CH.barConfig(pl, p.series.opened, { static: true, title: `Casos abiertos — ${p.label}` }), 1000, 440),
          resolved: CH.renderPng(CH.barConfig(pl, p.series.resolved, { static: true, title: `Casos resueltos — ${p.label}` }), 1000, 440),
          status: CH.renderPng(CH.barConfig(pl, p.series.status, { static: true, splitStacks: true, title: `Estado de los casos abiertos — ${p.label}` }), 1000, 440),
          vendor: CH.renderPng(CH.barConfig(pl, p.series.vendor, { static: true, title: `Casos escalados por fabricante — ${p.label}` }), 1000, 400),
          gestor: p.gestorChart.labels.length ? CH.renderPng(CH.barConfig(p.gestorChart.labels, p.gestorChart.series,
            { static: true, horizontal: true, title: `Casos por gestor (top 15) — ${p.label}` }), 1000, Math.max(320, 70 + p.gestorChart.labels.length * 34)) : null,
        };
        await nextFrame();
      }
      let teamCharts = null;
      let teamTables = null;
      if (includeTeam) {
        const T = state.team;
        const es = teamEngineerSeries(T);
        teamCharts = {
          engineers: es.series.length ? CH.renderPng(CH.barConfig(es.labels, es.series, { static: true, horizontal: true, title: `Casos restaurados por ingeniero — ${T.label}` }),
            1000, Math.max(300, 90 + es.labels.length * 40)) : null,
          trend: T.trend.length ? CH.renderPng(CH.barConfig(T.buckets.map((b) => b.label), T.trend, { static: true, title: `Evolución por ingeniero — ${T.label}` }), 1000, 460) : null,
        };
        teamTables = { trend: T.trend.length ? CH.chartTable(T.buckets.map((b) => b.label), T.trend) : null };
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
        dataUntil: state.summary.dataUntil,
      };
      if (includePeriod) meta.period = { data: state.period, note: periodNote(state.period), gestorCaption: gestorCaption(state.period) };
      if (includeTeam) meta.team = { data: state.team, note: teamNote(state.team) };
      meta.include = sections;
      const doc = SLMDocx.buildDocument(window.docx, C, r, meta, { logo, charts, periodCharts, periodTables, teamCharts, teamTables });
      const blob = await docx.Packer.toBlob(doc);
      const name = `INFORME-SLM-OSS-Sortis-${r.reportWeek.year}-Semana${String(r.reportWeek.week).padStart(2, '0')}-${safeFilePart(r.periodLabel.replace(/ \/ /g, '-'))}.docx`;
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
      for (const c of DETAIL_COLUMNS) if (c.key !== 'restored') cols.push(c.date ? { key: c.key, label: c.label, date: true } : c);
      const aoa = [cols.map((c) => c.label)];
      // export what is currently filtered in the "Extracted data" tab (all tickets when no filter)
      const list = hasDataFilter() ? state.data.filtered : state.tickets;
      for (const t of list) {
        aoa.push(cols.map((c) => {
          const v = t[c.key];
          if (c.date) return v === null || v === undefined ? null : C.toExcelSerial(v);
          if (c.fmt) return c.fmt(v);
          if (typeof v === 'string' && v.length > 32000) return v.slice(0, 32000);
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
    return !!(d.query.trim() || d.type || d.status || d.vendor || d.year || d.rgroup || d.ruser);
  }

  /** Second sheet with the weekly report tables. */
  function appendReportSheet(wb) {
    const r = state.report;
    if (!r) return;
    const wk = r.weeks.map((w) => `sem. ${w.week}`);
    const aoa = [[`Informe SLM-OSS · Semana ${r.reportWeek.week} de ${r.reportWeek.year} (${C.fmtDate(r.reportWeek.start)} – ${C.fmtDate(r.reportWeek.end - 86400000)}) · ${r.periodLabel}`], []];
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
    initPeriodControls();
    initTeamControls();
    initTicketDialog();
    initSectionPicker();
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
  else boot();
})();
