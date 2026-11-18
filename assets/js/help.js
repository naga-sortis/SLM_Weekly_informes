/* SLM Weekly Informes — help topics, "?" buttons and the guided tour */
(function () {
  'use strict';

  /**
   * Every section of the page explained: what it is, how to use it, and how the numbers are made.
   * Keys are used by data-help="…" attributes in the page and by the guided tour.
   */
  const TOPICS = {
    overview: {
      title: 'How this tool works',
      intro: 'Every week Orange sends two Oceane / BRISE exports. This page reads them, merges them, checks the figures and writes the SLM-OSS informe (Word).',
      steps: [
        'Upload the two weekly Excel files (step 1): the Incidencias file (XSP00025) and the OTs file (XSP00027 “USERS”). One file also works, as before.',
        'Choose the report week (step 2). The week, its dates and its month are used everywhere — on the page and in the Word report.',
        'Open the Manual review tab and review the tickets the tool cannot decide alone: suspected types, tickets resolved by another group and tickets without gestor.',
        'Check that all consistency checks are green (step 2). If one fails, the Word export is blocked until the figures add up.',
        'Download the Word report. Optionally download the extracted data (Excel) to cross-check.',
      ],
      points: [
        'On the server, save the files to the team database so the history of every week is kept and the manual reviews are shared with the team.',
        'Click any “?” on the page for the help of that section, or start the guided tour from the “Help” button at the top.',
      ],
    },
    files: {
      title: 'Step 1 · The weekly files',
      intro: 'Drop one or two Excel files on the upload box (or click “browse” and select both). The kind of each file is detected from its content.',
      steps: [
        'Incidencias file — mostly Failure tickets (e.g. “Tickets - SLM SORTIS-WEEK_XSP00025…xlsx”).',
        'OTs file — mostly work orders (e.g. “Tickets - SLM SORTIS-WEEK_…USERS.xlsx”, queue XSP00027).',
        'If the detection is wrong, change the kind of the file in the list below the upload box.',
        'Dropping a new file replaces the loaded file of the same kind (e.g. next week’s Incidencias file).',
      ],
      points: [
        'The two files are merged by Ticket ID: a ticket that appears in both files (or in several sheets) is counted once.',
        'Suspected tickets: a work order (or any non-Failure) found only in the Incidencias file, and a Failure found only in the OTs file. They are counted by the usual rules until somebody reviews them in the Manual review tab.',
        'With one file only, everything works exactly as before (the largest ticket sheet is used and every ticket is counted).',
        'The files are read by your browser. They are only sent to the server if you choose “Update the database”.',
      ],
    },
    database: {
      title: 'Team database (server only)',
      intro: 'When the page runs on the team server, the weekly files can be saved in a shared database, together with every manual review.',
      steps: [
        'Type your name in “Your name” — it is recorded with every import and every manual review.',
        'After uploading the files, a preview shows what would change: new tickets, updated tickets (each changed field is logged) and tickets that are no longer in the files (they are kept).',
        'Click “Update the database with both files” to save them. “Use database for the reports” builds the report from the whole history.',
        '“Import history” lists every import; “Backup” downloads a copy of the database.',
      ],
      points: [
        'Manual reviews (Incidencia / OT and “Resolved by”) are shared: what one person decides is seen by everybody, with name and date.',
        'In file mode (e.g. GitHub Pages) the database is not available and the manual reviews are kept in your browser only.',
      ],
    },
    activity: {
      title: 'Team activity (live)',
      intro: 'What everybody does on the team server — file imports, Incidencia / OT reviews and “Resolved by” changes — with the name and the time, updated every 10 seconds.',
      steps: [
        'Click “Team activity” (next to Backup) to open the list; the orange number counts what other users did since you last looked.',
        'Tick “Other users only” or choose a kind of activity to filter. Click a Ticket ID to open the ticket.',
        'When somebody else reviews a ticket, your page reloads the shared reviews and recalculates the figures by itself (a short message tells you who changed what).',
        'When somebody imports new files, a yellow bar offers to show the updated database — you decide when, so your work is not interrupted.',
      ],
      points: ['“● Live · synced hh:mm:ss” shows the last synchronisation. If the server cannot be reached it is retried automatically. In file mode (GitHub Pages) there is no shared activity.'],
    },
    settings: {
      title: 'Step 2 · Report week & settings',
      intro: 'These settings drive every table, chart and the Word report.',
      steps: [
        'Report week — the week of the informe (default: the latest week with tickets). Nothing created or resolved after the end of this week is counted, so a past week can be regenerated with the same figures.',
        'Weeks shown in tables — how many weeks (ending with the report week) the tables show.',
        'Trend charts — the last N months up to the report week, or one calendar year.',
        'SLM / Sortis group IDs — the Sortis queues (XSP00025, XSP00027). Used by the counting rule, the “Resolved by” review, the row “Abiertas por el SLM” and the Sortis team tab.',
        'Gestor from the description — optional: when “Current action” has no gestor, take the gestor named in the description (flagged in the report).',
      ],
      points: ['Weeks are ISO weeks (Monday–Sunday) calculated from the Creation date.'],
    },
    countRule: {
      title: 'Counting rule: only cases resolved by the Sortis team',
      intro: 'When it is on, a ticket is counted in the reports if it is still open, was returned (DEVUELTO), or was resolved by one of the Sortis groups (Restoration group ID XSP00025 / XSP00027).',
      steps: [
        'It is switched on automatically when both files are loaded, and off with a single file (as before). You can change it at any time.',
        'Tickets resolved by another group are not counted. They are listed in the Manual review tab under “Resolved by another group”.',
        'If the Sortis team really resolved one of them (another group took it over in Oceane), change its “Resolved by” to XSP00025 or XSP00027 with a comment — it is then counted.',
      ],
      points: ['A consistency check confirms that no case resolved by another group is counted.'],
    },
    checks: {
      title: 'Consistency checks',
      intro: 'Every report is verified before it can be sent to the customer. If a check fails, the Word export is blocked.',
      points: [
        'Nuevos = Resueltas + Sin resolver (no esc.) + Sin resolver (esc.) + Devueltas, for every week and both types.',
        'Backlog = Resueltos + Sin resolver, and each week’s backlog = the previous week’s pending.',
        '2.2 totals per week = “Nuevos durante la semana” (Incidencias); each gestor appears once.',
        '2.3 TOTAL = the cases listed in 2.4; every listed case was really pending at the cut-off.',
        'Each ticket is counted once (tickets in both files are merged); with the counting rule on, no case resolved by another group is counted; every manual “Resolved by” change has a comment.',
        'Monthly charts cover every case of the month exactly once; the report month contains the report week.',
      ],
    },
    word: {
      title: 'Download the Word report',
      intro: 'Builds the SLM-OSS informe with the report week, the tables, the charts and the selected sections.',
      steps: [
        'Open “Word report details” to fill the cover (title, revision, author), the highlights (2.5) and other works (2.6).',
        'Choose the sections to include — the numbering adapts automatically.',
        '“Anexo: ajustes manuales” (optional) lists the tickets of the report weeks reviewed manually, with the change, the reason, who and when. Leave it unticked if the customer should not see it.',
      ],
      points: ['The button is disabled while a consistency check fails.'],
    },
    sources: {
      title: 'Source sheets',
      intro: 'What each sheet of the uploaded files contains, before any report rule is applied — check it first.',
      points: [
        'File — the file each sheet comes from (Incidencias file / OTs file).',
        'Rows, unique tickets and repeated rows; creation dates from – to; Incidencias and OTs; ticket types and statuses; missing columns.',
        '“Tickets shared between sheets” shows how many tickets two sheets have in common (e.g. tickets present in both files).',
        'At the bottom you can browse, search and export the rows of any sheet exactly as in the source.',
      ],
    },
    weekly: {
      title: '2.1 Casos tratados & 2.3 Casos sin resolver',
      intro: 'The weekly tables of the informe, for incidencias and OTs, week by week.',
      points: [
        'Incidencias = ticket type Failure; OTs = every other type. A ticket whose Current action is a technician status (e.g. “JR - Trabajando”) is an OT. A manual review (Incidencia / OT) overrides the rules.',
        'Resuelta = status not Current and Restoration date before the end of the week (past weeks show the situation at the time).',
        'Escalada = Third party reference STA- / CSR (Ericsson), H- (Huawei), 1- (Nokia). Devuelta = Current action DEVUELTO.',
        'Backlog = cases created before the week, not returned and still unresolved when the week starts.',
        'With the counting rule on, only cases resolved by XSP00025 / XSP00027 (or still open / returned) are counted.',
      ],
    },
    gestor: {
      title: '2.2 Nº casos por gestor',
      intro: 'Every new incidencia of each week, under the gestor taken from “Current action” (GESTOR - PROBLEMA - TÉCNICO).',
      points: [
        'Spelling variants are merged (NFM-T / NFMT, ENM 3 / ENM3…). The weekly totals always equal “Nuevos durante la semana” of 2.1.',
        'Technician statuses (“JR - Trabajando”) are OTs and have their own row “OTs (técnico trabajando)”.',
        '“Sin gestor identificado” — incidencias whose action has no gestor. The panel below lists them with the reason; classify each one as Incidencia or OT and, if needed, change its “Resolved by”.',
      ],
    },
    nogestor: {
      title: 'Sin gestor identificado',
      intro: 'Incidencias of the period whose “Current action” does not name a gestor, with the reason of each one.',
      steps: [
        'Filter by reason or by review status, or search.',
        'Click Incidencia or OT to classify a case after checking it (click the Ticket ID to see every detail). It is recorded with your name and the date.',
        '“Resolved by” — click “Change…” if the ticket was resolved by XSP00025 / XSP00027 although Oceane shows another group (a comment is required).',
        'Undo restores the automatic rule.',
      ],
    },
    pending: {
      title: '2.4 Detalle de los casos sin resolver',
      intro: 'Every case still unresolved at the end of the report week, as listed in the Word report (with its queue and third-party reference).',
    },
    monthly: {
      title: 'Charts (section 3)',
      intro: 'New cases per week, then the monthly trends: cases opened, resolved, escalated (by priority, gestor and vendor) and their status.',
      points: ['Each chart shows its numbers under the bars, exactly as in the Word report. A week that spans two months is split by creation date.'],
    },
    period: {
      title: 'Month / Year summary',
      intro: 'All incidencias and OTs created in a month or a year: opened, resolved, pending, returned and escalated, with charts and the cases by gestor.',
      points: ['Tick “Include in the Word report” to add it as an extra section. It follows the same counting rule and manual reviews as the weekly report.'],
    },
    team: {
      title: 'Sortis team',
      intro: 'Tickets restored by the Sortis queues (Restoration group XSP00025 / XSP00027) from all sheets of both files, by engineer (Restoration user name).',
      points: [
        'A manual “Resolved by” change is taken into account here too: a ticket taken over by another group but resolved by the team appears under the chosen queue (and engineer, if given).',
        'Choose the year / month, the date used (restoration or creation) and the queue. The table at the bottom shows every restoration group of the period.',
      ],
    },
    data: {
      title: 'Extracted data',
      intro: 'Every ticket with every field, searchable and filterable. Click a Ticket ID to see all the columns of the Excel and who dealt with it.',
      points: [
        'Use “Columns” to show more fields: source file(s), suspected type, restoration group of the Excel, manual “Resolved by”, counted in the report…',
        '“Export ALL data” adds verification sheets (original rows, checks and how the files were read) to cross-check the analysis.',
      ],
    },
    review: {
      title: 'Manual review',
      intro: 'The tickets the tool cannot decide alone. Review them before downloading the Word report.',
      steps: [
        'Choose the period (by default the weeks shown in the tables) and keep “Pending review” selected.',
        'Suspected type — a non-Failure ticket only in the Incidencias file, or a Failure only in the OTs file. Open it (Ticket ID), then click Incidencia or OT.',
        'Resolved by another group — closed by a group other than XSP00025 / XSP00027, so not counted. If the Sortis team really resolved it, click “Change…” and choose XSP00025 or XSP00027, the engineer (optional) and explain why. If another group really resolved it, choose “Another group really resolved it” to mark it as checked.',
        'Sin gestor identificado — incidencias without gestor in Current action: classify them as Incidencia or OT.',
        'Bulk buttons apply the same decision to every pending ticket of the filtered list (each one is recorded).',
      ],
      points: [
        'Every decision is saved with your name and the date (on the server it is shared with the team and kept in the history of the ticket).',
        'Until a ticket is reviewed it is counted by the rules. “Export” downloads the list with every decision and comment.',
      ],
    },
    resolver: {
      title: 'Change “Resolved by”',
      intro: 'State which group really resolved the ticket. Used when another group took the ticket over in Oceane although the Sortis team resolved it.',
      steps: [
        'Choose XSP00025 or XSP00027 — the ticket is counted as resolved by that queue (in the reports and in the Sortis team tab).',
        'Optionally choose the engineer who resolved it.',
        'Explain the reason in the comment (required): it is kept with your name and the date, and shown in the review list, the ticket history, the Excel export and the optional Word annex.',
        'Or choose “Another group really resolved it” to mark the ticket as checked — it stays not counted.',
        '“Back to the Excel value” removes the manual change.',
      ],
    },
  };

  const TOUR = [
    { sel: '#uploadCard', topic: 'files' },
    { sel: '#dbBox', topic: 'database' },
    { sel: '#btnActivity', topic: 'activity' },
    { sel: '#settingsCard .settings-grid', topic: 'settings' },
    { sel: '#onlySortisField', topic: 'countRule' },
    { sel: '#checksBox', topic: 'checks' },
    { sel: '[data-tab="review"]', topic: 'review', tab: 'review' },
    { sel: '[data-tab="sources"]', topic: 'sources', tab: 'sources' },
    { sel: '[data-tab="weekly"]', topic: 'weekly', tab: 'weekly' },
    { sel: '[data-tab="gestor"]', topic: 'gestor', tab: 'gestor' },
    { sel: '[data-tab="pending"]', topic: 'pending', tab: 'pending' },
    { sel: '[data-tab="monthly"]', topic: 'monthly', tab: 'monthly' },
    { sel: '[data-tab="period"]', topic: 'period', tab: 'period' },
    { sel: '[data-tab="team"]', topic: 'team', tab: 'team' },
    { sel: '[data-tab="data"]', topic: 'data', tab: 'data' },
    { sel: '#btnWord', topic: 'word' },
  ];

  const $ = (s) => document.querySelector(s);

  function el(tag, attrs, ...children) {
    const n = document.createElement(tag);
    for (const [k, v] of Object.entries(attrs || {})) {
      if (v === null || v === undefined || v === false) continue;
      if (k === 'class') n.className = v;
      else if (k === 'text') n.textContent = v;
      else if (k.startsWith('on')) n.addEventListener(k.slice(2), v);
      else n.setAttribute(k, v === true ? '' : v);
    }
    for (const c of children.flat()) if (c !== null && c !== undefined && c !== false) n.appendChild(typeof c === 'string' ? document.createTextNode(c) : c);
    return n;
  }

  function topicBody(t) {
    return [
      el('p', { class: 'help-intro', text: t.intro }),
      t.steps && t.steps.length ? el('div', null, el('h4', { text: 'How to use it' }), el('ol', null, t.steps.map((x) => el('li', { text: x })))) : null,
      t.points && t.points.length ? el('div', null, el('h4', { text: 'Good to know' }), el('ul', null, t.points.map((x) => el('li', { text: x })))) : null,
    ].filter(Boolean);
  }

  /* ---------- help window ---------- */

  function open(key) {
    const t = TOPICS[key] || TOPICS.overview;
    const dlg = $('#helpDialog');
    if (!dlg) return;
    $('#hdTitle').textContent = t.title;
    $('#hdBody').replaceChildren(...topicBody(t));
    const sel = $('#hdTopics');
    sel.value = TOPICS[key] ? key : 'overview';
    if (typeof dlg.showModal === 'function') { if (!dlg.open) dlg.showModal(); } else dlg.setAttribute('open', '');
    $('#hdBody').scrollTop = 0;
  }

  /** A "?" button that opens the help of a topic. */
  function button(key) {
    const t = TOPICS[key];
    if (!t) return null;
    return el('button', { type: 'button', class: 'help-btn', 'aria-label': `Help: ${t.title}`, title: `Help: ${t.title}`, text: '?',
      onclick: (e) => { e.preventDefault(); e.stopPropagation(); open(key); } });
  }

  /* ---------- guided tour ---------- */

  const tour = { steps: [], i: 0 };

  function narrow() {
    return window.innerWidth < 640;
  }

  function visible(node) {
    if (!node) return false;
    if (node.closest('[hidden]')) return false;
    const r = node.getBoundingClientRect();
    return r.width > 0 && r.height > 0;
  }

  function startTour() {
    const dlg = $('#helpDialog');
    if (dlg && dlg.open) dlg.close();
    tour.steps = TOUR.filter((s) => visible(document.querySelector(s.sel)));
    tour.loaded = !$('#results').hidden;
    tour.i = 0;
    $('#tourPop').hidden = false;
    $('#tourSpot').hidden = false;
    showStep();
    window.addEventListener('resize', showStep);
    window.addEventListener('scroll', place, true);
    document.addEventListener('keydown', tourKeys);
  }

  function endTour() {
    $('#tourPop').hidden = true;
    $('#tourSpot').hidden = true;
    window.removeEventListener('resize', showStep);
    window.removeEventListener('scroll', place, true);
    document.removeEventListener('keydown', tourKeys);
  }

  function tourKeys(e) {
    if (e.key === 'Escape') endTour();
    else if (e.key === 'ArrowRight') go(1);
    else if (e.key === 'ArrowLeft') go(-1);
  }

  function go(d) {
    const n = tour.steps.length + (tour.loaded ? 0 : 1);
    tour.i = Math.max(0, Math.min(n - 1, tour.i + d));
    showStep();
  }

  function showStep() {
    const total = tour.steps.length + (tour.loaded ? 0 : 1);
    const last = tour.i === total - 1;
    const step = tour.steps[tour.i];
    let title, body;
    if (step) {
      if (step.tab) { const tab = document.querySelector(`[data-tab="${step.tab}"]`); if (tab && !tab.classList.contains('active')) tab.click(); }
      const node = document.querySelector(step.sel);
      if (narrow()) {
        // phones: target near the top, the tour box docked at the bottom
        node.scrollIntoView({ block: 'start', behavior: 'auto' });
        window.scrollBy(0, -100);
      } else {
        node.scrollIntoView({ block: 'center', behavior: 'auto' });
      }
      const t = TOPICS[step.topic];
      title = t.title;
      body = [el('p', { text: t.intro }), (t.steps || t.points) ? el('ul', null, (t.steps || t.points).slice(0, 3).map((x) => el('li', { text: x }))) : null];
    } else {
      title = 'Load the files to continue';
      body = [el('p', { text: 'Upload the weekly files (step 1) and start the tour again: it will then show the settings, the Manual review tab, every report section and the Word export.' })];
    }
    $('#tourTitle').textContent = title;
    $('#tourBody').replaceChildren(...body.filter(Boolean));
    $('#tourStep').textContent = `Step ${tour.i + 1} of ${total}`;
    $('#tourPrev').disabled = tour.i === 0;
    $('#tourNext').textContent = last ? 'Finish' : 'Next ›';
    $('#tourMore').hidden = !step;
    place();
  }

  function place() {
    const pop = $('#tourPop');
    const spot = $('#tourSpot');
    if (pop.hidden) return;
    const step = tour.steps[tour.i];
    const node = step && document.querySelector(step.sel);
    if (!node || !visible(node)) {
      spot.style.cssText = 'top:50%;left:50%;width:0;height:0';
      pop.style.top = `${Math.max(16, window.innerHeight / 2 - pop.offsetHeight / 2)}px`;
      pop.style.left = `${Math.max(16, window.innerWidth / 2 - pop.offsetWidth / 2)}px`;
      return;
    }
    const r = node.getBoundingClientRect();
    const pad = 6;
    spot.style.cssText = `top:${r.top - pad}px;left:${r.left - pad}px;width:${r.width + 2 * pad}px;height:${r.height + 2 * pad}px`;
    const w = pop.offsetWidth, h = pop.offsetHeight;
    if (narrow()) {
      pop.style.top = `${Math.max(12, window.innerHeight - h - 12)}px`;
      pop.style.left = `${Math.max(12, (window.innerWidth - w) / 2)}px`;
      return;
    }
    let top = r.bottom + 14;
    if (top + h > window.innerHeight - 12) top = r.top - h - 14;
    if (top < 12) top = Math.max(12, Math.min(window.innerHeight - h - 12, r.top + 12));
    const left = Math.max(12, Math.min(window.innerWidth - w - 12, r.left));
    pop.style.top = `${top}px`;
    pop.style.left = `${left}px`;
  }

  /* ---------- init ---------- */

  function init() {
    document.querySelectorAll('[data-help]').forEach((node) => {
      if (node.querySelector(':scope > .help-btn')) return;
      const b = button(node.dataset.help);
      if (b) node.appendChild(b);
    });
    const sel = $('#hdTopics');
    if (sel) {
      sel.replaceChildren(...Object.entries(TOPICS).map(([k, t]) => el('option', { value: k, text: t.title })));
      sel.addEventListener('change', () => open(sel.value));
    }
    const dlg = $('#helpDialog');
    if (dlg) {
      $('#hdClose').addEventListener('click', () => dlg.close());
      $('#hdDone').addEventListener('click', () => dlg.close());
      dlg.addEventListener('click', (e) => { if (e.target === dlg) dlg.close(); });
      $('#hdTour').addEventListener('click', startTour);
    }
    const hb = $('#helpBtn');
    if (hb) hb.addEventListener('click', () => open('overview'));
    $('#tourPrev').addEventListener('click', () => go(-1));
    $('#tourNext').addEventListener('click', () => {
      const total = tour.steps.length + (tour.loaded ? 0 : 1);
      if (tour.i >= total - 1) endTour(); else go(1);
    });
    $('#tourClose').addEventListener('click', endTour);
    $('#tourMore').addEventListener('click', () => { const s = tour.steps[tour.i]; endTour(); if (s) open(s.topic); });
  }

  window.SLMHelp = { TOPICS, open, button, startTour, init };
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
  else init();
})();
