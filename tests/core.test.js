/* Run with: node --test tests/  (Node 18+; no dependencies, synthetic data only) */
'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const C = require('../assets/js/core.js');

const HEADER = ['Ticket ID', 'Third party reference', 'Creation date', 'Acknowledgement date', 'Creation week', 'Creation month',
  'Creation year', 'Processing priority', 'Ticket type', 'Status', 'Initiator - Group ID', 'Initiator - Group abbreviation name',
  'Initiator - User name', 'Current action', 'Description', 'Restoration date'];

const serial = (y, m, d, h = 10, mi = 0) => Date.UTC(y, m - 1, d, h, mi) / 86400000 + 25569;

/** Build a sheet with a few title rows above the header, like the Oceane export. */
function sheet(rows) {
  return { name: 'asignados a', rows: [[], [null, 'Casos de Oceane'], [], HEADER, ...rows.map((r) => HEADER.map((h) => r[h] ?? null))] };
}

function ticket(id, created, o = {}) {
  const ms = Date.UTC(...created);
  const iso = C.isoWeekInfo(ms);
  return {
    'Ticket ID': id,
    'Creation date': serial(created[0], created[1] + 1, created[2], created[3] ?? 10),
    'Creation week': iso.week,
    'Creation month': created[1] + 1,
    'Creation year': created[0],
    'Processing priority': o.prio || 'P3',
    'Ticket type': o.type || 'Failure',
    Status: o.status || 'Current',
    'Initiator - Group ID': o.group || '515104',
    'Initiator - User name': o.user || 'USER',
    'Third party reference': o.ref || null,
    'Current action': o.action ?? 'NM RM - BACKUP - Leticia',
    Description: o.desc || 'desc',
    'Restoration date': o.restored ? serial(...o.restored) : null,
  };
}

test('ISO weeks at year boundaries', () => {
  assert.deepEqual([C.isoWeekInfo(Date.UTC(2025, 11, 29)).year, C.isoWeekInfo(Date.UTC(2025, 11, 29)).week], [2026, 1]);
  assert.deepEqual([C.isoWeekInfo(Date.UTC(2024, 11, 30)).year, C.isoWeekInfo(Date.UTC(2024, 11, 30)).week], [2025, 1]);
  assert.deepEqual([C.isoWeekInfo(Date.UTC(2021, 0, 3)).year, C.isoWeekInfo(Date.UTC(2021, 0, 3)).week], [2020, 53]);
  assert.equal(C.isoWeekInfo(Date.UTC(2026, 8, 25)).week, 39);
  const w = C.weekFromKey(202639);
  assert.equal(C.fmtDate(w.start), '21/09/2026');
  assert.equal(C.fmtDate(w.end - 86400000), '27/09/2026');
  assert.equal(C.weekRange(202551, 202602).length, 4); // 51, 52, 1, 2
});

test('Month label of a week', () => {
  const w39 = C.weekFromKey(202639);
  assert.equal(C.periodMonthsLabel(w39.start, w39.end), 'Septiembre 2026');
  const w40 = C.weekFromKey(202640); // 28/09 – 04/10
  assert.equal(C.periodMonthsLabel(w40.start, w40.end), 'Septiembre / Octubre 2026');
  const w1 = C.weekFromKey(202601); // 29/12/2025 – 04/01/2026
  assert.equal(C.periodMonthsLabel(w1.start, w1.end), 'Diciembre 2025 / Enero 2026');
});

test('Date parsing keeps the wall-clock time from Excel', () => {
  assert.equal(C.fmtDateTime(C.parseDate(serial(2026, 9, 25, 14, 6))), '2026/09/25 14:06:00');
  assert.equal(C.fmtDateTime(C.parseDate('2024/05/03 11:14:00')), '2024/05/03 11:14:00');
  assert.equal(C.fmtDateTime(C.parseDate('03/05/2024 11:14')), '2024/05/03 11:14:00'); // DD/MM/YYYY
  assert.equal(C.parseDate('31/02/2024'), null);
  assert.equal(C.parseDate(''), null);
  assert.equal(C.parseDate('abc'), null);
});

test('Third party reference → vendor', () => {
  assert.equal(C.vendorOf('STA-00265291'), 'Ericsson');
  assert.equal(C.vendorOf('H-37805479'), 'Huawei');
  assert.equal(C.vendorOf('1-03827757'), 'Nokia');
  assert.equal(C.vendorOf('1'), 'Otro');
  assert.equal(C.vendorOf('SP1175'), 'Otro');
  assert.equal(C.vendorOf(''), '');
});

test('Current action → gestor / problema / técnico', () => {
  assert.deepEqual(pick(C.parseAction('NM RM - NM3  BACKUP - Leticia')), ['NM RM', 'NM3  BACKUP', 'Leticia']);
  assert.deepEqual(pick(C.parseAction('NFM-T - Movimiento equipos - Julio')), ['NFM-T', 'Movimiento equipos', 'Julio']);
  assert.deepEqual(pick(C.parseAction('EC-1 - Solicitud FLM - Julio')), ['EC-1', 'Solicitud FLM', 'Julio']);
  assert.deepEqual(pick(C.parseAction('siENM - Problema conexión- Julio')), ['SIENM', 'Problema conexión', 'Julio']);
  assert.equal(C.parseAction('JR - Trabajando').gestor, '');
  assert.equal(C.parseAction('GV - ACCESO MENESES').gestor, '');
  assert.equal(C.parseAction('DEVUELTO').kind, 'devuelto');
  function pick(p) { return [p.gestor, p.problema, p.tecnico]; }
});

test('Gestor spelling variants are merged and counted together', () => {
  const ex = C.extractTickets([sheet([
    ticket('A1', [2026, 8, 21], { action: 'NFM-T - x - Julio' }),
    ticket('A2', [2026, 8, 22], { action: 'NFMT - y - Julio' }),
    ticket('A3', [2026, 8, 22], { action: 'NFM -T - z - Leticia' }),
    ticket('A4', [2026, 8, 23], { action: 'ENM 3 - a - Leticia' }),
    ticket('A5', [2026, 8, 23], { action: 'ENM3 - b - Leticia' }),
    ticket('A6', [2026, 8, 23], { action: 'JR - Trabajando' }),
  ])]);
  const r = C.computeReport(ex.tickets, { weeks: C.weekRange(202639, 202639), slmGroups: ['XSP00025'] });
  const byName = Object.fromEntries(r.gestores.map((g) => [g.gestor, g.total]));
  assert.equal(byName['NFM-T'], 3);
  assert.equal(Object.keys(byName).filter((k) => C.gestorKey(k) === 'ENM3').length, 1);
  assert.equal(byName['Sin gestor identificado'], 1);
  assert.equal(r.gestores.reduce((a, g) => a + g.total, 0), r.inc[0].nuevos);
  assert.ok(r.checks.every((c) => c.ok), JSON.stringify(r.checks.filter((c) => !c.ok)));
});

test('Weekly figures, backlog carry-over and cut-off', () => {
  const rows = [
    // week 38 (14–20/09/2026)
    ticket('W38-1', [2026, 8, 14], { restored: [2026, 9, 15, 12] }), // resolved in week 38
    ticket('W38-2', [2026, 8, 15], { status: 'Resolved', restored: [2026, 9, 23, 12] }), // resolved in week 39 → backlog
    ticket('W38-3', [2026, 8, 16], { ref: 'STA-1' }), // escalated, still open
    ticket('W38-4', [2026, 8, 16], { action: 'DEVUELTO', status: 'Closed' }), // returned
    ticket('W38-5', [2026, 8, 17], { type: 'Work order', group: 'XSP00025' }), // OT opened by SLM, open
    // week 39 (21–27/09/2026)
    ticket('W39-1', [2026, 8, 21], { group: 'XSP00025', user: 'ARCE LETICIA', status: 'Closed for dashboard', restored: [2026, 9, 22, 9] }),
    ticket('W39-2', [2026, 8, 25], { ref: 'H-123', status: 'Resolved', restored: [2026, 10, 2, 9] }), // resolved after week 39
    // week 40 – after the report week, must never be counted
    ticket('W40-1', [2026, 8, 29]),
  ];
  // first ticket: status must be closed to count as resolved
  rows[0].Status = 'Closed';
  const ex = C.extractTickets([sheet(rows)]);
  const r = C.computeReport(ex.tickets, { weeks: C.weekRange(202638, 202639), slmGroups: ['XSP00025', 'XSP00027'] });
  assert.ok(r.checks.every((c) => c.ok), JSON.stringify(r.checks.filter((c) => !c.ok)));
  const [i38, i39] = r.inc;
  assert.deepEqual([i38.nuevos, i38.resueltas, i38.sinResolverNoEsc, i38.sinResolverEsc, i38.devueltas], [4, 1, 1, 1, 1]);
  // SLM-opened ticket is counted like any other, and also shown in "Abiertas por el SLM"
  assert.deepEqual([i39.nuevos, i39.resueltas, i39.sinResolverNoEsc, i39.sinResolverEsc, i39.abiertasSLM], [2, 1, 0, 1, 1]);
  assert.deepEqual(i39.backlogNoEsc, { inicio: 1, resueltos: 1, sinResolver: 0 });
  assert.deepEqual(i39.backlogEsc, { inicio: 1, resueltos: 0, sinResolver: 1 });
  assert.equal(r.ot[0].abiertasSLM, 1);
  // pending at end of week 39: W38-3, W39-2 (INC) + W38-5 (OT); never W40-1
  assert.deepEqual(r.pending.map((t) => t.id).sort(), ['W38-3', 'W38-5', 'W39-2']);
  assert.equal(r.periodLabel, 'Septiembre 2026');
  assert.equal(r.reportWeek.week, 39);
  // monthly: September opened up to the cut-off (W40-1 is 29/09 → excluded)
  const sepOpened = r.monthly.opened.reduce((a, s) => a + s.data[s.data.length - 1], 0);
  assert.equal(sepOpened, 7);
});

test('Duplicates and invalid rows are reported', () => {
  const ex = C.extractTickets([sheet([
    ticket('D1', [2026, 8, 21]),
    ticket('D1', [2026, 8, 22]),
    { ...ticket('D2', [2026, 8, 22]), 'Creation date': 'not a date' },
  ])]);
  assert.equal(ex.tickets.length, 1);
  assert.equal(ex.duplicates, 1);
  assert.equal(ex.invalidDates, 1);
});

test('Wrong "Creation week" in the file is flagged and recalculated', () => {
  const t = ticket('X1', [2026, 8, 21]);
  t['Creation week'] = 12;
  const ex = C.extractTickets([sheet([t])]);
  assert.equal(ex.tickets[0].week, 39);
  assert.ok(ex.warnings.some((w) => w.includes('Creation week')));
});

test('A week spanning two months is reconciled with the monthly charts', () => {
  // Week 40 of 2026: Mon 28/09 – Sun 04/10
  const ex = C.extractTickets([sheet([
    ticket('S1', [2026, 8, 28]), ticket('S2', [2026, 8, 30], { type: 'Work order' }),
    ticket('O1', [2026, 9, 1]), ticket('O2', [2026, 9, 2]), ticket('O3', [2026, 9, 4], { type: 'Work order' }),
  ])]);
  const w = [C.weekFromKey(202639), C.weekFromKey(202640)];
  const r = C.computeReport(ex.tickets, { weeks: w, slmGroups: [], monthsBack: 3 });
  assert.ok(r.checks.every((c) => c.ok), JSON.stringify(r.checks.filter((c) => !c.ok)));
  assert.equal(r.periodLabel, 'Septiembre / Octubre 2026');
  assert.deepEqual(r.weekSplit.map((p) => [p.monthKey, p.inc, p.ot]), [[202609, 1, 1], [202610, 2, 1]]);
  assert.equal(r.inc[1].nuevos + r.ot[1].nuevos, 5);
  assert.match(r.monthLabels[r.monthLabels.length - 1], /^2026-10 \(hasta 04\/10\)$/);
  assert.match(C.weekSplitText(r), /5 casos nuevos.*2 en septiembre.*3 en octubre/);
});

test('Month / year view: all incidencias + OTs of the period', () => {
  const ex = C.extractTickets([sheet([
    ticket('M1', [2026, 7, 3], { status: 'Closed', restored: [2026, 8, 4, 9] }),
    ticket('M2', [2026, 7, 10], { type: 'Work order' }),
    ticket('M3', [2026, 7, 20], { ref: 'H-1', action: 'NFMT - x - Julio' }),
    ticket('M4', [2026, 7, 31], { action: 'DEVUELTO', status: 'Closed' }),
    ticket('J1', [2026, 6, 30], { status: 'Resolved', restored: [2026, 8, 2, 9] }), // July ticket resolved in August
    ticket('S1', [2026, 8, 1]),
    ticket('Y1', [2025, 7, 15], { type: 'Work order' }),
  ])]);
  const aug = C.computePeriod(ex.tickets, { year: 2026, month: 8, slmGroups: [] });
  assert.ok(aug.checks.every((c) => c.ok), JSON.stringify(aug.checks.filter((c) => !c.ok)));
  assert.equal(aug.label, 'Agosto 2026');
  assert.deepEqual([aug.summary.INC.abiertos, aug.summary.OT.abiertos], [3, 1]);
  assert.deepEqual([aug.summary.INC.resueltos, aug.summary.INC.pendientes, aug.summary.INC.devueltos], [1, 1, 1]);
  assert.equal(aug.summary.INC.escalados, 1);
  assert.equal(aug.summary.INC.resueltosEnPeriodo, 2); // M1 + J1 (opened in July)
  assert.equal(aug.buckets[0].label, 'sem. 31 (01–02/08)'); // week clipped to the month
  assert.ok(aug.gestores.some((g) => g.gestor === 'NFMT' && g.inc === 1));

  const y2026 = C.computePeriod(ex.tickets, { year: 2026, month: null, slmGroups: [] });
  assert.ok(y2026.checks.every((c) => c.ok));
  assert.equal(y2026.label, 'Año 2026');
  assert.equal(y2026.summary.INC.abiertos + y2026.summary.OT.abiertos, 6);
  assert.equal(y2026.buckets.length, 9); // Jan..Sep (no empty future months after the last data)
  const y2025 = C.computePeriod(ex.tickets, { year: 2025, month: null, slmGroups: [] });
  assert.deepEqual([y2025.summary.INC.abiertos, y2025.summary.OT.abiertos], [0, 1]);
  assert.deepEqual(C.periodsAvailable(ex.tickets).map((p) => p.year), [2026, 2025]);
});

test('Word builder accepts the month / year section', () => {
  let D;
  try { D = require('docx'); } catch (e) { return; } // docx is only vendored for the browser
  const ex = C.extractTickets([sheet([ticket('A', [2026, 8, 21]), ticket('B', [2026, 8, 22], { type: 'Work order' })])]);
  const r = C.computeReport(ex.tickets, { weeks: [C.weekFromKey(202639)], slmGroups: [] });
  const p = C.computePeriod(ex.tickets, { year: 2026, month: 9, slmGroups: [] });
  const R = require('../assets/js/docx-report.js');
  const doc = R.buildDocument(D, C, r, { title: 'T', period: { data: p, note: 'n' } }, {});
  assert.ok(doc);
});

test('“Sin gestor identificado” lists every ticket with the reason', () => {
  const ex = C.extractTickets([sheet([
    ticket('G1', [2026, 8, 21], { action: 'NM RM - BACKUP - Leticia' }),
    ticket('N1', [2026, 8, 21], { action: 'JR - Trabajando' }),
    ticket('N2', [2026, 8, 22], { action: '' }),
    ticket('N3', [2026, 8, 22], { action: 'DEVUELTO', status: 'Closed' }),
    ticket('N4', [2026, 8, 23], { action: 'Cerrado', status: 'Closed' }),
    ticket('O1', [2026, 8, 23], { type: 'Work order', action: 'IG - Trabajando' }),
  ])]);
  const r = C.computeReport(ex.tickets, { weeks: [C.weekFromKey(202639)], slmGroups: [] });
  assert.ok(r.checks.every((c) => c.ok), JSON.stringify(r.checks.filter((c) => !c.ok)));
  assert.deepEqual(r.noGestorTickets.map((t) => t.id).sort(), ['N1', 'N2', 'N3', 'N4']); // weekly 2.2 = incidencias only
  const reasons = Object.fromEntries(r.noGestorTickets.map((t) => [t.id, C.noGestorReason(t).code]));
  assert.deepEqual(reasons, { N1: 'trabajando', N2: 'empty', N3: 'devuelto', N4: 'cerrado' });
  const p = C.computePeriod(ex.tickets, { year: 2026, month: 9, slmGroups: [] });
  assert.ok(p.checks.every((c) => c.ok));
  assert.equal(p.noGestorTickets.length, 5); // incidencias + OTs
  assert.ok(p.noGestorTickets.every((t) => C.noGestorReason(t).label.length > 10));
});

test('Numbers under each chart: series, subtotals and totals', () => {
  const CH = require('../assets/js/charts.js').SLMCharts;
  const t = CH.chartTable(['Ago', 'Sep'], [
    { name: 'OTs', data: [5, 7] }, { name: 'Inc. P2', data: [1, 0] }, { name: 'Inc. P3', data: [3, 4] },
  ]);
  const byName = Object.fromEntries(t.rows.map((r) => [r.name, r]));
  assert.deepEqual(byName['Total incidencias'].data, [4, 4]);
  assert.deepEqual(byName['Total OTs'].data, [5, 7]);
  assert.deepEqual(byName.Total.data, [9, 11]);
  assert.equal(byName.Total.total, 20);
  const w = CH.chartTable(['s1'], [{ name: 'Incidencias', data: [3] }, { name: 'OTs', data: [2] }]);
  assert.deepEqual(w.rows.map((r) => r.name), ['Incidencias', 'OTs', 'Total']);
});

const HEADER_FULL = [...HEADER, 'Restoration group ID', 'Restoration group abbreviation name', 'Restoration user name',
  'Closure group ID', 'Closure group abbreviation name', 'Closure user name', 'Final nature'];
function sheetFull(name, rows) {
  return { name, rows: [[], HEADER_FULL, ...rows.map((r) => HEADER_FULL.map((h) => r[h] ?? null))] };
}
function rt(id, created, group, user, o = {}) {
  return { ...ticket(id, created, { status: 'Closed', restored: o.restored || [created[0], created[1] + 1, created[2], 15], ...o }),
    'Restoration group ID': group, 'Restoration group abbreviation name': group === 'XSP00025' ? 'SPOC_BO_EMS' : 'X',
    'Restoration user name': user, 'Closure group ID': o.closure || null, 'Closure user name': o.closureUser || null };
}

test('All columns are kept and duplicate rows are merged, not lost', () => {
  const a = rt('D1', [2026, 8, 1], 'XSP00027', 'PAREDES', { closure: '00XMOA' });
  const b = rt('D1', [2026, 8, 1], 'XSP00027', 'PAREDES', { closure: '515702' });
  const ex = C.extractTickets([sheetFull('asignados a', [a]), sheetFull('Ticket Work Order 2025', [b])]);
  const t = ex.tickets[0];
  assert.deepEqual(t.sheets, ['asignados a', 'Ticket Work Order 2025']);
  assert.deepEqual(t.closureGroups, ['00XMOA', '515702']);
  assert.deepEqual(t.rawAlt['Closure group ID'], ['515702']);
  assert.equal(t.restorationUser, 'PAREDES');
  assert.ok(t.raw.some(([k, v]) => k === 'Final nature'));
  assert.ok(t.raw.find(([k]) => k === 'Creation date')[1].startsWith('2026/09/01'));
});

test('Sortis team: restoration group / user from all sheets', () => {
  const rows = [
    rt('T1', [2026, 8, 1], 'XSP00025', 'ARCE LETICIA'),
    rt('T2', [2026, 8, 2], 'XSP00025', 'ARCE LETICIA', { type: 'Work order' }),
    rt('T3', [2026, 8, 3], 'XSP00027', 'PAREDES NAVARRO ADRIAN', { type: 'Work order', closureUser: 'PAREDES NAVARRO ADRIAN' }),
    rt('T4', [2026, 7, 30], 'XSP00027', null), // no user
    rt('X1', [2026, 8, 4], 'XSPINSYTE01', 'OTHER PERSON'),
    { ...ticket('U1', [2026, 8, 5]), 'Restoration group ID': 'XSP00027', 'Restoration user name': 'PAREDES NAVARRO ADRIAN' }, // no restoration date
  ];
  const ex = C.extractTickets([sheetFull('asignados a', rows)]);
  const all = C.computeTeam(ex.tickets, { groups: ['XSP00025', 'XSP00027'] });
  assert.ok(all.checks.every((c) => c.ok), JSON.stringify(all.checks.filter((c) => !c.ok)));
  assert.equal(all.total, 5); // U1 counted even without restoration date
  const by = Object.fromEntries(all.engineers.map((e) => [e.user, e]));
  assert.equal(by['ARCE LETICIA'].byGroup.XSP00025.inc, 1);
  assert.equal(by['ARCE LETICIA'].byGroup.XSP00025.ot, 1);
  assert.equal(by['PAREDES NAVARRO ADRIAN'].total, 2);
  assert.equal(by['PAREDES NAVARRO ADRIAN'].closedByThem, 1);
  assert.ok(by['(sin usuario de restauración)']);
  assert.ok(all.restorationGroups.some((g) => g.group === 'XSPINSYTE01' && !g.sortis));
  const sep = C.computeTeam(ex.tickets, { groups: ['XSP00025', 'XSP00027'], year: 2026, month: 9 });
  assert.ok(sep.checks.every((c) => c.ok));
  assert.equal(sep.total, 3); // T1–T3 restored in September; T4 restored 31/08 → August; U1 has no date
  assert.equal(sep.undated, 1);
  const q25 = C.computeTeam(ex.tickets, { groups: ['XSP00025', 'XSP00027'], queue: 'XSP00025' });
  assert.equal(q25.total, 2);
  assert.ok(q25.checks.every((c) => c.ok));
});

test('Word builder: sections can be left out and are renumbered', () => {
  let D;
  try { D = require('docx'); } catch (e) { return; }
  const ex = C.extractTickets([sheetFull('asignados a', [rt('A', [2026, 8, 21], 'XSP00025', 'ARCE LETICIA')])]);
  const r = C.computeReport(ex.tickets, { weeks: [C.weekFromKey(202639)], slmGroups: [] });
  const T = C.computeTeam(ex.tickets, { groups: ['XSP00025', 'XSP00027'] });
  const R = require('../assets/js/docx-report.js');
  const doc = R.buildDocument(D, C, r, { title: 'T', include: { s21: false, monthly: false }, team: { data: T, note: 'n' } }, {});
  assert.ok(doc);
});

test('CSR third party references are Ericsson', () => {
  assert.equal(C.vendorOf('CSR-00621504'), 'Ericsson');
  assert.equal(C.vendorOf('CSR 00401055'), 'Ericsson');
  assert.equal(C.vendorOf('CSR STA-00725162'), 'Ericsson');
  assert.equal(C.vendorOf('csr-1'), 'Ericsson');
  assert.equal(C.vendorOf('CSRX1'), 'Otro'); // another code that only starts with the letters
});

test('Trend charts can show one calendar year only', () => {
  const ex = C.extractTickets([sheet([
    ticket('A', [2025, 10, 20]), ticket('B', [2025, 11, 3]), ticket('C', [2026, 0, 15]), ticket('D', [2026, 8, 22]),
  ])]);
  const w = [C.weekFromKey(202639)];
  const def = C.computeReport(ex.tickets, { weeks: w, slmGroups: [], monthsBack: 12 });
  assert.equal(def.trendYear, null);
  assert.equal(C.monthShort(def.months[0]), '2025-10'); // last 12 months reach back into 2025
  const y26 = C.computeReport(ex.tickets, { weeks: w, slmGroups: [], trendYear: 2026 });
  assert.ok(y26.checks.every((c) => c.ok), JSON.stringify(y26.checks.filter((c) => !c.ok)));
  assert.deepEqual([C.monthShort(y26.months[0]), C.monthShort(y26.months[y26.months.length - 1])], ['2026-01', '2026-09']);
  assert.equal(y26.monthly.opened.reduce((a, s) => a + s.data.reduce((x, y) => x + y, 0), 0), 2); // only C and D
  const y25 = C.computeReport(ex.tickets, { weeks: w, slmGroups: [], trendYear: 2025 });
  assert.ok(y25.checks.every((c) => c.ok));
  assert.equal(y25.months.length, 12);
  assert.ok(y25.months.every((m) => Math.floor(m / 100) === 2025));
  assert.match(C.weekSplitText(y25), /año 2025/);
  const future = C.computeReport(ex.tickets, { weeks: w, slmGroups: [], trendYear: 2027, monthsBack: 7 });
  assert.equal(future.trendYear, null); // a year after the report week is never used
});
