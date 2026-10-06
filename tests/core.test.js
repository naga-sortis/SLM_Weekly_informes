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
