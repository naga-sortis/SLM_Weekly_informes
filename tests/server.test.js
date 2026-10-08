/* Server + database tests (synthetic data only). Run with: npm test */
'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const os = require('node:os');
const fs = require('node:fs');
const path = require('node:path');
const { createServer } = require('../server/server.js');
const { SlmDatabase } = require('../server/db.js');

const HEADER = ['Ticket ID', 'Third party reference', 'Creation date', 'Ticket type', 'Status', 'Current action', 'Restoration group ID', 'Restoration user name'];
const serial = (y, m, d, h = 10) => Date.UTC(y, m - 1, d, h) / 86400000 + 25569;
const row = (id, o = {}) => [id, o.ref ?? null, o.created ?? serial(2026, 9, 21), o.type ?? 'Failure', o.status ?? 'Current', o.action ?? '', o.rg ?? null, o.ru ?? null];
const workbook = (rows, extra = []) => ({ fileName: 'week.xlsx', user: 'Tester', sheets: [{ name: 'asignados a', rows: [[], ['Casos de Oceane'], HEADER, ...rows] }, ...extra] });

async function withServer(fn, opts = {}) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'slm-test-'));
  const server = createServer({ dbFile: path.join(dir, 'db.sqlite'), quiet: true, ...opts });
  await new Promise((r) => server.listen(0, '127.0.0.1', r));
  const base = `http://127.0.0.1:${server.address().port}`;
  try { await fn(base, server); } finally { await new Promise((r) => server.close(r)); server.db.close(); fs.rmSync(dir, { recursive: true, force: true }); }
}
const post = (base, p, body) => fetch(base + p, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) }).then(async (r) => ({ status: r.status, body: await r.json() }));

test('Import: dry run writes nothing; real import adds; next week updates, adds and keeps', async () => {
  await withServer(async (base) => {
    const w1 = workbook([row('A1'), row('A2', { type: 'Work order' }), row('A2', { type: 'Work order', rg: 'X' })]);
    const dry = await post(base, '/api/import?dryRun=1', w1);
    assert.equal(dry.status, 200);
    assert.deepEqual([dry.body.totals.added, dry.body.dryRun], [2, true]);
    assert.equal((await (await fetch(base + '/api/health')).json()).stats.tickets, 0);

    const r1 = await post(base, '/api/import', w1);
    assert.deepEqual(r1.body.totals, { added: 2, updated: 0, unchanged: 0, notInFile: 0 });

    // next week: A1 resolved, A2 identical, A3 new, (A1/A2 still there)
    const w2 = workbook([row('A1', { status: 'Resolved', rg: 'XSP00025', ru: 'ARCE LETICIA' }), row('A2', { type: 'Work order' }), row('A2', { type: 'Work order', rg: 'X' }), row('A3')]);
    const r2 = await post(base, '/api/import', w2);
    assert.deepEqual(r2.body.totals, { added: 1, updated: 1, unchanged: 1, notInFile: 0 });
    // week 3 without A2 → kept
    const r3 = await post(base, '/api/import', workbook([row('A3')]));
    assert.equal(r3.body.totals.notInFile, 2);

    const sheets = (await (await fetch(base + '/api/sheets')).json()).sheets;
    assert.equal(sheets.length, 1);
    const rows = sheets[0].rows;
    assert.deepEqual(rows[0].slice(0, 3), ['Ticket ID', 'Third party reference', 'Creation date']);
    assert.equal(rows.length - 1, 4); // A1, A2 (2 rows), A3
    assert.deepEqual(sheets[0].notInLatest.sort(), ['A1', 'A2']);
    const a1 = rows.find((r) => r[0] === 'A1');
    assert.equal(a1[rows[0].indexOf('Status')], 'Resolved');

    const hist = await (await fetch(base + '/api/tickets/A1/history')).json();
    assert.ok(hist.changes.some((c) => c.field === 'Status' && c.old_value === 'Current' && c.new_value === 'Resolved'));
    const imports = (await (await fetch(base + '/api/imports')).json()).imports;
    assert.equal(imports.length, 3);
    assert.equal(imports[0].user_name, 'Tester');
  });
});

test('Rebuilt sheets give the same tickets as the original file', async () => {
  const C = require('../assets/js/core.js');
  const wb = workbook([row('B1', { created: serial(2026, 9, 1) }), row('B2', { type: 'Work order', action: 'JR - Trabajando' }), row('B3', { ref: 'STA-1' })]);
  await withServer(async (base) => {
    await post(base, '/api/import', wb);
    const sheets = (await (await fetch(base + '/api/sheets')).json()).sheets;
    const a = C.extractTickets(wb.sheets).tickets.map((t) => [t.id, t.created, t.category, t.vendor, t.status]);
    const b = C.extractTickets(sheets).tickets.map((t) => [t.id, t.created, t.category, t.vendor, t.status]);
    assert.deepEqual(b, a);
  });
});

test('Manual classifications: set with name, change, remove, logged', async () => {
  await withServer(async (base) => {
    const put = (id, body) => fetch(`${base}/api/classifications/${id}`, { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) }).then(async (r) => ({ status: r.status, body: await r.json() }));
    assert.equal((await put('T1', { category: 'OT' })).status, 400); // name required
    assert.equal((await put('T1', { category: 'XX', user: 'Ana' })).status, 400);
    const ok = await put('T1', { category: 'OT', user: 'Ana' });
    assert.deepEqual([ok.status, ok.body.category, ok.body.user_name], [200, 'OT', 'Ana']);
    await put('T1', { category: 'INC', user: 'Luis' });
    let list = (await (await fetch(base + '/api/classifications')).json()).classifications;
    assert.deepEqual(list.map((c) => [c.ticket_id, c.category, c.user_name]), [['T1', 'INC', 'Luis']]);
    assert.equal((await fetch(`${base}/api/classifications/T1?user=Ana`, { method: 'DELETE' })).status, 200);
    list = (await (await fetch(base + '/api/classifications')).json()).classifications;
    assert.equal(list.length, 0);
    const hist = await (await fetch(base + '/api/tickets/T1/history')).json();
    assert.deepEqual(hist.classifications.map((c) => [c.action, c.category, c.user_name]), [['set', 'OT', 'Ana'], ['set', 'INC', 'Luis'], ['remove', null, 'Ana']]);
  });
});

test('Static files: page served, private files never served, security headers present', async () => {
  await withServer(async (base) => {
    const page = await fetch(base + '/');
    assert.equal(page.status, 200);
    assert.match(page.headers.get('content-security-policy'), /default-src 'self'/);
    assert.equal(page.headers.get('cache-control'), 'no-cache');
    for (const p of ['/server/db.js', '/package.json', '/tests/server.test.js', '/data/slm-informes.sqlite', '/..%2fpackage.json', '/.git/config', '/README.md']) {
      assert.equal((await fetch(base + p)).status, 404, p);
    }
    assert.equal((await fetch(base + '/vendor/xlsx.full.min.js')).status, 200);
    assert.equal((await fetch(base + '/assets/js/core.js')).status, 200);
  });
});

test('Optional login (SLM_BASIC_AUTH)', async () => {
  await withServer(async (base) => {
    assert.equal((await fetch(base + '/api/health')).status, 401);
    const auth = { Authorization: 'Basic ' + Buffer.from('slm:secret').toString('base64') };
    assert.equal((await fetch(base + '/api/health', { headers: auth })).status, 200);
  }, { basicAuth: 'slm:secret' });
});

test('Backup is a valid database', async () => {
  await withServer(async (base) => {
    await post(base, '/api/import', workbook([row('Z1')]));
    const buf = Buffer.from(await (await fetch(base + '/api/backup')).arrayBuffer());
    const f = path.join(os.tmpdir(), `slm-bk-${Date.now()}.sqlite`);
    fs.writeFileSync(f, buf);
    const db = new SlmDatabase(f);
    assert.equal(db.stats().tickets, 1);
    db.close(); fs.rmSync(f, { force: true });
  });
});

test('Only real changes are reported (HTML re-encoding is not a change); dates logged as dates', async () => {
  const { SlmDatabase } = require('../server/db.js');
  const db = new SlmDatabase(':memory:');
  const H = ['Ticket ID', 'Creation date', 'Ticket type', 'Status', 'Description', 'Restoration date'];
  const wb = (desc, status, restored) => ({ fileName: 'f.xlsx', sheets: [{ name: 'asignados a', rows: [H, ['X1', 46280.5, 'Failure', status, desc, restored]] }] });
  db.importWorkbook(wb('<span style="font-family:&quot;Aptos&quot;">Hola&nbsp;mundo</span>', 'Current', null));
  const same = db.importWorkbook(wb('<span style="font-family:"Aptos"">Hola mundo</span>', 'Current', null), { dryRun: true });
  assert.deepEqual(same.totals, { added: 0, updated: 0, unchanged: 1, notInFile: 0 });
  const real = db.importWorkbook(wb('<span>Hola mundo</span>', 'Resolved', 46290.5));
  assert.equal(real.totals.updated, 1);
  const ch = db.ticketHistory('X1').changes;
  assert.deepEqual(ch.map((c) => [c.field, c.old_value, c.new_value]), [['Status', 'Current', 'Resolved'], ['Restoration date', '', '2026/09/25 12:00:00']]);
  db.close();
});
