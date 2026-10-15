#!/usr/bin/env node
/*
 * Consistent copy of the team database, safe while the server is running.
 *
 *   node server/backup.js <backup-folder> [keep]     daily backup, keeps the newest <keep> files (default 30)
 *   node server/backup.js --copy <target-file>        one copy to an exact file (used when moving the database)
 *   node server/backup.js --check                     integrity check + counts, nothing is written
 *
 * The database is SLM_DB (default ./data/slm-informes.sqlite).
 */
'use strict';

const fs = require('node:fs');
const path = require('node:path');

const v = process.versions.node.split('.').map(Number);
if (v[0] < 22 || (v[0] === 22 && v[1] < 13)) {
  console.error(`Node.js 22.13 or newer is needed (this is ${process.versions.node}).`);
  process.exit(1);
}
const emitWarning = process.emitWarning;
process.emitWarning = function (w, ...rest) {
  if (/sqlite/i.test(String(w && w.message || w))) return;
  return emitWarning.call(process, w, ...rest);
};
const { DatabaseSync } = require('node:sqlite');

const source = path.resolve(process.env.SLM_DB || path.join(__dirname, '..', 'data', 'slm-informes.sqlite'));

function open() {
  if (!fs.existsSync(source)) { console.error(`No database at ${source}`); process.exit(2); }
  const db = new DatabaseSync(source, { readOnly: true });
  db.exec('PRAGMA busy_timeout = 15000');
  return db;
}

function counts(db) {
  const n = (t) => { try { return db.prepare(`SELECT COUNT(*) AS n FROM ${t}`).get().n; } catch (e) { return 0; } };
  return { tickets: db.prepare('SELECT COUNT(DISTINCT ticket_id) AS n FROM ticket_rows').get().n, imports: n('imports'), classifications: n('classifications') };
}

function copyTo(db, target) {
  fs.mkdirSync(path.dirname(target), { recursive: true });
  const tmp = `${target}.partial`;
  if (fs.existsSync(tmp)) fs.unlinkSync(tmp);
  db.prepare('VACUUM INTO ?').run(tmp);
  // Verify the copy before it replaces anything.
  const copy = new DatabaseSync(tmp, { readOnly: true });
  const ok = copy.prepare('PRAGMA integrity_check').get();
  const c = counts(copy);
  copy.close();
  if (Object.values(ok)[0] !== 'ok') { fs.unlinkSync(tmp); throw new Error('the copy failed its integrity check'); }
  fs.renameSync(tmp, target);
  return c;
}

const args = process.argv.slice(2);
try {
  if (args[0] === '--check') {
    const db = open();
    const ok = Object.values(db.prepare('PRAGMA integrity_check').get())[0];
    const c = counts(db);
    db.close();
    console.log(`${source}: integrity ${ok}; ${c.tickets} tickets, ${c.imports} imports, ${c.classifications} classifications`);
    process.exit(ok === 'ok' ? 0 : 3);
  } else if (args[0] === '--copy' && args[1]) {
    const db = open();
    const c = copyTo(db, path.resolve(args[1]));
    db.close();
    console.log(`Copied ${source} -> ${path.resolve(args[1])} (${c.tickets} tickets, ${c.imports} imports, ${c.classifications} classifications)`);
  } else if (args[0] && !args[0].startsWith('-')) {
    const dir = path.resolve(args[0]);
    const keep = Math.max(1, Number(args[1]) || 30);
    const stamp = new Date().toISOString().slice(0, 19).replace(/[:T]/g, '-');
    const db = open();
    const target = path.join(dir, `slm-informes-${stamp}.sqlite`);
    const c = copyTo(db, target);
    db.close();
    const old = fs.readdirSync(dir).filter((f) => /^slm-informes-[\d-]+\.sqlite$/.test(f)).sort().reverse().slice(keep);
    for (const f of old) fs.unlinkSync(path.join(dir, f));
    console.log(`Backup ${target} (${c.tickets} tickets, ${c.imports} imports, ${c.classifications} classifications); removed ${old.length} old backup(s)`);
  } else {
    console.error('Usage: node server/backup.js <backup-folder> [keep] | --copy <file> | --check');
    process.exit(64);
  }
} catch (err) {
  console.error(`Backup failed: ${err.message}`);
  process.exit(1);
}
