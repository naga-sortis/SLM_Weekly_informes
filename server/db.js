/*
 * SLM Weekly Informes — team database (SQLite, built into Node.js ≥ 22.13).
 *
 * Every weekly Excel is merged into the database, sheet by sheet:
 *   - a ticket not seen before is added;
 *   - a ticket already stored is updated when any of its values changed (the change is logged);
 *   - a ticket that is not in the new file is kept (history is never deleted).
 * Rows are stored exactly as read from the Excel (one JSON object per row, header → cell value),
 * so the page can rebuild the sheets and run the same report logic on the whole history.
 */
'use strict';

// node:sqlite prints an "ExperimentalWarning" on some Node versions; it is harmless, so hide only that one.
const emitWarning = process.emitWarning;
process.emitWarning = function (warning, ...rest) {
  const text = String(warning && warning.message || warning);
  const type = typeof rest[0] === 'string' ? rest[0] : (rest[0] && rest[0].type) || (warning && warning.name);
  if (type === 'ExperimentalWarning' && /sqlite/i.test(text)) return;
  return emitWarning.call(process, warning, ...rest);
};
const { DatabaseSync } = require('node:sqlite');
const fs = require('node:fs');
const path = require('node:path');
const C = require('../assets/js/core.js');

const SCHEMA = `
  PRAGMA journal_mode = WAL;
  PRAGMA foreign_keys = ON;
  CREATE TABLE IF NOT EXISTS imports (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    file_name TEXT NOT NULL,
    imported_at TEXT NOT NULL,
    user_name TEXT NOT NULL DEFAULT '',
    sheets TEXT NOT NULL,
    rows_read INTEGER NOT NULL,
    added INTEGER NOT NULL,
    updated INTEGER NOT NULL,
    unchanged INTEGER NOT NULL,
    not_in_file INTEGER NOT NULL
  );
  CREATE TABLE IF NOT EXISTS sheet_headers (
    sheet TEXT PRIMARY KEY,
    headers TEXT NOT NULL,
    last_import INTEGER
  );
  CREATE TABLE IF NOT EXISTS ticket_rows (
    sheet TEXT NOT NULL,
    ticket_id TEXT NOT NULL,
    data TEXT NOT NULL,              -- JSON array: every row of this ticket in this sheet (header → value)
    first_import INTEGER NOT NULL,
    last_change_import INTEGER NOT NULL,
    last_seen_import INTEGER NOT NULL,
    PRIMARY KEY (sheet, ticket_id)
  );
  CREATE TABLE IF NOT EXISTS changes (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    import_id INTEGER NOT NULL,
    sheet TEXT NOT NULL,
    ticket_id TEXT NOT NULL,
    field TEXT NOT NULL,
    old_value TEXT,
    new_value TEXT
  );
  CREATE INDEX IF NOT EXISTS changes_ticket ON changes(ticket_id);
  CREATE TABLE IF NOT EXISTS classifications (
    ticket_id TEXT PRIMARY KEY,
    category TEXT NOT NULL CHECK (category IN ('INC', 'OT')),
    user_name TEXT NOT NULL,
    at TEXT NOT NULL,
    note TEXT NOT NULL DEFAULT ''
  );
  CREATE TABLE IF NOT EXISTS classification_log (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    ticket_id TEXT NOT NULL,
    action TEXT NOT NULL,            -- set | remove
    category TEXT,
    user_name TEXT NOT NULL,
    at TEXT NOT NULL,
    note TEXT NOT NULL DEFAULT ''
  );
`;

function nowIso() {
  return new Date().toISOString();
}

function cleanValue(v) {
  if (v === undefined || v === '') return null;
  if (typeof v === 'string') return v.trim() === '' ? null : v;
  if (typeof v === 'number' || typeof v === 'boolean' || v === null) return v;
  return String(v);
}

/**
 * What a value *means*, for change detection: text is compared by its readable content
 * (HTML tags / entities and spacing ignored), so a weekly export that only re-encodes the HTML of a
 * description is not reported as a change. The stored value is always the one from the latest file.
 */
function meaning(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'number') return String(Math.round(v * 1e6) / 1e6);
  // decode entities first — repeatedly, for double-encoded text such as "&amp;quot;" — so that
  // "&lt;x&gt;" and "<x>" read the same; then drop tags and spacing
  let txt = String(v);
  for (let pass = 0; pass < 3; pass++) {
    const next = txt.replace(/&(#x[0-9a-f]+|#\d+|[a-z]+);/gi, (all, e) => {
      if (e[0] === '#') {
        const code = e[1].toLowerCase() === 'x' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
        return Number.isFinite(code) && code > 0 && code < 0x110000 ? String.fromCodePoint(code) : '';
      }
      return { nbsp: ' ', amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" }[e.toLowerCase()] ?? all;
    });
    if (next === txt) break;
    txt = next;
  }
  return txt
    .replace(/<[^>]*>/g, ' ')
    .replace(/[\s\u00a0]+/g, ' ')
    .trim();
}

function rowsMeaning(rows) {
  return JSON.stringify(rows.map((r) => Object.keys(r).sort().map((k) => [k, meaning(r[k])]).filter(([, v]) => v !== '')));
}

/** Change-log text of a value: dates as dates, everything else as written. */
function displayValue(field, v) {
  if (v === null || v === undefined || v === '') return '';
  if (typeof v === 'number' && /date|fecha/i.test(field)) {
    const ms = C.parseDate(v);
    if (ms !== null) return C.fmtDateTime(ms);
  }
  return String(v);
}

class SlmDatabase {
  constructor(file) {
    if (file !== ':memory:') fs.mkdirSync(path.dirname(path.resolve(file)), { recursive: true });
    this.file = file;
    this.db = new DatabaseSync(file);
    this.db.exec(SCHEMA);
  }

  close() {
    this.db.close();
  }

  stats() {
    const one = (sql) => this.db.prepare(sql).get();
    const last = one('SELECT id, file_name, imported_at, user_name FROM imports ORDER BY id DESC LIMIT 1');
    return {
      tickets: one('SELECT COUNT(DISTINCT ticket_id) AS n FROM ticket_rows').n,
      rows: one('SELECT COUNT(*) AS n FROM ticket_rows').n,
      sheets: this.db.prepare('SELECT t.sheet, COUNT(*) AS tickets FROM ticket_rows t JOIN sheet_headers h ON h.sheet = t.sheet GROUP BY t.sheet ORDER BY MIN(h.rowid)').all(),
      imports: one('SELECT COUNT(*) AS n FROM imports').n,
      lastImport: last || null,
      classifications: one('SELECT COUNT(*) AS n FROM classifications').n,
    };
  }

  /**
   * Merge one uploaded workbook.
   * @param {{fileName:string, user?:string, sheets:{name:string, rows:any[][]}[]}} payload
   *        rows = the sheet as an array of rows, including the header row.
   * @param {{dryRun?:boolean}} opts  dryRun: compute what would change, write nothing.
   */
  importWorkbook(payload, opts = {}) {
    if (!payload || !Array.isArray(payload.sheets) || !payload.sheets.length) throw new UserError('No sheets received.');
    const fileName = String(payload.fileName || 'file.xlsx').slice(0, 300);
    const user = String(payload.user || '').slice(0, 80);
    const dry = !!opts.dryRun;
    const sheetResults = [];
    const pending = []; // writes to perform
    let rowsRead = 0;

    for (const sheet of payload.sheets) {
      const name = String(sheet.name || '').slice(0, 120);
      if (!name || !Array.isArray(sheet.rows)) continue;
      const header = C.findHeader(sheet.rows);
      if (!header || header.map.ticketId === undefined) continue; // not a ticket sheet
      const headerRow = sheet.rows[header.index] || [];
      const cols = [];
      headerRow.forEach((h, c) => {
        const n = h === null || h === undefined ? '' : String(h).trim();
        if (n && !cols.some((x) => x.name === n)) cols.push({ c, name: n });
      });
      const idCol = header.map.ticketId;
      // group the rows of each ticket (a ticket may appear on several rows)
      const byTicket = new Map();
      for (let i = header.index + 1; i < sheet.rows.length; i++) {
        const r = sheet.rows[i];
        if (!Array.isArray(r)) continue;
        const id = r[idCol] === null || r[idCol] === undefined ? '' : String(r[idCol]).trim();
        if (!id || id.length > 60) continue;
        const obj = {};
        for (const { c, name: h } of cols) obj[h] = cleanValue(r[c]);
        if (!byTicket.has(id)) byTicket.set(id, []);
        byTicket.get(id).push(obj);
        rowsRead++;
      }
      const existing = new Map();
      for (const row of this.db.prepare('SELECT ticket_id, data FROM ticket_rows WHERE sheet = ?').all(name)) {
        existing.set(row.ticket_id, row.data);
      }
      let added = 0, updated = 0, unchanged = 0, changedFields = 0;
      const examples = [];
      for (const [id, rows] of byTicket) {
        const json = JSON.stringify(rows);
        const old = existing.get(id);
        if (old === undefined) {
          added++;
          pending.push({ kind: 'insert', sheet: name, id, json });
        } else if (old === json || rowsMeaning(JSON.parse(old)) === rowsMeaning(rows)) {
          unchanged++;
          pending.push({ kind: 'seen', sheet: name, id });
        } else {
          const diffs = diffRows(JSON.parse(old), rows);
          if (!diffs.length) { unchanged++; pending.push({ kind: 'seen', sheet: name, id }); continue; }
          updated++;
          changedFields += diffs.length;
          if (examples.length < 15) examples.push({ ticketId: id, changes: diffs.slice(0, 6) });
          pending.push({ kind: 'update', sheet: name, id, json, diffs });
        }
      }
      let notInFile = 0;
      for (const id of existing.keys()) if (!byTicket.has(id)) notInFile++;
      sheetResults.push({ sheet: name, tickets: byTicket.size, added, updated, unchanged, notInFile, changedFields, examples,
        headers: cols.map((x) => x.name) });
    }
    if (!sheetResults.length) throw new UserError('The file has no ticket sheet (no "Ticket ID" header row found).');

    const totals = sheetResults.reduce((a, s) => ({
      added: a.added + s.added, updated: a.updated + s.updated, unchanged: a.unchanged + s.unchanged, notInFile: a.notInFile + s.notInFile,
    }), { added: 0, updated: 0, unchanged: 0, notInFile: 0 });
    const result = { dryRun: dry, fileName, rowsRead, sheets: sheetResults, totals };
    if (dry) return result;

    this.db.exec('BEGIN');
    try {
      const imp = this.db.prepare(`INSERT INTO imports (file_name, imported_at, user_name, sheets, rows_read, added, updated, unchanged, not_in_file)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`).run(fileName, nowIso(), user, JSON.stringify(sheetResults.map((s) => s.sheet)), rowsRead,
        totals.added, totals.updated, totals.unchanged, totals.notInFile);
      const importId = Number(imp.lastInsertRowid);
      const ins = this.db.prepare('INSERT INTO ticket_rows (sheet, ticket_id, data, first_import, last_change_import, last_seen_import) VALUES (?, ?, ?, ?, ?, ?)');
      const upd = this.db.prepare('UPDATE ticket_rows SET data = ?, last_change_import = ?, last_seen_import = ? WHERE sheet = ? AND ticket_id = ?');
      const seen = this.db.prepare('UPDATE ticket_rows SET last_seen_import = ? WHERE sheet = ? AND ticket_id = ?');
      const chg = this.db.prepare('INSERT INTO changes (import_id, sheet, ticket_id, field, old_value, new_value) VALUES (?, ?, ?, ?, ?, ?)');
      for (const p of pending) {
        if (p.kind === 'insert') ins.run(p.sheet, p.id, p.json, importId, importId, importId);
        else if (p.kind === 'seen') seen.run(importId, p.sheet, p.id);
        else {
          upd.run(p.json, importId, importId, p.sheet, p.id);
          for (const d of p.diffs) chg.run(importId, p.sheet, p.id, d.field, d.old, d.new);
        }
      }
      const hdr = this.db.prepare(`INSERT INTO sheet_headers (sheet, headers, last_import) VALUES (?, ?, ?)
        ON CONFLICT(sheet) DO UPDATE SET headers = excluded.headers, last_import = excluded.last_import`);
      for (const s of sheetResults) {
        const prev = this.db.prepare('SELECT headers FROM sheet_headers WHERE sheet = ?').get(s.sheet);
        const merged = s.headers.slice();
        if (prev) for (const h of JSON.parse(prev.headers)) if (!merged.includes(h)) merged.push(h);
        hdr.run(s.sheet, JSON.stringify(merged), importId);
      }
      this.db.exec('COMMIT');
      result.importId = importId;
      return result;
    } catch (e) {
      this.db.exec('ROLLBACK');
      throw e;
    }
  }

  /** Rebuild every sheet (header row + one row per stored Excel row), plus which tickets were not in the last file. */
  exportSheets() {
    const sheets = [];
    for (const h of this.db.prepare('SELECT sheet, headers, last_import FROM sheet_headers ORDER BY rowid').all()) {
      const headers = JSON.parse(h.headers);
      const rows = [headers];
      const notInLatest = [];
      for (const r of this.db.prepare('SELECT ticket_id, data, last_seen_import FROM ticket_rows WHERE sheet = ? ORDER BY rowid').all(h.sheet)) {
        for (const obj of JSON.parse(r.data)) {
          for (const k of Object.keys(obj)) if (!headers.includes(k)) headers.push(k);
          rows.push(headers.map((k) => (obj[k] === undefined ? null : obj[k])));
        }
        if (r.last_seen_import !== h.last_import) notInLatest.push(r.ticket_id);
      }
      sheets.push({ name: h.sheet, rows, notInLatest });
    }
    return sheets;
  }

  imports() {
    return this.db.prepare('SELECT * FROM imports ORDER BY id DESC').all()
      .map((r) => ({ ...r, sheets: JSON.parse(r.sheets) }));
  }

  ticketHistory(ticketId) {
    const id = String(ticketId).slice(0, 60);
    return {
      rows: this.db.prepare(`SELECT t.sheet, f.file_name AS first_file, f.imported_at AS first_at, c.file_name AS last_change_file,
          c.imported_at AS last_change_at, s.file_name AS last_seen_file, s.imported_at AS last_seen_at
        FROM ticket_rows t JOIN imports f ON f.id = t.first_import JOIN imports c ON c.id = t.last_change_import
        JOIN imports s ON s.id = t.last_seen_import WHERE t.ticket_id = ?`).all(id),
      changes: this.db.prepare(`SELECT ch.sheet, ch.field, ch.old_value, ch.new_value, i.file_name, i.imported_at, i.user_name
        FROM changes ch JOIN imports i ON i.id = ch.import_id WHERE ch.ticket_id = ? ORDER BY ch.id`).all(id),
      classifications: this.db.prepare('SELECT action, category, user_name, at, note FROM classification_log WHERE ticket_id = ? ORDER BY id').all(id),
    };
  }

  classifications() {
    return this.db.prepare('SELECT ticket_id, category, user_name, at, note FROM classifications').all();
  }

  setClassification(ticketId, category, user, note) {
    const id = String(ticketId || '').trim().slice(0, 60);
    const cat = String(category || '').toUpperCase();
    const name = String(user || '').trim().slice(0, 80);
    const txt = String(note || '').slice(0, 500);
    if (!id) throw new UserError('Missing ticket ID.');
    if (cat !== 'INC' && cat !== 'OT') throw new UserError('Category must be INC or OT.');
    if (!name) throw new UserError('Please enter your name before classifying.');
    const at = nowIso();
    this.db.exec('BEGIN');
    try {
      this.db.prepare(`INSERT INTO classifications (ticket_id, category, user_name, at, note) VALUES (?, ?, ?, ?, ?)
        ON CONFLICT(ticket_id) DO UPDATE SET category = excluded.category, user_name = excluded.user_name, at = excluded.at, note = excluded.note`)
        .run(id, cat, name, at, txt);
      this.db.prepare('INSERT INTO classification_log (ticket_id, action, category, user_name, at, note) VALUES (?, ?, ?, ?, ?, ?)')
        .run(id, 'set', cat, name, at, txt);
      this.db.exec('COMMIT');
    } catch (e) { this.db.exec('ROLLBACK'); throw e; }
    return { ticket_id: id, category: cat, user_name: name, at, note: txt };
  }

  removeClassification(ticketId, user) {
    const id = String(ticketId || '').trim().slice(0, 60);
    const name = String(user || '').trim().slice(0, 80);
    if (!name) throw new UserError('Please enter your name.');
    this.db.exec('BEGIN');
    try {
      this.db.prepare('DELETE FROM classifications WHERE ticket_id = ?').run(id);
      this.db.prepare('INSERT INTO classification_log (ticket_id, action, category, user_name, at, note) VALUES (?, ?, NULL, ?, ?, ?)')
        .run(id, 'remove', name, nowIso(), '');
      this.db.exec('COMMIT');
    } catch (e) { this.db.exec('ROLLBACK'); throw e; }
    return { ticket_id: id };
  }

  /** Consistent copy of the whole database (for backups). */
  backupTo(file) {
    if (fs.existsSync(file)) fs.unlinkSync(file);
    this.db.prepare('VACUUM INTO ?').run(file);
    return file;
  }
}

/** Field-level differences between the stored rows of a ticket and the new ones. */
function diffRows(oldRows, newRows) {
  const diffs = [];
  const n = Math.max(oldRows.length, newRows.length);
  for (let i = 0; i < n; i++) {
    const a = oldRows[i] || {}, b = newRows[i] || {};
    const keys = new Set([...Object.keys(a), ...Object.keys(b)]);
    for (const k of keys) {
      if (meaning(a[k]) === meaning(b[k])) continue;
      diffs.push({ field: n > 1 ? `${k} (row ${i + 1})` : k, old: displayValue(k, a[k]), new: displayValue(k, b[k]) });
    }
  }
  if (oldRows.length !== newRows.length) diffs.push({ field: '(number of rows)', old: String(oldRows.length), new: String(newRows.length) });
  return diffs;
}

class UserError extends Error {}

module.exports = { SlmDatabase, UserError, diffRows };
