/*
 * SLM Weekly Informes — core logic (no DOM access).
 * Parses the Oceane ticket export, extracts the report fields and computes
 * every table/series used by the on-screen report and the Word document.
 * Works in the browser (window.SLMCore) and in Node (module.exports) so it can be tested.
 */
(function (root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.SLMCore = api;
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const DAY = 86400000;
  const WEEK = 7 * DAY;

  const MONTHS_ES = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio',
    'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'];

  /* ------------------------------------------------------------------ */
  /* Column mapping                                                      */
  /* ------------------------------------------------------------------ */

  // Field key -> accepted header names (compared after normalizeHeader()).
  const FIELD_ALIASES = {
    ticketId: ['ticket id', 'ticketid', 'id ticket', 'ticket'],
    thirdParty: ['third party reference', 'third-party reference', 'referencia third party', 'third party ref'],
    creationDate: ['creation date', 'fecha de creacion', 'fecha creacion'],
    creationWeek: ['creation week', 'semana de creacion', 'semana'],
    creationMonth: ['creation month', 'mes de creacion', 'mes'],
    creationYear: ['creation year', 'ano de creacion', 'ano'],
    priority: ['processing priority', 'prioridad', 'priority'],
    ticketType: ['ticket type', 'tipo', 'tipo de ticket'],
    status: ['status', 'estado'],
    groupId: ['initiator - group id', 'initiator group id', 'grupo que lo inicia'],
    groupName: ['initiator - group abbreviation name', 'initiator group abbreviation name'],
    userName: ['initiator - user name', 'initiator user name', 'persona que lo inicia'],
    currentAction: ['current action', 'accion actual'],
    description: ['description', 'descripcion'],
    restorationDate: ['restoration date', 'fecha de restauracion'],
    restorationGroupId: ['restoration group id'],
  };

  const REQUIRED_FIELDS = ['ticketId', 'creationDate', 'ticketType', 'status'];

  const FIELD_LABELS = {
    ticketId: 'Ticket ID',
    thirdParty: 'Third party reference',
    creationDate: 'Creation date',
    creationWeek: 'Creation week',
    creationMonth: 'Creation month',
    creationYear: 'Creation year',
    priority: 'Processing priority',
    ticketType: 'Ticket type',
    status: 'Status',
    groupId: 'Initiator - Group ID',
    groupName: 'Initiator - Group abbreviation name',
    userName: 'Initiator - User name',
    currentAction: 'Current action',
    description: 'Description',
    restorationDate: 'Restoration date',
    restorationGroupId: 'Restoration group ID',
  };

  function stripAccents(s) {
    return s.normalize('NFD').replace(/[̀-ͯ]/g, '');
  }

  function normalizeHeader(v) {
    if (v === null || v === undefined) return '';
    return stripAccents(String(v)).toLowerCase()
      .replace(/[\s ]+/g, ' ')
      .replace(/\s*-\s*/g, ' - ')
      .replace(/[:.]$/, '')
      .trim();
  }

  const ALIAS_LOOKUP = (() => {
    const m = new Map();
    for (const [field, names] of Object.entries(FIELD_ALIASES)) {
      for (const n of names) m.set(normalizeHeader(n), field);
    }
    return m;
  })();

  /** Find the header row (the one containing "Ticket ID") within the first rows of a sheet. */
  function findHeader(rows, maxScan = 60) {
    const limit = Math.min(rows.length, maxScan);
    for (let i = 0; i < limit; i++) {
      const row = rows[i];
      if (!Array.isArray(row)) continue;
      const map = {};
      row.forEach((cell, col) => {
        const field = ALIAS_LOOKUP.get(normalizeHeader(cell));
        if (field && map[field] === undefined) map[field] = col;
      });
      if (map.ticketId !== undefined && Object.keys(map).length >= 3) {
        return { index: i, map };
      }
    }
    return null;
  }

  /* ------------------------------------------------------------------ */
  /* Value helpers                                                       */
  /* ------------------------------------------------------------------ */

  function str(v) {
    if (v === null || v === undefined) return '';
    if (typeof v === 'number') return Number.isInteger(v) ? String(v) : String(v);
    return String(v).replace(/ /g, ' ').trim();
  }

  function validYMD(y, mo, d, h, mi, s) {
    return y >= 1900 && y <= 2200 && mo >= 1 && mo <= 12 && d >= 1 && d <= 31 &&
      new Date(Date.UTC(y, mo - 1, d)).getUTCDate() === d && // rejects 31/02 etc.
      h >= 0 && h <= 23 && mi >= 0 && mi <= 59 && s >= 0 && s <= 59;
  }

  /**
   * Convert a cell value to a UTC timestamp (ms) holding the *wall clock* time
   * shown in Excel. Everything in the app uses UTC getters, so no timezone/DST drift.
   */
  function parseDate(v, date1904) {
    if (v === null || v === undefined || v === '') return null;
    if (typeof v === 'number' && isFinite(v)) {
      if (v <= 0 || v > 2958465) return null;
      const serial = date1904 ? v + 1462 : v;
      return Math.round((serial - 25569) * 86400) * 1000;
    }
    if (v instanceof Date) {
      if (isNaN(v.getTime())) return null;
      return Date.UTC(v.getFullYear(), v.getMonth(), v.getDate(), v.getHours(), v.getMinutes(), v.getSeconds());
    }
    const s = String(v).trim();
    let m = s.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})(?:[ T]+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
    if (m) {
      const [y, mo, d, h, mi, se] = [+m[1], +m[2], +m[3], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0)];
      return validYMD(y, mo, d, h, mi, se) ? Date.UTC(y, mo - 1, d, h, mi, se) : null;
    }
    m = s.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})(?:[ ,T]+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
    if (m) { // European order: DD/MM/YYYY
      const [d, mo, y, h, mi, se] = [+m[1], +m[2], +m[3], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0)];
      return validYMD(y, mo, d, h, mi, se) ? Date.UTC(y, mo - 1, d, h, mi, se) : null;
    }
    if (/^\d+(\.\d+)?$/.test(s)) return parseDate(Number(s), date1904);
    return null;
  }

  function toInt(v) {
    if (v === null || v === undefined || v === '') return null;
    const n = typeof v === 'number' ? v : Number(String(v).trim());
    return Number.isFinite(n) ? Math.trunc(n) : null;
  }

  function pad(n) { return String(n).padStart(2, '0'); }

  function fmtDateTime(ms) {
    if (ms === null || ms === undefined) return '';
    const d = new Date(ms);
    return `${d.getUTCFullYear()}/${pad(d.getUTCMonth() + 1)}/${pad(d.getUTCDate())} ` +
      `${pad(d.getUTCHours())}:${pad(d.getUTCMinutes())}:${pad(d.getUTCSeconds())}`;
  }

  function fmtDate(ms) {
    if (ms === null || ms === undefined) return '';
    const d = new Date(ms);
    return `${pad(d.getUTCDate())}/${pad(d.getUTCMonth() + 1)}/${d.getUTCFullYear()}`;
  }

  /** Excel serial number for a UTC wall-clock timestamp (for writing .xlsx). */
  function toExcelSerial(ms) {
    return ms / DAY + 25569;
  }

  const ENTITIES = { nbsp: ' ', amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", euro: '€',
    aacute: 'á', eacute: 'é', iacute: 'í', oacute: 'ó', uacute: 'ú', ntilde: 'ñ',
    Aacute: 'Á', Eacute: 'É', Iacute: 'Í', Oacute: 'Ó', Uacute: 'Ú', Ntilde: 'Ñ', uuml: 'ü', ordm: 'º', ordf: 'ª' };

  /** Turn the HTML found in Oceane descriptions into readable plain text. */
  function htmlToText(v) {
    let s = str(v);
    if (!s) return '';
    s = s.replace(/<\s*(br|\/p|\/div|\/tr|\/li|\/h\d)\b[^>]*>/gi, '\n')
      .replace(/<\s*\/t[dh]\s*>/gi, ' ')
      .replace(/<[^>]*>/g, '')
      .replace(/&(#x[0-9a-f]+|#\d+|[a-z]+);/gi, (all, e) => {
        if (e[0] === '#') {
          const code = e[1].toLowerCase() === 'x' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
          return Number.isFinite(code) && code > 0 && code < 0x110000 ? String.fromCodePoint(code) : '';
        }
        return Object.prototype.hasOwnProperty.call(ENTITIES, e) ? ENTITIES[e] : all;
      })
      .replace(/[ \t ]+/g, ' ')
      .replace(/ *\n */g, '\n')
      .replace(/\n{3,}/g, '\n\n');
    return s.trim();
  }

  /** Remove characters that are illegal in XML (would corrupt a .docx). */
  function xmlSafe(s) {
    // eslint-disable-next-line no-control-regex
    return String(s ?? '').replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g, '');
  }

  /* ------------------------------------------------------------------ */
  /* Calendar helpers (ISO weeks, UTC based)                             */
  /* ------------------------------------------------------------------ */

  function isoWeekInfo(ms) {
    const d = new Date(ms);
    const dayIdx = (d.getUTCDay() + 6) % 7; // Monday = 0
    const monday = Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate() - dayIdx);
    const thursday = monday + 3 * DAY;
    const year = new Date(thursday).getUTCFullYear();
    const week = 1 + Math.floor((thursday - isoWeek1Monday(year)) / WEEK);
    return { year, week, start: monday, key: year * 100 + week };
  }

  function isoWeek1Monday(year) {
    const jan4 = Date.UTC(year, 0, 4);
    const idx = (new Date(jan4).getUTCDay() + 6) % 7;
    return jan4 - idx * DAY;
  }

  function weeksInIsoYear(year) {
    return Math.round((isoWeek1Monday(year + 1) - isoWeek1Monday(year)) / WEEK);
  }

  function weekFromKey(key) {
    const year = Math.floor(key / 100);
    const week = key % 100;
    const start = isoWeek1Monday(year) + (week - 1) * WEEK;
    return { year, week, key, start, end: start + WEEK };
  }

  function nextWeekKey(key) {
    const w = weekFromKey(key);
    return isoWeekInfo(w.start + WEEK).key;
  }

  /** Inclusive list of week descriptors between two week keys. */
  function weekRange(fromKey, toKey) {
    const out = [];
    if (fromKey > toKey) [fromKey, toKey] = [toKey, fromKey];
    let k = fromKey;
    let guard = 0;
    while (k <= toKey && guard++ < 600) {
      out.push(weekFromKey(k));
      k = nextWeekKey(k);
    }
    return out;
  }

  function monthKeyOf(ms) {
    const d = new Date(ms);
    return d.getUTCFullYear() * 100 + d.getUTCMonth() + 1;
  }

  function monthLabel(key) {
    return `${MONTHS_ES[(key % 100) - 1]} ${Math.floor(key / 100)}`;
  }

  function monthShort(key) {
    return `${Math.floor(key / 100)}-${pad(key % 100)}`;
  }

  function prevMonthKey(key) {
    const y = Math.floor(key / 100), m = key % 100;
    return m === 1 ? (y - 1) * 100 + 12 : y * 100 + m - 1;
  }

  function nextMonthKey(key) {
    const y = Math.floor(key / 100), m = key % 100;
    return m === 12 ? (y + 1) * 100 + 1 : y * 100 + m + 1;
  }

  /** The N months ending at (and including) endKey, oldest first. */
  function monthsEndingAt(endKey, n) {
    const out = [endKey];
    while (out.length < n) out.unshift(prevMonthKey(out[0]));
    return out;
  }

  /** ISO weeks whose Monday falls inside the given month (how the monthly informe groups weeks). */
  function weeksOfMonth(monthKey) {
    const y = Math.floor(monthKey / 100), m = monthKey % 100;
    const first = Date.UTC(y, m - 1, 1);
    const next = Date.UTC(m === 12 ? y + 1 : y, m === 12 ? 0 : m, 1);
    const out = [];
    let d = first + ((7 - ((new Date(first).getUTCDay() + 6) % 7)) % 7) * DAY; // first Monday
    for (; d < next; d += WEEK) {
      const info = isoWeekInfo(d);
      out.push(weekFromKey(info.key));
    }
    return out;
  }

  /* ------------------------------------------------------------------ */
  /* Field derivations                                                   */
  /* ------------------------------------------------------------------ */

  const VENDORS = ['Ericsson', 'Huawei', 'Nokia'];

  /** STA- → Ericsson, H- → Huawei, 1- → Nokia. Other non-empty references → "Otro". */
  function vendorOf(ref) {
    const r = str(ref).toUpperCase();
    if (!r) return '';
    if (/^STA\s*-/.test(r)) return 'Ericsson';
    if (/^H\s*-/.test(r)) return 'Huawei';
    if (/^1\s*-\s*\S/.test(r)) return 'Nokia';
    return 'Otro';
  }

  /** "Failure" → Incidencia; every other ticket type (Work order, requests…) → OT. */
  function categoryOf(type) {
    return /^fail/i.test(str(type)) ? 'INC' : 'OT';
  }

  function normStatus(s) {
    return stripAccents(str(s)).toLowerCase();
  }

  function isOpenStatus(s) {
    const n = normStatus(s);
    return n === 'current' || n === 'open' || n === 'abierto' || n === 'en curso' || n === '';
  }

  /** Status bucket used by the "Estado de los casos" chart. */
  function statusBucket(t) {
    if (t.devuelto) return 'Devuelto';
    const n = normStatus(t.status);
    if (isOpenStatus(t.status)) return 'Current';
    if (n.startsWith('closed') || n.startsWith('cerrad')) return 'Closed';
    if (n.startsWith('resolved') || n.startsWith('resuelt')) return 'Resolved';
    if (n.startsWith('restored') || n.startsWith('restaurad')) return 'Restored';
    return 'Closed';
  }

  /**
   * Split "GESTOR - PROBLEMA - TECNICO".
   * A hyphen only separates when it has a space on at least one side,
   * so names such as "NFM-T" or "U2020-CORE" stay intact.
   */
  function parseAction(action) {
    const raw = str(action);
    const res = { gestor: '', problema: '', tecnico: '', kind: 'empty' };
    if (!raw) return res;
    const up = stripAccents(raw).toUpperCase();
    if (/^DEVUELT[OA]S?\b/.test(up)) { res.kind = 'devuelto'; return res; }
    if (/^CERRAD[OA]S?$/.test(up)) { res.kind = 'cerrado'; return res; }
    // "NFM -T" / "NFM- T" are one name: glue a hyphen followed/preceded by a 1–2 char fragment.
    const glued = raw.replace(/([A-Za-z0-9])\s+-([A-Za-z0-9]{1,2})(?=\s|$)/g, '$1-$2');
    const parts = glued.split(/\s+-+\s*|\s*-+\s+/).map((p) => p.trim()).filter(Boolean);
    // Technician work status, e.g. "JR - Trabajando - Esperando autorizacion" or "GV - ACCESO MENESES"
    // (starts with the technician's initials instead of a gestor). "EC-1" style names are kept as gestor.
    if (/TRABAJANDO/.test(up) || /^[A-Z]{2}(\s*-+\s*(?!\d)|$)/.test(up)) {
      res.kind = 'trabajando';
      res.tecnico = (parts[0] || '').split('-')[0].trim().toUpperCase();
      res.problema = parts.slice(1).join(' - ');
      return res;
    }
    res.kind = 'gestor';
    res.gestor = (parts[0] || raw).replace(/\s+/g, ' ').toUpperCase();
    if (parts.length >= 3) {
      res.problema = parts.slice(1, -1).join(' - ');
      res.tecnico = parts[parts.length - 1];
    } else if (parts.length === 2) {
      res.problema = parts[1];
    }
    return res;
  }

  function titleCase(s) {
    return s.toLowerCase().replace(/(^|[\s-])(\S)/g, (m, a, b) => a + b.toUpperCase());
  }

  /* ------------------------------------------------------------------ */
  /* Extraction                                                          */
  /* ------------------------------------------------------------------ */

  /**
   * Inspect every sheet of a workbook (array-of-arrays per sheet) and report which ones
   * contain a ticket table.
   * @param {{name:string, rows:any[][]}[]} sheets
   */
  function detectSheets(sheets) {
    return sheets.map((s) => {
      const header = findHeader(s.rows);
      if (!header) return { name: s.name, ok: false, rows: 0 };
      let count = 0;
      for (let i = header.index + 1; i < s.rows.length; i++) {
        const r = s.rows[i];
        if (r && str(r[header.map.ticketId])) count++;
      }
      const missing = REQUIRED_FIELDS.filter((f) => header.map[f] === undefined);
      return { name: s.name, ok: missing.length === 0 && count > 0, rows: count, header, missing };
    });
  }

  /**
   * Extract tickets from one or more sheets.
   * @returns {{tickets: object[], warnings: string[], duplicates: number, invalidDates: number}}
   */
  function extractTickets(sheets, opts = {}) {
    const date1904 = !!opts.date1904;
    const seen = new Map();
    const warnings = [];
    let duplicates = 0;
    let invalidDates = 0;
    const missingOptional = new Set();

    for (const sheet of sheets) {
      const header = findHeader(sheet.rows);
      if (!header) { warnings.push(`Sheet "${sheet.name}": no "Ticket ID" header row found.`); continue; }
      const map = header.map;
      const missing = REQUIRED_FIELDS.filter((f) => map[f] === undefined);
      if (missing.length) {
        warnings.push(`Sheet "${sheet.name}": missing required column(s): ${missing.map((f) => FIELD_LABELS[f]).join(', ')}.`);
        continue;
      }
      for (const f of Object.keys(FIELD_ALIASES)) if (map[f] === undefined) missingOptional.add(FIELD_LABELS[f]);
      const get = (row, f) => (map[f] === undefined ? null : row[map[f]]);

      for (let i = header.index + 1; i < sheet.rows.length; i++) {
        const row = sheet.rows[i];
        if (!row) continue;
        const id = str(get(row, 'ticketId'));
        if (!id) continue;
        if (seen.has(id)) { duplicates++; continue; }
        const created = parseDate(get(row, 'creationDate'), date1904);
        if (created === null) { invalidDates++; continue; }
        const t = buildTicket(id, created, row, get, date1904);
        seen.set(id, t);
      }
    }

    const tickets = Array.from(seen.values());
    // Two-part actions ("NCE MCA - Geiarlenys"): if the 2nd part is a known technician, it is not the problem.
    const techs = new Set(tickets.filter((t) => t.tecnico && t.actionKind === 'gestor').map((t) => t.tecnico.toUpperCase()));
    for (const t of tickets) {
      if (t.actionKind === 'gestor' && !t.tecnico && t.problema && techs.has(t.problema.toUpperCase())) {
        t.tecnico = t.problema;
        t.problema = '';
      }
      if (t.tecnico && t.actionKind === 'gestor') t.tecnico = titleCase(t.tecnico);
    }
    canonicalizeGestores(tickets);
    inferGestoresFromDescription(tickets);
    tickets.sort((a, b) => b.created - a.created || (a.id < b.id ? -1 : 1));

    // Data-quality checks on the file itself (reported, never silently "fixed").
    const weekMismatch = tickets.filter((t) => t.fileWeek !== null && t.fileWeek !== isoWeekInfo(t.created).week).length;
    const monthMismatch = tickets.filter((t) => t.fileMonth !== null && t.fileMonth !== new Date(t.created).getUTCMonth() + 1).length;
    const restoredBefore = tickets.filter((t) => t.restored !== null && t.restored < t.created).length;
    if (weekMismatch) warnings.push(`${weekMismatch} ticket(s) have a "Creation week" that does not match their Creation date; the week is recalculated from the date.`);
    if (monthMismatch) warnings.push(`${monthMismatch} ticket(s) have a "Creation month" that does not match their Creation date; the month is recalculated from the date.`);
    if (restoredBefore) warnings.push(`${restoredBefore} ticket(s) have a Restoration date earlier than the Creation date; they are treated as resolved at creation.`);

    if (duplicates) warnings.push(`${duplicates} duplicated Ticket ID row(s) were ignored (first occurrence kept).`);
    if (invalidDates) warnings.push(`${invalidDates} row(s) without a valid Creation date were ignored.`);
    const optionalImportant = ['Third party reference', 'Processing priority', 'Initiator - Group ID',
      'Initiator - Group abbreviation name', 'Current action', 'Restoration date'];
    const miss = optionalImportant.filter((l) => missingOptional.has(l));
    if (miss.length) warnings.push(`Column(s) not found, values left empty: ${miss.join(', ')}.`);
    return { tickets, warnings, duplicates, invalidDates };
  }

  /** Key used to merge spelling variants of a gestor: "NFM-T", "NFM -T", "NFMT" → "NFMT". */
  function gestorKey(g) {
    return stripAccents(str(g)).toUpperCase().replace(/[^A-Z0-9]/g, '');
  }

  /** Give every spelling variant of a gestor the same (most frequent) name so repeats are counted together. */
  function canonicalizeGestores(tickets) {
    const variants = new Map();
    for (const t of tickets) {
      if (!t.gestor) continue;
      const k = gestorKey(t.gestor);
      if (!k) { t.gestor = ''; continue; }
      if (!variants.has(k)) variants.set(k, new Map());
      const m = variants.get(k);
      m.set(t.gestor, (m.get(t.gestor) || 0) + 1);
    }
    const display = new Map();
    for (const [k, m] of variants) {
      const best = Array.from(m.entries()).sort((a, b) => b[1] - a[1] || a[0].length - b[0].length || a[0].localeCompare(b[0]))[0][0];
      display.set(k, best);
    }
    for (const t of tickets) {
      if (!t.gestor) continue;
      t.gestorRaw = t.gestor;
      t.gestor = display.get(gestorKey(t.gestor));
    }
  }

  // First words that make a "gestor" in Current action a generic task rather than a system name.
  const GENERIC_GESTOR = /^(ACCESO|APERTURA|REVISION|SOLICITUD|CONECTIVIDAD|ESCUDO|BAJA|BLOQUEO|USER|USUARIO|ABIERTO|SOLUCIONADO|GESTOR|MANOS|CAIDA|CELDA|AUTORIZADO|ESTADISTIC|ENTIDAD|PDT|CC|OSP|REE|PENDIENTE|ERICSS|HUAWEI|NOKIA|ZTE|CISCO|SALTO|BALAS|LINAREJOS|SANITAS|MIRADOR|SHARING|CINTAS)/;

  function escapeRe(c) { return c.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }

  /**
   * When Current action has no gestor (e.g. "IG - Trabajando"), look for a known gestor name
   * in the description. Only names that already appear as gestor in this file are used, and
   * the gestor is assigned only when exactly one (most specific) name matches. Never overrides
   * the gestor written in Current action.
   */
  function buildGestorMatcher(tickets) {
    const freq = new Map();
    for (const t of tickets) if (t.gestor) freq.set(t.gestor, (freq.get(t.gestor) || 0) + 1);
    const vocab = [];
    for (const [name, n] of freq) {
      const key = gestorKey(name);
      if (n < 2 || key.length < 3 || GENERIC_GESTOR.test(key)) continue;
      // chars may be separated by spaces / hyphens / dots ("SAM5620" ~ "SAM 5620"); whole-word match
      const body = key.split('').map(escapeRe).join('[\\s\\-_.]*');
      vocab.push({ name, key, re: new RegExp(`(^|[^A-Z0-9])${body}(?![A-Z0-9])`) });
    }
    return (description) => {
      if (!description) return '';
      const text = stripAccents(description).toUpperCase();
      let hits = vocab.filter((v) => v.re.test(text));
      hits = hits.filter((h) => !hits.some((o) => o !== h && o.key.length > h.key.length && o.key.includes(h.key)));
      return hits.length === 1 ? hits[0].name : '';
    };
  }

  function inferGestoresFromDescription(tickets) {
    const match = buildGestorMatcher(tickets);
    for (const t of tickets) t.gestorDesc = t.gestor || t.devuelto ? '' : match(t.description);
  }

  function buildTicket(id, created, row, get, date1904) {
    const iso = isoWeekInfo(created);
    const cd = new Date(created);
    const thirdParty = str(get(row, 'thirdParty'));
    const vendor = vendorOf(thirdParty);
    const type = str(get(row, 'ticketType'));
    const status = str(get(row, 'status'));
    const action = str(get(row, 'currentAction'));
    const pa = parseAction(action);
    const restored = parseDate(get(row, 'restorationDate'), date1904);
    const open = isOpenStatus(status);
    const prio = str(get(row, 'priority')).toUpperCase();
    const t = {
      id,
      thirdParty,
      vendor,
      escalated: VENDORS.includes(vendor),
      created,
      // Week / month / year always derived from the Creation date (ISO week, ISO week-year),
      // so they are correct even when the file's own columns are stale or inconsistent.
      week: iso.week,
      weekYear: iso.year,
      month: cd.getUTCMonth() + 1,
      year: cd.getUTCFullYear(),
      fileWeek: toInt(get(row, 'creationWeek')),
      fileMonth: toInt(get(row, 'creationMonth')),
      // values used for grouping in the report (always computed → consistent ISO weeks)
      weekKey: iso.key,
      monthKey: monthKeyOf(created),
      priority: /^P\d$/.test(prio) ? prio : (prio || 'N/D'),
      type: type || 'N/D',
      category: categoryOf(type),
      status: status || 'N/D',
      groupId: str(get(row, 'groupId')),
      groupName: str(get(row, 'groupName')),
      userName: str(get(row, 'userName')),
      action,
      actionKind: pa.kind,
      gestor: pa.gestor,
      problema: pa.problema,
      tecnico: pa.tecnico,
      devuelto: pa.kind === 'devuelto',
      description: htmlToText(get(row, 'description')),
      restored,
      restorationGroupId: str(get(row, 'restorationGroupId')),
      open,
      // When the ticket stopped being pending (null while it is still open).
      resolvedAt: open ? null : (restored !== null && restored >= created ? restored : created),
    };
    t.statusBucket = statusBucket(t);
    return t;
  }

  /* ------------------------------------------------------------------ */
  /* Report computations                                                 */
  /* ------------------------------------------------------------------ */

  function resolvedBefore(t, cutoff) {
    return t.resolvedAt !== null && t.resolvedAt < cutoff;
  }

  function countBy(list, fn) {
    const m = new Map();
    for (const x of list) {
      const k = fn(x);
      if (k === null || k === undefined) continue;
      m.set(k, (m.get(k) || 0) + 1);
    }
    return m;
  }

  /**
   * Weekly section (2.x of the informe) for one category (INC / OT).
   * "Resolved" is evaluated as of the end of each week using the Restoration date,
   * so past weeks are reported as they stood at the time, regardless of later changes.
   */
  function weeklyForCategory(tickets, weeks, cat, slmGroups) {
    const list = tickets.filter((t) => t.category === cat);
    return weeks.map((w) => {
      const created = list.filter((t) => t.created >= w.start && t.created < w.end);
      const dev = created.filter((t) => t.devuelto);
      const rest = created.filter((t) => !t.devuelto);
      const res = rest.filter((t) => resolvedBefore(t, w.end));
      const unres = rest.filter((t) => !resolvedBefore(t, w.end));
      const backlog = list.filter((t) => t.created < w.start && !t.devuelto && !resolvedBefore(t, w.start));
      const blNo = backlog.filter((t) => !t.escalated);
      const blEsc = backlog.filter((t) => t.escalated);
      const inWeek = (t) => resolvedBefore(t, w.end);
      return {
        week: w,
        nuevos: created.length,
        resueltas: res.length,
        sinResolverNoEsc: unres.filter((t) => !t.escalated).length,
        sinResolverEsc: unres.filter((t) => t.escalated).length,
        devueltas: dev.length,
        abiertasSLM: created.filter((t) => slmGroups.has(t.groupId.toUpperCase())).length,
        backlogNoEsc: { inicio: blNo.length, resueltos: blNo.filter(inWeek).length, sinResolver: blNo.filter((t) => !inWeek(t)).length },
        backlogEsc: { inicio: blEsc.length, resueltos: blEsc.filter(inWeek).length, sinResolver: blEsc.filter((t) => !inWeek(t)).length },
      };
    });
  }

  /** "Septiembre 2026" or, for a week spanning two months, "Septiembre / Octubre 2026". */
  function periodMonthsLabel(startMs, endMsExclusive) {
    const a = monthKeyOf(startMs);
    const b = monthKeyOf(endMsExclusive - 1);
    if (a === b) return monthLabel(a);
    const ya = Math.floor(a / 100), yb = Math.floor(b / 100);
    const ma = MONTHS_ES[(a % 100) - 1], mb = MONTHS_ES[(b % 100) - 1];
    return ya === yb ? `${ma} / ${mb} ${ya}` : `${ma} ${ya} / ${mb} ${yb}`;
  }

  /**
   * Build the whole report for a window of weeks; the LAST week is the report week.
   * Nothing created or resolved after the end of the report week is counted anywhere.
   */
  function computeReport(tickets, options) {
    const weeks = options.weeks;
    const slmGroups = new Set((options.slmGroups || []).map((g) => g.trim().toUpperCase()).filter(Boolean));
    const monthsBack = Math.max(1, Math.min(24, options.monthsBack || 7));
    const lastWeek = weeks[weeks.length - 1];
    const cutoff = lastWeek.end;
    let dataUntil = options.dataUntil ?? null;
    if (dataUntil === null) {
      for (const t of tickets) {
        if (dataUntil === null || t.created > dataUntil) dataUntil = t.created;
        if (t.restored !== null && t.restored > dataUntil) dataUntil = t.restored;
      }
    }
    // Last moment actually covered by both the week and the data.
    const asOf = dataUntil === null ? cutoff - 1000 : Math.min(cutoff - 1000, dataUntil);
    const reportMonth = monthKeyOf(Math.max(lastWeek.start, asOf));
    const periodLabel = periodMonthsLabel(lastWeek.start, lastWeek.end);
    const scope = tickets.filter((t) => t.created < cutoff);

    const inc = weeklyForCategory(scope, weeks, 'INC', slmGroups);
    const ot = weeklyForCategory(scope, weeks, 'OT', slmGroups);

    // 2.2 Casos por gestor: EVERY new incidencia of each week, under the gestor taken from
    // "Current action" (GESTOR - PROBLEMA - TECNICO). Repeated gestor → +1 on the same row.
    // Incidencias without a gestor in the action go to "Sin gestor identificado" so the
    // column totals always equal "Nuevos durante la semana" (Incidencias) of table 2.1.
    const NO_GESTOR = 'Sin gestor identificado';
    const inferGestor = !!options.inferGestor;
    const gOf = (t) => t.gestor || (inferGestor ? t.gestorDesc : '') || '';
    let inferredCount = 0;
    const gestorMap = new Map();
    weeks.forEach((w, i) => {
      for (const t of scope) {
        if (t.category !== 'INC' || t.created < w.start || t.created >= w.end) continue;
        if (!t.gestor && gOf(t)) inferredCount++;
        const g = gOf(t) || NO_GESTOR;
        if (!gestorMap.has(g)) gestorMap.set(g, new Array(weeks.length).fill(0));
        gestorMap.get(g)[i]++;
      }
    });
    const gestores = Array.from(gestorMap.entries())
      .map(([gestor, counts]) => ({ gestor, counts, total: counts.reduce((a, b) => a + b, 0), unidentified: gestor === NO_GESTOR }))
      .sort((a, b) => (a.unidentified - b.unidentified) || a.gestor.localeCompare(b.gestor, 'es'));

    // 2.4 Detalle de los casos sin resolver (as of the end of the report week)
    const pending = scope
      .filter((t) => !t.devuelto && !resolvedBefore(t, cutoff))
      .sort((a, b) => b.created - a.created);

    // 3.x Monthly trends (up to the end of the report week)
    const months = monthsEndingAt(reportMonth, monthsBack);
    const monthSet = new Set(months);
    const prioSeries = (list, monthOf) => {
      const keys = ['OTs', 'Inc. P1', 'Inc. P2', 'Inc. P3', 'Inc. P4', 'Inc. otras'];
      const data = Object.fromEntries(keys.map((k) => [k, months.map(() => 0)]));
      for (const t of list) {
        const mk = monthOf(t);
        if (!monthSet.has(mk)) continue;
        const idx = months.indexOf(mk);
        let k = 'OTs';
        if (t.category === 'INC') k = /^P[1-4]$/.test(t.priority) ? `Inc. ${t.priority}` : 'Inc. otras';
        data[k][idx]++;
      }
      return prune(keys, data);
    };

    const opened = prioSeries(scope, (t) => t.monthKey);
    const resolved = prioSeries(scope.filter((t) => !t.devuelto && resolvedBefore(t, cutoff)),
      (t) => monthKeyOf(t.resolvedAt));
    const escalatedList = scope.filter((t) => t.escalated);
    const escalated = prioSeries(escalatedList, (t) => t.monthKey);

    const escByGestor = groupSeries(escalatedList, months, (t) => gOf(t) || NO_GESTOR, 12);
    const escByVendor = groupSeries(escalatedList, months, (t) => t.vendor, 12, VENDORS);

    // Status as it was at the end of the report week.
    const statusKeys = ['Inc. Current', 'Inc. Resolved', 'Inc. Restored', 'Inc. Closed', 'Inc. Devuelto',
      'OTs Current', 'OTs Cerradas', 'OTs Devueltas'];
    const statusData = Object.fromEntries(statusKeys.map((k) => [k, months.map(() => 0)]));
    for (const t of scope) {
      if (!monthSet.has(t.monthKey)) continue;
      const idx = months.indexOf(t.monthKey);
      const bucket = t.devuelto ? 'Devuelto' : (resolvedBefore(t, cutoff) ? t.statusBucket : 'Current');
      let k;
      if (t.category === 'INC') k = `Inc. ${bucket}`;
      else k = bucket === 'Devuelto' ? 'OTs Devueltas' : (bucket === 'Current' ? 'OTs Current' : 'OTs Cerradas');
      statusData[k][idx]++;
    }

    // How the report week's new cases fall into calendar months (a week can straddle two months,
    // so the monthly charts and the weekly figures are reconciled explicitly).
    const weekSplit = [];
    for (let d = lastWeek.start; d < lastWeek.end; d += DAY) {
      const mk = monthKeyOf(d);
      let part = weekSplit.find((x) => x.monthKey === mk);
      if (!part) { part = { monthKey: mk, from: d, to: d, inc: 0, ot: 0 }; weekSplit.push(part); }
      part.to = d;
    }
    for (const t of scope) {
      if (t.created < lastWeek.start || t.created >= lastWeek.end) continue;
      const part = weekSplit.find((x) => x.monthKey === t.monthKey);
      if (t.category === 'INC') part.inc++; else part.ot++;
    }
    const partialMonthNow = monthKeyOf(asOf) === reportMonth && monthKeyOf(asOf + DAY) === reportMonth;
    const monthLabels = months.map((m) => (m === reportMonth && partialMonthNow ? `${monthShort(m)} (hasta ${fmtDate(asOf).slice(0, 5)})` : monthShort(m)));

    const report = {
      weeks,
      weekSplit,
      monthLabels,
      reportWeek: lastWeek,
      cutoff,
      asOf,
      dataUntil,
      partialWeek: dataUntil !== null && dataUntil < cutoff - 1000,
      periodLabel,
      reportMonth,
      partialMonth: partialMonthNow,
      inc,
      ot,
      gestores,
      gestorInferred: inferGestor ? inferredCount : 0,
      gestorUnidentified: (gestores.find((g) => g.unidentified) || { total: 0 }).total,
      pending,
      months,
      monthly: {
        opened,
        resolved,
        escalated,
        escByGestor,
        escByVendor,
        status: prune(statusKeys, statusData),
      },
    };
    report.checks = validateReport(report, scope);
    return report;
  }

  /**
   * Consistency checks run on every report. Any failure is shown in the UI and blocks the
   * Word export, so a report with inconsistent figures can never be sent to the customer.
   */
  function validateReport(r, scope) {
    const checks = [];
    const add = (ok, label, detail) => checks.push({ ok: !!ok, label, detail: ok ? '' : detail });
    const n = r.weeks.length;

    // Weeks are consecutive ISO weeks, 7 days each
    let consecutive = true;
    for (let i = 0; i < n; i++) {
      const w = r.weeks[i];
      if (w.end - w.start !== WEEK || isoWeekInfo(w.start).key !== w.key || (i && w.start !== r.weeks[i - 1].end)) consecutive = false;
    }
    add(consecutive, 'Weeks are consecutive ISO weeks (Monday–Sunday)', 'The selected weeks are not consecutive.');

    const sum = (arr) => arr.reduce((a, b) => a + b, 0);
    for (const [cat, rows, name] of [['INC', r.inc, 'Incidencias'], ['OT', r.ot, 'OTs']]) {
      const list = scope.filter((t) => t.category === cat);
      const bad = [];
      const badBl = [];
      const badCarry = [];
      const badCount = [];
      rows.forEach((x, i) => {
        const w = r.weeks[i];
        if (x.nuevos !== x.resueltas + x.sinResolverNoEsc + x.sinResolverEsc + x.devueltas) bad.push(w.week);
        if (x.backlogNoEsc.inicio !== x.backlogNoEsc.resueltos + x.backlogNoEsc.sinResolver ||
            x.backlogEsc.inicio !== x.backlogEsc.resueltos + x.backlogEsc.sinResolver) badBl.push(w.week);
        if (i > 0) {
          const p = rows[i - 1];
          if (x.backlogNoEsc.inicio !== p.backlogNoEsc.sinResolver + p.sinResolverNoEsc ||
              x.backlogEsc.inicio !== p.backlogEsc.sinResolver + p.sinResolverEsc) badCarry.push(w.week);
        }
        const direct = list.filter((t) => t.created >= w.start && t.created < w.end).length;
        if (direct !== x.nuevos) badCount.push(w.week);
      });
      add(!badCount.length, `${name}: "Nuevos" equals the tickets created in each week`, `Mismatch in week(s) ${badCount.join(', ')}.`);
      add(!bad.length, `${name}: Nuevos = Resueltas + Sin resolver (no esc.) + Sin resolver (esc.) + Devueltas`, `Does not add up in week(s) ${bad.join(', ')}.`);
      add(!badBl.length, `${name}: Backlog = Resueltos + Sin resolver`, `Does not add up in week(s) ${badBl.join(', ')}.`);
      if (n > 1) add(!badCarry.length, `${name}: each week's backlog = previous backlog pending + previous week's new pending`, `Carry-over breaks in week(s) ${badCarry.join(', ')}.`);
      const last = rows[n - 1];
      const totalPending = last.sinResolverNoEsc + last.sinResolverEsc + last.backlogNoEsc.sinResolver + last.backlogEsc.sinResolver;
      const listed = r.pending.filter((t) => t.category === cat).length;
      add(totalPending === listed, `${name}: 2.3 TOTAL of the report week = cases listed in 2.4`, `2.3 TOTAL is ${totalPending} but 2.4 lists ${listed}.`);
    }

    // 2.2 vs 2.1
    const badG = [];
    r.weeks.forEach((w, i) => {
      const g = sum(r.gestores.map((x) => x.counts[i]));
      if (g !== r.inc[i].nuevos) badG.push(`sem. ${w.week} (${g} vs ${r.inc[i].nuevos})`);
    });
    add(!badG.length, '2.2 Casos por gestor: weekly totals = "Nuevos durante la semana" (Incidencias) in 2.1', `Mismatch: ${badG.join(', ')}.`);
    const dupNames = new Set();
    const keys = new Set();
    for (const g of r.gestores) { const k = gestorKey(g.gestor); if (keys.has(k)) dupNames.add(g.gestor); keys.add(k); }
    add(!dupNames.size, '2.2: each gestor appears only once (spelling variants merged)', `Repeated: ${Array.from(dupNames).join(', ')}.`);

    // Nothing after the cut-off, pending list really unresolved
    add(r.pending.every((t) => t.created < r.cutoff && !t.devuelto && !resolvedBefore(t, r.cutoff)),
      '2.4: every listed case was created before the end of the week and still unresolved then', 'Some listed cases are not pending at the cut-off.');
    const openedLast = sum(r.monthly.opened.map((s) => s.data[s.data.length - 1]));
    const directOpened = scope.filter((t) => t.monthKey === r.reportMonth).length;
    add(openedLast === directOpened, `3.1: cases opened in ${monthLabel(r.reportMonth)} = tickets created that month up to the cut-off`, `Chart shows ${openedLast}, data has ${directOpened}.`);
    const statusLast = sum(r.monthly.status.map((s) => s.data[s.data.length - 1]));
    add(statusLast === directOpened, '3.4: status chart covers every case of the month exactly once', `Chart shows ${statusLast}, data has ${directOpened}.`);

    // Weekly figures vs monthly charts for the report week
    const last = n - 1;
    const splitTotal = r.weekSplit.reduce((a, p) => a + p.inc + p.ot, 0);
    add(splitTotal === r.inc[last].nuevos + r.ot[last].nuevos,
      `Week ${r.reportWeek.week}: new cases split by month (${r.weekSplit.map((p) => `${monthShort(p.monthKey)}: ${p.inc + p.ot}`).join(', ')}) add up to the week total`,
      `Split gives ${splitTotal}, week total is ${r.inc[last].nuevos + r.ot[last].nuevos}.`);
    const badMonth = r.weekSplit.filter((p) => {
      const i = r.months.indexOf(p.monthKey);
      if (i < 0) return false;
      const bar = sum(r.monthly.opened.map((x) => x.data[i]));
      return bar < p.inc + p.ot;
    });
    add(!badMonth.length, '3.1: each monthly bar includes the report-week cases of that month', `Bar too low for ${badMonth.map((p) => monthShort(p.monthKey)).join(', ')}.`);

    // Dates
    add(monthKeyOf(r.reportWeek.start) === r.reportMonth || monthKeyOf(r.reportWeek.end - 1) === r.reportMonth,
      `Report month (${monthLabel(r.reportMonth)}) contains the report week ${r.reportWeek.week}`, 'Month and week do not match.');
    return checks;
  }

  function prune(keys, data) {
    const series = keys.filter((k) => data[k].some((v) => v > 0)).map((k) => ({ name: k, data: data[k] }));
    return series.length ? series : [{ name: keys[0], data: data[keys[0]] }];
  }

  /** Series per group (top N groups by volume, rest merged into "Otros"). */
  function groupSeries(list, months, keyFn, topN, fixedOrder) {
    const monthSet = new Set(months);
    const inRange = list.filter((t) => monthSet.has(t.monthKey));
    const totals = countBy(inRange, keyFn);
    let names = fixedOrder ? fixedOrder.filter((n) => totals.has(n)) :
      Array.from(totals.keys()).sort((a, b) => totals.get(b) - totals.get(a) || a.localeCompare(b, 'es'));
    let others = [];
    if (names.length > topN) { others = names.slice(topN - 1); names = names.slice(0, topN - 1); }
    const otherSet = new Set(others);
    const data = Object.fromEntries(names.map((n) => [n, months.map(() => 0)]));
    if (others.length) data.Otros = months.map(() => 0);
    for (const t of inRange) {
      const k = keyFn(t);
      const idx = months.indexOf(t.monthKey);
      if (otherSet.has(k)) data.Otros[idx]++;
      else if (data[k]) data[k][idx]++;
    }
    const keys = Object.keys(data);
    return keys.map((k) => ({ name: k, data: data[k] }));
  }

  /** Spanish sentence reconciling the report week with the monthly charts. */
  function weekSplitText(r) {
    const w = r.reportWeek;
    const total = r.weekSplit.reduce((a, p) => a + p.inc + p.ot, 0);
    const head = `La semana ${w.week} (${fmtDate(w.start).slice(0, 5)}–${fmtDate(w.end - DAY).slice(0, 5)}) tiene ${total} casos nuevos ` +
      `(${r.inc[r.inc.length - 1].nuevos} incidencias y ${r.ot[r.ot.length - 1].nuevos} OTs)`;
    let txt;
    if (r.weekSplit.length > 1) {
      txt = `${head}. Como la semana abarca dos meses, en los gráficos mensuales se reparten por fecha de creación: ` +
        r.weekSplit.map((p) => `${p.inc + p.ot} en ${MONTHS_ES[(p.monthKey % 100) - 1].toLowerCase()} (${p.from === p.to ? fmtDate(p.from).slice(0, 5) : `${fmtDate(p.from).slice(0, 5)}–${fmtDate(p.to).slice(0, 5)}`}: ${p.inc} incidencias, ${p.ot} OTs)`).join(' y ') + '.';
    } else {
      txt = `${head}, todos ellos en ${monthLabel(r.weekSplit[0].monthKey).toLowerCase()}.`;
    }
    if (r.partialMonth) txt += ` El mes de ${monthLabel(r.reportMonth).toLowerCase()} incluye solo los datos hasta el ${fmtDate(r.asOf)}.`;
    return txt;
  }

  /** Overall dataset facts for the UI. */
  function summarize(tickets) {
    let min = Infinity, max = -Infinity;
    let until = -Infinity;
    for (const t of tickets) {
      if (t.created < min) min = t.created;
      if (t.created > max) max = t.created;
      if (t.created > until) until = t.created;
      if (t.restored !== null && t.restored > until) until = t.restored;
    }
    const weekKeys = Array.from(new Set(tickets.map((t) => t.weekKey))).sort((a, b) => a - b);
    const monthKeys = Array.from(new Set(tickets.map((t) => t.monthKey))).sort((a, b) => a - b);
    return {
      total: tickets.length,
      incidencias: tickets.filter((t) => t.category === 'INC').length,
      ots: tickets.filter((t) => t.category === 'OT').length,
      open: tickets.filter((t) => t.open && !t.devuelto).length,
      escalated: tickets.filter((t) => t.escalated).length,
      minDate: tickets.length ? min : null,
      maxDate: tickets.length ? max : null,
      dataUntil: tickets.length ? until : null,
      weekKeys,
      monthKeys,
    };
  }

  return {
    FIELD_ALIASES, FIELD_LABELS, REQUIRED_FIELDS, VENDORS, MONTHS_ES,
    normalizeHeader, findHeader, detectSheets, extractTickets,
    parseDate, htmlToText, xmlSafe, vendorOf, categoryOf, parseAction, statusBucket,
    isoWeekInfo, weekFromKey, weekRange, weeksOfMonth, weeksInIsoYear, nextWeekKey,
    monthKeyOf, monthLabel, monthShort, monthsEndingAt, prevMonthKey, nextMonthKey,
    fmtDateTime, fmtDate, toExcelSerial,
    computeReport, summarize, weeklyForCategory, validateReport, gestorKey, periodMonthsLabel, buildGestorMatcher, weekSplitText,
  };
});
