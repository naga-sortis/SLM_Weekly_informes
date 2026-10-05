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
    const parts = raw.split(/\s+-+\s*|\s*-+\s+/).map((p) => p.trim()).filter(Boolean);
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
    tickets.sort((a, b) => b.created - a.created || (a.id < b.id ? -1 : 1));

    if (duplicates) warnings.push(`${duplicates} duplicated Ticket ID row(s) were ignored (first occurrence kept).`);
    if (invalidDates) warnings.push(`${invalidDates} row(s) without a valid Creation date were ignored.`);
    const optionalImportant = ['Third party reference', 'Processing priority', 'Initiator - Group ID',
      'Initiator - Group abbreviation name', 'Current action', 'Restoration date'];
    const miss = optionalImportant.filter((l) => missingOptional.has(l));
    if (miss.length) warnings.push(`Column(s) not found, values left empty: ${miss.join(', ')}.`);
    return { tickets, warnings, duplicates, invalidDates };
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
      // values as given in the file (fallback: computed from the creation date)
      week: toInt(get(row, 'creationWeek')) ?? iso.week,
      month: toInt(get(row, 'creationMonth')) ?? cd.getUTCMonth() + 1,
      year: toInt(get(row, 'creationYear')) ?? cd.getUTCFullYear(),
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

  function computeReport(tickets, options) {
    const weeks = options.weeks;
    const slmGroups = new Set((options.slmGroups || []).map((g) => g.trim().toUpperCase()).filter(Boolean));
    const monthsBack = Math.max(1, Math.min(24, options.monthsBack || 7));
    const lastWeek = weeks[weeks.length - 1];
    const cutoff = lastWeek.end;
    const reportMonth = options.monthKey || monthKeyOf(lastWeek.start);

    const inc = weeklyForCategory(tickets, weeks, 'INC', slmGroups);
    const ot = weeklyForCategory(tickets, weeks, 'OT', slmGroups);

    // 2.2 Casos por gestor (new incidencias of each week with a gestor in Current action)
    const gestorMap = new Map();
    weeks.forEach((w, i) => {
      for (const t of tickets) {
        if (t.category !== 'INC' || !t.gestor || t.created < w.start || t.created >= w.end) continue;
        if (!gestorMap.has(t.gestor)) gestorMap.set(t.gestor, new Array(weeks.length).fill(0));
        gestorMap.get(t.gestor)[i]++;
      }
    });
    const gestores = Array.from(gestorMap.entries())
      .map(([gestor, counts]) => ({ gestor, counts, total: counts.reduce((a, b) => a + b, 0) }))
      .sort((a, b) => a.gestor.localeCompare(b.gestor, 'es'));

    // 2.4 Detalle de los casos sin resolver (as of the end of the last week)
    const pending = tickets
      .filter((t) => t.created < cutoff && !t.devuelto && !resolvedBefore(t, cutoff))
      .sort((a, b) => b.created - a.created);

    // 3.x Monthly trends
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

    const opened = prioSeries(tickets, (t) => t.monthKey);
    const resolved = prioSeries(tickets.filter((t) => !t.devuelto && t.resolvedAt !== null && t.resolvedAt < cutoff),
      (t) => monthKeyOf(t.resolvedAt));
    const escalatedList = tickets.filter((t) => t.escalated);
    const escalated = prioSeries(escalatedList, (t) => t.monthKey);

    const escByGestor = groupSeries(escalatedList, months, (t) => t.gestor || 'SIN GESTOR', 12);
    const escByVendor = groupSeries(escalatedList, months, (t) => t.vendor, 12, VENDORS);

    const statusKeys = ['Inc. Current', 'Inc. Resolved', 'Inc. Restored', 'Inc. Closed', 'Inc. Devuelto',
      'OTs Current', 'OTs Cerradas', 'OTs Devueltas'];
    const statusData = Object.fromEntries(statusKeys.map((k) => [k, months.map(() => 0)]));
    for (const t of tickets) {
      if (!monthSet.has(t.monthKey)) continue;
      const idx = months.indexOf(t.monthKey);
      let k;
      if (t.category === 'INC') k = `Inc. ${t.statusBucket}`;
      else k = t.devuelto ? 'OTs Devueltas' : (t.statusBucket === 'Current' ? 'OTs Current' : 'OTs Cerradas');
      statusData[k][idx]++;
    }

    return {
      weeks,
      cutoff,
      reportMonth,
      inc,
      ot,
      gestores,
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

  /** Overall dataset facts for the UI. */
  function summarize(tickets) {
    let min = Infinity, max = -Infinity;
    for (const t of tickets) { if (t.created < min) min = t.created; if (t.created > max) max = t.created; }
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
    computeReport, summarize, weeklyForCategory,
  };
});
