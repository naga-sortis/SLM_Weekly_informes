#!/usr/bin/env node
/*
 * SLM Weekly Informes — production server.
 * Serves the web page and the team database API. No external packages: Node.js ≥ 22.13.
 *
 *   node server/server.js            (or: npm start)
 *
 * Environment variables:
 *   PORT            port to listen on                  (default 8080)
 *   HOST            interface to bind                  (default 0.0.0.0 = every interface)
 *   SLM_DB          SQLite database file               (default ./data/slm-informes.sqlite)
 *   SLM_BASIC_AUTH  "user:password" to require a login (default: no login — protect the server by network/VPN)
 *   SLM_MAX_UPLOAD_MB  largest upload accepted         (default 300)
 */
'use strict';

const http = require('node:http');
const fs = require('node:fs');
const path = require('node:path');
const zlib = require('node:zlib');
const crypto = require('node:crypto');
const { SlmDatabase, UserError } = require('./db.js');

const ROOT = path.resolve(__dirname, '..');
const VERSION = (() => {
  try { return JSON.parse(fs.readFileSync(path.join(ROOT, 'package.json'), 'utf8')).version; } catch (e) { return 'dev'; }
})();

// Only these files are ever served (never the database, the server code, tests or git data).
const PUBLIC = [/^index\.html$/, /^assets\/(css|js|img)\/[\w.-]+$/, /^vendor\/[\w.-]+$/];
const TYPES = {
  '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.css': 'text/css; charset=utf-8',
  '.png': 'image/png', '.svg': 'image/svg+xml', '.ico': 'image/x-icon', '.txt': 'text/plain; charset=utf-8', '.json': 'application/json',
};
const SECURITY_HEADERS = {
  'Content-Security-Policy': "default-src 'self'; script-src 'self'; style-src 'self' 'unsafe-inline'; img-src 'self' data: blob:; " +
    "connect-src 'self'; font-src 'self'; object-src 'none'; base-uri 'none'; form-action 'none'; frame-ancestors 'none'",
  'X-Content-Type-Options': 'nosniff',
  'Referrer-Policy': 'no-referrer',
  'X-Frame-Options': 'DENY',
  'Cross-Origin-Opener-Policy': 'same-origin',
  'Permissions-Policy': 'camera=(), microphone=(), geolocation=()',
};

function createServer(options = {}) {
  const db = options.db || new SlmDatabase(options.dbFile || process.env.SLM_DB || path.join(ROOT, 'data', 'slm-informes.sqlite'));
  const maxUpload = (Number(options.maxUploadMb || process.env.SLM_MAX_UPLOAD_MB) || 300) * 1024 * 1024;
  const basicAuth = options.basicAuth !== undefined ? options.basicAuth : (process.env.SLM_BASIC_AUTH || '');
  const log = options.quiet ? () => {} : (...a) => console.log(new Date().toISOString(), ...a);

  const server = http.createServer(async (req, res) => {
    const started = Date.now();
    try {
      for (const [k, v] of Object.entries(SECURITY_HEADERS)) res.setHeader(k, v);
      if (basicAuth && !authorized(req, basicAuth)) {
        res.writeHead(401, { 'WWW-Authenticate': 'Basic realm="SLM Weekly Informes", charset="UTF-8"', 'Content-Type': 'text/plain' });
        res.end('Login required');
        return;
      }
      const url = new URL(req.url, 'http://localhost');
      if (url.pathname.startsWith('/api/')) await handleApi(req, res, url, db, maxUpload);
      else serveStatic(req, res, url);
    } catch (err) {
      const status = err instanceof UserError ? 400 : (err.status || 500);
      if (status >= 500) log('ERROR', req.method, req.url, err && err.stack ? err.stack : err);
      if (!res.headersSent) sendJson(req, res, status, { error: status >= 500 ? 'Server error — see the server log.' : err.message });
      else res.end();
    } finally {
      if (req.url.startsWith('/api/')) log(req.method, req.url, res.statusCode, `${Date.now() - started}ms`);
    }
  });
  server.db = db;
  return server;
}

function authorized(req, expected) {
  const h = req.headers.authorization || '';
  if (!h.startsWith('Basic ')) return false;
  const given = Buffer.from(h.slice(6), 'base64');
  const want = Buffer.from(expected);
  return given.length === want.length && crypto.timingSafeEqual(given, want);
}

/* ---------------------------------------------------------------- static files */

function serveStatic(req, res, url) {
  if (req.method !== 'GET' && req.method !== 'HEAD') { res.writeHead(405, { Allow: 'GET, HEAD' }); res.end(); return; }
  let rel = decodeURIComponent(url.pathname).replace(/^\/+/, '');
  if (rel === '') rel = 'index.html';
  if (rel.includes('..') || rel.includes('\\') || !PUBLIC.some((re) => re.test(rel))) { res.writeHead(404, { 'Content-Type': 'text/plain' }); res.end('Not found'); return; }
  const file = path.join(ROOT, rel);
  let stat;
  try { stat = fs.statSync(file); } catch (e) { res.writeHead(404, { 'Content-Type': 'text/plain' }); res.end('Not found'); return; }
  if (!stat.isFile()) { res.writeHead(404); res.end(); return; }
  // Always revalidate, so users never mix old and new files after an update.
  const etag = `"${stat.size.toString(16)}-${Math.floor(stat.mtimeMs).toString(16)}"`;
  const headers = { 'Content-Type': TYPES[path.extname(file).toLowerCase()] || 'application/octet-stream', 'Cache-Control': 'no-cache', ETag: etag };
  if (req.headers['if-none-match'] === etag) { res.writeHead(304, headers); res.end(); return; }
  const gzipOk = /\bgzip\b/.test(req.headers['accept-encoding'] || '') && /\.(js|css|html|svg)$/.test(file) && stat.size > 1024;
  if (gzipOk) headers['Content-Encoding'] = 'gzip';
  else headers['Content-Length'] = stat.size;
  res.writeHead(200, headers);
  if (req.method === 'HEAD') { res.end(); return; }
  const stream = fs.createReadStream(file);
  (gzipOk ? stream.pipe(zlib.createGzip()) : stream).pipe(res);
}

/* ---------------------------------------------------------------- API */

async function handleApi(req, res, url, db, maxUpload) {
  const route = `${req.method} ${url.pathname}`;
  let m;
  if (route === 'GET /api/health') return sendJson(req, res, 200, { ok: true, version: VERSION, stats: db.stats() });
  if (route === 'GET /api/sheets') return sendJson(req, res, 200, { sheets: db.exportSheets(), stats: db.stats() });
  if (route === 'GET /api/imports') return sendJson(req, res, 200, { imports: db.imports() });
  if (route === 'POST /api/import') {
    const body = await readJson(req, maxUpload);
    const dryRun = url.searchParams.get('dryRun') === '1';
    return sendJson(req, res, 200, db.importWorkbook(body, { dryRun }));
  }
  if (route === 'GET /api/classifications') return sendJson(req, res, 200, { classifications: db.classifications() });
  if ((m = url.pathname.match(/^\/api\/classifications\/([^/]{1,60})$/))) {
    const id = decodeURIComponent(m[1]);
    if (req.method === 'PUT') {
      const body = await readJson(req, 64 * 1024);
      return sendJson(req, res, 200, db.setClassification(id, body.category, body.user, body.note));
    }
    if (req.method === 'DELETE') {
      return sendJson(req, res, 200, db.removeClassification(id, url.searchParams.get('user')));
    }
  }
  if ((m = url.pathname.match(/^\/api\/tickets\/([^/]{1,60})\/history$/)) && req.method === 'GET') {
    return sendJson(req, res, 200, db.ticketHistory(decodeURIComponent(m[1])));
  }
  if (route === 'GET /api/backup') {
    const tmp = path.join(require('node:os').tmpdir(), `slm-backup-${process.pid}-${Date.now()}.sqlite`);
    db.backupTo(tmp);
    const stamp = new Date().toISOString().slice(0, 16).replace(/[:T]/g, '-');
    res.writeHead(200, { 'Content-Type': 'application/octet-stream', 'Content-Disposition': `attachment; filename="slm-informes-backup-${stamp}.sqlite"`, 'Cache-Control': 'no-store' });
    const s = fs.createReadStream(tmp);
    s.on('close', () => fs.unlink(tmp, () => {}));
    s.pipe(res);
    return;
  }
  sendJson(req, res, 404, { error: 'Unknown API route' });
}

function readJson(req, limit) {
  return new Promise((resolve, reject) => {
    const type = req.headers['content-type'] || '';
    if (!type.startsWith('application/json')) { reject(Object.assign(new UserError('Expected JSON.'), { status: 415 })); return; }
    let input = req;
    if ((req.headers['content-encoding'] || '') === 'gzip') input = req.pipe(zlib.createGunzip());
    const chunks = [];
    let size = 0;
    input.on('data', (c) => {
      size += c.length;
      if (size > limit) { reject(Object.assign(new UserError(`Upload too large (limit ${Math.round(limit / 1048576)} MB).`), { status: 413 })); req.destroy(); return; }
      chunks.push(c);
    });
    input.on('end', () => {
      try { resolve(JSON.parse(Buffer.concat(chunks).toString('utf8'))); } catch (e) { reject(new UserError('Invalid JSON.')); }
    });
    input.on('error', (e) => reject(new UserError(`Could not read the upload: ${e.message}`)));
  });
}

function sendJson(req, res, status, obj) {
  const body = Buffer.from(JSON.stringify(obj));
  const headers = { 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store' };
  if (/\bgzip\b/.test(req.headers['accept-encoding'] || '') && body.length > 2048) {
    const gz = zlib.gzipSync(body);
    res.writeHead(status, { ...headers, 'Content-Encoding': 'gzip', 'Content-Length': gz.length });
    res.end(gz);
  } else {
    res.writeHead(status, { ...headers, 'Content-Length': body.length });
    res.end(body);
  }
}

if (require.main === module) {
  const port = Number(process.env.PORT) || 8080;
  const host = process.env.HOST || '0.0.0.0';
  const server = createServer();
  server.listen(port, host, () => {
    const st = server.db.stats();
    console.log(`SLM Weekly Informes ${VERSION} — http://${host === '0.0.0.0' ? '<server-ip>' : host}:${port}/`);
    console.log(`Database: ${server.db.file} (${st.tickets} tickets, ${st.imports} imports, ${st.classifications} classifications)`);
    if (!process.env.SLM_BASIC_AUTH) console.log('No login configured (SLM_BASIC_AUTH). Restrict access with the firewall / VPN.');
  });
  const stop = () => { server.close(() => { server.db.close(); process.exit(0); }); setTimeout(() => process.exit(0), 3000).unref(); };
  process.on('SIGINT', stop);
  process.on('SIGTERM', stop);
}

module.exports = { createServer };
