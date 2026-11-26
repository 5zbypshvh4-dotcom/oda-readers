#!/usr/bin/env node
/*
 * «Спільний час» — мінімальний сервер для прототипу.
 * Без залежностей: потрібен лише Node.js 18+.
 *
 *   node server.js            → http://localhost:3000
 *   PORT=8080 node server.js
 *
 * Дані зберігаються у data/polls.json (папка в .gitignore).
 */
'use strict';
const http = require('http');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

const PORT = Number(process.env.PORT) || 3000;
const DATA_DIR = path.join(__dirname, 'data');
const DATA_FILE = path.join(DATA_DIR, 'polls.json');
const INDEX = path.join(__dirname, 'index.html');

/* ---------- сховище ---------- */
let db = { polls: {} };
try { db = JSON.parse(fs.readFileSync(DATA_FILE, 'utf8')); } catch { /* перший запуск */ }
function persist() {
  fs.mkdirSync(DATA_DIR, { recursive: true });
  const tmp = DATA_FILE + '.tmp';
  fs.writeFileSync(tmp, JSON.stringify(db, null, 2));
  fs.renameSync(tmp, DATA_FILE);
}

/* ---------- утиліти ---------- */
const ALPH = 'abcdefghijkmnpqrstuvwxyz23456789';
const uid = n => Array.from(crypto.randomBytes(n), b => ALPH[b % 32]).join('');
const norm = s => String(s || '').trim().toLowerCase();
const str = (v, max) => String(v ?? '').trim().slice(0, max);
const isDate = s => /^\d{4}-\d{2}-\d{2}$/.test(s) && !isNaN(Date.parse(s + 'T00:00:00Z'));
const CELL_RE = /^\d{4}-\d{2}-\d{2}\|\d{1,4}$/;
const isClosed = p => !!p.closedAt || Date.now() >= p.deadline;
const safeEqual = (a, b) => {
  const x = Buffer.from(String(a || '')), y = Buffer.from(String(b || ''));
  return x.length === y.length && crypto.timingSafeEqual(x, y);
};

function send(res, code, obj) {
  res.writeHead(code, { 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store' });
  res.end(JSON.stringify(obj));
}
function readBody(req) {
  return new Promise((resolve, reject) => {
    let size = 0; const chunks = [];
    req.on('data', c => { size += c.length; if (size > 512 * 1024) { reject(new Error('Запит завеликий')); req.destroy(); } else chunks.push(c); });
    req.on('end', () => { try { resolve(chunks.length ? JSON.parse(Buffer.concat(chunks).toString('utf8')) : {}); } catch { reject(new Error('Некоректний JSON')); } });
    req.on('error', reject);
  });
}
function publicPoll(p) {
  const { adminToken, responses, ...poll } = p;
  return { poll, responses };
}

function validatePoll(b) {
  const d = {
    title: str(b.title, 120), organizer: str(b.organizer, 60),
    duration: Number(b.duration), dateFrom: str(b.dateFrom, 10), dateTo: str(b.dateTo, 10),
    hourFrom: Number(b.hourFrom), hourTo: Number(b.hourTo), skipWeekends: !!b.skipWeekends,
    deadline: Number(b.deadline), tz: str(b.tz || 'Europe/Kyiv', 64),
    expected: Array.isArray(b.expected) ? b.expected.map(x => str(x, 60)).filter(Boolean).slice(0, 50) : []
  };
  if (!d.title) return [null, 'Вкажіть тему дзвінка'];
  if (!d.organizer) return [null, "Вкажіть ім'я організатора"];
  if (![30, 45, 60, 90, 120].includes(d.duration)) return [null, 'Некоректна тривалість'];
  if (!isDate(d.dateFrom) || !isDate(d.dateTo) || d.dateTo < d.dateFrom) return [null, 'Перевірте діапазон дат'];
  if ((Date.parse(d.dateTo) - Date.parse(d.dateFrom)) / 864e5 > 20) return [null, 'Діапазон дат — не більше 3 тижнів'];
  if (!Number.isInteger(d.hourFrom) || !Number.isInteger(d.hourTo) || d.hourFrom < 0 || d.hourTo > 24 || d.hourTo <= d.hourFrom) return [null, 'Перевірте години'];
  if (d.duration > (d.hourTo - d.hourFrom) * 60) return [null, 'Дзвінок довший, ніж вікно годин'];
  if (!Number.isFinite(d.deadline) || d.deadline < Date.now() || d.deadline > Date.now() + 30 * 864e5) return [null, 'Некоректний дедлайн'];
  return [d, null];
}

/* ---------- API ---------- */
async function api(req, res, parts) {
  // parts: ['api', ...]
  if (parts[1] === 'health') return send(res, 200, { ok: true });

  if (parts[1] === 'polls' && parts.length === 2 && req.method === 'POST') {
    const [d, err] = validatePoll(await readBody(req));
    if (err) return send(res, 400, { error: err });
    const poll = { ...d, id: uid(8), adminToken: uid(16), createdAt: Date.now(), closedAt: null, final: null, responses: [] };
    db.polls[poll.id] = poll; persist();
    return send(res, 201, { id: poll.id, adminToken: poll.adminToken });
  }

  const p = parts[1] === 'polls' && db.polls[parts[2]];
  if (!p) return send(res, 404, { error: 'Дзвінок не знайдено' });

  if (parts.length === 3 && req.method === 'GET') return send(res, 200, publicPoll(p));

  if (req.method !== 'POST') return send(res, 405, { error: 'Метод не підтримується' });
  const body = await readBody(req);

  if (parts[3] === 'responses') {
    if (isClosed(p)) return send(res, 409, { error: 'Збір відповідей уже завершено' });
    const name = str(body.name, 60);
    if (!name) return send(res, 400, { error: "Вкажіть ім'я" });
    const cells = Array.isArray(body.cells) ? [...new Set(body.cells.filter(c => typeof c === 'string' && CELL_RE.test(c)))].slice(0, 3000) : [];
    const source = body.source === 'google' ? 'google' : 'manual';
    let r = body.id && p.responses.find(x => x.id === body.id);
    if (!r) r = p.responses.find(x => norm(x.name) === norm(name));
    if (r) Object.assign(r, { name, cells, source, updatedAt: Date.now() });
    else { r = { id: uid(10), name, cells, source, updatedAt: Date.now() }; p.responses.push(r); }
    persist();
    return send(res, 200, { id: r.id });
  }

  if (!safeEqual(body.adminToken, p.adminToken)) return send(res, 403, { error: 'Немає доступу' });

  if (parts[3] === 'check') return send(res, 200, { ok: true });
  if (parts[3] === 'close') { p.closedAt = p.closedAt || Date.now(); persist(); return send(res, 200, { ok: true }); }
  if (parts[3] === 'finalize') {
    const f = body.final;
    if (f === null) p.final = null;
    else {
      const start = Number(f?.start);
      const link = str(f?.link, 300);
      if (!isDate(f?.date) || !Number.isInteger(start) || start < 0 || start >= 1440) return send(res, 400, { error: 'Некоректний слот' });
      if (link && !/^https?:\/\//i.test(link)) return send(res, 400, { error: 'Некоректне посилання' });
      p.final = { date: f.date, start, link };
      p.closedAt = p.closedAt || Date.now();
    }
    persist();
    return send(res, 200, { ok: true });
  }
  return send(res, 404, { error: 'Не знайдено' });
}

/* ---------- сервер ---------- */
http.createServer(async (req, res) => {
  try {
    const url = new URL(req.url, 'http://x');
    const parts = url.pathname.split('/').filter(Boolean);
    if (parts[0] === 'api') return await api(req, res, parts);
    if (req.method === 'GET' && (url.pathname === '/' || url.pathname === '/index.html')) {
      res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
      return fs.createReadStream(INDEX).pipe(res);
    }
    send(res, 404, { error: 'Не знайдено' });
  } catch (e) {
    send(res, 400, { error: e.message || 'Помилка' });
  }
}).listen(PORT, () => console.log(`Спільний час: http://localhost:${PORT}`));
