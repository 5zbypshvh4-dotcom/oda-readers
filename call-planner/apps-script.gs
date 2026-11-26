/**
 * «Спільний час» — бекенд на Google Apps Script + Google Таблиця.
 * Безкоштовно, без серверів і без «засинання».
 *
 * Що робить:
 *  - зберігає дзвінки (вкладка Polls) і відповіді учасників (вкладка Responses)
 *    у Google Таблиці, яку видно лише тобі;
 *  - приймає запити від сторінки call-planner/index.html.
 *
 * ВСТАНОВЛЕННЯ (≈10 хвилин, один раз):
 *  1. Створи НОВУ Google Таблицю (sheets.new), назви її, напр., «Спільний час — дані».
 *     Окрему від таблиці сайту ODA Readers.
 *  2. Меню Розширення → Apps Script.
 *  3. Видали код-заглушку, встав УВЕСЬ цей файл, натисни 💾 Зберегти.
 *  4. Угорі вибери функцію setup і натисни ▶ Виконати. Дозволь доступ до таблиці
 *     (якщо Google напише «додаток не перевірено» → Додатково → Перейти до проєкту).
 *     У таблиці з'являться вкладки Polls і Responses.
 *  5. Розгорнути → Нове розгортання → тип «Веб-додаток».
 *     Виконувати як: Я. Хто має доступ: Усі (Anyone).
 *     Натисни «Розгорнути» і скопіюй URL веб-додатка (закінчується на /exec).
 *  6. Відкрий call-planner/index.html, знайди на початку скрипта рядок
 *       const API_URL = '';
 *     і встав URL між лапками. Збережи й опублікуй сайт (злий гілку в main).
 *
 * ОНОВЛЕННЯ КОДУ: після зміни цього файлу → Розгорнути → Керувати розгортаннями →
 * ✏️ → Версія: «Нова версія» → Розгорнути. URL при цьому не змінюється.
 */

var POLLS = 'Polls';
var RESP = 'Responses';
var POLL_COLS = ['id', 'adminToken', 'title', 'organizer', 'duration', 'dateFrom', 'dateTo', 'hourFrom', 'hourTo',
  'skipWeekends', 'deadline', 'tz', 'expected', 'createdAt', 'closedAt', 'final', 'deadlineText'];
var RESP_COLS = ['pollId', 'id', 'name', 'cells', 'source', 'updatedAt', 'slotsCount'];
var DURATIONS = [30, 45, 60, 90, 120];
var CELL_RE = /^\d{4}-\d{2}-\d{2}\|\d{1,4}$/;
var DATE_RE = /^\d{4}-\d{2}-\d{2}$/;

/* ---------------- точки входу ---------------- */

/** Запускається вручну один раз: створює вкладки. */
function setup() {
  sheet_(POLLS, POLL_COLS);
  sheet_(RESP, RESP_COLS);
}

function doGet() {
  return json_({ ok: true, service: 'call-planner' });
}

/** Усі запити зі сторінки приходять сюди як POST text/plain з JSON {action, ...}. */
function doPost(e) {
  var out;
  try {
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    out = handle_(body);
  } catch (err) {
    out = { error: String((err && err.message) || err) };
  }
  return json_(out);
}

function handle_(b) {
  switch (b.action) {
    case 'health': return { ok: true };
    case 'create': return withLock_(function () { return createPoll_(b); });
    case 'get': return getPoll_(b.id);
    case 'respond': return withLock_(function () { return respond_(b); });
    case 'check': adminCheck_(b); return { ok: true };
    case 'close': return withLock_(function () { return close_(b); });
    case 'finalize': return withLock_(function () { return finalize_(b); });
  }
  throw new Error('Невідома дія');
}

/* ---------------- дії ---------------- */

function createPoll_(b) {
  var d = validatePoll_(b);
  d.id = uid_(8);
  d.adminToken = uid_(16);
  d.createdAt = Date.now();
  d.closedAt = '';
  d.final = null;
  var sh = sheet_(POLLS, POLL_COLS);
  writeRow_(sh, sh.getLastRow() + 1, POLL_COLS, pollToRow_(d));
  return { id: d.id, adminToken: d.adminToken };
}

function getPoll_(id) {
  var found = findPoll_(id);
  var p = found.poll;
  delete p.adminToken;
  var responses = readAll_(sheet_(RESP, RESP_COLS), RESP_COLS)
    .filter(function (r) { return r.obj.pollId === p.id; })
    .map(function (r) { return respFromRow_(r.obj); });
  return { poll: p, responses: responses };
}

function respond_(b) {
  var p = findPoll_(b.id).poll;
  if (isClosed_(p)) throw new Error('Збір відповідей уже завершено');
  var name = str_(b.name, 60);
  if (!name) throw new Error("Вкажіть ім'я");
  var cells = [];
  var seen = {};
  (Array.isArray(b.cells) ? b.cells : []).forEach(function (c) {
    if (typeof c === 'string' && CELL_RE.test(c) && !seen[c] && cells.length < 3000) { seen[c] = 1; cells.push(c); }
  });
  var source = b.source === 'google' ? 'google' : 'manual';
  var rid = str_(b.responseId, 20);

  var sh = sheet_(RESP, RESP_COLS);
  var rows = readAll_(sh, RESP_COLS).filter(function (r) { return r.obj.pollId === p.id; });
  var hit = null;
  if (rid) hit = rows.filter(function (r) { return r.obj.id === rid; })[0] || null;
  if (!hit) hit = rows.filter(function (r) { return norm_(r.obj.name) === norm_(name); })[0] || null;

  var rec = { pollId: p.id, id: hit ? hit.obj.id : uid_(10), name: name, cells: cells, source: source, updatedAt: Date.now() };
  writeRow_(sh, hit ? hit.row : sh.getLastRow() + 1, RESP_COLS, [
    rec.pollId, rec.id, rec.name, JSON.stringify(rec.cells), rec.source, String(rec.updatedAt), String(cells.length)
  ]);
  return { id: rec.id };
}

function close_(b) {
  var found = adminCheck_(b);
  if (!found.poll.closedAt) found.poll.closedAt = Date.now();
  savePoll_(found);
  return { ok: true };
}

function finalize_(b) {
  var found = adminCheck_(b);
  var p = found.poll;
  var f = b.final;
  if (f === null) {
    p.final = null;
  } else {
    var start = Number(f && f.start);
    var link = str_(f && f.link, 300);
    if (!f || !DATE_RE.test(String(f.date)) || start !== Math.floor(start) || start < 0 || start >= 1440) throw new Error('Некоректний слот');
    if (link && !/^https?:\/\//i.test(link)) throw new Error('Некоректне посилання');
    p.final = { date: String(f.date), start: start, link: link };
    if (!p.closedAt) p.closedAt = Date.now();
  }
  savePoll_(found);
  return { ok: true };
}

/* ---------------- валідація ---------------- */

function validatePoll_(b) {
  var d = {
    title: str_(b.title, 120), organizer: str_(b.organizer, 60),
    duration: Number(b.duration), dateFrom: str_(b.dateFrom, 10), dateTo: str_(b.dateTo, 10),
    hourFrom: Number(b.hourFrom), hourTo: Number(b.hourTo), skipWeekends: !!b.skipWeekends,
    deadline: Number(b.deadline), tz: str_(b.tz || 'Europe/Kyiv', 64),
    expected: (Array.isArray(b.expected) ? b.expected : []).map(function (x) { return str_(x, 60); }).filter(String).slice(0, 50)
  };
  if (!d.title) throw new Error('Вкажіть тему дзвінка');
  if (!d.organizer) throw new Error("Вкажіть ім'я організатора");
  if (DURATIONS.indexOf(d.duration) < 0) throw new Error('Некоректна тривалість');
  if (!DATE_RE.test(d.dateFrom) || !DATE_RE.test(d.dateTo) || d.dateTo < d.dateFrom) throw new Error('Перевірте діапазон дат');
  if ((Date.parse(d.dateTo) - Date.parse(d.dateFrom)) / 864e5 > 20) throw new Error('Діапазон дат — не більше 3 тижнів');
  if (d.hourFrom !== Math.floor(d.hourFrom) || d.hourTo !== Math.floor(d.hourTo) || d.hourFrom < 0 || d.hourTo > 24 || d.hourTo <= d.hourFrom) throw new Error('Перевірте години');
  if (d.duration > (d.hourTo - d.hourFrom) * 60) throw new Error('Дзвінок довший, ніж вікно годин');
  if (!isFinite(d.deadline) || d.deadline < Date.now() || d.deadline > Date.now() + 30 * 864e5) throw new Error('Некоректний дедлайн');
  return d;
}

function adminCheck_(b) {
  var found = findPoll_(b.id);
  if (!b.adminToken || String(b.adminToken) !== found.poll.adminToken) throw new Error('Немає доступу');
  return found;
}

/* ---------------- таблиця ---------------- */

function sheet_(name, cols) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, cols.length).setNumberFormat('@').setValues([cols]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

/** Повертає [{row, obj}] для всіх рядків з даними. */
function readAll_(sh, cols) {
  var last = sh.getLastRow();
  if (last < 2) return [];
  var values = sh.getRange(2, 1, last - 1, cols.length).getValues();
  var out = [];
  for (var i = 0; i < values.length; i++) {
    if (!String(values[i][0] || '')) continue;
    var obj = {};
    for (var c = 0; c < cols.length; c++) obj[cols[c]] = String(values[i][c] === null || values[i][c] === undefined ? '' : values[i][c]);
    out.push({ row: i + 2, obj: obj });
  }
  return out;
}

/** Пише рядок як простий текст, щоб Таблиця не перетворювала дати й числа. */
function writeRow_(sh, row, cols, values) {
  sh.getRange(row, 1, 1, cols.length).setNumberFormat('@').setValues([values.map(function (v) { return v === null || v === undefined ? '' : String(v); })]);
}

function findPoll_(id) {
  id = String(id || '');
  var rows = readAll_(sheet_(POLLS, POLL_COLS), POLL_COLS);
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].obj.id === id) return { row: rows[i].row, poll: pollFromRow_(rows[i].obj) };
  }
  throw new Error('Дзвінок не знайдено');
}

function savePoll_(found) {
  writeRow_(sheet_(POLLS, POLL_COLS), found.row, POLL_COLS, pollToRow_(found.poll));
}

function pollToRow_(p) {
  return [p.id, p.adminToken, p.title, p.organizer, p.duration, p.dateFrom, p.dateTo, p.hourFrom, p.hourTo,
    p.skipWeekends ? 'TRUE' : 'FALSE', p.deadline, p.tz, JSON.stringify(p.expected || []), p.createdAt,
    p.closedAt || '', p.final ? JSON.stringify(p.final) : '',
    Utilities.formatDate(new Date(Number(p.deadline)), p.tz || 'Europe/Kyiv', 'yyyy-MM-dd HH:mm')];
}

function pollFromRow_(o) {
  return {
    id: o.id, adminToken: o.adminToken, title: o.title, organizer: o.organizer,
    duration: Number(o.duration), dateFrom: o.dateFrom, dateTo: o.dateTo,
    hourFrom: Number(o.hourFrom), hourTo: Number(o.hourTo), skipWeekends: o.skipWeekends === 'TRUE',
    deadline: Number(o.deadline), tz: o.tz || 'Europe/Kyiv', expected: parse_(o.expected, []),
    createdAt: Number(o.createdAt), closedAt: o.closedAt ? Number(o.closedAt) : null, final: parse_(o.final, null)
  };
}

function respFromRow_(o) {
  return { id: o.id, name: o.name, cells: parse_(o.cells, []), source: o.source, updatedAt: Number(o.updatedAt) };
}

/* ---------------- утиліти ---------------- */

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function isClosed_(p) { return !!p.closedAt || Date.now() >= p.deadline; }
function str_(v, max) { return String(v === null || v === undefined ? '' : v).trim().slice(0, max); }
function norm_(s) { return String(s || '').trim().toLowerCase(); }
function parse_(s, dflt) { if (!s) return dflt; try { return JSON.parse(s); } catch (e) { return dflt; } }

function uid_(n) {
  var alph = 'abcdefghijkmnpqrstuvwxyz23456789';
  var out = '';
  while (out.length < n) {
    var hex = Utilities.getUuid().replace(/-/g, '');
    for (var i = 0; i + 1 < hex.length && out.length < n; i += 2) out += alph[parseInt(hex.substr(i, 2), 16) % 32];
  }
  return out;
}
