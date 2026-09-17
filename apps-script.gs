/**
 * ODA Readers — Apps Script для Google Таблиці.
 *
 * Що робить:
 * 1. Приймає реєстрацію з сайту (імʼя, email, подія) і записує рядок у вкладку Signups.
 * 2. Якщо людина позначила чекбокс "Додати у свій Google Calendar" — створює один
 *    спільний Google Calendar-івент на подію (один раз) і додає її гостею туди ж.
 * 3. Раз на кілька годин (за таймером) автоматично перетворює посилання на пости
 *    в Instagram (вкладки Events і Gallery) на прямі посилання на фото — щоб банер
 *    подій і фотогалерея на сайті підтягували реальні картинки без ручного завантаження.
 *
 * Встановлення:
 * 1. Відкрий свою Google Таблицю (нативний формат, не .xlsx!).
 * 2. Меню Extensions → Apps Script.
 * 3. Видали код-заглушку, встав увесь код нижче, збережи.
 * 4. Зліва в редакторі — Services → натисни "+" → знайди "Calendar API" →
 *    Add. Без цього кроку пряме посилання "Підтвердити участь у
 *    календарі" не працюватиме (сайт тоді просто покаже запасний варіант
 *    "Перейти в календар" замість нього).
 * 5. Deploy → New deployment → тип "Web app".
 *    Execute as: Me. Who has access: Anyone.
 * 6. Deploy → дозволь доступ до Calendar і Sheets під своїм акаунтом.
 * 7. Скопіюй Web app URL і встав його в CONFIG.SIGNUP_ENDPOINT_URL у index.html сайту.
 * 8. Один раз запусти функцію createRefreshTrigger (вибери її у випадному списку
 *    вгорі редактора Apps Script і натисни Run) — це увімкне автоматичне оновлення
 *    фото з Instagram кожні 6 годин. Дозволь доступ, якщо попросить.
 * 9. Онов сторінку зі своєю таблицею — угорі зʼявиться меню "ODA Tools", де можна
 *    запустити оновлення фото вручну в будь-який момент.
 */

function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var eventsSheet = ss.getSheetByName('Events');
  var signupsSheet = ss.getSheetByName('Signups');

  // 1. Записуємо реєстрацію в таблицю
  signupsSheet.appendRow([
    new Date(), data.eventTitle, data.city, data.name, data.email,
    data.addToCalendar ? 'так' : 'ні'
  ]);

  if (!data.addToCalendar) {
    return jsonResponse({ok:true});
  }

  // 2. Знаходимо подію в таблиці Events
  var values = eventsSheet.getDataRange().getValues();
  var headers = values[0];
  var col = {};
  headers.forEach(function(h, i) { col[h] = i; });

  var rowIndex = -1;
  for (var i = 1; i < values.length; i++) {
    if (values[i][col.title] === data.eventTitle) { rowIndex = i; break; }
  }
  if (rowIndex === -1) {
    return jsonResponse({ok:false, error:'event not found'});
  }

  var row = values[rowIndex];
  var calendar = CalendarApp.getDefaultCalendar();
  var calEvent = null;
  var existingId = row[col.calendar_event_id];

  if (existingId) {
    calEvent = calendar.getEventById(existingId);
    // Захист від старого багу: якщо кешований івент має биту дату
    // (напр. 1970 рік через невірний парсинг часу) — не використовуємо
    // його повторно, перестворюємо нижче.
    if (calEvent && calEvent.getStartTime().getFullYear() < 2020) {
      calEvent.deleteEvent();
      calEvent = null;
    }
  }

  // 3. Якщо спільного івенту в календарі ще нема — створюємо один раз
  if (!calEvent) {
    var dateVal = row[col.date] instanceof Date ? row[col.date] : new Date(row[col.date]);
    var start = combineDateTime(dateVal, row[col.time_start]);
    var end = combineDateTime(dateVal, row[col.time_end]);
    calEvent = calendar.createEvent(row[col.title], start, end, {
      location: row[col.location],
      description: row[col.description]
    });
    eventsSheet.getRange(rowIndex + 1, col.calendar_event_id + 1).setValue(calEvent.getId());
  }

  // 4. Додаємо людину гостем у той самий спільний івент
  calEvent.addGuest(data.email);

  // 5. Пряме посилання "підтвердити участь" — те саме, що кнопка "Так" у
  //    email-запрошенні, але без походу в пошту (rst=1 = одразу "Так").
  var rsvpUrl = buildRsvpUrl(calEvent, calendar);

  return jsonResponse({ok:true, calendarRsvpUrl: rsvpUrl});
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function combineDateTime(dateVal, timeVal) {
  var d = new Date(dateVal.getTime());
  // Google Sheets зберігає введений час ("13:00") як Date-об'єкт
  // (внутрішньо — 30 грудня 1899 р.), а не текст — читаємо години/хвилини
  // напряму в цьому випадку, інакше беремо це як текстовий рядок "HH:MM".
  if (timeVal instanceof Date) {
    d.setHours(timeVal.getHours(), timeVal.getMinutes(), 0, 0);
  } else {
    var parts = String(timeVal).split(':');
    d.setHours(parseInt(parts[0], 10), parseInt(parts[1] || '0', 10), 0, 0);
  }
  return d;
}

/**
 * Будує пряме RSVP-посилання ("Так" з email-запрошення), використовуючи
 * справжній ідентифікатор події з Calendar API — той самий, який Google
 * підставляє в кнопку "Так" у листах. CalendarApp.getId() повертає інший
 * формат (iCalUID), який для цього посилання не підходить, тому шукаємо
 * подію через увімкнений розширений сервіс Calendar API.
 * Потребує: Apps Script → Services (+) → Calendar API.
 */
function buildRsvpUrl(calEvent, calendar) {
  var apiId = findApiEventId(calEvent.getId());
  if (!apiId) return '';
  var eid = Utilities.base64EncodeWebSafe(apiId + ' ' + calendar.getId()).replace(/=+$/, '');
  return 'https://calendar.google.com/calendar/event?action=RESPOND&eid=' + eid + '&rst=1';
}

function findApiEventId(iCalUID) {
  try {
    var res = Calendar.Events.list('primary', { iCalUID: iCalUID, maxResults: 1 });
    if (res.items && res.items.length) return res.items[0].id;
  } catch (err) {
    // Розширений сервіс Calendar API не увімкнено — просто не додаємо RSVP-лінк,
    // сайт покаже запасний варіант "Перейти в календар" замість нього.
  }
  return '';
}

/**
 * ==== АВТОМАТИЧНЕ РОЗПІЗНАВАННЯ ФОТО З INSTAGRAM ====
 * Читає посилання на пости в Instagram (з колонки image_url у Events і post_url
 * у Gallery) і зберігає пряме посилання на фото в сусідню колонку
 * (image_resolved / resolved_url), яку вже читає сайт.
 * Instagram віддає банерам-краулерам (для прев'ю посилань) картинку через
 * стандартний og:image тег — саме цей механізм тут і використовується.
 * Посилання на фото в Instagram діють кілька днів, тому функція запускається
 * за таймером (createRefreshTrigger) і оновлює їх раніше, ніж вони протухнуть.
 */
function refreshInstagramImages() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  resolveColumn(ss.getSheetByName('Events'), 'image_url', 'image_resolved');
  resolveColumn(ss.getSheetByName('Gallery'), 'post_url', 'resolved_url');
}

function resolveColumn(sheet, sourceCol, targetCol) {
  if (!sheet) return;
  var values = sheet.getDataRange().getValues();
  var headers = values[0];
  var srcIdx = headers.indexOf(sourceCol);
  var tgtIdx = headers.indexOf(targetCol);
  if (srcIdx === -1 || tgtIdx === -1) return;

  for (var i = 1; i < values.length; i++) {
    var src = String(values[i][srcIdx] || '').trim();
    if (!src || src.indexOf('instagram.com') === -1) continue;
    var resolved = resolveInstagramUrl(src);
    if (resolved) {
      sheet.getRange(i + 1, tgtIdx + 1).setValue(resolved);
    }
  }
}

function resolveInstagramUrl(url) {
  try {
    var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
    if (res.getResponseCode() >= 400) return '';
    var html = res.getContentText();
    var match = html.match(/<meta property="og:image" content="([^"]+)"/);
    if (match && match[1]) return match[1].replace(/&amp;/g, '&');
  } catch (err) {
    // Instagram іноді блокує автоматичні запити — просто пропускаємо цей рядок,
    // старе значення (якщо було) лишається, спробуємо ще раз наступного разу.
  }
  return '';
}

/** Запусти ОДИН РАЗ вручну (вибери у списку функцій угорі редактора → Run),
 *  щоб увімкнути автооновлення фото кожні 6 годин. */
function createRefreshTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'refreshInstagramImages') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('refreshInstagramImages')
    .timeBased()
    .everyHours(6)
    .create();
}

/** Додає меню "ODA Tools" у таблицю для ручного оновлення фото. */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('ODA Tools')
    .addItem('Оновити фото з Instagram зараз', 'refreshInstagramImages')
    .addToUi();
}
