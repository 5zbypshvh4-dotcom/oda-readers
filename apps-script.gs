/**
 * ODA Readers — Apps Script для Google Таблиці.
 *
 * Що робить:
 * 1. Приймає реєстрацію з сайту (імʼя, email, подія) і записує рядок у вкладку Signups.
 * 2. Створює один спільний Google Calendar-івент на подію (один раз) і додає
 *    кожну нову людину, що записалась, гостею в цей самий спільний івент.
 *
 * Встановлення:
 * 1. Відкрий свою Google Таблицю.
 * 2. Меню Extensions → Apps Script.
 * 3. Видали код-заглушку, встав увесь код нижче, збережи.
 * 4. Deploy → New deployment → тип "Web app".
 *    Execute as: Me. Who has access: Anyone.
 * 5. Deploy → дозволь доступ до Calendar і Sheets під своїм акаунтом.
 * 6. Скопіюй Web app URL і встав його в CONFIG.SIGNUP_ENDPOINT_URL у index.html сайту.
 */

function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var eventsSheet = ss.getSheetByName('Events');
  var signupsSheet = ss.getSheetByName('Signups');

  // 1. Записуємо реєстрацію в таблицю
  signupsSheet.appendRow([
    new Date(), data.eventTitle, data.city, data.name, data.email,
    data.wantsFuture ? 'так' : 'ні'
  ]);

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
    return ContentService.createTextOutput(JSON.stringify({ok:false, error:'event not found'}))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var row = values[rowIndex];
  var calendar = CalendarApp.getDefaultCalendar();
  var calEvent = null;
  var existingId = row[col.calendar_event_id];

  if (existingId) {
    calEvent = calendar.getEventById(existingId);
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

  return ContentService.createTextOutput(JSON.stringify({ok:true}))
    .setMimeType(ContentService.MimeType.JSON);
}

function combineDateTime(dateVal, timeStr) {
  var d = new Date(dateVal.getTime());
  var parts = String(timeStr).split(':');
  d.setHours(parseInt(parts[0], 10), parseInt(parts[1] || '0', 10), 0, 0);
  return d;
}
