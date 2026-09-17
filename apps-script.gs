/**
 * ODA Readers - Apps Script for the Google Sheet.
 *
 * What it does:
 * 1. Accepts signups from the site (name, email, event) and appends a row to the Signups tab.
 * 2. If the person checked "Add to my Google Calendar" - creates one shared Google Calendar
 *    event per event (once) and adds them as a guest to that same event.
 * 3. Every few hours (on a timer) automatically resolves Instagram post links (Events and
 *    Gallery tabs) into direct photo URLs, so the event banner and photo gallery on the site
 *    can show real pictures without manual uploads.
 *
 * Setup:
 * 1. Open your Google Sheet (native format, not .xlsx!).
 * 2. Menu Extensions -> Apps Script.
 * 3. Delete the placeholder code, paste in ALL the code below, save.
 * 4. On the left of the editor - Services -> click "+" -> find "Calendar API" -> Add.
 *    Without this step the direct "Confirm attendance in calendar" link won't work
 *    (the site will just fall back to the "Open in calendar" personal-copy link instead).
 * 5. Deploy -> New deployment -> type "Web app".
 *    Execute as: Me. Who has access: Anyone.
 * 6. Deploy -> allow access to Calendar and Sheets with your account.
 * 7. Copy the Web app URL and paste it into CONFIG.SIGNUP_ENDPOINT_URL in the site's index.html.
 * 8. Run the createRefreshTrigger function once manually (pick it from the function dropdown
 *    at the top of the Apps Script editor and click Run) - this turns on automatic photo
 *    refreshing from Instagram every 6 hours. Grant access if asked.
 * 9. Reload the page with your spreadsheet - an "ODA Tools" menu will appear at the top,
 *    where you can trigger a photo refresh manually at any time.
 */

function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var eventsSheet = ss.getSheetByName('Events');
  var signupsSheet = ss.getSheetByName('Signups');

  // 1. Record the signup in the sheet
  signupsSheet.appendRow([
    new Date(), data.eventTitle, data.city, data.name, data.email,
    data.addToCalendar ? 'так' : 'ні'
  ]);

  if (!data.addToCalendar) {
    return jsonResponse({ok:true});
  }

  // 2. Find the event in the Events sheet
  // Read display text only (one sheet call, not two): Google Sheets stores a time
  // value ("13:00") as an internal Date object anchored to Dec 30, 1899 - and for
  // that historical date, many Eastern European timezones (including
  // Bucharest/Chisinau) use an old Local Mean Time offset instead of the modern
  // one, so .getHours() on that object returns garbage. Reading the text exactly
  // as it's displayed in the sheet sidesteps that entirely, for every column.
  var displayValues = eventsSheet.getDataRange().getDisplayValues();
  var headers = displayValues[0];
  var col = {};
  headers.forEach(function(h, i) { col[h] = i; });

  var rowIndex = -1;
  for (var i = 1; i < displayValues.length; i++) {
    if (displayValues[i][col.title] === data.eventTitle) { rowIndex = i; break; }
  }
  if (rowIndex === -1) {
    return jsonResponse({ok:false, error:'event not found'});
  }

  var row = displayValues[rowIndex];
  var calendar = CalendarApp.getDefaultCalendar();
  var calEvent = null;
  var existingId = row[col.calendar_event_id];

  if (existingId) {
    calEvent = calendar.getEventById(existingId);
    // Guard against an old bug: if the cached event has a corrupted date
    // (e.g. year 1970 from a past bad time parse), don't reuse it -
    // recreate it below instead.
    try {
      if (calEvent && calEvent.getStartTime().getFullYear() < 2020) {
        calEvent.deleteEvent();
        calEvent = null;
      }
    } catch (err) {
      // Event was already deleted manually or is otherwise unavailable - just create a new one.
      calEvent = null;
    }
  }

  // 3. If the shared event doesn't exist in the calendar yet - create it once
  if (!calEvent) {
    var dateVal = new Date(row[col.date]);
    var start = combineDateTime(dateVal, row[col.time_start]);
    var end = combineDateTime(dateVal, row[col.time_end]);
    if (isNaN(start.getTime()) || isNaN(end.getTime()) || end <= start) {
      // Don't blow up the whole request over a bad date/time - report what's
      // wrong instead of an opaque exception from the Calendar API.
      return jsonResponse({
        ok:false,
        error:'bad event date/time',
        debug:{
          date: String(row[col.date]), dateType: typeof row[col.date],
          time_start: row[col.time_start],
          time_end: row[col.time_end],
          start: String(start), end: String(end)
        }
      });
    }
    calEvent = calendar.createEvent(row[col.title], start, end, {
      location: row[col.location],
      description: row[col.description]
    });
    eventsSheet.getRange(rowIndex + 1, col.calendar_event_id + 1).setValue(calEvent.getId());
  }

  // 4. Add the person as a guest to that same shared event
  calEvent.addGuest(data.email);

  // 5. Direct "confirm attendance" link - the same as the "Yes" button in the
  //    email invite, but without having to open email (rst=1 = accept right away).
  var rsvpUrl = buildRsvpUrl(calEvent, calendar);

  return jsonResponse({ok:true, calendarRsvpUrl: rsvpUrl});
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// timeStr - text exactly as shown in the sheet (e.g. "13:00"), taken via
// getDisplayValues(). Deliberately NOT read as a Date object: Google Sheets
// anchors time values to Dec 30, 1899, and for that historical date the
// Bucharest/Chisinau timezone uses an old Local Mean Time offset (+01:44:24)
// instead of the modern one - so .getHours() on that object returns
// nonsense values.
function combineDateTime(dateVal, timeStr) {
  var d = new Date(dateVal.getTime());
  var parts = String(timeStr).split(':');
  d.setHours(parseInt(parts[0], 10), parseInt(parts[1] || '0', 10), 0, 0);
  return d;
}

/**
 * Builds a direct RSVP link ("Yes" from the email invite), using the real
 * event id from the Calendar API - the same one Google puts into the "Yes"
 * button in emails. CalendarApp.getId() returns a different format
 * (iCalUID) that doesn't work for this link, so we look the event up
 * through the Calendar advanced service instead.
 * Requires: Apps Script -> Services (+) -> Calendar API.
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
    // Calendar API advanced service isn't enabled - just skip the RSVP link,
    // the site will fall back to the "Open in calendar" link instead.
  }
  return '';
}

/**
 * ==== AUTOMATIC INSTAGRAM PHOTO RESOLUTION ====
 * Reads Instagram post links (from the image_url column in Events and post_url
 * in Gallery) and stores a direct photo URL in the neighboring column
 * (image_resolved / resolved_url), which the site already reads.
 * Instagram serves link-preview crawlers a photo via the standard og:image
 * meta tag - that's the mechanism used here.
 * Instagram photo links expire after a few days, so this function runs on a
 * timer (createRefreshTrigger) and refreshes them before they go stale.
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
    var current = String(values[i][tgtIdx] || '');
    var resolved = resolveInstagramUrl(src);
    if (resolved) {
      sheet.getRange(i + 1, tgtIdx + 1).setValue(resolved);
    } else if (current && !looksLikeRealPhoto(current)) {
      // A previous run stored something bad (Instagram's generic
      // placeholder icon, or - when blocked - even the post's own page
      // URL echoed back instead of a photo) - clear it instead of leaving
      // a broken-looking image in the gallery.
      sheet.getRange(i + 1, tgtIdx + 1).setValue('');
    }
    // Small pause between requests - firing them back-to-back makes
    // Instagram rate-limit us faster (it starts blocking real photo
    // fetches sooner).
    Utilities.sleep(1500);
  }
}

// Real Instagram post photos are served from a scontent-*.cdninstagram.com
// (or fbcdn.net) host. Anything else - a generic static.cdninstagram.com
// placeholder icon, or even the post's own instagram.com URL echoed back -
// means the request got blocked/rate-limited rather than returning a real
// photo, so it's whitelisted in rather than blacklisting each bad pattern
// Instagram happens to return.
function looksLikeRealPhoto(url) {
  return url.indexOf('scontent') !== -1 &&
    (url.indexOf('cdninstagram.com') !== -1 || url.indexOf('fbcdn.net') !== -1);
}

function resolveInstagramUrl(url) {
  try {
    var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
    if (res.getResponseCode() >= 400) return '';
    var html = res.getContentText();
    var match = html.match(/<meta property="og:image" content="([^"]+)"/);
    if (match && match[1]) {
      var resolved = match[1].replace(/&amp;/g, '&');
      if (!looksLikeRealPhoto(resolved)) return '';
      return resolved;
    }
  } catch (err) {
    // Instagram sometimes blocks automated requests - just skip this row,
    // the old value (if any) stays, we'll try again next time.
  }
  return '';
}

/** Run this ONCE manually (pick it from the function list at the top of the
 *  editor -> Run) to turn on automatic photo refreshing every 6 hours. */
function createRefreshTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'refreshInstagramImages') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('refreshInstagramImages')
    .timeBased()
    .everyHours(6)
    .create();
}

/** Adds an "ODA Tools" menu to the sheet for manually refreshing photos. */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('ODA Tools')
    .addItem('Оновити фото з Instagram зараз', 'refreshInstagramImages')
    .addToUi();
}
