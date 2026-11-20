// Bind this script to the destination Sheet and deploy it as a web app that executes as its owner.
// The web app must accept anonymous requests; keep participant data pseudonymous and non-sensitive.
var EVENT_HEADERS = [
  'event_id',
  'participant_id',
  'session_id',
  'timestamp',
  'elapsed_ms',
  'condition',
  'event_type',
  'query',
  'selected_option',
  'method',
  'result_title',
  'product_id',
  'product_name',
  'category',
  'sustainable',
  'quantity',
  'event_data'
];

function doPost(e) {
  if (!e || !e.postData || !e.postData.contents) {
    throw new Error('Request body is required.');
  }
  if (e.postData.contents.length > 100000) {
    throw new Error('Request body is too large.');
  }

  var payload = JSON.parse(e.postData.contents);
  var events = payload && payload.events;
  if (!Array.isArray(events) || events.length < 1 || events.length > 50) {
    throw new Error('Send between 1 and 50 events per request.');
  }

  var rows = events.map(function (event) {
    requireText(event, 'event_id', 100);
    requireText(event, 'participant_id', 100);
    requireText(event, 'session_id', 100);
    requireText(event, 'timestamp', 50);
    requireText(event, 'condition', 10);
    requireText(event, 'type', 80);

    return [
      event.event_id,
      event.participant_id,
      event.session_id,
      event.timestamp,
      event.elapsed_ms,
      event.condition,
      event.type,
      event.query,
      event.label,
      event.method,
      event.title,
      event.product_id,
      event.product,
      event.category,
      event.sustainableProduct,
      event.quantity,
      JSON.stringify(event)
    ].map(toSafeCell);
  });

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadsheet) {
      throw new Error('Bind this Apps Script project to the destination spreadsheet.');
    }

    var sheet = spreadsheet.getSheetByName('Events') || spreadsheet.insertSheet('Events');
    var lastRow = sheet.getLastRow();
    if (lastRow === 0) {
      sheet.getRange(1, 1, 1, EVENT_HEADERS.length).setValues([EVENT_HEADERS]);
    } else {
      var firstHeader = sheet.getRange(1, 1).getDisplayValue();
      if (firstHeader !== EVENT_HEADERS[0]) {
        throw new Error('The Events sheet already has content with an unexpected header.');
      }
    }
    lastRow = sheet.getLastRow();
    var existingIds = lastRow > 1
      ? sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues().map(function (row) { return row[0]; })
      : [];
    var knownIds = Object.create(null);
    existingIds.forEach(function (id) { knownIds[id] = true; });
    var freshRows = rows.filter(function (row) {
      if (knownIds[row[0]]) return false;
      knownIds[row[0]] = true;
      return true;
    });
    if (freshRows.length) {
      sheet.getRange(lastRow + 1, 1, freshRows.length, EVENT_HEADERS.length).setValues(freshRows);
    }
    rows = freshRows;
  } finally {
    lock.releaseLock();
  }

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, accepted: rows.length }))
    .setMimeType(ContentService.MimeType.JSON);
}

function requireText(event, key, maxLength) {
  if (!event || typeof event[key] !== 'string' || !event[key] || event[key].length > maxLength) {
    throw new Error('Invalid or missing event field: ' + key);
  }
}

function toSafeCell(value) {
  if (value === null || typeof value === 'undefined') {
    return '';
  }
  var text = typeof value === 'string' ? value : JSON.stringify(value);
  return /^\s*[=+\-@]/.test(text) ? "'" + text : text;
}
