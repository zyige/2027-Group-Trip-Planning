// 2027 Couples Trip — response collector
// Paste this into Extensions > Apps Script on the "2027 Couples Trip — Responses" Google Sheet,
// then deploy as a Web App (see deployment steps Yigga has separately).

function doGet(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');
  var data = sheet.getDataRange().getValues();
  var rows = data.slice(1); // skip header row
  var responses = rows
    .filter(function (r) { return r[1]; }) // skip blank rows
    .map(function (r) {
      return {
        ts: r[0],
        couple: r[1],
        length: r[2],
        vibeTop: r[3],
        vibeOpen: r[4] ? String(r[4]).split(', ').filter(Boolean) : [],
        months: r[5] ? String(r[5]).split(', ').filter(Boolean) : [],
        budget: r[6],
        notes: r[7]
      };
    });
  return ContentService
    .createTextOutput(JSON.stringify({ responses: responses }))
    .setMimeType(ContentService.MimeType.JSON);
}

// Prevents spreadsheet formula injection: a cell value starting with
// =, +, -, or @ gets interpreted by Sheets as a formula, not text.
function sanitizeCell(v) {
  v = String(v || '');
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function doPost(e) {
  var payload = JSON.parse(e.postData.contents);
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');
  sheet.appendRow([
    new Date(),
    sanitizeCell(payload.couple),
    sanitizeCell(payload.length),
    sanitizeCell(payload.vibeTop),
    sanitizeCell((payload.vibeOpen || []).join(', ')),
    sanitizeCell((payload.months || []).join(', ')),
    sanitizeCell(payload.budget),
    sanitizeCell(payload.notes)
  ]);
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}
