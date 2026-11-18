// Финансовый планер — API для сайта.
// Отдаёт значения нужных листов таблицы только при правильном ключе.
// Ключ хранится в «Свойствах скрипта» (Project Settings → Script properties → API_KEY), в коде его нет.

const SHEET_ID = '13f7MNsJ94DHbw2VJPImBop9y0dGrUljzsJkgv0uRh2E';
const RANGES = {
  'Бюджет': 'A1:I64',
  'Портфель': 'A1:L31',
  'Подушка': 'A1:B16',
  'Цели': 'A1:J25',
  'Прогноз': 'A1:K30',
  'История': 'A1:F64'
};

function doPost(e) {
  let key = '';
  try { key = JSON.parse((e && e.postData && e.postData.contents) || '{}').key || ''; } catch (err) {}
  const expected = PropertiesService.getScriptProperties().getProperty('API_KEY');
  if (!expected || key !== expected) return json_({ ok: false, error: 'unauthorized' });
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const tz = ss.getSpreadsheetTimeZone();
  const sheets = {};
  Object.keys(RANGES).forEach(function (name) {
    const sh = ss.getSheetByName(name);
    if (!sh) { sheets[name] = []; return; }
    sheets[name] = sh.getRange(RANGES[name]).getValues().map(function (row) {
      return row.map(function (v) {
        return (v instanceof Date) ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : v;
      });
    });
  });
  return json_({ ok: true, updated: new Date().toISOString(), sheets: sheets });
}

function doGet() {
  return json_({ ok: false, error: 'use POST' });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Run once: creates a secret key, saves it in Script properties and prints it in the execution log.
// Copy the key from the log into the website. Running it again replaces the key.
function setupKey() {
  const key = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  PropertiesService.getScriptProperties().setProperty('API_KEY', key);
  console.log('API_KEY: ' + key);
}
