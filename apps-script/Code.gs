// Finplan API for the website: returns sheet values only with the correct key.
// The key is stored in Project Settings -> Script properties -> API_KEY (not in code).

const SHEET_ID = '13f7MNsJ94DHbw2VJPImBop9y0dGrUljzsJkgv0uRh2E';
const RANGES = {
  'Бюджет': 'A1:I64',
  'Портфель': 'A1:L31',
  'Подушка': 'A1:B16',
  'Цели': 'A1:J25',
  'Прогноз': 'A1:K30',
  'История': 'A1:F64',
  'Бумаги': 'A1:L60'
};

function doPost(e) {
  let key = '', mode = '';
  try { const b = JSON.parse((e && e.postData && e.postData.contents) || '{}'); key = b.key || ''; mode = b.mode || ''; } catch (err) {}
  const expected = PropertiesService.getScriptProperties().getProperty('API_KEY');
  if (!expected || key !== expected) return json_({ ok: false, error: 'unauthorized' });
  if (mode === 'broker') {
    try { return json_({ ok: true, updated: new Date().toISOString(), broker: brokerPortfolio_() }); }
    catch (err) { return json_({ ok: false, error: String(err && err.message || err).indexOf('no_token') >= 0 ? 'no_token' : 'broker_error', detail: String(err && err.message || err).slice(0, 200) }); }
  }
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

// ===================== MOEX quotes =====================
// ISS MOEX without a subscription is delayed by ~15 minutes.

const BOARDS = [
  { market: 'shares', board: 'TQBR', kind: 'share' },
  { market: 'shares', board: 'TQTF', kind: 'fund' },
  { market: 'bonds', board: 'TQOB', kind: 'ofz' },
  { market: 'bonds', board: 'TQCB', kind: 'bond' }
];
const MM_TICKERS = ['LQDT', 'AKMM', 'SBMM', 'TMON', 'CASH', 'AMNY', 'BCSD', 'PSMM', 'GPBM', 'MONY', 'SBMB', 'OPNS', 'YAMM'];
const GOLD_TICKERS = ['GOLD', 'TGLD', 'AKGD', 'SBGD', 'GLDX', 'VTBG'];

function fundType_(secid, name) {
  const n = (name || '').toLowerCase();
  if (MM_TICKERS.indexOf(secid) >= 0 || /денеж|ликвид|деньг|ден\.? ?рын|накопит|money/.test(n)) return 'Фонд денежного рынка';
  if (GOLD_TICKERS.indexOf(secid) >= 0 || /золот|gold/.test(n)) return 'Золото';
  if (/офз|гособл|госдолг/.test(n)) return 'ОФЗ';
  if (/облиг|bond|долг/.test(n)) return 'Корп. облигации';
  return 'Акции / фонды акций';
}

function fetchBoard_(b) {
  const url = 'https://iss.moex.com/iss/engines/stock/markets/' + b.market + '/boards/' + b.board +
    '/securities.json?iss.meta=off&iss.only=securities,marketdata' +
    '&securities.columns=SECID,SHORTNAME,PREVPRICE,FACEVALUE,ACCRUEDINT,FACEUNIT,CURRENCYID' +
    '&marketdata.columns=SECID,LAST,MARKETPRICE,LCURRENTPRICE,YIELD';
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) throw new Error('MOEX ' + b.board + ': HTTP ' + res.getResponseCode());
  const j = JSON.parse(res.getContentText());
  const col = function (block, name) { return j[block].columns.indexOf(name); };
  const md = {};
  j.marketdata.data.forEach(function (r) { md[r[col('marketdata', 'SECID')]] = r; });
  const rows = [];
  j.securities.data.forEach(function (s) {
    const secid = s[col('securities', 'SECID')];
    const name = s[col('securities', 'SHORTNAME')];
    const m = md[secid] || [];
    const pick = function (n) { const i = col('marketdata', n); return i >= 0 ? m[i] : null; };
    const price = pick('LAST') || pick('MARKETPRICE') || pick('LCURRENTPRICE') || s[col('securities', 'PREVPRICE')];
    if (!price) return;
    let unit = price, face = '', aci = '', yld = '', type;
    if (b.market === 'bonds') {
      const fu = s[col('securities', 'FACEUNIT')];
      if (fu && fu !== 'SUR' && fu !== 'RUB') return; // rouble bonds only
      face = s[col('securities', 'FACEVALUE')] || 1000;
      aci = s[col('securities', 'ACCRUEDINT')] || 0;
      unit = price * face / 100 + aci;
      yld = pick('YIELD') ? pick('YIELD') / 100 : '';
      type = b.kind === 'ofz' ? 'ОФЗ' : 'Корп. облигации';
    } else {
      const cur = s[col('securities', 'CURRENCYID')];
      if (cur && cur !== 'SUR' && cur !== 'RUB') return; // rouble securities only
      const isFund = b.kind === 'fund' || /ETF|БПИФ|ПИФ/i.test(name || '') || MM_TICKERS.indexOf(secid) >= 0 || GOLD_TICKERS.indexOf(secid) >= 0;
      type = isFund ? fundType_(secid, name) : 'Акции / фонды акций';
    }
    rows.push([secid, name, type, b.board, price, face, aci, Math.round(unit * 10000) / 10000, yld]);
  });
  return rows;
}

// Refreshes sheet "Котировки": all rouble shares, funds, OFZ and corporate bonds on MOEX.
function updateQuotes() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName('Котировки');
  if (!sh) sh = ss.insertSheet('Котировки');
  let rows = [];
  BOARDS.forEach(function (b) { rows = rows.concat(fetchBoard_(b)); });
  rows.sort(function (a, b) { return a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : 0; });
  const header = [['Тикер', 'Название', 'Тип', 'Режим', 'Цена (₽ или % номинала)', 'Номинал', 'НКД', 'Цена за 1 шт, ₽', 'Доходность к погашению']];
  sh.clearContents();
  sh.getRange(1, 1, 1, header[0].length).setValues(header).setFontWeight('bold');
  if (rows.length) sh.getRange(2, 1, rows.length, header[0].length).setValues(rows);
  sh.getRange('K1').setValue('Обновлено');
  sh.getRange('K2').setValue(new Date()).setNumberFormat('dd.MM.yyyy HH:mm');
  sh.getRange('L1').setValue('Источник: ISS Московской биржи, задержка ~15 мин.');
  sh.getRange('H2:H').setNumberFormat('#,##0.00');
  sh.getRange('I2:I').setNumberFormat('0.00%');
  sh.setFrozenRows(1);
}

// Run once: creates sheet "Бумаги", links it to "Портфель", loads quotes and refreshes them every 30 minutes.
function setupSecurities() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  updateQuotes();
  let sh = ss.getSheetByName('Бумаги');
  if (!sh) sh = ss.insertSheet('Бумаги', ss.getSheetByName('Портфель').getIndex());
  sh.getRange('A1').setValue('Биржевые бумаги').setFontSize(16).setFontWeight('bold');
  sh.getRange('A2').setValue('Выбери тикер из списка, укажи количество и среднюю цену покупки за 1 шт. Цена обновляется сама (Мосбиржа, задержка ~15 мин). Облигации: цена = % номинала × номинал + НКД.')
    .setFontSize(9).setFontColor('#5C6B7A');
  const head = ['Тикер', 'Количество, шт', 'Цена покупки за 1 шт, ₽', 'Ожид. доходность (необяз.)', 'Название', 'Тип', 'Цена сейчас за 1 шт, ₽', 'Стоимость, ₽', 'Вложено, ₽', 'Результат, ₽', 'Доходность для расчёта', 'Доля в бумагах'];
  sh.getRange(4, 1, 1, head.length).setValues([head]).setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#16202B').setWrap(true);
  const N0 = 5, N1 = 54, n = N1 - N0 + 1;
  sh.getRange(N0, 1, n, 4).setBackground('#FFF6CC').setFontColor('#0000FF');
  sh.getRange(N0, 2, n, 1).setNumberFormat('#,##0.##');
  sh.getRange(N0, 3, n, 1).setNumberFormat('#,##0.00 ₽');
  sh.getRange(N0, 4, n, 1).setNumberFormat('0.0%');
  const rule = SpreadsheetApp.newDataValidation().requireValueInRange(ss.getSheetByName('Котировки').getRange('A2:A'), true).setAllowInvalid(false).build();
  sh.getRange(N0, 1, n, 1).setDataValidation(rule);
  const f = [];
  for (let r = N0; r <= N1; r++) {
    f.push([
      '=IF(A' + r + '="","",IFERROR(VLOOKUP(A' + r + ',Котировки!$A:$I,2,FALSE),"нет котировки"))',
      '=IF(A' + r + '="","",IFERROR(VLOOKUP(A' + r + ',Котировки!$A:$I,3,FALSE),"Другое"))',
      '=IF(A' + r + '="","",IFERROR(VLOOKUP(A' + r + ',Котировки!$A:$I,8,FALSE),0))',
      '=IF(A' + r + '="","",B' + r + '*G' + r + ')',
      '=IF(A' + r + '="","",B' + r + '*C' + r + ')',
      '=IF(A' + r + '="","",H' + r + '-I' + r + ')',
      '=IF(A' + r + '="","",IF(D' + r + '<>"",D' + r + ',IFERROR(VLOOKUP(A' + r + ',Котировки!$A:$I,9,FALSE)*1,0)))',
      '=IF(OR(A' + r + '="",$H$55=0),"",H' + r + '/$H$55)'
    ]);
  }
  sh.getRange(N0, 5, n, 8).setFormulas(f);
  sh.getRange(N0, 7, n, 4).setNumberFormat('#,##0 ₽');
  sh.getRange(N0, 11, n, 2).setNumberFormat('0.0%');
  sh.getRange('A55').setValue('Итого').setFontWeight('bold');
  sh.getRange('H55').setFormula('=SUM(H5:H54)').setNumberFormat('#,##0 ₽').setFontWeight('bold');
  sh.getRange('I55').setFormula('=SUM(I5:I54)').setNumberFormat('#,##0 ₽').setFontWeight('bold');
  sh.getRange('J55').setFormula('=H55-I55').setNumberFormat('+#,##0 ₽;-#,##0 ₽;0 ₽').setFontWeight('bold');
  sh.getRange('K55').setFormula('=IF(H55=0,0,SUMPRODUCT(IFERROR(H5:H54*K5:K54,0))/H55)').setNumberFormat('0.0%').setFontWeight('bold');
  sh.getRange('A57').setValue('Котировки обновлены:');
  sh.getRange('C57').setFormula('=Котировки!K2').setNumberFormat('dd.MM.yyyy HH:mm');
  sh.setColumnWidth(1, 90); sh.setColumnWidths(2, 3, 120); sh.setColumnWidth(5, 220); sh.setColumnWidth(6, 170); sh.setColumnWidths(7, 6, 120);
  sh.setFrozenRows(4);

  // "Портфель": totals and structure include securities
  const p = ss.getSheetByName('Портфель');
  p.getRange('A31').setValue('В итогах учтён лист «Бумаги» (биржевые бумаги по текущим ценам).').setFontSize(9).setFontColor('#5C6B7A');
  p.getRange('C30').setFormula('=SUM(C5:C29)+Бумаги!I55');
  p.getRange('D30').setFormula('=SUM(D5:D29)+Бумаги!H55');
  p.getRange('G30').setFormula('=SUM(G5:G29)+SUMPRODUCT(IFERROR(Бумаги!H5:H54*Бумаги!K5:K54,0))');
  for (let r = 5; r <= 14; r++) {
    p.getRange('K' + r).setFormula('=SUMIF($B$5:$B$29,J' + r + ',$D$5:$D$29)+SUMIF(Бумаги!$F$5:$F$54,J' + r + ',Бумаги!$H$5:$H$54)');
  }

  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'updateQuotes') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('updateQuotes').timeBased().everyMinutes(30).create();
  console.log('Done: sheet "Бумаги" created, quotes loaded: ' + (ss.getSheetByName('Котировки').getLastRow() - 1));
}

// ===================== T-Invest API (read-only token in Script properties: TINVEST_TOKEN) =====================
const TINV_BASES = ['https://invest-public-api.tbank.ru/rest/', 'https://invest-public-api.tinkoff.ru/rest/'];

function tCall_(method, body) {
  const token = PropertiesService.getScriptProperties().getProperty('TINVEST_TOKEN');
  if (!token) throw new Error('no_token');
  let last = null;
  for (let k = 0; k < TINV_BASES.length; k++) {
    let res;
    try {
      res = UrlFetchApp.fetch(TINV_BASES[k] + 'tinkoff.public.invest.api.contract.v1.' + method, {
        method: 'post', contentType: 'application/json', payload: JSON.stringify(body || {}),
        headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true
      });
    } catch (err) { last = err; continue; }
    const code = res.getResponseCode();
    if (code === 200) return JSON.parse(res.getContentText());
    last = new Error('tinvest HTTP ' + code + ': ' + res.getContentText().slice(0, 150));
    if (code === 401 || code === 403) throw last;
  }
  throw last;
}

function q_(m) { return m ? Number(m.units || 0) + Number(m.nano || 0) / 1e9 : 0; }

function instrument_(uid) {
  const cache = CacheService.getScriptCache();
  const c = cache.get('i_' + uid);
  if (c) return JSON.parse(c);
  const r = tCall_('InstrumentsService/GetInstrumentBy', { idType: 'INSTRUMENT_ID_TYPE_UID', id: uid });
  const i = r.instrument || {};
  const o = { ticker: i.ticker || '', name: i.name || '' };
  cache.put('i_' + uid, JSON.stringify(o), 21600);
  return o;
}

function brokerType_(it, ticker, name) {
  if (it === 'share') return 'Акции / фонды акций';
  if (it === 'bond') return /^SU\d/.test(ticker) ? 'ОФЗ' : 'Корп. облигации';
  if (it === 'etf') return fundType_(ticker, name);
  if (it === 'currency') return /^RUB/.test(ticker) ? 'Наличные' : 'Валюта';
  return 'Другое';
}

// Positions of all open accounts. Prices are in the instrument currency (field "currency"); the site sums only rouble ones.
function brokerPortfolio_() {
  const acc = tCall_('UsersService/GetAccounts', {});
  const out = [];
  (acc.accounts || []).filter(function (a) { return a.status === 'ACCOUNT_STATUS_OPEN'; }).forEach(function (a) {
    const p = tCall_('OperationsService/GetPortfolio', { accountId: a.id, currency: 'RUB' });
    (p.positions || []).forEach(function (pos) {
      const qty = q_(pos.quantity);
      if (!qty) return;
      const inf = pos.instrumentUid ? instrument_(pos.instrumentUid) : { ticker: pos.figi, name: '' };
      const cur = ((pos.currentPrice || {}).currency || 'rub').toLowerCase();
      const price = q_(pos.currentPrice), nkd = q_(pos.currentNkd), avg = q_(pos.averagePositionPrice);
      out.push({
        account: a.name || 'Счёт', ticker: inf.ticker || pos.figi, name: inf.name || inf.ticker || pos.figi,
        instrumentType: pos.instrumentType, type: brokerType_(pos.instrumentType, inf.ticker, inf.name),
        qty: qty, price: price, nkd: nkd, avg: avg, currency: cur,
        value: qty * (price + nkd), invested: qty * avg, dailyYield: q_(pos.dailyYield)
      });
    });
  });
  return out;
}

// Check from the editor: prints how many positions were found (no token or amounts in the log).
function testBroker() {
  const b = brokerPortfolio_();
  const accs = {};
  b.forEach(function (x) { accs[x.account] = (accs[x.account] || 0) + 1; });
  console.log('Broker OK. Positions by account: ' + JSON.stringify(accs));
}
