/**
 * 配送日報 Google Sheets 連携 Web App
 * Script Properties: SPREADSHEET_ID, DEVICE_KEY
 */

var API_VERSION = 1;
var SHEET_NAMES = {
  dashboard: 'ダッシュボード', reports: '日報', deliveries: '納品明細',
  works: '作業明細', notes: 'メモ', logs: '同期ログ'
};
var HEADERS = {};
HEADERS[SHEET_NAMES.reports] = ['report_id', '日付', '氏名', '出発', '帰社', '配送時間(分)', '走行距離(km)', '配送件数', '1件あたり時間(分)', '1件あたり距離(km)', '社内作業件数', '翌日ルート', '特記事項', '端末保存日時', '同期日時'];
HEADERS[SHEET_NAMES.deliveries] = ['delivery_id', 'report_id', '日付', '順番', '開始', '完了', '滞在時間(分)', '納品先', '同期日時'];
HEADERS[SHEET_NAMES.works] = ['work_id', 'report_id', '日付', '順番', '開始', '完了', '作業時間(分)', '複合作業名', '選択項目', '詳細', '同期日時'];
HEADERS[SHEET_NAMES.notes] = ['note_id', '日付', '種類', '本文', '更新日時', '同期日時'];
HEADERS[SHEET_NAMES.logs] = ['request_id', 'action', 'entity_id', 'status', 'processed_at', 'error_code'];

function doPost(e) {
  var lock = LockService.getScriptLock();
  var requestId = '';
  var action = '';
  var targetId = '';
  try {
    var payload = parsePayload_(e);
    requestId = requireId_(payload.request_id, 'request_id');
    action = requireAction_(payload.action);
    validateApiVersion_(payload.api_version);
    validateDeviceKey_(payload.device_key);
    targetId = entityId_(action, payload.entity);

    lock.waitLock(30000);
    var spreadsheet = getSpreadsheet_();
    ensureSheets_(spreadsheet);
    cleanupSampleRows_(spreadsheet);
    var prior = findLog_(spreadsheet, requestId);
    if (prior && prior.status === 'processed') {
      return json_({ ok: true, status: 'processed', processed_at: prior.processedAt, duplicate: true });
    }

    upsertLog_(spreadsheet, requestId, action, targetId, 'pending', new Date(), '');
    if (action === 'upsert_report') upsertReport_(spreadsheet, payload.entity);
    if (action === 'delete_report') deleteReport_(spreadsheet, payload.entity);
    if (action === 'upsert_note') upsertNote_(spreadsheet, payload.entity);
    if (action === 'delete_note') deleteNote_(spreadsheet, payload.entity);

    var completedAt = new Date();
    upsertLog_(spreadsheet, requestId, action, targetId, 'processed', completedAt, '');
    SpreadsheetApp.flush();
    return json_({ ok: true, status: 'processed', processed_at: completedAt.toISOString() });
  } catch (error) {
    var code = safeErrorCode_(error);
    try {
      if (requestId) {
        var errorSpreadsheet = getSpreadsheet_();
        ensureSheets_(errorSpreadsheet);
        upsertLog_(errorSpreadsheet, requestId, action, targetId, 'error', new Date(), code);
      }
    } catch (ignored) {}
    return json_({ ok: false, status: 'error', error_code: code });
  } finally {
    try { lock.releaseLock(); } catch (ignoredRelease) {}
  }
}

/** JSONPによる状態確認。日報本文やキーなどの機密情報は返さない。 */
function doGet(e) {
  var callback = e && e.parameter ? String(e.parameter.callback || '') : '';
  if (!/^[A-Za-z_$][0-9A-Za-z_$\.]{0,100}$/.test(callback)) {
    return json_({ status: 'error', processed_at: '' });
  }
  var result = { status: 'pending', processed_at: '' };
  try {
    var requestId = requireId_(e.parameter.request_id, 'request_id');
    var spreadsheet = getSpreadsheet_();
    ensureSheets_(spreadsheet);
    var found = findLog_(spreadsheet, requestId);
    if (found) {
      result.status = found.status === 'processed' ? 'processed' : (found.status === 'error' ? 'error' : 'pending');
      result.processed_at = found.processedAt;
    }
  } catch (error) {
    result = { status: 'error', processed_at: '' };
  }
  return ContentService.createTextOutput(callback + '(' + JSON.stringify(result) + ');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

/** 初回セットアップ。エディタから1回実行する。 */
function initializeWorkbook() {
  var spreadsheet = getSpreadsheet_();
  ensureSheets_(spreadsheet);
  cleanupSampleRows_(spreadsheet);
  setupDashboard_(spreadsheet);
  SpreadsheetApp.flush();
  return spreadsheet.getUrl();
}

/** テンプレートの視覚確認用SAMPLEデータだけを削除する。 */
function cleanupSampleRows() {
  var spreadsheet = getSpreadsheet_();
  ensureSheets_(spreadsheet);
  cleanupSampleRows_(spreadsheet);
  SpreadsheetApp.flush();
}

function cleanupSampleRows_(spreadsheet) {
  [SHEET_NAMES.reports, SHEET_NAMES.deliveries, SHEET_NAMES.works, SHEET_NAMES.notes, SHEET_NAMES.logs].forEach(function (name) {
    var sheet = spreadsheet.getSheetByName(name);
    for (var row = sheet.getLastRow(); row >= 2; row--) {
      var values = sheet.getRange(row, 1, 1, Math.min(2, sheet.getLastColumn())).getDisplayValues()[0];
      if (values.some(function (value) { return String(value).indexOf('SAMPLE-') === 0; })) sheet.deleteRow(row);
    }
  });
}

function parsePayload_(e) {
  var raw = e && e.postData ? e.postData.contents : '';
  if (!raw || raw.length > 500000) throw appError_('INVALID_PAYLOAD');
  var value;
  try { value = JSON.parse(raw); } catch (error) { throw appError_('INVALID_JSON'); }
  if (!value || Object.prototype.toString.call(value) !== '[object Object]') throw appError_('INVALID_PAYLOAD');
  return value;
}

function validateApiVersion_(value) {
  if (Number(value) !== API_VERSION) throw appError_('UNSUPPORTED_API_VERSION');
}

function validateDeviceKey_(value) {
  var expected = PropertiesService.getScriptProperties().getProperty('DEVICE_KEY');
  if (!expected) throw appError_('SERVER_NOT_CONFIGURED');
  if (typeof value !== 'string' || value.length < 16 || !constantTimeEqual_(value, expected)) throw appError_('UNAUTHORIZED');
}

function constantTimeEqual_(left, right) {
  var max = Math.max(left.length, right.length);
  var diff = left.length ^ right.length;
  for (var i = 0; i < max; i++) diff |= (left.charCodeAt(i % (left.length || 1)) || 0) ^ (right.charCodeAt(i % (right.length || 1)) || 0);
  return diff === 0;
}

function requireAction_(value) {
  var allowed = ['upsert_report', 'delete_report', 'upsert_note', 'delete_note'];
  if (allowed.indexOf(value) < 0) throw appError_('UNKNOWN_ACTION');
  return value;
}

function entityId_(action, entity) {
  if (!entity || Object.prototype.toString.call(entity) !== '[object Object]') throw appError_('INVALID_ENTITY');
  if (action.indexOf('report') >= 0) return requireId_(entity.report_id || entity.date, 'report_id');
  return requireId_(entity.note_id || entity.id, 'note_id');
}

function upsertReport_(spreadsheet, source) {
  var report = validateReport_(source);
  var now = new Date();
  var deliveries = report.deliveries;
  var deliveryCount = deliveries.filter(function (item) { return Boolean(item.place); }).length;
  var workCount = report.works.filter(function (item) { return Boolean(item.label || item.start || item.end); }).length;
  var duration = minutesBetween_(report.departureTime, report.returnTime);
  var row = [
    report.id, report.date, report.userName, report.departureTime, report.returnTime,
    blankable_(duration), blankable_(report.mileage), deliveryCount,
    deliveryCount && duration !== null ? duration / deliveryCount : '',
    deliveryCount && report.mileage !== null ? report.mileage / deliveryCount : '', workCount,
    report.nextRoutes.join(' → '), report.specialNotes, report.savedAt, now
  ];
  upsertRowById_(spreadsheet.getSheetByName(SHEET_NAMES.reports), report.id, row);
  deleteRowsByValue_(spreadsheet.getSheetByName(SHEET_NAMES.deliveries), 2, report.id);
  deleteRowsByValue_(spreadsheet.getSheetByName(SHEET_NAMES.works), 2, report.id);

  var deliveryRows = [];
  deliveries.forEach(function (item, index) {
    if (!item.start && !item.end && !item.place) return;
    deliveryRows.push([
      item.id || report.id + '-delivery-' + (index + 1), report.id, report.date, index + 1,
      item.start, item.end, blankable_(minutesBetween_(item.start, item.end)), item.place, now
    ]);
  });
  appendRows_(spreadsheet.getSheetByName(SHEET_NAMES.deliveries), deliveryRows);

  var workRows = [];
  report.works.forEach(function (work, index) {
    if (!work.start && !work.end && !work.label) return;
    workRows.push([
      work.id || report.id + '-work-' + (index + 1), report.id, report.date, index + 1,
      work.start, work.end, blankable_(minutesBetween_(work.start, work.end)), work.label,
      work.items.join('＋'), work.detail, now
    ]);
  });
  appendRows_(spreadsheet.getSheetByName(SHEET_NAMES.works), workRows);
}

function deleteReport_(spreadsheet, source) {
  var id = requireId_(source.report_id || source.date, 'report_id');
  deleteRowsByValue_(spreadsheet.getSheetByName(SHEET_NAMES.reports), 1, id);
  deleteRowsByValue_(spreadsheet.getSheetByName(SHEET_NAMES.deliveries), 2, id);
  deleteRowsByValue_(spreadsheet.getSheetByName(SHEET_NAMES.works), 2, id);
}

function upsertNote_(spreadsheet, source) {
  var note = validateNote_(source);
  upsertRowById_(spreadsheet.getSheetByName(SHEET_NAMES.notes), note.id,
    [note.id, note.date, note.kind, note.body, note.updatedAt, new Date()]);
}

function deleteNote_(spreadsheet, source) {
  deleteRowsByValue_(spreadsheet.getSheetByName(SHEET_NAMES.notes), 1, requireId_(source.note_id || source.id, 'note_id'));
}

function validateReport_(source) {
  var date = requireDate_(source.date);
  var id = requireId_(source.report_id || date, 'report_id');
  var deliveries = requireArray_(source.deliveries, 500, 'deliveries').map(function (item) {
    item = requireObject_(item);
    var location = text_(item.location, 200);
    var other = text_(item.other, 200);
    return { id: optionalId_(item.delivery_id || item.id), start: time_(item.start), end: time_(item.end), place: location === 'その他' && other ? other : location };
  });
  var works = requireArray_(source.internalWorks || source.works, 500, 'works').map(function (item) {
    item = requireObject_(item);
    var items = stringArray_(item.items, 30, 100);
    var labels = items.map(function (name) { return detailedWorkLabel_(name, item); }).filter(Boolean);
    return {
      id: optionalId_(item.work_id || item.id), start: time_(item.start), end: time_(item.end),
      items: items, label: labels.join('＋'), detail: workDetail_(item)
    };
  });
  var mileage = source.mileage === '' || source.mileage === null || typeof source.mileage === 'undefined' ? null : Number(source.mileage);
  if (mileage !== null && (!isFinite(mileage) || mileage < 0 || mileage > 1000000)) throw appError_('INVALID_MILEAGE');
  return {
    id: id, date: date, userName: text_(source.userName || source.user_name, 100),
    departureTime: time_(source.departureTime || source.departure_time), returnTime: time_(source.returnTime || source.return_time),
    mileage: mileage, specialNotes: text_(source.specialNotes || source.special_notes, 10000),
    savedAt: isoDateTime_(source.savedAt || source.saved_at), deliveries: deliveries, works: works,
    nextRoutes: requireArray_(source.nextRoutes || source.next_routes, 100, 'nextRoutes').map(function (route) {
      route = requireObject_(route); var location = text_(route.location, 200); var other = text_(route.other, 200);
      return location === 'その他' && other ? other : location;
    }).filter(Boolean)
  };
}

function validateNote_(source) {
  return {
    id: requireId_(source.note_id || source.id, 'note_id'), date: requireDate_(source.date),
    kind: text_(source.kind, 100), body: requiredText_(source.body, 20000, 'body'),
    updatedAt: isoDateTime_(source.updatedAt || source.updated_at)
  };
}

function detailedWorkLabel_(name, item) {
  var map = { '片付け': ['cleanupItems', 'cleanupOther'], '検査': ['inspectionItems', 'inspectionOther'], '荷受け': ['receivingItems', 'receivingOther'] };
  if (name === 'その他') return text_(item.other, 500) || 'その他';
  if (!map[name]) return text_(name, 200);
  var values = stringArray_(item[map[name][0]], 30, 100).map(function (value) {
    return value === 'その他' && item[map[name][1]] ? 'その他（' + text_(item[map[name][1]], 200) + '）' : value;
  });
  return values.length ? name + '（' + values.join('、') + '）' : name;
}

function workDetail_(item) {
  var details = [];
  ['other', 'cleanupOther', 'inspectionOther', 'receivingOther'].forEach(function (key) {
    var value = text_(item[key], 500); if (value) details.push(key + ': ' + value);
  });
  return details.join(' / ');
}

function getSpreadsheet_() {
  var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw appError_('SERVER_NOT_CONFIGURED');
  try { return SpreadsheetApp.openById(id); } catch (error) { throw appError_('SPREADSHEET_UNAVAILABLE'); }
}

function ensureSheets_(spreadsheet) {
  [SHEET_NAMES.dashboard, SHEET_NAMES.reports, SHEET_NAMES.deliveries, SHEET_NAMES.works, SHEET_NAMES.notes, SHEET_NAMES.logs].forEach(function (name) {
    var sheet = spreadsheet.getSheetByName(name) || spreadsheet.insertSheet(name);
    if (HEADERS[name]) ensureHeader_(sheet, HEADERS[name]);
  });
}

function ensureHeader_(sheet, headers) {
  if (sheet.getMaxColumns() < headers.length) sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
  var current = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  var empty = current.every(function (value) { return value === ''; });
  if (!empty && current.join('\u001f') !== headers.join('\u001f')) throw appError_('HEADER_MISMATCH_' + sheet.getName());
  if (empty) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#d9ead3');
    sheet.setFrozenRows(1);
  }
}

function setupDashboard_(spreadsheet) {
  var sheet = spreadsheet.getSheetByName(SHEET_NAMES.dashboard);
  sheet.clear();
  sheet.getCharts().forEach(function (chart) { sheet.removeChart(chart); });
  sheet.getRange('A1').setValue('配送・社内業務 効率ダッシュボード').setFontSize(18).setFontWeight('bold');
  sheet.getRange('A2').setValue('対象月');
  sheet.getRange('B2').setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM'));
  sheet.getRange('A4:G4').setValues([['配送日数', '総配送時間(分)', '総走行距離(km)', '総配送件数', '1日平均時間(分)', '1件平均時間(分)', '1件平均距離(km)']]).setFontWeight('bold');
  sheet.getRange('A5').setFormula('=IFERROR(COUNT(FILTER(日報!F2:F,LEFT(日報!B2:B,7)=$B$2)),0)');
  sheet.getRange('B5').setFormula('=IFERROR(SUM(FILTER(日報!F2:F,LEFT(日報!B2:B,7)=$B$2)),0)');
  sheet.getRange('C5').setFormula('=IFERROR(SUM(FILTER(日報!G2:G,LEFT(日報!B2:B,7)=$B$2)),0)');
  sheet.getRange('D5').setFormula('=IFERROR(SUM(FILTER(日報!H2:H,LEFT(日報!B2:B,7)=$B$2)),0)');
  sheet.getRange('E5').setFormula('=IFERROR(B5/A5,"")');
  sheet.getRange('F5').setFormula('=IFERROR(B5/D5,"")');
  sheet.getRange('G5').setFormula('=IFERROR(C5/D5,"")');
  sheet.getRange('A8:F8').setValues([['日付', '配送時間(分)', '走行距離(km)', '配送件数', '1件時間(分)', '1件距離(km)']]).setFontWeight('bold');
  sheet.getRange('A9').setFormula('=IFERROR(FILTER({日報!B2:B,日報!F2:F,日報!G2:G,日報!H2:H,日報!I2:I,日報!J2:J},LEFT(日報!B2:B,7)=$B$2),"")');
  sheet.getRange('H8:K8').setValues([['日付', '走行距離(km)', '配送時間(分)', '配送件数']]).setFontWeight('bold');
  sheet.getRange('H9').setFormula('=IFERROR(FILTER({日報!B2:B,日報!G2:G,日報!F2:F,日報!H2:H},LEFT(日報!B2:B,7)=$B$2,日報!F2:F<>"",日報!G2:G<>"",日報!H2:H>0),"")');
  sheet.getRange('M8:R8').setValues([['複合作業名', '実施回数', '平均時間(分)', '合計時間(分)', '最短(分)', '最長(分)']]).setFontWeight('bold');
  sheet.getRange('M9').setFormula('=IFERROR(QUERY(FILTER({作業明細!H2:H,作業明細!G2:G},LEFT(作業明細!C2:C,7)=$B$2,作業明細!H2:H<>"",作業明細!G2:G<>""),"select Col1,count(Col2),avg(Col2),sum(Col2),min(Col2),max(Col2) group by Col1 order by sum(Col2) desc label Col1 \'\',count(Col2) \'\',avg(Col2) \'\',sum(Col2) \'\',min(Col2) \'\',max(Col2) \'\'",0),"")');
  sheet.getRange('S8:U8').setValues([['中央値(分)', '前月平均差(分)', '判定']]).setFontWeight('bold');
  sheet.getRange('S9').setFormula('=MAP(M9:M,LAMBDA(n,IF(n="","",IFERROR(MEDIAN(FILTER(作業明細!G$2:G,作業明細!H$2:H=n,LEFT(作業明細!C$2:C,7)=$B$2)),""))))');
  sheet.getRange('T9').setFormula('=MAP(M9:M,O9:O,LAMBDA(n,a,IF(n="","",IFERROR(a-AVERAGE(FILTER(作業明細!G$2:G,作業明細!H$2:H=n,LEFT(作業明細!C$2:C,7)=TEXT(EDATE(DATEVALUE($B$2&"-01"),-1),"yyyy-mm"))),""))))');
  sheet.getRange('U9').setFormula('=MAP(M9:M,N9:N,LAMBDA(n,c,IF(n="","",IF(c<3,"データ不足","参考値"))))');
  sheet.getRange('W8').setValue('合計時間上位3作業の週別平均').setFontWeight('bold');
  sheet.getRange('W9').setFormula('=IFERROR(QUERY(FILTER({作業明細!H2:H,TEXT(DATEVALUE(作業明細!C2:C)-WEEKDAY(DATEVALUE(作業明細!C2:C),2)+1,"yyyy-mm-dd"),作業明細!G2:G},LEFT(作業明細!C2:C,7)=$B$2,ISNUMBER(MATCH(作業明細!H2:H,$M$9:$M$11,0))),"select Col2,avg(Col3) group by Col2 pivot Col1 label Col2 \'週開始\'",0),"")');
  sheet.setFrozenRows(2);
  sheet.autoResizeColumns(1, 21);
  var deliveryChart = sheet.newChart().setChartType(Charts.ChartType.BUBBLE).addRange(sheet.getRange('H8:K2000')).setPosition(28, 1, 0, 0).setOption('title', '配送効率（距離 × 時間 × 件数）').setOption('hAxis.title', '走行距離(km)').setOption('vAxis.title', '配送時間(分)').build();
  var workChart = sheet.newChart().setChartType(Charts.ChartType.BUBBLE).addRange(sheet.getRange('M8:P500')).setPosition(28, 9, 0, 0).setOption('title', '改善候補マトリクス').setOption('hAxis.title', '実施回数').setOption('vAxis.title', '平均時間(分)').build();
  var weeklyChart = sheet.newChart().setChartType(Charts.ChartType.LINE).addRange(sheet.getRange('W9:AA100')).setPosition(47, 1, 0, 0).setOption('title', '上位作業の週別平均時間').setOption('vAxis.title', '平均時間(分)').build();
  sheet.insertChart(deliveryChart);
  sheet.insertChart(workChart);
  sheet.insertChart(weeklyChart);
}

function findLog_(spreadsheet, requestId) {
  var sheet = spreadsheet.getSheetByName(SHEET_NAMES.logs);
  if (sheet.getLastRow() < 2) return null;
  var finder = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).createTextFinder(requestId).matchEntireCell(true).findNext();
  if (!finder) return null;
  var row = sheet.getRange(finder.getRow(), 1, 1, 6).getValues()[0];
  return { row: finder.getRow(), status: String(row[3]), processedAt: isoOutput_(row[4]), errorCode: String(row[5] || '') };
}

function upsertLog_(spreadsheet, requestId, action, targetId, status, time, errorCode) {
  var sheet = spreadsheet.getSheetByName(SHEET_NAMES.logs);
  var existing = findLog_(spreadsheet, requestId);
  var row = [requestId, action || '', targetId || '', status, time, errorCode || ''];
  if (existing) sheet.getRange(existing.row, 1, 1, row.length).setValues([row]);
  else sheet.appendRow(row);
}

function upsertRowById_(sheet, id, row) {
  var found = sheet.getLastRow() >= 2 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).createTextFinder(id).matchEntireCell(true).findNext() : null;
  if (found) sheet.getRange(found.getRow(), 1, 1, row.length).setValues([row]);
  else sheet.getRange(sheet.getLastRow() + 1, 1, 1, row.length).setValues([row]);
}

function deleteRowsByValue_(sheet, column, value) {
  for (var row = sheet.getLastRow(); row >= 2; row--) {
    if (String(sheet.getRange(row, column).getValue()) === String(value)) sheet.deleteRow(row);
  }
}

function appendRows_(sheet, rows) {
  if (!rows.length) return;
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}

function minutesBetween_(start, end) {
  if (!start || !end) return null;
  var a = start.split(':').map(Number); var b = end.split(':').map(Number);
  var from = a[0] * 60 + a[1]; var to = b[0] * 60 + b[1];
  if (to < from) to += 1440;
  return to - from;
}

function time_(value) {
  if (value === null || typeof value === 'undefined' || value === '') return '';
  value = String(value);
  if (!/^(?:[01]\d|2[0-3]):[0-5]\d$/.test(value)) throw appError_('INVALID_TIME');
  return value;
}

function requireDate_(value) {
  value = String(value || '');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(value)) throw appError_('INVALID_DATE');
  var parsed = new Date(value + 'T00:00:00Z');
  if (isNaN(parsed.getTime()) || parsed.toISOString().slice(0, 10) !== value) throw appError_('INVALID_DATE');
  return value;
}

function isoDateTime_(value) {
  if (!value) return '';
  var parsed = new Date(value);
  if (isNaN(parsed.getTime())) throw appError_('INVALID_DATETIME');
  return parsed;
}

function requireId_(value, field) {
  value = String(value || '');
  if (!/^[0-9A-Za-z._:-]{1,160}$/.test(value)) throw appError_('INVALID_' + String(field).toUpperCase());
  return value;
}

function optionalId_(value) { return value ? requireId_(value, 'id') : ''; }
function requireObject_(value) { if (!value || Object.prototype.toString.call(value) !== '[object Object]') throw appError_('INVALID_ENTITY'); return value; }
function requireArray_(value, max, field) { if (typeof value === 'undefined' || value === null) return []; if (!Array.isArray(value) || value.length > max) throw appError_('INVALID_' + field.toUpperCase()); return value; }
function stringArray_(value, maxItems, maxLength) { return requireArray_(value, maxItems, 'array').map(function (item) { return text_(item, maxLength); }).filter(Boolean); }
function requiredText_(value, max, field) { var result = text_(value, max); if (!result) throw appError_('INVALID_' + field.toUpperCase()); return result; }
function text_(value, max) { if (value === null || typeof value === 'undefined') return ''; value = String(value).trim(); if (value.length > max) throw appError_('TEXT_TOO_LONG'); return /^[=+\-@]/.test(value) ? "'" + value : value; }
function blankable_(value) { return value === null || typeof value === 'undefined' || (typeof value === 'number' && isNaN(value)) ? '' : value; }
function isoOutput_(value) { if (!value) return ''; if (Object.prototype.toString.call(value) === '[object Date]') return value.toISOString(); return String(value); }
function appError_(code) { var error = new Error(code); error.code = code; return error; }
function safeErrorCode_(error) { var code = error && error.code ? String(error.code) : 'INTERNAL_ERROR'; return /^[A-Z0-9_]{1,100}$/.test(code) ? code : 'INTERNAL_ERROR'; }
function json_(value) { return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON); }
