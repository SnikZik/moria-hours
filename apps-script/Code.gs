// Moria work-hours backend. Records button presses into a Google Sheet.
// KEY and SHEET_ID are filled in by build.py; the real values are not committed.
var CONFIG = {
  SHEET_ID: '__SHEET_ID__',
  KEY: '__KEY__',
  EMPLOYEE: 'מוריה',
  MANAGER_EMAIL: '__MANAGER_EMAIL__', // gets an email on every clock in and clock out
  TZ: 'Asia/Jerusalem',
  STALE_HOURS: 16 // an open shift older than this is treated as forgotten
};

var LOG_TAB = 'יומן לחיצות';
var HOURS_TAB = 'שעות';
var MONTHLY_TAB = 'סיכום חודשי';
var LABELS = {
  start: 'התחלתי לעבוד',
  break_start: 'יצאתי להפסקה',
  break_end: 'חזרתי מהפסקה',
  end: 'סיימתי לעבוד',
  auto_close: 'משמרת לא נסגרה'
};
var DAYS = ['שני', 'שלישי', 'רביעי', 'חמישי', 'שישי', 'שבת', 'ראשון'];
// Personal lines per action, injected by build.py from messages.local.json (not committed).
var MESSAGES = /*__MESSAGES__*/ {};

function doGet(e) {
  return respond_(handle_(e.parameter || {}));
}

function doPost(e) {
  var body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) {}
  return respond_(handle_(body));
}

function respond_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function handle_(p) {
  if (p.key !== CONFIG.KEY) return { ok: false, error: 'unauthorized' };
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    if (!ss.getSheetByName(HOURS_TAB)) setup_(ss);
    var now = new Date();
    closeStale_(ss, now);
    if (!p.action || p.action === 'status') return status_(ss, now, true);
    var applied = apply_(ss, p.action, now);
    var result = status_(ss, now, !applied.error);
    if (applied.error) result.error = applied.error;
    else result.message = applied.message;
    return result;
  } catch (err) {
    return { ok: false, error: 'server', detail: String(err) };
  } finally {
    lock.releaseLock();
  }
}

function apply_(ss, action, now) {
  var state = readState_(ss).state;
  var allowed = {
    start: state === 'off',
    break_start: state === 'working',
    break_end: state === 'break',
    end: state === 'working' || state === 'break'
  };
  if (!(action in allowed)) return { error: 'bad_action' };
  if (!allowed[action]) return { error: 'invalid_state' };

  var hours = ss.getSheetByName(HOURS_TAB);
  if (action === 'start') {
    var row = hours.getLastRow() + 1;
    hours.getRange(row, 1, 1, 3).setValues([[now, dayName_(now), now]]);
    hours.getRange(row, 6).setValue(0);
    hours.getRange(row, 7).setFormula(
      '=IF(OR(C' + row + '="",D' + row + '=""),"",ROUND(MOD(D' + row + '-C' + row + ',1)*24-N(F' + row + ')/60,2))'
    );
  } else if (action === 'break_end') {
    closeBreak_(ss, now);
  } else if (action === 'end') {
    if (state === 'break') closeBreak_(ss, now);
    hours.getRange(hours.getLastRow(), 4).setValue(now);
  }
  var message = pickMessage_(ss, action, now);
  appendLog_(ss, action, now, message);
  if (action === 'start' || action === 'end') notifyManager_(ss, action);
  return { message: message };
}

// A line shown this month, on any button, is never shown again that month.
// The log tab is the record. When a button's list runs out it shows nothing rather than repeat.
function pickMessage_(ss, action, now) {
  var log = ss.getSheetByName(LOG_TAB);
  if (!log.getRange(1, 3).getValue()) log.getRange(1, 3).setValue('משפט').setFontWeight('bold');
  var month = dateKey_(now).slice(0, 7);
  var used = {};
  if (log.getLastRow() >= 2) {
    log.getRange(2, 1, log.getLastRow() - 1, 3).getValues().forEach(function (row) {
      if (row[0] instanceof Date && dateKey_(row[0]).slice(0, 7) === month) used[row[2]] = true;
    });
  }
  var open = (MESSAGES[action] || []).filter(function (line) { return !used[line]; });
  return open.length ? open[Math.floor(Math.random() * open.length)] : '';
}

function closeBreak_(ss, now) {
  var hours = ss.getSheetByName(HOURS_TAB);
  var row = hours.getLastRow();
  var from = readState_(ss).since;
  var minutes = Math.round((now - from) / 60000);
  var breaks = hours.getRange(row, 5).getDisplayValue();
  var span = time_(from) + ' עד ' + time_(now);
  hours.getRange(row, 5).setValue(breaks ? breaks + ', ' + span : span);
  hours.getRange(row, 6).setValue(Number(hours.getRange(row, 6).getValue() || 0) + minutes);
}

function closeStale_(ss, now) {
  var current = readState_(ss);
  if (current.state === 'off') return;
  var hours = ss.getSheetByName(HOURS_TAB);
  var row = hours.getLastRow();
  var started = hours.getRange(row, 3).getValue();
  if (!(started instanceof Date)) return;
  if (now - started < CONFIG.STALE_HOURS * 3600000) return;
  hours.getRange(row, 8).setValue('לא נלחץ סיום, צריך להשלים ידנית');
  appendLog_(ss, 'auto_close', now);
}

function readState_(ss) {
  var log = ss.getSheetByName(LOG_TAB);
  var row = log.getLastRow();
  if (row < 2) return { state: 'off', since: null };
  var values = log.getRange(row, 1, 1, 2).getValues()[0];
  var label = values[1];
  var state = 'off';
  if (label === LABELS.start || label === LABELS.break_end) state = 'working';
  if (label === LABELS.break_start) state = 'break';
  return { state: state, since: values[0] };
}

function status_(ss, now, ok) {
  var current = readState_(ss);
  var hours = ss.getSheetByName(HOURS_TAB);
  var row = hours.getLastRow();
  var out = { ok: ok, state: current.state, now: now.getTime(), employee: CONFIG.EMPLOYEE };
  if (row < 2) return out;

  var v = hours.getRange(row, 1, 1, 8).getValues()[0];
  var breakMinutes = Number(v[5] || 0);
  if (current.state !== 'off') {
    out.startedAt = v[2] instanceof Date ? v[2].getTime() : null;
    out.startLabel = hours.getRange(row, 3).getDisplayValue();
    out.breakMinutes = breakMinutes;
    if (current.state === 'break') out.breakStartedAt = current.since.getTime();
  } else if (v[3] instanceof Date && dateKey_(v[3]) === dateKey_(now)) {
    out.last = {
      start: hours.getRange(row, 3).getDisplayValue(),
      end: hours.getRange(row, 4).getDisplayValue(),
      breakMinutes: breakMinutes,
      netMinutes: Math.round((v[3] - v[2]) / 60000) - breakMinutes
    };
  }
  return out;
}

function appendLog_(ss, action, now, message) {
  ss.getSheetByName(LOG_TAB).appendRow([now, LABELS[action], message || '']);
}

function notifyManager_(ss, action) {
  if (!CONFIG.MANAGER_EMAIL) return;
  SpreadsheetApp.flush();
  var hours = ss.getSheetByName(HOURS_TAB);
  var row = hours.getLastRow();
  var cell = function (col) { return hours.getRange(row, col).getDisplayValue(); };
  var isStart = action === 'start';
  var body = 'תאריך ' + cell(1) + '\nהתחלה ' + cell(3);
  if (!isStart) body += '\nסיום ' + cell(4) + '\nהפסקות ' + (cell(6) || '0') + ' דקות\nסה"כ ' + cell(7) + ' שעות';
  try {
    MailApp.sendEmail({
      to: CONFIG.MANAGER_EMAIL,
      subject: CONFIG.EMPLOYEE + (isStart ? ' התחילה לעבוד ' + cell(3) : ' סיימה לעבוד ' + cell(4)),
      body: body
    });
  } catch (err) {
    console.error(err);
    hours.getRange(row, 8).setValue('מייל לא נשלח');
  }
}

function setup_(ss) {
  ss.setSpreadsheetTimeZone(CONFIG.TZ);
  var hours = ss.getSheets()[0].setName(HOURS_TAB);
  hours.getRange(1, 1, 1, 8).setValues([['תאריך', 'יום', 'התחלה', 'סיום', 'הפסקות', 'דקות הפסקה', 'שעות נטו', 'הערה']]);
  hours.getRange('A2:A').setNumberFormat('dd/MM/yyyy');
  hours.getRange('C2:D').setNumberFormat('HH:mm');
  hours.getRange('G2:G').setNumberFormat('0.00');
  hours.setColumnWidth(5, 170);
  hours.setColumnWidth(8, 220);
  styleTab_(hours);

  var log = ss.insertSheet(LOG_TAB);
  log.getRange(1, 1, 1, 3).setValues([['תאריך ושעה', 'פעולה', 'משפט']]);
  log.getRange('A2:A').setNumberFormat('dd/MM/yyyy HH:mm:ss');
  log.setColumnWidth(1, 170);
  log.setColumnWidth(2, 170);
  styleTab_(log);

  var monthly = ss.insertSheet(MONTHLY_TAB);
  monthly.getRange('A1').setFormula(
    "=QUERY('" + HOURS_TAB + "'!A:G,\"select year(A), month(A)+1, count(A), sum(G) where A is not null " +
    "group by year(A), month(A)+1 order by year(A), month(A)+1 " +
    "label year(A) 'שנה', month(A)+1 'חודש', count(A) 'משמרות', sum(G) 'שעות'\",1)"
  );
  styleTab_(monthly);
}

function styleTab_(sheet) {
  sheet.setRightToLeft(true);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, sheet.getMaxColumns()).setFontWeight('bold');
}

function dayName_(d) {
  return DAYS[Number(Utilities.formatDate(d, CONFIG.TZ, 'u')) - 1];
}

function time_(d) {
  return Utilities.formatDate(d, CONFIG.TZ, 'HH:mm');
}

function dateKey_(d) {
  return Utilities.formatDate(d, CONFIG.TZ, 'yyyy-MM-dd');
}
