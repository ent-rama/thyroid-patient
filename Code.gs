/**
 * ทะเบียนผู้ป่วย Head&Neck - Apps Script Web App
 * วางไฟล์นี้ใน Apps Script ที่สร้างจาก Google Sheet (Extensions > Apps Script)
 * 1 แถว = 1 เหตุการณ์ (ข้อมูลผู้ป่วย / ผ่าตัด / RAI)  ผู้ป่วยคนเดียวกันใช้ HN เดียวกัน
 */

var SHEET_NAME = 'Records';
var USERS_SHEET = 'Users';
var OWNER_EMAIL = 'nayika.to@gmail.com';

var HEADERS = ['ID', 'HN', 'ชื่อ', 'วันเกิด', 'อายุ', 'Tel', 'Dx', 'วันที่ Bx', 'Patho',
  'ประเภท', 'ผ่าตัดครั้งที่', 'ผ่าตัดอะไร', 'วันที่ผ่าตัด', 'วันที่ Discharge',
  'วันที่ RAI', 'ความแรง RAI (mCi)', 'Complication', 'หมายเหตุ', 'ผู้บันทึก', 'เวลาบันทึก'];

// ดัชนีคอลัมน์ (เริ่มที่ 0)
var C = { ID: 0, HN: 1, NAME: 2, BIRTH: 3, AGE: 4, TEL: 5, DX: 6, BX: 7, PATHO: 8,
  TYPE: 9, SXNO: 10, PROC: 11, SXDATE: 12, DC: 13, RAIDATE: 14, RAIDOSE: 15,
  COMP: 16, NOTE: 17, BY: 18, AT: 19 };

var PATIENT_COLS = [C.NAME, C.BIRTH, C.AGE, C.TEL, C.DX, C.BX, C.PATHO];
var DATE_COLS = [C.BIRTH, C.BX, C.SXDATE, C.DC, C.RAIDATE];

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || '';

  // ถ้ามี action = เรียกเป็น API (GET) → คืน JSON
  if (action) {
    return handleApi_(action, e.parameter);
  }

  // ไม่มี action → serve HTML ตามปกติ
  var email = Session.getActiveUser().getEmail();
  var t;
  if (!isAllowed_(email)) {
    t = HtmlService.createHtmlOutput(
      '<meta name="viewport" content="width=device-width,initial-scale=1">' +
      '<body style="font-family:sans-serif;padding:24px;text-align:center">' +
      '<h3>ไม่มีสิทธิ์เข้าใช้งาน</h3><p>บัญชี: ' + (email || '(ไม่ทราบ)') +
      '</p><p>กรุณาติดต่อผู้ดูแลเพื่อเพิ่มอีเมลในชีต Users</p></body>');
  } else {
    t = HtmlService.createHtmlOutputFromFile('index');
  }
  return t.setTitle('ทะเบียนผู้ป่วย Head&Neck')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  var body = {};
  try { body = JSON.parse(e.postData.contents); } catch(x) {}
  var action = body.action || '';
  return handleApi_(action, body);
}

function handleApi_(action, params) {
  var callback = params.callback || '';
  try {
    var result;
    switch (action) {
      case 'init':
        result = apiInit();
        break;
      case 'search':
        result = apiSearch(params.q || '');
        break;
      case 'getPatient':
        result = apiGetPatient(params.hn || '');
        break;
      case 'save':
        var saveData = params.data;
        if (typeof saveData === 'string') { try { saveData = JSON.parse(saveData); } catch(x){} }
        result = apiSave(saveData || params);
        break;
      case 'delete':
        result = apiDelete(params.id || '');
        break;
      default:
        result = { error: 'Unknown action: ' + action };
    }
    return jsonpOrJson_(result, callback);
  } catch (err) {
    return jsonpOrJson_({ error: err.message }, callback);
  }
}

function jsonpOrJson_(data, callback) {
  var json = JSON.stringify(data);
  if (callback) {
    // JSONP: ส่งกลับเป็น JavaScript เพื่อข้าม CORS
    return ContentService.createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------- ชีตและสิทธิ์ ----------
// ใช้ได้ทั้งสคริปต์ที่สร้างจากในชีต และสคริปต์ที่สร้างที่ script.google.com โดยตรง
var SHEET_ID = '1Don2BTnMJgo7jeJqRxHWvYmoIQiZIKmaguGmxbEjJcM';
function ss_() { return SpreadsheetApp.openById(SHEET_ID); }

function recordsSheet_() {
  var ss = ss_();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    // ใช้ชีตแรกเป็น Records (ไฟล์ที่สร้างไว้มีหัวคอลัมน์แล้ว)
    sh = ss.getSheets()[0];
    sh.setName(SHEET_NAME);
  }
  if (sh.getLastRow() === 0 || sh.getRange(1, 1).getValue() !== 'ID') {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  }
  sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#e3f2fd');
  sh.setFrozenRows(1);
  // คอลัมน์ HN, Tel, วันที่ เป็นข้อความ กัน Sheets แปลงค่าเอง
  [C.HN, C.TEL].concat(DATE_COLS).forEach(function (c) {
    sh.getRange(2, c + 1, sh.getMaxRows() - 1, 1).setNumberFormat('@');
  });
  return sh;
}

function usersSheet_() {
  var ss = ss_();
  var sh = ss.getSheetByName(USERS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(USERS_SHEET);
    sh.getRange(1, 1).setValue('อีเมลที่ได้รับอนุญาต').setFontWeight('bold');
    sh.getRange(2, 1, 2, 1).setValues([[OWNER_EMAIL], ['nayika.run@mahidol.ac.th']]);
    sh.setColumnWidth(1, 300);
  }
  return sh;
}

function isAllowed_(email) {
  if (!email) return false;
  email = email.toLowerCase();
  if (email === OWNER_EMAIL.toLowerCase() || email === 'nayika.run@mahidol.ac.th') return true;
  var sh = usersSheet_();
  var n = sh.getLastRow() - 1;
  if (n < 1) return false;
  var list = sh.getRange(2, 1, n, 1).getValues();
  for (var i = 0; i < list.length; i++) {
    if (String(list[i][0]).trim().toLowerCase() === email) return true;
  }
  return false;
}

function requireAuth_() {
  var email = Session.getActiveUser().getEmail();
  if (!isAllowed_(email)) throw new Error('ไม่มีสิทธิ์ใช้งาน (บัญชี: ' + (email || 'ไม่ทราบ — กรุณา login Google ด้วยบัญชีที่ได้รับอนุญาต') + ')');
  return email;
}

// ---------- แปลงข้อมูล ----------
function rowToObj_(r) {
  var o = {};
  Object.keys(C).forEach(function (k) {
    var v = r[C[k]];
    o[k] = (v instanceof Date) ? Utilities.formatDate(v, 'Asia/Bangkok', 'dd/MM/yyyy') : (v === null ? '' : String(v));
  });
  return o;
}

function readAll_() {
  var sh = recordsSheet_();
  var n = sh.getLastRow() - 1;
  if (n < 1) return [];
  var vals = sh.getRange(2, 1, n, HEADERS.length).getValues();
  return vals.map(function (r, i) { var o = rowToObj_(r); o._row = i + 2; return o; });
}

function patientFromRows_(rows) {
  // เอาข้อมูลผู้ป่วยจากแถวล่าสุดที่มีค่า
  var p = { HN: rows[0].HN, NAME: '', BIRTH: '', AGE: '', TEL: '', DX: '', BX: '', PATHO: '' };
  rows.forEach(function (r) {
    ['NAME', 'BIRTH', 'AGE', 'TEL', 'DX', 'BX', 'PATHO'].forEach(function (k) {
      if (r[k]) p[k] = r[k];
    });
  });
  return p;
}

// ---------- API ที่หน้าเว็บเรียก ----------
function apiInit() {
  var email = requireAuth_();
  recordsSheet_();
  return { email: email };
}

function apiSearch(q) {
  requireAuth_();
  q = String(q || '').trim().toLowerCase();
  var all = readAll_();
  var byHn = {};
  var order = [];
  all.forEach(function (r) {
    if (!r.HN) return;
    if (!byHn[r.HN]) { byHn[r.HN] = []; order.push(r.HN); }
    byHn[r.HN].push(r);
  });
  var out = [];
  order.forEach(function (hn) {
    var p = patientFromRows_(byHn[hn]);
    if (!q || hn.toLowerCase().indexOf(q) >= 0 || p.NAME.toLowerCase().indexOf(q) >= 0) {
      p.sx = byHn[hn].filter(function (r) { return r.TYPE === 'SX'; }).length;
      p.rai = byHn[hn].filter(function (r) { return r.TYPE === 'RAI'; }).length;
      out.push(p);
    }
  });
  return out.reverse().slice(0, 100);
}

function apiGetPatient(hn) {
  requireAuth_();
  hn = String(hn).trim();
  var rows = readAll_().filter(function (r) { return r.HN === hn; });
  if (!rows.length) return null;
  return { patient: patientFromRows_(rows), events: rows.filter(function (r) { return r.TYPE !== 'INFO'; }) };
}

/**
 * บันทึกเหตุการณ์ ob: {ID?, HN, NAME, BIRTH, AGE, TEL, DX, BX, PATHO, TYPE, SXNO, PROC,
 *   SXDATE, DC, RAIDATE, RAIDOSE, COMP, NOTE}  วันที่เป็น dd/mm/yyyy (พ.ศ.)
 */
function apiSave(ob) {
  var email = requireAuth_();
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var hn = String(ob.HN || '').trim();
    if (!hn) throw new Error('กรุณากรอก HN');
    if (['INFO', 'SX', 'RAI'].indexOf(ob.TYPE) < 0) throw new Error('ประเภทไม่ถูกต้อง');

    var sh = recordsSheet_();
    var now = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm');
    var row = new Array(HEADERS.length).fill('');
    Object.keys(C).forEach(function (k) { if (ob[k] !== undefined) row[C[k]] = ob[k]; });
    row[C.HN] = hn;
    row[C.BY] = email;
    row[C.AT] = now;

    var targetRow = 0;
    if (ob.ID) {
      var ids = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1).getValues();
      for (var i = 0; i < ids.length; i++) {
        if (String(ids[i][0]) === String(ob.ID)) { targetRow = i + 2; break; }
      }
    }
    var existsHn = false;
    if (sh.getLastRow() > 1) {
      existsHn = sh.getRange(2, C.HN + 1, sh.getLastRow() - 1, 1).getValues()
        .some(function (r) { return String(r[0]) === hn; });
    }
    if (ob.TYPE === 'INFO' && existsHn) {
      row[C.ID] = ''; // แก้ข้อมูลผู้ป่วย: ซิงค์ทุกแถวของ HN นี้ ไม่เพิ่มแถวใหม่
    } else if (targetRow) {
      sh.getRange(targetRow, 1, 1, HEADERS.length).setValues([row]);
    } else {
      row[C.ID] = Utilities.getUuid().slice(0, 8);
      sh.appendRow(row);
      targetRow = sh.getLastRow();
      // ให้ HN/Tel/วันที่ เป็นข้อความ
      [C.HN, C.TEL].concat(DATE_COLS).forEach(function (c) {
        sh.getRange(targetRow, c + 1).setNumberFormat('@').setValue(row[c]);
      });
    }

    // อัปเดตข้อมูลผู้ป่วยในทุกแถวของ HN เดียวกันให้ตรงกัน
    var last = sh.getLastRow();
    var hns = sh.getRange(2, C.HN + 1, last - 1, 1).getValues();
    PATIENT_COLS.forEach(function (c) {
      for (var j = 0; j < hns.length; j++) {
        if (String(hns[j][0]) === hn) sh.getRange(j + 2, c + 1).setNumberFormat('@').setValue(row[c]);
      }
    });
    SpreadsheetApp.flush();
    return { ok: true, id: row[C.ID] };
  } finally {
    lock.releaseLock();
  }
}

function apiDelete(id) {
  requireAuth_();
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = recordsSheet_();
    var n = sh.getLastRow() - 1;
    if (n < 1) return { ok: false };
    var ids = sh.getRange(2, 1, n, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(id)) { sh.deleteRow(i + 2); return { ok: true }; }
    }
    return { ok: false };
  } finally {
    lock.releaseLock();
  }
}
