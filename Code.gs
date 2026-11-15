/**
 * ทะเบียนผู้ป่วย Head&Neck - Apps Script Web App
 * โครงสร้างข้อมูลแยกเป็นสองชีต (ไม่ซ้ำซ้อน):
 *   Patients : 1 แถว = 1 ผู้ป่วย  (คีย์หลัก = HN, ห้ามซ้ำ)
 *   Events   : 1 แถว = 1 การผ่าตัดหรือ RAI  (อ้างถึงผู้ป่วยด้วย HN เท่านั้น)
 *   Records  : ข้อมูลเดิม เก็บไว้เป็นสำรอง ไม่ถูกใช้แล้ว
 */

var SHEET_ID = '1d3WMP3nm4Wb2ouP7FafaPQNzceTh8TvM3M31MI_diPA';
var PATIENTS_SHEET = 'Patients';
var EVENTS_SHEET = 'Events';
var OLD_SHEET = 'Records';
var USERS_SHEET = 'Users';
var OWNER_EMAIL = 'nayika.to@gmail.com';

var PAT_HEADERS = ['HN', 'ชื่อ', 'วันเกิด', 'อายุ', 'Tel', 'Dx', 'วันที่ Bx', 'Patho', 'ผู้บันทึก', 'เวลาบันทึก'];
var PAT = { HN: 0, NAME: 1, BIRTH: 2, AGE: 3, TEL: 4, DX: 5, BX: 6, PATHO: 7, BY: 8, AT: 9 };

var EVT_HEADERS = ['ID', 'HN', 'ประเภท', 'ผ่าตัดครั้งที่', 'ผ่าตัดอะไร', 'วันที่ผ่าตัด', 'วันที่ Discharge',
  'วันที่ RAI', 'ความแรง RAI (mCi)', 'Complication', 'หมายเหตุ', 'ผู้บันทึก', 'เวลาบันทึก'];
var EVT = { ID: 0, HN: 1, TYPE: 2, SXNO: 3, PROC: 4, SXDATE: 5, DC: 6, RAIDATE: 7, RAIDOSE: 8,
  COMP: 9, NOTE: 10, BY: 11, AT: 12 };

// ดัชนีคอลัมน์ของชีต Records เดิม (ใช้ตอนย้ายข้อมูลเท่านั้น)
var OLD = { ID: 0, HN: 1, NAME: 2, BIRTH: 3, AGE: 4, TEL: 5, DX: 6, BX: 7, PATHO: 8,
  TYPE: 9, SXNO: 10, PROC: 11, SXDATE: 12, DC: 13, RAIDATE: 14, RAIDOSE: 15,
  COMP: 16, NOTE: 17, BY: 18, AT: 19 };

// ---------- หน้าเว็บ / ทางเข้า API ----------
function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || '';
  if (action) {
    return handleApi_(action, e.parameter);
  }
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
  try { body = JSON.parse(e.postData.contents); } catch (x) {}
  return handleApi_(body.action || '', body);
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
        var data = params.data;
        if (typeof data === 'string') { try { data = JSON.parse(data); } catch (x) {} }
        result = apiSave(data || params);
        break;
      case 'delete':
        result = apiDelete(params.id || '');
        break;
      case 'migrate':
        result = migrateAsOwner_();
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
    return ContentService.createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

// ---------- ชีตและสิทธิ์ ----------
function ss_() { return SpreadsheetApp.openById(SHEET_ID); }

function getSheet_(name, headers, textCols) {
  var ss = ss_();
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0 || sh.getRange(1, 1).getValue() !== headers[0]) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#e3f2fd');
  sh.setFrozenRows(1);
  // คอลัมน์ HN, Tel, วันที่ เป็นข้อความ กัน Sheets แปลงค่าเอง
  textCols.forEach(function (c) {
    sh.getRange(2, c + 1, sh.getMaxRows() - 1, 1).setNumberFormat('@');
  });
  return sh;
}

function patientsSheet_() { return getSheet_(PATIENTS_SHEET, PAT_HEADERS, [PAT.HN, PAT.BIRTH, PAT.TEL, PAT.BX]); }
function eventsSheet_() { return getSheet_(EVENTS_SHEET, EVT_HEADERS, [EVT.HN, EVT.SXDATE, EVT.DC, EVT.RAIDATE]); }

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

// ---------- ตัวช่วยแปลงข้อมูล ----------
function str_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Bangkok', 'dd/MM/yyyy');
  return v == null ? '' : String(v);
}
function now_() { return Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm'); }

function readRows_(sh, width) {
  var n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, width).getValues();
}

function patientObj_(r) {
  return {
    HN: str_(r[PAT.HN]).trim(), NAME: str_(r[PAT.NAME]), BIRTH: str_(r[PAT.BIRTH]), AGE: str_(r[PAT.AGE]),
    TEL: str_(r[PAT.TEL]), DX: str_(r[PAT.DX]), BX: str_(r[PAT.BX]), PATHO: str_(r[PAT.PATHO])
  };
}

function eventObj_(r) {
  return {
    ID: str_(r[EVT.ID]), HN: str_(r[EVT.HN]).trim(), TYPE: str_(r[EVT.TYPE]), SXNO: str_(r[EVT.SXNO]),
    PROC: str_(r[EVT.PROC]), SXDATE: str_(r[EVT.SXDATE]), DC: str_(r[EVT.DC]),
    RAIDATE: str_(r[EVT.RAIDATE]), RAIDOSE: str_(r[EVT.RAIDOSE]), COMP: str_(r[EVT.COMP]), NOTE: str_(r[EVT.NOTE])
  };
}

// คืน {row, obj} ของผู้ป่วยตาม HN หรือ null
function findPatient_(hn) {
  var rows = readRows_(patientsSheet_(), PAT_HEADERS.length);
  for (var i = 0; i < rows.length; i++) {
    if (str_(rows[i][PAT.HN]).trim() === hn) return { row: i + 2, obj: patientObj_(rows[i]) };
  }
  return null;
}

// คืนหมายเลขแถวของเหตุการณ์ตาม ID หรือ 0
function findEventRow_(sh, id) {
  var n = sh.getLastRow() - 1;
  if (n < 1) return 0;
  var ids = sh.getRange(2, 1, n, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (str_(ids[i][0]) === String(id)) return i + 2;
  }
  return 0;
}

// ---------- API ที่หน้าเว็บเรียก ----------
function apiInit() {
  var email = requireAuth_();
  patientsSheet_();
  eventsSheet_();
  return { email: email };
}

function apiSearch(q) {
  requireAuth_();
  q = String(q || '').trim().toLowerCase();

  // นับจำนวนผ่าตัด/RAI ต่อ HN จากชีต Events
  var counts = {};
  readRows_(eventsSheet_(), EVT_HEADERS.length).forEach(function (r) {
    var hn = str_(r[EVT.HN]).trim();
    if (!hn) return;
    counts[hn] = counts[hn] || { sx: 0, rai: 0 };
    var t = str_(r[EVT.TYPE]);
    if (t === 'SX') counts[hn].sx++;
    else if (t === 'RAI') counts[hn].rai++;
  });

  var out = readRows_(patientsSheet_(), PAT_HEADERS.length).map(patientObj_).filter(function (p) {
    return p.HN && (!q || p.HN.toLowerCase().indexOf(q) >= 0 || p.NAME.toLowerCase().indexOf(q) >= 0);
  }).map(function (p) {
    var c = counts[p.HN] || { sx: 0, rai: 0 };
    p.sx = c.sx;
    p.rai = c.rai;
    return p;
  });
  return out.reverse().slice(0, 100);
}

function apiGetPatient(hn) {
  requireAuth_();
  hn = String(hn || '').trim();
  var found = findPatient_(hn);
  if (!found) return null;
  var events = readRows_(eventsSheet_(), EVT_HEADERS.length).map(eventObj_).filter(function (e) {
    return e.HN === hn;
  });
  return { patient: found.obj, events: events };
}

/**
 * ob.TYPE === 'INFO' → บันทึกข้อมูลผู้ป่วย (ob.isNew = true เมื่อเพิ่มใหม่)
 * ob.TYPE === 'SX' | 'RAI' → บันทึกเหตุการณ์ (ob.ID มีค่าเมื่อแก้ไข)
 */
function apiSave(ob) {
  var email = requireAuth_();
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (ob.TYPE === 'INFO') return savePatient_(ob, email);
    return saveEvent_(ob, email);
  } finally {
    lock.releaseLock();
  }
}

function savePatient_(ob, email) {
  var hn = String(ob.HN || '').trim();
  if (!hn) throw new Error('กรุณากรอก HN');
  if (!String(ob.NAME || '').trim()) throw new Error('กรุณากรอกชื่อ-สกุล');

  var found = findPatient_(hn);
  if (ob.isNew && found) {
    throw new Error('HN ' + hn + ' มีอยู่แล้ว (' + found.obj.NAME + ') — กรุณาค้นหาแล้วแก้ไขแทน');
  }
  if (!ob.isNew && !found) throw new Error('ไม่พบ HN ' + hn);

  var row = [hn, ob.NAME || '', ob.BIRTH || '', ob.AGE || '', ob.TEL || '', ob.DX || '',
    ob.BX || '', ob.PATHO || '', email, now_()];
  var sh = patientsSheet_();
  if (found) {
    sh.getRange(found.row, 1, 1, row.length).setValues([row]);
  } else {
    sh.getRange(sh.getLastRow() + 1, 1, 1, row.length).setValues([row]);
  }
  SpreadsheetApp.flush();
  return { ok: true, HN: hn };
}

function saveEvent_(ob, email) {
  var hn = String(ob.HN || '').trim();
  if (!hn) throw new Error('กรุณากรอก HN');
  if (['SX', 'RAI'].indexOf(ob.TYPE) < 0) throw new Error('ประเภทไม่ถูกต้อง');
  if (!findPatient_(hn)) throw new Error('ไม่พบ HN ' + hn + ' ในระบบ — ต้องเพิ่มผู้ป่วยก่อน');

  var row = new Array(EVT_HEADERS.length).fill('');
  row[EVT.HN] = hn;
  row[EVT.TYPE] = ob.TYPE;
  row[EVT.SXNO] = ob.SXNO || '';
  row[EVT.PROC] = ob.PROC || '';
  row[EVT.SXDATE] = ob.SXDATE || '';
  row[EVT.DC] = ob.DC || '';
  row[EVT.RAIDATE] = ob.RAIDATE || '';
  row[EVT.RAIDOSE] = ob.RAIDOSE || '';
  row[EVT.COMP] = ob.COMP || '';
  row[EVT.NOTE] = ob.NOTE || '';
  row[EVT.BY] = email;
  row[EVT.AT] = now_();

  var sh = eventsSheet_();
  var target = ob.ID ? findEventRow_(sh, ob.ID) : 0;
  if (target) {
    row[EVT.ID] = String(ob.ID);
    sh.getRange(target, 1, 1, row.length).setValues([row]);
  } else {
    row[EVT.ID] = Utilities.getUuid().slice(0, 8);
    sh.getRange(sh.getLastRow() + 1, 1, 1, row.length).setValues([row]);
  }
  SpreadsheetApp.flush();
  return { ok: true, id: row[EVT.ID] };
}

function apiDelete(id) {
  requireAuth_();
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = eventsSheet_();
    var r = findEventRow_(sh, id);
    if (!r) return { ok: false };
    sh.deleteRow(r);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

// ---------- ย้ายข้อมูลจาก Records เดิมไปเป็นสองชีต (รันครั้งเดียว) ----------
function migrateAsOwner_() {
  var email = (Session.getActiveUser().getEmail() || '').toLowerCase();
  if (email !== OWNER_EMAIL.toLowerCase()) throw new Error('เฉพาะเจ้าของระบบเท่านั้น');
  return migrateRecordsToSplit_();
}

function migrateRecordsToSplit_() {
  var old = ss_().getSheetByName(OLD_SHEET);
  if (!old) throw new Error('ไม่พบชีต Records');
  var pat = patientsSheet_();
  var evt = eventsSheet_();
  if (pat.getLastRow() > 1 || evt.getLastRow() > 1) {
    throw new Error('ชีต Patients/Events มีข้อมูลแล้ว ยกเลิกการย้าย');
  }
  var n = old.getLastRow() - 1;
  if (n < 1) return 'ไม่มีข้อมูลให้ย้าย';

  var rows = old.getRange(2, 1, n, 20).getValues();
  var patients = {};
  var order = [];
  var events = [];

  rows.forEach(function (r) {
    var hn = str_(r[OLD.HN]).trim();
    if (!hn) return;
    if (!patients[hn]) { patients[hn] = {}; order.push(hn); }
    // ใช้ค่าที่ไม่ว่างจากแถวล่าสุด เหมือนลอจิกเดิม
    ['NAME', 'BIRTH', 'AGE', 'TEL', 'DX', 'BX', 'PATHO'].forEach(function (k) {
      var v = str_(r[OLD[k]]);
      if (v) patients[hn][k] = v;
    });
    var t = str_(r[OLD.TYPE]);
    if (t === 'SX' || t === 'RAI') {
      var row = new Array(EVT_HEADERS.length).fill('');
      row[EVT.ID] = str_(r[OLD.ID]) || Utilities.getUuid().slice(0, 8);
      row[EVT.HN] = hn;
      row[EVT.TYPE] = t;
      row[EVT.SXNO] = str_(r[OLD.SXNO]);
      row[EVT.PROC] = str_(r[OLD.PROC]);
      row[EVT.SXDATE] = str_(r[OLD.SXDATE]);
      row[EVT.DC] = str_(r[OLD.DC]);
      row[EVT.RAIDATE] = str_(r[OLD.RAIDATE]);
      row[EVT.RAIDOSE] = str_(r[OLD.RAIDOSE]);
      row[EVT.COMP] = str_(r[OLD.COMP]);
      row[EVT.NOTE] = str_(r[OLD.NOTE]);
      row[EVT.BY] = str_(r[OLD.BY]);
      row[EVT.AT] = str_(r[OLD.AT]);
      events.push(row);
    }
  });

  var prow = order.map(function (hn) {
    var p = patients[hn];
    return [hn, p.NAME || '', p.BIRTH || '', p.AGE || '', p.TEL || '', p.DX || '', p.BX || '', p.PATHO || '', 'migration', now_()];
  });
  if (prow.length) pat.getRange(2, 1, prow.length, PAT_HEADERS.length).setValues(prow);
  if (events.length) evt.getRange(2, 1, events.length, EVT_HEADERS.length).setValues(events);
  SpreadsheetApp.flush();
  return 'ย้ายแล้ว: ผู้ป่วย ' + prow.length + ' ราย, เหตุการณ์ ' + events.length + ' รายการ (ชีต Records ยังอยู่เป็นสำรอง)';
}
