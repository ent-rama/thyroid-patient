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

var PAT_HEADERS = ['HN', 'ชื่อ', 'เพศ', 'วันเกิด', 'อายุ', 'Tel', 'Dx', 'วันที่ Bx', 'Patho', 'Patho เพิ่มเติม', 'เสียชีวิต', 'หมายเหตุเสียชีวิต', 'ผู้บันทึก', 'เวลาบันทึก'];
var PAT = { HN: 0, NAME: 1, GENDER: 2, BIRTH: 3, AGE: 4, TEL: 5, DX: 6, BX: 7, PATHO: 8, PATHONOTE: 9, DEAD: 10, DEADNOTE: 11, BY: 12, AT: 13 };

var EVT_HEADERS = ['ID', 'HN', 'ประเภท', 'ผ่าตัดครั้งที่', 'ผ่าตัดอะไร', 'วันที่รับไว้ รพ.', 'วันที่ผ่าตัด', 'วันที่ Discharge',
  'วันที่ RAI', 'ความแรง RAI (mCi)', 'Complication', 'หมายเหตุ', 'ผู้บันทึก', 'เวลาบันทึก'];
var EVT = { ID: 0, HN: 1, TYPE: 2, SXNO: 3, PROC: 4, ADMIT: 5, SXDATE: 6, DC: 7, RAIDATE: 8, RAIDOSE: 9,
  COMP: 10, NOTE: 11, BY: 12, AT: 13 };

// ดัชนีคอลัมน์ของชีต Records เดิม (ใช้ตอนย้ายข้อมูลเท่านั้น)
var OLD = { ID: 0, HN: 1, NAME: 2, BIRTH: 3, AGE: 4, TEL: 5, DX: 6, BX: 7, PATHO: 8,
  TYPE: 9, SXNO: 10, PROC: 11, SXDATE: 12, DC: 13, RAIDATE: 14, RAIDOSE: 15,
  COMP: 16, NOTE: 17, BY: 18, AT: 19 };

// ---------- หน้าเว็บ ----------
function doGet(e) {
  var email = Session.getActiveUser().getEmail();
  var t;
  if (!isAllowed_(email)) {
    var hint;
    if (email) {
      hint = 'กรุณาติดต่อผู้ดูแลเพื่อเพิ่มอีเมลนี้ในชีต Users';
    } else {
      var eff = '';
      try { eff = Session.getEffectiveUser().getEmail(); } catch (err) { eff = ''; }
      hint = eff
        ? '<b>ระบบมองไม่เห็นว่าใครกำลังเข้าใช้</b><br>สาเหตุ: การ Deploy ตั้งเป็น "ดำเนินการในชื่อ: ฉัน"<br>' +
          'ผู้ดูแลต้องแก้เป็น <b>"ดำเนินการในชื่อ: ผู้ใช้ที่เข้าถึงเว็บแอป"</b> แล้ว Deploy เวอร์ชันใหม่'
        : 'กรุณาลงชื่อเข้าใช้ Google ด้วยบัญชีที่ได้รับอนุญาต แล้วเปิดลิงก์นี้อีกครั้ง';
    }
    t = HtmlService.createHtmlOutput(
      '<meta name="viewport" content="width=device-width,initial-scale=1">' +
      '<body style="font-family:sans-serif;padding:24px;text-align:center;line-height:1.7">' +
      '<h3>ไม่มีสิทธิ์เข้าใช้งาน</h3><p>บัญชี: ' +
      (email ? String(email).replace(/[<>&]/g, '') : '(ไม่ทราบ)') +
      '</p><p>' + hint + '</p></body>');
  } else {
    t = HtmlService.createHtmlOutputFromFile('index');
  }
  return t.setTitle('ทะเบียนผู้ป่วย Head&Neck')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ---------- ชีตและสิทธิ์ ----------
function ss_() { return SpreadsheetApp.openById(SHEET_ID); }

function getSheet_(name, headers, textCols) {
  var ss = ss_();
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#e3f2fd');
  sh.setFrozenRows(1);
  // คอลัมน์ HN, Tel, วันที่ เป็นข้อความ กัน Sheets แปลงค่าเอง
  textCols.forEach(function (c) {
    sh.getRange(2, c + 1, sh.getMaxRows() - 1, 1).setNumberFormat('@');
  });
  return sh;
}

function patientsSheet_() { return getSheet_(PATIENTS_SHEET, PAT_HEADERS, [PAT.HN, PAT.BIRTH, PAT.TEL, PAT.BX]); }
function eventsSheet_() { return getSheet_(EVENTS_SHEET, EVT_HEADERS, [EVT.HN, EVT.ADMIT, EVT.SXDATE, EVT.DC, EVT.RAIDATE]); }

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
    HN: str_(r[PAT.HN]).trim(), NAME: str_(r[PAT.NAME]), GENDER: str_(r[PAT.GENDER]), BIRTH: str_(r[PAT.BIRTH]), AGE: str_(r[PAT.AGE]),
    TEL: str_(r[PAT.TEL]), DX: str_(r[PAT.DX]), BX: str_(r[PAT.BX]), PATHO: str_(r[PAT.PATHO]),
    PATHONOTE: str_(r[PAT.PATHONOTE]),
    DEAD: str_(r[PAT.DEAD]), DEADNOTE: str_(r[PAT.DEADNOTE])
  };
}

function eventObj_(r) {
  return {
    ID: str_(r[EVT.ID]), HN: str_(r[EVT.HN]).trim(), TYPE: str_(r[EVT.TYPE]), SXNO: str_(r[EVT.SXNO]),
    PROC: str_(r[EVT.PROC]), ADMIT: str_(r[EVT.ADMIT]), SXDATE: str_(r[EVT.SXDATE]), DC: str_(r[EVT.DC]),
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

  var row = [hn, ob.NAME || '', ob.GENDER || '', ob.BIRTH || '', ob.AGE || '', ob.TEL || '', ob.DX || '',
    ob.BX || '', ob.PATHO || '', ob.PATHONOTE || '', ob.DEAD ? 'เสียชีวิต' : '', ob.DEADNOTE || '', email, now_()];
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
  row[EVT.ADMIT] = ob.ADMIT || '';
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
    return [hn, p.NAME || '', '', p.BIRTH || '', p.AGE || '', p.TEL || '', p.DX || '', p.BX || '', p.PATHO || '', '', '', '', 'migration', now_()];
  });
  if (prow.length) pat.getRange(2, 1, prow.length, PAT_HEADERS.length).setValues(prow);
  if (events.length) evt.getRange(2, 1, events.length, EVT_HEADERS.length).setValues(events);
  SpreadsheetApp.flush();
  return 'ย้ายแล้ว: ผู้ป่วย ' + prow.length + ' ราย, เหตุการณ์ ' + events.length + ' รายการ (ชีต Records ยังอยู่เป็นสำรอง)';
}

// ================= สรุปข้อมูล / Dashboard =================

// แปลงข้อความวันที่ พ.ศ. (วว/ดด/ปปปป) เป็นเวลามิลลิวินาที UTC
function parseThaiDate_(v) {
  if (v instanceof Date) return Date.UTC(v.getFullYear(), v.getMonth(), v.getDate());
  var s = str_(v).trim();
  if (!s) return null;
  var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!m) return null;
  var d = +m[1], mo = +m[2], y = +m[3];
  if (y >= 2400) y -= 543;              // พ.ศ. -> ค.ศ.
  if (mo < 1 || mo > 12 || d < 1 || d > 31) return null;
  return Date.UTC(y, mo - 1, d);
}

// จำนวนวันจาก a ถึง b (null = คำนวณไม่ได้)
function diffDays_(a, b) {
  var x = parseThaiDate_(a), y = parseThaiDate_(b);
  if (x === null || y === null) return null;
  return Math.round((y - x) / 86400000);
}

// สถิติพื้นฐานของชุดตัวเลข
function stat_(arr) {
  var n = arr.length;
  if (!n) return { n: 0, mean: '', median: '', min: '', max: '' };
  var s = arr.slice().sort(function (a, b) { return a - b; });
  var sum = 0;
  for (var i = 0; i < n; i++) sum += s[i];
  var mid = n >> 1;
  var med = (n % 2) ? s[mid] : (s[mid - 1] + s[mid]) / 2;
  return {
    n: n,
    mean: Math.round(sum / n * 10) / 10,
    median: Math.round(med * 10) / 10,
    min: s[0], max: s[n - 1]
  };
}

function bump_(m, k) { m[k] = (m[k] || 0) + 1; }

function tally_(m) {
  var out = [];
  for (var k in m) if (m.hasOwnProperty(k)) out.push({ k: k, v: m[k] });
  out.sort(function (a, b) { return b.v - a.v || (a.k < b.k ? -1 : 1); });
  return out;
}

// อายุ: ใช้ช่องอายุก่อน ถ้าไม่มีคำนวณจากวันเกิด
function ageOf_(p) {
  var a = parseInt(str_(p.AGE), 10);
  if (!isNaN(a) && a >= 0 && a < 130) return a;
  var b = parseThaiDate_(p.BIRTH);
  if (b === null) return null;
  var bd = new Date(b), now = new Date();
  var y = now.getFullYear() - bd.getUTCFullYear();
  if ((now.getMonth() + 1) * 100 + now.getDate() < (bd.getUTCMonth() + 1) * 100 + bd.getUTCDate()) y--;
  return (y >= 0 && y < 130) ? y : null;
}

function ageBand_(a) {
  if (a === null) return 'ไม่ระบุ';
  if (a < 40) return 'น้อยกว่า 40 ปี';
  if (a < 60) return '40-59 ปี';
  return '60 ปีขึ้นไป';
}

var DAY_BUCKETS = ['0-1 วัน', '2 วัน', '3 วัน', '4 วัน', '5 วัน', '6 วัน', '7 วันขึ้นไป'];

function bucketLabel_(d) {
  if (d <= 1) return DAY_BUCKETS[0];
  if (d >= 7) return DAY_BUCKETS[6];
  return DAY_BUCKETS[d - 1];
}

function pct1_(n, d) { return d ? Math.round(n / d * 1000) / 10 : ''; }

// จำนวนและร้อยละในแต่ละช่วงวัน จากรายการจำนวนวัน
function distOf_(vals) {
  var m = {};
  DAY_BUCKETS.forEach(function (b) { m[b] = 0; });
  vals.forEach(function (d) { m[bucketLabel_(d)]++; });
  return DAY_BUCKETS.map(function (b) { return { k: b, v: m[b], pct: pct1_(m[b], vals.length) }; });
}

// สรุปของกลุ่มหนึ่ง (ปีเดียว หรือทุกปี): surg = วันหลังผ่าตัดรายครั้ง, pat = {HN: วันที่นานที่สุด}
// จำนวนคนในแต่ละช่วงวัน: 1 คนที่ผ่าตัดหลายครั้งนับอยู่ในทุกช่วงที่เขามีการผ่าตัด
function distPatientsAny_(g) {
  var c = {};
  DAY_BUCKETS.forEach(function (b) { c[b] = 0; });
  var hns = Object.keys(g.days);
  hns.forEach(function (hn) {
    var seen = {};
    g.days[hn].forEach(function (d) {
      var b = bucketLabel_(d);
      if (!seen[b]) { seen[b] = 1; c[b]++; }
    });
  });
  return DAY_BUCKETS.map(function (b) { return { k: b, v: c[b], pct: pct1_(c[b], hns.length) }; });
}

function groupSummary_(label, g, limit) {
  var patVals = Object.keys(g.pat).map(function (k) { return g.pat[k]; });
  var sp = 0, pp = 0;
  g.surg.forEach(function (d) { if (d <= limit) sp++; });
  patVals.forEach(function (d) { if (d <= limit) pp++; });
  return {
    year: label,
    patients: Object.keys(g.patSet).length,
    surgeries: g.total,
    kpiSurg: { pass: sp, total: g.surg.length, pct: pct1_(sp, g.surg.length) },
    kpiPat: { pass: pp, total: patVals.length, pct: pct1_(pp, patVals.length) },
    distSurg: distOf_(g.surg),
    distPat: distPatientsAny_(g),
    patientsKnown: Object.keys(g.days).length,
    multi: Object.keys(g.cnt).filter(function (k) { return g.cnt[k] > 1; }).length,
    unknownSurg: g.total - g.surg.length,
    unknownPat: Object.keys(g.patSet).length - patVals.length
  };
}

function newGroup_() { return { surg: [], pat: {}, patSet: {}, total: 0, days: {}, cnt: {} }; }

function addToGroup_(g, e, field) {
  field = field || '_sxDc';
  var d = e[field];
  g.total++;
  g.patSet[e.HN] = 1;
  g.cnt[e.HN] = (g.cnt[e.HN] || 0) + 1;
  if (d !== null && d !== undefined) {
    g.surg.push(d);
    (g.days[e.HN] = g.days[e.HN] || []).push(d);
    if (g.pat[e.HN] === undefined || d > g.pat[e.HN]) g.pat[e.HN] = d;
  }
}

var RAI_LIMIT = 180;
var RAI_BUCKETS = [
  { k: '0-30 วัน', wk: '0-4 สัปดาห์', max: 30 },
  { k: '31-60 วัน', wk: '4-9 สัปดาห์', max: 60 },
  { k: '61-90 วัน', wk: '9-13 สัปดาห์', max: 90 },
  { k: '91-120 วัน', wk: '13-17 สัปดาห์', max: 120 },
  { k: '121-150 วัน', wk: '17-21 สัปดาห์', max: 150 },
  { k: '151-180 วัน', wk: '21-26 สัปดาห์', max: 180 },
  { k: 'มากกว่า 180 วัน', wk: 'มากกว่า 26 สัปดาห์', max: Infinity }
];

function raiDistOf_(vals) {
  var c = RAI_BUCKETS.map(function () { return 0; });
  vals.forEach(function (d) {
    for (var i = 0; i < RAI_BUCKETS.length; i++) {
      if (d <= RAI_BUCKETS[i].max) { c[i]++; break; }
    }
  });
  return RAI_BUCKETS.map(function (b, i) {
    return { k: b.k, wk: b.wk, v: c[i], pct: pct1_(c[i], vals.length) };
  });
}

function weeksStat_(st) {
  if (!st.n) return { n: 0, mean: '', median: '', min: '', max: '' };
  function w(x) { return Math.round(x / 7 * 10) / 10; }
  return { n: st.n, mean: w(st.mean), median: w(st.median), min: w(st.min), max: w(st.max) };
}

function raiGroupSummary_(label, g) {
  var patVals = Object.keys(g.pat).map(function (k) { return g.pat[k]; });
  var sp = 0, pp = 0;
  g.surg.forEach(function (d) { if (d <= RAI_LIMIT) sp++; });
  patVals.forEach(function (d) { if (d <= RAI_LIMIT) pp++; });
  var st = stat_(g.surg);
  return {
    year: label,
    patients: Object.keys(g.patSet).length,
    rai: g.total,
    kpiRai: { pass: sp, total: g.surg.length, pct: pct1_(sp, g.surg.length) },
    kpiPat: { pass: pp, total: patVals.length, pct: pct1_(pp, patVals.length) },
    days: st,
    weeks: weeksStat_(st),
    distRai: raiDistOf_(g.surg),
    distPat: raiDistOf_(patVals),
    unknownRai: g.total - g.surg.length,
    unknownPat: Object.keys(g.patSet).length - patVals.length
  };
}

var PATHO_KEYS = ['PTC', 'FTC', 'ATC', 'MTC', 'Bethesda II', 'Bethesda III'];

function normPatho_(v) {
  v = str_(v).trim();
  if (!v) return 'ไม่ระบุ';
  return PATHO_KEYS.indexOf(v) >= 0 ? v : 'อื่นๆ';
}

function splitItems_(v) {
  return str_(v).split(',').map(function (x) { return x.trim(); }).filter(function (x) { return x; });
}

function hasNeck_(items) {
  return items.some(function (x) { return /neck dissection|MRND/i.test(x); });
}

function hasTotal_(items) {
  // Total thyroidectomy รวมถึง Completion thyroidectomy และที่พิมพ์ย่อว่า TTD
  return items.some(function (x) { return /total thyroidectomy|completion|\bTTD\b/i.test(x); });
}

function hasTrach_(items) {
  return items.some(function (x) { return /tracheostomy/i.test(x); });
}

function hasLobec_(items) {
  return items.some(function (x) { return /lobectomy/i.test(x); });
}

var SX_CATS = [
  'Lobectomy อย่างเดียว',
  'TTD อย่างเดียว',
  'TTD + ต่อมน้ำเหลือง',
  'TTD + ต่อมน้ำเหลือง + tracheostomy',
  'Lobectomy + ต่อมน้ำเหลือง',
  'Lobectomy + tracheostomy + ต่อมน้ำเหลือง',
  'Lobectomy + tracheostomy',
  'TTD + tracheostomy',
  'อื่นๆ'
];
var SX_CATS_ALWAYS = 6;   // 6 กลุ่มแรกแสดงเสมอ ที่เหลือแสดงเมื่อมีข้อมูล

// จัดการผ่าตัด 1 ครั้งเข้ากลุ่มเดียว (ไม่ซ้อนกัน) จากรายการหัตถการ
function surgCategory_(items) {
  var total = hasTotal_(items), lob = hasLobec_(items), neck = hasNeck_(items), tr = hasTrach_(items);
  if (total) {
    if (neck && tr) return SX_CATS[3];
    if (neck) return SX_CATS[2];
    if (tr) return SX_CATS[7];
    return SX_CATS[1];
  }
  if (lob) {
    if (neck && tr) return SX_CATS[5];
    if (neck) return SX_CATS[4];
    if (tr) return SX_CATS[6];
    return SX_CATS[0];
  }
  return SX_CATS[8];
}

function dayStat_(arr) {
  var st = stat_(arr);
  return { n: st.n, mean: st.mean, median: st.median, min: st.min, max: st.max };
}

// สรุปชุด complication: events ที่มี COMP เป็นรายการคั่นด้วย ,  denom = จำนวนคนทั้งหมดของกลุ่มนั้น
function compSummary_(events, denom) {
  var byItem = {}, patAny = {};
  events.forEach(function (e) {
    var seen = {};
    splitItems_(e.COMP).forEach(function (it) {
      if (!byItem[it]) byItem[it] = { k: it, pats: {}, surgeries: 0 };
      byItem[it].pats[e.HN] = 1;
      if (!seen[it]) { seen[it] = 1; byItem[it].surgeries++; }
      patAny[e.HN] = 1;
    });
  });
  var items = Object.keys(byItem).map(function (k) {
    var n = Object.keys(byItem[k].pats).length;
    return { k: k, patients: n, times: byItem[k].surgeries, pct: pct1_(n, denom) };
  }).sort(function (a, b) { return b.patients - a.patients || (a.k < b.k ? -1 : 1); });
  var withAny = Object.keys(patAny).length;
  return {
    denom: denom, items: items,
    withAny: withAny, withAnyPct: pct1_(withAny, denom),
    none: denom - withAny, nonePct: pct1_(denom - withAny, denom)
  };
}

// สรุปมะเร็งไทรอยด์: เพศต่อปี, Patho, การผ่าตัด, วันนอนแยกชนิด, complication, RAI, เสียชีวิต
function buildCancerSummary_(pats, evByHn) {
  var N = pats.length;
  // --- 1) ต่อปี แยกเพศ (ปี Bx ถ้าไม่มีใช้ปีของการรักษาครั้งแรก) ---
  var ym = {};
  pats.forEach(function (p) {
    var t = parseThaiDate_(p.BX);
    if (t === null) {
      (evByHn[p.HN] || []).forEach(function (e) {
        if (e._t !== null && e._t !== undefined && (t === null || e._t < t)) t = e._t;
      });
    }
    var y = (t === null) ? 'ไม่ระบุปี' : String(new Date(t).getUTCFullYear() + 543);
    if (!ym[y]) ym[y] = { year: y, total: 0, male: 0, female: 0, other: 0 };
    ym[y].total++;
    var g = str_(p.GENDER).trim();
    if (g === 'ชาย') ym[y].male++; else if (g === 'หญิง') ym[y].female++; else ym[y].other++;
  });
  var yearRows = Object.keys(ym).sort(function (a, b) {
    if (a === 'ไม่ระบุปี') return 1;
    if (b === 'ไม่ระบุปี') return -1;
    return (+b) - (+a);
  }).map(function (k) { return ym[k]; });
  var tot = { year: 'รวม', total: N, male: 0, female: 0, other: 0 };
  yearRows.forEach(function (r) { tot.male += r.male; tot.female += r.female; tot.other += r.other; });

  // --- 2) Patho ---
  var pm = {};
  pats.forEach(function (p) { bump_(pm, normPatho_(p.PATHO)); });
  var pathoOrder = PATHO_KEYS.concat(['อื่นๆ', 'ไม่ระบุ']);
  var patho = pathoOrder.filter(function (k) {
    return pm[k] || ['PTC', 'FTC', 'ATC', 'MTC'].indexOf(k) >= 0;
  }).map(function (k) { return { k: k, v: pm[k] || 0, pct: pct1_(pm[k] || 0, N) }; });

  // --- 3) การผ่าตัด ---
  var sxPatients = [], sxAll = [], patSx = {};
  pats.forEach(function (p) {
    var list = (evByHn[p.HN] || []).filter(function (e) { return e.TYPE === 'SX'; });
    if (list.length) { sxPatients.push(p); patSx[p.HN] = list; list.forEach(function (e) { sxAll.push(e); }); }
  });
  var cnt = { '1 ครั้ง': 0, '2 ครั้ง': 0, '3 ครั้ง': 0, '4 ครั้งขึ้นไป': 0 };
  var patterns = {};
  sxPatients.forEach(function (p) {
    var list = patSx[p.HN], n = list.length;
    cnt[n >= 4 ? '4 ครั้งขึ้นไป' : n + ' ครั้ง']++;
    if (n >= 2) {
      var pat = list.map(function (e) { return surgCategory_(splitItems_(e.PROC)); }).join(' → ');
      var key = n + '|' + pat;
      if (!patterns[key]) patterns[key] = { n: n, pattern: pat, patients: 0 };
      patterns[key].patients++;
    }
  });
  var countRows = Object.keys(cnt).map(function (k) { return { k: k, v: cnt[k], pct: pct1_(cnt[k], sxPatients.length) }; });
  var patternRows = Object.keys(patterns).map(function (k) { return patterns[k]; })
    .sort(function (a, b) { return a.n - b.n || b.patients - a.patients || (a.pattern < b.pattern ? -1 : 1); });

  // --- 4) วันนอนหลังผ่าตัด แยกตามกลุ่มการผ่าตัด (ไม่ซ้อนกัน) ---
  var allDays = [];
  var cat = {};
  SX_CATS.forEach(function (c) { cat[c] = { surgeries: 0, pats: {}, days: [] }; });
  var ttd = { surgeries: 0, pats: {}, days: [] };
  sxAll.forEach(function (e) {
    var d = e._sxDc, known = (d !== null && d !== undefined);
    var items = splitItems_(e.PROC);
    if (known) allDays.push(d);
    var c = cat[surgCategory_(items)];
    c.surgeries++; c.pats[e.HN] = 1;
    if (known) c.days.push(d);
    if (hasTotal_(items)) {
      ttd.surgeries++; ttd.pats[e.HN] = 1;
      if (known) ttd.days.push(d);
    }
  });
  function grpOut(label, g) {
    return { k: label, surgeries: g.surgeries, patients: Object.keys(g.pats).length, days: dayStat_(g.days) };
  }
  var catRows = SX_CATS.map(function (c, i) { return { i: i, g: cat[c], k: c }; })
    .filter(function (x) { return x.i < SX_CATS_ALWAYS || x.g.surgeries > 0; })
    .map(function (x) { return grpOut(x.k, x.g); });
  var ttdOut = grpOut('TTD ทั้งหมด', ttd);
  ttdOut.pctOfPatients = pct1_(ttdOut.patients, sxPatients.length);
  ttdOut.pctOfSurgeries = pct1_(ttdOut.surgeries, sxAll.length);

  // --- 5) complication หลังผ่าตัด / 6) RAI / 7) complication หลัง RAI ---
  var raiEvents = [], raiPats = {};
  pats.forEach(function (p) {
    (evByHn[p.HN] || []).forEach(function (e) {
      if (e.TYPE === 'RAI') { raiEvents.push(e); raiPats[p.HN] = 1; }
    });
  });
  var nRaiPat = Object.keys(raiPats).length;

  // --- 8) เสียชีวิต ---
  var deadBy = {}, patBy = {}, nDead = 0;
  pats.forEach(function (p) {
    var k = normPatho_(p.PATHO);
    bump_(patBy, k);
    if (str_(p.DEAD).trim()) { nDead++; bump_(deadBy, k); }
  });
  var deadRows = pathoOrder.filter(function (k) { return deadBy[k]; })
    .map(function (k) { return { k: k, v: deadBy[k], total: patBy[k] || 0, pctOfPatho: pct1_(deadBy[k], patBy[k] || 0) }; });

  return {
    patients: N,
    yearGender: yearRows, yearGenderTotal: tot,
    patho: patho,
    surgery: {
      patients: sxPatients.length, times: sxAll.length,
      byCount: countRows, patterns: patternRows
    },
    los: {
      patients: sxPatients.length, surgeries: sxAll.length,
      all: dayStat_(allDays),
      byCat: catRows,
      ttd: ttdOut
    },
    compSx: compSummary_(sxAll, sxPatients.length),
    rai: { patients: nRaiPat, pct: pct1_(nRaiPat, N), comp: compSummary_(raiEvents, nRaiPat) },
    dead: { total: nDead, pct: pct1_(nDead, N), byPatho: deadRows }
  };
}

// คำนวณสถิติทั้งหมดจากชีต Patients + Events
function computeStats_() {
  var pats = readRows_(patientsSheet_(), PAT_HEADERS.length).map(patientObj_)
    .filter(function (p) { return p.HN; });
  var evts = readRows_(eventsSheet_(), EVT_HEADERS.length).map(eventObj_)
    .filter(function (e) { return e.HN; });

  var byHn = {};
  pats.forEach(function (p) { byHn[p.HN] = p; });

  // ----- กลุ่มนับหัวคน -----
  var gender = {}, band = {}, patho = {}, ages = [];
  pats.forEach(function (p) {
    bump_(gender, str_(p.GENDER).trim() || 'ไม่ระบุ');
    var a = ageOf_(p);
    if (a !== null) ages.push(a);
    bump_(band, ageBand_(a));
    bump_(patho, str_(p.PATHO).trim() || 'ไม่ระบุ');
  });

  // ----- จัดเหตุการณ์ตาม HN เรียงตามวันที่ -----
  var evByHn = {};
  evts.forEach(function (e) {
    e._t = parseThaiDate_(e.TYPE === 'RAI' ? e.RAIDATE : e.SXDATE);
    (evByHn[e.HN] = evByHn[e.HN] || []).push(e);
  });
  Object.keys(evByHn).forEach(function (hn) {
    evByHn[hn].sort(function (a, b) { return (a._t === null ? 0 : a._t) - (b._t === null ? 0 : b._t); });
  });

  // ----- กลุ่มนับรายครั้ง -----
  var proc = {}, doseTally = {};
  var losEach = [], sxDcEach = [], sxRaiEach = [], doses = [];
  var sxRows = [], raiRows = [];
  var nSx = 0, nRai = 0, badDate = 0;

  Object.keys(evByHn).forEach(function (hn) {
    var lastSx = null;
    evByHn[hn].forEach(function (e) {
      if (e.TYPE === 'SX') {
        nSx++; lastSx = e;
        e._raiDate = ''; e._sxRaiDays = null;
        str_(e.PROC).split(',').forEach(function (x) { x = x.trim(); if (x) bump_(proc, x); });
        var d1 = diffDays_(e.ADMIT, e.DC);
        var d2 = diffDays_(e.SXDATE, e.DC);
        if (d1 !== null && d1 < 0) { badDate++; d1 = null; }
        if (d2 !== null && d2 < 0) { badDate++; d2 = null; }
        e._los = d1; e._sxDc = d2;
        if (d1 !== null) losEach.push(d1);
        if (d2 !== null) sxDcEach.push(d2);
        sxRows.push(e);
      } else if (e.TYPE === 'RAI') {
        nRai++;
        var dz = parseFloat(str_(e.RAIDOSE));
        if (!isNaN(dz)) { doses.push(dz); bump_(doseTally, str_(e.RAIDOSE).trim() + ' mCi'); }
        var d3 = lastSx ? diffDays_(lastSx.SXDATE, e.RAIDATE) : null;
        if (d3 !== null && d3 < 0) { badDate++; d3 = null; }
        e._sxRai = d3;
        e._pairSxNo = lastSx ? str_(lastSx.SXNO) : '';
        e._pairSxDate = lastSx ? str_(lastSx.SXDATE) : '';
        if (d3 !== null) sxRaiEach.push(d3);
        if (lastSx && !lastSx._raiDate) { lastSx._raiDate = e.RAIDATE; lastSx._sxRaiDays = d3; }
        raiRows.push(e);
      }
    });
  });

  // ----- กลุ่มนับรายคน (1 คน = 1 ค่า) -----
  var sxPerPat = [], raiPerPat = [], losSumPat = [], sxDcSumPat = [], firstRaiPat = [];
  var sxCountTally = {}, raiCountTally = {};
  var withSx = 0, withRai = 0, noEvent = 0;
  var patRows = [];

  pats.forEach(function (p) {
    var list = evByHn[p.HN] || [];
    var nsx = 0, nrai = 0, losSum = null, dcSum = null, firstRai = null, firstRaiDate = '';
    list.forEach(function (e) {
      if (e.TYPE === 'SX') {
        nsx++;
        if (e._los !== null) losSum = (losSum === null ? 0 : losSum) + e._los;
        if (e._sxDc !== null) dcSum = (dcSum === null ? 0 : dcSum) + e._sxDc;
      } else if (e.TYPE === 'RAI') {
        nrai++;
        if (firstRai === null && e._sxRai !== null) { firstRai = e._sxRai; firstRaiDate = e.RAIDATE; }
      }
    });
    if (!list.length) noEvent++;
    if (nsx) { withSx++; sxPerPat.push(nsx); bump_(sxCountTally, nsx + ' ครั้ง'); }
    if (nrai) { withRai++; raiPerPat.push(nrai); bump_(raiCountTally, nrai + ' ครั้ง'); }
    if (losSum !== null) losSumPat.push(losSum);
    if (dcSum !== null) sxDcSumPat.push(dcSum);
    if (firstRai !== null) firstRaiPat.push(firstRai);
    patRows.push({ p: p, nsx: nsx, nrai: nrai, losSum: losSum, dcSum: dcSum,
      firstRai: firstRai, firstRaiDate: firstRaiDate });
  });

  // ----- ตัวชี้วัด: ผ่าตัด -> Discharge ไม่เกิน 3 วัน -----
  var KPI_DAYS = 3;
  var kpi = { days: KPI_DAYS, pass: 0, total: 0, pct: '' };
  sxDcEach.forEach(function (d) { kpi.total++; if (d <= KPI_DAYS) kpi.pass++; });
  kpi.pct = kpi.total ? Math.round(kpi.pass / kpi.total * 1000) / 10 : '';
  sxRows.forEach(function (e) {
    e._kpi = (e._sxDc === null) ? '' : (e._sxDc <= KPI_DAYS ? 'ผ่าน' : 'ไม่ผ่าน');
  });

  // ต่อคน: ผ่านเมื่อทุกครั้งที่ผ่าตัด (ที่คำนวณได้) ออก รพ. ภายในเกณฑ์
  var kpiPat = { days: KPI_DAYS, pass: 0, total: 0, pct: '' };
  patRows.forEach(function (r) {
    var n = 0, fail = 0;
    (evByHn[r.p.HN] || []).forEach(function (e) {
      if (e.TYPE === 'SX' && e._sxDc !== null) { n++; if (e._sxDc > KPI_DAYS) fail++; }
    });
    r.kpi = n ? (fail ? 'ไม่ผ่าน' : 'ผ่าน') : '';
    if (n) { kpiPat.total++; if (!fail) kpiPat.pass++; }
  });
  kpiPat.pct = kpiPat.total ? Math.round(kpiPat.pass / kpiPat.total * 1000) / 10 : '';
  kpi.patient = kpiPat;

  // ----- สรุปรายปี (นับปีจากวันที่ผ่าตัด พ.ศ.) -----
  var yAll = newGroup_(), yMap = {};
  sxRows.forEach(function (e) {
    var y = (e._t === null || e._t === undefined) ? 'ไม่ระบุปี' : String(new Date(e._t).getUTCFullYear() + 543);
    if (!yMap[y]) yMap[y] = newGroup_();
    addToGroup_(yMap[y], e);
    addToGroup_(yAll, e);
  });
  var yearKeys = Object.keys(yMap).sort(function (a, b) {
    if (a === 'ไม่ระบุปี') return 1;
    if (b === 'ไม่ระบุปี') return -1;
    return (+b) - (+a);
  });
  var years = yearKeys.map(function (y) { return groupSummary_(y, yMap[y], KPI_DAYS); });
  var yearTotal = groupSummary_('ทุกปี', yAll, KPI_DAYS);

  // ----- RAI รายปี (นับปีจากวันที่ RAI) แต่ละครั้งนับจากการผ่าตัดก่อนหน้า -----
  var rAll = newGroup_(), rMap = {};
  raiRows.forEach(function (e) {
    var y = (e._t === null || e._t === undefined) ? 'ไม่ระบุปี' : String(new Date(e._t).getUTCFullYear() + 543);
    if (!rMap[y]) rMap[y] = newGroup_();
    addToGroup_(rMap[y], e, '_sxRai');
    addToGroup_(rAll, e, '_sxRai');
  });
  var rKeys = Object.keys(rMap).sort(function (a, b) {
    if (a === 'ไม่ระบุปี') return 1;
    if (b === 'ไม่ระบุปี') return -1;
    return (+b) - (+a);
  });
  var raiOut = {
    limit: RAI_LIMIT,
    withRaiPct: pct1_(withRai, pats.length),
    total: raiGroupSummary_('ทุกปี', rAll),
    years: rKeys.map(function (y) { return raiGroupSummary_(y, rMap[y]); })
  };

  return {
    generated: now_(),
    patients: pats.length,
    gender: tally_(gender),
    ageBand: tally_(band),
    ageStat: stat_(ages),
    patho: tally_(patho),

    nSx: nSx, nRai: nRai, withSx: withSx, withRai: withRai, noEvent: noEvent,
    proc: tally_(proc),
    dose: tally_(doseTally), doseStat: stat_(doses),
    sxCountTally: tally_(sxCountTally), raiCountTally: tally_(raiCountTally),

    // มุมมอง "ต่อคน"
    byPatient: {
      sxCount: stat_(sxPerPat),
      raiCount: stat_(raiPerPat),
      losSum: stat_(losSumPat),
      sxDcSum: stat_(sxDcSumPat),
      firstRai: stat_(firstRaiPat)
    },
    // มุมมอง "ต่อครั้ง"
    byEvent: {
      los: stat_(losEach),
      sxDc: stat_(sxDcEach),
      sxRai: stat_(sxRaiEach)
    },

    kpi: kpi,
    years: years,
    yearTotal: yearTotal,
    rai: raiOut,
    tc: buildCancerSummary_(pats, evByHn),
    _sxRows: sxRows, _raiRows: raiRows, _patRows: patRows, _byHn: byHn
  };
}

// เรียกจากหน้าเว็บ: ส่งเฉพาะตัวเลขสรุป
function apiStats() {
  requireAuth_();
  var s = computeStats_();
  delete s._sxRows; delete s._raiRows; delete s._patRows; delete s._byHn;
  return s;
}

// ---------- ส่งออกลง Google Sheets ----------
function resetSheet_(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.clear();
  return sh;
}

function writeTable_(sh, headers, rows, textCols) {
  sh.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#e3f2fd');
  sh.setFrozenRows(1);
  (textCols || []).forEach(function (c) {
    sh.getRange(2, c + 1, Math.max(sh.getMaxRows() - 1, 1), 1).setNumberFormat('@');
  });
  if (rows.length) sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
  for (var i = 0; i < headers.length; i++) sh.autoResizeColumn(i + 1);
}

function num_(v) { return (v === null || v === undefined || v === '') ? '' : v; }

function writeSummarySheet_(ss, s) {
  var sh = resetSheet_(ss, 'Dashboard');
  var R = [], bold = [];
  function blank() { R.push(['', '', '']); }
  function head(t) { bold.push(R.length + 1); R.push([t, '', '']); }
  function tallyBlock(title, list, denom, unit) {
    blank(); head(title);
    R.push(['รายการ', 'จำนวน (' + unit + ')', 'ร้อยละ']);
    if (!list.length) { R.push(['- ไม่มีข้อมูล -', '', '']); return; }
    list.forEach(function (x) {
      R.push([x.k, x.v, denom ? Math.round(x.v / denom * 1000) / 10 : '']);
    });
  }
  function statBlock(title, st, unit) {
    blank(); head(title);
    if (!st.n) { R.push(['ยังไม่มีข้อมูลพอคำนวณ', '', '']); return; }
    R.push(['จำนวนที่คำนวณได้', st.n, '']);
    R.push(['ค่าเฉลี่ย', st.mean, unit]);
    R.push(['มัธยฐาน', st.median, unit]);
    R.push(['ต่ำสุด', st.min, unit]);
    R.push(['สูงสุด', st.max, unit]);
  }

  head('สรุปข้อมูลผู้ป่วย Thyroid');
  R.push(['อัปเดตเมื่อ', '', '']);

  blank(); head('ภาพรวม');
  R.push(['ผู้ป่วยทั้งหมด', s.patients, 'คน']);
  R.push(['เคยผ่าตัด', s.withSx, 'คน']);
  R.push(['เคยได้ RAI', s.withRai, 'คน']);
  R.push(['ยังไม่มีข้อมูลการรักษา', s.noEvent, 'คน']);
  R.push(['การผ่าตัดทั้งหมด', s.nSx, 'ครั้ง']);
  R.push(['การให้ RAI ทั้งหมด', s.nRai, 'ครั้ง']);

  blank(); head('ตัวชี้วัด: ผ่าตัด → Discharge ไม่เกิน ' + s.kpi.days + ' วัน');
  R.push(['ก. ต่อการผ่าตัด 1 ครั้ง', '', '']);
  R.push(['ผ่านเกณฑ์', s.kpi.pass, 'ครั้ง']);
  R.push(['ไม่ผ่านเกณฑ์', s.kpi.total - s.kpi.pass, 'ครั้ง']);
  R.push(['การผ่าตัดที่คำนวณได้', s.kpi.total, 'ครั้ง']);
  R.push(['ร้อยละผ่านเกณฑ์', s.kpi.pct, '%']);
  R.push(['ข. ต่อคนไข้ (ผ่านเมื่อทุกครั้งที่ผ่าตัดผ่านเกณฑ์)', '', '']);
  R.push(['ผ่านเกณฑ์', s.kpi.patient.pass, 'คน']);
  R.push(['ไม่ผ่านเกณฑ์', s.kpi.patient.total - s.kpi.patient.pass, 'คน']);
  R.push(['คนไข้ที่คำนวณได้', s.kpi.patient.total, 'คน']);
  R.push(['ร้อยละผ่านเกณฑ์', s.kpi.patient.pct, '%']);

  // ===== สรุปผู้ป่วยมะเร็งไทรอยด์ =====
  var tc = s.tc;
  function compBlock(title, c, unit) {
    blank(); head(title);
    R.push(['ฐานคำนวณ ' + c.denom + ' คน', '', '']);
    c.items.forEach(function (x) { R.push([x.k, x.patients + ' คน (' + x.times + ' ครั้ง)', x.pct + '%']); });
    R.push(['ไม่มี complication', c.none, c.nonePct + '%']);
  }
  blank(); head('ผู้ป่วยมะเร็งไทรอยด์ ต่อปี แยกเพศ (นับปีจากวัน Bx)');
  R.push(['ปี พ.ศ.', 'ทั้งหมด (คน)', 'ชาย / หญิง']);
  tc.yearGender.concat([tc.yearGenderTotal]).forEach(function (r) {
    R.push([r.year, r.total, r.male + ' / ' + r.female + (r.other ? ' (ไม่ระบุ ' + r.other + ')' : '')]);
  });
  blank(); head('ชนิด Patho (ทุกคนที่ลงทะเบียน ' + tc.patients + ' คน)');
  R.push(['Patho', 'จำนวน (คน)', 'ร้อยละ']);
  tc.patho.forEach(function (x) { R.push([x.k, x.v, x.pct]); });

  blank(); head('การผ่าตัด');
  R.push(['ผ่าตัด (คน)', tc.surgery.patients, 'คน']);
  R.push(['ผ่าตัด (ครั้ง)', tc.surgery.times, 'ครั้ง']);
  tc.surgery.byCount.forEach(function (x) { R.push(['ผ่าตัด ' + x.k, x.v, x.pct + '%']); });
  blank(); head('คนที่ผ่าตัดหลายครั้ง ทำอะไรบ้าง');
  if (!tc.surgery.patterns.length) R.push(['- ไม่มี -', '', '']);
  tc.surgery.patterns.forEach(function (x) { R.push([x.n + ' ครั้ง: ' + x.pattern, x.patients, 'คน']); });

  blank(); head('นอน รพ. หลังผ่าตัด (นับจากวันที่ผ่าตัด)');
  R.push(['ผู้ป่วยที่ผ่าตัดทั้งหมด ' + tc.los.patients + ' คน ' + tc.los.surgeries + ' ครั้ง: เฉลี่ย (วัน)', tc.los.all.n ? tc.los.all.mean : '-', tc.los.all.n ? 'มัธยฐาน ' + tc.los.all.median : '']);
  R.push(['กลุ่มการผ่าตัด', 'ครั้ง / คน', 'เฉลี่ย (วัน)']);
  tc.los.byCat.forEach(function (x) {
    R.push([x.k, x.surgeries + ' / ' + x.patients, x.days.n ? x.days.mean : '-']);
  });
  R.push(['TTD ทั้งหมด: เฉลี่ย (วัน)', tc.los.ttd.days.n ? tc.los.ttd.days.mean : '-', tc.los.ttd.days.n ? 'มัธยฐาน ' + tc.los.ttd.days.median : '']);
  R.push(['TTD ทั้งหมด เทียบผู้ป่วยที่ผ่าตัด', tc.los.ttd.patients + ' / ' + tc.los.patients + ' คน', tc.los.ttd.pctOfPatients + '%']);
  R.push(['TTD ทั้งหมด เทียบการผ่าตัดทั้งหมด', tc.los.ttd.surgeries + ' / ' + tc.los.surgeries + ' ครั้ง', tc.los.ttd.pctOfSurgeries + '%']);

  compBlock('Complication หลังผ่าตัด', tc.compSx);
  blank(); head('ผู้ป่วยที่ได้ RAI');
  R.push(['ได้ RAI (คน)', tc.rai.patients, tc.rai.pct + '% ของผู้ป่วยมะเร็งไทรอยด์ทั้งหมด']);
  compBlock('Complication หลัง RAI', tc.rai.comp);
  blank(); head('ผู้ป่วยเสียชีวิต');
  R.push(['เสียชีวิต (คน)', tc.dead.total, tc.dead.pct + '% ของผู้ป่วยทั้งหมด']);
  tc.dead.byPatho.forEach(function (x) { R.push([x.k, x.v + ' คน', x.pctOfPatho + '% ของ ' + x.k + ' (' + x.total + ' คน)']); });

  tallyBlock('เพศ', s.gender, s.patients, 'คน');
  tallyBlock('ช่วงอายุ', s.ageBand, s.patients, 'คน');
  statBlock('อายุ (ปี)', s.ageStat, 'ปี');
  tallyBlock('ผลพยาธิวิทยา (Patho)', s.patho, s.patients, 'คน');
  tallyBlock('ชนิดการผ่าตัด (นับทุกครั้ง)', s.proc, s.nSx, 'ครั้ง');
  tallyBlock('จำนวนครั้งที่ผ่าตัด ต่อคน', s.sxCountTally, s.withSx, 'คน');
  tallyBlock('จำนวนครั้งที่ได้ RAI ต่อคน', s.raiCountTally, s.withRai, 'คน');
  tallyBlock('ความแรง RAI', s.dose, s.nRai, 'ครั้ง');

  blank(); head('===== มุมมอง ก: ต่อคน (1 คน = 1 ค่า) =====');
  statBlock('จำนวนครั้งผ่าตัด ต่อคน (เฉพาะคนที่เคยผ่าตัด)', s.byPatient.sxCount, 'ครั้ง');
  statBlock('วันนอนโรงพยาบาลรวมทั้งหมด ต่อคน', s.byPatient.losSum, 'วัน');
  statBlock('วันหลังผ่าตัดรวมทั้งหมด ต่อคน', s.byPatient.sxDcSum, 'วัน');
  statBlock('ผ่าตัด → RAI ครั้งแรก ต่อคน', s.byPatient.firstRai, 'วัน');

  blank(); head('===== มุมมอง ข: ต่อครั้ง (1 ครั้ง = 1 ค่า) =====');
  statBlock('นอนโรงพยาบาล (รับไว้ → Discharge) ต่อการผ่าตัด 1 ครั้ง', s.byEvent.los, 'วัน');
  statBlock('ผ่าตัด → Discharge ต่อการผ่าตัด 1 ครั้ง', s.byEvent.sxDc, 'วัน');
  statBlock('ผ่าตัด → RAI ต่อการให้ RAI 1 ครั้ง', s.byEvent.sxRai, 'วัน');

  if (s.badDate) {
    blank();
    head('ข้อมูลที่ข้ามไป');
    R.push(['วันที่เรียงลำดับไม่ถูกต้อง (เช่น Discharge มาก่อนผ่าตัด)', s.badDate, 'รายการ']);
  }

  sh.getRange(1, 1, R.length, 3).setValues(R);
  sh.getRange(2, 2).setNumberFormat('@').setValue(s.generated);
  bold.forEach(function (r) {
    sh.getRange(r, 1, 1, 3).setFontWeight('bold').setBackground('#e3f2fd');
  });
  sh.getRange(1, 1, 1, 3).setFontSize(13);
  sh.setColumnWidth(1, 330); sh.setColumnWidth(2, 110); sh.setColumnWidth(3, 90);
}

function writePatientSheet_(ss, s) {
  var H = ['HN', 'ชื่อ', 'เพศ', 'อายุ', 'Dx', 'Patho', 'Patho เพิ่มเติม',
    'จำนวนครั้งผ่าตัด', 'จำนวนครั้ง RAI', 'วันนอน รพ. รวม (วัน)',
    'วันหลังผ่าตัดรวม (วัน)', 'ผ่าตัด→RAI ครั้งแรก (วัน)', 'ตัวชี้วัด ≤' + s.kpi.days + ' วัน (ต่อคน)', 'เสียชีวิต', 'หมายเหตุเสียชีวิต'];
  var rows = s._patRows.map(function (r) {
    var p = r.p, a = ageOf_(p);
    return [p.HN, p.NAME || '', p.GENDER || '', a === null ? '' : a, p.DX || '',
      p.PATHO || '', p.PATHONOTE || '', r.nsx, r.nrai,
      num_(r.losSum), num_(r.dcSum), num_(r.firstRai), r.kpi || '', p.DEAD || '', p.DEADNOTE || ''];
  });
  writeTable_(resetSheet_(ss, 'ตารางรายคน'), H, rows, [0]);
}

function writeSurgerySheet_(ss, s) {
  var H = ['HN', 'ชื่อ', 'เพศ', 'อายุ', 'Patho', 'ผ่าตัดครั้งที่', 'ผ่าตัดอะไร',
    'วันที่รับไว้ รพ.', 'วันที่ผ่าตัด', 'วันที่ Discharge', 'นอน รพ. (วัน)',
    'ผ่าตัด→Discharge (วัน)', 'ตัวชี้วัด ≤' + s.kpi.days + ' วัน',
    'วันที่ RAI ถัดไป', 'ผ่าตัด→RAI (วัน)', 'Complication', 'หมายเหตุ'];
  var rows = s._sxRows.map(function (e) {
    var p = s._byHn[e.HN] || {}, a = ageOf_(p);
    return [e.HN, p.NAME || '', p.GENDER || '', a === null ? '' : a, p.PATHO || '',
      e.SXNO || '', e.PROC || '', e.ADMIT || '', e.SXDATE || '', e.DC || '',
      num_(e._los), num_(e._sxDc), e._kpi || '',
      e._raiDate || '', num_(e._sxRaiDays), e.COMP || '', e.NOTE || ''];
  });
  writeTable_(resetSheet_(ss, 'ตารางผ่าตัด'), H, rows, [0, 7, 8, 9, 13]);
}

function writeRaiSheet_(ss, s) {
  var H = ['HN', 'ชื่อ', 'เพศ', 'อายุ', 'Patho', 'RAI ครั้งที่', 'วันที่ RAI',
    'ความแรง (mCi)', 'ผ่าตัดครั้งที่ก่อนหน้า', 'วันที่ผ่าตัดครั้งนั้น',
    'ผ่าตัด→RAI (วัน)', 'Complication', 'หมายเหตุ'];
  var seq = {};
  var rows = s._raiRows.map(function (e) {
    var p = s._byHn[e.HN] || {}, a = ageOf_(p);
    seq[e.HN] = (seq[e.HN] || 0) + 1;
    return [e.HN, p.NAME || '', p.GENDER || '', a === null ? '' : a, p.PATHO || '',
      seq[e.HN], e.RAIDATE || '', e.RAIDOSE || '',
      e._pairSxNo || '', e._pairSxDate || '', num_(e._sxRai), e.COMP || '', e.NOTE || ''];
  });
  writeTable_(resetSheet_(ss, 'ตาราง RAI'), H, rows, [0, 6, 9]);
}

function writeYearSheet_(ss, s) {
  var sh = resetSheet_(ss, 'สรุปรายปี');
  var R = [], hdr = [];
  function line(a) { while (a.length < 6) a.push(''); R.push(a); }
  function head(a) { hdr.push(R.length + 1); line(a); }
  var all = [s.yearTotal].concat(s.years);

  head(['สรุปรายปี (นับปีจากวันที่ผ่าตัด พ.ศ.)', '', '', '', '', '']);
  line(['นอน รพ. = จำนวนวันหลังผ่าตัดจนถึง Discharge · คนที่ผ่าตัดหลายครั้งนับอยู่ในทุกช่วงที่เขามีการผ่าตัด (ผลรวม % ของคนอาจเกิน 100)', '', '', '', '', '']);
  line(['']);
  head(['ปี พ.ศ.', 'ผู้ป่วย (คน)', 'ผ่าตัด (ครั้ง)', 'ผ่านเกณฑ์ ต่อครั้ง (%)', 'ผ่านเกณฑ์ ต่อคน (%)', 'เกณฑ์ (วัน)']);
  all.forEach(function (y) {
    line([y.year, y.patients, y.surgeries,
      y.kpiSurg.total ? y.kpiSurg.pass + '/' + y.kpiSurg.total + ' = ' + y.kpiSurg.pct + '%' : '-',
      y.kpiPat.total ? y.kpiPat.pass + '/' + y.kpiPat.total + ' = ' + y.kpiPat.pct + '%' : '-',
      '≤ ' + s.kpi.days]);
  });
  all.forEach(function (y) {
    line(['']);
    head(['ปี ' + y.year + ' — ช่วงวันหลังผ่าตัด', 'จำนวนครั้ง', '% ของการผ่าตัดทั้งหมด', 'จำนวนคน', '% ของคนไข้ทั้งหมด', '']);
    for (var i = 0; i < DAY_BUCKETS.length; i++) {
      line([y.distSurg[i].k, y.distSurg[i].v, y.distSurg[i].pct, y.distPat[i].v, y.distPat[i].pct, '']);
    }
    if (y.unknownSurg || y.unknownPat) {
      line(['คำนวณไม่ได้ (ไม่มีวันผ่าตัดหรือ Discharge)', y.unknownSurg, '', y.unknownPat, '', '']);
    }
  });

  // ----- RAI รายปี -----
  var ra = [s.rai.total].concat(s.rai.years);
  line(['']);
  head(['RAI (กลืนแร่) รายปี — นับปีจากวันที่ RAI', '', '', '', '', '']);
  line(['จำนวนวัน = วัน RAI ลบวันผ่าตัดครั้งล่าสุดก่อนหน้า · เกณฑ์ ไม่เกิน ' + s.rai.limit + ' วัน หลังผ่าตัด · ต่อคนใช้ค่าที่นานที่สุดของคนนั้น', '', '', '', '', '']);
  line(['']);
  head(['ปี พ.ศ.', 'ผู้ป่วยที่ได้ RAI (คน)', 'RAI (ครั้ง)', 'ภายใน ' + s.rai.limit + ' วัน ต่อครั้ง', 'ภายใน ' + s.rai.limit + ' วัน ต่อคน', 'เฉลี่ย / มัธยฐาน (วัน)']);
  ra.forEach(function (y) {
    line([y.year, y.patients, y.rai,
      y.kpiRai.total ? y.kpiRai.pass + '/' + y.kpiRai.total + ' = ' + y.kpiRai.pct + '%' : '-',
      y.kpiPat.total ? y.kpiPat.pass + '/' + y.kpiPat.total + ' = ' + y.kpiPat.pct + '%' : '-',
      y.days.n ? y.days.mean + ' / ' + y.days.median : '-']);
  });
  ra.forEach(function (y) {
    line(['']);
    head(['RAI ปี ' + y.year + ' — ช่วงหลังผ่าตัด', 'ช่วงสัปดาห์', 'จำนวนครั้ง', '% ของครั้ง', 'จำนวนคน', '% ของคน']);
    for (var i = 0; i < RAI_BUCKETS.length; i++) {
      line([y.distRai[i].k, y.distRai[i].wk, y.distRai[i].v, y.distRai[i].pct, y.distPat[i].v, y.distPat[i].pct]);
    }
    if (y.unknownRai || y.unknownPat) {
      line(['คำนวณไม่ได้ (ไม่มีวันผ่าตัดก่อนหน้า)', '', y.unknownRai, '', y.unknownPat, '']);
    }
  });

  sh.getRange(1, 1, R.length, 6).setValues(R);
  hdr.forEach(function (r) { sh.getRange(r, 1, 1, 6).setFontWeight('bold').setBackground('#e3f2fd'); });
  sh.setColumnWidth(1, 330);
  for (var c = 2; c <= 6; c++) sh.setColumnWidth(c, 150);
}

function apiExportDashboard() {
  requireAuth_();
  var s = computeStats_();
  var ss = ss_();
  writeSummarySheet_(ss, s);
  writePatientSheet_(ss, s);
  writeYearSheet_(ss, s);
  writeSurgerySheet_(ss, s);
  writeRaiSheet_(ss, s);
  SpreadsheetApp.flush();
  return 'ส่งออกแล้ว 5 แท็บ: Dashboard, สรุปรายปี, ตารางรายคน, ตารางผ่าตัด, ตาราง RAI';
}

// รันฟังก์ชันนี้ในหน้า Apps Script เพื่อตรวจว่าการ Deploy มองเห็นผู้ใช้หรือไม่
function whoAmI() {
  var a = '', b = '';
  try { a = Session.getActiveUser().getEmail(); } catch (e) { a = 'error'; }
  try { b = Session.getEffectiveUser().getEmail(); } catch (e) { b = 'error'; }
  Logger.log('ActiveUser (คนที่กำลังเข้า): ' + (a || '(ว่าง)'));
  Logger.log('EffectiveUser (เจ้าของสคริปต์): ' + (b || '(ว่าง)'));
  return { active: a, effective: b };
}
