/**
 * ME1 IPD — Voice Form backend (Google Apps Script, โปรเจกต์แยก / standalone)
 *
 * ขั้นตอน (ละเอียดใน README.md):
 *   1) script.google.com > New project > วางไฟล์นี้ทั้งหมด > Save
 *   2) Project Settings > Script Properties: เพิ่ม TEAM_CODE = รหัสทีม
 *   3) เลือกฟังก์ชัน setupSpreadsheet > Run  → สร้าง Google Sheet ตามต้นแบบ "แบบบันทึก ME1 IPD"
 *      (ดูลิงก์ชีตใน Execution log และใน Script Properties: SHEET_ID)
 *   4) Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone)
 */

var CFG = {
  LOG_SHEET: 'OK ME1 IPD',
  ADMIN_SHEET: 'OK Admin',
  APP_LOG_SHEET: 'App_Log',
  GLOSSARY_SHEET: 'App_Glossary',
  DRUGS_SHEET: 'App_Drugs',
  HEADER_ROWS: 2,
  TEMPLATE_ROWS: 2000,          // จำนวนแถวที่ใส่ dropdown/สูตรไว้ล่วงหน้า
  ADMIN_COL: { time: 'B', processing: 'D', staff: 'E' }
};

var STAFF_SKIP = [/มิย|สค|ลาออก|เข้า IPD/, /^กรณีไม่ลงข้อมูล/];

var DEFAULT_GLOSSARY = [
  ['key', 'คีย์, ขี่, ขี้, คี'],
  ['ไม่ off', 'ไม่ออฟ, ไม่อ๊อฟ'],
  ['off', 'ออฟ, อ๊อฟ, อ็อฟ'],
  ['extra', 'เอ็กซ์ตร้า, เอ็กซ์ตรา, เอ็กตร้า'],
  ['get', 'เก็ท, เก็ต'],
  ['prn', 'พีอาร์เอ็น'],
  ['nss', 'เอ็นเอสเอส'],
  ['iv', 'ไอวี'],
  ['inj', 'อินเจ็ค, อินเจ็คชั่น'],
  ['drip', 'ดริป']
];

var STOPWORDS = ('key off inj extra get iv ml q mg prn tab day g cc dose hs syr for one cont code x1 unit er pc ac ' +
  'robot profile hm atb drip water stat ed od bid tid qid mo not and the with ward stock adult sc im po ivd cap amp ' +
  'vial sol susp oint ok set no screen copy new old ipd opd admit note am pm hr hrs min wk iu mcg ear eye').split(' ');

/* ================= เปิดชีต ================= */

function getSS_() {
  var id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  var active = SpreadsheetApp.getActive();          // กรณีวางสคริปต์ไว้ในชีตโดยตรง
  if (active) return active;
  throw new Error('ยังไม่ได้สร้างชีต — รันฟังก์ชัน setupSpreadsheet ก่อน');
}

/* ================= HTTP ================= */

function doGet(e) {
  try {
    var p = (e && e.parameter) || {};
    checkCode_(p.code);
    if (p.action === 'init') return json_(init_());
    return json_({ ok: false, error: 'unknown action' });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

function doPost(e) {
  try {
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    checkCode_(body.code);
    if (body.action === 'save') return json_(save_(body.record || {}));
    return json_({ ok: false, error: 'unknown action' });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function checkCode_(code) {
  var expected = PropertiesService.getScriptProperties().getProperty('TEAM_CODE');
  if (!expected) throw new Error('ยังไม่ได้ตั้ง TEAM_CODE ใน Script Properties');
  if (String(code || '') !== String(expected)) throw new Error('รหัสทีมไม่ถูกต้อง');
}

/* ================= init: ส่งตัวเลือกให้หน้าเว็บ ================= */

function init_() {
  var ss = getSS_();
  var admin = ss.getSheetByName(CFG.ADMIN_SHEET);
  if (!admin) throw new Error('ไม่พบแท็บ ' + CFG.ADMIN_SHEET);
  return {
    ok: true,
    times: colValues_(admin, CFG.ADMIN_COL.time).filter(function (v) { return v !== 'ไม่ทราบ'; }),
    processing: colValues_(admin, CFG.ADMIN_COL.processing),
    staff: staffList_(admin),              // ส่งเฉพาะ id + รหัส/ชื่อเล่น ไม่ส่งชื่อเต็ม
    drugs: drugList_(ss),
    glossary: glossary_(ss)
  };
}

function colValues_(sheet, col) {
  var last = sheet.getLastRow();
  if (last < 2) return [];
  return sheet.getRange(col + '2:' + col + last).getDisplayValues()
    .map(function (r) { return String(r[0]).trim(); })
    .filter(function (v) { return v !== ''; });
}

// "03 สุธิดา (แมว)" → "03 แมว"
function nicknameLabel_(full) {
  var m = full.match(/^\s*(\d+)?\s*([^()]*?)\s*(?:\(([^)]+)\))?\s*$/);
  if (!m) return full;
  var code = m[1] || '', name = (m[2] || '').trim(), nick = (m[3] || '').trim();
  return (code ? code + ' ' : '') + (nick || name);
}

function staffList_(admin) {
  var col = CFG.ADMIN_COL.staff, last = admin.getLastRow(), out = [];
  if (last < 2) return out;
  var vals = admin.getRange(col + '2:' + col + last).getDisplayValues();
  for (var i = 0; i < vals.length; i++) {
    var full = String(vals[i][0]).trim();
    if (!full) continue;
    if (STAFF_SKIP.some(function (re) { return re.test(full); })) continue;
    out.push({ id: i + 2, label: nicknameLabel_(full) });
  }
  return out;
}

function staffFullById_(ss, id) {
  var admin = ss.getSheetByName(CFG.ADMIN_SHEET);
  var row = parseInt(id, 10);
  if (!row || row < 2 || row > admin.getLastRow()) return '';
  return String(admin.getRange(CFG.ADMIN_COL.staff + row).getDisplayValue()).trim();
}

// รายชื่อยา = App_Drugs (ชุดตั้งต้นจากบันทึก ก.ย. 69) + คำอังกฤษในคอลัมน์ C ของบันทึกใหม่
function drugList_(ss) {
  var count = {}, stop = {};
  STOPWORDS.forEach(function (w) { stop[w] = 1; });
  function add(w, n) { count[w] = (count[w] || 0) + n; }

  var sh = ss.getSheetByName(CFG.LOG_SHEET), last = sh.getLastRow();
  if (last > CFG.HEADER_ROWS) {
    sh.getRange('C' + (CFG.HEADER_ROWS + 1) + ':C' + last).getDisplayValues().forEach(function (r) {
      (String(r[0]).match(/[A-Za-z][A-Za-z0-9\-]*/g) || []).forEach(function (w) {
        var wl = w.toLowerCase().replace(/-+$/, '');
        if (wl.length < 3 || stop[wl] || /^[qx]\d+$/.test(wl)) return;
        add(wl, 1);
        var base = wl.replace(/\d+$/, '').replace(/-+$/, '');
        if (base.length >= 3 && base !== wl && !stop[base]) add(base, 1);
      });
    });
  }
  var ds = ss.getSheetByName(CFG.DRUGS_SHEET);
  if (ds && ds.getLastRow() > 1) {
    var vals = ds.getRange(2, 1, ds.getLastRow() - 1, 1).getDisplayValues();
    vals.forEach(function (r, i) {
      var w = String(r[0]).trim().toLowerCase();
      if (w) add(w, 0.001 * (vals.length - i));        // รักษาลำดับความถี่เดิม
    });
  }
  return Object.keys(count).sort(function (a, b) { return count[b] - count[a]; }).slice(0, 1000);
}

function glossary_(ss) {
  var sh = ss.getSheetByName(CFG.GLOSSARY_SHEET);
  if (!sh) {
    sh = ss.insertSheet(CFG.GLOSSARY_SHEET);
    writeGlossary_(sh);
  }
  var last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 2).getDisplayValues()
    .map(function (r) {
      return { to: String(r[0]).trim(), from: String(r[1]).split(',').map(function (s) { return s.trim(); }).filter(String) };
    })
    .filter(function (g) { return g.to && g.from.length; });
}

function writeGlossary_(sh) {
  sh.getRange(1, 1, 1, 2).setValues([['คำที่ถูก', 'คำที่ระบบได้ยินผิด (คั่นด้วย ,)']]).setFontWeight('bold');
  sh.getRange(2, 1, DEFAULT_GLOSSARY.length, 2).setValues(DEFAULT_GLOSSARY);
  sh.setColumnWidth(1, 140); sh.setColumnWidth(2, 360); sh.setFrozenRows(1);
}

/* ================= save: บันทึก 1 รายการ ================= */

function save_(rec) {
  var detail = String(rec.detail || '').trim();
  if (!detail) throw new Error('ยังไม่ได้กรอกรายละเอียด ME');
  if (!rec.processing) throw new Error('ยังไม่ได้เลือก Processing error');

  var d = String(rec.date || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  var date = d ? new Date(+d[1], +d[2] - 1, +d[3]) : new Date();

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var ss = getSS_();
    var log = ss.getSheetByName(CFG.APP_LOG_SHEET);
    if (!log) { log = ss.insertSheet(CFG.APP_LOG_SHEET); writeAppLogHeader_(log); }

    // กันบันทึกซ้ำ (ส่งซ้ำตอนเน็ตหลุด)
    if (rec.clientId && log.getLastRow() > 1) {
      var ids = log.getRange(2, 5, log.getLastRow() - 1, 1).getValues();
      for (var k = 0; k < ids.length; k++) {
        if (String(ids[k][0]) === String(rec.clientId)) return { ok: true, row: log.getRange(k + 2, 2).getValue(), duplicate: true };
      }
    }
    var staffKey = rec.staffKeyId ? staffFullById_(ss, rec.staffKeyId) : '';
    var staffDispense = rec.staffDispenseId ? staffFullById_(ss, rec.staffDispenseId) : '';

    var sh = ss.getSheetByName(CFG.LOG_SHEET);
    var row = nextEmptyRow_(sh);
    // เขียนเฉพาะ A–E และ G (F = เภสัชกร เว้นไว้) ไม่แตะคอลัมน์อื่นที่มีสูตร
    sh.getRange(row, 1, 1, 5).setValues([[date, rec.time || '', detail, rec.processing, staffKey]]);
    sh.getRange(row, 7).setValue(staffDispense);

    log.appendRow([new Date(), row, detail, String(rec.raw || ''), String(rec.clientId || '')]);
    return { ok: true, row: row };
  } finally {
    lock.releaseLock();
  }
}

function writeAppLogHeader_(log) {
  log.getRange(1, 1, 1, 5).setValues([['บันทึกเมื่อ', 'แถวที่', 'รายละเอียด (ที่บันทึก)', 'ข้อความจากเสียงก่อนแก้', 'รหัสอ้างอิงจากเครื่อง']]).setFontWeight('bold');
  log.setFrozenRows(1);
}

// แถวว่างถัดไป = แถวหลังสุดที่คอลัมน์ C มีข้อมูล + 1
function nextEmptyRow_(sh) {
  var last = sh.getLastRow();
  if (last <= CFG.HEADER_ROWS) return CFG.HEADER_ROWS + 1;
  var vals = sh.getRange('C1:C' + last).getValues();
  for (var i = vals.length - 1; i >= CFG.HEADER_ROWS; i--) {
    if (String(vals[i][0]).trim() !== '') return i + 2;
  }
  return CFG.HEADER_ROWS + 1;
}

/* ================= สร้าง Google Sheet ตามต้นแบบ ================= */

function setupSpreadsheet() {
  var props = PropertiesService.getScriptProperties();
  var existing = props.getProperty('SHEET_ID');
  if (existing) {
    Logger.log('มีชีตอยู่แล้ว: https://docs.google.com/spreadsheets/d/' + existing +
      '\nถ้าต้องการสร้างใหม่ ให้ลบ SHEET_ID ใน Script Properties ก่อน');
    return;
  }
  var ss = SpreadsheetApp.create('แบบบันทึก ME1 IPD (แอปเสียง)');
  ss.setSpreadsheetTimeZone('Asia/Bangkok');

  var log = ss.getSheets()[0];
  log.setName(CFG.LOG_SHEET);
  var admin = ss.insertSheet(CFG.ADMIN_SHEET);
  buildAdmin_(admin);
  buildLog_(log, admin);

  writeGlossary_(ss.insertSheet(CFG.GLOSSARY_SHEET));
  var drugs = ss.insertSheet(CFG.DRUGS_SHEET);
  drugs.getRange(1, 1).setValue('ชื่อยา (ตั้งต้นจากบันทึก ก.ย. 69 — เพิ่มได้)').setFontWeight('bold');
  drugs.getRange(2, 1, TPL.drugs.length, 1).setValues(TPL.drugs.map(function (d) { return [d]; }));
  drugs.setColumnWidth(1, 260); drugs.setFrozenRows(1);
  writeAppLogHeader_(ss.insertSheet(CFG.APP_LOG_SHEET));

  props.setProperty('SHEET_ID', ss.getId());
  Logger.log('สร้างชีตแล้ว: ' + ss.getUrl());
}

function buildAdmin_(admin) {
  Object.keys(TPL.admin).forEach(function (col) {
    var vals = TPL.admin[col];
    admin.getRange(col + '1:' + col + vals.length).setValues(vals.map(function (v) { return [v]; }));
  });
  Object.keys(TPL.adminExtra).forEach(function (col) {
    var vals = TPL.adminExtra[col];
    vals.forEach(function (v, i) {
      var cell = admin.getRange(col + (i + 1));
      if (String(v).charAt(0) === '=') cell.setFormula(v); else cell.setValue(v);
    });
  });
  admin.getRange('1:1').setFontWeight('bold').setWrap(true);
  admin.setFrozenRows(1);
  admin.setColumnWidths(1, 12, 180);
}

function buildLog_(log, admin) {
  var n = CFG.TEMPLATE_ROWS, first = CFG.HEADER_ROWS + 1, lastRow = first + n - 1;
  if (log.getMaxRows() < lastRow) log.insertRowsAfter(log.getMaxRows(), lastRow - log.getMaxRows());
  if (log.getMaxColumns() < 29) log.insertColumnsAfter(log.getMaxColumns(), 29 - log.getMaxColumns());

  log.getRange(1, 1, 1, TPL.row1.length).setValues([TPL.row1]).setWrap(true).setFontColor('#8a1c15');
  log.getRange(2, 1, 1, TPL.row2.length).setValues([TPL.row2]).setFontWeight('bold').setWrap(true).setBackground('#efece4');
  log.setFrozenRows(2); log.setFrozenColumns(2);
  log.getRange('A' + first + ':A' + lastRow).setNumberFormat('dd/mm/yyyy');
  log.getRange('C' + first + ':C' + lastRow).setWrap(true);
  var widths = { 1: 95, 2: 80, 3: 300, 4: 120, 5: 150, 6: 140, 7: 150, 8: 120, 9: 150, 10: 130, 11: 130, 12: 130, 13: 110, 14: 90, 15: 110, 16: 90, 17: 150 };
  Object.keys(widths).forEach(function (c) { log.setColumnWidth(+c, widths[c]); });

  // Dropdown อ้างอิงรายการใน OK Admin (เหมือนไฟล์ต้นแบบ) — ใส่ค่าอื่นได้แต่จะมีเตือน
  function listRule(col) {
    var len = TPL.admin[col].length;
    return SpreadsheetApp.newDataValidation()
      .requireValueInRange(admin.getRange(col + '2:' + col + len), true)
      .setAllowInvalid(true).build();
  }
  var map = { B: 'B', D: 'D', E: 'E', F: 'F', G: 'E', H: 'G', I: 'H', J: 'I', K: 'F', L: 'F', M: 'J', O: 'L', P: 'B' };
  Object.keys(map).forEach(function (target) {
    log.getRange(target + first + ':' + target + lastRow).setDataValidation(listRule(map[target]));
  });

  // คอลัมน์ช่วยสำหรับ Admin (S–AA) — สูตรเดียวต่อคอลัมน์ คำนวณเฉพาะแถวที่มีรายละเอียด
  function af(col, expr) { log.getRange(col + first).setFormula('=ARRAYFORMULA(IF(C' + first + ':C="",,' + expr + '))'); }
  af('S', 'IF(D' + first + ':D="","NO","YES")');
  af('T', 'IF(H' + first + ':H="","NO","YES")');
  af('U', 'IF(I' + first + ':I="","NO","YES")');
  af('V', "IF(O" + first + ":O='OK Admin'!$L$2,\"NO\",\"YES\")");
  af('W', 'IF(E' + first + ':E="","NO","YES")');
  af('X', 'IF(G' + first + ':G="","NO","YES")');
  af('Y', 'IF(F' + first + ':F="","NO","YES")');
  af('Z', 'IF(J' + first + ':J="","NO","YES")');
  var J = 'J' + first + ':J';
  af('AA', 'IF((' + J + '="")+(' + J + "='OK Admin'!$I$2),\"no ME\"," +
           'IF(' + J + "='OK Admin'!$I$3,\"ไม่ทราบ\"," +
           'IF(' + J + "='OK Admin'!$I$4,\"A-B\"," +
           'IF((' + J + "='OK Admin'!$I$5)+(" + J + "='OK Admin'!$I$6),\"C-D\",\"E up\"))))");
  log.getRange('S2:AA2').setBackground('#e3f2ee');
}

/* ================= ทดสอบใน editor ================= */
function testInit() { Logger.log(JSON.stringify(init_()).slice(0, 2000)); }

/* ================= ข้อมูลต้นแบบจาก "แบบบันทึก ME1 IPD_กย 69.xlsx" ================= */
var TPL = {
 "row1": [
  " ",
  "",
  "",
  "",
  "",
  "",
  "",
  "",
  "กรณี TE หรือ DE cat C up\nกรุณาลง RM center เท่านั้น",
  "",
  "",
  "",
  "",
  "",
  "",
  "",
  "",
  "",
  "ไม่ต้องกรอก สำหรับ Admin เท่านั้น\nแปลงข้อมูลอัตโนมัติ ครับ",
  "",
  "",
  "",
  "",
  "",
  "Admin = ภก.ปวัฒน์, ภญ.ปิยนาฏ, ชยาภรณ์ บุญอยู่\nเริ่ม 5/5/63 ปรับปรุงล่าสุด 15/5/63",
  "",
  "",
  "",
  "Check list Before interprete"
 ],
 "row2": [
  "ว/ด/ป ที่เกิด ME",
  "เวลาที่เกิด ME",
  "ระบุรายละเอียด ME",
  "Processing error",
  "ชื่อ จนท. Key",
  "ชื่อเภสัช Screen",
  "ชื่อ จนท. จัดยา",
  "Transcribing error",
  "Dispensing error",
  "Cate ระบุกรณี TE, DE",
  "TE, DE Screen",
  "TE, DE Check",
  "Ward ระบุกรณี TE, DE",
  "AN ระบุกรณี TE, DE",
  "HAD",
  "เวลาที่แก้ไข ME",
  "ผู้บันทึก",
  "",
  "PreDE Y/N",
  "TE Y/N",
  "DE Y/N",
  "HAD Y/N",
  "KEY Y/N",
  "จัดยา Y/N",
  "เภสัช Screen Y/N",
  "Cat Y/N",
  "Cat range",
  "HAD ID",
  ""
 ],
 "admin": {
  "A": [
   "วันที่เกิด ME",
   "วัน/เดือน/ปี"
  ],
  "B": [
   "เกิดใน / นอก เวลาราชการ",
   "ไม่ทราบ",
   "ในเวลา",
   "นอกเวลา"
  ],
  "C": [
   "ระบุรายละเอียด"
  ],
  "D": [
   "Processing error",
   "no Processing error",
   "1. ชื่อ ผป ",
   "2. ชื่่อยา",
   "3. ความแรง",
   "4. จำนวน",
   "5. วิธีบริหาร",
   "6. ขาด/เกิน",
   "7. รูปแบบ",
   "8. ติดฉลากผิด",
   "9. บรรจุภัณฑ์",
   "10. ใส่สูตรผิด",
   "11. ไม่ off ยา",
   "12. ยาเสื่อมสภาพ",
   "13. ยามี DI",
   "14. มีประวัติแพ้",
   "15. อื่นๆ"
  ],
  "E": [
   "รายชื่อ จนท.",
   "ไม่ทราบ",
   "เภสัชกร staff",
   "Trainee",
   "01 อภิญญาณ (ปุ๊ก)",
   "02 วิมลศิริ (อาม)",
   "03 สุธิดา (แมว)",
   "04 รจนชัย (รจ)",
   "05 พรพิมล (อาย)",
   "06 ขวัญรัตน์ (ปอย)",
   "07 เต็มศิริ (เต็ม)",
   "08 อิสรา (เอม)",
   "09 ศศิธร (มาร์)",
   "10 อารียา (มิ้ง)",
   "11 ญาณิกา (กิ่ง)",
   "12 วารุณี (มิ้น)",
   "13 สุวนันท์ (น้อย)",
   "14 วิลาวัณย์ (แฮม)",
   "เกษร (นา)",
   "เชนิยา (กลอย)",
   "ดาวอัมพร (ดาว)",
   "ธนพร (เฟรช)",
   "นริศรา (อ้อน)",
   "นันท์ธิญาภรณ์ (นิกกี้)",
   "เนตรอัปสร (เอมชิ)",
   "บัวทอง (อี๊ด)",
   "ไพลิน (เล็ก)",
   "มยุรี (มะ)",
   "รัชนก (ปาย)",
   "วัฒนาพร (มด)",
   "วานีต้า (ด้า)",
   "วาสนา (วาส)",
   "วิจิตตรา (น้อย)",
   "สุรภา (โอ๋เอ๋)",
   "สายฝน (ตุ๋ง)",
   "สิรารัตน์ (อิ้ง)",
   "สุคนธา (จุ๋ม)",
   "สุพรรษา (ปู)",
   "อนงค์ (นง)",
   "แพรวผกา (แพรว)",
   "รังสิมา (น้ำอุ่น)",
   "นัชชา (นัชชา)",
   "จันทิมา (เปตอง)",
   "กรณีไม่ลงข้อมูล(Blank)"
  ],
  "F": [
   "รายชื่อเภสัชกร",
   "ไม่ทราบ",
   "Trainee",
   "01 มัณฑนาประสารเกตุ",
   "02 กรัณย์กรฤทธิ์ภักดี",
   "03 พรรณทิพย์นิกรวัฒน์",
   "04 เปรมชัยเม่นสิน",
   "05 ปวัฒน์ผุดวาย",
   "06 เขมจิราเยาวกุลพัฒนา",
   "07 พิมพรัศมิ์ยังคง",
   "08 จุฑามาศวุฒิกรสัมมากิจ",
   "09 อัจฉราภรณ์ ทองเย็น",
   "10 อชิรญาบุตรดาน้อย",
   "11 ปิยะนาฎเชิงสะอาด",
   "12 ปิยะนัฐไทยปิยะ",
   "13 พรทิพย์วรกิจพูนผล",
   "14 วรรณวัฒน์สินเจริญ",
   "15 วรวิชวรผล",
   "16 ขวัญชนกทับปะระ",
   "17 สุทธิดา ห้วยล้ำ",
   "18 วิศิษฏ์ ริมเขต",
   "ก้านตองตันทวีวงศ์",
   "จำรัสลักษณ์ ขวัญนวล",
   "จิราวดีสุทธิ",
   "จุฑารัตน์ จันทร์กลับ",
   "เจษฎาพรภักดี",
   "ชวัลพัชร์ กาญจน์ลัทธ์",
   "นภัสววรณ อุรุวงศ์",
   "นิรมลศรีสุข",
   "พรรณพร กิริยา",
   "พิลาศลักษณ์ภู่",
   "ภรณ์ทิพย์สร้อยพิทักษ์",
   "ภัคพรรณคำแฝง",
   "ภัทรชนม์พิเนตสิริ",
   "รุ่งรัตน์เศรษฐโกมุท",
   "วรรณวรีเศรษฐวิวัฒนกุล",
   "วัจนีย์สวัสดิ์พงษ์",
   "วัยวรรธน์บุณยมานพ",
   "วีราภรณ์ธารณามัย",
   "ศรีรัตน์ศุภวรรธนะกุล",
   "ศิระยาเล็กเจริญ",
   "ศิริพร ตันติวิภานุวงศ์",
   "ศุภศักดิ์ บุญประเสริฐ",
   "สุภาภรณ์ศรีสุพรรณวิทยา",
   "ธัญวรัตม์  พุฒกลั่น",
   "อังคณาแสงนภากาศ",
   "สุภานัน นันทวงษ์",
   "ปิยธิดา พิมราช",
   "รุจิราภา ปันแดง"
  ],
  "G": [
   "Transcribing error",
   "no TE",
   "1. ผิดคน",
   "2. ผิดชนิด (รายการยา)**",
   "3. ผิดขนาด/ ความแรง/ ความเข้มข้น/ ความถี่**",
   "4. ผิดเวลา (เช่น ac หรือ pc / ผิดเวลาที่กำหนด)",
   "5. อัตราเร็วในการให้ยาผิด (เร็วไป / ช้าไป)",
   "6. ระยะเวลาในการให้ยาผิด (ยาวไป / สั้นไป)",
   "7. ผิดวิถีทาง / ผิดตำแหน่ง (IV IM ID Sc SL Oral Inh)**",
   "8. รูปแบบยาผิด (Tab / Syr / Inj / Suppo / Oint / Lotion)",
   "9. คัดลอกยาขาด / เกิน / ซ้ำซ้อน",
   "10. copy ไม่ระบุชื่อ",
   "11. อื่นๆ"
  ],
  "H": [
   "Dispensing error",
   "no DE",
   "1. ผิดคน",
   "2. ผิดชนิด (รายการยา)**",
   "3. ผิดขนาด/ ความแรง/ ความเข้มข้น/ ให้ซ้ำ/ ความถี่**",
   "4. วิธีบริหารยาผิด",
   "5. ผิดเวลา (เช่น ac หรือ pc / ผิดเวลาที่กำหนด)",
   "6. ผิดวิถีทาง / ผิดตำแหน่ง (IV IM ID Sc SL Oral Inhaler)",
   "7. รูปแบบยาผิด (Tab / Syr / Inj / Suppo / Oint / Lotion)**",
   "8. ผิดเทคนิค (ผสมยาที่เข้ากันไม่ได้ บด-แบ่งยาที่ห้าม)",
   "9. ผิดจำนวน / ผิดปริมาณ",
   "10. รายการไม่ครบ / เกิน / จ่ายยาทีไม่มีคำสั่งใช้ / ซ้ำซ้อน",
   "11. ไม่ off / one day - continue",
   "12. ภาชนะบรรจุไม่เหมาะสม (เช่น ไม่ใส่ซองสีชา)",
   "13. ยาเสื่อมสภาพ (หมดอายุ / เก็บไม่เหมาะสม)",
   "14. จ่ายยาที่เกิด Drug interaction",
   "15. จ่ายยาที่ผู้ป่วยมีประวัติแพ้ยา",
   "16. อื่นๆ"
  ],
  "I": [
   "Cate ระบุกรณี TE, DE",
   "no TE, DE ไม่ต้องระบุ",
   "ไม่ทราบ",
   "B",
   "C",
   "D",
   "E",
   "F",
   "G",
   "H",
   "I"
  ],
  "J": [
   "Ward",
   "no TE, DE ไม่ต้องระบุ",
   "ไม่ทราบ",
   "อญ 1",
   "อญ 2",
   "อช 1",
   "อช 2",
   "PP",
   "LR",
   "เด็กโต",
   "SNB",
   "NICU",
   "Ortho ญ",
   "Ortho ช",
   "ศญ ",
   "ศช 1",
   "ศช 2",
   "Trauma",
   "พ 2",
   "พ 3",
   "พ 4",
   "พ 5",
   "VIP",
   "ICU-Med",
   "ICU-Surg",
   "ENT",
   "นรีเวช",
   "Stroke",
   "ICU-Med ชั้น 7",
   "ARI",
   "Anes"
  ],
  "K": [
   "AN ระบุกรณี TE, DE",
   "ไม่ทราบ"
  ],
  "L": [
   "HAD",
   "no HAD",
   "doBUTamine inj",
   "doPAmine inj",
   "KCl inj",
   "Levophed inj",
   "Morphine inj",
   "NTG inj",
   "Pethidine inj",
   "Warfarin",
   "Heparin"
  ]
 },
 "adminExtra": {
  "M": [
   "='OK ME1 IPD'!P2",
   "=B2",
   "=B3",
   "=B4"
  ],
  "O": [
   "='OK ME1 IPD'!S2",
   "Yes",
   "No"
  ],
  "P": [
   "='OK ME1 IPD'!T2",
   "Yes",
   "No"
  ],
  "Q": [
   "='OK ME1 IPD'!U2",
   "Yes",
   "No"
  ],
  "R": [
   "='OK ME1 IPD'!V2",
   "Yes",
   "No"
  ],
  "S": [
   "",
   "Yes",
   "No"
  ],
  "T": [
   "='OK ME1 IPD'!W2",
   "Yes",
   "No"
  ],
  "U": [
   "='OK ME1 IPD'!X2",
   "Yes",
   "No"
  ],
  "V": [
   "='OK ME1 IPD'!Y2",
   "Yes",
   "No"
  ],
  "W": [
   "='OK ME1 IPD'!Z2",
   "Yes",
   "No"
  ],
  "X": [
   "='OK ME1 IPD'!AA2",
   "no ME",
   "ไม่ทราบ",
   "A-B",
   "C-D",
   "E up"
  ]
 },
 "drugs": [
  "nss",
  "cef-3",
  "cef",
  "para",
  "lasix",
  "losec",
  "dexa",
  "tazocin",
  "inhalex",
  "kcl",
  "hydralazine",
  "cefa",
  "folic",
  "hydrocortisone",
  "azithro",
  "dilantin",
  "onsia",
  "omeprazole",
  "d5w",
  "ergo",
  "acetin",
  "osel",
  "clinda",
  "ors",
  "cpm",
  "manidipine",
  "keppra",
  "caco3",
  "caco",
  "ampi",
  "vanco",
  "amlo",
  "ativan",
  "insulin",
  "mgso4",
  "mgso",
  "tramol",
  "transamine",
  "celebrex",
  "bco",
  "mero",
  "atorvas",
  "meropenem",
  "cefazolin",
  "penfill",
  "norflex",
  "mom",
  "tazo",
  "swi",
  "asa",
  "metro",
  "sitaflox",
  "naproxen",
  "chart",
  "carvedilol",
  "asa81",
  "plavix",
  "levophed",
  "senokot",
  "genta",
  "budesonide",
  "dom",
  "prednisolone",
  "fentanyl",
  "ichart",
  "amoxy",
  "losec20",
  "rate",
  "cef3",
  "bactrim",
  "gabapentin",
  "quetiapine",
  "plasil",
  "trazodone",
  "elixir",
  "cefotaxime",
  "order",
  "add",
  "nacl",
  "metoprolol",
  "methylpred",
  "nss100",
  "swi10",
  "para500",
  "lasix500",
  "fluimucil",
  "topiramate",
  "warfarin",
  "albumin",
  "vit",
  "cefixime",
  "doxa",
  "cream",
  "lotion",
  "avamys",
  "pred",
  "doxy",
  "brom",
  "azathiopine",
  "azithromycin",
  "acetar",
  "flagyl",
  "augmentin",
  "filgrastim",
  "fortum",
  "tigecycline",
  "aldactone",
  "losartan",
  "cipro",
  "heavy",
  "kalimate",
  "rhina",
  "alc",
  "buscopan",
  "carvedilol6",
  "singulair",
  "erlotinib",
  "ome",
  "patch",
  "cpr",
  "e20",
  "amlo10",
  "mtv",
  "biapenem",
  "com",
  "biosulin",
  "risperidone",
  "ccu",
  "aripiprazole",
  "isdn",
  "mfm",
  "milk",
  "ampicillin",
  "melcam",
  "simvas",
  "sitafloxacin",
  "filgrastrim",
  "domperidone",
  "tige",
  "heparin",
  "clozapine",
  "fbc",
  "dimen",
  "moxiflox",
  "sulbactam",
  "fer",
  "acyclovir",
  "apixaban",
  "seretide",
  "evo",
  "olazapine",
  "onco",
  "lansoprazole",
  "brufen",
  "spironolactone",
  "amikin",
  "lorazepam",
  "case2",
  "case",
  "controloc",
  "air",
  "hcq",
  "mucillin",
  "mavang",
  "xylocaine",
  "silverol",
  "lactulose",
  "sticker",
  "berodual",
  "clida",
  "bisoprolol",
  "fosfo",
  "levofloxacin",
  "manitol",
  "ivermectin",
  "nss1000",
  "ketorolac",
  "key1",
  "key",
  "sal",
  "alfacalcidol",
  "piogli",
  "cordarine",
  "swi10ml",
  "budesonid",
  "quatiepine",
  "enaril",
  "gpo",
  "tranxene",
  "gemfibrozil600",
  "gemfibrozil",
  "fluoxetine",
  "sulam",
  "mydrocortisone",
  "theophylline",
  "cephalexin",
  "respule",
  "ceftazidime",
  "viga",
  "cal",
  "methylprednisolone",
  "cef4",
  "haldol",
  "mgcl2",
  "mgcl",
  "laxic",
  "abc",
  "vopar",
  "bicabonate",
  "key3",
  "senna",
  "loperamide",
  "tizanidine",
  "marcaine",
  "vopar250",
  "seretide25",
  "cordarone",
  "bromhex",
  "sodamont",
  "tranxamine",
  "doxazocin",
  "lasix40inj",
  "gaba",
  "atenolol50",
  "atenolol",
  "cloxa",
  "deax",
  "wafarin",
  "sr10",
  "insulin70",
  "transmine",
  "etham",
  "isordil",
  "water10ml",
  "sodiumbicarb",
  "glucose",
  "spirava",
  "sulbutamol",
  "ceftaz",
  "ampho",
  "clobet",
  "lasix40",
  "clexane0",
  "clexane",
  "mmf",
  "cephalexine",
  "ponstan",
  "apixan",
  "enoxa",
  "trauma",
  "ntg",
  "wcu",
  "pyridostigmine",
  "propranolol",
  "urea",
  "nss1000ml",
  "metopolol",
  "sim20",
  "sim",
  "colchicine",
  "omnicef",
  "metformin",
  "clona",
  "albendazole",
  "scale",
  "calcitriol",
  "eent",
  "wster",
  "atrovas",
  "nefopam",
  "mero1g",
  "amlo5",
  "drop",
  "norp",
  "vopa",
  "ketoconazole",
  "fosfomycin",
  "granule",
  "lercarnidipine",
  "clotrimazole",
  "pip",
  "taz",
  "sodium",
  "valproate",
  "espogen",
  "hypercrit",
  "cmt",
  "ergocalciferol",
  "smofperipheral",
  "pgs",
  "mero1",
  "bilastine",
  "lrs",
  "ciprol",
  "clindamycin",
  "pantoprazole",
  "d5nss",
  "linezolid",
  "dextromethophan",
  "hydrocor",
  "primalut",
  "glycerine",
  "ergol",
  "betahistine6mg",
  "mag",
  "phod",
  "hydrocorti",
  "panto",
  "thyroxine",
  "actrpid",
  "seroquel",
  "levoflox",
  "ventolin",
  "cc---15",
  "bactoban",
  "meto",
  "carbetocin",
  "pph",
  "azitho",
  "fluid",
  "sertraline50",
  "sertraline",
  "calcium",
  "unj",
  "manidipne20",
  "manidipne",
  "sitafox",
  "viscous",
  "amitrip",
  "norep",
  "levo",
  "dopamine",
  "dopa",
  "adrenaline",
  "empagliflozin",
  "alfuzosin",
  "alpha",
  "hemax",
  "allopurinol",
  "metroprolol",
  "luid",
  "etoricoxib",
  "elixer",
  "omepazole",
  "amiodarone",
  "d5s",
  "glu",
  "mrc",
  "box",
  "vitk",
  "albu",
  "osel75",
  "hydro",
  "allo",
  "airx",
  "manidipin",
  "quetiapine25",
  "remde",
  "ezetimibe",
  "fluorometholone",
  "chloram",
  "ceftriaxone",
  "paracetamol",
  "methylpre",
  "ppn",
  "prnkey",
  "dulpicate",
  "artane",
  "cypohep",
  "glycerin",
  "ccbid",
  "lora1",
  "lora",
  "ampi500",
  "penfll",
  "augmenitn",
  "azathoprine",
  "acetazolamide",
  "brufen400",
  "transamin",
  "norepinephrine",
  "alfuzocin",
  "pioglitazone",
  "then",
  "ceftri",
  "gluconate",
  "bupivacaine",
  "pza",
  "risper",
  "hydra",
  "rifam",
  "escitalopram",
  "clonazepam",
  "alummilke",
  "berclomine",
  "voriconazole",
  "clopidogrel",
  "vanvo",
  "irrigate",
  "doxycycline",
  "macarin",
  "d5w100",
  "insulatard",
  "unasyn",
  "air-x",
  "vitalapid",
  "phenobarbital",
  "extem",
  "tue",
  "fri",
  "ciproflox",
  "hadol",
  "moduretic",
  "sodamint",
  "isobaric",
  "tyroxine",
  "jelly",
  "avexa",
  "actrapid",
  "synto",
  "julie",
  "tolperisone",
  "lasix250",
  "nss50",
  "nss100ml",
  "alum",
  "calcitonin",
  "norflox",
  "donepezil",
  "colistin",
  "ibuprofen",
  "lerca",
  "mmi",
  "q2d",
  "spiriva",
  "amof",
  "hyoscine",
  "glagine",
  "arcoxia",
  "smof",
  "ris",
  "aug",
  "prrn",
  "inh",
  "alfa",
  "lercar",
  "oral",
  "heprin",
  "ccukey",
  "prevacid",
  "ooff",
  "warter",
  "med",
  "fluco",
  "glaritus",
  "labeterol"
 ]
};
