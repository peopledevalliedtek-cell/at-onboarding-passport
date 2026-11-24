/**
 * AT Digital Onboarding Passport — backend
 * เก็บข้อมูลทั้งหมดไว้ใน Google Sheet ของคุณเอง ผ่าน Google Apps Script Web App
 *
 * ตั้งค่าครั้งแรก (ดู README.md ของ repo สำหรับขั้นตอนละเอียด):
 * 1. สร้าง Google Sheet ใหม่ (ชื่ออะไรก็ได้)
 * 2. เมนู Extensions > Apps Script แล้ววางไฟล์นี้ทั้งหมด (ลบโค้ดเดิมทิ้งก่อน)
 * 3. รันฟังก์ชัน setupSheets once (เลือกจาก dropdown ด้านบน แล้วกด Run) เพื่อสร้างชีตทั้งหมด
 *    พร้อมหัวตารางภาษาไทยและจัดรูปแบบให้สวยงาม — อนุญาตสิทธิ์ตามที่ขอ
 * 4. Deploy > New deployment > Web app > Execute as: Me, Who has access: Anyone
 * 5. คัดลอก Web app URL ไปวางในหน้า "ตั้งค่าฐานข้อมูล" ของเว็บแอป
 *
 * โครงสร้างข้อมูล: ใช้หลักการ "JSON ทั้งก้อนคือฐานข้อมูลจริง" เก็บไว้ในคอลัมน์สุดท้าย
 * ของชีต "พนักงาน" และ "Template การ Onboarding" (แก้ไขข้อมูลผ่านหน้าเว็บเท่านั้น
 * ไม่ควรแก้ในคอลัมน์ JSON โดยตรง) ส่วนชีตอื่นๆ ทั้งหมดเป็น "ตารางสรุปภาษาไทย" ที่ระบบ
 * สร้าง/อัปเดตให้อัตโนมัติทุกครั้งที่มีการบันทึกพนักงาน สำหรับเปิดดู/ทำรายงานได้โดยตรง
 */

var SHEETS = {
  SETTINGS: 'ตั้งค่าระบบ',
  TEMPLATES: 'Template การ Onboarding',
  EMPLOYEES: 'พนักงาน',
  CHECKLIST: 'Checklist Onboarding',
  KPI: 'ตัวชี้วัด KPI ทดลองงาน',
  REFLECTION: 'แบบสะท้อน 2 ทาง',
  USERS: 'ผู้ใช้งานระบบ'
};

var BRAND_RED = '#c0272d';
var BRAND_DARK = '#3a1012';

/* ========================= doGet / doPost ========================= */

function doGet(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) || 'list';
    if (action === 'list') {
      return jsonResponse_({
        ok: true,
        settings: readSettings_(),
        templates: readTemplates_(),
        employees: readEmployees_(),
        users: readUsers_()
      });
    }
    return jsonResponse_({ ok: false, error: 'unknown action: ' + action });
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var action = body.action;
    if (action === 'setSettings') { writeSettings_(body.data); return jsonResponse_({ ok: true }); }
    if (action === 'setUsers') { writeUsers_(body.data); return jsonResponse_({ ok: true }); }
    if (action === 'setTemplate') { upsertTemplate_(body.id, body.data); return jsonResponse_({ ok: true }); }
    if (action === 'deleteTemplate') { deleteTemplate_(body.id); return jsonResponse_({ ok: true }); }
    if (action === 'setEmployee') { upsertEmployee_(body.id, body.data); return jsonResponse_({ ok: true }); }
    if (action === 'deleteEmployee') { deleteEmployee_(body.id); return jsonResponse_({ ok: true }); }
    if (action === 'uploadPhoto') {
      var photo = uploadEmployeePhoto_(body.id, body.filename, body.mimeType, body.base64);
      return jsonResponse_({ ok: true, photoUrl: photo.photoUrl });
    }
    return jsonResponse_({ ok: false, error: 'unknown action: ' + action });
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err) });
  }
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ========================= one-time setup / formatting ========================= */

/** Run this once manually from the Apps Script editor after pasting the code in. */
function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet_(ss, SHEETS.SETTINGS, ['หัวข้อ', 'ค่า']);
  ensureSheet_(ss, SHEETS.TEMPLATES, [
    'รหัส Template', 'ชื่อ Template', 'แผนก / สายงาน', 'ระดับตำแหน่ง',
    'ระยะเวลาทดลองงานเริ่มต้น (วัน)', 'จำนวนรายการ Checklist', 'จำนวนตัวชี้วัด KPI',
    'อัปเดตล่าสุด', 'ข้อมูลเต็ม (JSON — ใช้งานในระบบ ไม่ต้องแก้ตรงนี้)'
  ]);
  ensureSheet_(ss, SHEETS.EMPLOYEES, [
    'รหัสอ้างอิง', 'รหัสพนักงาน', 'ชื่อ-นามสกุล', 'ชื่อเล่น', 'ตำแหน่ง', 'แผนก / ทีม',
    'วันเริ่มงาน', 'ระยะเวลาทดลองงาน (วัน)', 'วันครบกำหนด', 'ความคืบหน้า (%)', 'สถานะ',
    'Template ที่ใช้', 'พี่เลี้ยง', 'หัวหน้าโดยตรง', 'HR ผู้ดูแล', 'รูปโปรไฟล์ (ลิงก์)', 'อัปเดตล่าสุด',
    'ข้อมูลเต็ม (JSON — ใช้งานในระบบ ไม่ต้องแก้ตรงนี้)'
  ]);
  ensureSheet_(ss, SHEETS.CHECKLIST, [
    'รหัสพนักงาน', 'ชื่อ-นามสกุล', 'เฟส', 'รายการตรวจสอบ', 'สถานะ',
    'บันทึกปัญหา / รายละเอียดเพิ่มเติม', 'วันที่ดำเนินการ'
  ]);
  ensureSheet_(ss, SHEETS.KPI, [
    'รหัสพนักงาน', 'ชื่อ-นามสกุล', 'ตัวชี้วัด', 'เป้าหมายช่วงทดลองงาน', 'ดูจากที่ไหน', 'บันทึกเพิ่มเติม'
  ]);
  ensureSheet_(ss, SHEETS.REFLECTION, [
    'รหัสพนักงาน', 'ชื่อ-นามสกุล', 'รอบ', 'ผู้ตอบ', 'ข้อที่', 'คำถาม', 'คำตอบ', 'ขยายความ', 'วันที่ตอบ'
  ]);
  ensureSheet_(ss, SHEETS.USERS, ['บทบาท', 'ชื่อผู้ใช้งาน', 'รหัสผ่าน']);

  // seed default settings + users + one starter template if sheets are empty
  if (readSettings_().length === 0 || Object.keys(flattenSettings_(readSettings_())).length === 0) {
    writeSettings_({
      'ชื่อระบบ': 'AT Digital Onboarding Passport',
      'หน่วยงาน': 'People & Organization Management · ทีม Onboarding / People Development',
      'บริษัท 1': 'บริษัท อัลลายด์เทค (ประเทศไทย) จำกัด',
      'บริษัท 2': 'บริษัท อัลลายด์เทคเอ็นจิเนียริ่ง จำกัด'
    });
  }
  if (readUsers_().length === 0) {
    writeUsers_([
      { role: 'hr', name: 'ฝ่าย People & Organization Management', pin: 'hr2026' },
      { role: 'supervisor', name: 'หัวหน้างาน / พี่เลี้ยง', pin: 'lead2026' }
    ]);
  }
  if (readTemplates_().length === 0) {
    seedStarterTemplates_();
  }

  prettifyAllSheets_(ss);
  // Put the most useful sheets first for anyone opening the spreadsheet directly.
  ss.setActiveSheet(ss.getSheetByName(SHEETS.EMPLOYEES));
  SpreadsheetApp.flush();
}

function genId_(prefix) {
  return prefix + '_' + Utilities.getUuid().slice(0, 8);
}

function defaultChecklist_() {
  return {
    phase1: { title: 'Phase 1 · ก่อนวันเริ่มงาน (Pre-Onboarding)', items: [
      'จัดทำสัญญาจ้างและเตรียมเอกสารสมัครงานให้ครบถ้วน',
      'แจ้งฝ่าย IT จัดเตรียมคอมพิวเตอร์ อีเมล และสิทธิ์เข้าใช้งานระบบ (ถ้ามี)',
      'จัดเตรียมโต๊ะทำงานและอุปกรณ์สำนักงานพื้นฐาน (ถ้ามี)',
      'ตรวจสอบข้อมูลส่วนบุคคลในระบบ HR ถูกต้องครบถ้วน',
      'โทรคอนเฟิร์มวันเริ่มงาน เวลา สถานที่ ผู้รับการติดต่อ ทบทวนการแต่งกาย',
      'แจ้งต้นสังกัด / พี่เลี้ยง เพื่อเตรียมดูแลพนักงานใหม่'
    ]},
    phase2: { title: 'Phase 2 · เริ่มต้นการทำงาน (Day 1–7)', items: [
      'ปฐมนิเทศกฎระเบียบ สวัสดิการ วัฒนธรรมองค์กร และความปลอดภัย',
      'เซ็นสัญญาจ้าง',
      'แจ้งรายละเอียดหน้าที่การทำงานตาม Job Description',
      'พนักงานตอบได้ว่าใครคือพี่เลี้ยง และใครคือหัวหน้าโดยตรง',
      'พาเดินแนะนำตัวกับหัวหน้างานและพี่เลี้ยงรวมถึงทีมที่เกี่ยวข้อง',
      'ส่งมอบพนักงานให้พี่เลี้ยง (Buddy) ดูแลเรื่องการทานอาหารกลางวัน',
      'ฝ่าย IT แนะนำการเข้าถึงระบบไอทีของบริษัท (ถ้ามี)',
      'พาทัวร์อาคารสถานที่บริษัท'
    ]},
    phase3: { title: 'Phase 3 · ติดตามและประเมินผล (วันที่ 30–60)', items: [
      'ตรวจสอบการใช้งานระบบงานต่างๆ เช่น เข้า-ออกงาน / การลา / IT',
      'สอบถามความเป็นอยู่ สายบังคับบัญชา เพื่อนสนิทหรือพี่เลี้ยงคือใคร',
      'ทบทวนหัวข้อที่ถูกสอนงานจากพี่เลี้ยง / หัวหน้างาน',
      'ติดตามผลการปฏิบัติงาน ทำงานอะไรมาแล้วบ้าง',
      'สอบถามปัญหาในแผนก / ความช่วยเหลือที่ต้องการ',
      'สอบถามว่ามีคำพูดหรือการกระทำที่ทำให้ไม่สบายใจหรือไม่ มีอะไรอยากให้ HR รู้',
      'ทบทวนหัวข้อปฐมนิเทศ เช่น ผังองค์กร สายบังคับบัญชา',
      'ทบทวนขั้นตอนแจ้งลาป่วยฉุกเฉิน: แจ้งใคร ทางไหน ก่อนกี่โมง',
      'ตรวจสอบวันลาหยุดงาน/เหตุจำเป็นเพิ่มเติม (ถ้ามี)'
    ]},
    phase4: { title: 'Phase 4 · ติดตามและประเมินผลหลังผ่านทดลองงาน (วันที่ 90–สิ้นสุดระยะ)', items: [
      'ตรวจสอบการสอนงานจากเอกสาร OJT',
      'สัมภาษณ์เรื่องมาตรฐาน JD / KPI / WI ของตนเอง',
      'ทบทวนวัฒนธรรมองค์กรและ Core Value ว่ามีอะไรบ้าง',
      'ยกตัวอย่างการเห็น Core Value จากคนในแผนกทำจริง 1 เรื่อง',
      'ทบทวนกิจกรรมที่ผ่านมา ดีหรือไม่ดีอย่างไร',
      'แจ้งสวัสดิการที่จะได้รับหลังผ่านทดลองงาน',
      'สอบถามปัญหาในแผนก / ความช่วยเหลือ / การสนับสนุนที่ต้องการ'
    ]}
  };
}
function defaultKpi_() {
  return [
    { metric: 'ทักษะทางเทคนิคและการปฏิบัติงานจริงตามตำแหน่งงาน', target: '100%', source: 'ใบตรวจรับงาน / การสุ่มตรวจหน้างาน' },
    { metric: 'ความรู้และการวิเคราะห์แก้ปัญหาหน้างาน', target: '> 70%', source: 'รายงานการแก้ไขปัญหา / แบบทดสอบความรู้' },
    { metric: 'การวางแผนและควบคุมงานให้แล้วเสร็จตามกำหนด', target: 'ภายในวันสุดท้ายของทุกสัปดาห์', source: 'แผนงานส่วนบุคคล / รายงานสรุปผลประจำสัปดาห์' },
    { metric: 'การสื่อสาร รายงานความคืบหน้า และทำงานร่วมกับทีม', target: 'ภายในวันสุดท้ายของทุกสัปดาห์', source: 'บันทึกรายงานในกลุ่มสื่อสาร / บันทึกการประชุม' }
  ];
}
function defaultReflectionRounds_() {
  return [
    { round: 1, day: 7, label: 'รอบ 1 · วันที่ 7' },
    { round: 2, day: 30, label: 'รอบ 2 · วันที่ 30' },
    { round: 3, day: 90, label: 'รอบ 3 · วันที่ 90' }
  ];
}
function seedStarterTemplates_() {
  var t1 = {
    id: genId_('tpl'), name: 'ช่างเทคนิค / ปฏิบัติการ (มาตรฐาน)', department: 'Asset Reliability / Operations',
    positionLevel: 'ระดับช่าง / ปฏิบัติการ', durationDays: 119,
    checklist: defaultChecklist_(), kpi: defaultKpi_(), reflectionRounds: defaultReflectionRounds_()
  };
  var t2 = {
    id: genId_('tpl'), name: 'สำนักงาน / เจ้าหน้าที่ (มาตรฐาน)', department: 'สำนักงานทั่วไป',
    positionLevel: 'ระดับเจ้าหน้าที่ / Officer', durationDays: 90,
    checklist: defaultChecklist_(), kpi: defaultKpi_(), reflectionRounds: defaultReflectionRounds_()
  };
  upsertTemplate_(t1.id, t1);
  upsertTemplate_(t2.id, t2);
}

function ensureSheet_(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  var existingHeaders = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  var blank = existingHeaders.every(function (v) { return v === '' || v === null; });
  if (blank) sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  return sheet;
}

function prettifyAllSheets_(ss) {
  Object.keys(SHEETS).forEach(function (key) {
    prettifySheet_(ss.getSheetByName(SHEETS[key]));
  });
  // Default sheet created by Google Sheets ("Sheet1") — hide it if still empty & unused.
  var sheet1 = ss.getSheetByName('Sheet1') || ss.getSheetByName('ชีต1');
  if (sheet1 && sheet1.getLastRow() === 0) {
    try { ss.deleteSheet(sheet1); } catch (e) { /* only sheet left, ignore */ }
  }
}

function prettifySheet_(sheet) {
  if (!sheet) return;
  var lastCol = sheet.getLastColumn() || sheet.getMaxColumns();
  var headerRange = sheet.getRange(1, 1, 1, lastCol);
  headerRange
    .setBackground(BRAND_RED)
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(10.5)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center')
    .setWrap(true);
  sheet.setFrozenRows(1);
  sheet.setRowHeight(1, 34);
  sheet.setTabColor(BRAND_RED);

  // Reasonable column widths: narrower for short id/date columns, wider for names/JSON.
  for (var c = 1; c <= lastCol; c++) {
    var header = String(sheet.getRange(1, c).getValue());
    var width = 150;
    if (/JSON/i.test(header)) width = 90;
    else if (/^(ชื่อ|รายการ|ตัวชี้วัด|คำถาม|บันทึก|ขยายความ|ดูจากที่ไหน|หน่วยงาน|บริษัท|ค่า)/.test(header)) width = 240;
    else if (/(วัน|รหัส|สถานะ|ข้อที่|ผู้ตอบ|รอบ)/.test(header)) width = 110;
    sheet.setColumnWidth(c, width);
  }

  // Alternating row banding over a generous range so new rows inherit it.
  var maxRows = Math.max(sheet.getMaxRows(), 200);
  var bandRange = sheet.getRange(2, 1, maxRows - 1, lastCol);
  try {
    var existingBandings = sheet.getBandings();
    existingBandings.forEach(function (b) { b.remove(); });
    bandRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);
  } catch (e) { /* banding can fail on a 0-row range; safe to ignore */ }

  sheet.setHiddenGridlines(false);
}

/* ========================= settings ========================= */

function readSettings_() {
  var sheet = sheetByName_(SHEETS.SETTINGS);
  var values = sheet.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < values.length; i++) {
    if (!values[i][0]) continue;
    out.push({ key: String(values[i][0]), value: values[i][1] });
  }
  return out;
}
function flattenSettings_(rows) {
  var o = {};
  rows.forEach(function (r) { o[r.key] = r.value; });
  return o;
}
function writeSettings_(dataObj) {
  var sheet = sheetByName_(SHEETS.SETTINGS);
  sheet.getRange(2, 1, Math.max(sheet.getMaxRows() - 1, 1), 2).clearContent();
  var rows = Object.keys(dataObj).map(function (k) { return [k, dataObj[k]]; });
  if (rows.length) sheet.getRange(2, 1, rows.length, 2).setValues(rows);
}

/* ========================= users ========================= */

function readUsers_() {
  var sheet = sheetByName_(SHEETS.USERS);
  var values = sheet.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < values.length; i++) {
    if (!values[i][0]) continue;
    out.push({ role: String(values[i][0]), name: values[i][1], pin: values[i][2] });
  }
  return out;
}
function writeUsers_(list) {
  var sheet = sheetByName_(SHEETS.USERS);
  sheet.getRange(2, 1, Math.max(sheet.getMaxRows() - 1, 1), 3).clearContent();
  var rows = list.map(function (u) { return [u.role, u.name, u.pin]; });
  if (rows.length) sheet.getRange(2, 1, rows.length, 3).setValues(rows);
}

/* ========================= templates ========================= */

function readTemplates_() {
  var sheet = sheetByName_(SHEETS.TEMPLATES);
  var values = sheet.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < values.length; i++) {
    var json = values[i][8];
    if (!json) continue;
    try { out.push(JSON.parse(json)); } catch (e) { /* skip bad row */ }
  }
  return out;
}

function upsertTemplate_(id, data) {
  var sheet = sheetByName_(SHEETS.TEMPLATES);
  var row = findRowByFirstCol_(sheet, id);
  var itemCount = 0;
  Object.keys(data.checklist || {}).forEach(function (k) { itemCount += (data.checklist[k].items || []).length; });
  var values = [
    id, data.name || '', data.department || '', data.positionLevel || '',
    data.durationDays || '', itemCount, (data.kpi || []).length,
    new Date().toISOString(), JSON.stringify(data)
  ];
  if (row === -1) sheet.appendRow(values);
  else sheet.getRange(row, 1, 1, values.length).setValues([values]);
}

function deleteTemplate_(id) {
  var sheet = sheetByName_(SHEETS.TEMPLATES);
  var row = findRowByFirstCol_(sheet, id);
  if (row !== -1) sheet.deleteRow(row);
}

/* ========================= employees ========================= */

function readEmployees_() {
  var sheet = sheetByName_(SHEETS.EMPLOYEES);
  var values = sheet.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < values.length; i++) {
    var json = values[i][17];
    if (!json) continue;
    try { out.push(JSON.parse(json)); } catch (e) { /* skip bad row */ }
  }
  return out;
}

function upsertEmployee_(id, data) {
  var sheet = sheetByName_(SHEETS.EMPLOYEES);
  var row = findRowByFirstCol_(sheet, id);
  var progress = employeeProgress_(data);
  var endDate = addDays_(data.startDate, data.durationDays);
  var values = [
    id, data.employeeCode || '', data.fullName || '', data.nickname || '', data.position || '', data.department || '',
    data.startDate || '', data.durationDays || '', endDate, progress, statusLabel_(data.status),
    templateNameFor_(data.templateId), contactLabel_(data.mentor), contactLabel_(data.supervisor), contactLabel_(data.hrContact),
    data.photoUrl || '', new Date().toISOString(), JSON.stringify(data)
  ];
  if (row === -1) sheet.appendRow(values);
  else sheet.getRange(row, 1, 1, values.length).setValues([values]);

  rewriteChecklistSummary_(id, data);
  rewriteKpiSummary_(id, data);
  rewriteReflectionSummary_(id, data);
}

function deleteEmployee_(id) {
  deleteRowsMatchingFirstCol_(sheetByName_(SHEETS.EMPLOYEES), id);
  deleteRowsMatchingFirstCol_(sheetByName_(SHEETS.CHECKLIST), id);
  deleteRowsMatchingFirstCol_(sheetByName_(SHEETS.KPI), id);
  deleteRowsMatchingFirstCol_(sheetByName_(SHEETS.REFLECTION), id);
}

/* ---- profile photo upload (stored in Google Drive, linked onto the employee record) ---- */

var PHOTO_DRIVE_ROOT_FOLDER_NAME = 'AT Onboarding Passport - Profile Photos';

function getEmployeeJson_(id) {
  var sheet = sheetByName_(SHEETS.EMPLOYEES);
  var row = findRowByFirstCol_(sheet, id);
  if (row === -1) return null;
  var json = sheet.getRange(row, 18).getValue();
  try { return JSON.parse(json); } catch (e) { return null; }
}

function getOrCreateRootFolder_(name) {
  var it = DriveApp.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(name);
}

function uploadEmployeePhoto_(id, filename, mimeType, base64) {
  if (!id) throw new Error('missing id');
  if (!base64) throw new Error('missing image data');
  var folder = getOrCreateRootFolder_(PHOTO_DRIVE_ROOT_FOLDER_NAME);

  // Replace any previous photo for this employee so the folder doesn't pile up old copies.
  var prefix = 'profile_' + id;
  var existing = folder.getFiles();
  while (existing.hasNext()) {
    var f = existing.next();
    if (f.getName().indexOf(prefix) === 0) f.setTrashed(true);
  }

  var bytes = Utilities.base64Decode(base64);
  var ext = (mimeType && mimeType.indexOf('png') !== -1) ? '.png' : '.jpg';
  var blob = Utilities.newBlob(bytes, mimeType || 'image/jpeg', prefix + ext);
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  var photoUrl = 'https://lh3.googleusercontent.com/d/' + file.getId() + '=s400';
  var data = getEmployeeJson_(id) || {};
  data.photoUrl = photoUrl;
  upsertEmployee_(id, data);
  return { photoUrl: photoUrl };
}

function templateNameFor_(templateId) {
  var tpl = readTemplates_().filter(function (t) { return t.id === templateId; })[0];
  return tpl ? tpl.name : '';
}
function contactLabel_(c) {
  if (!c || !c.name) return '';
  return c.name + (c.position ? ' (' + c.position + ')' : '') + (c.phone ? ' · ' + c.phone : '');
}
function statusLabel_(s) {
  return { active: 'กำลังติดตาม', done: 'ครบกำหนดแล้ว', offboarded: 'พ้นสภาพ' }[s] || (s || 'กำลังติดตาม');
}
function employeeProgress_(emp) {
  var total = 0, done = 0;
  var cl = emp.checklist || {};
  Object.keys(cl).forEach(function (k) {
    (cl[k] || []).forEach(function (item) { total++; if (item.done) done++; });
  });
  return total ? Math.round(done / total * 100) : 0;
}
function addDays_(dateStr, n) {
  if (!dateStr) return '';
  var d = new Date(dateStr + 'T00:00:00');
  d.setDate(d.getDate() + (parseInt(n, 10) || 0));
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

/* ---- flattened Thai summary tables, regenerated on every employee save ---- */

function rewriteChecklistSummary_(id, data) {
  var sheet = sheetByName_(SHEETS.CHECKLIST);
  deleteRowsMatchingFirstCol_(sheet, id);
  var rows = [];
  var cl = data.checklist || {};
  Object.keys(cl).forEach(function (phaseKey) {
    (cl[phaseKey] || []).forEach(function (item) {
      rows.push([
        id, data.fullName || '', phaseLabel_(phaseKey), item.label || '',
        item.done ? 'เสร็จแล้ว' : 'ยังไม่เสร็จ', item.note || '', item.date || ''
      ]);
    });
  });
  if (rows.length) sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 7).setValues(rows);
}
function phaseLabel_(key) {
  return { phase1: 'เฟส 1 · ก่อนเริ่มงาน', phase2: 'เฟส 2 · วันที่ 1-7', phase3: 'เฟส 3 · วันที่ 30-60', phase4: 'เฟส 4 · วันที่ 90+' }[key] || key;
}

function rewriteKpiSummary_(id, data) {
  var sheet = sheetByName_(SHEETS.KPI);
  deleteRowsMatchingFirstCol_(sheet, id);
  var rows = (data.kpi || []).map(function (k) {
    return [id, data.fullName || '', k.metric || '', k.target || '', k.source || '', k.notes || ''];
  });
  if (rows.length) sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 6).setValues(rows);
}

function rewriteReflectionSummary_(id, data) {
  var sheet = sheetByName_(SHEETS.REFLECTION);
  deleteRowsMatchingFirstCol_(sheet, id);
  var rows = [];
  var reflections = data.reflections || {};
  Object.keys(reflections).forEach(function (round) {
    var rec = reflections[round] || {};
    ['employee', 'mentor'].forEach(function (who) {
      var answers = rec[who];
      if (!answers) return;
      var whoLabel = who === 'employee' ? 'พนักงานใหม่' : 'พี่เลี้ยง/หัวหน้างาน';
      var extra = rec[who + 'Extra'] || '';
      var date = rec[who + 'Date'] || '';
      answers.forEach(function (a, i) {
        rows.push([id, data.fullName || '', 'รอบ ' + round, whoLabel, i + 1, a.q || '', answerLabel_(a.answer), a.note || '', date]);
      });
      if (extra) rows.push([id, data.fullName || '', 'รอบ ' + round, whoLabel, '-', '(ข้อเสนอแนะเพิ่มเติม)', '', extra, date]);
    });
  });
  if (rows.length) sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 9).setValues(rows);
}
function answerLabel_(a) {
  return { yes: 'ใช่', partial: 'บางส่วน', no: 'ไม่ใช่' }[a] || (a || '');
}

/* ========================= shared helpers ========================= */

function sheetByName_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) { setupSheets(); sheet = ss.getSheetByName(name); }
  return sheet;
}
function findRowByFirstCol_(sheet, id) {
  var values = sheet.getRange(1, 1, sheet.getLastRow() || 1, 1).getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) return i + 1;
  }
  return -1;
}
function deleteRowsMatchingFirstCol_(sheet, id) {
  var values = sheet.getRange(1, 1, sheet.getLastRow() || 1, 1).getValues();
  for (var i = values.length - 1; i >= 1; i--) {
    if (String(values[i][0]) === String(id)) sheet.deleteRow(i + 1);
  }
}
