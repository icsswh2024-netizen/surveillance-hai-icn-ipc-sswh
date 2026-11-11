/**
 * แบบบันทึกผู้ป่วย HAI + แก้ไข Dropdown — Google Apps Script (ไฟล์เดียวคุมทั้งหมด)
 * ผูกกับไฟล์ชีตข้อมูลหลัก: 1KEK2bvXLDAY7awAapTnF32yZz1qvfqfWXje2ccG2pLc
 *
 * ทำ 2 อย่างในไฟล์เดียว (เพราะอยู่ชีตเดียวกัน):
 *   1) action = "addCase" → เขียนเคส (1+ บรรทัด) ลงแท็บ "ทะเบียน-1"  (เมนู แบบบันทึกผู้ป่วย HAI)
 *   2) action = "setList" → เขียนค่าตัวเลือกลงคอลัมน์ในแท็บ "List-ทะเบียน-1" (เมนู แก้ไข Dropdown)
 *
 * ติดตั้ง (ครั้งเดียว ใช้ URL /exec เดียวกับทั้งสองเมนู):
 *   1) เปิดไฟล์ชีต → Extensions → Apps Script → วางโค้ดนี้ → บันทึก
 *   2) Deploy → New deployment → Web app · Execute as = Me · Who has access = Anyone → คัดลอก /exec
 *      (ถ้าเคย Deploy แล้ว ใช้ Manage deployments → Edit → New version)
 *   3) วาง /exec ลงช่อง URL ในเมนู "แบบบันทึกผู้ป่วย HAI" หรือ "แก้ไข Dropdown" (ใช้ร่วมกัน)
 *
 * วิธีจับคู่คอลัมน์ (addCase): ใช้แถวหัวตารางของ "ทะเบียน-1" จับคู่กับคีย์ที่ส่งมา
 * แบบไม่สนตัวพิมพ์/ช่องว่าง/จุด (เช่น "วันนอน" = "วันนนอน", "op" = "OP")
 */

const REC_SHEET = 'ทะเบียน-1-2570';  // แท็บเคสปีงบปัจจุบัน (เปลี่ยนชื่อทุกต้นปีงบ เช่น ปีถัดไป → 'ทะเบียน-1-2571')

function recJson_(obj, cb) {
  const out = JSON.stringify(obj);
  if (cb) return ContentService.createTextOutput(cb + '(' + out + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(out).setMimeType(ContentService.MimeType.JSON);
}
function recNorm_(s) { return String(s == null ? '' : s).toLowerCase().replace(/[\s.]/g, ''); }

// ===== ยืนยันผลสงสัยติดเชื้อ (ICWN) จากหน้าเว็บ → เขียนกลับไฟล์ฟอร์ม ICWN =====
var ICN_FILE_ID = '1HU80TyjUXsvyiiBP8gI41qmTfqaLV2k72NbWJ6lh8FA';
var ICN_SHEET = 'การตอบแบบฟอร์ม 1';
function icnHn_(v) { return String(v == null ? '' : v).replace(/\D/g, '').replace(/^0+/, ''); }
function icnVerify_(p) {
  try {
    var sh = SpreadsheetApp.openById(ICN_FILE_ID).getSheetByName(ICN_SHEET);
    if (!sh) return recJson_({ ok: false, note: 'ไม่พบแท็บฟอร์ม ICWN "' + ICN_SHEET + '"' });
    var lastCol = sh.getLastColumn();
    var head = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (x) { return String(x).trim(); });
    var cRes = head.indexOf('ผลการวินิจฉัย'), cHn = head.indexOf('HN');
    if (cRes < 0) return recJson_({ ok: false, note: 'ไม่พบคอลัมน์ "ผลการวินิจฉัย"' });
    // คอลัมน์ audit (สร้างต่อท้ายถ้ายังไม่มี)
    var cBy = head.indexOf('ยืนยันโดย'); if (cBy < 0) { cBy = lastCol; sh.getRange(1, cBy + 1).setValue('ยืนยันโดย'); lastCol++; }
    var cAt = head.indexOf('เวลายืนยัน'); if (cAt < 0) { cAt = lastCol; sh.getRange(1, cAt + 1).setValue('เวลายืนยัน'); lastCol++; }
    var hn = icnHn_(p.hn), rowNum = parseInt(p.row, 10), ok = false;
    var lastRow = sh.getLastRow();
    function hnAt(r) { return cHn < 0 ? '' : icnHn_(sh.getRange(r, cHn + 1).getValue()); }
    if (rowNum >= 2 && rowNum <= lastRow && (!hn || hnAt(rowNum) === hn)) { ok = true; }
    else if (hn && lastRow > 1) {   // เผื่อแถวเลื่อน: ค้นด้วย HN
      var hns = sh.getRange(2, cHn + 1, lastRow - 1, 1).getValues();
      for (var i = 0; i < hns.length; i++) { if (icnHn_(hns[i][0]) === hn) { rowNum = i + 2; ok = true; break; } }
    }
    if (!ok) return recJson_({ ok: false, note: 'ไม่พบรายการผู้ป่วยในฟอร์ม (HN/แถวไม่ตรง)' });
    sh.getRange(rowNum, cRes + 1).setValue(String(p.result || ''));
    sh.getRange(rowNum, cBy + 1).setValue(String(p.by || ''));
    sh.getRange(rowNum, cAt + 1).setValue(new Date());
    return recJson_({ ok: true, row: rowNum, result: p.result });
  } catch (err) { return recJson_({ ok: false, error: String(err) }); }
}
const REC_LIST_SHEET = 'List-ทะเบียน-1';

function doGet(e) {
  const cb = e && e.parameter && e.parameter.callback;
  const action = e && e.parameter && e.parameter.action;
  if (action === 'login') return recJson_(recLogin_((e.parameter || {}).u, (e.parameter || {}).h), cb);   // รวม login ไว้ไฟล์เดียว
  if (action === 'ping') {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(REC_SHEET);
    const headers = sh ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (x) { return String(x).trim(); }) : [];
    return recJson_({ ok: true, version: 'record-v1', login: 'login-v1', lists: true, sheet: REC_SHEET, found: !!sh, headers: headers }, cb);
  }
  if (action === 'monthlyHeaders') {
    // คืนหัวคอลัมน์ของแท็บ MMW/HDU ให้ฟอร์มสร้างช่องกรอกอัตโนมัติ
    const name = MONTHLY_SHEETS[String(e.parameter.sheet || '').toUpperCase()];
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = name ? ss.getSheetByName(name) : null;
    const headers = sh ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (x) { return String(x).trim(); }).filter(function (x) { return x !== ''; }) : [];
    return recJson_({ ok: true, sheet: name || '', found: !!sh, headers: headers }, cb);
  }
  return recJson_({ ok: true, version: 'record-v1', lists: true }, cb);
}

// เขียนค่าตัวเลือกลงคอลัมน์หนึ่งในแท็บ List-ทะเบียน-1 (สร้างคอลัมน์ใหม่ได้ถ้ายังไม่มี)
function recSetList_(body) {
  const header = String(body.header || '').trim();
  if (!header) return recJson_({ ok: false, note: 'ไม่มีชื่อคอลัมน์' });
  const values = Array.isArray(body.values) ? body.values.map(function (v) { return [v]; }) : [];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(REC_LIST_SHEET);
  if (!sh) sh = ss.insertSheet(REC_LIST_SHEET);
  const lastCol = Math.max(1, sh.getLastColumn());
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (x) { return String(x).trim(); });
  let col = -1;
  for (var i = 0; i < headers.length; i++) { if (headers[i] === header) { col = i + 1; break; } }
  if (col < 0) {
    if (!body.create) return recJson_({ ok: false, note: 'ไม่พบคอลัมน์ "' + header + '"' });
    col = (headers.join('') === '' ? 1 : lastCol + 1);
    sh.getRange(1, col, 1, 1).setValue(header);
  }
  // ล้างค่าของเดิมใต้หัวตาราง แล้วเขียนค่าใหม่
  const maxRow = sh.getMaxRows();
  if (maxRow > 1) sh.getRange(2, col, maxRow - 1, 1).clearContent();
  if (values.length) sh.getRange(2, col, values.length, 1).setValues(values);
  return recJson_({ ok: true, action: 'setList', header: header, count: values.length, col: col });
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (body.action === 'setList') return recSetList_(body);
    if (body.action === 'icnVerify') return icnVerify_(body);         // แอดมินยืนยันผลสงสัยติดเชื้อ (ICWN) ผ่านเว็บ
    if (body.action === 'addMonthly') return recAddMonthly_(body);   // บันทึกรายการหัตถการ/ยอดเดือน → แท็บ MMW/HDU (เดิม)
    if (body.action === 'addUnified') return recAddUnified_(body);    // บันทึกรวมทุกหน่วย → แท็บเดียว "หัตถการ"
    if (body.action === 'editCase') return recEditCase_(body);        // แก้ไขเคสเดิม (ค้นแถวจากค่าที่เคยบันทึก แล้วอัปเดต)
    if (body.action !== 'addCase') return recJson_({ ok: false, note: 'unknown action: ' + body.action });
    // รองรับทั้งบรรทัดเดียว (record) และหลายบรรทัด (records = เคส + เชื้อเพิ่มบรรทัดต่อ)
    const recs = Array.isArray(body.records) ? body.records : [body.record || {}];
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(REC_SHEET);
    if (!sh) return recJson_({ ok: false, note: 'ไม่พบแท็บ "' + REC_SHEET + '"' });
    const lastCol = sh.getLastColumn();
    const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (x) { return String(x).trim(); });
    const hmap = {};
    headers.forEach(function (h, i) { if (h !== '') hmap[recNorm_(h)] = i; });
    let added = 0, matched = 0;
    recs.forEach(function (rec) {
      if (!rec || !Object.keys(rec).length) return;
      const row = new Array(lastCol).fill('');
      Object.keys(rec).forEach(function (k) {
        const idx = hmap[recNorm_(k)];
        if (idx !== undefined) { row[idx] = rec[k]; matched++; }
      });
      sh.appendRow(row);
      added++;
    });
    return recJson_({ ok: true, action: 'addCase', added: added, matched: matched, lastRow: sh.getLastRow() });
  } catch (err) {
    return recJson_({ ok: false, error: String(err) });
  }
}

// ===== แก้ไขเคสเดิม =====
// body.match = ค่าคอลัมน์ระบุตัวของแถวเดิม (Year/Month/Unit/HN/SITE/HAI) ตอนที่บันทึกไว้
// body.record = ค่าทั้งหมดที่ต้องการให้เป็น (คอลัมน์ไหนไม่มีในฟอร์มจะไม่ถูกแตะ)
function recEditCase_(body) {
  const match = body.match || {};
  const rec = body.record || {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(REC_SHEET);
  if (!sh) return recJson_({ ok: false, note: 'ไม่พบแท็บ "' + REC_SHEET + '"' });
  const lastCol = sh.getLastColumn();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return recJson_({ ok: false, action: 'editCase', found: false, note: 'ไม่มีข้อมูลให้แก้ไข' });
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (x) { return String(x).trim(); });
  const hmap = {};
  headers.forEach(function (h, i) { if (h !== '') hmap[recNorm_(h)] = i; });
  const mkeys = Object.keys(match).filter(function (k) { return hmap[recNorm_(k)] !== undefined; });
  if (!mkeys.length) return recJson_({ ok: false, action: 'editCase', found: false, note: 'ไม่มีคีย์ค้นหา' });
  const data = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  let foundRow = -1;
  for (let r = 0; r < data.length; r++) {
    let ok = true;
    for (let i = 0; i < mkeys.length; i++) {
      const idx = hmap[recNorm_(mkeys[i])];
      if (String(data[r][idx]).trim() !== String(match[mkeys[i]]).trim()) { ok = false; break; }
    }
    if (ok) { foundRow = r + 2; break; }
  }
  if (foundRow < 0) return recJson_({ ok: false, action: 'editCase', found: false, note: 'ไม่พบแถวที่ตรงกับรายการเดิม' });
  const cur = sh.getRange(foundRow, 1, 1, lastCol).getValues()[0];
  let matched = 0;
  Object.keys(rec).forEach(function (k) {
    const idx = hmap[recNorm_(k)];
    if (idx !== undefined) { cur[idx] = rec[k]; matched++; }
  });
  sh.getRange(foundRow, 1, 1, lastCol).setValues([cur]);
  return recJson_({ ok: true, action: 'editCase', found: true, row: foundRow, matched: matched });
}

// ===== แท็บรวมทุกหน่วย "หัตถการ" (union ของ MMW + HDU) =====
var PROC_SHEET = 'หัตถการ';
var PROC_HEADERS = [
  'Timestamp', 'หน่วยงาน', 'แบบฟอร์ม', 'รายงานยอดประจำเดือน',
  // MMW: วันคาอุปกรณ์
  'จำนวนแบบบันทึกใบ IC (ใบ)', 'จำนวนวันนอนทั้งหมด (วัน)',
  'ผู้ป่วยคาสายสวนปัสสาวะ (ราย)', 'วันคาสายสวนปัสสาวะ (วัน)',
  'ผู้ป่วยใส่เครื่องช่วยหายใจ (ราย)', 'วันใส่เครื่องช่วยหายใจ (วัน)',
  'ผู้ป่วยใส่ ET/TT ไม่ใส่เครื่อง (ราย)', 'วัน ET/TT ไม่ใส่เครื่อง (วัน)',
  'ผู้ป่วยใส่ CVC (ราย)', 'วันใส่ CVC (วัน)',
  'ผู้ป่วยให้สารน้ำ (ราย)', 'วันให้สารน้ำ (วัน)',
  'ผู้ป่วยให้ NSS Lock (ราย)', 'วันให้ NSS Lock (วัน)',
  // HDU: หน่วยไต
  'HD (ครั้ง)', 'CAPD (ครั้ง)', 'AVF,AVG (ครั้ง)', 'Perm (ครั้ง)', 'DLC (ครั้ง)',
  'หน่วยไต2 HD (ครั้ง)', 'หน่วยไต2 AVF,AVG (ครั้ง)', 'หน่วยไต2 Perm (ครั้ง)', 'หน่วยไต2 DLC (ครั้ง)',
  'CAPD ติดเชื้อ peritonitis (ราย)', 'CAPD ติดเชื้อ Exit site (ราย)',
  'Tenckhoff catheter (ครั้ง)', 'ทำสาย Perm (ครั้ง)', 'ทำสาย DLC (ครั้ง)',
  'วันคาสาย Perm (วัน)', 'วันคาสาย DLC (วัน)',
  'หน่วยไต2 วันคาสาย Perm (วัน)', 'หน่วยไต2 วันคาสาย DLC (วัน)',
  'ลงชื่อผู้รายงาน'
];
function recAddUnified_(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(PROC_SHEET);
  if (!sh) { sh = ss.insertSheet(PROC_SHEET); sh.getRange(1, 1, 1, PROC_HEADERS.length).setValues([PROC_HEADERS]); sh.setFrozenRows(1); }
  const lastCol = Math.max(sh.getLastColumn(), PROC_HEADERS.length);
  let headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (x) { return String(x).trim(); });
  if (!headers.join('').trim()) { sh.getRange(1, 1, 1, PROC_HEADERS.length).setValues([PROC_HEADERS]); headers = PROC_HEADERS.slice(); }
  const hmap = {};
  headers.forEach(function (h, i) { if (h !== '') hmap[recNorm_(h)] = i; });
  const rec = body.record || {};
  if (rec.Timestamp == null && rec.timestamp == null) rec.Timestamp = new Date();
  const row = new Array(headers.length).fill('');
  let matched = 0, unmatched = [];
  Object.keys(rec).forEach(function (k) {
    const idx = hmap[recNorm_(k)];
    if (idx !== undefined) { row[idx] = rec[k]; matched++; } else if (String(rec[k]).trim() !== '') unmatched.push(k);
  });
  sh.appendRow(row);
  return recJson_({ ok: true, action: 'addUnified', sheet: PROC_SHEET, matched: matched, unmatched: unmatched, lastRow: sh.getLastRow() });
}

// บันทึกรายการหัตถการ/ยอดประจำเดือน ต่อท้ายแท็บ MMW หรือ HDU (จับคู่ตามหัวคอลัมน์จริง)
var MONTHLY_SHEETS = { MMW: 'MMW', HDU: 'HDU' };   // เพิ่มหน่วยใหม่ที่นี่ได้ในอนาคต
function recAddMonthly_(body) {
  const name = MONTHLY_SHEETS[String(body.sheet || '').toUpperCase()];
  if (!name) return recJson_({ ok: false, note: 'sheet ไม่ถูกต้อง: ' + body.sheet });
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(name);
  if (!sh) return recJson_({ ok: false, note: 'ไม่พบแท็บ "' + name + '"' });
  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (x) { return String(x).trim(); });
  const hmap = {};
  headers.forEach(function (h, i) { if (h !== '') hmap[recNorm_(h)] = i; });
  const rec = body.record || {};
  if (rec.Timestamp == null && rec.timestamp == null) rec.Timestamp = new Date();   // เติมเวลาอัตโนมัติ
  const row = new Array(lastCol).fill('');
  let matched = 0, unmatched = [];
  Object.keys(rec).forEach(function (k) {
    const idx = hmap[recNorm_(k)];
    if (idx !== undefined) { row[idx] = rec[k]; matched++; } else if (String(rec[k]).trim() !== '') unmatched.push(k);
  });
  sh.appendRow(row);
  return recJson_({ ok: true, action: 'addMonthly', sheet: name, matched: matched, unmatched: unmatched, lastRow: sh.getLastRow() });
}

// ========================================================================
// LOGIN (รวมมาจาก Login.gs — ให้ไฟล์เดียวจบ ไม่ต้องมี doGet ซ้ำอีกไฟล์)
//   ใช้แท็บ "Users" ในไฟล์ทะเบียน · เว็บส่งเฉพาะ SHA-256(password+SALT)
//   ⚠️ SALT ต้องตรงกับหน้าเว็บ (index.html)
// ========================================================================
var REC_LOGIN = { SHEET: 'Users', SALT: 'sswh-ic-2569' };
var REC_MENU_KEYS = ['dashboard', 'wards', 'mdr', 'service', 'indicator', 'analysis',
  'report', 'org', 'patient', 'compare', 'surv', 'verify', 'special', 'postdc', 'control', 'risk', 'record'];
var REC_MENU_LABELS = {
  dashboard: 'แดชบอร์ด', wards: 'สรุปรายหอผู้ป่วย', mdr: 'เชื้อดื้อยา (MDR)', service: 'สถิติบริการ',
  indicator: 'ตัวชี้วัด', analysis: 'บทวิเคราะห์', report: 'สรุปรายงาน', org: 'เชื้อก่อโรค HAI',
  patient: 'วิเคราะห์ผู้ป่วย', compare: 'เปรียบเทียบ', surv: 'ประสิทธิภาพการเฝ้าระวัง',
  verify: 'ตรวจสอบสงสัยการติดเชื้อ', special: 'เฝ้าระวังหน่วยพิเศษ', postdc: 'เฝ้าระวัง Post discharge',
  control: 'Control Chart', risk: 'Risk register', record: 'แบบบันทึกผู้ป่วย HAI'
};
function recNormMenu_(s) { return String(s == null ? '' : s).trim().toLowerCase().replace(/[\s()·\-–_.]/g, ''); }
function recHeaderKey_(h) {
  var n = recNormMenu_(h); if (!n) return null;
  if (n === 'all' || n === 'ทุกเมนู' || n === 'ทั้งหมด') return 'all';
  if (REC_MENU_KEYS.indexOf(n) >= 0) return n;
  for (var k in REC_MENU_LABELS) { if (recNormMenu_(REC_MENU_LABELS[k]) === n) return k; }
  return null;
}
function recChecked_(v) {
  if (v === true) return true;
  var s = String(v == null ? '' : v).trim().toLowerCase();
  return s === 'true' || s === '1' || s === 'yes' || s === 'y' || s === 'x' || s === '✓' || s === '✔' || s === 'true()';
}
function recLogin_(u, h) {
  u = String(u == null ? '' : u).trim().toLowerCase();
  h = String(h == null ? '' : h).trim().toLowerCase();
  if (!u || !h) return { ok: false };
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REC_LOGIN.SHEET);
  if (!sh) return { ok: false, error: 'ไม่พบแท็บ "' + REC_LOGIN.SHEET + '"' };
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return { ok: false };
  var hdr = data[0].map(function (x) { return String(x).trim().toLowerCase(); });
  var ci = { user: hdr.indexOf('username'), pass: hdr.indexOf('password'), name: hdr.indexOf('ชื่อ-สกุล'), role: hdr.indexOf('role'), menus: hdr.indexOf('menus') };
  if (ci.name < 0) ci.name = hdr.indexOf('name');
  if (ci.user < 0 || ci.pass < 0) return { ok: false, error: 'หัวคอลัมน์ต้องมี username และ password' };
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (String(row[ci.user] == null ? '' : row[ci.user]).trim().toLowerCase() !== u) continue;
    var calc = recSha256Hex_(String(row[ci.pass] == null ? '' : row[ci.pass]) + REC_LOGIN.SALT);
    if (calc === h) return { ok: true, username: u, name: ci.name >= 0 ? String(row[ci.name] || '') : u, role: ci.role >= 0 ? String(row[ci.role] || '') : '', menus: recResolveMenus_(hdr, row, ci) };
    return { ok: false };
  }
  return { ok: false };
}
function recResolveMenus_(hdr, row, ci) {
  var allowed = [], seen = {}, sawMenuCol = false;
  for (var c = 0; c < hdr.length; c++) {
    var key = recHeaderKey_(hdr[c]); if (!key) continue;
    sawMenuCol = true;
    if (recChecked_(row[c])) { if (key === 'all') return 'all'; if (!seen[key]) { seen[key] = true; allowed.push(key); } }
  }
  if (sawMenuCol) return allowed.join(',');
  if (ci.menus >= 0) { var t = String(row[ci.menus] || '').trim(); if (t) return t; }
  return 'all';
}
function recSha256Hex_(s) {
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8), hex = '';
  for (var i = 0; i < raw.length; i++) { var b = raw[i] < 0 ? raw[i] + 256 : raw[i], x = b.toString(16); hex += (x.length === 1 ? '0' : '') + x; }
  return hex;
}
