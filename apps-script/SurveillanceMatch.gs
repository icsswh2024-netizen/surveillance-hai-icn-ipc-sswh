/**
 * จับคู่ "ประสิทธิภาพการเฝ้าระวัง" (Yes/No) ลงทะเบียน-1
 * ──────────────────────────────────────────────────────────────
 * หลักการ (sensitivity ของระบบเฝ้าระวัง) — ผลลัพธ์ 3 สถานะ:
 *   มีแท็บอ้างอิง = ฟอร์มรายงานสงสัยติดเชื้อ (ICWN) อยู่ "คนละไฟล์"
 *   จับด้วย HN หรือ ชื่อ+สกุล และ "หน่วย/หอ" ต้องตรงกันด้วย (Unit ↔ หน่วย)
 *     - เจอในฟอร์ม + หน่วยตรง   → "รายงานแล้วตรง"
 *     - เจอในฟอร์ม + หน่วยไม่ตรง → "รายงานแล้วไม่ตรง"
 *     - ไม่เจอในฟอร์ม            → "ไม่รายงาน"
 *   (ปิดเงื่อนไขหน่วยได้ด้วย SM.MATCH_UNIT = false → เจอก็ถือว่า "รายงานแล้วตรง")
 *
 * ติดตั้ง (ครั้งเดียว):
 *   1) เปิดไฟล์ชีตข้อมูลหลัก (ทะเบียน-1, 1KEK2bv...) → Extensions → Apps Script
 *   2) เพิ่มไฟล์นี้ (วางโค้ด) → บันทึก
 *   3) ครั้งแรกจะขอสิทธิ์เข้าถึง Spreadsheet (เพราะเปิดไฟล์ฟอร์มข้ามไฟล์) → อนุญาต
 *      ** บัญชีที่รันต้องมีสิทธิ์เปิดไฟล์ฟอร์ม ICWN ด้วย **
 *   4) รีเฟรชหน้าชีต → มีเมนู "IC Tools"
 *   5) IC Tools → "จับคู่ประสิทธิภาพเฝ้าระวัง (Yes/No)"
 *
 * จับคู่แบบ: ตัดคำนำหน้า (นาย/นาง/น.ส./ด.ช./ด.ญ./พระ) + ช่องว่าง/จุด,
 *   HN ตัดศูนย์นำหน้า (0414052 = 414052). เข้าเงื่อนไขถ้า HN ตรง หรือ ชื่อ+สกุล ตรง.
 */

// ========================= CONFIG =========================
var SM = {
  REC_SHEET: 'ทะเบียน-1-2570',          // แท็บเคสปีงบปัจจุบัน (เปลี่ยนชื่อทุกต้นปีงบ เช่น 'ทะเบียน-1-2571')
  REC_NAME: 'ชื่อ', REC_SURNAME: 'สกุล', REC_HN: 'HN', REC_OUT: 'ประสิทธิภาพ',
  REC_UNIT: 'Unit',                   // รหัสหอในทะเบียน-1 (mmw, icu, su, ...)
  // แท็บอ้างอิง = ฟอร์มรายงานสงสัยติดเชื้อ ICWN (คนละไฟล์)
  REF_FILE_ID: '1HU80TyjUXsvyiiBP8gI41qmTfqaLV2k72NbWJ6lh8FA',
  REF_SHEET: 'การตอบแบบฟอร์ม 1',
  REF_FULLNAME: 'ชื่อ-สกุลผู้ป่วย',   // ชื่อ+สกุล รวมช่องเดียว
  REF_HN: 'HN',
  REF_UNIT: 'หน่วย',                  // คอลัมน์หน่วยในฟอร์ม (MMW, ICU, ...) — ใช้คอลัมน์แรก
  // ตัวเลือก
  MATCH_UNIT: true,      // true = ต้องให้ "หน่วย/หอ" ตรงกันด้วยจึงนับว่า "รายงานแล้วตรง"
  REF_ONLY_HAI: false,   // true = นับเฉพาะแถวที่ "ผลการวินิจฉัย" = HAI (คอลัมน์ M)
  REF_DX: 'ผลการวินิจฉัย',
  WRITE_NO: true,        // true = เขียน "ไม่รายงาน" ให้เคสที่ไม่ตรงด้วย
  // ป้ายผลลัพธ์ 3 สถานะ (ตรงกับคอลัมน์ "ประสิทธิภาพ" ในเว็บ)
  OUT_RIGHT: 'รายงานแล้วตรง', OUT_WRONG: 'รายงานแล้วไม่ตรง', OUT_NONE: 'ไม่รายงาน',
  // ===== จับคู่ "ผลการวินิจฉัย" (ฟอร์ม) ให้ตรงกับคอลัมน์ HAI (ทะเบียน) =====
  REC_HAI: 'HAI',            // คอลัมน์ HAI ในทะเบียน-1 (ค่า 'hai' = ยืนยันติดเชื้อ)
  REC_HAI_VAL: 'hai',        // ค่าที่ถือว่าเป็นเคส HAI
  RES_HAI: 'HAI',            // เขียนลงฟอร์มเมื่อเจอในทะเบียน
  RES_NO: 'No HAI'           // เขียนลงฟอร์มเมื่อไม่เจอ
};
// =========================================================

function onOpen() {
  SpreadsheetApp.getUi().createMenu('IC Tools')
    .addItem('จับคู่ประสิทธิภาพเฝ้าระวัง (ตรง/ไม่ตรง/ไม่รายงาน)', 'smMatchSurveillance')
    .addItem('จับคู่ผลการวินิจฉัย ↔ HAI (ฟอร์มสงสัย)', 'smMatchResult')
    .addToUi();
}

var SM_TITLES = /^(นาย|นางสาว|น\.?ส\.?|นาง|ด\.?ช\.?|ด\.?ญ\.?|เด็กชาย|เด็กหญิง|พระ|สามเณร|ว่าที่ร\.?ต\.?)/;
function smName_(s) {
  var t = String(s == null ? '' : s).trim();
  // ตัดคำนำหน้าออก (วนเผื่อซ้อน เช่น "พระ ...") แล้วลบช่องว่าง/จุด
  for (var i = 0; i < 2; i++) t = t.replace(SM_TITLES, '').trim();
  return t.toLowerCase().replace(/[\s.\-]/g, '');
}
function smHn_(s) { return String(s == null ? '' : s).replace(/\D/g, '').replace(/^0+/, ''); }
// normalize หน่วย/หอ: ตัดช่องว่าง/จุด/ขีด แล้ว lowercase (mmw = MMW)
function smUnit_(s) { return String(s == null ? '' : s).trim().toLowerCase().replace(/[\s.\-]/g, ''); }
// หาดัชนีคอลัมน์ "แรก" ที่ชื่อตรง (เผื่อหัวคอลัมน์ซ้ำ เช่น "หน่วย" มี 2 ช่อง)
function smColFirst_(sheet, header) {
  var hs = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var target = String(header).toLowerCase().replace(/[\s.]/g, '');
  for (var i = 0; i < hs.length; i++) { if (String(hs[i]).trim().toLowerCase().replace(/[\s.]/g, '') === target) return i; }
  return -1;
}

function smHeaderMap_(sheet) {
  var hs = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var map = {};
  hs.forEach(function (h, i) { var k = String(h).trim(); if (k) { map[k] = i; map[k.toLowerCase().replace(/[\s.]/g, '')] = i; } });
  return map;
}
function smCol_(map, header) {
  if (!header) return -1;
  if (map[header] !== undefined) return map[header];
  var n = header.toLowerCase().replace(/[\s.]/g, '');
  return map[n] !== undefined ? map[n] : -1;
}

function smMatchSurveillance() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rec = ss.getSheetByName(SM.REC_SHEET);
  if (!rec) { ui.alert('ไม่พบแท็บ "' + SM.REC_SHEET + '"'); return; }

  // ---- เปิดไฟล์ฟอร์มอ้างอิง (ข้ามไฟล์) ----
  var refSS;
  try { refSS = SpreadsheetApp.openById(SM.REF_FILE_ID); }
  catch (e) { ui.alert('เปิดไฟล์ฟอร์มอ้างอิงไม่ได้ — ตรวจสอบว่าบัญชีนี้มีสิทธิ์เข้าถึงไฟล์\n' + e); return; }
  var ref = refSS.getSheetByName(SM.REF_SHEET);
  if (!ref) { ui.alert('ไม่พบแท็บ "' + SM.REF_SHEET + '" ในไฟล์ฟอร์ม'); return; }

  // ---- สร้างชุดกุญแจจากฟอร์ม (HN + ชื่อ + หน่วย) ----
  var rMap = smHeaderMap_(ref);
  var ciName = smCol_(rMap, SM.REF_FULLNAME), ciHn = smCol_(rMap, SM.REF_HN), ciDx = smCol_(rMap, SM.REF_DX);
  var ciUnit = smColFirst_(ref, SM.REF_UNIT);   // คอลัมน์ "หน่วย" แรกในฟอร์ม
  if (ciName < 0 && ciHn < 0) { ui.alert('ฟอร์มไม่พบคอลัมน์ ชื่อ-สกุล/HN'); return; }
  if (SM.MATCH_UNIT && ciUnit < 0) { ui.alert('เปิดเงื่อนไขหน่วยไว้ แต่ฟอร์มไม่พบคอลัมน์ "' + SM.REF_UNIT + '"'); return; }
  // keysAny = เจอในฟอร์ม (ไม่สนหน่วย) · keysUnit = เจอ + หน่วยตรง
  var rLast = ref.getLastRow(), keysAny = {}, keysUnit = {}, refCnt = 0;
  if (rLast > 1) {
    var rData = ref.getRange(2, 1, rLast - 1, ref.getLastColumn()).getValues();
    rData.forEach(function (row) {
      if (SM.REF_ONLY_HAI && ciDx >= 0 && String(row[ciDx]).trim().toLowerCase().replace(/\s/g, '') !== 'hai') return;
      var nm = ciName >= 0 ? smName_(row[ciName]) : '', hn = ciHn >= 0 ? smHn_(row[ciHn]) : '';
      if (!nm && !hn) return;
      var u = ciUnit >= 0 ? smUnit_(row[ciUnit]) : '';
      if (hn) { keysAny['hn:' + hn] = true; if (u) keysUnit['hn:' + hn + '|' + u] = true; }
      if (nm) { keysAny['nm:' + nm] = true; if (u) keysUnit['nm:' + nm + '|' + u] = true; }
      refCnt++;
    });
  }

  // ---- วนทะเบียน-1 เขียน Yes/No ----
  var mMap = smHeaderMap_(rec);
  var cN = smCol_(mMap, SM.REC_NAME), cS = smCol_(mMap, SM.REC_SURNAME), cH = smCol_(mMap, SM.REC_HN), cOut = smCol_(mMap, SM.REC_OUT);
  var cU = smCol_(mMap, SM.REC_UNIT);
  if (cOut < 0) { ui.alert('ทะเบียน-1 ไม่พบคอลัมน์ "' + SM.REC_OUT + '"'); return; }
  if (SM.MATCH_UNIT && cU < 0) { ui.alert('ทะเบียน-1 ไม่พบคอลัมน์ "' + SM.REC_UNIT + '"'); return; }
  var mLast = rec.getLastRow();
  if (mLast < 2) { ui.alert('ทะเบียน-1 ไม่มีข้อมูล'); return; }

  var mData = rec.getRange(2, 1, mLast - 1, rec.getLastColumn()).getValues();
  var out = [], right = 0, wrong = 0, none = 0, skip = 0;
  mData.forEach(function (row) {
    var nm = smName_((cN >= 0 ? row[cN] : '') + (cS >= 0 ? row[cS] : '')), hn = cH >= 0 ? smHn_(row[cH]) : '';
    if (!nm && !hn) { out.push([row[cOut]]); skip++; return; }  // ข้ามบรรทัดต่อ (ไม่มีชื่อ/HN)
    var reportedAny = (hn && keysAny['hn:' + hn]) || (nm && keysAny['nm:' + nm]);
    if (!reportedAny) {                                   // ไม่เจอในฟอร์มเลย = ไม่รายงาน
      if (SM.WRITE_NO) { out.push([SM.OUT_NONE]); none++; } else { out.push([row[cOut]]); }
      return;
    }
    if (!SM.MATCH_UNIT) { out.push([SM.OUT_RIGHT]); right++; return; }  // ไม่เปิดเงื่อนไขหน่วย = เจอก็ตรง
    var ru = cU >= 0 ? smUnit_(row[cU]) : '';
    var unitOk = ru && ((hn && keysUnit['hn:' + hn + '|' + ru]) || (nm && keysUnit['nm:' + nm + '|' + ru]));
    if (unitOk) { out.push([SM.OUT_RIGHT]); right++; }    // เจอ + หน่วยตรง = รายงานแล้วตรง
    else { out.push([SM.OUT_WRONG]); wrong++; }           // เจอ แต่หน่วยไม่ตรง = รายงานแล้วไม่ตรง
  });
  rec.getRange(2, cOut + 1, out.length, 1).setValues(out);

  ui.alert('จับคู่เสร็จแล้ว ✅\nรายการในฟอร์ม: ' + refCnt +
    '\nรายงานแล้วตรง: ' + right + '  ·  รายงานแล้วไม่ตรง: ' + wrong + '  ·  ไม่รายงาน: ' + none +
    '  ·  ข้าม(บรรทัดต่อ): ' + skip +
    '\nกุญแจ: HN หรือ ชื่อ+สกุล' + (SM.MATCH_UNIT ? ' + หน่วยต้องตรง (Unit ↔ ' + SM.REF_UNIT + ')' : '') +
    (SM.REF_ONLY_HAI ? '\n(นับเฉพาะ ผลการวินิจฉัย = HAI)' : ''));
}

// ========================================================================
// จับคู่ "ผลการวินิจฉัย" (ฟอร์มสงสัย ICWN) ให้ตรงกับคอลัมน์ "HAI" ในทะเบียน-1
//   - สร้างชุดกุญแจ (HN / ชื่อ+สกุล) จากเคสในทะเบียนที่ HAI = 'hai'
//   - วนทุกแถวในฟอร์ม: เจอ → 'HAI', ไม่เจอ → 'No HAI' (เขียนลงคอลัมน์ ผลการวินิจฉัย)
// ========================================================================
function smMatchResult() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rec = ss.getSheetByName(SM.REC_SHEET);
  if (!rec) { ui.alert('ไม่พบแท็บ "' + SM.REC_SHEET + '"'); return; }

  // ---- สร้างชุดกุญแจ HAI จากทะเบียน (เคสที่ HAI = 'hai') ----
  var mMap = smHeaderMap_(rec);
  var cN = smCol_(mMap, SM.REC_NAME), cS = smCol_(mMap, SM.REC_SURNAME), cH = smCol_(mMap, SM.REC_HN), cHai = smCol_(mMap, SM.REC_HAI);
  if (cHai < 0) { ui.alert('ทะเบียน-1 ไม่พบคอลัมน์ "' + SM.REC_HAI + '"'); return; }
  var mLast = rec.getLastRow(), haiKeys = {}, haiCnt = 0;
  if (mLast > 1) {
    var mData = rec.getRange(2, 1, mLast - 1, rec.getLastColumn()).getValues();
    mData.forEach(function (row) {
      if (String(row[cHai]).trim().toLowerCase() !== String(SM.REC_HAI_VAL).toLowerCase()) return;  // เฉพาะเคส HAI
      var nm = smName_((cN >= 0 ? row[cN] : '') + (cS >= 0 ? row[cS] : '')), hn = cH >= 0 ? smHn_(row[cH]) : '';
      if (hn) haiKeys['hn:' + hn] = true;
      if (nm) haiKeys['nm:' + nm] = true;
      haiCnt++;
    });
  }

  // ---- เปิดฟอร์มสงสัย (ข้ามไฟล์) แล้ววนเขียน ผลการวินิจฉัย ----
  var refSS;
  try { refSS = SpreadsheetApp.openById(SM.REF_FILE_ID); }
  catch (e) { ui.alert('เปิดไฟล์ฟอร์มอ้างอิงไม่ได้ — ตรวจสอบสิทธิ์เข้าถึง\n' + e); return; }
  var ref = refSS.getSheetByName(SM.REF_SHEET);
  if (!ref) { ui.alert('ไม่พบแท็บ "' + SM.REF_SHEET + '" ในไฟล์ฟอร์ม'); return; }
  var rMap = smHeaderMap_(ref);
  var ciName = smCol_(rMap, SM.REF_FULLNAME), ciHn = smCol_(rMap, SM.REF_HN), ciOut = smColFirst_(ref, SM.REF_DX);
  if (ciOut < 0) { ui.alert('ฟอร์มไม่พบคอลัมน์ "' + SM.REF_DX + '"'); return; }
  if (ciName < 0 && ciHn < 0) { ui.alert('ฟอร์มไม่พบคอลัมน์ ชื่อ-สกุล/HN'); return; }
  var rLast = ref.getLastRow();
  if (rLast < 2) { ui.alert('ฟอร์มไม่มีข้อมูล'); return; }

  var rData = ref.getRange(2, 1, rLast - 1, ref.getLastColumn()).getValues();
  var out = [], hai = 0, no = 0, skip = 0;
  rData.forEach(function (row) {
    var nm = ciName >= 0 ? smName_(row[ciName]) : '', hn = ciHn >= 0 ? smHn_(row[ciHn]) : '';
    if (!nm && !hn) { out.push([row[ciOut]]); skip++; return; }   // ไม่มีชื่อ/HN → คงค่าเดิม
    var hit = (hn && haiKeys['hn:' + hn]) || (nm && haiKeys['nm:' + nm]);
    if (hit) { out.push([SM.RES_HAI]); hai++; } else { out.push([SM.RES_NO]); no++; }
  });
  ref.getRange(2, ciOut + 1, out.length, 1).setValues(out);

  ui.alert('จับคู่ผลการวินิจฉัยเสร็จแล้ว ✅' +
    '\nเคส HAI ในทะเบียน: ' + haiCnt +
    '\nเขียน "' + SM.RES_HAI + '": ' + hai + '  ·  "' + SM.RES_NO + '": ' + no + '  ·  ข้าม: ' + skip +
    '\nกุญแจ: HN หรือ ชื่อ+สกุล (ตัดคำนำหน้า/ช่องว่าง)');
}
