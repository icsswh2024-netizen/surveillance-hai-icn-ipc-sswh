/**
 * จับคู่ "ประสิทธิภาพการเฝ้าระวัง" (Yes/No) ลงทะเบียน-1
 * ──────────────────────────────────────────────────────────────
 * หลักการ: มีแท็บอ้างอิง (รายชื่อที่ระบบเฝ้าระวังตรวจพบ) ในไฟล์เดียวกัน
 *   ถ้าเคสในทะเบียน-1 มี ชื่อ+สกุล+HN ตรงกับรายการในแท็บอ้างอิง → ประสิทธิภาพ = "Yes"
 *   ถ้าไม่ตรง (แต่เป็นเคสที่มีข้อมูล) → "No"
 *
 * ติดตั้ง (ครั้งเดียว):
 *   1) เปิดไฟล์ชีตข้อมูล (1KEK2bv...) → Extensions → Apps Script
 *   2) เพิ่มไฟล์นี้ (วางโค้ด) → บันทึก
 *   3) รีเฟรชหน้าชีต จะมีเมนู "IC Tools" โผล่บนแถบเมนู
 *   4) ตั้งชื่อแท็บอ้างอิง + ชื่อหัวคอลัมน์ ในส่วน CONFIG ด้านล่างให้ตรงกับชีตจริง
 *   5) เมนู IC Tools → "จับคู่ประสิทธิภาพเฝ้าระวัง (Yes/No)"
 *
 * หมายเหตุ: จับคู่แบบไม่สนตัวพิมพ์/ช่องว่าง/จุด และ HN ตัดศูนย์นำหน้าออกให้เทียบได้
 */

// ========================= CONFIG (แก้ให้ตรงชีตจริง) =========================
var SM_CONFIG = {
  REC_SHEET: 'ทะเบียน-1',          // แท็บทะเบียนหลัก (ปลายทางที่จะเขียน Yes/No)
  REF_SHEET: 'เฝ้าระวัง',           // << ชื่อแท็บอ้างอิง (รายชื่อที่เฝ้าระวังจับได้) — แก้ให้ตรง
  // หัวคอลัมน์ในทะเบียน-1 (ปกติไม่ต้องแก้)
  REC_NAME: 'ชื่อ', REC_SURNAME: 'สกุล', REC_HN: 'HN', REC_OUT: 'ประสิทธิภาพ',
  // หัวคอลัมน์ในแท็บอ้างอิง — ถ้าชื่อหัวต่างกัน แก้ตรงนี้ (เว้นว่างได้ถ้าไม่มีคอลัมน์นั้น)
  REF_NAME: 'ชื่อ', REF_SURNAME: 'สกุล', REF_HN: 'HN',
  WRITE_NO: true                   // true = เขียน "No" ให้เคสที่ไม่ตรงด้วย (false = เว้นว่างไว้)
};
// ===========================================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('IC Tools')
    .addItem('จับคู่ประสิทธิภาพเฝ้าระวัง (Yes/No)', 'smMatchSurveillance')
    .addToUi();
}

function smNorm_(s) { return String(s == null ? '' : s).toLowerCase().replace(/[\s.\-]/g, ''); }
// HN: ตัดศูนย์นำหน้า + ตัดอักขระไม่ใช่ตัวเลข เพื่อให้ 0012345 = 12345
function smHn_(s) { return String(s == null ? '' : s).replace(/\D/g, '').replace(/^0+/, ''); }
// กุญแจจับคู่: ชื่อ+สกุล+HN
function smKey_(name, surname, hn) { return smNorm_(name) + '|' + smNorm_(surname) + '|' + smHn_(hn); }

function smHeaderMap_(sheet) {
  var lastCol = sheet.getLastColumn();
  var hs = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var map = {};
  hs.forEach(function (h, i) { var k = String(h).trim(); if (k) map[k] = i; map[smNorm_(h)] = i; });
  return map;
}
function smCol_(map, header) {
  if (header == null || header === '') return -1;
  if (map[header] !== undefined) return map[header];
  var n = smNorm_(header);
  return map[n] !== undefined ? map[n] : -1;
}

function smMatchSurveillance() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var cfg = SM_CONFIG;

  var ref = ss.getSheetByName(cfg.REF_SHEET);
  if (!ref) { ui.alert('ไม่พบแท็บอ้างอิง "' + cfg.REF_SHEET + '" — แก้ชื่อใน CONFIG (REF_SHEET) ให้ตรง'); return; }
  var rec = ss.getSheetByName(cfg.REC_SHEET);
  if (!rec) { ui.alert('ไม่พบแท็บ "' + cfg.REC_SHEET + '"'); return; }

  // ----- สร้างชุดกุญแจจากแท็บอ้างอิง -----
  var rMap = smHeaderMap_(ref);
  var ciN = smCol_(rMap, cfg.REF_NAME), ciS = smCol_(rMap, cfg.REF_SURNAME), ciH = smCol_(rMap, cfg.REF_HN);
  if (ciN < 0 && ciH < 0) { ui.alert('แท็บอ้างอิงไม่พบคอลัมน์ ชื่อ/HN — ตรวจ CONFIG (REF_NAME/REF_HN)'); return; }
  var rLast = ref.getLastRow();
  var keys = {};
  if (rLast > 1) {
    var rData = ref.getRange(2, 1, rLast - 1, ref.getLastColumn()).getValues();
    rData.forEach(function (row) {
      var nm = ciN >= 0 ? row[ciN] : '', su = ciS >= 0 ? row[ciS] : '', hn = ciH >= 0 ? row[ciH] : '';
      if (smNorm_(nm) === '' && smHn_(hn) === '') return;
      keys[smKey_(nm, su, hn)] = true;                 // คีย์เต็ม ชื่อ+สกุล+HN
      if (smHn_(hn)) keys['hn:' + smHn_(hn)] = true;    // สำรอง: จับด้วย HN อย่างเดียว
    });
  }

  // ----- วนทะเบียน-1 เขียน Yes/No -----
  var mMap = smHeaderMap_(rec);
  var cN = smCol_(mMap, cfg.REC_NAME), cS = smCol_(mMap, cfg.REC_SURNAME), cH = smCol_(mMap, cfg.REC_HN), cOut = smCol_(mMap, cfg.REC_OUT);
  if (cOut < 0) { ui.alert('ทะเบียน-1 ไม่พบคอลัมน์ "' + cfg.REC_OUT + '"'); return; }
  var mLast = rec.getLastRow();
  if (mLast < 2) { ui.alert('ทะเบียน-1 ไม่มีข้อมูล'); return; }

  var mData = rec.getRange(2, 1, mLast - 1, rec.getLastColumn()).getValues();
  var out = [], yes = 0, no = 0, skip = 0;
  mData.forEach(function (row) {
    var nm = cN >= 0 ? row[cN] : '', su = cS >= 0 ? row[cS] : '', hn = cH >= 0 ? row[cH] : '';
    // ข้ามแถวที่ไม่มีทั้งชื่อและ HN (เช่น บรรทัดต่อของเชื้อเพิ่ม)
    if (smNorm_(nm) === '' && smHn_(hn) === '') { out.push([row[cOut]]); skip++; return; }
    var hit = keys[smKey_(nm, su, hn)] || (smHn_(hn) && keys['hn:' + smHn_(hn)]);
    if (hit) { out.push(['Yes']); yes++; }
    else if (cfg.WRITE_NO) { out.push(['No']); no++; }
    else { out.push([row[cOut]]); }
  });
  rec.getRange(2, cOut + 1, out.length, 1).setValues(out);

  ui.alert('จับคู่เสร็จแล้ว ✅\nYes: ' + yes + '  ·  No: ' + no + '  ·  ข้าม(บรรทัดต่อ): ' + skip +
    '\nแท็บอ้างอิง: ' + cfg.REF_SHEET + '  ·  กุญแจ: ชื่อ+สกุล+HN (สำรอง HN)');
}
