/**
 * จับคู่ "ประสิทธิภาพการเฝ้าระวัง" (Yes/No) ลงทะเบียน-1
 * ──────────────────────────────────────────────────────────────
 * หลักการ (sensitivity ของระบบเฝ้าระวัง):
 *   มีแท็บอ้างอิง = ฟอร์มรายงานสงสัยติดเชื้อ (ICWN) อยู่ "คนละไฟล์"
 *   ถ้าเคสในทะเบียน-1 ตรงกับรายการในฟอร์ม (จับด้วย HN หรือ ชื่อ+สกุล)
 *     → ประสิทธิภาพ = "Yes"  (ระบบเฝ้าระวังตรวจพบ/มีการรายงาน)
 *   ถ้าไม่ตรง → "No"
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
  REC_SHEET: 'ทะเบียน-1',
  REC_NAME: 'ชื่อ', REC_SURNAME: 'สกุล', REC_HN: 'HN', REC_OUT: 'ประสิทธิภาพ',
  // แท็บอ้างอิง = ฟอร์มรายงานสงสัยติดเชื้อ ICWN (คนละไฟล์)
  REF_FILE_ID: '1HU80TyjUXsvyiiBP8gI41qmTfqaLV2k72NbWJ6lh8FA',
  REF_SHEET: 'การตอบแบบฟอร์ม 1',
  REF_FULLNAME: 'ชื่อ-สกุลผู้ป่วย',   // ชื่อ+สกุล รวมช่องเดียว
  REF_HN: 'HN',
  // ตัวเลือก
  REF_ONLY_HAI: false,   // true = นับเฉพาะแถวที่ "ผลการวินิจฉัย" = HAI (คอลัมน์ M)
  REF_DX: 'ผลการวินิจฉัย',
  WRITE_NO: true         // true = เขียน "No" ให้เคสที่ไม่ตรงด้วย
};
// =========================================================

function onOpen() {
  SpreadsheetApp.getUi().createMenu('IC Tools')
    .addItem('จับคู่ประสิทธิภาพเฝ้าระวัง (Yes/No)', 'smMatchSurveillance')
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

  // ---- สร้างชุดกุญแจจากฟอร์ม (HN + ชื่อ) ----
  var rMap = smHeaderMap_(ref);
  var ciName = smCol_(rMap, SM.REF_FULLNAME), ciHn = smCol_(rMap, SM.REF_HN), ciDx = smCol_(rMap, SM.REF_DX);
  if (ciName < 0 && ciHn < 0) { ui.alert('ฟอร์มไม่พบคอลัมน์ ชื่อ-สกุล/HN'); return; }
  var rLast = ref.getLastRow(), keys = {}, refCnt = 0;
  if (rLast > 1) {
    var rData = ref.getRange(2, 1, rLast - 1, ref.getLastColumn()).getValues();
    rData.forEach(function (row) {
      if (SM.REF_ONLY_HAI && ciDx >= 0 && String(row[ciDx]).trim().toLowerCase().replace(/\s/g, '') !== 'hai') return;
      var nm = ciName >= 0 ? smName_(row[ciName]) : '', hn = ciHn >= 0 ? smHn_(row[ciHn]) : '';
      if (!nm && !hn) return;
      if (hn) keys['hn:' + hn] = true;
      if (nm) keys['nm:' + nm] = true;
      refCnt++;
    });
  }

  // ---- วนทะเบียน-1 เขียน Yes/No ----
  var mMap = smHeaderMap_(rec);
  var cN = smCol_(mMap, SM.REC_NAME), cS = smCol_(mMap, SM.REC_SURNAME), cH = smCol_(mMap, SM.REC_HN), cOut = smCol_(mMap, SM.REC_OUT);
  if (cOut < 0) { ui.alert('ทะเบียน-1 ไม่พบคอลัมน์ "' + SM.REC_OUT + '"'); return; }
  var mLast = rec.getLastRow();
  if (mLast < 2) { ui.alert('ทะเบียน-1 ไม่มีข้อมูล'); return; }

  var mData = rec.getRange(2, 1, mLast - 1, rec.getLastColumn()).getValues();
  var out = [], yes = 0, no = 0, skip = 0;
  mData.forEach(function (row) {
    var nm = smName_((cN >= 0 ? row[cN] : '') + (cS >= 0 ? row[cS] : '')), hn = cH >= 0 ? smHn_(row[cH]) : '';
    if (!nm && !hn) { out.push([row[cOut]]); skip++; return; }  // ข้ามบรรทัดต่อ
    var hit = (hn && keys['hn:' + hn]) || (nm && keys['nm:' + nm]);
    if (hit) { out.push(['Yes']); yes++; }
    else if (SM.WRITE_NO) { out.push(['No']); no++; }
    else { out.push([row[cOut]]); }
  });
  rec.getRange(2, cOut + 1, out.length, 1).setValues(out);

  ui.alert('จับคู่เสร็จแล้ว ✅\nรายการในฟอร์ม: ' + refCnt +
    '\nYes: ' + yes + '  ·  No: ' + no + '  ·  ข้าม(บรรทัดต่อ): ' + skip +
    '\nกุญแจ: HN หรือ ชื่อ+สกุล (ตัดคำนำหน้า/ช่องว่าง)' +
    (SM.REF_ONLY_HAI ? '\n(นับเฉพาะ ผลการวินิจฉัย = HAI)' : ''));
}
