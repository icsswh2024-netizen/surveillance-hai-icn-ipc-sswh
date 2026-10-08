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

const REC_SHEET = 'ทะเบียน-1';

function recJson_(obj, cb) {
  const out = JSON.stringify(obj);
  if (cb) return ContentService.createTextOutput(cb + '(' + out + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(out).setMimeType(ContentService.MimeType.JSON);
}
function recNorm_(s) { return String(s == null ? '' : s).toLowerCase().replace(/[\s.]/g, ''); }
const REC_LIST_SHEET = 'List-ทะเบียน-1';

function doGet(e) {
  const cb = e && e.parameter && e.parameter.callback;
  const action = e && e.parameter && e.parameter.action;
  if (action === 'ping') {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(REC_SHEET);
    const headers = sh ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (x) { return String(x).trim(); }) : [];
    return recJson_({ ok: true, version: 'record-v1', lists: true, sheet: REC_SHEET, found: !!sh, headers: headers }, cb);
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
