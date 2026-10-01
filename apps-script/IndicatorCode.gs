/**
 * ตัวชี้วัด IPC — Google Apps Script
 * ผูกกับไฟล์: 0-ตัวชี้วัดงาน IPC (1uXfgiUX2_AJzCQam9uTdj1gMfJpxTLVONf0ndILMoGE)
 *
 * ใช้ 2 อย่าง:
 *   1) setupIndicatorSheet()  — รันครั้งเดียวในตัวแก้ไข (สร้าง/เติมแท็บ "ทะเบียนตัวชี้วัด")
 *   2) doGet                   — ให้เว็บอ่านทะเบียน/ค่าตัวชี้วัดผ่าน JSONP (ถ้าไม่ใช้ gviz)
 *
 * ติดตั้ง: เปิดไฟล์ชีต → Extensions → Apps Script → วางโค้ดนี้ → บันทึก
 *   - จะรัน setupIndicatorSheet: เลือกฟังก์ชันแล้วกด Run (อนุญาตสิทธิ์ครั้งแรก)
 *   - จะใช้ doGet: Deploy → New deployment → Web app · Execute as = Me · Anyone → คัดลอก /exec
 */

const INDICATOR_SHEET = 'ทะเบียนตัวชี้วัด';
const INDICATOR_HEADERS = ['รหัส', 'ชื่อตัวชี้วัด', 'ตัวตั้ง', 'ตัวหาร', 'หน่วย', 'ความถี่', 'source', 'metricKey', 'NQA', 'HA6', 'เขต', 'สปสช.', 'เป้าหมาย', 'เป้าหมาย_NQA', 'เป้าหมาย_HA6', 'เป้าหมาย_เขต', 'เป้าหมาย_สปสช.', 'วันที่อัปเดตเป้า', 'ผู้รับผิดชอบ', 'ข้อมูลเดิมที่ใช้'];

function jsonOut_(obj, cb) {
  const out = JSON.stringify(obj);
  if (cb) return ContentService.createTextOutput(cb + '(' + out + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(out).setMimeType(ContentService.MimeType.JSON);
}

/*** ===== อ่านทะเบียนตัวชี้วัด (JSONP) ===== ***/
function doGet(e) {
  const cb = e && e.parameter && e.parameter.callback;
  const action = e && e.parameter && e.parameter.action;
  if (action === 'ping') {
    const names = SpreadsheetApp.getActiveSpreadsheet().getSheets().map(function (s) { return s.getName(); });
    return jsonOut_({ ok: true, version: 'indicator-v2', write: true, sheets: names }, cb);
  }
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(INDICATOR_SHEET);
  if (!sh) return jsonOut_({ ok: false, note: 'ไม่พบแท็บ "' + INDICATOR_SHEET + '"', rows: [] }, cb);
  const data = sh.getDataRange().getValues();
  if (!data.length) return jsonOut_({ ok: true, rows: [] }, cb);
  const head = data.shift().map(function (x) { return String(x).trim(); });
  const rows = data.filter(function (r) { return String(r[0] || '').trim() !== ''; }).map(function (r) {
    const o = {};
    head.forEach(function (h, i) { o[h] = r[i]; });
    return o;
  });
  return jsonOut_({ ok: true, rows: rows }, cb);
}

/*** ===== เพิ่ม/แก้ไข/ลบ ตัวชี้วัด (เขียนกลับชีต) =====
 * action = 'upsert'  → body.record = { 'รหัส':..., 'ชื่อตัวชี้วัด':..., ... }
 *                       (เพิ่มใหม่ถ้ายังไม่มีรหัส / แก้ไขถ้ามีอยู่แล้ว)
 *          'delete'  → body.code = 'IPC-xx'
 * คอลัมน์ที่ยังไม่มีในหัวตารางจะถูกสร้างให้อัตโนมัติ
 ***/
function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const action = body.action;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(INDICATOR_SHEET);
    if (!sh) { sh = ss.insertSheet(INDICATOR_SHEET); sh.getRange(1, 1, 1, INDICATOR_HEADERS.length).setValues([INDICATOR_HEADERS]); }
    let data = sh.getDataRange().getValues();
    if (!data.length) { sh.getRange(1, 1, 1, INDICATOR_HEADERS.length).setValues([INDICATOR_HEADERS]); data = [INDICATOR_HEADERS.slice()]; }
    let head = data[0].map(function (x) { return String(x).trim(); });
    const codeCol = head.indexOf('รหัส') < 0 ? 0 : head.indexOf('รหัส');

    if (action === 'upsert') {
      const rec = body.record || {};
      const code = String(rec['รหัส'] || '').trim();
      if (!code) return jsonOut_({ ok: false, note: 'ไม่มีรหัสตัวชี้วัด' });
      // สร้างคอลัมน์ใหม่อัตโนมัติถ้ายังไม่มี
      let headChanged = false;
      Object.keys(rec).forEach(function (k) { if (head.indexOf(k) < 0) { head.push(k); headChanged = true; } });
      if (headChanged) sh.getRange(1, 1, 1, head.length).setValues([head]);
      const arr = head.map(function (h) { return rec[h] != null ? rec[h] : ''; });
      let foundRow = -1;
      for (let i = 1; i < data.length; i++) { if (String(data[i][codeCol] || '').trim() === code) { foundRow = i + 1; break; } }
      if (foundRow > 0) sh.getRange(foundRow, 1, 1, arr.length).setValues([arr]);
      else sh.appendRow(arr);
      return jsonOut_({ ok: true, action: 'upsert', code: code, updated: foundRow > 0 });
    }

    if (action === 'delete') {
      const code = String(body.code || '').trim();
      if (!code) return jsonOut_({ ok: false, note: 'ไม่มีรหัสตัวชี้วัด' });
      let removed = 0;
      for (let i = data.length - 1; i >= 1; i--) { if (String(data[i][codeCol] || '').trim() === code) { sh.deleteRow(i + 1); removed++; } }
      return jsonOut_({ ok: true, action: 'delete', code: code, removed: removed });
    }

    return jsonOut_({ ok: false, note: 'unknown action: ' + action });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

/*** ===== ติดตั้ง/เติมแท็บ "ทะเบียนตัวชี้วัด" (รันครั้งเดียว) =====
 * ปลอดภัย: ถ้ามีข้อมูลอยู่แล้ว (มากกว่าหัวตาราง) จะไม่เขียนทับ
 ***/
function setupIndicatorSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(INDICATOR_SHEET);
  if (sh && sh.getLastRow() > 1) {
    Logger.log('มีแท็บ ' + INDICATOR_SHEET + ' และมีข้อมูลอยู่แล้ว — ไม่เขียนทับ');
    return;
  }
  if (!sh) sh = ss.insertSheet(INDICATOR_SHEET);
  const rows = [
    INDICATOR_HEADERS,
    ['IPC-01', 'อัตราการติดเชื้อในโรงพยาบาล (HAI) รวม', 'จำนวนครั้งการติดเชื้อ HAI', 'จำนวนวันนอนผู้ป่วย', '/1000 วันนอน', 'เดือน/ไตรมาส/ปี', 'auto', 'hai_rate', '✓', '✓', '✓', '✓', '<1.08', 'ICN', 'ทะเบียน-1'],
    ['IPC-02', 'อัตราการติดเชื้อทางเดินปัสสาวะจากคาสายสวน (CAUTI)', 'จำนวนครั้ง CAUTI', 'จำนวนวันคาสายสวนปัสสาวะ', '/1000 วันคาสาย', 'เดือน/ไตรมาส/ปี', 'auto', 'cauti_rate', '✓', '✓', '✓', '', '<2.92', 'ICN', 'ทะเบียน-1 (RATE_META)'],
    ['IPC-03', 'อัตราการเกิดปอดอักเสบจากการใช้เครื่องช่วยหายใจ (VAP)', 'จำนวนครั้ง VAP', 'จำนวนวันใช้เครื่องช่วยหายใจ', '/1000 วันเครื่องช่วยหายใจ', 'เดือน/ไตรมาส/ปี', 'auto', 'vap_rate', '✓', '✓', '✓', '', '<1.34', 'ICN', 'ทะเบียน-1 (RATE_META)'],
    ['IPC-04', 'อัตราการติดเชื้อในกระแสเลือดจากสายสวนหลอดเลือด (CLABSI)', 'จำนวนครั้ง CLABSI', 'จำนวนวันคาสายสวนหลอดเลือด', '/1000 วันคาสาย', 'เดือน/ไตรมาส/ปี', 'auto', 'clabsi_rate', '✓', '✓', '✓', '', '<2.19', 'ICN', 'ทะเบียน-1 (RATE_META)'],
    ['IPC-05', 'อัตราการติดเชื้อตำแหน่งผ่าตัด (SSI)', 'จำนวนครั้ง SSI', 'จำนวนครั้งการผ่าตัด', 'ร้อยละ', 'ไตรมาส/ปี', 'auto', 'ssi_rate', '✓', '✓', '✓', '', '<0.22', 'ICN', 'ทะเบียน-1 / Post-discharge'],
    ['IPC-06', 'ร้อยละการติดเชื้อดื้อยา (MDRO) ต่อการติดเชื้อทั้งหมด', 'จำนวนเคสติดเชื้อดื้อยา', 'จำนวนเคสติดเชื้อ HAI ทั้งหมด', 'ร้อยละ', 'เดือน/ไตรมาส/ปี', 'auto', 'mdr_pct', '✓', '✓', '✓', '✓', 'ลดลงจากปีก่อน', 'ICN', 'เมนู MDR เดิม'],
    ['IPC-07', 'อัตราการติดเชื้อดื้อยา (MDRO) ต่อวันนอน', 'จำนวนเคสติดเชื้อดื้อยา', 'จำนวนวันนอนผู้ป่วย', '/1000 วันนอน', 'เดือน/ไตรมาส/ปี', 'auto', 'mdro_rate', '', '✓', '✓', '', 'ลดลงจากปีก่อน', 'ICN', 'เมนู MDR เดิม'],
    ['IPC-08', 'จำนวนบุคลากรสัมผัสเลือด/สารคัดหลั่งจากการปฏิบัติหน้าที่', 'จำนวนเหตุการณ์สัมผัส', '-', 'ครั้ง', 'ไตรมาส/ปี', 'external', 'staff_sharp', '✓', '✓', '', '', '<10 ครั้ง', 'ICN/อาชีวอนามัย', 'ชีต IC-อุบัติเหตุเจ้าหน้าที่'],
    ['IPC-09', 'อัตราการล้างมือถูกต้องตามโอกาส (Hand Hygiene Compliance)', 'จำนวนครั้งที่ล้างมือถูกต้อง', 'จำนวนโอกาสที่ควรล้างมือ (สุ่มสังเกต)', 'ร้อยละ', 'เดือน/ไตรมาส', 'manual', 'hh_compliance', '✓', '✓', '✓', '✓', '≥85%', 'ICN', 'เริ่มเก็บใหม่ (5 moments)'],
    ['IPC-10', 'ความครบถ้วนการปฏิบัติตาม Bundle (CAUTI/VAP/CLABSI)', 'จำนวนรายการ Bundle ที่ปฏิบัติครบ', 'จำนวนรายการ Bundle ที่ตรวจทั้งหมด', 'ร้อยละ', 'เดือน/ไตรมาส', 'manual', 'bundle_compliance', '✓', '✓', '', '', '≥90%', 'ICN/หอผู้ป่วย', 'เริ่มเก็บใหม่ (audit)'],
    ['IPC-11', 'ความครอบคลุมการเฝ้าระวังการติดเชื้อ (Surveillance Coverage)', 'จำนวนหน่วย/ผู้ป่วยที่เฝ้าระวัง', 'จำนวนหน่วย/ผู้ป่วยเป้าหมายทั้งหมด', 'ร้อยละ', 'ไตรมาส/ปี', 'manual', 'surv_coverage', '', '✓', '', '', '100%', 'ICN', 'เริ่มเก็บใหม่'],
    ['IPC-12', 'อัตราบุคลากรได้รับวัคซีนป้องกันโรค (ไข้หวัดใหญ่/HBV)', 'จำนวนบุคลากรที่ได้รับวัคซีนครบ', 'จำนวนบุคลากรกลุ่มเป้าหมาย', 'ร้อยละ', 'ปี', 'manual', 'staff_vaccine', '', '✓', '', '', '≥90%', 'ICN/อาชีวอนามัย', 'เริ่มเก็บใหม่'],
    ['IPC-13', 'จำนวนบุคลากรติดเชื้อวัณโรคจากการปฏิบัติงาน', 'จำนวนบุคลากรติดเชื้อวัณโรค', '-', 'คน', 'ปี', 'manual', 'staff_tb', '', '✓', '', '', '0 คน', 'ICN/อาชีวอนามัย', 'ชีตตัวชี้วัดเดิม (ข้อ 8)'],
    ['IPC-14', 'การเฝ้าระวังคุณภาพสิ่งแวดล้อม (น้ำดื่ม/น้ำประปา/น้ำเสีย)', 'ผลการตรวจผ่านเกณฑ์', '-', 'ผ่าน/ไม่ผ่าน', 'ปี', 'manual', 'env_surv', '', '✓', '', '', 'ผ่าน', 'ICN', 'ชีตตัวชี้วัดเดิม (ข้อ 11)'],
    ['IPC-15', 'ประสิทธิภาพการเฝ้าระวังการติดเชื้อ (ความไวการวินิจฉัย)', 'จำนวน Case วินิจฉัยถูกต้อง', 'จำนวน Case สงสัยทั้งหมด', 'ร้อยละ', 'เดือน/ไตรมาส', 'manual', 'surv_sensitivity', '', '✓', '', '', '≥80%', 'ICN', 'ชีตประสิทธิภาพการเฝ้าระวัง']
  ];
  // แทรก 4 คอลัมน์เป้าแยกหมวด (เว้นว่างให้กรอกภายหลัง) หลังคอลัมน์ "เป้าหมาย" (index 12)
  // ยกเว้นแถวหัวตาราง (index 0) ที่เป็น INDICATOR_HEADERS อยู่แล้ว
  for (var i = 1; i < rows.length; i++) {
    rows[i].splice(13, 0, '', '', '', '', '');
  }
  sh.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sh.getRange(1, 1, 1, rows[0].length).setFontWeight('bold').setBackground('#ecfdf5');
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, rows[0].length);
  Logger.log('สร้างแท็บ ' + INDICATOR_SHEET + ' และเติม ' + (rows.length - 1) + ' ตัวชี้วัดเรียบร้อย');
}
