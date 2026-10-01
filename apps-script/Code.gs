/**
 * Risk Register — Google Apps Script Web App
 * ใช้คู่กับ index.html (เมนู Risk register)
 *
 * ติดตั้ง:
 *   1) เปิดชีต Risk-register → Extensions → Apps Script → วางโค้ดนี้ทั้งไฟล์ → บันทึก
 *   2) Deploy → New deployment → Web app · Execute as = Me · Who has access = Anyone → Deploy
 *   3) คัดลอก Web app URL (ลงท้าย /exec) ไปตั้งใน index.html (ตัวแปร RISK_SAVE_URL)
 *   * แก้โค้ดครั้งใด ต้อง Deploy → Manage deployments → New version ทุกครั้ง
 *
 * แท็บที่ใช้:
 *   - RiskRegister : เก็บข้อมูลที่บันทึกจากแดชบอร์ด (สร้างเองอัตโนมัติ)
 *   - ตัวเลือก      : ตัวเลือก dropdown (หัวคอลัมน์ตาม map ใน readOptions_)
 *   - เกณฑ์         : เป้าหมายรายปี (A=site, หัวคอลัมน์มีเลขปี เช่น "เป้าหมาย 2568")
 */

/*** ===== ตั้งค่า ===== ***/
const SHEET_NAME   = 'RiskRegister';
const OPTION_SHEET = 'ตัวเลือก';
const TARGET_SHEET = 'เกณฑ์';

const HEADERS = ['Source','Data Added Date','Risk Title','Risk Description','Quarter',
  'Likelihood','Consequence','Risk Level','Risk Transfer & Prevention','Risk Monitor & Control',
  'Risk Mitigation','QI plan','Risk Owner','Review Frequency','Data last review',
  'Residual risk level','Risk status','Risk ID','เป้าหมาย'];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) { sh = ss.insertSheet(SHEET_NAME); sh.appendRow(HEADERS); }
  return sh;
}

// วันที่ → สตริง yyyy-MM-dd (ตัดเวลา) ตามเขตเวลาของสคริปต์
function dstr_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return v;
}

function jsonOut_(obj, cb) {
  const out = JSON.stringify(obj);
  if (cb) return ContentService.createTextOutput(cb + '(' + out + ')')
    .setMimeType(ContentService.MimeType.JAVASCRIPT); // JSONP
  return ContentService.createTextOutput(out)
    .setMimeType(ContentService.MimeType.JSON);
}

/*** ===== บันทึก (เขียนทับทั้งแท็บ) ===== ***/
function doPost(e) {
  if (!e || !e.postData) return jsonOut_({ ok: false, note: 'เรียกผ่าน Web app POST เท่านั้น' });
  const body = JSON.parse(e.postData.contents);

  // กรอกลงแท็บ "ปีงบ 25xx" ตามแบบฟอร์ม RM ทางการ (เขียนที่ A4 · เก็บหัวตารางเดิม)
  if (body.action === 'fillYear') {
    const name = 'ปีงบ ' + body.year;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh2 = ss.getSheetByName(name);
    if (!sh2) return jsonOut_({ ok: false, note: 'ไม่พบแท็บ "' + name + '"' });
    const aoa = body.aoa || [];
    const START = 4, COLS = 17; // หัวตาราง 3 แถวแรกของ template
    const last = sh2.getLastRow();
    if (last >= START) sh2.getRange(START, 1, last - START + 1, COLS).clearContent();
    if (aoa.length) {
      sh2.getRange(START, 1, aoa.length, COLS).setValues(aoa);
      // บังคับคอลัมน์วันที่เป็นวันที่ล้วน (ไม่ติดเวลา): B=Data Added, O=Data last review
      sh2.getRange(START, 2, aoa.length, 1).setNumberFormat('dd/mm/yyyy');
      sh2.getRange(START, 15, aoa.length, 1).setNumberFormat('dd/mm/yyyy');
    }
    return jsonOut_({ ok: true, filled: aoa.length, sheet: name });
  }

  const sh = sheet_();
  sh.clear();
  sh.appendRow(HEADERS);
  (body.rows || []).forEach(function (r) {
    sh.appendRow([r.source, r.dataAddedDate, r.riskTitle, r.riskDescription, r.quarter,
      r.likelihood, r.consequence, r.riskLevel, r.prevention, r.monitorControl,
      r.mitigation, r.qiPlan, r.riskOwner, r.reviewFrequency, r.dataLastReview,
      r.residualLevel, r.riskStatus, r.riskId, r.target]);
  });
  // บังคับคอลัมน์วันที่เป็นวันที่ล้วน (ไม่ติดเวลา): B=Data Added Date, O=Data last review
  var nRows = (body.rows || []).length;
  if (nRows) {
    sh.getRange(2, 2, nRows, 1).setNumberFormat('dd/mm/yyyy');
    sh.getRange(2, 15, nRows, 1).setNumberFormat('dd/mm/yyyy');
  }
  return jsonOut_({ ok: true, saved: (body.rows || []).length });
}

/*** ===== เรียกดู / ตัวเลือก / เป้าหมาย (รองรับ JSONP) ===== ***/
function doGet(e) {
  const action = e && e.parameter && e.parameter.action;
  const cb     = e && e.parameter && e.parameter.callback;

  // ตรวจสอบว่า Deploy เวอร์ชันใหม่แล้ว + ดูชื่อแท็บทั้งหมด (เปิด ...exec?action=ping)
  if (action === 'ping') {
    const names = SpreadsheetApp.getActiveSpreadsheet().getSheets().map(function (s) { return s.getName(); });
    return jsonOut_({ ok: true, version: 'fillYear-v2', sheets: names }, cb);
  }

  if (action === 'options') {
    const o = readOptions_();
    o.targets = readTargets_();
    return jsonOut_(o, cb);
  }

  const sh = sheet_();
  const data = sh.getDataRange().getValues();
  data.shift();
  const rows = data.map(function (row) {
    return {
      source: row[0], dataAddedDate: dstr_(row[1]), riskTitle: row[2], riskDescription: row[3],
      quarter: row[4], likelihood: row[5], consequence: row[6], riskLevel: row[7],
      prevention: row[8], monitorControl: row[9], mitigation: row[10], qiPlan: row[11],
      riskOwner: row[12], reviewFrequency: row[13], dataLastReview: dstr_(row[14]),
      residualLevel: row[15], riskStatus: row[16], riskId: row[17], target: row[18]
    };
  });
  return jsonOut_({ rows: rows }, cb);
}

/*** ===== อ่านตัวเลือก dropdown จากแท็บ "ตัวเลือก" ===== ***/
function readOptions_() {
  const out = {
    riskId: [], source: [], owner: [], reviewFreq: [], residual: [], status: [],
    prevention: [], monitor: [], mitigation: [], qi: []
  };
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(OPTION_SHEET);
  if (!sh) return out;
  const data = sh.getDataRange().getValues();
  if (!data.length) return out;
  const head = data.shift().map(function (x) { return String(x).trim().toLowerCase(); });
  const map = {
    'risk id': 'riskId', 'riskid': 'riskId',
    'source': 'source',
    'risk owner list': 'owner', 'risk owner': 'owner', 'owner': 'owner',
    'review frequency': 'reviewFreq', 'review freq': 'reviewFreq', 'reviewfreq': 'reviewFreq',
    'residual': 'residual', 'residual risk level': 'residual',
    'risk status': 'status', 'status': 'status',
    'risk transfer & prevention': 'prevention', 'prevention': 'prevention',
    'risk monitor & control': 'monitor', 'monitor & control': 'monitor',
    'risk mitigation': 'mitigation', 'mitigation': 'mitigation',
    'qi plan': 'qi', 'qi': 'qi'
  };
  head.forEach(function (h, ci) {
    const key = map[h]; if (!key) return;
    for (var r = 0; r < data.length; r++) {
      var v = String(data[r][ci] || '').trim();
      if (v) out[key].push(v);
    }
  });
  return out;
}

/*** ===== อ่านเป้าหมายจากแท็บ "เกณฑ์" (A=site, หัวมีเลขปี 25xx) ===== ***/
function readTargets_() {
  const out = {};
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TARGET_SHEET);
  if (!sh) return out;
  const data = sh.getDataRange().getValues();
  if (!data.length) return out;
  const head = data.shift().map(function (x) { return String(x).trim(); });
  // ทุกคอลัมน์ (นอกจาก A) ที่หัวมีเลขปี 25xx = 1 ตัวเลือกเป้าหมาย (key = หัวคอลัมน์เต็ม เช่น "ประเทศ 2568")
  const tgtCols = [];
  head.forEach(function (h, ci) {
    if (ci > 0 && /25\d\d/.test(h)) tgtCols.push({ ci: ci, key: h });
  });
  data.forEach(function (row) {
    const s = String(row[0] || '').trim();
    if (!s) return;
    const o = out[s] = {};
    tgtCols.forEach(function (tc) { o[tc.key] = String(row[tc.ci] || '').trim(); });
  });
  return out;
}

/*** ===== ติดตั้งแท็บ "ทะเบียนตัวชี้วัด" + เติมข้อมูลตั้งต้น (รันครั้งเดียว) =====
 * วิธีใช้: เลือกฟังก์ชัน setupIndicatorSheet ในตัวแก้ไข Apps Script แล้วกด Run
 * ปลอดภัย: ถ้ามีแท็บอยู่แล้วและมีข้อมูล จะไม่เขียนทับ (กันข้อมูลหาย)
 ***/
const INDICATOR_SHEET = 'ทะเบียนตัวชี้วัด';
function setupIndicatorSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(INDICATOR_SHEET);
  if (sh && sh.getLastRow() > 1) {
    SpreadsheetApp.getUi().alert('มีแท็บ "' + INDICATOR_SHEET + '" และมีข้อมูลอยู่แล้ว — ไม่เขียนทับ');
    return;
  }
  if (!sh) sh = ss.insertSheet(INDICATOR_SHEET);
  const rows = [
    ['รหัส','ชื่อตัวชี้วัด','ตัวตั้ง','ตัวหาร','หน่วย','ความถี่','source','metricKey','NQA','HA6','เขต','สปสช.','เป้าหมาย','ผู้รับผิดชอบ','ข้อมูลเดิมที่ใช้'],
    ['IPC-01','อัตราการติดเชื้อในโรงพยาบาล (HAI) รวม','จำนวนครั้งการติดเชื้อ HAI','จำนวนวันนอนผู้ป่วย','/1000 วันนอน','เดือน/ไตรมาส/ปี','auto','hai_rate','✓','✓','✓','✓','<1.08','ICN','ทะเบียน-1'],
    ['IPC-02','อัตราการติดเชื้อทางเดินปัสสาวะจากคาสายสวน (CAUTI)','จำนวนครั้ง CAUTI','จำนวนวันคาสายสวนปัสสาวะ','/1000 วันคาสาย','เดือน/ไตรมาส/ปี','auto','cauti_rate','✓','✓','✓','','<2.92','ICN','ทะเบียน-1 (RATE_META)'],
    ['IPC-03','อัตราการเกิดปอดอักเสบจากการใช้เครื่องช่วยหายใจ (VAP)','จำนวนครั้ง VAP','จำนวนวันใช้เครื่องช่วยหายใจ','/1000 วันเครื่องช่วยหายใจ','เดือน/ไตรมาส/ปี','auto','vap_rate','✓','✓','✓','','<1.34','ICN','ทะเบียน-1 (RATE_META)'],
    ['IPC-04','อัตราการติดเชื้อในกระแสเลือดจากสายสวนหลอดเลือด (CLABSI)','จำนวนครั้ง CLABSI','จำนวนวันคาสายสวนหลอดเลือด','/1000 วันคาสาย','เดือน/ไตรมาส/ปี','auto','clabsi_rate','✓','✓','✓','','<2.19','ICN','ทะเบียน-1 (RATE_META)'],
    ['IPC-05','อัตราการติดเชื้อตำแหน่งผ่าตัด (SSI)','จำนวนครั้ง SSI','จำนวนครั้งการผ่าตัด','ร้อยละ','ไตรมาส/ปี','auto','ssi_rate','✓','✓','✓','','<0.22','ICN','ทะเบียน-1 / Post-discharge'],
    ['IPC-06','ร้อยละการติดเชื้อดื้อยา (MDRO) ต่อการติดเชื้อทั้งหมด','จำนวนเคสติดเชื้อดื้อยา','จำนวนเคสติดเชื้อ HAI ทั้งหมด','ร้อยละ','เดือน/ไตรมาส/ปี','auto','mdr_pct','✓','✓','✓','✓','ลดลงจากปีก่อน','ICN','เมนู MDR เดิม'],
    ['IPC-07','อัตราการติดเชื้อดื้อยา (MDRO) ต่อวันนอน','จำนวนเคสติดเชื้อดื้อยา','จำนวนวันนอนผู้ป่วย','/1000 วันนอน','เดือน/ไตรมาส/ปี','auto','mdro_rate','','✓','✓','','ลดลงจากปีก่อน','ICN','เมนู MDR เดิม'],
    ['IPC-08','จำนวนบุคลากรสัมผัสเลือด/สารคัดหลั่งจากการปฏิบัติหน้าที่','จำนวนเหตุการณ์สัมผัส','-','ครั้ง','ไตรมาส/ปี','external','staff_sharp','✓','✓','','','ลดลงจากปีก่อน','ICN/อาชีวอนามัย','ชีต IC-อุบัติเหตุเจ้าหน้าที่'],
    ['IPC-09','อัตราการล้างมือถูกต้องตามโอกาส (Hand Hygiene Compliance)','จำนวนครั้งที่ล้างมือถูกต้อง','จำนวนโอกาสที่ควรล้างมือ (สุ่มสังเกต)','ร้อยละ','เดือน/ไตรมาส','manual','hh_compliance','✓','✓','✓','✓','≥85%','ICN','เริ่มเก็บใหม่ (5 moments)'],
    ['IPC-10','ความครบถ้วนการปฏิบัติตาม Bundle (CAUTI/VAP/CLABSI)','จำนวนรายการ Bundle ที่ปฏิบัติครบ','จำนวนรายการ Bundle ที่ตรวจทั้งหมด','ร้อยละ','เดือน/ไตรมาส','manual','bundle_compliance','✓','✓','','','≥90%','ICN/หอผู้ป่วย','เริ่มเก็บใหม่ (audit)'],
    ['IPC-11','ความครอบคลุมการเฝ้าระวังการติดเชื้อ (Surveillance Coverage)','จำนวนหน่วย/ผู้ป่วยที่เฝ้าระวัง','จำนวนหน่วย/ผู้ป่วยเป้าหมายทั้งหมด','ร้อยละ','ไตรมาส/ปี','manual','surv_coverage','','✓','','','100%','ICN','เริ่มเก็บใหม่'],
    ['IPC-12','อัตราบุคลากรได้รับวัคซีนป้องกันโรค (ไข้หวัดใหญ่/HBV)','จำนวนบุคลากรที่ได้รับวัคซีนครบ','จำนวนบุคลากรกลุ่มเป้าหมาย','ร้อยละ','ปี','manual','staff_vaccine','','✓','','','≥90%','ICN/อาชีวอนามัย','เริ่มเก็บใหม่'],
    ['IPC-13','ความทันเวลาการสอบสวนและควบคุมการระบาด (Outbreak Response)','จำนวนเหตุการณ์ที่สอบสวนทันเวลา','จำนวนเหตุการณ์ระบาดทั้งหมด','ร้อยละ','ต่อเหตุการณ์/ปี','manual','outbreak_timely','','✓','','','100%','ICN','เริ่มเก็บใหม่'],
    ['IPC-14','ความเหมาะสมการให้ยาปฏิชีวนะป้องกันก่อนผ่าตัด (SAP)','จำนวนการผ่าตัดที่ให้ยาเหมาะสม','จำนวนการผ่าตัดที่ควรได้รับยา','ร้อยละ','ไตรมาส','manual','sap_appropriate','✓','✓','','✓','≥95%','ICN/ศัลยกรรม/เภสัช','เริ่มเก็บใหม่ (ASP)'],
    ['IPC-15','การปฏิบัติตามมาตรฐานสิ่งแวดล้อม/IC Round','จำนวนรายการผ่านเกณฑ์','จำนวนรายการที่ตรวจทั้งหมด','ร้อยละ','ไตรมาส','manual','ic_round','','✓','','','≥90%','ICN','เริ่มเก็บใหม่']
  ];
  sh.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sh.getRange(1, 1, 1, rows[0].length).setFontWeight('bold').setBackground('#ecfdf5');
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, rows[0].length);
  SpreadsheetApp.getUi().alert('สร้างแท็บ "' + INDICATOR_SHEET + '" และเติม ' + (rows.length - 1) + ' ตัวชี้วัดเรียบร้อย');
}
