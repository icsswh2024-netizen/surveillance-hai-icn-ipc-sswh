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
  'Residual risk level','Risk status','Risk ID'];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) { sh = ss.insertSheet(SHEET_NAME); sh.appendRow(HEADERS); }
  return sh;
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
    if (aoa.length) sh2.getRange(START, 1, aoa.length, COLS).setValues(aoa);
    return jsonOut_({ ok: true, filled: aoa.length, sheet: name });
  }

  const sh = sheet_();
  sh.clear();
  sh.appendRow(HEADERS);
  (body.rows || []).forEach(function (r) {
    sh.appendRow([r.source, r.dataAddedDate, r.riskTitle, r.riskDescription, r.quarter,
      r.likelihood, r.consequence, r.riskLevel, r.prevention, r.monitorControl,
      r.mitigation, r.qiPlan, r.riskOwner, r.reviewFrequency, r.dataLastReview,
      r.residualLevel, r.riskStatus, r.riskId]);
  });
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
      source: row[0], dataAddedDate: row[1], riskTitle: row[2], riskDescription: row[3],
      quarter: row[4], likelihood: row[5], consequence: row[6], riskLevel: row[7],
      prevention: row[8], monitorControl: row[9], mitigation: row[10], qiPlan: row[11],
      riskOwner: row[12], reviewFrequency: row[13], dataLastReview: row[14],
      residualLevel: row[15], riskStatus: row[16], riskId: row[17]
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
