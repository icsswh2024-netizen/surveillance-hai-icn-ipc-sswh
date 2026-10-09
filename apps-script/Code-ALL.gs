/**
 * โค้ดรวม (ไฟล์เดียว) สำหรับชีต "รายงานการติดเชื้อในโรงพยาบาล"
 * ใช้ doGet/doPost ตัวเดียว route ตาม action — รองรับทุกเมนูของเว็บ:
 *   • Risk register : doPost (save / fillYear) · doGet (rows / options / ping)
 *   • แบบบันทึกผู้ป่วย HAI : doPost action="addCase"  → แท็บ "ทะเบียน-1"
 *   • แก้ไข Dropdown      : doPost action="setList"  → แท็บ "List-ทะเบียน-1"
 *
 * ⚠️ ใช้แทน "รหัส.gs" เดิมทั้งไฟล์ (อย่าวางเพิ่มเป็นไฟล์ที่สอง เพราะ doGet/doPost จะซ้ำ)
 * ติดตั้ง: วางทับโค้ดเดิม → บันทึก → Deploy → Manage deployments → Edit → New version → Deploy
 *         URL /exec เดิมใช้ได้ต่อ · ใช้ URL เดียวกันกับทุกเมนู (Risk / แบบบันทึก / แก้ Dropdown)
 */

/*** ===== ตั้งค่า ===== ***/
const SHEET_NAME   = 'RiskRegister';
const OPTION_SHEET = 'ตัวเลือก';
const TARGET_SHEET = 'เกณฑ์';
const REC_SHEET      = 'ทะเบียน-1';
const REC_LIST_SHEET = 'List-ทะเบียน-1';

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
function dstr_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return v;
}
function jsonOut_(obj, cb) {
  const out = JSON.stringify(obj);
  if (cb) return ContentService.createTextOutput(cb + '(' + out + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(out).setMimeType(ContentService.MimeType.JSON);
}
function recNorm_(s) { return String(s == null ? '' : s).toLowerCase().replace(/[\s.]/g, ''); }

/*** ===== doGet : ping / options / rows ===== ***/
function doGet(e) {
  const action = e && e.parameter && e.parameter.action;
  const cb     = e && e.parameter && e.parameter.callback;

  if (action === 'ping') {
    const names = SpreadsheetApp.getActiveSpreadsheet().getSheets().map(function (s) { return s.getName(); });
    // มี marker ครบทุกเมนู: Risk (version) + บันทึก/Dropdown (record/lists)
    return jsonOut_({ ok: true, version: 'fillYear-v2', record: 'record-v1', lists: true, sheets: names }, cb);
  }
  if (action === 'options') {
    const o = readOptions_(); o.targets = readTargets_();
    return jsonOut_(o, cb);
  }
  // ค่าเริ่มต้น: ส่งแถว RiskRegister
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

/*** ===== doPost : route ตาม action ===== ***/
function doPost(e) {
  if (!e || !e.postData) return jsonOut_({ ok: false, note: 'เรียกผ่าน Web app POST เท่านั้น' });
  const body = JSON.parse(e.postData.contents);

  if (body.action === 'addCase') return recAddCase_(body);
  if (body.action === 'setList') return recSetList_(body);

  // กรอกลงแท็บ "ปีงบ 25xx" ตามแบบฟอร์ม RM ทางการ
  if (body.action === 'fillYear') {
    const name = 'ปีงบ ' + body.year;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh2 = ss.getSheetByName(name);
    if (!sh2) return jsonOut_({ ok: false, note: 'ไม่พบแท็บ "' + name + '"' });
    const aoa = body.aoa || [];
    const START = 4, COLS = 17;
    const last = sh2.getLastRow();
    if (last >= START) sh2.getRange(START, 1, last - START + 1, COLS).clearContent();
    if (aoa.length) {
      sh2.getRange(START, 1, aoa.length, COLS).setValues(aoa);
      sh2.getRange(START, 2, aoa.length, 1).setNumberFormat('dd/mm/yyyy');
      sh2.getRange(START, 15, aoa.length, 1).setNumberFormat('dd/mm/yyyy');
    }
    return jsonOut_({ ok: true, filled: aoa.length, sheet: name });
  }

  // ค่าเริ่มต้น: บันทึก RiskRegister (เขียนทับทั้งแท็บ)
  const sh = sheet_();
  sh.clear();
  sh.appendRow(HEADERS);
  (body.rows || []).forEach(function (r) {
    sh.appendRow([r.source, r.dataAddedDate, r.riskTitle, r.riskDescription, r.quarter,
      r.likelihood, r.consequence, r.riskLevel, r.prevention, r.monitorControl,
      r.mitigation, r.qiPlan, r.riskOwner, r.reviewFrequency, r.dataLastReview,
      r.residualLevel, r.riskStatus, r.riskId, r.target]);
  });
  var nRows = (body.rows || []).length;
  if (nRows) {
    sh.getRange(2, 2, nRows, 1).setNumberFormat('dd/mm/yyyy');
    sh.getRange(2, 15, nRows, 1).setNumberFormat('dd/mm/yyyy');
  }
  return jsonOut_({ ok: true, saved: (body.rows || []).length });
}

/*** ===== แบบบันทึกผู้ป่วย HAI → แท็บ "ทะเบียน-1" ===== ***/
function recAddCase_(body) {
  const recs = Array.isArray(body.records) ? body.records : [body.record || {}];
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REC_SHEET);
  if (!sh) return jsonOut_({ ok: false, note: 'ไม่พบแท็บ "' + REC_SHEET + '"' });
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
  return jsonOut_({ ok: true, action: 'addCase', added: added, matched: matched, lastRow: sh.getLastRow() });
}

/*** ===== แก้ไข Dropdown → แท็บ "List-ทะเบียน-1" ===== ***/
function recSetList_(body) {
  const header = String(body.header || '').trim();
  if (!header) return jsonOut_({ ok: false, note: 'ไม่มีชื่อคอลัมน์' });
  const values = Array.isArray(body.values) ? body.values.map(function (v) { return [v]; }) : [];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(REC_LIST_SHEET);
  if (!sh) sh = ss.insertSheet(REC_LIST_SHEET);
  const lastCol = Math.max(1, sh.getLastColumn());
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (x) { return String(x).trim(); });
  let col = -1;
  for (var i = 0; i < headers.length; i++) { if (headers[i] === header) { col = i + 1; break; } }
  if (col < 0) {
    if (!body.create) return jsonOut_({ ok: false, note: 'ไม่พบคอลัมน์ "' + header + '"' });
    col = (headers.join('') === '' ? 1 : lastCol + 1);
    sh.getRange(1, col, 1, 1).setValue(header);
  }
  const maxRow = sh.getMaxRows();
  if (maxRow > 1) sh.getRange(2, col, maxRow - 1, 1).clearContent();
  if (values.length) sh.getRange(2, col, values.length, 1).setValues(values);
  return jsonOut_({ ok: true, action: 'setList', header: header, count: values.length, col: col });
}

/*** ===== อ่านตัวเลือก dropdown Risk จากแท็บ "ตัวเลือก" ===== ***/
function readOptions_() {
  const out = { riskId: [], source: [], owner: [], reviewFreq: [], residual: [], status: [], prevention: [], monitor: [], mitigation: [], qi: [] };
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(OPTION_SHEET);
  if (!sh) return out;
  const data = sh.getDataRange().getValues();
  if (!data.length) return out;
  const head = data.shift().map(function (x) { return String(x).trim().toLowerCase(); });
  const map = {
    'risk id': 'riskId', 'riskid': 'riskId', 'source': 'source',
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
    for (var r = 0; r < data.length; r++) { var v = String(data[r][ci] || '').trim(); if (v) out[key].push(v); }
  });
  return out;
}

/*** ===== อ่านเป้าหมายจากแท็บ "เกณฑ์" ===== ***/
function readTargets_() {
  const out = {};
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TARGET_SHEET);
  if (!sh) return out;
  const data = sh.getDataRange().getValues();
  if (!data.length) return out;
  const head = data.shift().map(function (x) { return String(x).trim(); });
  const tgtCols = [];
  head.forEach(function (h, ci) { if (ci > 0 && /25\d\d/.test(h)) tgtCols.push({ ci: ci, key: h }); });
  data.forEach(function (row) {
    const s = String(row[0] || '').trim();
    if (!s) return;
    const o = out[s] = {};
    tgtCols.forEach(function (tc) { o[tc.key] = String(row[tc.ci] || '').trim(); });
  });
  return out;
}
