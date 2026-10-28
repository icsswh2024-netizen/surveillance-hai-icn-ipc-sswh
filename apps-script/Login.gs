/**
 * ระบบเข้าสู่ระบบ (Login) — ตรวจรหัสผ่านฝั่ง Apps Script
 * ──────────────────────────────────────────────────────────────
 * หลักการ:
 *   - เก็บผู้ใช้ในแท็บ "Users" ของไฟล์ 0-2569-รายงานการติดเชื้อในโรงพยาบาล
 *   - หน้าเว็บส่งเฉพาะ "แฮช SHA-256(password + SALT)" มา ไม่ส่งรหัสจริง
 *   - Apps Script คำนวณแฮชของรหัสในชีตแล้วเทียบ → รหัสจริงไม่วิ่งผ่านเน็ต
 *
 * ติดตั้ง (ครั้งเดียว):
 *   1) เปิดไฟล์ชีต (1KEK2bv...) → Extensions → Apps Script
 *   2) เพิ่มไฟล์นี้ (วางโค้ด) → บันทึก
 *   3) Deploy → New deployment → ประเภท Web app
 *        - Execute as: Me
 *        - Who has access: Anyone
 *      คัดลอก URL /exec
 *   4) นำ URL ไปวางในหน้าเว็บ (หน้าเข้าสู่ระบบ → "ตั้งค่า URL Apps Script")
 *
 * แท็บ "Users" (แถวแรก = หัวคอลัมน์) — รองรับ 2 แบบ:
 *
 *   แบบ A) คอลัมน์ "menus" เดียว (พิมพ์รายการเมนูคั่นด้วย comma หรือ "all")
 *   | username | password | ชื่อ-สกุล | role  | menus                       |
 *   | admin    | 1234     | ICN ก.   | admin | all                         |
 *   | icwn1    | abcd     | พยาบาล ข | user  | dashboard,wards,org,record  |
 *
 *   แบบ B) คอลัมน์แยกทุกเมนู + ติ๊ก checkbox (แนะนำ — ง่ายกว่า)
 *   | username | password | ชื่อ-สกุล | role | all | dashboard | wards | mdr | ... |
 *   | admin    | 1234     | ICN ก.   | admin| ✔   |           |       |     |     |
 *   | icwn1    | abcd     | พยาบาล ข | user |     | ✔         | ✔     |     |     |
 *     - ติ๊ก checkbox (เมนู Insert → Checkbox ในชีต) คอลัมน์ไหน = ให้สิทธิ์เมนูนั้น
 *     - ติ๊กคอลัมน์ "all" = เข้าได้ทุกเมนู
 *     - ค่าที่นับว่า "ติ๊ก": TRUE / ✔ / 1 / yes / x
 *
 *   key เมนู (ใช้เป็นหัวคอลัมน์แบบ B): dashboard, wards, mdr, service, indicator,
 *     analysis, report, org, patient, compare, surv, special, postdc, control, risk, record
 *
 * ⚠️ หมายเหตุความปลอดภัย: เป็นการกั้นเมนูระดับ UI สำหรับใช้ภายในหน่วยงาน
 *    ไม่ใช่ระบบความปลอดภัยระดับสูง และ "ไม่" ปกป้องข้อมูลในชีตที่แชร์ public
 */

var LOGIN_CFG = {
  SHEET: 'Users',
  SALT: 'sswh-ic-2569'   // ⚠️ ต้องตรงกับค่า SALT ในหน้าเว็บ (index.html)
};

// key เมนูทั้งหมด (ใช้เป็นหัวคอลัมน์ checkbox แบบ B)
var LOGIN_MENU_KEYS = ['dashboard', 'wards', 'mdr', 'service', 'indicator', 'analysis',
  'report', 'org', 'patient', 'compare', 'surv', 'special', 'postdc', 'control', 'risk', 'record'];

// นับว่า "ติ๊ก/อนุญาต" ไหม (รองรับ checkbox TRUE, ✔, 1, yes, x)
function smChecked_(v) {
  if (v === true) return true;
  var s = String(v == null ? '' : v).trim().toLowerCase();
  return s === 'true' || s === '1' || s === 'yes' || s === 'y' || s === 'x' ||
    s === '✓' || s === '✔' || s === 'true()';
}

function doGet(e) {
  var p = (e && e.parameter) || {};
  var cb = p.callback || 'callback';
  var out;
  try {
    if (p.action === 'login') out = smLogin_(p.u, p.h);
    else if (p.action === 'ping') out = { ok: true, login: 'login-v1' };
    else out = { ok: false, error: 'unknown action' };
  } catch (err) {
    out = { ok: false, error: String(err) };
  }
  return ContentService
    .createTextOutput(cb + '(' + JSON.stringify(out) + ')')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function smLogin_(u, h) {
  u = String(u == null ? '' : u).trim().toLowerCase();
  h = String(h == null ? '' : h).trim().toLowerCase();
  if (!u || !h) return { ok: false };

  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOGIN_CFG.SHEET);
  if (!sh) return { ok: false, error: 'ไม่พบแท็บ "' + LOGIN_CFG.SHEET + '"' };

  var data = sh.getDataRange().getValues();
  if (data.length < 2) return { ok: false };
  var hdr = data[0].map(function (x) { return String(x).trim().toLowerCase(); });
  var ci = {
    user: hdr.indexOf('username'),
    pass: hdr.indexOf('password'),
    name: hdr.indexOf('ชื่อ-สกุล'),
    role: hdr.indexOf('role'),
    menus: hdr.indexOf('menus')
  };
  if (ci.name < 0) ci.name = hdr.indexOf('name');
  if (ci.user < 0 || ci.pass < 0) return { ok: false, error: 'หัวคอลัมน์ต้องมี username และ password' };

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (String(row[ci.user] == null ? '' : row[ci.user]).trim().toLowerCase() !== u) continue;
    var stored = String(row[ci.pass] == null ? '' : row[ci.pass]);
    var calc = sha256Hex_(stored + LOGIN_CFG.SALT);
    if (calc === h) {
      return {
        ok: true,
        username: u,
        name: ci.name >= 0 ? String(row[ci.name] || '') : u,
        role: ci.role >= 0 ? String(row[ci.role] || '') : '',
        menus: smResolveMenus_(hdr, row, ci)
      };
    }
    return { ok: false }; // พบ username แต่รหัสไม่ตรง
  }
  return { ok: false }; // ไม่พบ username
}

// รวมสิทธิ์เมนู: คอลัมน์ "all" ติ๊ก → "all"; ไม่งั้นรวมคอลัมน์เมนูที่ติ๊ก;
// ถ้าไม่มีคอลัมน์ checkbox เลย ใช้คอลัมน์ "menus" (ข้อความ) แทน
function smResolveMenus_(hdr, row, ci) {
  var allI = hdr.indexOf('all');
  if (allI >= 0 && smChecked_(row[allI])) return 'all';
  var allowed = [];
  for (var k = 0; k < LOGIN_MENU_KEYS.length; k++) {
    var key = LOGIN_MENU_KEYS[k];
    var col = hdr.indexOf(key);
    if (col >= 0 && smChecked_(row[col])) allowed.push(key);
  }
  if (allowed.length) return allowed.join(',');
  // ไม่มี checkbox → ใช้คอลัมน์ menus (ข้อความ) ถ้ามี
  if (ci.menus >= 0) { var t = String(row[ci.menus] || '').trim(); if (t) return t; }
  return 'all';
}

function sha256Hex_(s) {
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8);
  var hex = '';
  for (var i = 0; i < raw.length; i++) {
    var b = raw[i] < 0 ? raw[i] + 256 : raw[i];
    var x = b.toString(16);
    hex += (x.length === 1 ? '0' : '') + x;
  }
  return hex;
}
