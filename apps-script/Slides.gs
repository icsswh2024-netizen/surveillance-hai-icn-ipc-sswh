/**
 * Slides.gs — ตัวช่วยสร้างแท็บ "Slides" สำหรับกำหนดเด็คนำเสนอของแดชบอร์ด HAI
 * แต่ละแถว = 1 สไลด์ (หัวข้อหลัก / หัวข้อรอง / ชนิด / พารามิเตอร์ / หมายเหตุ)
 *
 * วิธีใช้:
 *   1) เปิดชีต 0-2569-รายงานการติดเชื้อในโรงพยาบาล → เมนู ส่วนขยาย (Extensions) → Apps Script
 *   2) วางไฟล์นี้ แล้วเลือกฟังก์ชัน createSlidesTab แล้วกด Run (อนุญาตสิทธิ์ครั้งแรก)
 *   3) จะได้แท็บ "Slides" พร้อม "ดรอปดาวน์" ให้เลือก ชนิด และ พารามิเตอร์ ได้ทันที
 *
 * เมื่อกรอกเสร็จ รีเฟรชหน้าเว็บแดชบอร์ด → เมนู สรุปการรายงาน › นำเสนอ จะสร้างเด็คตามแท็บนี้
 * ถ้าลบแท็บ "Slides" ทิ้ง ระบบจะกลับไปใช้เด็ค 81 สไลด์ในตัวเหมือนเดิม
 */

var SLIDES_TAB = 'Slides';

// ชนิดสไลด์ที่เว็บรองรับ (ค่า = ที่เว็บอ่าน, คำอธิบาย = ที่โชว์ให้ผู้ใช้เข้าใจ)
var SLIDE_TYPES = [
  ['cover',          'สไลด์ปก'],
  ['agenda',         'หัวข้อการนำเสนอ (พารามิเตอร์=เลขข้อที่ไฮไลต์)'],
  ['kpi',            'การ์ด KPI ภาพรวม (เคส/อัตรา/MDR/SSI)'],
  ['site-table',     'ตารางตำแหน่งติดเชื้อ + อัตรา'],
  ['trend',          'กราฟแนวโน้มรายเดือน (แท่ง+เส้น)'],
  ['ward-all',       'ตารางทุกหอผู้ป่วย + อัตรา'],
  ['site-rate',      'กราฟอัตรารายเดือนของตำแหน่ง (ต้องมีพารามิเตอร์)'],
  ['site-month-h1',  'อัตราครึ่งปีแรกของตำแหน่ง (ต้องมีพารามิเตอร์)'],
  ['site-month-h2',  'อัตราครึ่งปีหลังของตำแหน่ง (ต้องมีพารามิเตอร์)'],
  ['site-ward',      'ตำแหน่งนั้นแยกรายหอ (ต้องมีพารามิเตอร์)'],
  ['site-org',       'เชื้อก่อโรคของตำแหน่งนั้น (ต้องมีพารามิเตอร์)'],
  ['ssi-wound',      'SSI แยกประเภทแผล (wound class)'],
  ['ssi-anat',       'SSI แยกอวัยวะ/หัตถการ'],
  ['omph-kpi',       'KPI สะดืออักเสบ (Omphalitis)'],
  ['mdr-compare',    'กราฟเทียบ HAI vs MDR'],
  ['mdr-kpi',        'KPI เชื้อดื้อยา'],
  ['mdr-org',        'เชื้อดื้อยาที่พบบ่อย'],
  ['mdr-site',       'MDR แยกตามตำแหน่ง'],
  ['mdr-of',         'จำนวน MDR ของตำแหน่งหนึ่ง (ต้องมีพารามิเตอร์)'],
  ['device',         'ตารางอุปกรณ์ (วันคา/เฉลี่ยต่อคน)'],
  ['device-cath',    'KPI สายสวนปัสสาวะ'],
  ['device-ven',     'KPI เครื่องช่วยหายใจ'],
  ['org-top',        'เชื้อที่พบบ่อย + เชื้อดื้อยา (คู่กัน)'],
  ['needle',         'อุบัติการณ์เข็มทิ่มตำรายไตรมาส'],
  ['summary',        'สรุปผล'],
  ['placeholder',    'สไลด์หัวข้อว่าง (ใช้ช่องหมายเหตุเป็นเนื้อหา)']
];

// ตำแหน่งติดเชื้อที่ใช้เป็นพารามิเตอร์ของชนิด site-* และ mdr-of
var SLIDE_PARAMS = ['VAP', 'CAUTI', 'CLABSI', 'SSI', 'HAP', 'BSI', 'Omphalitis', 'SSTI', 'UTI'];

// แถวตัวอย่างเริ่มต้น (ลำดับ, หัวข้อหลัก, หัวข้อรอง, ชนิด, พารามิเตอร์, หมายเหตุ)
var SLIDE_SAMPLE = [
  [1,  'บทนำ',   'โรงพยาบาลศรีสังวรสุโขทัย',            'cover',       '',    ''],
  [2,  'บทนำ',   'หัวข้อการนำเสนอ',                      'agenda',      '',    ''],
  [3,  'ภาพรวม', 'สถานการณ์การติดเชื้อในโรงพยาบาล',     'kpi',         '',    ''],
  [4,  'ภาพรวม', 'แนวโน้มการติดเชื้อรายเดือน',           'trend',       '',    ''],
  [5,  'ภาพรวม', 'ตำแหน่งที่พบการติดเชื้อ',              'site-table',  '',    ''],
  [6,  'รายหอ',  'อัตราการติดเชื้อรายหอผู้ป่วย',          'ward-all',    '',    ''],
  [7,  'VAP',    'อัตรา VAP รายเดือน',                   'site-rate',   'VAP', ''],
  [8,  'VAP',    'VAP แยกรายหอผู้ป่วย',                   'site-ward',   'VAP', ''],
  [9,  'CAUTI',  'อัตรา CAUTI รายเดือน',                 'site-rate',   'CAUTI', ''],
  [10, 'ดื้อยา', 'เทียบการติดเชื้อกับเชื้อดื้อยา',        'mdr-compare', '',    ''],
  [11, 'ดื้อยา', 'เชื้อดื้อยาที่พบบ่อย',                  'mdr-org',     '',    ''],
  [12, 'อื่นๆ',  'การติดเชื้อโคโรนาไวรัส (COVID-19)',    'placeholder', '',    'ยังไม่มีข้อมูลในระบบ'],
  [13, 'สรุป',   'สรุปผลและข้อเสนอแนะ',                  'summary',     '',    '']
];

/** สร้าง/รีเซ็ตแท็บ Slides พร้อมดรอปดาวน์ */
function createSlidesTab() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SLIDES_TAB);
  if (sh) {
    var ui = SpreadsheetApp.getUi();
    var resp = ui.alert('พบแท็บ "' + SLIDES_TAB + '" อยู่แล้ว',
      'ต้องการล้างและสร้างใหม่หรือไม่? (ข้อมูลเดิมจะหาย)', ui.ButtonSet.YES_NO);
    if (resp !== ui.Button.YES) return;
    sh.clear();
    sh.clearDataValidations();
  } else {
    sh = ss.insertSheet(SLIDES_TAB);
  }

  // หัวคอลัมน์
  var header = ['ลำดับ', 'หัวข้อหลัก', 'หัวข้อรอง', 'ชนิด', 'พารามิเตอร์', 'หมายเหตุ'];
  sh.getRange(1, 1, 1, header.length).setValues([header])
    .setFontWeight('bold').setBackground('#0f766e').setFontColor('#ffffff')
    .setHorizontalAlignment('center');
  sh.setFrozenRows(1);

  // แถวตัวอย่าง
  sh.getRange(2, 1, SLIDE_SAMPLE.length, header.length).setValues(SLIDE_SAMPLE);

  // ---- ดรอปดาวน์ "ชนิด" (คอลัมน์ D) ----
  var typeValues = SLIDE_TYPES.map(function (t) { return t[0]; });
  var typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(typeValues, true)
    .setAllowInvalid(false)
    .setHelpText('เลือกชนิดสไลด์ที่ต้องการแสดง')
    .build();
  sh.getRange(2, 4, 500, 1).setDataValidation(typeRule);

  // ---- ดรอปดาวน์ "พารามิเตอร์" (คอลัมน์ E) ----
  var paramRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(SLIDE_PARAMS, true)
    .setAllowInvalid(true)   // อนุญาตเว้นว่าง/ใส่เลขเองได้ (เช่น agenda, site-org)
    .setHelpText('ตำแหน่งติดเชื้อ (เฉพาะชนิด site-* และ mdr-of) — ชนิดอื่นเว้นว่างได้')
    .build();
  sh.getRange(2, 5, 500, 1).setDataValidation(paramRule);

  // ความกว้างคอลัมน์
  sh.setColumnWidth(1, 55);
  sh.setColumnWidth(2, 130);
  sh.setColumnWidth(3, 300);
  sh.setColumnWidth(4, 150);
  sh.setColumnWidth(5, 120);
  sh.setColumnWidth(6, 260);

  // ---- แผ่นคำอธิบายชนิด (ไว้ดูอ้างอิง) วางด้านขวา เริ่มคอลัมน์ H ----
  sh.getRange(1, 8, 1, 2).setValues([['ชนิด (ค่า)', 'ความหมาย']])
    .setFontWeight('bold').setBackground('#e2e8f0');
  sh.getRange(2, 8, SLIDE_TYPES.length, 2).setValues(SLIDE_TYPES);
  sh.setColumnWidth(8, 150);
  sh.setColumnWidth(9, 420);

  SpreadsheetApp.getActiveSpreadsheet().toast('สร้างแท็บ "' + SLIDES_TAB + '" พร้อมดรอปดาวน์เรียบร้อย', 'สำเร็จ', 5);
}
