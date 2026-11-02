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
  ['rate-year',      'กราฟอัตรา รายปี (แกน X=ปี, Y=อัตรา) — พารามิเตอร์=ตำแหน่ง/HAI'],
  ['rate-quarter',   'กราฟอัตรา รายไตรมาส (แกน X=ไตรมาส, Y=อัตรา)'],
  ['rate-month',     'กราฟอัตรา รายเดือน (แกน X=เดือน, Y=อัตรา)'],
  ['hist-trend',     'ย้อนหลังหลายปี: กราฟเส้น (พารามิเตอร์=ชื่อแท็บ เช่น HAI, Site, VAP, CAUTI, SSI)'],
  ['hist-bar',       'ย้อนหลังหลายปี: กราฟแท่ง (พารามิเตอร์=ชื่อแท็บ)'],
  ['hist-group',     'ย้อนหลังเทียบรายปี: แท่งกลุ่ม (พารามิเตอร์=แท็บ เช่น HAI-ward, VAP-M, VAP-W, VAP-G)'],
  ['hist-live',      'เส้นเดียว ปีเก่า+ปีปัจจุบัน (พารามิเตอร์=HAI/MDR/VAP/CAUTI/SSI/CLABSI...)'],
  ['multi',          'หลายกราฟในสไลด์เดียว (พารามิเตอร์=รายการคั่น ; เช่น hist-trend:HAI ; hist-group:VAP-M)'],
  ['count-year',     'กราฟจำนวนเคส รายปี (แกน X=ปี, Y=จำนวน)'],
  ['count-quarter',  'กราฟจำนวนเคส รายไตรมาส (แกน X=ไตรมาส, Y=จำนวน)'],
  ['count-month',    'กราฟจำนวนเคส รายเดือน (แกน X=เดือน, Y=จำนวน)'],
  ['pct-year',       'กราฟร้อยละ รายปี (แกน Y=%) — พารามิเตอร์=MDR หรือตำแหน่ง'],
  ['pct-quarter',    'กราฟร้อยละ รายไตรมาส (แกน Y=%)'],
  ['pct-month',      'กราฟร้อยละ รายเดือน (แกน Y=%)'],
  ['donut-site',     'โดนัท สัดส่วนตำแหน่งติดเชื้อ'],
  ['pie-site',       'วงกลม สัดส่วนตำแหน่งติดเชื้อ'],
  ['donut-org',      'โดนัท สัดส่วนเชื้อก่อโรค (พารามิเตอร์=ตำแหน่ง/เว้นว่าง=ทุกตำแหน่ง)'],
  ['pie-org',        'วงกลม สัดส่วนเชื้อก่อโรค'],
  ['bar-site',       'แท่ง จำนวนเคสแต่ละตำแหน่งติดเชื้อ'],
  ['hbar-ward',      'แท่งแนวนอน อัตรารายหอ (พารามิเตอร์=HAI/ตำแหน่ง)'],
  ['hbar-org',       'แท่งแนวนอน เชื้อก่อโรค (พารามิเตอร์=ตำแหน่ง)'],
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
// ===== รายการสำหรับโหมดประกอบเอง =====
var SLIDE_VARS = ['HAI', 'MDR', 'VAP', 'CAUTI', 'CLABSI', 'SSI', 'HAP', 'BSI', 'Omphalitis', 'UTI',
  'หอ', 'เชื้อ', 'กลุ่มงาน',                                  // มิติเปรียบเทียบ (รายหอ/เชื้อก่อโรค/กลุ่มงาน)
  'VAP,CAUTI,SSI', 'VAP,CAUTI,CLABSI,SSI',                     // หลายตัวแปรในกราฟเดียว (คั่นด้วย ,)
  'Site', 'HAI-ward', 'VAP-M', 'VAP-W', 'VAP-G', 'VAP-IPD',    // ชื่อแท็บย้อนหลัง (ดึงทั้งแท็บ)
  'CAUTI-M', 'CAUTI-W', 'CAUTI-G', 'SSI-C', 'SSI-W', 'SSI-G', 'SSI-O', 'OMP-W', 'BSI-HDU'];   // ตัวแปร
var SLIDE_METRICS = ['อัตรา', 'จำนวน', 'ร้อยละ'];        // ค่าที่แสดง
var SLIDE_PERIODS = ['ปี', 'ไตรมาส', 'เดือน'];           // ช่วงเวลา
var SLIDE_HISTORY = ['ใช้', 'ไม่'];                       // ใช้ข้อมูลเก่า
var SLIDE_SHAPES = ['เส้น', 'แท่ง', 'พื้นที่', 'ผสม(แท่ง+เส้น)', 'โดนัท', 'วงกลม', 'แท่งแนวนอน'];   // รูปแบบกราฟ (ไทย; เว็บรองรับทั้งไทย/อังกฤษ)
var SLIDE_PARAMS = ['HAI', 'MDR', 'VAP', 'CAUTI', 'CLABSI', 'SSI', 'HAP', 'BSI', 'Omphalitis', 'SSTI', 'UTI',
  // ชื่อแท็บในชีตย้อนหลัง "สรุปรายปี-HAI" (ใช้กับชนิด hist-*)
  'Site', 'HAI-ward', 'VAP-M', 'VAP-W', 'VAP-G', 'VAP-IPD', 'CAUTI-M', 'CAUTI-W', 'CAUTI-G',
  'SSI-C', 'SSI-W', 'SSI-G', 'OMP', 'OMP-W', 'CLABSI'];

// แถวตัวอย่างเริ่มต้น = โครงนำเสนอเต็ม (ลำดับ, หัวข้อหลัก, หัวข้อรอง, ชนิด, พารามิเตอร์, หมายเหตุ)
// สะท้อนเด็ค 81 สไลด์ในตัว — แก้/ลบ/สลับแถวได้ตามต้องการ
var S2 = 'อัตราการติดเชื้อในโรงพยาบาล', S3 = 'การติดเชื้อ VAP', S4 = 'การติดเชื้อ CAUTI',
    S5 = 'SSI / Omphalitis', S6 = 'เชื้อดื้อยาต้านจุลชีพ (MDR)', S7 = 'อุบัติเหตุในเจ้าหน้าที่',
    S8 = 'การติดเชื้อโคโรนาไวรัส (COVID-19)', S9 = 'วัณโรคในเจ้าหน้าที่',
    S10 = 'IC Round และการทำความสะอาดมือ', S11 = 'การทบทวนความรู้ของบุคลากร',
    S12 = 'มาตรการป้องกัน CAUTI', S13 = 'มาตรการป้องกัน VAP', S14 = 'อุบัติเหตุเข็มทิ่มตำ';
var CANVA = 'มีในเด็ค Canva (ไม่มีข้อมูลในระบบ)';
var SLIDE_SAMPLE = [
  // 1. บทนำ
  ['บทนำ', 'โรงพยาบาลศรีสังวรสุโขทัย', 'cover', '', ''],
  ['บทนำ', 'หัวข้อการนำเสนอ', 'agenda', '', ''],
  ['ภาพรวม', 'สถานการณ์การติดเชื้อในโรงพยาบาล (ภาพรวม)', 'kpi', '', ''],
  // 2. อัตราการติดเชื้อในโรงพยาบาล
  [S2, 'เปรียบเทียบอัตราการติดเชื้อในโรงพยาบาลโดยรวม', 'trend', '', ''],
  [S2, 'อัตราการติดเชื้อ จำแนกตามตำแหน่งการติดเชื้อ', 'site-table', '', ''],
  [S2, 'อัตราการติดเชื้อ จำแนกตามหอผู้ป่วยอายุรกรรม', 'placeholder', '', 'ดูรวมในสไลด์ "จำแนกตามหอผู้ป่วยโดยรวม"'],
  [S2, 'อัตราการติดเชื้อ จำแนกตามหอผู้ป่วยศัลยกรรม', 'placeholder', '', 'ดูรวมในสไลด์ "จำแนกตามหอผู้ป่วยโดยรวม"'],
  [S2, 'อัตราการติดเชื้อ จำแนกตามหอผู้ป่วยผู้คลอด กุมารเวชกรรม และ NICU', 'placeholder', '', 'ดูรวมในสไลด์ "จำแนกตามหอผู้ป่วยโดยรวม"'],
  [S2, 'อัตราการติดเชื้อ จำแนกตามหอผู้ป่วยพิเศษและ Cohort', 'placeholder', '', 'ดูรวมในสไลด์ "จำแนกตามหอผู้ป่วยโดยรวม"'],
  [S2, 'อัตราการติดเชื้อ จำแนกตามหอผู้ป่วยโดยรวม', 'ward-all', '', ''],
  // 3. VAP
  [S3, 'อัตราการติดเชื้อ VAP (ปอดอักเสบจากเครื่องช่วยหายใจ)', 'site-rate', 'VAP', ''],
  [S3, 'อัตราการติดเชื้อ VAP จำแนกตามแผนก', 'placeholder', '', CANVA],
  [S3, 'อัตราการติดเชื้อ VAP รายเดือน (ต.ค. – มี.ค.)', 'site-month-h1', 'VAP', ''],
  [S3, 'อัตราการติดเชื้อ VAP รายเดือน (เม.ย. – ก.ย.)', 'site-month-h2', 'VAP', ''],
  [S3, 'อัตราการติดเชื้อ VAP จำแนกตามหอผู้ป่วย', 'site-ward', 'VAP', ''],
  [S3, 'ร้อยละของเชื้อที่เป็นสาเหตุของ VAP', 'site-org', 'VAP', ''],
  // 4. CAUTI
  [S4, 'อัตราการติดเชื้อ CAUTI (ทางเดินปัสสาวะจากสายสวน)', 'site-rate', 'CAUTI', ''],
  [S4, 'อัตราการติดเชื้อ CAUTI รายเดือน (ต.ค. – มี.ค.)', 'site-month-h1', 'CAUTI', ''],
  [S4, 'อัตราการติดเชื้อ CAUTI รายเดือน (เม.ย. – ก.ย.)', 'site-month-h2', 'CAUTI', ''],
  [S4, 'อัตราการติดเชื้อ CAUTI จำแนกตามหอผู้ป่วย', 'site-ward', 'CAUTI', ''],
  [S4, 'ร้อยละของเชื้อที่เป็นสาเหตุของ CAUTI', 'site-org', 'CAUTI', ''],
  // 5. SSI / Omphalitis
  [S5, 'อัตราการติดเชื้อแผลผ่าตัด SSI', 'site-rate', 'SSI', ''],
  [S5, 'อัตราการติดเชื้อ SSI จำแนกตามประเภทแผลผ่าตัด', 'ssi-wound', '', ''],
  [S5, 'อัตราการติดเชื้อ SSI จำแนกตามหอผู้ป่วย', 'site-ward', 'SSI', ''],
  [S5, 'จำนวนครั้งและประเภทการผ่าตัดที่เกิด SSI', 'ssi-anat', '', ''],
  [S5, 'ร้อยละของเชื้อที่เป็นสาเหตุของ SSI', 'site-org', 'SSI', ''],
  [S5, 'ภาวะ Omphalitis (การติดเชื้อบริเวณสะดือทารก)', 'omph-kpi', '', ''],
  [S5, 'อัตราการติดเชื้อ Omphalitis จำแนกตามหอผู้ป่วย', 'site-ward', 'Omphalitis', ''],
  // 6. MDR
  [S6, 'หัวข้อถัดไป: การติดเชื้อดื้อยาต้านจุลชีพ', 'agenda', '2', ''],
  [S6, 'เปรียบเทียบอัตราการติดเชื้อกับการติดเชื้อดื้อยา', 'mdr-compare', '', ''],
  [S6, 'อัตราการติดเชื้อดื้อยาต้านจุลชีพในโรงพยาบาล', 'mdr-kpi', '', ''],
  [S6, 'ร้อยละของเชื้อที่เป็นสาเหตุของการติดเชื้อดื้อยา', 'mdr-org', '', ''],
  [S6, 'การติดเชื้อดื้อยา จำแนกตามตำแหน่งการติดเชื้อ', 'mdr-site', '', ''],
  [S6, 'อัตราการติดเชื้อดื้อยา VAP', 'mdr-of', 'VAP', ''],
  [S6, 'อัตราการติดเชื้อดื้อยา CAUTI', 'mdr-of', 'CAUTI', ''],
  [S6, 'ระบบเฝ้าระวัง: การรายงานผู้ป่วยสงสัย และ Device Day', 'device', '', ''],
  // 7. อุบัติเหตุในเจ้าหน้าที่
  [S7, 'หัวข้อถัดไป: อุบัติการณ์ในเจ้าหน้าที่', 'agenda', '3', ''],
  [S7, 'จำนวนครั้งการเกิดอุบัติเหตุในเจ้าหน้าที่', 'needle', '', ''],
  [S7, 'จำแนกตามวิชาชีพ เหตุการณ์ และหน่วยงาน', 'placeholder', '', CANVA],
  [S7, 'จำแนกตามวิชาชีพ', 'placeholder', '', CANVA],
  [S7, 'จำแนกตามเหตุการณ์', 'placeholder', '', CANVA],
  [S7, 'จำแนกตามกิจกรรม', 'placeholder', '', CANVA],
  [S7, 'จำแนกตามช่วงเวลาที่เกิด', 'placeholder', '', CANVA],
  [S7, 'จำแนกตามลักษณะการดูแล', 'placeholder', '', CANVA],
  // 8. COVID-19
  [S8, 'หัวข้อถัดไป: การติดเชื้อโคโรนาไวรัส (COVID-19)', 'agenda', '4', ''],
  [S8, 'จำนวนผู้ป่วยติดเชื้อ COVID-19 ในหอผู้ป่วยแยกโรค', 'placeholder', '', CANVA],
  [S8, 'จำนวนผู้ป่วยติดเชื้อ COVID-19 เสียชีวิต ในหอแยกโรค', 'placeholder', '', CANVA],
  [S8, 'จำนวนเจ้าหน้าที่ติดเชื้อ COVID-19 จำแนกตามประเภทงาน', 'placeholder', '', CANVA],
  [S8, 'เจ้าหน้าที่ติดเชื้อ COVID-19 จำแนกตามประเภทงาน', 'placeholder', '', CANVA],
  [S8, 'จำนวนเจ้าหน้าที่ให้บริการที่ติดเชื้อ COVID-19 ปี 2567', 'placeholder', '', CANVA],
  [S8, 'จำนวนเจ้าหน้าที่ทั่วไปที่ติดเชื้อ COVID-19 ปี 2567', 'placeholder', '', CANVA],
  [S8, 'จำนวนเจ้าหน้าที่กลุ่มการพยาบาลที่ติดเชื้อ COVID-19', 'placeholder', '', CANVA],
  [S8, 'จำนวนเจ้าหน้าที่ติดเชื้อ COVID-19 จำแนกตามแหล่งสัมผัส', 'placeholder', '', CANVA],
  // 9. วัณโรค
  [S9, 'หัวข้อถัดไป: วัณโรคในเจ้าหน้าที่', 'agenda', '5', ''],
  [S9, 'จำนวนครั้งการเกิดอุบัติการณ์วัณโรคในเจ้าหน้าที่', 'placeholder', '', CANVA],
  // 10. IC Round / ล้างมือ
  [S10, 'หัวข้อถัดไป: IC Round และการทำความสะอาดมือ', 'agenda', '6', ''],
  [S10, 'ผล IC Round: พื้นที่ ขยะ การทำลายเชื้อ น้ำยา และการแพร่กระจายเชื้อ', 'placeholder', '', CANVA],
  [S10, 'ผล IC Round: PPE, Multiple doses, ผ้าเปื้อน, IV และ TB', 'placeholder', '', CANVA],
  [S10, 'ผล IC Round: VAP, CAUTI และ MDRO', 'placeholder', '', CANVA],
  [S10, 'ผล VAP, CAUTI, MDRO จำแนกตามหอผู้ป่วย', 'placeholder', '', CANVA],
  [S10, 'ผล VAP, CAUTI, MDRO จำแนกตามหอผู้ป่วย (เพิ่มเติม)', 'placeholder', '', CANVA],
  [S10, 'ผลการทำความสะอาดมือ เปรียบเทียบตามโอกาส (5 Moments)', 'placeholder', '', CANVA],
  [S10, 'ผลการปฏิบัติตามขั้นตอนการทำความสะอาดมือ', 'placeholder', '', CANVA],
  // 11. ทบทวนความรู้
  [S11, 'RN: ความรู้การวินิจฉัยการติดเชื้อ — ร้อยละการตอบแบบทดสอบ', 'placeholder', '', CANVA],
  [S11, 'RN: ความรู้การวินิจฉัยการติดเชื้อ — ค่าเฉลี่ยคะแนน', 'placeholder', '', CANVA],
  [S11, 'RN: ความรู้การวินิจฉัยการติดเชื้อ — ระดับความรู้', 'placeholder', '', CANVA],
  [S11, 'PN/NA: การดูแลสิ่งแวดล้อมเชื้อดื้อยา — ร้อยละการตอบแบบทดสอบ', 'placeholder', '', CANVA],
  [S11, 'PN/NA: การดูแลสิ่งแวดล้อมเชื้อดื้อยา — ค่าเฉลี่ยคะแนน', 'placeholder', '', CANVA],
  [S11, 'PN/NA: การดูแลสิ่งแวดล้อมเชื้อดื้อยา — ระดับความรู้', 'placeholder', '', CANVA],
  // 12. มาตรการป้องกัน CAUTI
  [S12, 'มาตรการป้องกัน CAUTI (ภาพรวม)', 'placeholder', '', ''],
  [S12, 'ข้อมูลวันใส่สายสวนปัสสาวะ จำนวนผู้ป่วย และค่าเฉลี่ยวันใส่สาย', 'device-cath', '', ''],
  [S12, 'โปสเตอร์ประชาสัมพันธ์และการเตือนพิจารณาถอดสายสวนปัสสาวะ', 'placeholder', '', '🖼️ รูปภาพ — ดูในเด็ค Canva'],
  [S12, 'บัตรสอดเตือนที่จุด Charge', 'placeholder', '', '🖼️ รูปภาพ — ดูในเด็ค Canva'],
  [S12, 'ตัวอย่างการใช้บัตรสอดเตือนที่จุด Charge', 'placeholder', '', '🖼️ รูปภาพ — ดูในเด็ค Canva'],
  // 13. มาตรการป้องกัน VAP
  [S13, 'ข้อมูลวันใช้เครื่องช่วยหายใจ จำนวนผู้ป่วย และค่าเฉลี่ยวันใช้เครื่อง', 'device-ven', '', ''],
  [S13, 'ค่าเฉลี่ยวันใช้เครื่องช่วยหายใจ', 'device-ven', '', ''],
  // 14. เข็มทิ่มตำ + สรุป
  [S14, 'อุบัติเหตุเข็มทิ่มตำในเจ้าหน้าที่', 'needle', '', ''],
  [S14, 'สรุปผลและข้อเสนอแนะ', 'summary', '', '']
];

// ID ของชีต "0-2569-รายงานการติดเชื้อในโรงพยาบาล" (ใช้สำรองเมื่อรันนอกบริบทชีต)
var SLIDES_SHEET_ID = '1KEK2bvXLDAY7awAapTnF32yZz1qvfqfWXje2ccG2pLc';

/** สร้าง/รีเซ็ตแท็บ Slides พร้อมดรอปดาวน์ (รันได้ทั้งจากเมนู Extensions และจากตัวแก้ไข Apps Script) */
function createSlidesTab() {
  // เปิดชีต: ใช้ไฟล์ที่ผูกอยู่ก่อน ถ้าไม่มีให้เปิดด้วย ID
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) ss = SpreadsheetApp.openById(SLIDES_SHEET_ID);
  // ถ้ามีแท็บเดิม → ล้างแล้วสร้างทับเลย (ไม่ถามยืนยัน เพื่อไม่ให้ค้างรอกล่อง dialog จนหมดเวลา)
  var sh = ss.getSheetByName(SLIDES_TAB);
  if (sh) { sh.clear(); sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearDataValidations(); } else { sh = ss.insertSheet(SLIDES_TAB); }

  // ===== คอลัมน์แบบประกอบเอง (A..K) + ช่องว่าง L + ตารางอ้างอิง M..N =====
  // [ลำดับ, หัวข้อหลัก, หัวข้อรอง, ชนิด, พารามิเตอร์, ตัวแปร, ค่า, ช่วงเวลา, ข้อมูลเก่า, กราฟ, หมายเหตุ]
  var header = ['ลำดับ', 'หัวข้อหลัก', 'หัวข้อรอง', 'ชนิด', 'พารามิเตอร์', 'ตัวแปร', 'ค่า', 'ช่วงเวลา', 'ข้อมูลเก่า', 'กราฟ', 'หมายเหตุ'];
  // แถวประกอบเอง (ไม่ต้องใส่ "ชนิด" — ใส่ ตัวแปร/ค่า/ช่วงเวลา/ข้อมูลเก่า/กราฟ แทน)
  var COMPOSE = [
    // section, title, variable, metric, period, history, shape
    ['ภาพรวม', 'อัตรา HAI หลายปี (เก่า+ปัจจุบัน)', 'HAI', 'อัตรา', 'ปี', 'ใช้', 'เส้น'],
    ['VAP', 'อัตรา VAP รายเดือน', 'VAP', 'อัตรา', 'เดือน', 'ไม่', 'เส้น'],
    ['CAUTI', 'อัตรา CAUTI รายไตรมาส', 'CAUTI', 'อัตรา', 'ไตรมาส', 'ไม่', 'แท่ง'],
    ['ดื้อยา', 'ร้อยละ MDR รายปี', 'MDR', 'ร้อยละ', 'ปี', 'ไม่', 'เส้น'],
    ['เชื้อ', 'สัดส่วนเชื้อ CAUTI', 'CAUTI', 'จำนวน', '', '', 'โดนัท']
  ];
  // รวมแถว: composable ก่อน แล้วตามด้วย SLIDE_SAMPLE (ชนิดเฉพาะ cover/agenda/kpi/...)
  var rowsMain = [];
  COMPOSE.forEach(function (c) { rowsMain.push(['', c[0], c[1], '', '', c[2], c[3], c[4], c[5], c[6], '']); });     // ชนิด/พารามิเตอร์ ว่าง
  SLIDE_SAMPLE.forEach(function (r) { rowsMain.push(['', r[0], r[1], r[2], r[3], '', '', '', '', '', r[4]]); });    // ตัวแปร..กราฟ ว่าง
  var nData = Math.max(rowsMain.length, SLIDE_TYPES.length);
  var grid = [header.concat(['', 'ชนิด (ค่า)', 'ความหมาย'])];
  for (var i = 0; i < nData; i++) {
    var a = rowsMain[i] || ['', '', '', '', '', '', '', '', '', '', ''];
    var t = SLIDE_TYPES[i] || ['', ''];
    a[0] = i < rowsMain.length ? (i + 1) : '';   // เลขลำดับ
    grid.push(a.concat(['', t[0], t[1]]));        // + ช่องว่าง L + อ้างอิง M,N
  }
  sh.getRange(1, 1, grid.length, 14).setValues(grid);   // เขียนทีเดียว

  // จัดรูปแบบหัวคอลัมน์
  sh.getRange(1, 1, 1, 11).setFontWeight('bold').setBackground('#0f766e').setFontColor('#ffffff').setHorizontalAlignment('center');
  sh.getRange(1, 13, 1, 2).setFontWeight('bold').setBackground('#e2e8f0');
  sh.setFrozenRows(1);

  // ดรอปดาวน์ — เฉพาะช่วงที่ใช้งาน (แถวตัวอย่าง + เผื่ออีก 80 แถว)
  var vEnd = rowsMain.length + 80;
  function dv(col, list, allowInvalid, help) {
    sh.getRange(2, col, vEnd, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(list, true).setAllowInvalid(allowInvalid).setHelpText(help).build());
  }
  dv(4, SLIDE_TYPES.map(function (t) { return t[0]; }), true, 'ชนิดเฉพาะ (cover/agenda/kpi/...) — เว้นว่างถ้าใช้แบบประกอบเอง');
  dv(5, SLIDE_PARAMS, true, 'พารามิเตอร์ของชนิดเฉพาะ (เช่น hist-trend ใส่ชื่อแท็บ)');
  dv(6, SLIDE_VARS, true, 'ตัวแปรที่ต้องการ (HAI/VAP/CAUTI/...)');
  dv(7, SLIDE_METRICS, true, 'ค่าที่แสดง: อัตรา/จำนวน/ร้อยละ');
  dv(8, SLIDE_PERIODS, true, 'ช่วงเวลา: ปี/ไตรมาส/เดือน');
  dv(9, SLIDE_HISTORY, true, 'ใช้ข้อมูลเก่า (ย้อนหลัง) หรือไม่');
  dv(10, SLIDE_SHAPES, true, 'รูปแบบกราฟ');

  // ความกว้างคอลัมน์
  var widths = [50, 120, 260, 130, 110, 110, 80, 90, 90, 110, 200, 18, 150, 420];
  for (var c = 0; c < widths.length; c++) sh.setColumnWidth(c + 1, widths[c]);
  var sample = rowsMain;   // สำหรับ log ท้ายฟังก์ชัน

  try { ss.toast('สร้างแท็บ "' + SLIDES_TAB + '" พร้อมดรอปดาวน์เรียบร้อย', 'สำเร็จ', 5); } catch (e) { }
  Logger.log('สร้างแท็บ "%s" เรียบร้อย (%s แถว)', SLIDES_TAB, sample.length);
}

/**
 * ฉบับเบาพิเศษ: เขียนเฉพาะหัวคอลัมน์ + แถวตัวอย่าง (ไม่ใส่ดรอปดาวน์/ตารางอ้างอิง)
 * ใช้เมื่อ createSlidesTab ช้าจนหมดเวลา (ชีตมีสูตรหนัก/IMPORTRANGE)
 * พิมพ์ "ชนิด" เองได้ตามรายการในคำอธิบาย — เว็บอ่านได้เหมือนกันทุกประการ
 */
function createSlidesTabLite() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) ss = SpreadsheetApp.openById(SLIDES_SHEET_ID);
  var sh = ss.getSheetByName(SLIDES_TAB);
  if (sh) { sh.clear(); sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearDataValidations(); } else { sh = ss.insertSheet(SLIDES_TAB); }
  var grid = [['ลำดับ', 'หัวข้อหลัก', 'หัวข้อรอง', 'ชนิด', 'พารามิเตอร์', 'ตัวแปร', 'ค่า', 'ช่วงเวลา', 'ข้อมูลเก่า', 'กราฟ', 'หมายเหตุ']];
  SLIDE_SAMPLE.forEach(function (r, i) { grid.push([i + 1, r[0], r[1], r[2], r[3], '', '', '', '', '', r[4]]); });
  sh.getRange(1, 1, grid.length, 11).setValues(grid);   // เขียนครั้งเดียว ไม่มี validation
  sh.getRange(1, 1, 1, 11).setFontWeight('bold').setBackground('#0f766e').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  Logger.log('createSlidesTabLite: เขียน %s แถว', SLIDE_SAMPLE.length);
}

/**
 * ★ ปลอดภัย: ไม่ลบข้อมูลที่กรอกไว้ ★
 * ใส่ "ดรอปดาวน์" ให้คอลัมน์ที่มีอยู่แล้วในแท็บ Slides โดยจับจากชื่อหัวคอลัมน์
 * (ไม่ clear, ไม่เขียนทับค่า) — ใช้เมื่อกรอกข้อมูลเองไว้แล้ว แต่อยากได้ดรอปดาวน์
 */
function applySlidesValidation() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) ss = SpreadsheetApp.openById(SLIDES_SHEET_ID);
  var sh = ss.getSheetByName(SLIDES_TAB);
  if (!sh) { Logger.log('ไม่พบแท็บ "%s"', SLIDES_TAB); return; }
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var nRows = Math.max(sh.getLastRow() - 1, 0) + 200;   // แถวที่มี + เผื่ออีก 200
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (x) { return String(x).trim(); });
  function colOf(names) { for (var i = 0; i < headers.length; i++) if (names.indexOf(headers[i]) >= 0) return i + 1; return 0; }
  function setDV(c, list, help) {
    if (!c) return;
    sh.getRange(2, c, nRows, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(list, true).setAllowInvalid(true).setHelpText(help).build());
  }
  setDV(colOf(['ชนิด', 'type', 'ชนิดข้อมูล']), SLIDE_TYPES.map(function (t) { return t[0]; }), 'ชนิดเฉพาะ (cover/agenda/kpi/...)');
  setDV(colOf(['พารามิเตอร์', 'param']), SLIDE_PARAMS, 'พารามิเตอร์/ชื่อแท็บย้อนหลัง');
  setDV(colOf(['ตัวแปร', 'ตัวชี้วัด', 'variable']), SLIDE_VARS, 'ตัวแปร (HAI/VAP/หอ/เชื้อ/กลุ่มงาน/...)');
  setDV(colOf(['ค่า', 'ค่าที่แสดง', 'metric']), SLIDE_METRICS, 'ค่า: อัตรา/จำนวน/ร้อยละ');
  setDV(colOf(['ช่วงเวลา', 'ช่วง', 'period']), SLIDE_PERIODS, 'ช่วงเวลา: ปี/ไตรมาส/เดือน');
  setDV(colOf(['ข้อมูลเก่า', 'ย้อนหลัง', 'history']), SLIDE_HISTORY, 'ใช้ข้อมูลเก่า: ใช้/ไม่');
  setDV(colOf(['กราฟ', 'รูปแบบ', 'shape']), SLIDE_SHAPES, 'รูปแบบกราฟ');
  Logger.log('ใส่ดรอปดาวน์ให้คอลัมน์ที่มีอยู่แล้ว (ไม่ลบข้อมูล)');
}

// ⚠️ createSlidesTab / createSlidesTabLite จะ "ลบข้อมูลเดิมทั้งแท็บ" แล้วเขียนตัวอย่างใหม่
//    ถ้ากรอกข้อมูลเองไว้แล้ว ให้ใช้ applySlidesValidation() แทน (ไม่ลบข้อมูล)
//
// หมายเหตุ: โปรเจกต์นี้มี onOpen อยู่แล้วใน SurveillanceMatch.gs — ห้ามประกาศ onOpen ซ้ำ
