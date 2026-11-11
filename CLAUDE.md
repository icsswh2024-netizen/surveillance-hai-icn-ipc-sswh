# คำแนะนำโปรเจกต์ (CLAUDE.md)

ระบบเฝ้าระวังการติดเชื้อในโรงพยาบาล (HAI) — โรงพยาบาลศรีสังวรสุโขทัย
ไฟล์หลักคือ `index.html` (single-file, deploy ขึ้น GitHub Pages ผ่าน branch `main`)

## กฎสำคัญ

### 1) ❌ งดใช้ localStorage (คำสั่งผู้ใช้)
- **ห้ามเขียนโค้ดใหม่ที่ใช้ `localStorage` / `sessionStorage`** เพื่อเก็บสถานะ ข้อมูล หรือการตั้งค่า
- เหตุผล: ผูกกับเครื่อง/เบราว์เซอร์เดียว → เปิดอีกอุปกรณ์แล้วไม่เห็นข้อมูลเดิม ไม่เหมาะกับการใช้งานหลายเครื่อง/หลายคน
- แทนที่ด้วย: **แหล่งข้อมูลกลางที่ทุกเครื่องเห็นเหมือนกัน** เช่น
  - อ่านจาก Google Sheet โดยตรง (gviz CSV) — เป็นแหล่งความจริงกลาง
  - บันทึก/อัปเดตผ่าน Apps Script Web App (REC_URL)
  - ค่าคอนฟิกฝังในโค้ด (const) หรือดึงจากชีต
- หมายเหตุ: โค้ดเดิมบางส่วนยังใช้ localStorage อยู่ (แบบร่างฟอร์ม, URL, session login, การตั้งค่าธีม/sidebar) — ค่อยทยอยย้ายออกเมื่อผู้ใช้สั่ง ไม่รื้อโดยพลการ

## ขั้นตอนมาตรฐานเมื่อแก้ไข
1. แก้ `index.html` แล้ว **บั๊มป์ build stamp** (`#buildStamp` เช่น `v2026-11-11b`)
2. commit + push ขึ้น `main` ด้วย `git push origin HEAD:main`
3. ลงท้าย commit ด้วย footer ที่กำหนด (Co-Authored-By + Claude-Session)

## สถาปัตยกรรม Apps Script
- `RecordCode.gs` = Web App (REC_URL) — login + บันทึกเคส (addCase/editCase) + verify + หัตถการ
- `SurveillanceMatch.gs` / `Slides.gs` = เมนู IC Tools (ไม่มี web endpoint)
- โปรเจกต์หนึ่งมี doGet/doPost ได้ **เพียงชุดเดียว**

## แท็บชีต (ปีงบปัจจุบัน 2570)
- `ทะเบียน-1-2570` = เคส · `ตัวหาร-2570` = ตัวหาร · + ปีเก่า `ทะเบียน-1-2569` / `ตัวหาร-2569`
- config อยู่ที่ `SHEET_NAME` + `EXTRA_SHEETS` (index.html) และ `REC_SHEET` (RecordCode.gs)
