# OR Appointment — GitHub + Google Apps Script

ชุดนี้เป็นเวอร์ชันพร้อมนำไปใช้งานสำหรับสร้างเว็บ OR Appointment บน GitHub Pages โดยยังใช้ Google Sheet เดิมผ่าน Google Apps Script เดิม ทำให้เว็บเก่าและเว็บใหม่ใช้งานข้อมูลชุดเดียวกันได้พร้อมกัน

## ไฟล์ในชุด

- `index.html` — หน้าเว็บ OR Appointment สำหรับ GitHub Pages
- `code.gs` — Google Apps Script backend ที่เพิ่ม API สำหรับ GitHub และยังรองรับ OR Queue Connector แบบเดิม

## โครงสร้างการทำงาน

```text
OR Queue
   ↓
Chrome Extension (Modal Safe)
   ↓ POST แบบเดิม
Google Apps Script เดิม
   ↓
Incoming / PENDING
   ↓
┌──────────────────────────────┐
│ เว็บเดิม Apps Script         │
│ เว็บใหม่ GitHub Pages        │
└──────────────────────────────┘
   ↓
ตรวจสอบ → ยืนยัน
   ↓
Database / Calendar ใน Sheet เดิม
```

## 1. อัปเดต Google Apps Script เดิม

1. เปิดโปรเจกต์ Apps Script เดิม
2. สำรอง `code.gs` เดิมไว้ก่อน
3. นำ `code.gs` ในชุดนี้ไปแทนเนื้อหา `code.gs` เดิม
4. ไปที่ **Deploy → Manage deployments**
5. เลือก Web App deployment เดิม → **Edit**
6. เลือก **New version** แล้วกด Deploy
7. ใช้ URL `/exec` เดิมต่อได้ ไม่จำเป็นต้องเปลี่ยน URL

ฟังก์ชัน `doPost()` แบบเดิมของ OR Queue ยังถูกเก็บไว้ ดังนั้น Chrome Extension เดิมยังส่งข้อมูลเข้า `Incoming` ได้เหมือนเดิม

## 2. ตั้งค่า URL ของ API ใน GitHub

เปิด `index.html` แล้วหา:

```js
const API_URL = 'PASTE_YOUR_EXISTING_APPS_SCRIPT_EXEC_URL_HERE';
```

เปลี่ยนเป็น URL `/exec` ของ Apps Script เดิม เช่น:

```js
const API_URL = 'https://script.google.com/macros/s/XXXXXXXXXXXX/exec';
```

**ไม่ต้องเอา URL ไปเปิดใน browser เพื่อทดสอบ JSON** เพราะ `doGet()` ของ Apps Script ยังคงแสดงเว็บ OR Appointment เดิม การทดสอบ API ให้ใช้ปุ่ม **🔌 ทดสอบ API** ในหน้า Login ของเว็บ GitHub

## 3. นำเว็บขึ้น GitHub Pages

สร้าง repository แล้ววางไฟล์:

```text
OR-Appointment/
└── index.html
```

จากนั้นเปิด GitHub Pages ของ repository นั้น

## 4. การทดสอบ

### ทดสอบ API

เปิดเว็บ GitHub → ที่หน้า Login กด **🔌 ทดสอบ API**

ถ้าปกติจะแสดง:

```text
✅ API เชื่อมต่อสำเร็จ
```

### ทดสอบระบบหลัก

1. Login
2. ตรวจ Calendar และรายการผู้ป่วย
3. เพิ่มนัดหมายจากเว็บ GitHub
4. ตรวจ Google Sheet ว่ามีข้อมูล
5. เปิดเว็บ Apps Script เดิมและตรวจว่าเห็นข้อมูลเดียวกัน
6. แก้ไข / ลบ / กู้คืน / ลบถาวร

### ทดสอบ OR Queue

1. เปิด OR Queue และเข้าสู่ระบบ
2. เปิด modal ของผู้ป่วย
3. ใช้ Chrome Extension เดิมกดเก็บข้อมูลจาก modal
4. ส่งรายการเข้า Pending
5. เปิด OR Appointment
6. กด **รอตรวจสอบ OR Queue**
7. ตรวจสอบข้อมูล
8. กด **ยืนยันลง Calendar** เพื่อบันทึกเข้า Database/Calendar
9. หรือกด **ปฏิเสธ** พร้อมระบุเหตุผล

ระบบ GitHub ใช้ API เหล่านี้สำหรับ Pending:

- `getPendingImports`
- `approvePendingImport`
- `rejectPendingImport`

ดังนั้นไม่จำเป็นต้องแก้ Chrome Extension เดิมสำหรับ workflow นี้

## 5. ความปลอดภัย

- การ Login ของเว็บ GitHub ใช้ session token ชั่วคราวใน `sessionStorage`
- token ถูกเก็บใน Apps Script Cache และหมดอายุภายใน 6 ชั่วโมง
- `ping` เปิดไว้โดยไม่ต้อง Login เพื่อใช้ทดสอบการเชื่อมต่อเท่านั้น
- ก่อนใช้งานจริงควรเปลี่ยน username/password ที่ hard-code ใน `code.gs` เป็นค่าที่ปลอดภัยกว่า และควรย้ายไป Script Properties

## 6. สำคัญ

เว็บเก่าและเว็บใหม่ใช้ฐานข้อมูล Google Sheet ชุดเดียวกัน ดังนั้น **ไม่ต้องย้ายข้อมูลเดิม** และไม่ต้องเปลี่ยน URL เว็บเก่าที่เครื่องอื่นบันทึกไว้

เมื่อมีการเพิ่มข้อมูลจากเว็บใดเว็บหนึ่ง อีกเว็บจะเห็นข้อมูลเดียวกันหลังจากโหลดข้อมูลใหม่
