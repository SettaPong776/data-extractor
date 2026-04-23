# Data Extractor — แยกข้อมูลตารางจาก PDF/Word → Excel

โปรแกรมเว็บแอปสำหรับสกัดข้อมูลตารางจากไฟล์ PDF และ Word (.docx) แล้วจัดเรียงคอลัมน์ให้ตรงกับ Excel ที่ต้องการ จากนั้นส่งออกเป็นไฟล์ `.xlsx`

## ✨ คุณสมบัติ

- 📤 **อัปโหลด** ไฟล์ PDF หรือ Word (.docx) — รองรับ Drag & Drop
- 📊 **สกัดตาราง** อัตโนมัติจากทุกหน้า พร้อม auto-merge ตารางข้ามหน้า
- 🔄 **จัดเรียงคอลัมน์** — ลาก Drag & Drop, เปลี่ยนชื่อ, ซ่อน/แสดงคอลัมน์
- 📥 **ส่งออก Excel** (.xlsx) พร้อม auto-size column width
- 🔒 **ปลอดภัย** — ทำงาน Client-side ทั้งหมด ข้อมูลไม่ถูกส่งไปเซิร์ฟเวอร์

## 🚀 วิธีใช้งาน

1. เปิดไฟล์ `index.html` ในเบราว์เซอร์ หรือรันผ่าน Web Server (เช่น XAMPP)
2. ลากไฟล์ PDF/DOCX มาวางในพื้นที่อัปโหลด
3. ตรวจสอบข้อมูลที่สกัดได้ใน Preview
4. จัดเรียงคอลัมน์ให้ตรงกับ Excel ปลายทาง
5. ดาวน์โหลดไฟล์ Excel (.xlsx)

## 🛠 เทคโนโลยี

| Library | หน้าที่ |
|---------|---------|
| [PDF.js](https://mozilla.github.io/pdf.js/) | อ่านและสกัดข้อมูลจาก PDF |
| [mammoth.js](https://github.com/mwilliamson/mammoth.js) | แปลง .docx → HTML แล้วอ่านตาราง |
| [SheetJS](https://sheetjs.com/) | สร้างไฟล์ Excel (.xlsx) |

## 📁 โครงสร้างโปรเจค

```
convert/
├── index.html              # หน้าเว็บหลัก
├── css/
│   └── style.css           # Dark-mode UI + Glassmorphism
├── js/
│   ├── app.js              # Main application controller
│   ├── pdf-extractor.js    # PDF table extraction engine
│   ├── word-extractor.js   # Word/DOCX table extraction
│   ├── column-mapper.js    # Column mapping UI (Drag & Drop)
│   └── excel-exporter.js   # Excel export
└── README.md
```

## 📝 หมายเหตุ

- รองรับ PDF ที่สร้างจากโปรแกรม (Word→PDF, Excel→PDF) ✅
- PDF ที่สแกนมาจากกระดาษ (Scanned) ❌ ต้องใช้ OCR
