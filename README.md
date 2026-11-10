# ระบบติดตามความก้าวหน้างานวิจัยแพทย์ประจำบ้าน (Resident Research Progression Tracker)

> **Ophthalmology Residency Research Progression Tracker** — ระบบบริหารจัดการ ติดตามไทม์ไลน์ และแจ้งเตือนความก้าวหน้างานวิจัยแพทย์ประจำบ้าน ภาควิชาจักษุวิทยา แบบ Real-Time ครอบคลุมหลักสูตร 3 ชั้นปี (R1–R3) พร้อมระบบเชื่อมต่อ Google Sheets และ Google Apps Script 100% Serverless

[![Live Web App](https://img.shields.io/badge/Live_Web_App-Google_Apps_Script-0F9D58?style=for-the-badge&logo=google)](https://script.google.com/macros/s/AKfycbyLb-1J1Btou1977X2384IN1bELDa_5TV_6szhWMargyoWa_KvCW2TtKA6XmbI1pQ/exec)
[![GitHub Pages](https://img.shields.io/badge/Demo-GitHub_Pages-222222?style=for-the-badge&logo=github)](https://sunshinesunshine.github.io/resident-research-tracker/)
[![Version](https://img.shields.io/badge/version-1.2.0-blue?style=for-the-badge)](VERSION)

---

## 🔗 ลิงก์ระบบใช้งานจริง (Live Production Links)

| บริการ | ลิงก์เข้าใช้งาน | คำอธิบาย |
| :--- | :--- | :--- |
| **🌐 Web App (ระบบจริงภาควิชา)** | **[คลิกเปิดใช้งาน Web App](https://script.google.com/macros/s/AKfycbyLb-1J1Btou1977X2384IN1bELDa_5TV_6szhWMargyoWa_KvCW2TtKA6XmbI1pQ/exec)** | รันบน Google Apps Script (Serverless) เชื่อมต่อฐานข้อมูลกลางภาควิชา 2-Way Sync |
| **💻 Interactive Demo (GitHub Pages)** | **[คลิกเปิด GitHub Pages](https://sunshinesunshine.github.io/resident-research-tracker/)** | หน้าเว็บตัวอย่างแบบ Standalone ทำงานบนเบราว์เซอร์พร้อม LocalStorage |
| **📊 Google Sheets (ฐานข้อมูลกลาง)** | **[คลิกเปิดดู Google Sheet](https://docs.google.com/spreadsheets/d/18f71XkkgACgGwhWGii2wIA0BnAlsnox0_sy61VEkFXk/edit)** | ตาราง Sheet ข้อมูลดิบสำหรับประธานหลักสูตรและเลขาภาควิชา |
| **⚙️ Google Apps Script Console** | **[คลิกเปิด Script Editor](https://script.google.com/d/1v7zKapUPDvNooT1_qeq4YN0J3VwnTU7HciqurM8hkwbn1U0ZmB3JinNr/edit)** | คอนโซลจัดการโค้ดหลังบ้าน (Backend Serverless) |

---

## ✨ จุดเด่นของระบบ (Core Features)

### 1. 📊 Executive Dashboard & Kinetic KPI
- **High-Impact Metric Cards**: ตัวเลขนับอนิเมชันด้วย GSAP, แสดงจำนวน Resident รวม, สัดส่วนผลงานที่ตามแผน, และจำนวนผู้ที่ติดสถานะเฝ้าระวัง
- **Distribution Visualizer**: กราฟแท่งจำแนกสถานะและสัดส่วนความคืบหน้ารายชั้นปี (R1, R2, R3)
- **Automatic Forum Countdown**: นาฬิกานับถอยหลังสู่งานสัมมนาครั้งถัดไปแบบอัตโนมัติ

### 2. 📅 Dual-View Progression (Gantt & Kanban)
- **Interactive Multi-Cohort Gantt**: แผนภาพแกนต์ชาร์ตแสดงระยะเวลาของ 6 ขั้นตอนตามหลักสูตร พร้อมระบุหมุด Milestone ชัดเจน
- **Smart Kanban Board**: บอร์ดติดตามงานแบบ Drag & Drop จำแนกตาม 6 สเต็ปของงานวิจัย พร้อมตัวกรองตามชั้นปี, สถานะ, และอาจารย์ที่ปรึกษา

### 3. 🚨 Early Warning Radar (ระบบเฝ้าระวังอัตโนมัติ)
- คำนวณวันคงเหลือเทียบกับกำหนดส่งจริง (Strict Deadline Tracking)
- จำแนกสถานะเป็น **"ตามแผน (On Track)"**, **"ใกล้ถึงกำหนด (Due Soon ≤ 30 วัน)"**, และ **"เลยกำหนด (Overdue)"** โดยอัตโนมัติ

### 4. 👨‍⚕️ 7 Subspecialties & 33 Faculty Advisors
รองรับอาจารย์ที่ปรึกษาครบทั้ง 33 ท่านใน 7 อนุสาขาของภาควิชาจักษุวิทยา:
- **Cornea (5 ท่าน)**: พญ.นวลจิรา, พญ.กนกทิพย์, พญ.เบญจมา, นพ.วีรภัทร, พญ.ชิดชนก
- **Glaucoma (5 ท่าน)**: พญ.ดวงดาว, พญ.กุลวรรณ, พญ.อรอร, พญ.สุพัตรา, นพ.ปิติพงศ์
- **Neuro-Ophthalmology (3 ท่าน)**: พญ.ปิยวดี, นพ.ปรเมศวร์, นพ.จิเมธัส
- **OPRC (4 ท่าน)**: พญ.สุณิสา, พญ.มิ่งขวัญ, พญ.อรวีณัฏฐ์, นพ.สรรเสริญ
- **Pediatric Ophthalmology (2 ท่าน)**: พญ.อาภา, พญ.จิราภา
- **Retina & Vitreous (11 ท่าน)**: พญ.สุดาวดี, พญ.วงศ์ศิริ, นพ.วรภัทร, พญ.อุมาภรณ์, พญ.พิริญา, นพ.เอกชัย, พญ.กาญจน์ธิดา, นพ.วรินทร์, พญ.นวลพร, พญ.ทิฆัมพร, พญ.ปาณิศา
- **Uveitis & Immunology (3 ท่าน)**: นพ.โยธิน, นพ.นพคุณ, พญ.ณิศวรรณ

### 5. 📱 Resident Self-Service Portal
- แพทย์ประจำบ้านสามารถเลือกชื่อตนเองเพื่ออัปเดตเปอร์เซ็นต์ความก้าวหน้า
- แนบลิงก์ Google Drive ของเอกสารโครงร่าง (Proposal) หรือแบบตอบรับจริยธรรม (COA) ได้ด้วยตนเอง

### 6. 🔐 Admin Security Gate
- ระบบความปลอดภัยด้วย **Admin PIN** (รหัสเริ่มต้น: `1234`)
- ป้องกันการแก้ไขกรอบเวลาหลักสูตร, การเลื่อนปีการศึกษา, หรือการรีเซ็ตระบบโดยไม่ได้รับอนุญาต

---

## 🛠️ โครงสร้างไฟล์ใน Repository (Project Structure)

```text
resident-research-tracker/
├── index.html                           # Single-Page Application (HTML5 + Tailwind CSS + GSAP)
├── Code.gs                              # Google Apps Script Backend (Serverless CRUD & Sheet Sync)
├── appsscript.json                      # Apps Script Manifest (Timezone Asia/Bangkok, Permissions)
├── Resident_Research_Progress_Tracker.xlsx # เทมเพลตฐานข้อมูล Excel 5 แผ่นงาน
├── generate_sheet.py                    # สคริปต์ Python สำหรับสร้าง/รีเซ็ตชีต Excel
├── VERSION                              # หมายเลขเวอร์ชันปัจจุบัน (v1.2.0)
├── CHANGELOG.md                         # บันทึกประวัติการพัฒนา
└── README.md                            # คู่มือการใช้งานระบบ
```

---

## 🚀 สถาปัตยกรรมเทคโนโลยี (Tech Stack)

- **Frontend**: HTML5, Tailwind CSS (Utility-first), Vanilla JavaScript (Zero framework bloat), Plus Jakarta Sans / Prompt / Sarabun typography
- **Kinetic Animation**: GSAP (GreenSock Animation Platform) v3.12.5
- **Backend**: Google Apps Script (Serverless V8 Runtime)
- **Database**: Google Sheets API (2-Way Synchronized, 5 Normalized Sheets)
- **Deployment**: Google Apps Script Web App (Exec endpoint) & GitHub Pages

---

## 📄 License & Attribution

พัฒนาเพื่อสนับสนุนงานฝึกอบรมแพทย์ประจำบ้าน ภาควิชาจักษุวิทยา  
Designed and orchestrated with high precision by **Dr. Chayanat Ouijalern** & **Nid (@nid)**.
