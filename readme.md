\# 🏆 NextTalent — ระบบสะสม XP และทักษะบุคลากร



ระบบสะสมคะแนน XP และ Level แบบ Gamification สำหรับบุคลากร โดยให้ผู้ใช้อัพโหลดใบรับรอง (Certificate) เป็น PDF ระบบจะทำ \*\*OCR + AI วิเคราะห์ทักษะ\*\* อัตโนมัติ คำนวณ XP ชั่วคราว แล้วส่งเข้าสู่กระบวนการอนุมัติจาก Admin ก่อนนับเข้า Level จริง



พัฒนาด้วย \*\*Google Apps Script\*\* ทำงานเป็น \*\*Single-Page Web App\*\* ไฟล์เดียว (`App.html`) ไม่มี backend แยก ไม่ต้องมี server ภายนอก ใช้ Google Sheets เป็นฐานข้อมูล



\---



\## 📋 สารบัญ



\- \[ภาพรวมระบบ](#-ภาพรวมระบบ)

\- \[สถาปัตยกรรม](#-สถาปัตยกรรม)

\- \[โครงสร้างไฟล์](#-โครงสร้างไฟล์)

\- \[โครงสร้าง Google Sheets](#-โครงสร้าง-google-sheets)

\- \[สูตรคำนวณ XP / Level](#-สูตรคำนวณ-xp--level)

\- \[Badge System](#-badge-system)

\- \[ฟังก์ชันฝั่ง Server](#-ฟังก์ชันฝั่ง-server-google-scriptrun)

\- \[การติดตั้ง (Setup)](#-การติดตั้ง-setup)

\- \[Flow การทำงาน](#-flow-การทำงาน)

\- \[หน้าตา UI](#-หน้าตา-ui)

\- \[Error Handling](#-error-handling)

\- \[ข้อจำกัดที่ทราบอยู่แล้ว](#-ข้อจำกัดที่ทราบอยู่แล้ว)

\- \[แผนต่อยอด](#-แผนต่อยอด)



\---



\## 🎯 ภาพรวมระบบ



| ฟีเจอร์ | รายละเอียด |

|---|---|

| \*\*อัพโหลดใบรับรอง\*\* | รับไฟล์ PDF → แปลง Base64 → ส่งไป OCR + AI วิเคราะห์ |

| \*\*OCR\*\* | เรียก n8n webhook แบบ multipart/form-data |

| \*\*AI วิเคราะห์ทักษะ\*\* | เรียก KKU AI Gateway วิเคราะห์ `skill` (IT/บริหาร/ทั่วไป) และ `level` (มหาวิทยาลัย/ประเทศ/นานาชาติ) |

| \*\*คำนวณ XP\*\* | คำนวณชั่วคราวจาก Config (baseXP × rarity) ก่อนบันทึกจริง |

| \*\*อนุมัติ/ปฏิเสธ\*\* | Admin ตรวจสอบ แก้ไข skill/level และอนุมัติ/ปฏิเสธ ก่อนนับ XP เข้า Level |

| \*\*Level แบบคู่ขนาน\*\* | `growupLevel` (รวมทุกทักษะ) และ `specialList` (แยกตามทักษะ) คำนวณคนละสูตร |

| \*\*Badge / Gamification\*\* | Badge 4 ระดับ (Rookie/Rising Star/Expert/Master) ตาม Level พร้อม popup animation ตอนได้ XP |

| \*\*Dashboard\*\* | Leaderboard, สัดส่วนทักษะ, Skill Gap Analysis เทียบเป้าหมายองค์กร |

| \*\*ค้นหาบุคลากร\*\* | ค้นหาตามชื่อ/ทักษะ เพื่อหาคนที่มี skill ตรงกับที่ต้องการ |



\---



\## 🏗 สถาปัตยกรรม



```

┌─────────────────────────────────────────────────┐

│                Google Apps Script                │

│                                                   │

│  doGet() ──► App.html (Single Page, ทุก view)    │

│                                                   │

│  Client (JS)  ◄──google.script.run──►  Server(.gs) │

│  ไม่มี doPost(), ไม่ผ่าน fetch()/HTTP              │

└───────────────┬───────────────────┬──────────────┘

&#x20;               │                   │

&#x20;               ▼                   ▼

&#x20;       ┌───────────────┐   ┌──────────────────┐

&#x20;       │ Google Sheets │   │  Google Drive      │

&#x20;       │ Users/        │   │  (เก็บไฟล์ PDF     │

&#x20;       │ Submissions/  │   │   ใบรับรอง)        │

&#x20;       │ Config        │   └──────────────────┘

&#x20;       └───────────────┘

&#x20;               │

&#x20;               ▼

&#x20;    ┌────────────────────┐      ┌─────────────────────┐

&#x20;    │  n8n Webhook (OCR) │─────►│ KKU AI Gateway       │

&#x20;    │  multipart/form-   │      │ วิเคราะห์ skill/level │

&#x20;    │  data → JSON       │      │ (chat/completions)   │

&#x20;    └────────────────────┘      └─────────────────────┘

```



\*\*หลักการสำคัญ:\*\*

\- \*\*Single Web App / Single URL\*\* — ทุก view (หน้าหลัก, Admin, Dashboard, ค้นหา) อยู่ในไฟล์ `App.html` ไฟล์เดียว สลับด้วย JS (`switchView()`) ไม่ใช่คนละหน้า/คนละ URL

\- \*\*ไม่มี `doPost()`\*\* — เดิมวางแผนให้รับ REST request แต่ Apps Script Web App ตอบ POST ด้วย 302 redirect ทำให้ browser แปลงเป็น GET (ตาม HTTP spec) แล้ววิ่งเข้า `doGet()` แทน จึงเปลี่ยนมาใช้ `google.script.run` ทั้งหมด

\- \*\*ไม่มี Login\*\* — ใช้ dropdown เลือก user แทน

\- \*\*API Key ปลอดภัย\*\* — เก็บใน `PropertiesService` (Script Properties) ไม่ hardcode ในโค้ด



\---



\## 📁 โครงสร้างไฟล์



```

NextTalent (Apps Script Project)

├── NextTalent.gs      # ฟังก์ชันฝั่ง server ทั้งหมด (doGet, business logic, Sheets I/O)

├── App.html           # Single-page app ทั้งหมด (HTML + CSS + Client-side JS)

└── (Google Sheets)    # ผูกกับ Spreadsheet: Users / Submissions / Config

```



\---



\## 🗂 โครงสร้าง Google Sheets



\### Sheet: `Users`



| คอลัมน์ | คำอธิบาย |

|---|---|

| `userId` | รหัส user (unique) |

| `name` | ชื่อ-นามสกุล |

| `department` | หน่วยงาน |

| `photoUrl` | ลิงก์รูปโปรไฟล์ (ไม่บังคับ) |



Seed ตัวอย่าง: `u001` มานะ ใจดี, `u002` สมหญิง รักเรียน, `u003` วิชัย ตั้งใจทำงาน



\### Sheet: `Submissions`



| คอลัมน์ | คำอธิบาย |

|---|---|

| `id` | รหัสรายการ (unique) |

| `timestamp` | เวลาที่ส่งผล |

| `userId` | อ้างอิงจาก Sheet Users |

| `fileUrl` | ลิงก์ไฟล์ใน Drive |

| `ocrText` | ข้อความที่ได้จาก OCR |

| `skill` | IT / บริหาร / ทั่วไป |

| `level` | มหาวิทยาลัย / ประเทศ / นานาชาติ |

| `rarity` | ตัวคูณ rarity ที่ใช้ตอนคำนวณ (เก็บย้อนหลัง) |

| `xp` | XP ที่คำนวณได้ |

| `status` | `submitted` (รออนุมัติ) / `approved` (นับ XP) / `rejected` (ไม่นับ XP แต่เก็บประวัติ) |



\### Sheet: `Config` (key-value ตาม category)



| category | key | value |

|---|---|---|

| `level` | มหาวิทยาลัย | 10 |

| `level` | ประเทศ | 30 |

| `level` | นานาชาติ | 50 |

| `rarity` | IT | 1.0 |

| `rarity` | บริหาร | 1.5 |

| `rarity` | ทั่วไป | 1.0 |

| `curve` | base | 100 |

| `curve` | exponent | 1.5 |

| `specialCurve` | base | 50 |

| `specialCurve` | exponent | 1.5 |

| `benchmark` | `<skill>` | เป้าหมาย XP เฉลี่ยต่อคน (ใช้ใน Skill Gap Analysis, ไม่บังคับ) |



> ปรับค่าทั้งหมดได้จากชีตโดยตรง ไม่ต้องแก้โค้ด



\---



\## 🧮 สูตรคำนวณ XP / Level



ระบบคำนวณ Level แบบ\*\*คู่ขนาน 2 รูปแบบ\*\*:



\### 1. `growupLevel` — ภาพรวมทุกทักษะ



```

XP ต่อใบรับรอง = baseXP(level) × rarity(skill)

growupLevel.XP = Σ xp ของทุกรายการ status=approved (ไม่แยก skill)

XP สะสมที่ต้องใช้ขึ้น Level N = curve.base × N^curve.exponent

```



ค่าเริ่มต้น: `base=100, exponent=1.5`



\### 2. `specialList` — แยกตาม skill (IT / บริหาร / ทั่วไป)



```

specialList\[S].XP = Σ xp ของรายการ approved ที่ skill=S เท่านั้น

XP สะสมที่ต้องใช้ขึ้น Level N = specialCurve.base × N^specialCurve.exponent

```



ค่าเริ่มต้น: `base=50, exponent=1.5` (คนละ base กับ growupLevel เพราะ XP ต่อ skill สะสมได้น้อยกว่า XP รวม)



> \*\*`pendingXp`\*\* (รายการ `submitted` ที่รออนุมัติ) จะคำนวณแยกไว้ต่างหากทั้งใน growupLevel และในแต่ละ skill — ไม่กระทบ Level จนกว่าจะ approved



\---



\## 🎖 Badge System



ใช้เกณฑ์ Level เดียวกันทั้ง 2 ระบบ แต่แสดงผลต่างกัน:



| Level | Badge | สี (unlocked) |

|---|---|---|

| 1–4 | Rookie 🌱 | `#65a30d` เขียว |

| 5–9 | Rising Star ⭐ | `#f59e0b` อำพัน |

| 10–19 | Expert 🔥 | `#ea580c` ส้ม |

| 20+ | Master 👑 | `#eab308` ทอง |



\- \*\*growupLevel badge\*\* — แสดงไอคอนเดียว (badge ปัจจุบันที่ได้รับ)

\- \*\*specialList badge\*\* — แสดงเป็น shelf ครบ 4 ระดับต่อ skill (locked = สีเทาอ่อน / unlocked = สีของตัวเอง)



\---



\## ⚙️ ฟังก์ชันฝั่ง Server (`google.script.run`)



Client เรียกฝั่ง server ตรงๆ ไม่ผ่าน `fetch()`/`doPost()`



| ฟังก์ชัน | หน้าที่ |

|---|---|

| `getUsersList()` | คืนรายชื่อ user ทั้งหมด สำหรับ dropdown |

| `getUserState(userId)` | คืน profile, growupLevel, specialList, submissions ทั้งหมดของ user |

| `analyzeCertificate(fileBase64, fileName, mimeType)` | เซฟไฟล์ลง Drive → OCR → AI วิเคราะห์ → คำนวณ XP ชั่วคราว (ยังไม่เซฟ) |

| `submitResult(userId, result)` | บันทึกผลลง Submissions (status=`submitted`) แล้วคืน state ล่าสุด |

| `getPendingSubmissions()` | คืนรายการที่รออนุมัติทั้งหมด (สำหรับหน้า Admin) |

| `getConfigOptions()` | คืนตัวเลือก skill/level ทั้งหมดจาก Config (สำหรับ dropdown แก้ไข) |

| `updateSubmission(id, skill, level)` | แก้ไข skill/level ของรายการที่รออนุมัติ |

| `approveSubmission(id)` | เปลี่ยน status เป็น `approved` (นับ XP เข้า Level) |

| `rejectSubmission(id)` | เปลี่ยน status เป็น `rejected` (เก็บประวัติ ไม่นับ XP) |

| `addUser(name, department, photoUrl)` | เพิ่ม user ใหม่เข้า Sheet Users |

| `getDashboardStats()` | คืนสถิติรวม: totalUsers, leaderboard, สัดส่วน skill ที่ approved |

| `getSkillGapAnalysis()` | เทียบ XP เฉลี่ยต่อ skill กับ benchmark ใน Config หา critical gaps ต่อหน่วยงาน |

| `searchUsers(query)` | ค้นหา user ตามชื่อ/ทักษะ |



> Error ที่ `throw` ในฝั่ง server จะเด้งเข้า `.withFailureHandler(...)` ของฝั่ง client โดยอัตโนมัติ



\---



\## 🚀 การติดตั้ง (Setup)



\### 1. สร้างโปรเจกต์



สร้าง Google Apps Script ผูกกับ Google Sheets ไฟล์ใหม่ (Container-bound) หรือ Standalone + ระบุ Spreadsheet ID เอง



\### 2. ตั้งค่า Script Properties



ไปที่ \*\*Project Settings → Script Properties\*\* เพิ่ม:



| Key | Value |

|---|---|

| `AI\_GATEWAY\_API\_KEY` | API key ของ KKU AI Gateway |

| `AI\_GATEWAY\_BASE\_URL` | `https://gen.ai.kku.ac.th/api/v1` |

| `OCR\_WEBHOOK\_URL` | `https://automate.kku.ac.th/webhook-test/lab4-pdf` |



\### 3. สร้าง Sheet และ seed ข้อมูล



รันฟังก์ชันเหล่านี้จาก Apps Script Editor \*\*ครั้งเดียว\*\* ตามลำดับ:



```javascript

setupSheet();      // สร้าง Sheet Users / Submissions / Config พร้อม header

setupSeedData();   // ใส่ค่า default ใน Config + user ตัวอย่างใน Users

```



> ทั้งสองฟังก์ชันจะ\*\*ไม่เขียนทับ\*\*ถ้า Sheet นั้นมีข้อมูลอยู่แล้ว (ปลอดภัยต่อการรันซ้ำ)



\### 4. วางไฟล์ `App.html`



สร้างไฟล์ HTML ชื่อ `App` ใน Apps Script Editor แล้ววางโค้ด UI ทั้งหมด (มีทั้ง view หน้าหลัก, Dashboard, Admin, ค้นหา ในไฟล์เดียว)



\### 5. Deploy เป็น Web App



\*\*Deploy → New deployment → Web app\*\*



| ตั้งค่า | ค่า |

|---|---|

| Execute as | Me (เจ้าของ script) |

| Who has access | เฉพาะคนในองค์กร (หรือตามนโยบายที่ต้องการ) |



จะได้ URL เดียวสำหรับใช้งานทั้งระบบ (หน้าหลัก/Admin/Dashboard/ค้นหา อยู่ในหน้าเดียวกัน สลับด้วยปุ่มในแถบเมนู)



\---



\## 🔄 Flow การทำงาน



\### สำหรับผู้ใช้ทั่วไป



```

1\. เลือกชื่อจาก dropdown "เลือก User"

2\. ระบบโหลด Level/XP/Badge/รายการที่เคยส่งของ user คนนั้น

3\. กด "อัพโหลดใบรับรอง" → เลือกไฟล์ PDF

4\. รอผล OCR + AI วิเคราะห์ (synchronous, มี loading state)

5\. เห็นผลชั่วคราว: skill, level, rarity, XP ที่จะได้

6\. กด "ส่งผล" → บันทึกลง Sheet (status = submitted)

&#x20;  → popup +XP animation

&#x20;  → รายการไปอยู่ในสถานะ "รออนุมัติ" (ไม่กระทบ Level จนกว่า Admin จะอนุมัติ)

```



\### สำหรับ Admin



```

1\. เข้าแท็บ "Admin" → ดูรายการรออนุมัติทั้งหมด (ทุก user)

2\. กด "ดู/แก้ไข" เพื่อ preview ไฟล์ PDF + แก้ skill/level ถ้า AI วิเคราะห์ผิด

3\. กด "อนุมัติ" → status = approved → XP นับเข้า Level จริง

&#x20;  หรือกด "ปฏิเสธ" → status = rejected → เก็บประวัติ ไม่นับ XP

4\. ถ้า user ที่เปิดหน้าหลักอยู่พอดีตรงกับ submission ที่ถูกดำเนินการ

&#x20;  → หน้าหลักของ user นั้นจะรีเฟรชอัตโนมัติ + toast แจ้งเตือน

```



\---



\## 🖥 หน้าตา UI



ระบบเป็น Single-Page App มี 4 view สลับกันด้วยแถบเมนู (sidebar):



| View | เนื้อหา |

|---|---|

| \*\*หน้าหลัก\*\* | เลือก user, การ์ด Level (growupLevel + medal), specialList (XP/badge shelf แยกตาม skill), อัพโหลดใบรับรอง, คลังใบรับรอง (inventory) |

| \*\*แดชบอร์ด\*\* | สถิติรวม, Leaderboard Top 5, สัดส่วนทักษะที่ approved, Skill Gap Analysis เทียบ benchmark, Critical Gaps ต่อหน่วยงาน |

| \*\*Admin\*\* | รายการรออนุมัติ (ดู/แก้ไข/อนุมัติ/ปฏิเสธ), เพิ่ม User ใหม่ |

| \*\*ค้นหา\*\* | ค้นหาบุคลากรตามชื่อ/ทักษะ พร้อม badge และ skill tags |



\*\*ธีมสี:\*\* ม่วง-ชมพู-ทอง (gamification, dark mode) — responsive รองรับทั้ง Desktop และ Mobile



\---



\## ⚠️ Error Handling



| จุดที่อาจเกิด error | วิธีจัดการ |

|---|---|

| OCR webhook ตอบไม่ใช่ JSON | fallback แสดง raw response text เพื่อ debug |

| OCR ไม่พบข้อความ (PDF สแกน/รูปภาพ) | throw error แจ้งชัดเจน |

| AI Gateway ตอบไม่ใช่ JSON ที่ parse ได้ | throw error พร้อม raw content |

| AI วิเคราะห์ skill/level ไม่ตรงกับ Config | throw error "ไม่พบค่า rarity/baseXP" |

| Error ทั้งหมดฝั่ง server | เด้งเข้า `.withFailureHandler()` ฝั่ง client → แสดงกล่อง error สีแดง + ปุ่ม "ลองใหม่" (ไม่ auto-retry, ไม่บันทึกแถว error ลง Sheet) |



\---



\## 🚧 ข้อจำกัดที่ทราบอยู่แล้ว



\- ไม่มีระบบ Login จริง (ใช้ dropdown เลือก user แทน) — เหมาะสำหรับ prototype/internal use เท่านั้น

\- OCR รองรับเฉพาะไฟล์ PDF ที่มีข้อความ (ไม่รองรับ PDF ที่เป็นภาพสแกนล้วนถ้า OCR webhook ไม่ได้ทำ image OCR)

\- การเรียก OCR + AI เป็นแบบ synchronous ทั้ง pipeline — ถ้าไฟล์ใหญ่/webhook ช้า ผู้ใช้ต้องรอจนกว่าจะเสร็จ (มี timeout ตาม UrlFetchApp ของ Apps Script)

\- ไม่มีการ retry อัตโนมัติเมื่อเกิด error



\---



\## 🔮 แผนต่อยอด



\- \[ ] ระบบ Authentication จริง (Google Workspace SSO)

\- \[ ] แจ้งเตือนผ่าน Email/LINE Notify เมื่อมีการอนุมัติ/ปฏิเสธ

\- \[ ] Export รายงาน Skill Gap เป็น PDF/Excel

\- \[ ] รองรับไฟล์ประเภทอื่น (JPG/PNG) นอกจาก PDF

\- \[ ] ระบบ audit log การแก้ไข/อนุมัติของ Admin



\---



\## 📄 License



Internal use — Khon Kaen University (KKU)

