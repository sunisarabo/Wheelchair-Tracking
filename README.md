# WC Tracker — HKT Ground Handling

ระบบบันทึกการให้บริการรถเข็น (Wheelchair Service) สนามบินภูเก็ต
Google Apps Script Web App + Google Sheets

## ไฟล์

| ไฟล์ | ใช้ทำอะไร |
|---|---|
| `Code.gs` | Backend (Apps Script) — API flights / cases, บันทึก TripLog, สรุป DailySummary |
| `Index.html` | หน้าเว็บแอป (มือถือ) — เลือกพนักงาน → เลือกรถเข็น → ลิงก์เคส/ไฟลท์ → จับเวลา |

## ไฟล์ Google Sheets ที่เชื่อมต่อ

1. **Daily Flight Schedule Record 2026** (`FLIGHT_SS_ID`)
   — ตารางบิน tab รายเดือน (`AUG2026`, `SEP2026`, ...) ใช้ดึงรายการไฟลท์ของวันนี้
2. **MMM YYYY PORTER SUMMARY** (`PORTER_SS_ID`)
   — ไฟล์ assign case ของ OCC/VIP (tab รายวัน) ใช้:
   - ดึงเคสที่ assign วันนี้มาแสดงในแอป (เที่ยวบิน, porter, WCHR/WCHS/WCHC, gate, ETA/ETD)
   - เขียนเวลา **รับเคส** (ตอนกดเริ่ม) และ **ส่งเคส** (ตอนกดจบ) กลับไปที่ชีทอัตโนมัติ
     (เขียนเฉพาะช่องที่ยังว่าง — ไม่ทับข้อมูลที่กรอกเองแล้ว)
   - ไฟล์นี้เปลี่ยนทุกเดือน → `PORTER_AUTO_BY_NAME = true` จะหาไฟล์เดือนปัจจุบัน
     จากชื่อไฟล์ใน Drive ให้อัตโนมัติ (fallback เป็น `PORTER_SS_ID`)
3. **WC Tracker Log** — ไฟล์ที่ผูกกับ Apps Script นี้ (TripLog + DailySummary)

## วิธีติดตั้ง / อัปเดต

1. เปิดไฟล์ "WC Tracker Log" → Extensions → Apps Script
2. วาง `Code.gs` ทับของเดิม และวาง `Index.html` ในไฟล์ HTML ชื่อ `Index`
3. Run `setupSheets()` (ครั้งแรก) — อนุญาตสิทธิ์ Sheets + Drive
4. Run `listPorterTabs()` แล้วดู Execution log:
   - ถ้าเจอ `✅ Today tab resolved` = พร้อมใช้
   - ถ้าไม่เจอ ให้ดูชื่อ tab ใน log แล้วใส่รูปแบบชื่อใน `PORTER_TAB_OVERRIDE`
5. Run `testFlights()` / `testCases()` เช็คว่าดึงข้อมูลวันนี้ได้
6. Deploy → New deployment → Web App (Execute as **Me**, access **Anyone**)
7. เอา URL ไปใส่ `GAS_URL` ใน `Index.html` แล้ว deploy ใหม่อีกครั้ง

## Mapping คอลัมน์

ถ้าไฟล์ต้นทางเปลี่ยน layout ให้แก้ค่า index (0-based) ใน `Code.gs`:

- `COL` — Daily Flight Schedule Record (Date=0, Airlines=1, ... Status=20, TYPE OF FLIGHT=21)
- `PCOL` — Porter Summary tab รายวัน (ที่=1, IATA=2, เที่ยวบิน=3, porter=4, ...
  ได้รับแจ้งเคส=13, รับเคส MARK=14, รับเคส เวลา=15, ส่งเคส=16, ประเภท Services=17)
