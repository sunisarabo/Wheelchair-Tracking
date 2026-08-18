// ═══════════════════════════════════════════════════════════════════
// WC TRACKER — Google Apps Script Backend  v3.1
// Phuket Airport · Ground Handling
// ═══════════════════════════════════════════════════════════════════
//
// ── สิ่งที่ต้องตั้งค่า ──
//
//  FLIGHT_SS_ID  : Spreadsheet ID ของ "Daily Flight Schedule Record 2026"
//  PORTER_SS_ID  : Spreadsheet ID ของ "AUG 2026 PORTER SUMMARY" (ไฟล์ assign case)
//                  ↳ ไฟล์นี้เปลี่ยนทุกเดือน — เปิด PORTER_AUTO_BY_NAME ให้หาไฟล์
//                    เดือนปัจจุบันจากชื่อไฟล์อัตโนมัติ (เช่น "SEP 2026 PORTER SUMMARY")
//  LOG_SS_ID     : ใช้ SpreadsheetApp.getActiveSpreadsheet() (ไฟล์ GAS นี้เอง)
//
// ── วิธีติดตั้ง ──
//  1. สร้าง Google Sheets ใหม่ → "WC Tracker Log"
//  2. Extensions > Apps Script
//  3. วาง Code นี้ทั้งหมด + สร้างไฟล์ HTML ชื่อ "Index" วางโค้ดหน้าเว็บ
//  4. เช็ค FLIGHT_SS_ID / PORTER_SS_ID ด้านล่าง
//  5. Run setupSheets() ครั้งแรก (อนุญาตสิทธิ์ Sheets + Drive)
//  6. Run listPorterTabs() แล้วดู Log — เช็คว่าโค้ดหา tab รายวันของ
//     Porter Summary เจอหรือไม่ ถ้าไม่เจอให้ใส่ชื่อ tab ใน PORTER_TAB_OVERRIDE
//  7. Deploy > New Deployment > Web App
//     Execute as: Me | Who has access: Anyone
//  8. Copy URL ไปใส่ใน HTML App (GAS_URL)
// ═══════════════════════════════════════════════════════════════════

// ── CONFIG ──────────────────────────────────────────────────────────
const FLIGHT_SS_ID = '1eN4KFB_yEOWZGn3NViMfq-AD_2Dy-IF7ECaP1C7KQiw';
// ↑ Daily Flight Schedule Record 2026 (tab รายเดือน เช่น AUG2026)

const PORTER_SS_ID = '1kvcr_cDgPUa70Oy4DjADZRNuIJs3QXYjbsCDjbOhLCw';
// ↑ AUG 2026 PORTER SUMMARY (ไฟล์ assign case — เปลี่ยนทุกเดือน)

const PORTER_AUTO_BY_NAME = true;
// ↑ true = หาไฟล์ Porter Summary ของเดือนปัจจุบันจากชื่อ "MMM YYYY PORTER SUMMARY"
//   ใน Drive อัตโนมัติ (ถ้าหาไม่เจอจะ fallback ไปใช้ PORTER_SS_ID)

const PORTER_TAB_OVERRIDE = '';
// ↑ ปกติโค้ดจะหา tab รายวันเอง (จากชื่อ tab หรือวันที่ในหัวตาราง)
//   ถ้าหาไม่เจอ ให้ใส่ชื่อ tab ตรงนี้ เช่น '18' หรือ '18AUG'

const PORTER_WRITEBACK = true;
// ↑ true = ตอน start/end trip ที่ลิงก์กับเคส จะเขียนเวลา "รับเคส/ส่งเคส"
//   กลับไปที่ Porter Summary ให้อัตโนมัติ (เขียนเฉพาะช่องที่ยังว่างเท่านั้น)

const LOG_SHEET   = 'TripLog';
const STAT_SHEET  = 'DailySummary';

// ── COLUMN MAPPING (Daily Flight Schedule Record) ───────────────────
// Layout: (Date) | Airlines | FLT No. | Routing | STA | STD | A/C TYPE | A/C Reg.
//         | ETA | ETD | Estimate | Landing | On Block | Off Block | Airborne
//         | Bay/Gate | REMARK | TSAT | FT | CTOT | Status | TYPE OF FLIGHT | ...
const COL = {
  DATE:      0,   // (blank header) — Date object
  AIRLINE:   1,   // Airlines  เช่น EY, PG, SQ
  FLTNO:     2,   // FLT No.   เช่น "410/411", "270"
  ROUTING:   3,   // Routing   เช่น "AUH-HKT-AUH", "HKT-BKK"
  STA:       4,   // STA (LT)  เช่น 0715, "1900(-1)"
  STD:       5,   // STD (LT)
  ACTYPE:    6,   // A/C TYPE  เช่น B78X, A319
  ETA:       8,   // ETA (LT)
  ETD:       9,   // ETD (LT)
  REMARK:   16,   // REMARK
  STATUS:   20,   // Status         : "On time" / "Delayed" / "Cancelled"
  FLTTYPE:  21,   // TYPE OF FLIGHT : "Quickturn" / "Departure" / "Arrival"
};

// ── COLUMN MAPPING (PORTER SUMMARY — tab รายวัน) ────────────────────
// A(0)=STATUS | B(1)=ที่ | C(2)=สายการบิน IATA | D(3)=เที่ยวบินที่
// E(4)=ชื่อพนักงาน (porter) | F(5)=สถานะ | G(6)=ETA | H(7)=เวลาสแตนบายขาเข้า
// I(8)=ETD | J(9)=เวลาบอร์ดดิ้ง | K(10)=เช็ค Porter standby gate (checkbox)
// L(11)=ชื่อ LP | M(12)=ประตูทางออก | N(13)=ได้รับแจ้งเคส (เวลา)
// O(14)=รับเคส MARK (checkbox) | P(15)=รับเคส เวลา | Q(16)=ส่งเคส (เวลา)
// R(17)=ประเภท Services (WCHR/WCHS/WCHC) | S(18)=ขาเข้า ✓ | T(19)=ขาออก ✓
// U(20)=ตรวจสอบแล้ว ✓ | V(21)=ระยะเวลารอ (สูตร) | W(22)=REMARK | X(23)=SEAT NO.
const PCOL = {
  NO:            1,
  AIRLINE:       2,
  FLTNO:         3,
  PORTER:        4,
  STATUS:        5,
  ETA:           6,
  ETD:           8,
  GATE:         12,
  NOTIFIED:     13,
  PICKUP_MARK:  14,
  PICKUP_TIME:  15,
  DELIVER_TIME: 16,
  SVC:          17,
  ARR_CHK:      18,
  DEP_CHK:      19,
  REMARK:       22,
};

// ── HEADERS ─────────────────────────────────────────────────────────
const LOG_HEADERS = [
  'Trip ID','วันที่','รถเข็น','ประเภท WC',
  'รหัสพนักงาน','ชื่อเล่น','ชื่อเต็ม',
  'เที่ยวบิน','สายการบิน','Routing',
  'STA','STD',
  'เวลาเริ่ม','เวลาจบ','ระยะเวลา (นาที)','สถานะ',
  'บันทึกเมื่อ',
  'Case No','SSR','Porter (ชีท)'
];

// ═══════════════════════════════════════════════════════════════════
// doGet — Web App UI + API (flights / cases / health)
// ═══════════════════════════════════════════════════════════════════
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || '';

  // API: today's flight list
  if (action === 'flights') {
    try { return jsonOk({ flights: getTodayFlights() }); }
    catch (err) { return jsonErr('flights error: ' + err.message); }
  }

  // API: today's assigned cases (Porter Summary)
  if (action === 'cases') {
    try { return jsonOk(getTodayCases()); }
    catch (err) { return jsonErr('cases error: ' + err.message); }
  }

  // API: health check
  if (action === 'health') {
    return jsonOk({ app: 'WC Tracker HKT', version: '3.1', time: new Date().toISOString() });
  }

  // Default: serve the web-app UI
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('WC Service — HKT')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ═══════════════════════════════════════════════════════════════════
// doPost — รับ trip start / end จาก HTML App
// ═══════════════════════════════════════════════════════════════════
function doPost(e) {
  try {
    const data   = JSON.parse(e.postData.contents);
    const action = data.action;
    const trip   = data.trip;

    if      (action === 'start') return jsonOk(handleStart(trip));
    else if (action === 'end')   return jsonOk(handleEnd(trip));
    else                          return jsonErr('Unknown action: ' + action);

  } catch (err) {
    Logger.log('doPost error: ' + err.message);
    return jsonErr(err.message);
  }
}

// ═══════════════════════════════════════════════════════════════════
// getTodayFlights — ดึง flight จาก Daily Flight Schedule Record
// ═══════════════════════════════════════════════════════════════════
function getTodayFlights() {
  const today   = new Date();
  const tabName = getFlightTabName(today);

  Logger.log('Looking for tab: ' + tabName + ' | today: ' + today.toDateString());

  let ss;
  try {
    ss = SpreadsheetApp.openById(FLIGHT_SS_ID);
  } catch (e) {
    throw new Error('Cannot open Flight Spreadsheet. Check FLIGHT_SS_ID. ' + e.message);
  }

  const sheet = ss.getSheetByName(tabName);
  if (!sheet) {
    Logger.log('Tab not found: ' + tabName);
    return [];
  }

  const data = sheet.getDataRange().getValues();
  const flights = [];

  const tDay   = today.getDate();
  const tMonth = today.getMonth();
  const tYear  = today.getFullYear();

  for (let r = 1; r < data.length; r++) {
    const row     = data[r];
    if (!isSameBkkDate(row[COL.DATE], tDay, tMonth, tYear)) continue;

    const airline  = String(row[COL.AIRLINE] || '').trim();
    const fltno    = String(row[COL.FLTNO]   || '').trim();
    const routing  = String(row[COL.ROUTING] || '').trim();

    if (!airline || !fltno) continue;

    // ตัด Cancelled / Postponed — เช็คจาก Status, TYPE OF FLIGHT และ ETA/ETD
    // (บางแถวเจ้าหน้าที่พิมพ์ "Cancelled" ลงช่อง ETA/ETD แทน)
    const cancelCheck = [COL.ETA, COL.ETD, COL.STATUS, COL.FLTTYPE]
      .map(i => String(row[i] || ''))
      .join(' ');
    if (/cancel|postpon/i.test(cancelCheck)) continue;

    const sta = fmtTime(row[COL.STA]);
    const std = fmtTime(row[COL.STD]);

    // "EY 410/411" → เอาแค่ขาแรก
    const fltCode = (airline + fltno.split('/')[0]).replace(/\s/g, '');

    // ระบุ DEP / ARR / QT — ใช้ column "TYPE OF FLIGHT" ก่อน, fallback จาก Routing
    const type = getFlightTypeFromCol(row[COL.FLTTYPE]) || getFlightType(routing);

    flights.push({
      code:    fltCode,
      airline: airline,
      fltno:   fltno,
      routing: routing,
      sta:     sta,
      std:     std,
      type:    type,           // 'DEP' | 'ARR' | 'QT'
      display: buildDisplay(airline, fltno, routing, sta, std, type)
    });
  }

  // sort by time (DEP by STD, ARR by STA)
  flights.sort((a, b) => {
    const ta = a.type === 'DEP' ? a.std : a.sta;
    const tb = b.type === 'DEP' ? b.std : b.sta;
    return ta.localeCompare(tb);
  });

  Logger.log('Flights found: ' + flights.length);
  return flights;
}

// ── helper: เทียบค่า date ใน cell กับวันนี้ (Bangkok time) ──
function isSameBkkDate(rawDate, tDay, tMonth, tYear) {
  let d = null;
  if (rawDate instanceof Date && !isNaN(rawDate)) d = rawDate;
  else if (rawDate) {
    const p = new Date(rawDate);
    if (!isNaN(p)) d = p;
  }
  if (!d) return false;
  const bkk = new Date(d.getTime() + 7 * 3600 * 1000);
  return bkk.getUTCDate()     === tDay
      && bkk.getUTCMonth()    === tMonth
      && bkk.getUTCFullYear() === tYear;
}

// ── helper: ได้ชื่อ tab เช่น AUG2026 จากวันที่ ──
function getFlightTabName(date) {
  const months = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
  return months[date.getMonth()] + date.getFullYear();
}

// ── helper: แปลง "TYPE OF FLIGHT" → DEP/ARR/QT ──
function getFlightTypeFromCol(val) {
  const s = String(val || '').toLowerCase();
  if (s.indexOf('quick') >= 0) return 'QT';
  if (s.indexOf('dep')   >= 0) return 'DEP';
  if (s.indexOf('arr')   >= 0) return 'ARR';
  return null;
}

// ── helper: ระบุประเภทเที่ยวบิน จาก Routing ──
function getFlightType(routing) {
  const parts = routing.split('-');
  const hasHktFirst = parts[0] === 'HKT';
  const hktCount    = parts.filter(p => p === 'HKT').length;

  if (hktCount >= 2) return 'QT';    // Quickturn ผ่าน HKT
  if (hasHktFirst)   return 'DEP';   // ออกจาก HKT
  return 'ARR';                       // เข้า HKT
}

// ── helper: สร้างข้อความแสดงใน dropdown ──
function buildDisplay(airline, fltno, routing, sta, std, type) {
  const time = type === 'DEP' ? ('STD ' + std) : ('STA ' + sta);
  const dir  = type === 'DEP' ? '→' : type === 'ARR' ? '←' : '⇄';
  return airline + fltno.split('/')[0] + ' ' + dir + ' ' + routing + ' [' + time + ']';
}

// ═══════════════════════════════════════════════════════════════════
// PORTER SUMMARY — assign case integration
// ═══════════════════════════════════════════════════════════════════

// ── หาไฟล์ Porter Summary ของเดือนปัจจุบัน ──
function getPorterSpreadsheet(date) {
  if (PORTER_AUTO_BY_NAME) {
    try {
      const name  = getPorterFileName(date);   // เช่น "AUG 2026 PORTER SUMMARY"
      const files = DriveApp.getFilesByName(name);
      if (files.hasNext()) return SpreadsheetApp.open(files.next());
      Logger.log('Porter file not found by name: "' + name + '" — fallback to PORTER_SS_ID');
    } catch (e) {
      Logger.log('Porter file search error: ' + e.message + ' — fallback to PORTER_SS_ID');
    }
  }
  return SpreadsheetApp.openById(PORTER_SS_ID);
}

function getPorterFileName(date) {
  const months = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
  return months[date.getMonth()] + ' ' + date.getFullYear() + ' PORTER SUMMARY';
}

// ── หา tab รายวันของวันนี้ ──
// 1) PORTER_TAB_OVERRIDE  2) เดาจากชื่อ tab  3) หา Date ในหัวตาราง (แถว 1-8)
function findPorterDayTab(ss, date) {
  if (PORTER_TAB_OVERRIDE) {
    const s = ss.getSheetByName(PORTER_TAB_OVERRIDE);
    if (s) return s;
  }

  const months = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
  const d   = date.getDate();
  const dd  = String(d).padStart(2, '0');
  const mmm = months[date.getMonth()];
  const m1  = date.getMonth() + 1;
  const yy  = String(date.getFullYear()).slice(-2);
  const candidates = [
    String(d), dd,
    d + mmm, dd + mmm, d + ' ' + mmm, dd + ' ' + mmm,
    dd + mmm + yy, d + mmm + yy,
    d + '/' + m1, dd + '/' + String(m1).padStart(2, '0'),
    dd + '-' + String(m1).padStart(2, '0'),
  ].map(normTabName);

  const sheets = ss.getSheets();
  for (const sh of sheets) {
    if (candidates.indexOf(normTabName(sh.getName())) >= 0) return sh;
  }

  // fallback: หา Date object ที่ตรงกับวันนี้ ในหัวตาราง (บริเวณ "ใส่วันที่นะครับ")
  const tDay = date.getDate(), tMonth = date.getMonth(), tYear = date.getFullYear();
  for (const sh of sheets) {
    const rows = Math.min(8, sh.getLastRow());
    const cols = Math.min(10, sh.getLastColumn());
    if (rows < 1 || cols < 1) continue;
    const vals = sh.getRange(1, 1, rows, cols).getValues();
    for (const row of vals) {
      for (const v of row) {
        if (v instanceof Date && !isNaN(v)
            && v.getDate() === tDay && v.getMonth() === tMonth && v.getFullYear() === tYear) {
          return sh;
        }
      }
    }
  }
  return null;
}

function normTabName(s) {
  return String(s).toUpperCase().replace(/[\s._]/g, '');
}

// ── ดึงเคสที่ assign ของวันนี้ ──
function getTodayCases() {
  const today = new Date();
  let ss;
  try {
    ss = getPorterSpreadsheet(today);
  } catch (e) {
    throw new Error('Cannot open Porter Summary. Check PORTER_SS_ID. ' + e.message);
  }

  const sheet = findPorterDayTab(ss, today);
  if (!sheet) {
    Logger.log('Porter day tab not found for ' + today.toDateString()
      + ' — run listPorterTabs() and set PORTER_TAB_OVERRIDE if needed');
    return { cases: [], tab: null };
  }

  const data  = sheet.getDataRange().getValues();
  const cases = [];

  for (let r = 0; r < data.length; r++) {
    const row  = data[r];
    const no   = row[PCOL.NO];
    const iata = String(row[PCOL.AIRLINE] || '').trim();
    const flt  = String(row[PCOL.FLTNO]   || '').trim();

    // แถวเคสจริง = "ที่" เป็นตัวเลข และมี IATA code (กันแถวคำแนะนำ/หัวตาราง)
    const noNum = (typeof no === 'number') ? no : parseInt(no, 10);
    if (!noNum || isNaN(noNum) || !iata || !flt) continue;

    cases.push({
      row:       r + 1,                                    // แถวจริงใน Sheet (1-based)
      no:        noNum,
      airline:   iata,
      fltno:     flt,
      flight:    (iata + flt.split('/')[0]).replace(/\s/g, ''),
      porter:    String(row[PCOL.PORTER] || '').trim(),
      status:    String(row[PCOL.STATUS] || '').trim(),
      eta:       fmtTime(row[PCOL.ETA]),
      etd:       fmtTime(row[PCOL.ETD]),
      gate:      String(row[PCOL.GATE] || '').trim(),
      svc:       String(row[PCOL.SVC]  || '').trim(),      // WCHR / WCHS / WCHC
      dir:       row[PCOL.DEP_CHK] === true ? 'DEP' : (row[PCOL.ARR_CHK] === true ? 'ARR' : ''),
      notified:  fmtTime(row[PCOL.NOTIFIED]),
      pickup:    fmtTime(row[PCOL.PICKUP_TIME]),
      delivered: fmtTime(row[PCOL.DELIVER_TIME]),
      remark:    String(row[PCOL.REMARK] || '').trim(),
    });
  }

  Logger.log('Cases found: ' + cases.length + ' | tab: ' + sheet.getName());
  return { cases: cases, tab: sheet.getName() };
}

// ── เขียนเวลา รับเคส / ส่งเคส กลับไปที่ Porter Summary ──
// phase: 'pickup' (ตอน start) | 'deliver' (ตอน end)
// เขียนเฉพาะช่องที่ยังว่าง — ไม่ทับข้อมูลที่ OCC/LP กรอกเองแล้ว
function markCase(trip, phase, timeStr) {
  if (!PORTER_WRITEBACK) return;
  if (!trip || !trip.caseRow) return;

  try {
    const today = new Date();
    const ss    = getPorterSpreadsheet(today);
    const sheet = findPorterDayTab(ss, today);
    if (!sheet) { Logger.log('markCase: day tab not found'); return; }

    let row = parseInt(trip.caseRow, 10);
    const maxRow = sheet.getLastRow();

    // ตรวจว่าแถวยังเป็นเคสเดิม (กันกรณี OCC แทรก/ลบแถวหลังแอปโหลดข้อมูลไป)
    const rowMatches = (r) => {
      if (r < 1 || r > maxRow) return false;
      const v = sheet.getRange(r, 1, 1, PCOL.SVC + 1).getValues()[0];
      return String(v[PCOL.AIRLINE]).trim() === String(trip.caseAirline || '').trim()
          && String(v[PCOL.FLTNO]).trim()   === String(trip.caseFltno   || '').trim();
    };

    if (!rowMatches(row)) {
      // หาแถวใหม่จาก airline + fltno + porter
      row = -1;
      const data = sheet.getDataRange().getValues();
      for (let r = 0; r < data.length; r++) {
        if (String(data[r][PCOL.AIRLINE]).trim() === String(trip.caseAirline || '').trim()
         && String(data[r][PCOL.FLTNO]).trim()   === String(trip.caseFltno   || '').trim()
         && (!trip.casePorter || String(data[r][PCOL.PORTER]).trim() === String(trip.casePorter).trim())) {
          row = r + 1;
          break;
        }
      }
      if (row < 0) { Logger.log('markCase: case row not found for ' + trip.caseAirline + trip.caseFltno); return; }
    }

    const t = timeStr || Utilities.formatDate(new Date(), 'Asia/Bangkok', 'HH:mm');

    if (phase === 'pickup') {
      const cur = sheet.getRange(row, PCOL.PICKUP_TIME + 1).getValue();
      if (!cur) {
        sheet.getRange(row, PCOL.PICKUP_MARK + 1).setValue(true);
        sheet.getRange(row, PCOL.PICKUP_TIME + 1).setValue(t);
        Logger.log('markCase pickup: row ' + row + ' = ' + t);
      }
    } else if (phase === 'deliver') {
      const cur = sheet.getRange(row, PCOL.DELIVER_TIME + 1).getValue();
      if (!cur) {
        sheet.getRange(row, PCOL.DELIVER_TIME + 1).setValue(t);
        Logger.log('markCase deliver: row ' + row + ' = ' + t);
      }
    }
  } catch (e) {
    // write-back พังต้องไม่ทำให้การบันทึก trip ล้ม
    Logger.log('markCase error (' + phase + '): ' + e.message);
  }
}

// ═══════════════════════════════════════════════════════════════════
// handleStart
// ═══════════════════════════════════════════════════════════════════
function handleStart(trip) {
  if (!trip || !trip.id) throw new Error('No trip data');

  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const log = getOrCreateSheet(ss, LOG_SHEET, LOG_HEADERS);
  const now = new Date();
  const today = Utilities.formatDate(now, 'Asia/Bangkok', 'dd/MM/yyyy');
  const startTime = trip.startTime || Utilities.formatDate(now, 'Asia/Bangkok', 'HH:mm:ss');

  log.appendRow([
    trip.id,
    today,
    trip.ctrl     || '',
    trip.type     || '',
    trip.sid      || '',
    trip.snick    || '',
    trip.sname    || '',
    trip.flight   || '',
    trip.flAirline|| '',
    trip.flRouting|| '',
    trip.flSTA    || '',
    trip.flSTD    || '',
    startTime,
    '',   // endTime
    '',   // duration
    'กำลังบริการ',
    Utilities.formatDate(now, 'Asia/Bangkok', 'dd/MM/yyyy HH:mm:ss'),
    trip.caseNo     || '',
    trip.caseSvc    || '',
    trip.casePorter || ''
  ]);

  // write-back "รับเคส" → Porter Summary
  markCase(trip, 'pickup', startTime.slice(0, 5));

  Logger.log('START: ' + trip.id + ' | ' + trip.snick + ' | ' + trip.ctrl
    + (trip.caseNo ? ' | case #' + trip.caseNo : ''));
  return { action: 'start', tripId: trip.id };
}

// ═══════════════════════════════════════════════════════════════════
// handleEnd
// ═══════════════════════════════════════════════════════════════════
function handleEnd(trip) {
  if (!trip || !trip.id) throw new Error('No trip data');

  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const log = getOrCreateSheet(ss, LOG_SHEET, LOG_HEADERS);
  const data = log.getDataRange().getValues();
  const now  = new Date();
  const dur  = parseFloat(trip.duration) || 0;
  const endTime = trip.endTime || Utilities.formatDate(now, 'Asia/Bangkok', 'HH:mm:ss');

  let found = false;
  for (let r = 1; r < data.length; r++) {
    if (String(data[r][0]) !== String(trip.id)) continue;

    const row = r + 1;
    log.getRange(row, 14).setValue(endTime);               // เวลาจบ
    log.getRange(row, 15).setValue(dur);                   // ระยะเวลา
    log.getRange(row, 16).setValue('เสร็จสิ้น');           // สถานะ
    log.getRange(row, 17).setValue(
      Utilities.formatDate(now, 'Asia/Bangkok', 'dd/MM/yyyy HH:mm:ss'));

    // Color-code duration cell
    const durCell = log.getRange(row, 15);
    if (dur > 60)      { durCell.setBackground('#fee2e2').setFontColor('#dc2626'); }
    else if (dur > 30) { durCell.setBackground('#fef3c7').setFontColor('#d97706'); }
    else               { durCell.setBackground('#dcfce7').setFontColor('#16a34a'); }

    found = true;
    break;
  }

  if (!found) {
    // START ไม่ได้ sync → append แถวใหม่
    Logger.log('END not found — appending: ' + trip.id);
    const today = Utilities.formatDate(now, 'Asia/Bangkok', 'dd/MM/yyyy');
    log.appendRow([
      trip.id, today, trip.ctrl||'', trip.type||'',
      trip.sid||'', trip.snick||'', trip.sname||'',
      trip.flight||'', trip.flAirline||'', trip.flRouting||'',
      trip.flSTA||'', trip.flSTD||'',
      trip.startTime||'', endTime, dur,
      'เสร็จสิ้น',
      Utilities.formatDate(now, 'Asia/Bangkok', 'dd/MM/yyyy HH:mm:ss'),
      trip.caseNo||'', trip.caseSvc||'', trip.casePorter||''
    ]);
  }

  // write-back "ส่งเคส" → Porter Summary
  markCase(trip, 'deliver', endTime.slice(0, 5));

  updateDailySummary(ss);

  Logger.log('END: ' + trip.id + ' | ' + trip.snick + ' | ' + dur + ' min');
  return { action: 'end', tripId: trip.id, duration: dur };
}

// ═══════════════════════════════════════════════════════════════════
// updateDailySummary
// ═══════════════════════════════════════════════════════════════════
function updateDailySummary(ss) {
  try {
    const log = ss.getSheetByName(LOG_SHEET);
    if (!log) return;
    const stat = getOrCreateSheet(ss, STAT_SHEET, [
      'วันที่','Trips ทั้งหมด','Trips เสร็จ','เฉลี่ย (น.)','นานสุด (น.)','สั้นสุด (น.)','> 60 น.','อัปเดต'
    ]);

    const today    = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy');
    const allRows  = log.getDataRange().getValues().slice(1);
    const todayRows = allRows.filter(r => String(r[1]) === today);
    const doneRows  = todayRows.filter(r => String(r[15]) === 'เสร็จสิ้น');
    const durs      = doneRows.map(r => parseFloat(r[14])).filter(d => !isNaN(d) && d > 0);

    const rowData = [
      today,
      todayRows.length,
      doneRows.length,
      durs.length ? (durs.reduce((a,b)=>a+b,0)/durs.length).toFixed(1) : '',
      durs.length ? Math.max(...durs).toFixed(1) : '',
      durs.length ? Math.min(...durs).toFixed(1) : '',
      durs.filter(d=>d>60).length,
      Utilities.formatDate(new Date(), 'Asia/Bangkok', 'HH:mm:ss')
    ];

    const statData = stat.getDataRange().getValues();
    let statRow = -1;
    for (let r = 1; r < statData.length; r++) {
      if (String(statData[r][0]) === today) { statRow = r + 1; break; }
    }

    if (statRow > 0) stat.getRange(statRow, 1, 1, rowData.length).setValues([rowData]);
    else             stat.appendRow(rowData);

  } catch(e) {
    Logger.log('updateDailySummary error: ' + e.message);
  }
}

// ═══════════════════════════════════════════════════════════════════
// Helpers
// ═══════════════════════════════════════════════════════════════════
function getOrCreateSheet(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (headers && headers.length) {
      const hRow = sh.getRange(1, 1, 1, headers.length);
      hRow.setValues([headers]);
      hRow.setBackground('#1d6fe8').setFontColor('#ffffff').setFontWeight('bold');
      sh.setFrozenRows(1);
    }
  } else if (headers && headers.length && sh.getLastColumn() < headers.length) {
    // v3.1 เพิ่มคอลัมน์ Case No / SSR / Porter — เติม header ให้ sheet เดิม
    const hRow = sh.getRange(1, 1, 1, headers.length);
    hRow.setValues([headers]);
    hRow.setBackground('#1d6fe8').setFontColor('#ffffff').setFontWeight('bold');
  }
  return sh;
}

// ── helper: แปลงเวลา → "HH:mm"
//    รองรับ: Date object, 715, "0715", "1900(-1)", "7:15", "07.15" ──
function fmtTime(val) {
  if (val instanceof Date && !isNaN(val)) {
    return Utilities.formatDate(val, 'Asia/Bangkok', 'HH:mm');
  }
  if (!val && val !== 0) return '';
  const s = String(val).trim();

  // "7:15" / "07.15"
  let m = s.match(/^(\d{1,2})[:.](\d{2})/);
  if (m) return m[1].padStart(2,'0') + ':' + m[2];

  // ตัวเลข 3-4 หลักตัวแรกในข้อความ เช่น "1900(-1)" → 1900
  m = s.match(/\d{3,4}/);
  if (!m) return s;
  const n = parseInt(m[0], 10);
  const h = Math.floor(n / 100);
  const mm = n % 100;
  if (h > 23 || mm > 59) return s;
  return String(h).padStart(2,'0') + ':' + String(mm).padStart(2,'0');
}

function jsonOk(obj) {
  return ContentService.createTextOutput(JSON.stringify({ ok: true, ...obj }))
    .setMimeType(ContentService.MimeType.JSON);
}

function jsonErr(msg) {
  return ContentService.createTextOutput(JSON.stringify({ ok: false, error: msg }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ═══════════════════════════════════════════════════════════════════
// setupSheets — รันครั้งแรก
// ═══════════════════════════════════════════════════════════════════
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  getOrCreateSheet(ss, LOG_SHEET, LOG_HEADERS);
  getOrCreateSheet(ss, STAT_SHEET, [
    'วันที่','Trips ทั้งหมด','Trips เสร็จ','เฉลี่ย (น.)','นานสุด (น.)','สั้นสุด (น.)','> 60 น.','อัปเดต'
  ]);
  ss.rename('WC Tracker Log — HKT Ground Handling');
  SpreadsheetApp.getUi().alert(
    '✅ ตั้งค่าเสร็จ!\n\n' +
    'Sheets ที่สร้าง:\n• TripLog\n• DailySummary\n\n' +
    'ขั้นตอนต่อไป:\n1. เช็ค FLIGHT_SS_ID / PORTER_SS_ID\n' +
    '2. Run listPorterTabs() เช็คว่าเจอ tab รายวัน\n' +
    '3. Deploy > New Deployment > Web App\n' +
    '4. Copy URL ใส่ใน HTML App'
  );
}

// ═══════════════════════════════════════════════════════════════════
// listPorterTabs — debug: ดูชื่อ tab ทั้งหมดใน Porter Summary
//                  และเช็คว่า resolver หา tab วันนี้เจอไหม
// ═══════════════════════════════════════════════════════════════════
function listPorterTabs() {
  const today = new Date();
  const ss = getPorterSpreadsheet(today);
  Logger.log('Porter file: ' + ss.getName());
  ss.getSheets().forEach(sh => Logger.log('  tab: "' + sh.getName() + '"'));
  const day = findPorterDayTab(ss, today);
  Logger.log(day
    ? '✅ Today tab resolved: "' + day.getName() + '"'
    : '❌ Today tab NOT found — ใส่ชื่อ tab ใน PORTER_TAB_OVERRIDE');
}

// ═══════════════════════════════════════════════════════════════════
// testFlights / testCases — ทดสอบดึงข้อมูลวันนี้
// ═══════════════════════════════════════════════════════════════════
function testFlights() {
  Logger.log('=== TEST FLIGHTS ===');
  try {
    const flights = getTodayFlights();
    Logger.log('Total: ' + flights.length);
    flights.slice(0, 5).forEach(f => Logger.log(JSON.stringify(f)));
  } catch(e) {
    Logger.log('ERROR: ' + e.message);
  }
}

function testCases() {
  Logger.log('=== TEST CASES ===');
  try {
    const res = getTodayCases();
    Logger.log('Tab: ' + res.tab + ' | Total: ' + res.cases.length);
    res.cases.slice(0, 5).forEach(c => Logger.log(JSON.stringify(c)));
  } catch(e) {
    Logger.log('ERROR: ' + e.message);
  }
}

// ═══════════════════════════════════════════════════════════════════
// testPost — ทดสอบ start/end trip
// ═══════════════════════════════════════════════════════════════════
function testPost() {
  const trip = {
    id:'TRP-TEST-001', ctrl:'AOTGA-WH-15', no:'WH-15', type:'ใหญ่',
    sid:'2101042', snick:'สุทิน', sname:'สุทิน ครองรั้ว',
    flight:'PG270', flAirline:'PG', flRouting:'HKT-BKK',
    flSTA:'', flSTD:'0735',
    startMs: Date.now()-1800000, startTime:'08:30:00', status:'active'
  };
  Logger.log('=== TEST START ===');
  Logger.log(JSON.stringify(handleStart(trip)));

  trip.endTime='09:00:00'; trip.duration='30.0'; trip.status='done';
  Logger.log('=== TEST END ===');
  Logger.log(JSON.stringify(handleEnd(trip)));
}
