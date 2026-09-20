// ═══════════════════════════════════════════════════════════════════
// WC TRACKER — Google Apps Script Backend  v5.0
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
//
// ── v5.0: แยกบทบาท LP / Porter ──
//  LP     : จ่ายเคสให้ porter จากในแอปได้เลย (เขียนชื่อกลับเข้า Porter Summary)
//           และเปิดเคสใหม่จากตารางบินได้
//  Porter : เห็นเฉพาะเคสที่ถูกจ่ายให้ตัวเอง แล้วกดเริ่ม/จบงานเอง
//           งานที่ค้างอยู่เก็บฝั่ง server — เริ่มเครื่องหนึ่งไปกดจบอีกเครื่องได้
// ═══════════════════════════════════════════════════════════════════

// ── CONFIG ──────────────────────────────────────────────────────────
const FLIGHT_SS_ID = '1eN4KFB_yEOWZGn3NViMfq-AD_2Dy-IF7ECaP1C7KQiw';
// ↑ Daily Flight Schedule Record 2026 (tab รายเดือน เช่น AUG2026)

const PORTER_SS_ID = '1cGLIpFPthPedkWjcQM_A4PFyKOGvuHAFCiJr-Jvs8II';
// ↑ SEP 2026 PORTER SUMMARY (ไฟล์ assign case — เปลี่ยนทุกเดือน)
//   tab รายวันชื่อรูปแบบ 01SEP26 / 14SEP26 (โค้ดหาเองอัตโนมัติ)

const PORTER_AUTO_BY_NAME = true;
// ↑ true = หาไฟล์ Porter Summary ของเดือนปัจจุบันจากชื่อ "MMM YYYY PORTER SUMMARY"
//   ใน Drive อัตโนมัติ (ถ้าหาไม่เจอจะ fallback ไปใช้ PORTER_SS_ID)

const PORTER_TAB_OVERRIDE = '';
// ↑ ปกติโค้ดจะหา tab รายวันเอง (จากชื่อ tab หรือวันที่ในหัวตาราง)
//   ถ้าหาไม่เจอ ให้ใส่ชื่อ tab ตรงนี้ เช่น '18' หรือ '18AUG'

// ── PRE-WHEELCHAIR (ยอดจองล่วงหน้า + ตารางเวร porter) ───────────────
const PRE_SS_ID = '1m_l9V0OAS2iW1KmsMEfzOZiQI_s_QybsilkeLxCAbSY';
// ↑ SEP 2026 PRE-WHEELCHAIR — อยู่ในโฟลเดอร์ <PRE_ROOT_FOLDER_ID>/SEP 2026
const PRE_ROOT_FOLDER_ID = '1xiTMfPDQ0nRLf0eKYjhn2SIeiP8-HRwb';
const PRE_AUTO_BY_NAME = true;
// ↑ true = หาไฟล์ "MMM YYYY PRE-WHEELCHAIR" ของเดือนปัจจุบันเองจาก Drive

const PORTER_WRITEBACK = true;
// ↑ true = ตอน start/end trip ที่ลิงก์กับเคส จะเขียนเวลา "รับเคส/ส่งเคส"
//   กลับไปที่ Porter Summary ให้อัตโนมัติ (เขียนเฉพาะช่องที่ยังว่างเท่านั้น)

const LOG_SHEET   = 'TripLog';
const STAT_SHEET  = 'DailySummary';
const QUEUE_SHEET = 'Queue';
const WC_SHEET    = 'WCStatus';
const ASSIGN_SHEET = 'AssignLog';   // ประวัติการจ่ายงานของ LP (ใครจ่ายเคสไหนให้ใคร เมื่อไหร่)

// งานที่เริ่มมานานกว่านี้ = "ค้าง" ไม่ใช่ "กำลังทำอยู่"
// เคสรถเข็นหนึ่งเคสใช้เวลาเป็นสิบนาที ไม่ใช่เป็นวัน — ที่ค้างคือคนลืมกดจบ
// หรือมือถือดับ/ปิดแอประหว่างงาน ตั้งเผื่อกะยาวสุด (17:00–05:00) ไว้แล้ว
const STALE_HOURS = 12;

const WC_HEADERS = ['รถเข็น','สถานะ','อาการ/หมายเหตุ','แจ้งโดย','แจ้งเมื่อ','อัปเดตล่าสุด'];
// ↑ สถานะ: 'ซ่อม' = ใช้งานไม่ได้ · 'ใช้งานได้' = ซ่อมเสร็จ/กลับมาใช้ได้

const QUEUE_HEADERS = ['วันที่','ลำดับ','รหัสพนักงาน','ชื่อเล่น','กะ','สถานะ','Trips วันนี้','อัปเดต'];
// ↑ "ลำดับ" ใช้เป็นตัวตัดสินเมื่อจำนวนเคสเท่ากัน (คนที่ว่างก่อน/เข้าคิวก่อนขึ้นก่อน)
//   ลำดับจริงของคิว = เรียงตาม "Trips วันนี้" น้อยสุดขึ้นก่อน แยกตามกะ

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

// ── COLUMN MAPPING (PORTER SUMMARY — บล็อก STAFF RECORD ด้านขวา) ─────
// Z(25)=NO. | AA(26)=SKED | AB(27)=NAME | AC(28)=CASE SUMMARY
// = รายชื่อพนักงานที่ "มาทำงานวันนั้น" ที่ LP/OCC กรอกไว้
const SCOL = { NO: 25, SKED: 26, NAME: 27, CASES: 28 };

// ── COLUMN MAPPING (PRE-WHEELCHAIR — tab รายวัน เช่น 20SEP26) ───────
// บล็อก A: ไฟลท์ + ยอดจองล่วงหน้า (0-20)
// บล็อก B: จำนวนเคสแต่ละช่วงเวลา / MANPOWER (23-26, สองตารางซ้อนกัน)
// บล็อก C: Porter Daily Assignment = ตารางเวรจริง (28-32)
const FCOL = {
  AIRLINE:  0, FLTNO: 1, ROUTING: 2,
  STA_RAW:  3, STD_RAW: 4, STA: 5, STD: 6,
  CT_OPEN:  7, CT_CLOSE: 8,
  A_WCHR:   9, A_WCHS: 10, A_WCHC: 11, A_AVIH: 12, A_MAAS: 13, A_TOTAL: 14,
  D_WCHR:  15, D_WCHS: 16, D_WCHC: 17, D_AVIH: 18, D_MAAS: 19, D_TOTAL: 20,
  WIN_LBL: 23, WIN_1: 25, WIN_2: 26,   // demand: ขาเข้า/ขาออก · manpower: SKED/OT
  R_NAME:  28, R_IN: 29, R_OUT: 30, R_OT_IN: 31, R_OT_OUT: 32,
};

// รหัสสถานะในตารางเวร (ไม่ใช่เวลา = ไม่ได้มาทำงาน)
const OFF_CODES = ['OFF','SL','BL','AL','ML','PL','X','-'];

// ── HEADERS ─────────────────────────────────────────────────────────
const LOG_HEADERS = [
  'Trip ID','วันที่','รถเข็น','ประเภท WC',
  'รหัสพนักงาน','ชื่อเล่น','ชื่อเต็ม',
  'เที่ยวบิน','สายการบิน','Routing',
  'STA','STD',
  'เวลาเริ่ม','เวลาจบ','ระยะเวลา (นาที)','สถานะ',
  'บันทึกเมื่อ',
  'Case No','SSR','Porter (ชีท)',
  'Case Row','Case Airline','Case Flt'
];

// ประวัติการจ่ายงาน (LP กดจ่ายเคสผ่านแอป)
const ASSIGN_HEADERS = [
  'บันทึกเมื่อ','วันที่','แถวในชีท','เคส','เที่ยวบิน','SSR',
  'จ่ายให้','คนเดิม','จ่ายโดย','หมายเหตุ'
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

  // API: today's on-duty staff (STAFF RECORD block)
  if (action === 'staff') {
    try { return jsonOk(getTodayStaff()); }
    catch (err) { return jsonErr('staff error: ' + err.message); }
  }

  // API: สถานะรถเข็น (เฉพาะคันที่แจ้งซ่อมอยู่)
  if (action === 'wcstatus') {
    try { return jsonOk({ repair: getWCRepair() }); }
    catch (err) { return jsonErr('wcstatus error: ' + err.message); }
  }

  // API: สรุปเคสวันนี้
  if (action === 'summary') {
    try { return jsonOk({ summary: getTodaySummary() }); }
    catch (err) { return jsonErr('summary error: ' + err.message); }
  }

  // API: ยอดจองล่วงหน้า + ตารางเวร + demand/manpower
  if (action === 'pre') {
    try { return jsonOk(getPreBooking()); }
    catch (err) { return jsonErr('pre error: ' + err.message); }
  }

  // API: ตารางเวร porter วันนี้ (จาก PRE-WHEELCHAIR)
  if (action === 'roster') {
    try {
      const pre = getPreBooking();
      return jsonOk({ tab: pre.tab, roster: pre.roster });
    } catch (err) { return jsonErr('roster error: ' + err.message); }
  }

  // API: ข้อมูลทั้งหน้าจอในครั้งเดียว (ลดจำนวน request จากมือถือ)
  if (action === 'board') {
    try { return jsonOk(getBoard()); }
    catch (err) { return jsonErr('board error: ' + err.message); }
  }

  // API: today's porter queue
  if (action === 'queue') {
    try { return jsonOk({ queue: getQueue() }); }
    catch (err) { return jsonErr('queue error: ' + err.message); }
  }

  // API: งานที่ยังไม่จบ (ใช้กู้สถานะเมื่อรีเฟรช / เปลี่ยนเครื่อง)
  if (action === 'active') {
    try { return jsonOk({ active: getActiveTrips() }); }
    catch (err) { return jsonErr('active error: ' + err.message); }
  }

  // API: งานที่เปิดค้างนานเกินกำหนด (ต้องให้ LP กดปิด)
  if (action === 'stale') {
    try { return jsonOk({ stale: getStaleTrips() }); }
    catch (err) { return jsonErr('stale error: ' + err.message); }
  }

  // API: งานของ porter คนเดียว — ?action=myjobs&sid=...&name=...
  if (action === 'myjobs') {
    try {
      const p = (e && e.parameter) || {};
      return jsonOk(getMyJobs(p.sid || '', p.name || ''));
    } catch (err) { return jsonErr('myjobs error: ' + err.message); }
  }

  // API: health check
  if (action === 'health') {
    return jsonOk({ app: 'WC Tracker HKT', version: '5.0', time: new Date().toISOString() });
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

    if      (action === 'start')     return jsonOk(handleStart(trip));
    else if (action === 'end')       return jsonOk(handleEnd(trip));
    else if (action === 'assign')    return jsonOk(assignCase(data.assign || {}));
    else if (action === 'case_new')  return jsonOk(createCase(data.case || {}));
    else if (action === 'close_stale') return jsonOk({ closed: closeStaleTrips(data.by || '') });
    else if (action === 'queue_set') return jsonOk({ queue: setQueue(data.queue || []) });
    else if (action === 'wc_set')     return jsonOk({ repair: setWCStatus(data.wc || {}) });
    else                              return jsonErr('Unknown action: ' + action);

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

    const porterRaw = String(row[PCOL.PORTER] || '').trim();
    const svc       = String(row[PCOL.SVC] || '').trim().toUpperCase();
    const eta       = fmtTime(row[PCOL.ETA]);
    const etd       = fmtTime(row[PCOL.ETD]);

    // ขาเข้า/ขาออก: ใช้ checkbox ก่อน — ไม่มีก็ดูว่ามี ETA (เข้า) หรือ ETD (ออก)
    let dir = row[PCOL.DEP_CHK] === true ? 'DEP'
            : (row[PCOL.ARR_CHK] === true ? 'ARR' : '');
    if (!dir) dir = etd ? 'DEP' : (eta ? 'ARR' : '');

    cases.push({
      row:       r + 1,                                    // แถวจริงใน Sheet (1-based)
      no:        noNum,
      airline:   iata,
      fltno:     flt,
      flight:    (iata === 'ETC') ? 'ETC' : (iata + flt.split('/')[0]).replace(/\s/g, ''),
      porter:    porterRaw,
      porters:   splitPorters(porterRaw),                  // 1 เคสมีได้หลายคน ("A, B")
      state:     caseState(row[PCOL.STATUS], row[PCOL.DELIVER_TIME], row[PCOL.PICKUP_TIME]),
      status:    String(row[PCOL.STATUS] || '').trim(),
      assigned:  !!porterRaw,
      eta:       eta,
      etd:       etd,
      gate:      String(row[PCOL.GATE] || '').trim(),
      svc:       svc,                                      // WCHR/WCHS/WCHC/MAAS/AVIH/ETC
      dir:       dir,
      notified:  fmtTime(row[PCOL.NOTIFIED]),
      pickup:    fmtTime(row[PCOL.PICKUP_TIME]),
      delivered: fmtTime(row[PCOL.DELIVER_TIME]),
      remark:    String(row[PCOL.REMARK] || '').trim(),
    });
  }

  Logger.log('Cases found: ' + cases.length + ' | tab: ' + sheet.getName());
  return { cases: cases, staff: readStaffRecord(sheet), tab: sheet.getName() };
}

// ── ช่อง porter อาจมีหลายคน: "Chatchai, Maleekee" → ['Chatchai','Maleekee'] ──
function splitPorters(raw) {
  return String(raw || '').split(/[,/]+/)
    .map(function (x) { return x.trim(); })
    .filter(function (x) { return x.length > 0; });
}

// ── สถานะเคส: STANDBY (ยังไม่เริ่ม) / ON PROCESS (กำลังทำ) / COMPLETED (จบ) ──
function caseState(statusCell, deliverCell, pickupCell) {
  const st = String(statusCell || '').trim().toUpperCase();
  if (st.indexOf('COMPLET') >= 0) return 'COMPLETED';
  if (st.indexOf('PROCESS') >= 0) return 'ON PROCESS';
  if (st.indexOf('CANCEL')  >= 0) return 'CANCELLED';
  // ไม่มีค่าในช่องสถานะ → เดาจากเวลาที่กรอก
  if (deliverCell) return 'COMPLETED';
  if (pickupCell)  return 'ON PROCESS';
  return 'STANDBY';
}

// ═══════════════════════════════════════════════════════════════════
// สรุปเคสประจำวัน — ให้ตัวเลขชุดเดียวกับบล็อกสรุปในชีท
// ═══════════════════════════════════════════════════════════════════
function getTodaySummary() {
  const res   = getTodayCases();
  const cases = res.cases || [];
  const staff = res.staff || [];

  const sum = {
    tab:        res.tab,
    total:      cases.length,
    arr:        0, dep: 0,
    standby:    0, process: 0, completed: 0,
    unassigned: 0,
    bySvc:      {},   // WCHR / WCHS / WCHC / MAAS / AVIH / ETC
    byAirline:  {},   // 6E: 7, AI: 4, ...
    byStaff:    {},   // ชื่อในชีท → จำนวนเคส
    onDuty:     staff.filter(function (s) { return s.onDuty; }).length,
    dayOff:     staff.filter(function (s) { return !s.onDuty; }).length
  };

  cases.forEach(function (c) {
    if (c.dir === 'ARR') sum.arr++;
    else if (c.dir === 'DEP') sum.dep++;

    if      (c.state === 'COMPLETED')  sum.completed++;
    else if (c.state === 'ON PROCESS') sum.process++;
    else if (c.state === 'STANDBY')    sum.standby++;

    if (!c.assigned) sum.unassigned++;

    const svc = c.svc || '-';
    sum.bySvc[svc] = (sum.bySvc[svc] || 0) + 1;

    const air = c.airline || '-';
    sum.byAirline[air] = (sum.byAirline[air] || 0) + 1;

    c.porters.forEach(function (p) {
      sum.byStaff[p] = (sum.byStaff[p] || 0) + 1;
    });
  });

  // เทียบกับยอดจองล่วงหน้า + กำลังคนต่อช่วงเวลา
  try {
    const pre = getPreBooking();
    sum.preBooked = pre.total;
    sum.demand    = pre.demand;
    sum.manpower  = pre.manpower;
    sum.window    = currentWindow(pre);
    sum.rosterOn  = pre.roster.filter(function (r) { return r.onDuty; }).length;
    sum.rosterOff = pre.roster.filter(function (r) { return r.status !== 'DUTY'; }).length;
    sum.rosterAll = pre.roster.length;
  } catch (e) {
    Logger.log('summary/pre error: ' + e.message);
  }

  Logger.log('Summary: ' + sum.total + ' cases (ARR ' + sum.arr + ' / DEP ' + sum.dep
    + ') · standby ' + sum.standby + ' · on process ' + sum.process
    + ' · completed ' + sum.completed);
  return sum;
}

// ── รายชื่อพนักงานที่มาทำงานวันนี้ (บล็อก STAFF RECORD) ──
function getTodayStaff() {
  const today = new Date();
  const ss    = getPorterSpreadsheet(today);
  const sheet = findPorterDayTab(ss, today);
  if (!sheet) {
    Logger.log('Staff: day tab not found for ' + today.toDateString());
    return { staff: [], tab: null };
  }
  return { staff: readStaffRecord(sheet), tab: sheet.getName() };
}

// อ่านบล็อก STAFF RECORD จาก tab รายวัน
// หาหัวตาราง "NAME" (คู่กับ "NO."/"SKED") เองใน 12 แถวแรก — ถ้าไม่เจอใช้ SCOL
function readStaffRecord(sheet) {
  const data = sheet.getDataRange().getValues();
  let colName = SCOL.NAME, colNo = SCOL.NO, colSked = SCOL.SKED, colCases = SCOL.CASES;
  let headRow = -1;

  for (let r = 0; r < Math.min(12, data.length); r++) {
    for (let c = 0; c < data[r].length; c++) {
      if (String(data[r][c] || '').trim().toUpperCase() === 'NAME') {
        colName  = c;
        colNo    = c - 2;
        colSked  = c - 1;
        colCases = c + 1;
        headRow  = r;
        break;
      }
    }
    if (headRow >= 0) break;
  }

  // ชีทมาร์กคนที่ "วันหยุด" ด้วยพื้นหลังสีแดง (** RED MARK = DAY OFF)
  let bg = [];
  try {
    bg = sheet.getRange(1, colName + 1, data.length, 1).getBackgrounds();
  } catch (e) {
    Logger.log('readStaffRecord: cannot read backgrounds — ' + e.message);
  }

  const staff = [];
  const seen  = {};
  for (let r = (headRow >= 0 ? headRow + 1 : 1); r < data.length; r++) {
    const row  = data[r];
    const name = String(row[colName] || '').trim();
    if (!name) continue;
    if (name.toUpperCase() === 'NAME') continue;
    if (name.indexOf('**') === 0) continue;        // บรรทัดหมายเหตุ
    if (seen[name]) continue;

    const noRaw = colNo >= 0 ? row[colNo] : '';
    const no    = (typeof noRaw === 'number') ? noRaw : parseInt(noRaw, 10);
    if (!no || isNaN(no)) continue;                // แถวรายชื่อจริงต้องมีเลข NO.

    const off = bg.length > r ? isDayOffColor(bg[r][0]) : false;

    seen[name] = true;
    staff.push({
      no:     no,
      name:   name,
      sked:   colSked >= 0 ? String(row[colSked] || '').trim() : '',
      cases:  Number(row[colCases]) || 0,
      onDuty: !off,
      dayOff: off
    });
  }

  staff.sort(function (a, b) { return a.no - b.no; });
  Logger.log('Staff record: ' + staff.length + ' (on duty '
    + staff.filter(function (s) { return s.onDuty; }).length + ')');
  return staff;
}

// พื้นหลังโทนแดง = วันหยุด (รองรับหลายเฉด เช่น #ff0000, #e06666, #f4cccc)
function isDayOffColor(hex) {
  const h = String(hex || '').trim().toLowerCase();
  if (!h || h === '#ffffff' || h === 'white' || h.length < 7) return false;
  const r = parseInt(h.substr(1, 2), 16);
  const g = parseInt(h.substr(3, 2), 16);
  const b = parseInt(h.substr(5, 2), 16);
  if (isNaN(r) || isNaN(g) || isNaN(b)) return false;
  return r > 180 && (r - g) > 35 && (r - b) > 35;
}

// ═══════════════════════════════════════════════════════════════════
// PRE-WHEELCHAIR — ยอดจองล่วงหน้า + ตารางเวร porter + demand/manpower
// ═══════════════════════════════════════════════════════════════════
// ไฟล์: "<MMM> <YYYY> PRE-WHEELCHAIR" (โฟลเดอร์รายเดือนใน PRE_ROOT_FOLDER_ID)
// tab รายวัน: 01SEP26 … 30SEP26 (รูปแบบเดียวกับ Porter Summary)

function getPreSpreadsheet(date) {
  if (PRE_AUTO_BY_NAME) {
    try {
      const months = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
      const key    = months[date.getMonth()] + ' ' + date.getFullYear() + ' PRE-WHEELCHAIR';
      // ชื่อไฟล์จริงมีช่องว่างท้ายชื่อ — ใช้ contains แทน getFilesByName
      const it = DriveApp.searchFiles('title contains "' + key + '"');
      if (it.hasNext()) return SpreadsheetApp.open(it.next());
      Logger.log('PRE file not found by name: "' + key + '" — fallback to PRE_SS_ID');
    } catch (e) {
      Logger.log('PRE file search error: ' + e.message + ' — fallback to PRE_SS_ID');
    }
  }
  return SpreadsheetApp.openById(PRE_SS_ID);
}

// tab รายวันใช้รูปแบบเดียวกับ Porter Summary (01SEP26) → ใช้ resolver เดิม
function findPreDayTab(ss, date) {
  return findPorterDayTab(ss, date);
}

// ── อ่าน tab รายวันของ PRE-WHEELCHAIR ทั้ง 3 บล็อกในรอบเดียว ──
function getPreBooking() {
  const today = new Date();
  let ss;
  try {
    ss = getPreSpreadsheet(today);
  } catch (e) {
    throw new Error('Cannot open PRE-WHEELCHAIR. Check PRE_SS_ID. ' + e.message);
  }

  const sheet = findPreDayTab(ss, today);
  if (!sheet) {
    Logger.log('PRE day tab not found for ' + today.toDateString());
    return { tab: null, flights: [], demand: [], manpower: [], roster: [], total: 0 };
  }

  const data = sheet.getDataRange().getValues();

  const flights  = [];
  const demand   = [];
  const manpower = [];
  const roster   = [];
  let inManpower = false;   // บล็อก B มีสองตารางซ้อน: demand ก่อน แล้ว MANPOWER

  for (let r = 0; r < data.length; r++) {
    const row = data[r];

    // ── บล็อก A: ไฟลท์ + ยอดจอง ──
    const air = String(row[FCOL.AIRLINE] || '').trim();
    const flt = String(row[FCOL.FLTNO]   || '').trim();
    if (air && flt && air.toUpperCase() !== 'AIRLINES' && !/^\*/.test(air)) {
      const arr = {
        wchr: num(row[FCOL.A_WCHR]), wchs: num(row[FCOL.A_WCHS]), wchc: num(row[FCOL.A_WCHC]),
        avih: num(row[FCOL.A_AVIH]), maas: num(row[FCOL.A_MAAS]), total: num(row[FCOL.A_TOTAL])
      };
      const dep = {
        wchr: num(row[FCOL.D_WCHR]), wchs: num(row[FCOL.D_WCHS]), wchc: num(row[FCOL.D_WCHC]),
        avih: num(row[FCOL.D_AVIH]), maas: num(row[FCOL.D_MAAS]), total: num(row[FCOL.D_TOTAL])
      };
      if (arr.total || dep.total) {          // เก็บเฉพาะไฟลท์ที่มียอดจอง
        flights.push({
          airline: air,
          fltno:   flt,
          routing: String(row[FCOL.ROUTING] || '').trim(),
          sta:     fmtTime(row[FCOL.STA]),
          std:     fmtTime(row[FCOL.STD]),
          ctOpen:  fmtTime(row[FCOL.CT_OPEN]),
          ctClose: fmtTime(row[FCOL.CT_CLOSE]),
          arr:     arr,
          dep:     dep,
          total:   arr.total + dep.total
        });
      }
    }

    // ── บล็อก B: ช่วงเวลา (demand ก่อน → MANPOWER) ──
    const lbl = String(row[FCOL.WIN_LBL] || '').trim();
    if (lbl.toUpperCase() === 'MANPOWER') inManpower = true;
    if (/^\d{1,2}:\d{2}\s*-\s*\d{1,2}:\d{2}$/.test(lbl)) {
      const a = num(row[FCOL.WIN_1]);
      const b = num(row[FCOL.WIN_2]);
      const bucket = inManpower ? manpower : demand;
      // แถวช่วงเวลาถูก merge สองบรรทัด — กันซ้ำ
      if (!bucket.length || bucket[bucket.length - 1].win !== lbl) {
        bucket.push(inManpower ? { win: lbl, sked: a, ot: b }
                               : { win: lbl, arr: a, dep: b, total: a + b });
      }
    }

    // ── บล็อก C: Porter Daily Assignment (ตารางเวร) ──
    const nm = String(row[FCOL.R_NAME] || '').trim();
    if (nm && nm.toUpperCase() !== 'NAME' && nm.indexOf('Porter Daily') < 0) {
      const inV  = shiftVal(row[FCOL.R_IN]);
      const outV = shiftVal(row[FCOL.R_OUT]);
      if (inV) {
        const off = OFF_CODES.indexOf(inV.toUpperCase()) >= 0;
        roster.push({
          name:    nm,
          in:      off ? '' : inV,
          out:     off ? '' : outV,
          otIn:    shiftVal(row[FCOL.R_OT_IN]),
          otOut:   shiftVal(row[FCOL.R_OT_OUT]),
          status:  off ? inV.toUpperCase() : 'DUTY',
          shift:   off ? inV.toUpperCase() : (inV + '-' + outV),
          onDuty:  off ? false : inShift(inV, outV, today),
          onOT:    off ? false : inShift(shiftVal(row[FCOL.R_OT_IN]), shiftVal(row[FCOL.R_OT_OUT]), today)
        });
      }
    }
  }

  const total = demand.reduce(function (a, d) { return a + d.total; }, 0);
  Logger.log('PRE ' + sheet.getName() + ': ' + flights.length + ' flights, '
    + total + ' pre-booked, roster ' + roster.length
    + ' (on duty now ' + roster.filter(function (x) { return x.onDuty; }).length + ')');

  return { tab: sheet.getName(), flights: flights, demand: demand,
           manpower: manpower, roster: roster, total: total };
}

// ── helpers ──
function num(v) {
  const n = (typeof v === 'number') ? v : parseFloat(String(v || '').replace(/,/g, ''));
  return isNaN(n) ? 0 : n;
}

// ช่องเวลาในตารางเวรเป็นได้ทั้ง Date, "07:00" หรือรหัส OFF/SL/BL (#N/A = ไม่มี)
function shiftVal(v) {
  if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v, 'Asia/Bangkok', 'HH:mm');
  const s = String(v == null ? '' : v).trim();
  if (!s || s.indexOf('#') === 0) return '';
  return s;
}

// เวลาปัจจุบัน (Bangkok) อยู่ในช่วงกะหรือไม่ — รองรับกะข้ามวัน (17:00-05:00)
function inShift(inStr, outStr, now) {
  const a = hhmmToMin(inStr), b = hhmmToMin(outStr);
  if (a == null || b == null) return false;
  const t = hhmmToMin(Utilities.formatDate(now || new Date(), 'Asia/Bangkok', 'HH:mm'));
  if (a === b) return false;
  return (a < b) ? (t >= a && t < b) : (t >= a || t < b);   // ข้ามเที่ยงคืน
}

function hhmmToMin(s) {
  const m = String(s || '').match(/^(\d{1,2}):(\d{2})/);
  if (!m) return null;
  return parseInt(m[1], 10) * 60 + parseInt(m[2], 10);
}

// ── ช่วงเวลาปัจจุบัน + demand/manpower ของช่วงนั้น ──
function currentWindow(pre) {
  const now = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'HH:mm');
  const t   = hhmmToMin(now);
  function hit(list) {
    for (let i = 0; i < list.length; i++) {
      const m = String(list[i].win).match(/^(\d{1,2}:\d{2})\s*-\s*(\d{1,2}:\d{2})$/);
      if (!m) continue;
      const a = hhmmToMin(m[1]), b = hhmmToMin(m[2]);
      if (a < b ? (t >= a && t < b) : (t >= a || t < b)) return list[i];
    }
    return null;
  }
  return { now: now, demand: hit(pre.demand || []), manpower: hit(pre.manpower || []) };
}

// ═══════════════════════════════════════════════════════════════════
// BOARD — รวมทุกอย่างที่หน้าแอปต้องใช้ ไว้ใน request เดียว
// ═══════════════════════════════════════════════════════════════════
function getBoard() {
  const out = { time: Utilities.formatDate(new Date(), 'Asia/Bangkok', 'HH:mm'), errors: [] };

  // แต่ละแหล่งพังได้อิสระ — ที่เหลือยังใช้งานต่อได้
  try {
    const c = getTodayCases();
    out.cases = c.cases; out.staff = c.staff; out.caseTab = c.tab;
  } catch (e) { out.cases = []; out.staff = []; out.errors.push('cases: ' + e.message); }

  try {
    const pre = getPreBooking();
    out.pre      = { tab: pre.tab, flights: pre.flights, total: pre.total };
    out.demand   = pre.demand;
    out.manpower = pre.manpower;
    out.roster   = pre.roster;
    out.window   = currentWindow(pre);
  } catch (e) {
    out.pre = { flights: [], total: 0 }; out.demand = []; out.manpower = [];
    out.roster = []; out.errors.push('pre: ' + e.message);
  }

  try {
    const t = scanOpenTrips();
    out.active = t.active; out.stale = t.stale;
  } catch (e) { out.active = []; out.stale = []; out.errors.push('active: ' + e.message); }
  try { out.queue  = getQueue(); }    catch (e) { out.queue = [];  out.errors.push('queue: ' + e.message); }
  try { out.repair = getWCRepair(); } catch (e) { out.repair = []; out.errors.push('repair: ' + e.message); }

  if (out.errors.length) Logger.log('getBoard errors: ' + out.errors.join(' | '));
  return out;
}

// ═══════════════════════════════════════════════════════════════════
// WC STATUS — แจ้งซ่อม / ซ่อมเสร็จ (sheet "WCStatus")
// ═══════════════════════════════════════════════════════════════════
// - แจ้งซ่อม  : รถคันนั้นจะถูกกันออกจากการเลือกใช้งานในแอปทุกเครื่อง
// - ซ่อมเสร็จ : กลับมาใช้งานได้ (เก็บแถวไว้เป็นประวัติ)

function getWCSheet() {
  return getOrCreateSheet(SpreadsheetApp.getActiveSpreadsheet(), WC_SHEET, WC_HEADERS);
}

// คืนเฉพาะคันที่ "ซ่อม" อยู่ตอนนี้
function getWCRepair() {
  const sh   = getWCSheet();
  const data = sh.getDataRange().getValues();
  const list = [];
  for (let r = 1; r < data.length; r++) {
    if (String(data[r][1] || '').trim() !== 'ซ่อม') continue;
    list.push({
      ctrl:  String(data[r][0] || '').trim(),
      note:  String(data[r][2] || '').trim(),
      by:    String(data[r][3] || '').trim(),
      since: String(data[r][4] || '').trim()
    });
  }
  Logger.log('WC under repair: ' + list.length);
  return list;
}

// wc = { ctrl, status: 'ซ่อม' | 'ใช้งานได้', note, by }
function setWCStatus(wc) {
  if (!wc || !wc.ctrl) throw new Error('No wc data');

  const sh     = getWCSheet();
  const now    = new Date();
  const stamp  = Utilities.formatDate(now, 'Asia/Bangkok', 'dd/MM/yyyy HH:mm:ss');
  const status = (wc.status === 'ซ่อม') ? 'ซ่อม' : 'ใช้งานได้';
  const data   = sh.getDataRange().getValues();

  let row = -1;
  for (let r = 1; r < data.length; r++) {
    if (String(data[r][0]).trim() === String(wc.ctrl).trim()) { row = r + 1; break; }
  }

  if (row > 0) {
    sh.getRange(row, 2).setValue(status);
    sh.getRange(row, 3).setValue(wc.note || '');
    sh.getRange(row, 4).setValue(wc.by || '');
    if (status === 'ซ่อม') sh.getRange(row, 5).setValue(stamp);   // แจ้งเมื่อ
    sh.getRange(row, 6).setValue(stamp);
    sh.getRange(row, 1, 1, WC_HEADERS.length)
      .setBackground(status === 'ซ่อม' ? '#fee2e2' : '#dcfce7');
  } else {
    sh.appendRow([wc.ctrl, status, wc.note || '', wc.by || '', stamp, stamp]);
    sh.getRange(sh.getLastRow(), 1, 1, WC_HEADERS.length)
      .setBackground(status === 'ซ่อม' ? '#fee2e2' : '#dcfce7');
  }

  Logger.log('WC ' + wc.ctrl + ' → ' + status + (wc.note ? ' (' + wc.note + ')' : ''));
  return getWCRepair();
}

// เช็คว่ารถคันนี้แจ้งซ่อมอยู่ไหม
function isWCUnderRepair(ctrl) {
  if (!ctrl) return false;
  return getWCRepair().some(function (w) { return w.ctrl === String(ctrl).trim(); });
}

// ═══════════════════════════════════════════════════════════════════
// PORTER QUEUE — คิวรับเคสอัตโนมัติ (เก็บใน sheet "Queue" — ทุกเครื่องเห็นตรงกัน)
// ═══════════════════════════════════════════════════════════════════
// การทำงาน:
//  - Supervisor จัดคิวคนเข้าเวรจากหน้า "คิว" ในแอป (เลือกกะ A/B/C) → queue_set
//  - ลำดับคิว = จำนวนเคสวันนี้น้อยสุดขึ้นก่อน (เสมอกัน → คนที่ว่างก่อนขึ้นก่อน)
//  - เริ่มบริการ (start)  → คนนั้นสถานะ "กำลังบริการ" + นับ trip
//  - จบบริการ (end)       → กลับเป็น "ว่าง" (ลำดับตัดสินเสมอถูกดันไปท้าย)
//  - คิวเป็นรายวัน — ข้ามวันแล้วเริ่มคิวใหม่

function queueTodayStr() {
  return Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy');
}

function getQueueSheet(ss) {
  return getOrCreateSheet(ss || SpreadsheetApp.getActiveSpreadsheet(), QUEUE_SHEET, QUEUE_HEADERS);
}

// ── อ่านคิวของวันนี้ (เรียงตามลำดับ) ──
function getQueue() {
  const sh    = getQueueSheet();
  const today = queueTodayStr();
  const data  = sh.getDataRange().getValues();
  const list  = [];
  for (let r = 1; r < data.length; r++) {
    if (String(data[r][0]) !== today) continue;
    list.push({
      order:  Number(data[r][1]) || 0,
      sid:    String(data[r][2] || ''),
      nick:   String(data[r][3] || ''),
      shift:  String(data[r][4] || 'A'),
      status: String(data[r][5] || 'ว่าง'),
      trips:  Number(data[r][6]) || 0,
    });
  }
  // เคสน้อยสุดขึ้นก่อน → เสมอกันใช้ลำดับ (คนที่ว่างก่อน/เข้าคิวก่อน)
  list.sort((a, b) => (a.trips - b.trips) || (a.order - b.order));
  return list;
}

// ── บันทึกคิวใหม่ทั้งชุด (จากหน้า "คิว" ในแอป) ──
// items = [{sid, nick, shift}, ...]
// สถานะ/จำนวน trip/ลำดับตัดสินเสมอ ของคนที่อยู่ในคิวเดิมจะถูกเก็บไว้
function setQueue(items) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sh    = getQueueSheet(ss);
  const today = queueTodayStr();
  const now   = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'HH:mm:ss');

  // เก็บสถานะเดิมไว้ก่อนล้าง
  const prev = {};
  getQueue().forEach(q => { prev[q.sid] = q; });

  // ลบแถวของวันนี้ (และวันเก่า — เก็บชีทให้สะอาด)
  const data = sh.getDataRange().getValues();
  for (let r = data.length - 1; r >= 1; r--) {
    sh.deleteRow(r + 1);
  }

  const rows = (items || []).map((it, i) => {
    const old = prev[String(it.sid)] || {};
    return [
      today,
      old.order || (i + 1),
      String(it.sid || ''),
      String(it.nick || ''),
      String(it.shift || old.shift || 'A'),
      old.status || 'ว่าง',
      old.trips  || 0,
      now
    ];
  });
  if (rows.length) sh.getRange(2, 1, rows.length, QUEUE_HEADERS.length).setValues(rows);

  Logger.log('setQueue: ' + rows.length + ' คน');
  return getQueue();
}

// ── hook ตอน start/end trip ──
function queueOnStart(sid) {
  try {
    if (!sid) return;
    const sh    = getQueueSheet();
    const today = queueTodayStr();
    const data  = sh.getDataRange().getValues();
    for (let r = 1; r < data.length; r++) {
      if (String(data[r][0]) === today && String(data[r][2]) === String(sid)) {
        sh.getRange(r + 1, 6).setValue('กำลังบริการ');
        sh.getRange(r + 1, 7).setValue((Number(data[r][6]) || 0) + 1);
        sh.getRange(r + 1, 8).setValue(Utilities.formatDate(new Date(), 'Asia/Bangkok', 'HH:mm:ss'));
        return;
      }
    }
  } catch (e) { Logger.log('queueOnStart error: ' + e.message); }
}

function queueOnEnd(sid) {
  try {
    if (!sid) return;
    const sh    = getQueueSheet();
    const today = queueTodayStr();
    const data  = sh.getDataRange().getValues();
    let maxOrder = 0;
    for (let r = 1; r < data.length; r++) {
      if (String(data[r][0]) === today) maxOrder = Math.max(maxOrder, Number(data[r][1]) || 0);
    }
    for (let r = 1; r < data.length; r++) {
      if (String(data[r][0]) === today && String(data[r][2]) === String(sid)) {
        sh.getRange(r + 1, 2).setValue(maxOrder + 1);   // ตัวตัดสินเสมอ: ไปท้ายสุด
        sh.getRange(r + 1, 6).setValue('ว่าง');
        sh.getRange(r + 1, 8).setValue(Utilities.formatDate(new Date(), 'Asia/Bangkok', 'HH:mm:ss'));
        return;
      }
    }
  } catch (e) { Logger.log('queueOnEnd error: ' + e.message); }
}

// ═══════════════════════════════════════════════════════════════════
// ACTIVE TRIPS — งานที่ยังไม่จบ (อ่านจาก TripLog)
// ═══════════════════════════════════════════════════════════════════
// เก็บฝั่ง server เพื่อให้ porter เริ่มงานบนเครื่องหนึ่งแล้วไปกดจบบนอีกเครื่องได้
// และรีเฟรชหน้าจอแล้วงานที่ค้างอยู่ไม่หาย
const TCOL = {
  ID: 0, DATE: 1, CTRL: 2, TYPE: 3,
  SID: 4, SNICK: 5, SNAME: 6,
  FLIGHT: 7, AIRLINE: 8, ROUTING: 9, STA: 10, STD: 11,
  START: 12, END: 13, DUR: 14, STATUS: 15, SAVED: 16,
  CASE_NO: 17, CASE_SVC: 18, CASE_PORTER: 19,
  CASE_ROW: 20, CASE_AIRLINE: 21, CASE_FLTNO: 22,
};

function getActiveTrips() { return scanOpenTrips().active; }

// งานที่เปิดค้างนานเกิน STALE_HOURS — ไม่นับเป็นงานที่กำลังทำ
// ถ้าปล่อยไว้จะล็อกรถเข็นคันนั้นไม่ให้ใครใช้ และโชว์ตัวจับเวลาเดินเป็นร้อยชั่วโมง
function getStaleTrips() { return scanOpenTrips().stale; }

// อ่าน TripLog รอบเดียว แล้วแยกเป็น "กำลังทำ" กับ "ค้าง"
function scanOpenTrips() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return { active: [], stale: [] };
  const log = ss.getSheetByName(LOG_SHEET);
  if (!log || log.getLastRow() < 2) return { active: [], stale: [] };

  const data   = log.getDataRange().getValues();
  const active = [], stale = [];
  const cutoff = Date.now() - STALE_HOURS * 3600 * 1000;

  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    if (!row[TCOL.ID]) continue;
    if (row[TCOL.END]) continue;                                   // จบไปแล้ว
    if (String(row[TCOL.STATUS] || '').indexOf('กำลัง') < 0) continue;

    const startMs = tripStartMs(row[TCOL.DATE], row[TCOL.START]);
    const bucket  = (startMs !== null && startMs < cutoff) ? stale : active;

    bucket.push({
      id:          String(row[TCOL.ID]),
      date:        fmtDateCell(row[TCOL.DATE]),
      ctrl:        String(row[TCOL.CTRL]  || ''),
      type:        String(row[TCOL.TYPE]  || ''),
      sid:         String(row[TCOL.SID]   || ''),
      snick:       String(row[TCOL.SNICK] || ''),
      sname:       String(row[TCOL.SNAME] || ''),
      flight:      String(row[TCOL.FLIGHT]  || ''),
      flAirline:   String(row[TCOL.AIRLINE] || ''),
      flRouting:   String(row[TCOL.ROUTING] || ''),
      flSTA:       String(row[TCOL.STA] || ''),
      flSTD:       String(row[TCOL.STD] || ''),
      startTime:   fmtClockCell(row[TCOL.START]),
      caseNo:      row[TCOL.CASE_NO]     || '',
      caseSvc:     String(row[TCOL.CASE_SVC]    || ''),
      casePorter:  String(row[TCOL.CASE_PORTER] || ''),
      caseRow:     row[TCOL.CASE_ROW]     || '',
      caseAirline: String(row[TCOL.CASE_AIRLINE] || ''),
      caseFltno:   String(row[TCOL.CASE_FLTNO]   || ''),
      logRow:      r + 1,
      ageHours:    startMs === null ? null : Math.round((Date.now() - startMs) / 360000) / 10,
    });
  }
  return { active: active, stale: stale };
}

// "20/09/2026" + "08:12:30" → epoch ms (เวลาไทย) · อ่านไม่ออกคืน null
// อ่านไม่ออก = ไม่รู้ว่าเก่าแค่ไหน จึงถือว่ายังใช้งานอยู่ ดีกว่าไปปิดงานจริงของใคร
function tripStartMs(dateVal, timeVal) {
  const d = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(fmtDateCell(dateVal));
  const t = /^(\d{1,2}):(\d{2})(?::(\d{2}))?$/.exec(fmtClockCell(timeVal));
  if (!d || !t) return null;
  const ms = new Date(Number(d[3]), Number(d[2]) - 1, Number(d[1]),
                      Number(t[1]), Number(t[2]), Number(t[3] || 0)).getTime();
  return isNaN(ms) ? null : ms;
}

// ── ปิดงานที่ค้าง ──
// ไม่เดาเวลาจบ (ไม่มีใครรู้ว่าจบจริงตอนไหน) — ทำแค่ปลดล็อกรถเข็น
// และกันไม่ให้แถวนั้นโผล่เป็น "กำลังให้บริการ" อีก ประวัติยังอยู่ครบ
function closeStaleTrips(by) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('ไม่พบ Spreadsheet ที่ผูกกับสคริปต์นี้');
  const log = ss.getSheetByName(LOG_SHEET);
  if (!log) return [];

  const stale = getStaleTrips();
  const now   = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm:ss');
  const who   = String(by || '').trim();

  stale.forEach(function (t) {
    log.getRange(t.logRow, TCOL.STATUS + 1)
       .setValue('ค้าง — ไม่ได้กดจบ' + (who ? ' (ปิดโดย ' + who + ')' : ''));
    log.getRange(t.logRow, TCOL.SAVED + 1).setValue(now);
    Logger.log('ปิดงานค้าง: ' + t.id + ' | ' + t.snick + ' | ' + t.ctrl
      + ' | เริ่ม ' + t.date + ' ' + t.startTime + ' (' + t.ageHours + ' ชม.)');
  });

  Logger.log('ปิดงานค้างทั้งหมด ' + stale.length + ' รายการ');
  return stale;
}

// เซลล์วันที่ใน TripLog เขียนเป็นข้อความ dd/MM/yyyy แต่ชีทอาจแปลงเป็น Date เอง
function fmtDateCell(v) {
  if (v instanceof Date && !isNaN(v)) {
    return Utilities.formatDate(v, 'Asia/Bangkok', 'dd/MM/yyyy');
  }
  return String(v || '').trim();
}

// เวลาเริ่ม/จบ เขียนเป็น HH:mm:ss แต่ชีทอาจแปลงเป็น Date เช่นกัน
function fmtClockCell(v) {
  if (v instanceof Date && !isNaN(v)) {
    return Utilities.formatDate(v, 'Asia/Bangkok', 'HH:mm:ss');
  }
  return String(v || '').trim();
}

// งานของ porter คนหนึ่ง: เคสที่ LP จ่ายให้ + งานที่กำลังทำอยู่
// name = ชื่ออังกฤษอย่างที่สะกดในชีท (แอปส่งมาจากตาราง ROMAN)
function getMyJobs(sid, name) {
  const res  = getTodayCases();
  const want = normPorterName(name);

  const mine = (res.cases || []).filter(function (c) {
    if (!want) return false;
    return (c.porters || []).some(function (p) { return porterMatches(p, want); });
  });

  const active = getActiveTrips().filter(function (t) {
    return !sid || String(t.sid) === String(sid);
  });

  return { name: name || '', sid: sid || '', tab: res.tab, jobs: mine, active: active };
}

function normPorterName(x) {
  return String(x || '').toLowerCase().replace(/[^a-z0-9฀-๿]/g, '');
}

// ชีทสะกดชื่อไม่นิ่ง ("Ameen" / "A Meen" / "Wuttichai K.") — เทียบแบบตัดอักขระพิเศษ
// แล้วยอมให้เป็นคำขึ้นต้นของกันและกันได้ เหมือนที่ฝั่งแอปทำ
function porterMatches(sheetName, wantNorm) {
  const n = normPorterName(sheetName);
  if (!n || !wantNorm) return false;
  return n === wantNorm || n.indexOf(wantNorm) === 0 || wantNorm.indexOf(n) === 0;
}

// ═══════════════════════════════════════════════════════════════════
// PORTER SUMMARY — write-back: หาแถวเคส / อัปเดตเวลา / จ่ายงาน
// ═══════════════════════════════════════════════════════════════════

// เปิด tab รายวันของ Porter Summary (ใช้ร่วมกันทุกฟังก์ชันที่เขียนกลับ)
function openPorterDayTab() {
  const today = new Date();
  const ss    = getPorterSpreadsheet(today);
  const sheet = findPorterDayTab(ss, today);
  if (!sheet) {
    throw new Error('ไม่พบ tab ของวันนี้ใน Porter Summary — '
      + 'รัน listPorterTabs() แล้วตั้ง PORTER_TAB_OVERRIDE ถ้าจำเป็น');
  }
  return sheet;
}

// ── หาแถวของเคสในชีท ──
// แอปส่ง caseRow ที่อ่านไว้ตอนโหลดมาด้วย แต่ OCC อาจแทรก/ลบแถวระหว่างนั้น
// จึงตรวจว่าแถวเดิมยังเป็นเคสเดียวกันไหม ถ้าไม่ใช่ค่อยไล่หาใหม่จาก IATA + เที่ยวบิน
// คืน -1 ถ้าหาไม่เจอ
function locateCaseRow(sheet, ref) {
  const wantAir = String((ref && ref.caseAirline) || '').trim().toUpperCase();
  const wantFlt = String((ref && ref.caseFltno)   || '').trim().toUpperCase();
  if (!wantAir || !wantFlt) return -1;

  const maxRow = sheet.getLastRow();
  const wantNo = parseInt((ref && ref.caseNo), 10);
  const row    = parseInt((ref && ref.caseRow), 10);

  const sameCase = function (vals) {
    return String(vals[PCOL.AIRLINE] || '').trim().toUpperCase() === wantAir
        && String(vals[PCOL.FLTNO]   || '').trim().toUpperCase() === wantFlt;
  };

  if (row >= 1 && row <= maxRow) {
    const vals = sheet.getRange(row, 1, 1, PCOL.SVC + 1).getValues()[0];
    if (sameCase(vals)) return row;
  }

  // แถวเดิมไม่ใช่เคสนี้แล้ว → ไล่หาใหม่ทั้งชีท
  const data = sheet.getDataRange().getValues();
  let fallback = -1;
  for (let r = 0; r < data.length; r++) {
    if (!sameCase(data[r])) continue;
    // เที่ยวบินเดียวกันมีได้หลายเคส — ใช้เลข "ที่" ช่วยแยกถ้ามี
    const no = parseInt(data[r][PCOL.NO], 10);
    if (wantNo && no === wantNo) return r + 1;
    if (fallback < 0) fallback = r + 1;
  }
  return fallback;
}

// ── เขียนเวลา "รับเคส" / "ส่งเคส" กลับเข้า Porter Summary ──
// เขียนทับของเดิมไม่ได้ (กันแอปไปลบสิ่งที่ LP กรอกเอง) และพังแล้วต้องไม่ทำให้ trip ล้ม
function markCase(trip, phase, timeStr) {
  if (!PORTER_WRITEBACK) return;
  if (!trip || !trip.caseRow) return;

  try {
    const sheet = openPorterDayTab();
    const row   = locateCaseRow(sheet, trip);
    if (row < 0) {
      Logger.log('markCase: ไม่พบแถวเคส ' + trip.caseAirline + trip.caseFltno);
      return;
    }

    const t = timeStr || Utilities.formatDate(new Date(), 'Asia/Bangkok', 'HH:mm');

    if (phase === 'pickup') {
      if (!sheet.getRange(row, PCOL.PICKUP_TIME + 1).getValue()) {
        sheet.getRange(row, PCOL.PICKUP_MARK + 1).setValue(true);
        sheet.getRange(row, PCOL.PICKUP_TIME + 1).setValue(t);
        Logger.log('markCase pickup: แถว ' + row + ' = ' + t);
      }
      setCaseStatus(sheet, row, 'ON PROCESS');

    } else if (phase === 'deliver') {
      if (!sheet.getRange(row, PCOL.DELIVER_TIME + 1).getValue()) {
        sheet.getRange(row, PCOL.DELIVER_TIME + 1).setValue(t);
        Logger.log('markCase deliver: แถว ' + row + ' = ' + t);
      }
      setCaseStatus(sheet, row, 'COMPLETED');
    }
  } catch (e) {
    // write-back พังต้องไม่ทำให้การบันทึก trip ล้ม
    Logger.log('markCase error (' + phase + '): ' + e.message);
  }
}

// ── ช่องสถานะ: ไม่ถอยหลัง (COMPLETED แล้วไม่กลับไป ON PROCESS) และไม่แตะเคสที่ถูกยกเลิก ──
function setCaseStatus(sheet, row, status) {
  const cell = sheet.getRange(row, PCOL.STATUS + 1);
  const cur  = String(cell.getValue() || '').trim().toUpperCase();
  if (cur.indexOf('CANCEL') >= 0) return;
  if (cur.indexOf('COMPLET') >= 0 && status !== 'COMPLETED') return;
  if (cur === status) return;
  cell.setValue(status);
}

// ═══════════════════════════════════════════════════════════════════
// ASSIGN — LP จ่ายเคสให้ porter จากในแอป
// ═══════════════════════════════════════════════════════════════════
// a = { caseRow, caseNo, caseAirline, caseFltno, porters:[ชื่อในชีท], by, note }
// porters ว่าง = ถอนงานคืน (ล้างช่องชื่อพนักงาน)
function assignCase(a) {
  if (!a) throw new Error('ไม่มีข้อมูลการจ่ายงาน');

  const sheet = openPorterDayTab();
  const row   = locateCaseRow(sheet, a);
  if (row < 0) {
    throw new Error('ไม่พบเคส ' + (a.caseAirline || '') + (a.caseFltno || '')
      + ' ในชีทวันนี้ — อาจถูกลบหรือย้ายแถว ลองรีเฟรชแล้วจ่ายใหม่');
  }

  const cur = sheet.getRange(row, 1, 1, PCOL.SVC + 1).getValues()[0];

  // เคสที่ส่งผู้โดยสารแล้วไม่ควรเปลี่ยนคน — ชื่อในชีทคือคนที่ทำจริง
  if (cur[PCOL.DELIVER_TIME]) {
    throw new Error('เคส #' + (cur[PCOL.NO] || a.caseNo) + ' ส่งเรียบร้อยแล้ว — เปลี่ยนคนไม่ได้');
  }

  const before = String(cur[PCOL.PORTER] || '').trim();
  const names  = (a.porters || [])
    .map(function (x) { return String(x || '').trim(); })
    .filter(function (x) { return x.length > 0; });
  const after = names.join(', ');

  // เคสที่รับผู้โดยสารไปแล้ว เพิ่มคนช่วยได้ แต่ถอดคนที่กำลังทำอยู่ออกไม่ได้
  if (cur[PCOL.PICKUP_TIME] && before) {
    const keep = splitPorters(before).filter(function (n) {
      return names.map(function (x) { return x.toUpperCase(); }).indexOf(n.toUpperCase()) < 0;
    });
    if (keep.length) {
      throw new Error('เคส #' + (cur[PCOL.NO] || a.caseNo) + ' กำลังให้บริการอยู่โดย '
        + keep.join(', ') + ' — เอาชื่อออกไม่ได้ (เพิ่มคนช่วยได้)');
    }
  }

  sheet.getRange(row, PCOL.PORTER + 1).setValue(after);

  // ยังไม่เริ่มงาน: มีคนแล้ว = STANDBY, ถอนคนออก = ล้างสถานะ
  if (!cur[PCOL.PICKUP_TIME]) {
    const st = String(cur[PCOL.STATUS] || '').trim().toUpperCase();
    if (after && !st) sheet.getRange(row, PCOL.STATUS + 1).setValue('STANDBY');
    if (!after && st === 'STANDBY') sheet.getRange(row, PCOL.STATUS + 1).setValue('');
  }

  logAssign({
    row:     row,
    caseNo:  cur[PCOL.NO] || a.caseNo || '',
    flight:  String(cur[PCOL.AIRLINE] || '') + String(cur[PCOL.FLTNO] || ''),
    svc:     String(cur[PCOL.SVC] || ''),
    after:   after || '(ถอนงาน)',
    before:  before,
    by:      a.by || '',
    note:    a.note || ''
  });

  Logger.log('assignCase: แถว ' + row + ' #' + (cur[PCOL.NO] || '') + ' → "' + after + '" โดย ' + (a.by || '-'));
  return { row: row, caseNo: cur[PCOL.NO] || a.caseNo || '', porter: after, before: before };
}

// ── LP เปิดเคสใหม่จากไฟลท์ที่ยังไม่มีในชีท ──
// c = { airline, fltno, eta, etd, gate, svc, dir, porters, by, remark }
// ใช้แถวว่างที่มีเลข "ที่" อยู่แล้วเท่านั้น — ไม่แทรกแถวเอง เพราะจะทำให้
// บล็อก STAFF RECORD ทางขวาและสูตรในชีทเลื่อนตาม
function createCase(c) {
  if (!c) throw new Error('ไม่มีข้อมูลเคส');
  const airline = String(c.airline || '').trim().toUpperCase();
  const fltno   = String(c.fltno   || '').trim();
  if (!airline || !fltno) throw new Error('ต้องระบุสายการบินและเที่ยวบิน');

  const sheet = openPorterDayTab();
  const data  = sheet.getDataRange().getValues();

  let target = -1;
  for (let r = 0; r < data.length; r++) {
    const no = parseInt(data[r][PCOL.NO], 10);
    if (!no || isNaN(no)) continue;
    const air = String(data[r][PCOL.AIRLINE] || '').trim();
    const flt = String(data[r][PCOL.FLTNO]   || '').trim();
    if (!air && !flt) { target = r + 1; break; }     // แถวที่มีเลขแต่ยังไม่มีไฟลท์
  }
  if (target < 0) {
    throw new Error('ชีทวันนี้ไม่มีแถวว่างเหลือ — เพิ่มแถวในชีทก่อนแล้วค่อยกดใหม่');
  }

  const names = (c.porters || [])
    .map(function (x) { return String(x || '').trim(); })
    .filter(function (x) { return x.length > 0; })
    .join(', ');

  sheet.getRange(target, PCOL.AIRLINE + 1).setValue(airline);
  sheet.getRange(target, PCOL.FLTNO   + 1).setValue(fltno);
  if (names)    sheet.getRange(target, PCOL.PORTER + 1).setValue(names);
  if (c.eta)    sheet.getRange(target, PCOL.ETA    + 1).setValue(c.eta);
  if (c.etd)    sheet.getRange(target, PCOL.ETD    + 1).setValue(c.etd);
  if (c.gate)   sheet.getRange(target, PCOL.GATE   + 1).setValue(c.gate);
  if (c.svc)    sheet.getRange(target, PCOL.SVC    + 1).setValue(String(c.svc).toUpperCase());
  if (c.remark) sheet.getRange(target, PCOL.REMARK + 1).setValue(c.remark);
  if (c.dir === 'ARR') sheet.getRange(target, PCOL.ARR_CHK + 1).setValue(true);
  if (c.dir === 'DEP') sheet.getRange(target, PCOL.DEP_CHK + 1).setValue(true);
  if (names) sheet.getRange(target, PCOL.STATUS + 1).setValue('STANDBY');

  const caseNo = sheet.getRange(target, PCOL.NO + 1).getValue();

  logAssign({
    row: target, caseNo: caseNo, flight: airline + fltno, svc: c.svc || '',
    after: names || '(ยังไม่จ่าย)', before: '', by: c.by || '', note: 'เปิดเคสใหม่จากแอป'
  });

  Logger.log('createCase: แถว ' + target + ' #' + caseNo + ' ' + airline + fltno);
  return { row: target, caseNo: caseNo, flight: airline + fltno, porter: names };
}

// ── ประวัติการจ่ายงาน (เก็บในไฟล์ log ของแอป ไม่ยุ่งกับ Porter Summary) ──
function logAssign(rec) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return;
    const sh  = getOrCreateSheet(ss, ASSIGN_SHEET, ASSIGN_HEADERS);
    const now = new Date();
    sh.appendRow([
      Utilities.formatDate(now, 'Asia/Bangkok', 'dd/MM/yyyy HH:mm:ss'),
      Utilities.formatDate(now, 'Asia/Bangkok', 'dd/MM/yyyy'),
      rec.row || '', rec.caseNo || '', rec.flight || '', rec.svc || '',
      rec.after || '', rec.before || '', rec.by || '', rec.note || ''
    ]);
  } catch (e) {
    Logger.log('logAssign error: ' + e.message);
  }
}

// ═══════════════════════════════════════════════════════════════════
// handleStart
// ═══════════════════════════════════════════════════════════════════
function handleStart(trip) {
  if (!trip || !trip.id) throw new Error('No trip data');
  if (isWCUnderRepair(trip.ctrl)) throw new Error('รถเข็น ' + trip.ctrl + ' แจ้งซ่อมอยู่ — เลือกคันอื่น');

  // กันเริ่มซ้ำจากคนละเครื่อง
  const open = getActiveTrips();
  const dupWC = open.filter(function (t) { return t.ctrl && t.ctrl === trip.ctrl; })[0];
  if (dupWC) {
    throw new Error('รถเข็น ' + trip.ctrl + ' กำลังถูกใช้โดย ' + (dupWC.snick || dupWC.sid)
      + ' (เริ่ม ' + dupWC.startTime + ') — เลือกคันอื่น');
  }
  if (trip.caseNo) {
    const dupCase = open.filter(function (t) {
      return String(t.caseNo) === String(trip.caseNo) && String(t.sid) === String(trip.sid);
    })[0];
    if (dupCase) {
      throw new Error('เคส #' + trip.caseNo + ' เริ่มไปแล้วเมื่อ ' + dupCase.startTime
        + ' — ไปกดจบงานที่หน้า "งานของฉัน"');
    }
  }

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
    trip.casePorter || '',
    trip.caseRow     || '',
    trip.caseAirline || '',
    trip.caseFltno   || ''
  ]);

  // write-back "รับเคส" → Porter Summary
  markCase(trip, 'pickup', startTime.slice(0, 5));

  // อัปเดตคิว: คนนี้กำลังบริการ
  queueOnStart(trip.sid);

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
      trip.caseNo||'', trip.caseSvc||'', trip.casePorter||'',
      trip.caseRow||'', trip.caseAirline||'', trip.caseFltno||''
    ]);
  }

  // write-back "ส่งเคส" → Porter Summary
  markCase(trip, 'deliver', endTime.slice(0, 5));

  // อัปเดตคิว: กลับเป็นว่าง + ไปต่อท้ายคิว
  queueOnEnd(trip.sid);

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
  if (!ss) {
    throw new Error('ไม่พบ Spreadsheet ที่ผูกกับสคริปต์นี้ — '
      + 'ต้องสร้างสคริปต์จาก Google Sheets (Extensions → Apps Script) ไม่ใช่สคริปต์เดี่ยว');
  }

  const made = [
    [LOG_SHEET,   LOG_HEADERS],
    [STAT_SHEET,  ['วันที่','Trips ทั้งหมด','Trips เสร็จ','เฉลี่ย (น.)','นานสุด (น.)','สั้นสุด (น.)','> 60 น.','อัปเดต']],
    [QUEUE_SHEET,  QUEUE_HEADERS],
    [WC_SHEET,     WC_HEADERS],
    [ASSIGN_SHEET, ASSIGN_HEADERS]
  ].map(function (x) { getOrCreateSheet(ss, x[0], x[1]); return x[0]; });

  ss.rename('WC Tracker Log — HKT Ground Handling');

  const msg = '✅ ตั้งค่าเสร็จ\n\n'
    + 'Sheets พร้อมใช้: ' + made.join(' · ') + '\n\n'
    + 'ขั้นตอนต่อไป:\n'
    + '1. Run listPorterTabs() — เช็คว่าเจอ tab รายวันของ Porter Summary\n'
    + '2. Run testPre() — เช็คตารางเวร + ยอดจองล่วงหน้า\n'
    + '3. Run testCases() / testSummary() — เช็คเคสและสรุปวันนี้\n'
    + '4. Run testFreeCaseRows() — เหลือแถวว่างให้ LP เปิดเคสใหม่กี่แถว\n'
    + '5. Run testMyJobs() — เช็คการจับคู่ชื่อ porter กับเคสในชีท\n'
    + '   (แก้ชื่อในฟังก์ชันเป็นคนที่มีเคสจริงวันนี้ก่อนรัน)';

  Logger.log(msg);

  // getUi() ใช้ได้เฉพาะตอนรันจากเมนู/editor ของไฟล์ Sheets เท่านั้น —
  // รันจาก trigger, web app หรือ clasp จะ throw ปล่อยผ่านไปใช้ Logger แทน
  try {
    SpreadsheetApp.getUi().alert(msg);
  } catch (e) {
    Logger.log('(ไม่มี UI ในบริบทนี้ — ดูผลจาก log ด้านบนได้เลย)');
  }

  return made;
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

function testStaff() {
  Logger.log('=== TEST STAFF ON DUTY ===');
  try {
    const res = getTodayStaff();
    Logger.log('Tab: ' + res.tab + ' | Total: ' + res.staff.length);
    res.staff.forEach(function (s) { Logger.log(s.no + '. ' + s.name + ' (' + s.cases + ' เคส)'); });
  } catch (e) {
    Logger.log('ERROR: ' + e.message);
  }
}

function testPre() {
  Logger.log('=== TEST PRE-WHEELCHAIR ===');
  try {
    const p = getPreBooking();
    Logger.log('Tab ' + p.tab + ' | ไฟลท์ที่มียอดจอง ' + p.flights.length + ' | รวม ' + p.total + ' เคส');
    p.demand.forEach(function (d) { Logger.log('  demand ' + d.win + ' → ARR ' + d.arr + ' / DEP ' + d.dep); });
    p.manpower.forEach(function (m) { Logger.log('  manpower ' + m.win + ' → SKED ' + m.sked + ' / OT ' + m.ot); });
    Logger.log('Roster ' + p.roster.length + ' คน · เข้าเวรตอนนี้ '
      + p.roster.filter(function (r) { return r.onDuty; }).length);
    p.roster.slice(0, 10).forEach(function (r) {
      Logger.log('  ' + r.name + ' · ' + r.shift + (r.onDuty ? ' ← ตอนนี้' : '') );
    });
    const w = currentWindow(p);
    Logger.log('ช่วงนี้ ' + w.now + ' → demand ' + JSON.stringify(w.demand) + ' | manpower ' + JSON.stringify(w.manpower));
  } catch (e) {
    Logger.log('ERROR: ' + e.message);
  }
}

function testSummary() {
  Logger.log('=== TEST SUMMARY ===');
  try {
    const s = getTodaySummary();
    Logger.log('Tab ' + s.tab + ' | total ' + s.total + ' (ARR ' + s.arr + ' / DEP ' + s.dep + ')');
    Logger.log('STANDBY ' + s.standby + ' · ON PROCESS ' + s.process + ' · COMPLETED ' + s.completed);
    Logger.log('ยังไม่ assign: ' + s.unassigned + ' | on duty ' + s.onDuty + ' / day off ' + s.dayOff);
    Logger.log('SSR: ' + JSON.stringify(s.bySvc));
    Logger.log('Airline: ' + JSON.stringify(s.byAirline));
  } catch (e) {
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

// ── v5.0: งานที่ยังค้างอยู่ + งานของ porter คนหนึ่ง ──
function testActive() {
  const t = scanOpenTrips();
  const line = function (x) {
    return '  ' + x.id + ' | ' + x.snick + ' | ' + x.ctrl
      + ' | เริ่ม ' + x.date + ' ' + x.startTime
      + (x.ageHours === null ? '' : ' (' + x.ageHours + ' ชม.)')
      + (x.caseNo ? ' | เคส #' + x.caseNo : '') + ' | ✈ ' + x.flight;
  };

  Logger.log('กำลังให้บริการอยู่จริง: ' + t.active.length);
  t.active.forEach(function (x) { Logger.log(line(x)); });

  Logger.log('งานค้าง (เปิดเกิน ' + STALE_HOURS + ' ชม. ไม่ได้กดจบ): ' + t.stale.length);
  t.stale.forEach(function (x) { Logger.log(line(x)); });
  if (t.stale.length) {
    Logger.log('  ↳ รถเข็นพวกนี้ยังถูกกันไว้ — รัน closeStaleTrips() หรือกดปุ่มในแอป (LP) เพื่อปลดล็อก');
  }
  return t;
}

// แก้ชื่อตรงนี้เป็นชื่ออังกฤษอย่างที่สะกดในชีท แล้วกด Run เพื่อเช็คการจับคู่
function testMyJobs() {
  const name = 'Chatchai';
  const r = getMyJobs('', name);
  Logger.log('เคสของ "' + name + '" ใน tab ' + r.tab + ': ' + r.jobs.length + ' เคส');
  r.jobs.forEach(function (c) {
    Logger.log('  #' + c.no + ' ' + c.flight + ' ' + c.svc + ' · ' + c.state + ' · ' + c.porter);
  });
  Logger.log('งานที่กำลังทำอยู่ทั้งระบบ: ' + r.active.length);
  return r;
}

// เช็คว่าแถวว่างในชีทวันนี้ยังพอให้ LP เปิดเคสใหม่ได้ไหม
function testFreeCaseRows() {
  const sheet = openPorterDayTab();
  const data  = sheet.getDataRange().getValues();
  let used = 0, free = 0;
  for (let r = 0; r < data.length; r++) {
    const no = parseInt(data[r][PCOL.NO], 10);
    if (!no || isNaN(no)) continue;
    const air = String(data[r][PCOL.AIRLINE] || '').trim();
    const flt = String(data[r][PCOL.FLTNO]   || '').trim();
    if (!air && !flt) free++; else used++;
  }
  Logger.log('tab ' + sheet.getName() + ': ใช้ไปแล้ว ' + used + ' แถว · ว่าง ' + free + ' แถว');
  return { tab: sheet.getName(), used: used, free: free };
}
