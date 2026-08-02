/**
 * O83 Work Alert — Google Apps Script
 * ใช้กับ GitHub Pages index.html
 *
 * ข้อมูลที่แก้ได้โดยไม่ต้องแก้โค้ด:
 * - Sheet "พนักงาน" สำหรับรายชื่อ
 * - Sheet "กะ" สำหรับรายการกะและลำดับการแสดงผล
 */

// ══════════════════════════════════════════════
//  DEFAULT CONFIG
// ══════════════════════════════════════════════
const DEFAULT_CONFIG = {
  SPREADSHEET_ID: '1fC5Dvs-gEJ04wiF2bikD4rMPaK7S24PrRcuDJCBpnRM',
  SHEET_NAME: 'รายการเข้างาน',
  STAFF_SHEET_NAME: 'พนักงาน',
  SHIFT_SHEET_NAME: 'กะ',
  DRIVE_FOLDER_ID: '1sGEbvwt00KYw_i2b2G9DWDwrdBaNk1i3',
  PROJECT_NAME: 'ดิ ออริจิ้น รามอินทรา83 สเตชั่น',
  MAX_REMARK_CHARS: 200,
  TIMEZONE: 'Asia/Bangkok'
};

const DEFAULT_SHIFTS = [
  ['07:00-16:00', true, 1],
  ['08:00-17:00', true, 2],
  ['09:00-18:00', true, 3],
  ['10:00-19:00', true, 4],
  ['13:00-22:00', true, 5],
  ['22:00-07:00', true, 6]
];

// ══════════════════════════════════════════════
//  CONFIG HELPERS
// ══════════════════════════════════════════════
function getConfig_() {
  const props = PropertiesService.getScriptProperties();
  return {
    SPREADSHEET_ID: (props.getProperty('SPREADSHEET_ID') || DEFAULT_CONFIG.SPREADSHEET_ID).trim(),
    SHEET_NAME: (props.getProperty('SHEET_NAME') || DEFAULT_CONFIG.SHEET_NAME).trim(),
    STAFF_SHEET_NAME: (props.getProperty('STAFF_SHEET_NAME') || DEFAULT_CONFIG.STAFF_SHEET_NAME).trim(),
    SHIFT_SHEET_NAME: (props.getProperty('SHIFT_SHEET_NAME') || DEFAULT_CONFIG.SHIFT_SHEET_NAME).trim(),
    DRIVE_FOLDER_ID: (props.getProperty('DRIVE_FOLDER_ID') || DEFAULT_CONFIG.DRIVE_FOLDER_ID).trim(),
    PROJECT_NAME: (props.getProperty('PROJECT_NAME') || DEFAULT_CONFIG.PROJECT_NAME).trim(),
    MAX_REMARK_CHARS: Number(props.getProperty('MAX_REMARK_CHARS') || DEFAULT_CONFIG.MAX_REMARK_CHARS),
    TIMEZONE: (props.getProperty('TIMEZONE') || DEFAULT_CONFIG.TIMEZONE).trim()
  };
}

function checkSetup() {
  const cfg = getConfig_();
  const report = {
    SPREADSHEET_ID: cfg.SPREADSHEET_ID,
    SHEET_NAME: cfg.SHEET_NAME,
    STAFF_SHEET_NAME: cfg.STAFF_SHEET_NAME,
    SHIFT_SHEET_NAME: cfg.SHIFT_SHEET_NAME,
    DRIVE_FOLDER_ID: cfg.DRIVE_FOLDER_ID || 'EMPTY',
    PROJECT_NAME: cfg.PROJECT_NAME,
    MAX_REMARK_CHARS: cfg.MAX_REMARK_CHARS,
    TIMEZONE: cfg.TIMEZONE,
    shiftCount: getShiftList_().length
  };
  console.log(JSON.stringify(report, null, 2));
  return report;
}

// ══════════════════════════════════════════════
//  SETUP SHEETS
// ══════════════════════════════════════════════
function setupSystem() {
  const cfg = getConfig_();
  const ss = SpreadsheetApp.openById(cfg.SPREADSHEET_ID);

  let logSheet = ss.getSheetByName(cfg.SHEET_NAME);
  if (!logSheet) logSheet = ss.insertSheet(cfg.SHEET_NAME);
  if (logSheet.getLastRow() === 0) {
    logSheet.appendRow([
      'วันที่บันทึก', 'วันที่ทำงาน', 'รอบ', 'ชื่อพนักงาน', 'ชื่อเล่น',
      'ตำแหน่ง', 'กะเข้างาน', 'สถานะ', 'ประเภทการลา/มาสาย',
      'หมายเหตุ', 'Timestamp', 'Link รูปภาพ'
    ]);
  }

  let staffSheet = ss.getSheetByName(cfg.STAFF_SHEET_NAME);
  if (!staffSheet) {
    staffSheet = ss.insertSheet(cfg.STAFF_SHEET_NAME);
    staffSheet.appendRow(['name', 'nick', 'pos', 'type']);
    staffSheet.getRange(2, 1, 9, 4).setValues([
      ['จิราภรณ์', 'เจน', 'ผช.ผจก.อาคาร', 'mgr'],
      ['กัญญาภรณ์', 'เมย์', 'ธุรการ', 'op'],
      ['ศิรดา', 'การ์ตูน', 'ธุรการ', 'op'],
      ['นันท์สินี', 'นัน', 'DWธุรการ', 'op'],
      ['วรัญญู', 'ออฟ', 'Support ช่างอาคาร', 'tech'],
      ['อิทธิพงศ์', 'หนุ่ม', 'Support ช่างอาคาร', 'tech'],
      ['นพรัตน์', 'บอล', 'DWช่างอาคาร', 'tech'],
      ['บดินทร์', 'ไอติม', 'DWช่างอาคาร', 'tech'],
      ['ทรงพล', 'ป้าบอล', 'ช่างอาคาร', 'tech']
    ]);
  }

  let shiftSheet = ss.getSheetByName(cfg.SHIFT_SHEET_NAME);
  if (!shiftSheet) shiftSheet = ss.insertSheet(cfg.SHIFT_SHEET_NAME);
  if (shiftSheet.getLastRow() === 0) {
    shiftSheet.getRange(1, 1, 1, 3).setValues([['กะ', 'เปิดใช้งาน', 'ลำดับ']]);
    shiftSheet.getRange(2, 1, DEFAULT_SHIFTS.length, 3).setValues(DEFAULT_SHIFTS);
    shiftSheet.setFrozenRows(1);
  }

  console.log('setupSystem complete');
}

// ══════════════════════════════════════════════
//  WEB APP — GET
// ══════════════════════════════════════════════
function doGet(e) {
  try {
    setupSystem();
    const action = String((e && e.parameter && e.parameter.action) || 'ping');

    if (action === 'getStaff') {
      const config = getAppConfig_();
      return outputJsonOrJsonp_(e, {
        success: true,
        staff: getStaffList_(),
        config,
        // ฟิลด์ด้านบนช่วยให้หน้าเว็บรุ่นเก่าค่อย ๆ อัปเดตได้โดยไม่พัง
        shifts: config.shifts,
        projectName: config.projectName,
        maxRemarkChars: config.maxRemarkChars
      });
    }

    if (action === 'getConfig') {
      return outputJsonOrJsonp_(e, { success: true, config: getAppConfig_() });
    }

    if (action === 'check') {
      return outputJsonOrJsonp_(e, { success: true, setup: checkSetup() });
    }

    return outputJsonOrJsonp_(e, {
      success: true,
      message: 'O83 GAS is running',
      action,
      time: formatTimestamp_(new Date(), getConfig_().TIMEZONE)
    });
  } catch (err) {
    console.log('doGet error:', err.stack || err.message);
    return outputJsonOrJsonp_(e, { success: false, error: err.message });
  }
}

function getAppConfig_() {
  const cfg = getConfig_();
  return {
    projectName: cfg.PROJECT_NAME,
    maxRemarkChars: cfg.MAX_REMARK_CHARS,
    timezone: cfg.TIMEZONE,
    shifts: getShiftList_()
  };
}

function getStaffList_() {
  const cfg = getConfig_();
  const ss = SpreadsheetApp.openById(cfg.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(cfg.STAFF_SHEET_NAME);
  if (!sheet) throw new Error('ไม่พบ Sheet รายชื่อพนักงาน: ' + cfg.STAFF_SHEET_NAME);

  const values = sheet.getDataRange().getValues();
  if (!values.length) return [];

  const headers = values[0].map(String);
  const list = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const item = {};
    headers.forEach((h, j) => item[h] = row[j]);
    if (item.nick) list.push(item);
  }
  return list;
}

function getShiftList_() {
  const cfg = getConfig_();
  const ss = SpreadsheetApp.openById(cfg.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(cfg.SHIFT_SHEET_NAME);
  if (!sheet) return [];

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];

  const headers = values[0].map(h => String(h || '').trim().toLowerCase());
  const shiftIndex = findHeaderIndex_(headers, ['กะ', 'shift', 'ชื่อกะ', 'เวลา']);
  const activeIndex = findHeaderIndex_(headers, ['เปิดใช้งาน', 'active', 'ใช้งาน', 'สถานะ']);
  const orderIndex = findHeaderIndex_(headers, ['ลำดับ', 'sort', 'order']);
  const items = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const shift = String(row[shiftIndex === -1 ? 0 : shiftIndex] || '').trim();
    if (!shift || (activeIndex !== -1 && !isActive_(row[activeIndex]))) continue;

    const orderValue = orderIndex === -1 ? i : Number(row[orderIndex]);
    items.push({ shift, order: Number.isFinite(orderValue) ? orderValue : i });
  }

  items.sort((a, b) => a.order - b.order);
  const seen = {};
  return items.filter(item => {
    if (seen[item.shift]) return false;
    seen[item.shift] = true;
    return true;
  }).map(item => item.shift);
}

function findHeaderIndex_(headers, aliases) {
  for (let i = 0; i < aliases.length; i++) {
    const index = headers.indexOf(String(aliases[i]).toLowerCase());
    if (index !== -1) return index;
  }
  return -1;
}

function isActive_(value) {
  if (value === '' || value === null || typeof value === 'undefined') return true;
  if (value === false) return false;
  const normalized = String(value).trim().toLowerCase();
  return !['false', '0', 'no', 'n', 'ไม่', 'ไม่ใช่', 'ปิด', 'ปิดใช้งาน', 'inactive'].includes(normalized);
}

function outputJsonOrJsonp_(e, obj) {
  const json = JSON.stringify(obj);
  const callback = e && e.parameter && e.parameter.callback;
  if (callback && /^[a-zA-Z_$][0-9a-zA-Z_$\.]*$/.test(callback)) {
    return ContentService.createTextOutput(callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

// ══════════════════════════════════════════════
//  WEB APP — POST
// ══════════════════════════════════════════════
function doPost(e) {
  try {
    const payload = parsePostPayload_(e);
    const result = handleWorkAlertSubmit_(payload);
    return outputPostResult_(result);
  } catch (err) {
    console.log('doPost error:', err.stack || err.message);
    return outputPostResult_({ success: false, error: err.message });
  }
}

function parsePostPayload_(e) {
  if (!e) throw new Error('ไม่มี event object จาก doPost');
  if (e.parameter && e.parameter.payload) return JSON.parse(e.parameter.payload);

  if (e.postData && e.postData.contents) {
    const contents = e.postData.contents;
    try {
      return JSON.parse(contents);
    } catch (jsonErr) {
      const params = parseQueryString_(contents);
      if (params.payload) return JSON.parse(params.payload);
      throw jsonErr;
    }
  }
  throw new Error('ไม่พบ payload ใน doPost');
}

function parseQueryString_(qs) {
  return String(qs || '').split('&').reduce((acc, part) => {
    const idx = part.indexOf('=');
    if (idx === -1) return acc;
    const key = decodeURIComponent(part.slice(0, idx).replace(/\+/g, ' '));
    const val = decodeURIComponent(part.slice(idx + 1).replace(/\+/g, ' '));
    acc[key] = val;
    return acc;
  }, {});
}

function outputPostResult_(result) {
  const safeJson = JSON.stringify(result).replace(/</g, '\\u003c');
  const html = `<!doctype html><html><head><meta charset="utf-8"></head><body><pre>${escapeHtml_(safeJson)}</pre><script>try{parent.postMessage(${safeJson}, '*');}catch(e){}<\\/script></body></html>`;
  return HtmlService.createHtmlOutput(html);
}

function escapeHtml_(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

// ══════════════════════════════════════════════
//  SUBMIT WORK ALERT — บันทึก Sheet + Drive
// ══════════════════════════════════════════════
function handleWorkAlertSubmit_(data) {
  setupSystem();
  const cfg = getConfig_();
  if (!data || typeof data !== 'object') throw new Error('payload ไม่ถูกต้อง');

  const staffData = Array.isArray(data.staffData)
    ? data.staffData.map(s => normalizeStaffRecord_(s, cfg))
    : [];
  const roundTh = data.round === 'eve' ? 'รอบเย็น' : 'รอบเช้า';
  const workDate = data.date || formatThaiDate_(new Date(), cfg.TIMEZONE);
  const timestamp = formatTimestamp_(new Date(), cfg.TIMEZONE);
  let driveUrl = '-';

  if (data.imageBase64) {
    if (!cfg.DRIVE_FOLDER_ID) throw new Error('ยังไม่ได้ตั้งค่า DRIVE_FOLDER_ID');
    driveUrl = saveImageToDrive_(data.imageBase64);
  }

  appendLogRows_(staffData, {
    date: data.date || workDate,
    workDate,
    roundTh,
    timestamp,
    imageUrl: driveUrl
  });

  return {
    success: true,
    message: 'บันทึกข้อมูลแล้ว',
    driveUrl,
    staffCount: staffData.length,
    time: timestamp
  };
}

function normalizeStaffRecord_(s, cfg) {
  const max = cfg.MAX_REMARK_CHARS || 200;
  return {
    name: String(s.name || '').trim(),
    nick: String(s.nick || '').trim(),
    pos: String(s.pos || '').trim(),
    shift: s.shift || null,
    offType: s.offType || null,
    lateType: s.lateType || null,
    remark: String(s.remark || '').trim().slice(0, max)
  };
}

function appendLogRows_(staffData, meta) {
  const cfg = getConfig_();
  const ss = SpreadsheetApp.openById(cfg.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(cfg.SHEET_NAME);
  if (!sheet) throw new Error('ไม่พบ Sheet บันทึก: ' + cfg.SHEET_NAME);
  if (!staffData.length) return;

  const rows = staffData.map(s => {
    let status = 'ยังไม่ระบุ';
    let shiftTime = '-';
    let type = '-';

    if (s.shift === 'off') {
      status = 'หยุด/ลา';
      type = s.offType || '-';
    } else if (s.shift) {
      status = s.lateType ? 'ทำงาน/มาสาย' : 'ทำงาน';
      shiftTime = s.shift;
      type = s.lateType || '-';
    } else if (s.lateType) {
      status = 'มาสาย';
      type = s.lateType;
    }

    return [
      meta.date,
      meta.workDate,
      meta.roundTh,
      s.name,
      s.nick,
      s.pos,
      shiftTime,
      status,
      type,
      s.remark || '-',
      meta.timestamp,
      meta.imageUrl || '-'
    ];
  });

  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}

function formatThaiDate_(date, timezone) {
  const parts = Utilities.formatDate(date, timezone || 'Asia/Bangkok', 'd/M/yyyy').split('/');
  return `${parts[0]}/${parts[1]}/${Number(parts[2]) + 543}`;
}

function formatTimestamp_(date, timezone) {
  return Utilities.formatDate(date, timezone || 'Asia/Bangkok', 'dd/MM/yyyy HH:mm:ss');
}

function formatFileTimestamp_(date, timezone) {
  return Utilities.formatDate(date, timezone || 'Asia/Bangkok', 'yyyyMMdd_HHmmss');
}

// ══════════════════════════════════════════════
//  DRIVE SAVE
// ══════════════════════════════════════════════
function saveImageToDrive_(base64Image) {
  const cfg = getConfig_();
  if (!cfg.DRIVE_FOLDER_ID) throw new Error('ยังไม่ได้ตั้งค่า DRIVE_FOLDER_ID');

  const bytes = Utilities.base64Decode(base64Image);
  const blob = Utilities.newBlob(bytes, 'image/png',
    formatFileTimestamp_(new Date(), cfg.TIMEZONE) + '_Smart_time_dashboard_announcement.png');
  const folder = DriveApp.getFolderById(cfg.DRIVE_FOLDER_ID);
  const file = folder.createFile(blob);
  return file.getUrl();
}

function sanitizeFileName_(s) {
  return String(s).replace(/[\\/:*?"<>|\s]+/g, '_').slice(0, 80);
}

// ══════════════════════════════════════════════
//  TEST FUNCTIONS
// ══════════════════════════════════════════════
function testGetStaff() {
  const staff = getStaffList_();
  console.log(JSON.stringify(staff, null, 2));
  return staff;
}

function testGetShifts() {
  const shifts = getShiftList_();
  console.log(JSON.stringify(shifts, null, 2));
  return shifts;
}

function testSaveImageToDrive() {
  const testBase64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8/5+hHgAHggJ/PchI6QAAAABJRU5ErkJggg==';
  const url = saveImageToDrive_(testBase64, 'test');
  console.log('Drive URL:', url);
}
