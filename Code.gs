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

const SUBMISSION_SHEET_NAME = 'ประวัติการส่ง';
const CHANGE_SHEET_NAME = 'ประวัติการเปลี่ยนแปลง';
const SUMMARY_SHEET_NAME = 'สรุปประจำเดือน';
const SUBMISSION_HEADERS = [
  'Submission ID', 'วันที่ทำงาน', 'วันที่ทำงาน Key', 'รอบ', 'เวอร์ชัน',
  'สถานะล่าสุด', 'จำนวนการเปลี่ยนแปลง', 'จำนวนพนักงาน', 'Timestamp',
  'ชื่อไฟล์', 'File ID', 'Link รูปภาพ', 'ประเภทการส่ง'
];
const CHANGE_HEADERS = [
  'วันที่ทำงาน', 'วันที่ทำงาน Key', 'ชื่อพนักงาน', 'ชื่อเล่น', 'รอบ', 'เวอร์ชัน',
  'ประเภทการเปลี่ยนแปลง', 'สถานะเดิม', 'สถานะใหม่', 'รายละเอียดเดิม',
  'รายละเอียดใหม่', 'Timestamp', 'Submission ID', 'Link รูปภาพ'
];
const SUMMARY_HEADERS = [
  'เดือน', 'ชื่อพนักงาน', 'มาสาย (ครั้ง)', 'ขาดงาน (ครั้ง)', 'ลา (ครั้ง)',
  'วันที่มาสาย', 'วันที่ขาดงาน', 'วันที่ลา', 'วันที่มีการอัปเดต', 'อัปเดตล่าสุด'
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
    SUBMISSION_SHEET_NAME,
    CHANGE_SHEET_NAME,
    SUMMARY_SHEET_NAME,
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

  ensureHeaderColumns_(logSheet, [
    'Submission ID', 'วันที่ทำงาน Key', 'เวอร์ชัน', 'สถานะเวอร์ชัน',
    'จำนวนการเปลี่ยนแปลง', 'ประเภทการส่ง'
  ]);

  const submissionSheet = getOrCreateSheet_(ss, SUBMISSION_SHEET_NAME, SUBMISSION_HEADERS);
  ensureHeaderColumns_(submissionSheet, SUBMISSION_HEADERS);
  const changeSheet = getOrCreateSheet_(ss, CHANGE_SHEET_NAME, CHANGE_HEADERS);
  ensureHeaderColumns_(changeSheet, CHANGE_HEADERS);
  const summarySheet = getOrCreateSheet_(ss, SUMMARY_SHEET_NAME, SUMMARY_HEADERS);
  ensureHeaderColumns_(summarySheet, SUMMARY_HEADERS);
  [submissionSheet, changeSheet, summarySheet].forEach(sheet => sheet.setFrozenRows(1));

  console.log('setupSystem complete');
}

function getOrCreateSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() === 0) sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  return sheet;
}

function ensureHeaderColumns_(sheet, headers) {
  if (!headers || !headers.length) return;
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return;
  }
  const existing = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0]
    .map(value => String(value || '').trim());
  const missing = headers.filter(header => existing.indexOf(header) === -1);
  if (missing.length) {
    sheet.getRange(1, existing.length + 1, 1, missing.length).setValues([missing]);
  }
}

function getHeaderMap_(sheet) {
  if (!sheet || sheet.getLastRow() === 0) return {};
  const headers = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0];
  return headers.reduce((map, header, index) => {
    const key = String(header || '').trim();
    if (key) map[key] = index + 1;
    return map;
  }, {});
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
  const scriptClose = '<' + '/script>';
  const html = `<!doctype html><html><head><meta charset="utf-8"></head><body><pre>${escapeHtml_(safeJson)}</pre><script>try{window.top.postMessage(${safeJson}, '*');}catch(e){}${scriptClose}</body></html>`;
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
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    setupSystem();
    const cfg = getConfig_();
    if (!data || typeof data !== 'object') throw new Error('payload ไม่ถูกต้อง');

    const staffData = Array.isArray(data.staffData)
      ? data.staffData.map(s => normalizeStaffRecord_(s, cfg))
      : [];
    if (!staffData.length) throw new Error('ไม่พบข้อมูลพนักงานในรายการส่ง');

    const now = new Date();
    const roundTh = data.round === 'eve' ? 'รอบเย็น' : 'รอบเช้า';
    const roundCode = data.round === 'eve' ? 'EVE' : 'MORNING';
    const workDateKey = getWorkDateKey_(data, cfg.TIMEZONE);
    const workDate = data.date || formatDisplayDateFromKey_(workDateKey);
    const timestamp = formatTimestamp_(now, cfg.TIMEZONE);
    const currentSnapshot = buildStaffSnapshot_(staffData);
    const previous = getLatestSubmission_(workDateKey);
    const previousSnapshot = previous ? getSubmissionSnapshot_(previous) : {};
    const changes = previous ? diffStaffSnapshots_(previousSnapshot, currentSnapshot) : [];

    if (previous && !changes.length) {
      return {
        success: true,
        noChange: true,
        version: previous.version,
        driveUrl: previous.imageUrl || '-',
        message: 'ไม่มีการเปลี่ยนแปลง จึงไม่ได้สร้างไฟล์หรือบันทึกซ้ำ',
        time: timestamp
      };
    }

    const version = previous ? previous.version + 1 : 1;
    const submissionId = createSubmissionId_(workDateKey, version, now, cfg.TIMEZONE);
    const submissionType = previous ? 'อัปเดต' : 'ตารางหลัก';
    let image = { id: '', url: '-', name: '' };

    if (data.imageBase64) {
      if (!cfg.DRIVE_FOLDER_ID) throw new Error('ยังไม่ได้ตั้งค่า DRIVE_FOLDER_ID');
      const fileName = buildImageFileName_(workDateKey, version, roundCode, now, cfg.TIMEZONE, 'LATEST');
      image = saveImageToDrive_(data.imageBase64, fileName);
    }

    if (previous) markSubmissionHistory_(previous, 'HISTORY');

    const meta = {
      date: workDate,
      workDate,
      workDateKey,
      roundTh,
      timestamp,
      imageUrl: image.url,
      imageFileId: image.id,
      imageFileName: image.name,
      submissionId,
      version,
      versionStatus: 'LATEST',
      changeCount: changes.length,
      submissionType
    };

    appendLogRows_(staffData, meta);
    appendSubmissionHistory_(meta, staffData.length);
    if (changes.length) appendChangeRows_(changes, meta, previousSnapshot);
    rebuildMonthlySummary_();

    return {
      success: true,
      noChange: false,
      version,
      submissionId,
      driveUrl: image.url,
      fileName: image.name,
      staffCount: staffData.length,
      changeCount: changes.length,
      message: previous ? 'บันทึกการเปลี่ยนแปลงแล้ว' : 'บันทึกตารางหลักแล้ว',
      time: timestamp
    };
  } finally {
    lock.releaseLock();
  }
}

function getWorkDateKey_(data, timezone) {
  const tz = timezone || 'Asia/Bangkok';
  if (data && data.dateObj) {
    const parsed = new Date(data.dateObj);
    if (!isNaN(parsed.getTime())) return Utilities.formatDate(parsed, tz, 'yyyyMMdd');
  }
  const text = String((data && data.date) || '').trim();
  const match = text.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/);
  if (match) {
    let year = Number(match[3]);
    if (year > 2400) year -= 543;
    const parsed = new Date(year, Number(match[2]) - 1, Number(match[1]));
    if (!isNaN(parsed.getTime())) return Utilities.formatDate(parsed, tz, 'yyyyMMdd');
  }
  return Utilities.formatDate(new Date(), tz, 'yyyyMMdd');
}

function formatDisplayDateFromKey_(key) {
  const match = String(key || '').match(/^(\d{4})(\d{2})(\d{2})$/);
  if (!match) return formatThaiDate_(new Date(), getConfig_().TIMEZONE);
  return `${Number(match[3])}/${Number(match[2])}/${Number(match[1]) + 543}`;
}

function createSubmissionId_(workDateKey, version, date, timezone) {
  return `${workDateKey}-v${String(version).padStart(2, '0')}-${formatFileTimestamp_(date, timezone).replace('_', '')}`;
}

function buildStaffSnapshot_(staffData) {
  return staffData.reduce((map, staff) => {
    const key = normalizeNameKey_(staff.name);
    if (key) map[key] = {
      name: staff.name,
      nick: staff.nick,
      pos: staff.pos,
      shift: staff.shift || '',
      offType: staff.offType || '',
      lateType: staff.lateType || '',
      remark: staff.remark || ''
    };
    return map;
  }, {});
}

function normalizeNameKey_(name) {
  return String(name || '').trim().replace(/\s+/g, ' ').toLowerCase();
}

function snapshotValue_(staff) {
  return [staff.shift || '', staff.offType || '', staff.lateType || '', staff.remark || '']
    .map(value => String(value).trim()).join('|');
}

function diffStaffSnapshots_(previous, current) {
  const changes = [];
  Object.keys(current).forEach(key => {
    const oldStaff = previous[key] || { name: current[key].name, nick: current[key].nick, pos: current[key].pos };
    const newStaff = current[key];
    if (previous[key] && snapshotValue_(oldStaff) === snapshotValue_(newStaff)) return;
    changes.push({
      name: newStaff.name,
      nick: newStaff.nick,
      oldStaff,
      newStaff,
      type: detectChangeType_(oldStaff, newStaff),
      oldLabel: describeStaffState_(oldStaff),
      newLabel: describeStaffState_(newStaff)
    });
  });
  return changes;
}

function detectChangeType_(oldStaff, newStaff) {
  const oldOff = String(oldStaff.offType || '').trim();
  const newOff = String(newStaff.offType || '').trim();
  const oldLate = String(oldStaff.lateType || '').trim();
  const newLate = String(newStaff.lateType || '').trim();
  if (newOff === 'ขาดงาน') return 'ขาดงาน';
  if (isLeaveType_(newOff)) return 'ลา';
  if (newLate || oldLate) return 'มาสาย';
  if (oldOff || newOff) return 'ขาด/ลา';
  if (String(oldStaff.shift || '') !== String(newStaff.shift || '')) return 'เปลี่ยนกะ';
  return 'แก้ไขข้อมูล';
}

function isLeaveType_(value) {
  const text = String(value || '').trim();
  return text.indexOf('ลา') === 0 || text === 'หยุดพักร้อน';
}

function describeStaffState_(staff) {
  if (!staff) return '-';
  if (staff.offType) return `หยุด/ลา: ${staff.offType}`;
  if (staff.lateType) return `มาสาย: ${staff.lateType}`;
  if (staff.shift) return `ทำงาน: ${staff.shift}`;
  return 'ยังไม่ระบุ';
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

  ensureHeaderColumns_(sheet, [
    'Submission ID', 'วันที่ทำงาน Key', 'เวอร์ชัน', 'สถานะเวอร์ชัน',
    'จำนวนการเปลี่ยนแปลง', 'ประเภทการส่ง'
  ]);
  const headerMap = getHeaderMap_(sheet);
  const width = Math.max(sheet.getLastColumn(), 18);

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

    const row = new Array(width).fill('');
    const values = {
      'วันที่บันทึก': meta.date,
      'วันที่ทำงาน': meta.workDate,
      'รอบ': meta.roundTh,
      'ชื่อพนักงาน': s.name,
      'ชื่อเล่น': s.nick,
      'ตำแหน่ง': s.pos,
      'กะเข้างาน': shiftTime,
      'สถานะ': status,
      'ประเภทการลา/มาสาย': type,
      'หมายเหตุ': s.remark || '-',
      'Timestamp': meta.timestamp,
      'Link รูปภาพ': meta.imageUrl || '-',
      'Submission ID': meta.submissionId,
      'วันที่ทำงาน Key': meta.workDateKey,
      'เวอร์ชัน': meta.version,
      'สถานะเวอร์ชัน': meta.versionStatus,
      'จำนวนการเปลี่ยนแปลง': meta.changeCount,
      'ประเภทการส่ง': meta.submissionType
    };
    Object.keys(values).forEach(header => {
      const column = headerMap[header];
      if (column) row[column - 1] = values[header];
    });
    return row;
  });

  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}

function appendSubmissionHistory_(meta, staffCount) {
  const cfg = getConfig_();
  const ss = SpreadsheetApp.openById(cfg.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SUBMISSION_SHEET_NAME);
  const row = [
    meta.submissionId,
    meta.workDate,
    meta.workDateKey,
    meta.roundTh,
    meta.version,
    meta.versionStatus,
    meta.changeCount,
    staffCount,
    meta.timestamp,
    meta.imageFileName || '-',
    meta.imageFileId || '-',
    meta.imageUrl || '-',
    meta.submissionType
  ];
  sheet.getRange(sheet.getLastRow() + 1, 1, 1, row.length).setValues([row]);
}

function appendChangeRows_(changes, meta, previousSnapshot) {
  if (!changes || !changes.length) return;
  const cfg = getConfig_();
  const ss = SpreadsheetApp.openById(cfg.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CHANGE_SHEET_NAME);
  const rows = changes.map(change => [
    meta.workDate,
    meta.workDateKey,
    change.name,
    change.nick,
    meta.roundTh,
    meta.version,
    change.type,
    change.oldLabel,
    change.newLabel,
    describeStaffState_(change.oldStaff),
    describeStaffState_(change.newStaff),
    meta.timestamp,
    meta.submissionId,
    meta.imageUrl || '-'
  ]);
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}

function getLatestSubmission_(workDateKey) {
  const cfg = getConfig_();
  const ss = SpreadsheetApp.openById(cfg.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SUBMISSION_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return null;
  const values = sheet.getDataRange().getValues();
  let latest = null;
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (String(row[2] || '').trim() !== String(workDateKey)) continue;
    const version = Number(row[4]) || 0;
    if (!latest || version > latest.version) {
      latest = {
        submissionId: String(row[0] || ''),
        workDate: String(row[1] || ''),
        workDateKey: String(row[2] || ''),
        roundTh: String(row[3] || ''),
        version,
        status: String(row[5] || ''),
        imageFileName: String(row[9] || ''),
        fileId: String(row[10] || '').replace(/^-$/, ''),
        imageUrl: String(row[11] || '').replace(/^-$/, ''),
        timestamp: String(row[8] || '')
      };
    }
  }
  return latest;
}

function getSubmissionSnapshot_(submission) {
  if (!submission || !submission.submissionId) return {};
  const cfg = getConfig_();
  const ss = SpreadsheetApp.openById(cfg.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(cfg.SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return {};
  const headers = getHeaderMap_(sheet);
  const values = sheet.getDataRange().getValues();
  const get = (row, header, fallback) => {
    const index = headers[header];
    return index ? row[index - 1] : fallback;
  };
  return values.slice(1).reduce((map, row) => {
    if (String(get(row, 'Submission ID', '') || '') !== submission.submissionId) return map;
    const name = String(get(row, 'ชื่อพนักงาน', '') || '').trim();
    const key = normalizeNameKey_(name);
    if (!key) return map;
    const status = String(get(row, 'สถานะ', '') || '');
    const type = String(get(row, 'ประเภทการลา/มาสาย', '') || '');
    const shift = String(get(row, 'กะเข้างาน', '') || '');
    map[key] = {
      name,
      nick: String(get(row, 'ชื่อเล่น', '') || '').trim(),
      pos: String(get(row, 'ตำแหน่ง', '') || '').trim(),
      shift: shift === '-' ? '' : shift,
      offType: status === 'หยุด/ลา' ? type.replace(/^-$/, '') : '',
      lateType: status.indexOf('มาสาย') !== -1 ? type.replace(/^-$/, '') : '',
      remark: String(get(row, 'หมายเหตุ', '') || '').replace(/^-$/, '').trim()
    };
    return map;
  }, {});
}

function markSubmissionHistory_(submission, state) {
  if (!submission || !submission.submissionId) return;
  const cfg = getConfig_();
  const ss = SpreadsheetApp.openById(cfg.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SUBMISSION_SHEET_NAME);
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0] || '') !== submission.submissionId) continue;
    sheet.getRange(i + 1, 6).setValue(state);
    if (state === 'HISTORY' && submission.fileId) {
      const renamed = renameDriveFile_(submission.fileId, submission.imageFileName, state);
      if (renamed) sheet.getRange(i + 1, 10).setValue(renamed);
    }
    break;
  }
}

function renameDriveFile_(fileId, oldName, state) {
  try {
    const file = DriveApp.getFileById(fileId);
    const currentName = oldName || file.getName();
    const newName = state === 'HISTORY'
      ? currentName.replace('_LATEST_', '_HISTORY_')
      : currentName.replace('_HISTORY_', '_LATEST_');
    if (newName !== currentName) file.setName(newName);
    return newName;
  } catch (err) {
    console.log('renameDriveFile_ skipped:', err.message);
    return '';
  }
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

function buildImageFileName_(workDateKey, version, roundCode, date, timezone, state) {
  const versionText = String(version).padStart(2, '0');
  const sentAt = formatFileTimestamp_(date, timezone);
  return `${workDateKey}_v${versionText}_${sentAt}_${roundCode}_${state}_Smart_time_dashboard_announcement.png`;
}

function rebuildMonthlySummary_() {
  const cfg = getConfig_();
  const ss = SpreadsheetApp.openById(cfg.SPREADSHEET_ID);
  const summarySheet = ss.getSheetByName(SUMMARY_SHEET_NAME);
  const submissionSheet = ss.getSheetByName(SUBMISSION_SHEET_NAME);
  const logSheet = ss.getSheetByName(cfg.SHEET_NAME);
  const changeSheet = ss.getSheetByName(CHANGE_SHEET_NAME);
  if (!summarySheet || !submissionSheet || !logSheet) return;

  const latestByDate = {};
  if (submissionSheet.getLastRow() >= 2) {
    submissionSheet.getDataRange().getValues().slice(1).forEach(row => {
      const workDateKey = String(row[2] || '').trim();
      const version = Number(row[4]) || 0;
      if (!workDateKey) return;
      if (!latestByDate[workDateKey] || version > latestByDate[workDateKey].version) {
        latestByDate[workDateKey] = { submissionId: String(row[0] || ''), version };
      }
    });
  }

  const headerMap = getHeaderMap_(logSheet);
  const rows = logSheet.getLastRow() >= 2 ? logSheet.getDataRange().getValues().slice(1) : [];
  const summary = {};
  const get = (row, header, fallback) => {
    const index = headerMap[header];
    return index ? row[index - 1] : fallback;
  };

  rows.forEach(row => {
    const workDateKey = String(get(row, 'วันที่ทำงาน Key', '') || '').trim();
    const submissionId = String(get(row, 'Submission ID', '') || '').trim();
    if (!workDateKey || !submissionId || !latestByDate[workDateKey] || latestByDate[workDateKey].submissionId !== submissionId) return;
    const name = String(get(row, 'ชื่อพนักงาน', '') || '').trim();
    if (!name) return;
    const month = workDateKey.slice(0, 6);
    const key = `${month}|${normalizeNameKey_(name)}`;
    const item = summary[key] || {
      month,
      name,
      late: 0,
      absent: 0,
      leave: 0,
      lateDates: [],
      absentDates: [],
      leaveDates: [],
      updateDates: [],
      lastUpdated: ''
    };
    const status = String(get(row, 'สถานะ', '') || '');
    const type = String(get(row, 'ประเภทการลา/มาสาย', '') || '');
    const dateText = formatDisplayDateFromKey_(workDateKey);
    if (status.indexOf('มาสาย') !== -1) {
      item.late++;
      if (item.lateDates.indexOf(dateText) === -1) item.lateDates.push(dateText);
    }
    if (status === 'หยุด/ลา') {
      if (type === 'ขาดงาน') {
        item.absent++;
        if (item.absentDates.indexOf(dateText) === -1) item.absentDates.push(dateText);
      } else if (isLeaveType_(type)) {
        item.leave++;
        if (item.leaveDates.indexOf(dateText) === -1) item.leaveDates.push(dateText);
      }
    }
    summary[key] = item;
  });

  if (changeSheet && changeSheet.getLastRow() >= 2) {
    changeSheet.getDataRange().getValues().slice(1).forEach(row => {
      const workDateKey = String(row[1] || '').trim();
      const name = String(row[2] || '').trim();
      if (!workDateKey || !name) return;
      const month = workDateKey.slice(0, 6);
      const key = `${month}|${normalizeNameKey_(name)}`;
      const item = summary[key] || {
        month,
        name,
        late: 0,
        absent: 0,
        leave: 0,
        lateDates: [],
        absentDates: [],
        leaveDates: [],
        updateDates: [],
        lastUpdated: ''
      };
      const dateText = formatDisplayDateFromKey_(workDateKey);
      const changeType = String(row[6] || '').trim();
      const timestamp = String(row[11] || '').trim();
      if (item.updateDates.indexOf(dateText) === -1) item.updateDates.push(dateText);
      if (!item.lastUpdated || timestamp > item.lastUpdated) item.lastUpdated = timestamp;
      if (changeType === 'มาสาย' && item.lateDates.indexOf(dateText) === -1) item.lateDates.push(dateText);
      if (changeType === 'ขาดงาน' && item.absentDates.indexOf(dateText) === -1) item.absentDates.push(dateText);
      if (changeType === 'ลา' && item.leaveDates.indexOf(dateText) === -1) item.leaveDates.push(dateText);
      summary[key] = item;
    });
  }

  const output = Object.keys(summary).sort().map(key => {
    const item = summary[key];
    return [
      item.month,
      item.name,
      item.late,
      item.absent,
      item.leave,
      item.lateDates.join(', '),
      item.absentDates.join(', '),
      item.leaveDates.join(', '),
      item.updateDates.join(', '),
      item.lastUpdated || '-'
    ];
  });
  const lastRow = summarySheet.getLastRow();
  if (lastRow > 1) summarySheet.getRange(2, 1, lastRow - 1, SUMMARY_HEADERS.length).clearContent();
  if (output.length) summarySheet.getRange(2, 1, output.length, SUMMARY_HEADERS.length).setValues(output);
}

// ══════════════════════════════════════════════
//  DRIVE SAVE
// ══════════════════════════════════════════════
function saveImageToDrive_(base64Image, fileName) {
  const cfg = getConfig_();
  if (!cfg.DRIVE_FOLDER_ID) throw new Error('ยังไม่ได้ตั้งค่า DRIVE_FOLDER_ID');

  const bytes = Utilities.base64Decode(base64Image);
  const finalName = fileName || (formatFileTimestamp_(new Date(), cfg.TIMEZONE) + '_Smart_time_dashboard_announcement.png');
  const blob = Utilities.newBlob(bytes, 'image/png', finalName);
  const folder = DriveApp.getFolderById(cfg.DRIVE_FOLDER_ID);
  const file = folder.createFile(blob);
  return { id: file.getId(), url: file.getUrl(), name: file.getName() };
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
  const file = saveImageToDrive_(testBase64, 'test.png');
  console.log('Drive file:', file);
}
