/**
 * ACADEMIC INFORMATION DASHBOARD
 * กลุ่มบริหารงานวิชาการ โรงเรียนน้ำขุ่นวิทยา
 * สังกัดองค์การบริหารส่วนจังหวัดอุบลราชธานี
 * Google Apps Script (V8)
 *
 * 1. ตรวจสอบ SHEET_ID / FOLDER_ID ด้านล่าง
 * 2. Dashboard และระบบนักเรียนจะตรวจสอบแหล่งข้อมูลและตั้งค่าพื้นฐานอัตโนมัติ
 * 3. รัน setupSystem() จาก Apps Script Editor เมื่อต้องการเปิดใช้บัญชีผู้ดูแล/Trigger
 * 4. รายงานผลใช้สำเนา Google Docs ชั่วคราวและส่งออก PDF โดยไม่แก้เอกสารต้นฉบับ
 * 5. หลังเพิ่มการเชื่อม Google Docs ให้รัน setupSystem() อีกครั้งเพื่ออนุญาตสิทธิ์ Drive/Docs
 * 6. Deploy > New deployment > Web app > Execute as: Me
 */
const APP = Object.freeze({
  SHEET_ID: '1HU0m95sQLeUH4AXsFztqr4dEdbkw8cjZClpNg_dKelE',
  FOLDER_ID: '1S-Fglwkr7m4tBbyNXR2-n6HlKqoAhYUX',
  SCHOOL: 'โรงเรียนน้ำขุ่นวิทยา',
  DEPARTMENT: 'กลุ่มบริหารงานวิชาการ',
  AFFILIATION: 'สังกัดองค์การบริหารส่วนจังหวัดอุบลราชธานี',
  DATA_SHEET: 'STUDENTS',
  STATUS_SHEET: 'แสดงผล 0 ร มส',
  // ใช้ Google Docs ต้นฉบับเป็นแม่แบบ โดยสร้างสำเนาชั่วคราวและไม่แก้ไขเอกสารต้นฉบับ
  REPORT_TEMPLATE_DOC_ID: '1eCBuEJEZsgt_7OKYxPTHvZHJLPOdqEwphutcTiiz-DU',
  REPORT_ROWS_PER_PAGE: 12,
  QUEUE_SHEET: 'ALERT_QUEUE',
  MAX_ROWS: 15000,
  MAX_COLUMNS: 35,
  MAX_IMPORT_BYTES: 5 * 1024 * 1024,
  MAX_IMAGE_BYTES: 2 * 1024 * 1024,
  CACHE_SECONDS: 300, // Dashboard cache 5 นาที
  LOOKUP_CACHE_SECONDS: 600, // ดัชนีค้นหา 10 นาที; ล้างทันทีเมื่อแก้ไขชีตหรือกดรีเฟรช
  CACHE_CHUNK_CHARS: 25000,
  SESSION_SECONDS: 21600, // 6 hours
  ADMIN_USERNAME: 'admin',
  MIN_PASSWORD_LENGTH: 4,
  MAX_PASSWORD_LENGTH: 64,
  MAX_STATUS_SHEETS: 20,
  MAX_STATUS_ROWS: 15000,
  MAX_STATUS_COLUMNS: 80
});

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('สารสนเทศกลุ่มบริหารงานวิชาการ | โรงเรียนน้ำขุ่นวิทยา')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * รันจากตัวแก้ไขสคริปต์โดยเจ้าของระบบเท่านั้น
 * ใช้รันครั้งแรกหรือเมื่อเพิ่มสิทธิ์เชื่อมต่อ Google Docs/Drive
 */
function setupSystem() {
  assertScriptOwner_();
  const ss = SpreadsheetApp.openById(APP.SHEET_ID);
  const sourceInfo = validateAcademicSourceSheets_(ss);
  const p = PropertiesService.getScriptProperties();
  const warnings = [];

  try {
    DriveApp.getFolderById(APP.FOLDER_ID).getName();
  } catch (folderError) {
    warnings.push('โฟลเดอร์ภาพยังเข้าไม่ถึง: ' + (folderError.message || folderError));
  }

  // ตรวจสิทธิ์ Google Docs ต้นแบบโดยไม่แก้ไขเอกสารจริง
  try {
    const reportTemplate = DriveApp.getFileById(APP.REPORT_TEMPLATE_DOC_ID);
    if (reportTemplate.getMimeType() !== 'application/vnd.google-apps.document') {
      throw new Error('ไฟล์ที่กำหนดไม่ใช่ Google Docs');
    }
    const reportDoc = DocumentApp.openById(APP.REPORT_TEMPLATE_DOC_ID);
    const reportTables = reportDoc.getBody().getTables();
    if (reportTables.length !== 2 ||
        reportTables[0].getNumRows() < 3 ||
        reportTables[1].getNumRows() !== APP.REPORT_ROWS_PER_PAGE + 1 ||
        reportTables[1].getRow(0).getNumCells() !== 10) {
      throw new Error('โครงสร้างเอกสารต้นแบบไม่ตรงกับแบบรายงานที่ระบบต้องใช้');
    }
  } catch (reportError) {
    warnings.push('เชื่อมต่อ Google Docs ต้นแบบรายงานไม่ได้: ' +
      (reportError.message || reportError));
  }

  if (p.getProperty('CONFIGURED_SHEET_ID') !== APP.SHEET_ID) {
    p.deleteProperty('DATA_SHEET_NAME');
  }
  p.setProperties({
    CONFIGURED_SHEET_ID: APP.SHEET_ID,
    SOURCE_SCHEMA_VERSION: '7',
    DATA_REVISION: Utilities.getUuid().slice(0, 12),
    SYSTEM_READY: '1'
  });

  try { ensureQueueSheet_(ss); }
  catch (queueError) {
    warnings.push('สร้างชีทคิวแจ้งเตือนไม่สำเร็จ: ' + (queueError.message || queueError));
  }

  Logger.log('ชื่อผู้ใช้ผู้ดูแลระบบ: ' + APP.ADMIN_USERNAME);
  if (!p.getProperty('ADMIN_PASSWORD_HASH')) {
    const initialPassword = issueAdminPassword_();
    Logger.log('======================================================');
    Logger.log('รหัสผ่านเริ่มต้น (แสดงครั้งแรกเท่านั้น): ' + initialPassword);
    Logger.log('กรุณาคัดลอกเก็บไว้ในที่ปลอดภัย แล้วลบบันทึกหากมีการแชร์');
    Logger.log('======================================================');
  } else {
    Logger.log('มีรหัสผู้ดูแลระบบอยู่แล้ว ไม่ได้เปลี่ยนรหัสเดิม');
  }

  try { ensureQueueTrigger_(); }
  catch (triggerError) {
    warnings.push('สร้าง Trigger แจ้งเตือนไม่สำเร็จ: ' + (triggerError.message || triggerError));
  }
  try { ensureAcademicSourceEditTrigger_(ss); }
  catch (editTriggerError) {
    warnings.push('สร้าง Trigger ล้าง Cache ไม่สำเร็จ: ' + (editTriggerError.message || editTriggerError));
  }

  Logger.log('setupSystem สำเร็จ: ใช้แท็บรายชื่อนักเรียน ' + sourceInfo.classTabs +
    ' ห้อง และชีท ' + APP.STATUS_SHEET);
  warnings.forEach(w => Logger.log('คำเตือน: ' + w));
  return warnings.length
    ? 'setupSystem สำเร็จ แต่มีคำเตือน: ' + warnings.join(' | ')
    : 'setupSystem สำเร็จ';
}

/** ใช้กรณีลืมรหัส: รันจาก Apps Script Editor เท่านั้น */
function resetAdminPassword() {
  assertScriptOwner_();
  const pw = issueAdminPassword_();
  Logger.log('ชื่อผู้ใช้: ' + APP.ADMIN_USERNAME);
  Logger.log('รหัสผ่านใหม่: ' + pw);
  Logger.log('รหัสผ่านเดิมและเซสชันเดิมถูกยกเลิก');
  return 'สร้างรหัสใหม่แล้ว ดูที่ Execution log';
}

/**
 * ป้องกันไม่ให้ผู้ชม Web App เรียก setupSystem/resetAdminPassword ผ่าน Console
 * ในการรันจาก Apps Script Editor ผู้ใช้ที่สั่งรันต้องเป็นเจ้าของสคริปต์
 */
function assertScriptOwner_() {
  const active = Session.getActiveUser().getEmail();
  const owner = Session.getEffectiveUser().getEmail();
  if (!active || !owner || active.toLowerCase() !== owner.toLowerCase()) {
    throw new Error('ฟังก์ชันนี้อนุญาตให้เจ้าของสคริปต์เรียกจาก Apps Script Editor เท่านั้น');
  }
}

function issueAdminPassword_() {
  const p = PropertiesService.getScriptProperties();
  const password = Utilities.getUuid().replace(/-/g, '').slice(0, 12);
  const salt = Utilities.getUuid().replace(/-/g, '');
  p.setProperties({
    ADMIN_PASSWORD_SALT: salt,
    ADMIN_PASSWORD_HASH: digest_(salt + ':' + password),
    ADMIN_AUTH_VERSION: Utilities.getUuid()
  });
  return password;
}

function digest_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s)
    .map(b => ('0' + (b & 255).toString(16)).slice(-2)).join('');
}

/** เข้าสู่ระบบด้วย username admin และรหัสผ่านที่ตั้งไว้ */
function loginAdmin(username, password) {
  const p = PropertiesService.getScriptProperties();
  const salt = p.getProperty('ADMIN_PASSWORD_SALT');
  const expected = p.getProperty('ADMIN_PASSWORD_HASH');
  if (!salt || !expected) throw new Error('ยังไม่ได้ติดตั้งระบบ กรุณารัน setupSystem()');

  // ลดความเสี่ยงการสุ่มรหัสผ่าน โดยเฉพาะเมื่อผู้ดูแลเลือกรหัสผ่านเพียง 4 ตัว
  const cache = CacheService.getScriptCache();
  if (cache.get('ADMIN_LOGIN_LOCKED') === '1') {
    throw new Error('เข้าสู่ระบบผิดหลายครั้ง กรุณาลองใหม่หลังจาก 5 นาที');
  }
  const validUser = username === APP.ADMIN_USERNAME;
  const validPassword = typeof password === 'string' &&
    password.length >= APP.MIN_PASSWORD_LENGTH &&
    password.length <= APP.MAX_PASSWORD_LENGTH &&
    digest_(salt + ':' + password) === expected;
  if (!validUser || !validPassword) {
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const attempts = Number(cache.get('ADMIN_LOGIN_FAILS') || 0) + 1;
      if (attempts >= 5) {
        cache.put('ADMIN_LOGIN_LOCKED', '1', 300);
        cache.remove('ADMIN_LOGIN_FAILS');
        throw new Error('เข้าสู่ระบบผิดครบ 5 ครั้ง กรุณาลองใหม่หลังจาก 5 นาที');
      }
      cache.put('ADMIN_LOGIN_FAILS', String(attempts), 300);
    } finally { lock.releaseLock(); }
    throw new Error('ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง');
  }
  cache.remove('ADMIN_LOGIN_FAILS');
  const token = Utilities.getUuid() + Utilities.getUuid();
  cache.put('AUTH_' + digest_(token),
    p.getProperty('ADMIN_AUTH_VERSION'), APP.SESSION_SECONDS);
  return { success: true, username: APP.ADMIN_USERNAME, token: token, expiresIn: APP.SESSION_SECONDS };
}

/** เปลี่ยนรหัสผ่านได้จากเมนูความปลอดภัย: ขั้นต่ำ 4 สูงสุด 64 ตัวอักษร */
function changeAdminPassword(sessionToken, oldPassword, newPassword) {
  requireAdmin_(sessionToken);
  if (typeof oldPassword !== 'string' ||
      typeof newPassword !== 'string' ||
      newPassword.length < APP.MIN_PASSWORD_LENGTH ||
      newPassword.length > APP.MAX_PASSWORD_LENGTH) {
    throw new Error('รหัสผ่านใหม่ต้องมีความยาว 4–64 ตัวอักษร');
  }
  if (newPassword === oldPassword) throw new Error('กรุณาตั้งรหัสผ่านใหม่ที่แตกต่างจากรหัสเดิม');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const p = PropertiesService.getScriptProperties();
    const salt = p.getProperty('ADMIN_PASSWORD_SALT');
    const expected = p.getProperty('ADMIN_PASSWORD_HASH');
    if (!salt || !expected || digest_(salt + ':' + oldPassword) !== expected) {
      throw new Error('รหัสผ่านปัจจุบันไม่ถูกต้อง');
    }
    const newSalt = Utilities.getUuid().replace(/-/g, '');
    p.setProperties({
      ADMIN_PASSWORD_SALT: newSalt,
      ADMIN_PASSWORD_HASH: digest_(newSalt + ':' + newPassword),
      ADMIN_AUTH_VERSION: Utilities.getUuid()
    });
    const cache = CacheService.getScriptCache();
    cache.remove('AUTH_' + digest_(sessionToken));
    cache.remove('ADMIN_LOGIN_FAILS');
    cache.remove('ADMIN_LOGIN_LOCKED');
    return { success: true, message: 'เปลี่ยนรหัสผ่านเรียบร้อย กรุณาเข้าสู่ระบบใหม่' };
  } finally { lock.releaseLock(); }
}

function requireAdmin_(token) {
  if (typeof token !== 'string' || token.length !== 72) {
    throw new Error('กรุณาเข้าสู่ระบบผู้ดูแลก่อนดำเนินการ');
  }
  const p = PropertiesService.getScriptProperties();
  const version = CacheService.getScriptCache().get('AUTH_' + digest_(token));
  if (!version || version !== p.getProperty('ADMIN_AUTH_VERSION')) {
    throw new Error('เซสชันหมดอายุ กรุณาเข้าสู่ระบบใหม่');
  }
}

function logoutAdmin(token) {
  if (typeof token === 'string') {
    CacheService.getScriptCache().remove('AUTH_' + digest_(token));
  }
  return { success: true };
}

function getAdminSettings(token) {
  requireAdmin_(token);
  const p = PropertiesService.getScriptProperties();
  return {
    success: true,
    telegramReady: Boolean(p.getProperty('TELEGRAM_BOT_TOKEN') && p.getProperty('TELEGRAM_CHAT_ID')),
    lineReady: Boolean(p.getProperty('LINE_CHANNEL_ACCESS_TOKEN') && p.getProperty('LINE_TO')),
    telegramChatId: p.getProperty('TELEGRAM_CHAT_ID') || '',
    lineTo: p.getProperty('LINE_TO') || ''
  };
}

/** ไม่ส่ง token / secret ใด ๆ กลับไปยังเว็บ */
function saveNotificationSettings(sessionToken, settings) {
  requireAdmin_(sessionToken);
  settings = settings || {};
  const p = PropertiesService.getScriptProperties();
  const updates = {};
  if (settings.telegramToken) updates.TELEGRAM_BOT_TOKEN = validateText_(settings.telegramToken, 300);
  if (settings.telegramChatId) updates.TELEGRAM_CHAT_ID = validateText_(settings.telegramChatId, 120);
  if (settings.lineToken) updates.LINE_CHANNEL_ACCESS_TOKEN = validateText_(settings.lineToken, 700);
  if (settings.lineTo) updates.LINE_TO = validateText_(settings.lineTo, 120);
  if (Object.keys(updates).length) p.setProperties(updates);
  if (settings.clearTelegram === true) {
    p.deleteProperty('TELEGRAM_BOT_TOKEN');
    p.deleteProperty('TELEGRAM_CHAT_ID');
  }
  if (settings.clearLine === true) {
    p.deleteProperty('LINE_CHANNEL_ACCESS_TOKEN');
    p.deleteProperty('LINE_TO');
  }
  return getAdminSettings(sessionToken);
}

function validateText_(v, max) {
  const s = String(v).trim();
  if (!s || s.length > max) throw new Error('รูปแบบข้อมูลการตั้งค่าไม่ถูกต้อง');
  return s;
}

function getDashboardData(forceRefresh) {
  const started = Date.now();
  try {
    // ตรวจสถานะระบบก่อน แต่ไม่เปิด Spreadsheet จนกว่าจะ Cache miss
    // ลดเวลาเปิด Dashboard ในกรณีที่มีข้อมูลอยู่ใน Cache แล้ว
    let ss = ensureAcademicSetup_();
    const props = PropertiesService.getScriptProperties();
    const rev = props.getProperty('DATA_REVISION') || '0';
    const key = 'DASH_V10_GRADE_GENDER_0RMS_' + rev;
    let cache = null;
    try { cache = CacheService.getScriptCache(); } catch (cacheInitError) {
      console.warn('Dashboard cache unavailable: ' + (cacheInitError.message || cacheInitError));
    }

    if (forceRefresh === true) {
      try { clearStatusLookupCache_(rev); } catch (clearError) {
        console.warn('Lookup cache clear skipped: ' + (clearError.message || clearError));
      }
      if (cache) {
        try { cache.remove(key); } catch (removeError) {}
      }
    } else if (cache) {
      try {
        const stored = cache.get(key);
        if (stored) {
          const cachedResult = JSON.parse(stored);
          if (cachedResult && cachedResult.success === true) {
            cachedResult.serverTime = new Date().toISOString();
            cachedResult.cached = true;
            cachedResult.elapsedMs = Date.now() - started;
            return cachedResult;
          }
        }
      } catch (cacheError) {
        console.warn('Dashboard cache read skipped: ' + (cacheError.message || cacheError));
        try { cache.remove(key); } catch (removeError) {}
      }
    }

    // อ่านด้วย batchGet ก่อน; เปิด SpreadsheetApp เฉพาะกรณี API ไม่พร้อมและต้อง fallback
    const result = buildAcademicDashboardData_(ss);
    result.success = true;
    result.school = APP.SCHOOL;
    result.department = APP.DEPARTMENT;
    result.affiliation = APP.AFFILIATION;
    result.serverTime = new Date().toISOString();
    result.cached = false;
    result.elapsedMs = Date.now() - started;

    if (cache) {
      try { cache.put(key, JSON.stringify(result), APP.CACHE_SECONDS); }
      catch (cacheError) {
        console.warn('Dashboard cache write skipped: ' + (cacheError.message || cacheError));
      }
    }
    return result;
  } catch (e) {
    console.error('getDashboardData: ' + (e.stack || e));
    return { success: false, error: e.message || String(e), elapsedMs: Date.now() - started };
  }
}

/**
 * V6: สรุปผลเป็นระดับชั้น ม.1–ม.6 และแยกเพศจากคำนำหน้าชื่อ
 * จำนวนติด 0/ร/มส ในกราฟนับเป็นจำนวนนักเรียนไม่ซ้ำในแต่ละสถานะ
 * ส่วน KPI totalStatus ยังคงนับเป็นจำนวนรายการรายวิชา
 */
function buildAcademicDashboardData_(ss) {
  const source = getStatusLookupData_(ss);
  const students = source.students || [];
  const statusByStudent = source.statusByStudent || {};
  const statusRecordCounts = source.statusRecordCounts || { zero: 0, r: 0, ms: 0 };
  const statusRecordsByGrade = source.statusRecordsByGrade || {};
  const grades = {};
  const rooms = {};
  const summary = {
    totalLevels: 6, totalRooms: 30, grandTotal: students.length,
    totalMale: 0, totalFemale: 0, totalUnknown: 0,
    zero: Number(statusRecordCounts.zero) || 0,
    r: Number(statusRecordCounts.r) || 0,
    ms: Number(statusRecordCounts.ms) || 0,
    totalStatus: 0,
    zeroStudents: 0, rStudents: 0, msStudents: 0,
    studentsWithStatus: 0, statusMale: 0, statusFemale: 0, statusUnknown: 0,
    studentsWithBehaviorScore: students.filter(x => x.hasBehaviorScore).length,
    studentsWithoutCitizenId: source.studentsWithoutCitizenId || 0,
    unmatchedStatusRecords: Number(source.unmatchedStatusRecords) || 0
  };

  for (let level = 1; level <= 6; level++) {
    const grade = 'ม.' + level;
    const recordCounts = statusRecordsByGrade[grade] || {};
    grades[grade] = {
      grade, level,
      male: 0, female: 0, unknown: 0, total: 0,
      behaviorScored: 0,
      statusMale: 0, statusFemale: 0, statusUnknown: 0, statusTotal: 0,
      zeroMale: 0, zeroFemale: 0, zeroUnknown: 0,
      rMale: 0, rFemale: 0, rUnknown: 0,
      msMale: 0, msFemale: 0, msUnknown: 0,
      zeroRecords: Number(recordCounts.zeroRecords) || 0,
      rRecords: Number(recordCounts.rRecords) || 0,
      msRecords: Number(recordCounts.msRecords) || 0
    };
  }

  getTargetClassNames_().forEach(className => {
    const cls = parseTargetClass_(className, '');
    rooms[className] = {
      grade: cls.grade, room: cls.room, label: className,
      male: 0, female: 0, unknown: 0, total: 0,
      behaviorScored: 0
    };
  });

  const rosterByStudentKey = Object.create(null);
  students.forEach(student => {
    const grade = grades[student.grade];
    const room = rooms[student.classLabel];
    if (!grade || !room) return;

    const gender = student.gender || 'unknown';
    grade[gender]++;
    grade.total++;
    room[gender]++;
    room.total++;

    if (student.hasBehaviorScore) {
      grade.behaviorScored++;
      room.behaviorScored++;
    }

    const studentKey = student.classLabel + '|' + (student.studentNo || student.normName);
    rosterByStudentKey[studentKey] = student;
  });

  // ใช้รายการสถานะที่จับคู่กับนักเรียนแล้วจากดัชนีร่วม
  // ลดการเก็บ statusRecords ซ้ำสองชุดใน Cache
  Object.keys(statusByStudent).forEach(studentKey => {
    const student = rosterByStudentKey[studentKey];
    if (!student) return;

    const grade = grades[student.grade];
    if (!grade) return;

    const gender = student.gender || 'unknown';
    const records = statusByStudent[studentKey] || [];
    const types = Object.create(null);
    records.forEach(record => { types[record.type] = true; });

    grade.statusTotal++;
    if (gender === 'male') grade.statusMale++;
    else if (gender === 'female') grade.statusFemale++;
    else grade.statusUnknown++;

    Object.keys(types).forEach(type => {
      const field = type + (gender === 'male' ? 'Male' : gender === 'female' ? 'Female' : 'Unknown');
      if (Object.prototype.hasOwnProperty.call(grade, field)) grade[field]++;
    });
  });

  const gradeData = Object.keys(grades)
    .sort((a, b) => grades[a].level - grades[b].level)
    .map(key => {
      const item = grades[key];
      item.statusTotal = item.statusMale + item.statusFemale + item.statusUnknown;
      item.zeroStudents = item.zeroMale + item.zeroFemale + item.zeroUnknown;
      item.rStudents = item.rMale + item.rFemale + item.rUnknown;
      item.msStudents = item.msMale + item.msFemale + item.msUnknown;
      return item;
    });

  const studentData = Object.keys(rooms).map(key => rooms[key]).sort((a, b) => {
    const ca = parseTargetClass_(a.label, ''), cb = parseTargetClass_(b.label, '');
    return (Number(ca.grade.match(/\d+/)[0]) - Number(cb.grade.match(/\d+/)[0])) ||
      (Number(ca.room) - Number(cb.room));
  });

  summary.totalMale = gradeData.reduce((n, x) => n + x.male, 0);
  summary.totalFemale = gradeData.reduce((n, x) => n + x.female, 0);
  summary.totalUnknown = gradeData.reduce((n, x) => n + x.unknown, 0);
  summary.grandTotal = summary.totalMale + summary.totalFemale + summary.totalUnknown;
  summary.totalStatus = summary.zero + summary.r + summary.ms;
  summary.studentsWithStatus = Object.keys(statusByStudent).length;
  summary.statusMale = gradeData.reduce((n, x) => n + x.statusMale, 0);
  summary.statusFemale = gradeData.reduce((n, x) => n + x.statusFemale, 0);
  summary.statusUnknown = gradeData.reduce((n, x) => n + x.statusUnknown, 0);
  summary.zeroStudents = gradeData.reduce((n, x) => n + x.zeroStudents, 0);
  summary.rStudents = gradeData.reduce((n, x) => n + x.rStudents, 0);
  summary.msStudents = gradeData.reduce((n, x) => n + x.msStudents, 0);

  return {
    gradeData,
    studentData,
    summary,
    source: {
      classTabs: 30,
      statusSheet: APP.STATUS_SHEET,
      genderSource: 'คำนำหน้าชื่อในรายชื่อนักเรียน',
      statusUnit: 'จำนวนรายการรายวิชาที่ยังไม่แก้ไข',
      statusStudentUnit: 'จำนวนนักเรียนไม่ซ้ำในแต่ละสถานะ',
      readMode: source.readMode || 'SpreadsheetApp'
    }
  };
}

/**
 * สร้าง/อ่าน cache ข้อมูลสำหรับ Dashboard และการตรวจสอบสถานะร่วมกัน
 * Cache และ Lock เป็นตัวเร่งความเร็ว ไม่ใช่เงื่อนไขบังคับในการให้บริการ
 */
function getStatusLookupData_(existingSpreadsheet) {
  const props = PropertiesService.getScriptProperties();
  const revision = props.getProperty('DATA_REVISION') || '0';
  const baseKey = 'STATUS_LOOKUP_V10_' + revision;
  let cache = null;
  try { cache = CacheService.getScriptCache(); } catch (e) {
    console.warn('Lookup cache unavailable: ' + (e.message || e));
  }

  if (cache) {
    const cached = readStatusLookupCache_(cache, baseKey);
    if (cached) return cached;
  }

  const lock = LockService.getScriptLock();
  let locked = false;
  try { locked = lock.tryLock(5000); } catch (lockError) {
    console.warn('Lookup lock skipped: ' + (lockError.message || lockError));
  }

  try {
    if (cache) {
      const cachedAgain = readStatusLookupCache_(cache, baseKey);
      if (cachedAgain) return cachedAgain;
    }

    const batchRows = readTargetSourceRowsBatch_(existingSpreadsheet || null);
    const ss = existingSpreadsheet || batchRows.spreadsheet || null;
    const roster = readTargetRoster_(ss, batchRows);
    const statusRecords = readTargetStatusRecords_(ss, batchRows);

    const safeStudents = roster.students.map(s => ({
      citizenHash: s.citizen ? digest_('CITIZEN:' + s.citizen) : '',
      citizenId: s.citizen || '',
      name: s.name,
      normName: s.normName,
      studentNo: s.studentNo,
      classLabel: s.classLabel,
      grade: s.grade,
      room: s.room,
      gender: s.gender,
      behaviorScore: s.behaviorScore,
      hasBehaviorScore: s.hasBehaviorScore
    }));

    const indexes = buildStatusLookupIndexes_(safeStudents, statusRecords);
    const payload = {
      students: safeStudents,
      studentIndexByCitizenHash: indexes.studentIndexByCitizenHash,
      studentIndexByName: indexes.studentIndexByName,
      statusByStudent: indexes.statusByStudent,
      statusRecordCounts: indexes.statusRecordCounts,
      statusRecordsByGrade: indexes.statusRecordsByGrade,
      unmatchedStatusRecords: indexes.unmatchedStatusRecords,
      studentsWithoutCitizenId: roster.studentsWithoutCitizenId,
      readMode: batchRows.readMode || 'SpreadsheetApp direct reads'
    };

    // เขียน Cache แบบ best-effort; หากข้อมูลใหญ่เกินโควตาให้คืนข้อมูลจริงทันที
    if (cache) {
      try {
        const serialized = JSON.stringify(payload);
        const chunkSize = APP.CACHE_CHUNK_CHARS;
        const chunks = [];
        for (let i = 0; i < serialized.length; i += chunkSize) {
          chunks.push(serialized.slice(i, i + chunkSize));
        }

        const maxChunks = 80;
        if (chunks.length > 0 && chunks.length <= maxChunks) {
          const cacheValues = {};
          chunks.forEach((chunk, i) => cacheValues[baseKey + '_' + i] = chunk);
          cache.remove(baseKey + '_META');
          cache.putAll(cacheValues, APP.LOOKUP_CACHE_SECONDS);
          cache.put(baseKey + '_META', String(chunks.length), APP.LOOKUP_CACHE_SECONDS);
        } else {
          cache.remove(baseKey + '_META');
          console.warn('Lookup cache skipped: ' + chunks.length + ' chunks');
        }
      } catch (cacheError) {
        try { cache.remove(baseKey + '_META'); } catch (removeError) {}
        console.warn('Lookup cache write skipped: ' + (cacheError.message || cacheError));
      }
    }

    return payload;
  } finally {
    if (locked) {
      try { lock.releaseLock(); } catch (releaseError) {}
    }
  }
}

function readStatusLookupCache_(cache, baseKey) {
  try {
    const meta = cache.get(baseKey + '_META');
    if (!meta) return null;
    const count = Number(meta);
    if (!(count > 0 && count <= 80)) return null;
    const keys = Array.from({ length: count }, (_, i) => baseKey + '_' + i);
    const values = cache.getAll(keys);
    if (!values || !keys.every(key => typeof values[key] === 'string')) return null;
    const json = keys.map(key => values[key]).join('');
    return JSON.parse(json);
  } catch (e) {
    console.warn('Lookup cache read skipped: ' + (e.message || e));
    return null;
  }
}

function buildStatusLookupIndexes_(students, statusRecords) {
  const studentIndexByCitizenHash = Object.create(null);
  const studentIndexByName = Object.create(null);
  const rosterByClassNo = Object.create(null);
  const rosterByClassName = Object.create(null);

  students.forEach((student, index) => {
    if (student.citizenHash) {
      if (Object.prototype.hasOwnProperty.call(studentIndexByCitizenHash, student.citizenHash)) {
        studentIndexByCitizenHash[student.citizenHash] = -1;
      } else {
        studentIndexByCitizenHash[student.citizenHash] = index;
      }
    }
    if (student.normName) {
      if (Object.prototype.hasOwnProperty.call(studentIndexByName, student.normName)) {
        studentIndexByName[student.normName] = -1;
      } else {
        studentIndexByName[student.normName] = index;
      }
      const classNameKey = student.classLabel + '|' + student.normName;
      if (Object.prototype.hasOwnProperty.call(rosterByClassName, classNameKey)) {
        rosterByClassName[classNameKey] = null; // ชื่อซ้ำในห้องเดียวกัน: ไม่จับคู่ด้วยชื่อเพียงอย่างเดียว
      } else {
        rosterByClassName[classNameKey] = student;
      }
    }
    if (student.studentNo) rosterByClassNo[student.classLabel + '|' + student.studentNo] = student;
  });

  const statusByStudent = Object.create(null);
  const statusRecordCounts = { zero: 0, r: 0, ms: 0 };
  const statusRecordsByGrade = {};
  for (let level = 1; level <= 6; level++) {
    statusRecordsByGrade['ม.' + level] = {
      zeroRecords: 0, rRecords: 0, msRecords: 0
    };
  }
  let unmatchedStatusRecords = 0;

  (statusRecords || []).forEach(record => {
    if (Object.prototype.hasOwnProperty.call(statusRecordCounts, record.type)) {
      statusRecordCounts[record.type]++;
    }

    const cls = parseTargetClass_(record.classLabel, '');
    if (cls && statusRecordsByGrade[cls.grade]) {
      const field = record.type + 'Records';
      if (Object.prototype.hasOwnProperty.call(statusRecordsByGrade[cls.grade], field)) {
        statusRecordsByGrade[cls.grade][field]++;
      }
    }

    let student = record.studentNo
      ? rosterByClassNo[record.classLabel + '|' + record.studentNo]
      : null;

    if (!student && record.normName) {
      const byName = rosterByClassName[record.classLabel + '|' + record.normName];
      if (byName) student = byName;
    }

    if (!student) {
      unmatchedStatusRecords++;
      return;
    }

    const key = student.classLabel + '|' + (student.studentNo || student.normName);
    if (!statusByStudent[key]) statusByStudent[key] = [];
    // เก็บเฉพาะฟิลด์ที่ใช้แสดงในรายงาน ลดขนาด Cache และเวลา JSON.parse
    statusByStudent[key].push({
      type: record.type,
      subject: record.subject || '',
      subjectCode: record.subjectCode || '',
      teacher: record.teacher || '',
      year: record.year || '',
      term: record.term || '',
      editStatus: record.editStatus || ''
    });
  });

  return {
    studentIndexByCitizenHash,
    studentIndexByName,
    statusByStudent,
    statusRecordCounts,
    statusRecordsByGrade,
    unmatchedStatusRecords
  };
}

/**
 * อ่านข้อมูล 30 ห้องเรียนและชีทสถานะด้วย batchGet 1 คำขอ
 * หาก Sheets API ไม่พร้อม จะ fallback เป็น SpreadsheetApp และอ่านตาม getLastRow()
 */
function readTargetSourceRowsBatch_(ss) {
  const classNames = getTargetClassNames_();
  const names = classNames.concat([APP.STATUS_SHEET]);
  let cache = null;
  try { cache = CacheService.getScriptCache(); } catch (e) {}

  // พยายามอ่าน 30 ห้อง + ชีทสถานะด้วย API เพียง 1 คำขอ
  // หาก API/สิทธิ์ใช้งานไม่ได้ จะสลับไปอ่านด้วย SpreadsheetApp โดยอัตโนมัติ
  let batchDisabled = false;
  if (cache) {
    try { batchDisabled = cache.get('SHEETS_BATCHGET_DISABLED') === '1'; } catch (cacheReadError) {}
  }
  if (!batchDisabled) {
    try {
      const ranges = names.map(name => {
        const quoted = "'" + String(name).replace(/'/g, "''") + "'";
        return quoted + (name === APP.STATUS_SHEET
          ? '!A1:L' + APP.MAX_STATUS_ROWS
          : '!A1:F' + APP.MAX_ROWS);
      });
      const query = ranges.map(range => 'ranges=' + encodeURIComponent(range)).join('&') +
        '&valueRenderOption=FORMATTED_VALUE&majorDimension=ROWS' +
        '&fields=' + encodeURIComponent('valueRanges(range,values)');
      const url = 'https://sheets.googleapis.com/v4/spreadsheets/' +
        encodeURIComponent(APP.SHEET_ID) + '/values:batchGet?' + query;
      const response = UrlFetchApp.fetch(url, {
        method: 'get',
        headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
        muteHttpExceptions: true
      });
      const code = response.getResponseCode();
      if (code < 200 || code >= 300) {
        throw new Error('Sheets API batchGet HTTP ' + code + ': ' +
          response.getContentText().slice(0, 180));
      }
      const result = JSON.parse(response.getContentText() || '{}');
      if (!Array.isArray(result.valueRanges) || result.valueRanges.length !== names.length) {
        throw new Error('Sheets API ส่งข้อมูลกลับมาไม่ครบทุกแท็บ');
      }
      const classRows = Object.create(null);
      classNames.forEach((name, index) => {
        classRows[name] = result.valueRanges[index].values || [];
      });
      if (cache) {
        try { cache.remove('SHEETS_BATCHGET_DISABLED'); } catch (removeError) {}
      }
      return {
        classRows,
        statusRows: result.valueRanges[names.length - 1].values || [],
        batch: true,
        readMode: 'Sheets API batchGet (1 request)'
      };
    } catch (error) {
      // จำสถานะชั่วคราว ป้องกันการลอง API ที่ยังไม่พร้อมซ้ำทุกครั้ง
      if (cache) {
        try { cache.put('SHEETS_BATCHGET_DISABLED', '1', 300); } catch (cacheError) {}
      }
      console.warn('Sheets API batchGet fallback: ' + (error.message || error));
    }
  }

  // Fallback ที่คงผลลัพธ์เดิม ใช้เฉพาะแถวที่มีข้อมูลจริงตาม getLastRow()
  // เปิด Spreadsheet เฉพาะเมื่อ batchGet ใช้งานไม่ได้
  if (!ss) ss = SpreadsheetApp.openById(APP.SHEET_ID);
  const classRows = Object.create(null);
  classNames.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) throw new Error('ไม่พบแท็บข้อมูล ' + name);
    const lastRow = Math.max(3, Math.min(sh.getLastRow() || 3, APP.MAX_ROWS));
    classRows[name] = sh.getRange(1, 1, lastRow, 6).getDisplayValues();
  });

  const statusSheet = ss.getSheetByName(APP.STATUS_SHEET);
  if (!statusSheet) throw new Error('ไม่พบชีท ' + APP.STATUS_SHEET);
  const statusLastRow = Math.max(5, Math.min(statusSheet.getLastRow() || 5, APP.MAX_STATUS_ROWS));
  const statusRows = statusSheet.getRange(1, 1, statusLastRow, 12).getDisplayValues();
  return {
    classRows,
    statusRows,
    batch: false,
    spreadsheet: ss,
    readMode: 'SpreadsheetApp direct reads (fallback)'
  };
}

function clearStatusLookupCache_(revisionOverride) {
  try {
    const props = PropertiesService.getScriptProperties();
    const revision = revisionOverride || props.getProperty('DATA_REVISION') || '0';
    const baseKey = 'STATUS_LOOKUP_V10_' + revision;
    const cache = CacheService.getScriptCache();
    const meta = cache.get(baseKey + '_META');
    if (meta) {
      const count = Number(meta);
      const keys = [baseKey + '_META'];
      for (let i = 0; i < count && i < 80; i++) keys.push(baseKey + '_' + i);
      cache.removeAll(keys);
    }
  } catch (e) {
    console.warn('Lookup cache clear skipped: ' + (e.message || e));
  }
}

/** เรียกหลังแก้ข้อมูลต้นทาง เพื่อเปลี่ยน revision และบังคับให้สร้าง cache ใหม่ */
function invalidateAcademicDataCache_() {
  const props = PropertiesService.getScriptProperties();
  const oldRevision = props.getProperty('DATA_REVISION') || '0';
  clearStatusLookupCache_(oldRevision);
  props.setProperty('DATA_REVISION', Utilities.getUuid().slice(0, 12));
}

function ensureAcademicSourceEditTrigger_(ss) {
  const exists = ScriptApp.getProjectTriggers().some(trigger =>
    trigger.getHandlerFunction() === 'handleAcademicSourceEdit_');
  if (!exists) {
    ScriptApp.newTrigger('handleAcademicSourceEdit_').forSpreadsheet(ss).onEdit().create();
  }
}

/** ล้าง cache อัตโนมัติเมื่อมีการแก้ข้อมูลใน 30 ห้องเรียนหรือชีท 0/ร/มส */
function handleAcademicSourceEdit_(event) {
  if (!event || !event.range) return;
  const sheetName = event.range.getSheet().getName();
  if (sheetName !== APP.STATUS_SHEET && getTargetClassNames_().indexOf(sheetName) < 0) return;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return;
  try { invalidateAcademicDataCache_(); }
  finally { lock.releaseLock(); }
}

/** ทดสอบการอ่านข้อมูลจากแท็บห้องเรียนผ่าน SpreadsheetApp */
function testFastRead() {
  assertScriptOwner_();
  const started = Date.now();
  const ss = SpreadsheetApp.openById(APP.SHEET_ID);
  const rows = readTargetSourceRowsBatch_(ss);
  const elapsed = Date.now() - started;
  const result = {
    success: Boolean(rows),
    mode: rows ? rows.readMode : 'อ่านข้อมูลไม่สำเร็จ',
    elapsedMs: elapsed,
    classTabs: getTargetClassNames_().length,
    statusRows: rows && rows.statusRows ? rows.statusRows.length : null
  };
  Logger.log(JSON.stringify(result));
  return result;
}

/**
 * ตรวจสอบแหล่งข้อมูลจริงจาก Apps Script Editor โดยไม่แก้ไขข้อมูลในชีต
 * รายงานจำนวนข้อมูล วิธีอ่าน และเวลาที่ใช้ เพื่อแยกปัญหาการเชื่อมต่อออกจากปัญหาหน้าเว็บ
 */
function diagnoseDashboard() {
  assertScriptOwner_();
  const started = Date.now();
  try {
    const ss = SpreadsheetApp.openById(APP.SHEET_ID);
    const validation = validateAcademicSourceSheets_(ss);
    const batchRows = readTargetSourceRowsBatch_(ss);
    const roster = readTargetRoster_(ss, batchRows);
    const statusRecords = readTargetStatusRecords_(ss, batchRows);
    return {
      success: true,
      spreadsheetName: ss.getName(),
      classTabs: validation.classTabs,
      statusSheet: APP.STATUS_SHEET,
      readMode: batchRows.readMode || 'SpreadsheetApp direct reads',
      students: roster.students.length,
      studentsWithoutCitizenId: roster.studentsWithoutCitizenId,
      statusRecords: statusRecords.length,
      elapsedMs: Date.now() - started,
      setupReady: PropertiesService.getScriptProperties().getProperty('SOURCE_SCHEMA_VERSION') === '7'
    };
  } catch (e) {
    return {
      success: false,
      error: e.message || String(e),
      elapsedMs: Date.now() - started
    };
  }
}

function getTargetClassNames_() {
  const names = [];
  for (let grade = 1; grade <= 6; grade++) {
    for (let room = 1; room <= 5; room++) names.push('ม.' + grade + '/' + room);
  }
  return names;
}

function validateAcademicSourceSheets_(ss) {
  const missing = [];
  const invalid = [];
  getTargetClassNames_().forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) { missing.push(name); return; }
    const lastRow = Math.max(3, sh.getLastRow());
    const preview = sh.getRange(1, 1, Math.min(lastRow, 12), 6).getDisplayValues();
    const schema = findClassHeader_(preview);
    if (!schema || schema.name < 0 || schema.citizen < 0 || schema.behavior < 0) invalid.push(name);
  });
  const statusSheet = ss.getSheetByName(APP.STATUS_SHEET);
  let statusSchema = null;
  if (statusSheet) {
    const preview = statusSheet.getRange(1, 1, Math.min(Math.max(statusSheet.getLastRow(), 1), 15),
      Math.min(Math.max(statusSheet.getLastColumn(), 1), 20)).getDisplayValues();
    statusSchema = findStatusHeader_(preview);
  }
  if (missing.length) throw new Error('ไม่พบแท็บห้องเรียน: ' + missing.join(', '));
  if (invalid.length) throw new Error('หัวตารางชื่อ–สกุล / เลขประชาชน / คะแนนพฤติกรรมไม่ครบในแท็บ: ' + invalid.join(', '));
  if (!statusSheet || !statusSchema) throw new Error('ไม่พบหัวตารางที่จำเป็นในชีท ' + APP.STATUS_SHEET);
  return { classTabs: 30, statusSheet: APP.STATUS_SHEET };
}

/**
 * ตรวจสอบสิทธิ์และโครงสร้างข้อมูลแบบ Lazy initialization
 * ทำงานอัตโนมัติเมื่อเปิด Dashboard/ตรวจสอบนักเรียนครั้งแรก
 * ไม่สร้าง Trigger และไม่แก้ไขข้อมูลใน Google Sheets
 */
function ensureAcademicSetup_() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('CONFIGURED_SHEET_ID') === APP.SHEET_ID &&
      props.getProperty('SOURCE_SCHEMA_VERSION') === '7') {
    // ระบบพร้อมแล้ว: ไม่เปิด Spreadsheet ซ้ำทุกครั้ง
    return null;
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    // ตรวจซ้ำหลังได้ Lock ป้องกันการตั้งค่าพร้อมกันหลายคำขอ
    if (props.getProperty('CONFIGURED_SHEET_ID') === APP.SHEET_ID &&
        props.getProperty('SOURCE_SCHEMA_VERSION') === '7') {
      return null;
    }

    const ss = SpreadsheetApp.openById(APP.SHEET_ID);
    validateAcademicSourceSheets_(ss);
    props.setProperties({
      CONFIGURED_SHEET_ID: APP.SHEET_ID,
      SOURCE_SCHEMA_VERSION: '7',
      DATA_REVISION: Utilities.getUuid().slice(0, 12),
      SYSTEM_READY: '1'
    });
    return ss; // คืน Spreadsheet เฉพาะครั้งที่ต้องตั้งค่าครั้งแรก
  } finally {
    lock.releaseLock();
  }
}

function normalizeHeaderKey_(value) {
  return normalizeThaiDigits_(String(value || '')).replace(/[\s\-_.()\/]+/g, '').toLowerCase();
}

/** ค้นหาหัวตารางในแท็บห้องเรียน (รองรับความต่างของบางห้อง เช่น ไม่มีคอลัมน์สังกัดสี) */
function findClassHeader_(rows) {
  for (let r = 0; r < Math.min(rows.length, 12); r++) {
    const h = (rows[r] || []).map(normalizeHeaderKey_);
    const col = fn => h.findIndex(fn);
    const name = col(x => /^(ชื่อสกุล|ชื่อสกุลนักเรียน|ชื่อนักเรียน|ชื่อนามสกุล|ชื่อและนามสกุล)$/.test(x));
    const citizen = col(x => /^(เลขประชาชน|เลขบัตรประชาชน|เลขประจำตัวประชาชน|เลชประชาชน|เลขบัตรปชช)$/.test(x));
    const no = col(x => /^(ที่|เลขที่|ลำดับ)$/.test(x));
    if (name < 0 || (citizen < 0 && no < 0)) continue;
    return {
      headerRow: r,
      no,
      studentId: col(x => /^(เลขประจำตัว|รหัสนักเรียน|เลขรหัสนักเรียน)$/.test(x)),
      citizen,
      name,
      behavior: col(x => /^(คะแนนพฤติกรรม|คะแนนความประพฤติ|คะแนนพฤติกรรมคงเหลือ|คะแนนคงเหลือ)$/.test(x))
    };
  }
  return null;
}

/** อ่านรายชื่อนักเรียนเฉพาะ ม.1/1–ม.6/5 จากแท็บห้องเรียนจริง */
function readTargetRoster_(ss, batchRows) {
  const students = [];
  let studentsWithoutCitizenId = 0;

  getTargetClassNames_().forEach(className => {
    const cachedRows = batchRows && batchRows.classRows &&
      Object.prototype.hasOwnProperty.call(batchRows.classRows, className)
      ? batchRows.classRows[className] : null;
    const sh = ss ? ss.getSheetByName(className) : null;
    if (!cachedRows && !sh) throw new Error('ไม่พบแท็บห้องเรียน ' + className);
    const rows = cachedRows !== null
      ? cachedRows
      : sh.getRange(1, 1, Math.max(3, Math.min(sh.getLastRow(), APP.MAX_ROWS)), 6).getDisplayValues();
    const schema = findClassHeader_(rows.slice(0, 12));
    if (!schema || schema.name < 0) throw new Error('อ่านหัวตารางรายชื่อนักเรียนไม่ได้: ' + className);
    const cls = parseTargetClass_(className, '');
    if (!cls) throw new Error('ชื่อแท็บห้องเรียนไม่ถูกต้อง: ' + className);
    for (let rowIndex = schema.headerRow + 1; rowIndex < rows.length; rowIndex++) {
      const row = rows[rowIndex];
      const name = String(row[schema.name] || '').trim();
      if (!name) continue;
      const citizen = schema.citizen >= 0 ? normalizeCitizen_(row[schema.citizen]) : '';
      const studentNo = schema.no >= 0 ? normalizeStudentNo_(row[schema.no]) : '';
      const behaviorScore = schema.behavior >= 0 ? String(row[schema.behavior] || '').trim() : '';
      const student = {
        citizen,
        name,
        normName: normalizeStudentName_(name),
        gender: classifyStudentGender_(name),
        studentNo,
        classLabel: className,
        grade: cls.grade,
        room: String(cls.room),
        behaviorScore,
        hasBehaviorScore: behaviorScore !== '' && behaviorScore !== '-'
      };
      students.push(student);
      if (!citizen) studentsWithoutCitizenId++;
    }
  });
  return { students, studentsWithoutCitizenId };
}

/** จำแนกเพศจากคำนำหน้าชื่อจริงในทะเบียน (ไม่เดาจากชื่อจริง/นามสกุล) */
function classifyStudentGender_(value) {
  const name = String(value || '').normalize('NFKC').trim().replace(/\s+/g, ' ');
  if (/^(?:เด็กชาย|ด\.?\s*ช\.?|นาย)/.test(name)) return 'male';
  if (/^(?:เด็กหญิง|ด\.?\s*ญ\.?|นางสาว|นาง)/.test(name)) return 'female';
  return 'unknown';
}

function normalizeCitizen_(value) {
  const s = normalizeThaiDigits_(value).replace(/\D/g, '');
  return /^\d{13}$/.test(s) ? s : '';
}

function normalizeStudentNo_(value) {
  return normalizeThaiDigits_(value).trim().replace(/\.0$/, '').replace(/\s+/g, '');
}

/** ค้นหาหัวตารางของชีทแสดงผล 0 ร มส */
function findStatusHeader_(rows) {
  for (let r = 0; r < Math.min(rows.length, 15); r++) {
    const h = (rows[r] || []).map(normalizeHeaderKey_);
    const col = fn => h.findIndex(fn);
    const name = col(x => /^(ชื่อนักเรียน|ชื่อสกุล|ชื่อนามสกุล)$/.test(x));
    const classCol = col(x => /^(ชั้น|ระดับชั้น|ชั้นเรียน)$/.test(x));
    const studentNo = col(x => /^(เลขที่|ที่)$/.test(x));
    const status = col(x => /^(สถานะ|ผลการเรียน)$/.test(x));
    if (name < 0 || classCol < 0 || status < 0) continue;
    return {
      headerRow: r, name, classCol, studentNo, status,
      subjectCode: col(x => /^(รหัสวิชา|รหัสรายวิชา)$/.test(x)),
      subject: col(x => /^(รายวิชา|ชื่อวิชา)$/.test(x)),
      teacher: col(x => /^(ครูประจำวิชา|ครูผู้สอน)$/.test(x)),
      year: col(x => /^ปีการศึกษา$/.test(x)),
      term: col(x => /^(ภาคเรียนที่|เทอม)$/.test(x)),
      editStatus: col(x => /^สถานะการแก้ไข$/.test(x))
    };
  }
  return null;
}

/** อ่านเฉพาะรายการ 0 / ร / มส ในห้อง ม.1/1–ม.6/5 จากชีทแสดงผล 0 ร มส */
function readTargetStatusRecords_(ss, batchRows) {
  const hasCachedRows = batchRows && Array.isArray(batchRows.statusRows);
  const sh = ss ? ss.getSheetByName(APP.STATUS_SHEET) : null;
  if (!hasCachedRows && !sh) throw new Error('ไม่พบชีท ' + APP.STATUS_SHEET);
  const rows = hasCachedRows
    ? batchRows.statusRows
    : sh.getRange(1, 1, Math.max(5, Math.min(sh.getLastRow(), APP.MAX_STATUS_ROWS)), 12).getDisplayValues();
  const schema = findStatusHeader_(rows.slice(0, 15));
  if (!schema) throw new Error('ไม่พบหัวตารางข้อมูล 0 / ร / มส ในชีท ' + APP.STATUS_SHEET);
  const output = [];
  const completedStatusKey = normalizeHeaderKey_('แก้ไขเสร็จแล้ว');
  for (let rowIndex = schema.headerRow + 1; rowIndex < rows.length; rowIndex++) {
    const row = rows[rowIndex];
    const name = String(row[schema.name] || '').trim();
    const type = normalizeStatusType_(row[schema.status]);
    if (!name || !type) continue;
    const cls = parseTargetClass_(row[schema.classCol], '');
    if (!cls) continue;
    const editStatus = schema.editStatus >= 0 ? String(row[schema.editStatus] || '').trim() : '';
    if (normalizeHeaderKey_(editStatus) === completedStatusKey) continue;
    output.push({
      classLabel: cls.label,
      studentNo: schema.studentNo >= 0 ? normalizeStudentNo_(row[schema.studentNo]) : '',
      name,
      normName: normalizeStudentName_(name),
      type,
      subjectCode: schema.subjectCode >= 0 ? String(row[schema.subjectCode] || '').trim() : '',
      subject: schema.subject >= 0 ? String(row[schema.subject] || '').trim() : '',
      teacher: schema.teacher >= 0 ? String(row[schema.teacher] || '').trim() : '',
      year: schema.year >= 0 ? String(row[schema.year] || '').trim() : '',
      term: schema.term >= 0 ? String(row[schema.term] || '').trim() : '',
      editStatus
    });
  }
  return output;
}

function normalizeStatusType_(value) {
  const s = normalizeThaiDigits_(value).trim().replace(/\s+/g, '');
  if (s === '0' || s === 'ติด0') return 'zero';
  if (s === 'ร' || s === 'ติดร') return 'r';
  if (s === 'มส' || s === 'ติดมส') return 'ms';
  return '';
}

function aggregateStudentData_(raw) {
  const output = {
    studentData: [],
    summary: { totalLevels: 0, totalRooms: 0, totalMale: 0,
      totalFemale: 0, totalUnknown: 0, grandTotal: 0 }
  };
  if (!raw || raw.length === 0 || raw.every(r => r.every(v => v === ''))) return output;
  const h = findHeader_(raw);
  if (!h) throw new Error('ไม่พบคอลัมน์ระดับชั้น/ชั้น กรุณาตรวจสอบหัวตาราง');
  const agg = {};
  for (let r = h.row + 1; r < raw.length; r++) {
    const row = raw[r];
    // Dashboard แสดงเฉพาะห้อง ม.1/1 ถึง ม.6/5 เท่านั้น
    const cls = parseTargetClass_(row[h.grade], h.room >= 0 ? row[h.room] : '');
    if (!cls) continue;
    const key = cls.label;
    if (!agg[key]) agg[key] = { grade: cls.grade, room: cls.room, male: 0, female: 0, unknown: 0 };
    const item = agg[key];

    const male = h.male >= 0 ? number_(row[h.male]) : 0;
    const female = h.female >= 0 ? number_(row[h.female]) : 0;
    const total = h.total >= 0 ? number_(row[h.total]) : 0;
    const hasAggregate = (h.male >= 0 && row[h.male] !== '' && row[h.male] != null) ||
      (h.female >= 0 && row[h.female] !== '' && row[h.female] != null) ||
      (h.total >= 0 && row[h.total] !== '' && row[h.total] != null);
    if (hasAggregate) {
      item.male += male;
      item.female += female;
      item.unknown += Math.max(0, total - male - female);
    } else {
      const gender = h.gender >= 0 ? String(row[h.gender] ?? '') : '';
      const g = gender.replace(/\s/g, '').toLowerCase();
      if (/^(ช|ชาย|ด\.ช\.|เด็กชาย|นาย|m|male)$/.test(g)) item.male++;
      else if (/^(ญ|หญิง|ด\.ญ\.|เด็กหญิง|นางสาว|นาง|f|female)$/.test(g)) item.female++;
      else item.unknown++;
    }
  }
  Object.keys(agg).forEach(key => {
    const v = agg[key];
    const total = v.male + v.female + v.unknown;
    output.studentData.push({grade:v.grade, room:v.room, label:key, rooms:1, male:v.male, female:v.female,
      unknown:v.unknown, total, type:'sec'});
    output.summary.totalRooms++;
    output.summary.totalMale += v.male;
    output.summary.totalFemale += v.female;
    output.summary.totalUnknown += v.unknown;
  });
  output.studentData.sort((a,b) => {
    const ca = parseTargetClass_(a.label, '');
    const cb = parseTargetClass_(b.label, '');
    return (Number(ca && ca.grade.match(/\d+/)[0]) - Number(cb && cb.grade.match(/\d+/)[0])) ||
      (Number(ca && ca.room) - Number(cb && cb.room));
  });
  output.summary.totalLevels = new Set(output.studentData.map(x => x.grade)).size;
  output.summary.grandTotal = output.summary.totalMale + output.summary.totalFemale + output.summary.totalUnknown;
  return output;
}

/** แปลงรูปแบบชั้น/ห้องให้เป็น ม.n/r และยอมรับเฉพาะ ม.1/1–ม.6/5 */
function normalizeThaiDigits_(value) {
  const thai = '๐๑๒๓๔๕๖๗๘๙';
  return String(value ?? '').replace(/[๐-๙]/g, d => String(thai.indexOf(d))).trim().replace(/\s+/g, ' ');
}

function parseTargetClass_(gradeValue, roomValue) {
  const gradeText = normalizeThaiDigits_(gradeValue);
  const roomText = normalizeThaiDigits_(roomValue);
  const candidates = [gradeText, roomText].filter(Boolean);
  let gradeNo = 0, roomNo = 0;
  const prefix = '(?:มัธยมศึกษา(?:ปีที่)?|มัธยมปีที่|ม\\.?)';
  const combined = new RegExp('^' + prefix + '\\s*(\\d+)\\s*(?:/|ห้อง)\\s*(\\d+)$', 'i');
  for (const candidate of candidates) {
    const m = candidate.match(combined);
    if (m) { gradeNo = Number(m[1]); roomNo = Number(m[2]); break; }
  }
  if (!gradeNo) {
    const gradeOnly = gradeText.match(new RegExp('^' + prefix + '\\s*(\\d+)$', 'i'));
    if (gradeOnly) gradeNo = Number(gradeOnly[1]);
  }
  if (!roomNo && roomText) {
    const roomMatch = roomText.match(/(?:ห้อง\s*)?(\d+)\s*$/);
    if (roomMatch) roomNo = Number(roomMatch[1]);
  }
  if (gradeNo < 1 || gradeNo > 6 || roomNo < 1 || roomNo > 5) return null;
  return {grade:'ม.' + gradeNo, room:roomNo, label:'ม.' + gradeNo + '/' + roomNo};
}

function number_(v) {
  const s = String(v ?? '').replace(/,/g, '').trim();
  const n = Number(s);
  return Number.isFinite(n) && n > 0 ? Math.floor(n) : 0;
}

function normalizeGrade_(value) {
  const s = String(value ?? '').trim().replace(/\s+/g, ' ');
  if (!s) return '';
  const schemes = [
    [/^(?:มัธยมศึกษา(?:ปีที่)?|มัธยมปีที่|ม\.?)[\s.]*(\d+)$/i, 'ม.'],
    [/^(?:ประถมศึกษา(?:ปีที่)?|ประถมปีที่|ป\.?)[\s.]*(\d+)$/i, 'ป.'],
    [/^(?:อนุบาล(?:ปีที่)?|อ\.?)[\s.]*(\d+)$/i, 'อ.']
  ];
  for (const [re, prefix] of schemes) {
    const m = s.match(re);
    if (m) return prefix + Number(m[1]);
  }
  return s;
}

function gradeSort_(g) {
  const m = g.match(/^(อ\.|ป\.|ม\.)(\d+)/);
  if (!m) return 9000;
  return ({'อ.': 0, 'ป.': 100, 'ม.': 200}[m[1]] || 9000) + Number(m[2]);
}

function findHeader_(rows) {
  for (let i=0; i<Math.min(rows.length,20); i++) {
    const headings = rows[i].map(x => String(x ?? '').trim().replace(/\s+/g,'').toLowerCase());
    const find = fn => headings.findIndex(fn);
    const grade = find(s => s === 'ชั้น' || s === 'ระดับชั้น' || s === 'ชั้นเรียน' || s === 'ชั้น/ห้อง' || s === 'ชั้นห้อง' || s === 'ห้องเรียน' || s === 'class' || s === 'classroom' || s.includes('ระดับชั้น'));
    if (grade === -1) continue;
    return {
      row:i, grade,
      room:find(s => s === 'ห้อง' || s === 'ห้องเรียน' || s === 'ห้องที่' || s === 'room'),
      roomCount:find(s => s === 'จำนวนห้อง' || s === 'ห้อง(ห้อง)'),
      gender:find(s => s === 'เพศ' || s === 'เพศนักเรียน'),
      male:find(s => s === 'ชาย' || s === 'นักเรียนชาย' || s === 'จำนวนชาย' || s === 'ชาย(คน)'),
      female:find(s => s === 'หญิง' || s === 'นักเรียนหญิง' || s === 'จำนวนหญิง' || s === 'หญิง(คน)'),
      total:find(s => s === 'รวม' || s === 'รวม(คน)' || s === 'จำนวนนักเรียน' || s === 'รวมทั้งสิ้น')
    };
  }
  return null;
}

function studentSheetPriority_(name) {
  if (/รายชื่อ|ทะเบียน|นักเรียน|student/i.test(name)) return 10;
  if (/0\s*ร\s*มส|แก้\s*0|พฤติกรรม|คะแนน|สรุป/i.test(name)) return -10;
  return 0;
}

function sheetHasTargetClasses_(sheet) {
  if (!sheet || sheet.getLastRow() < 2 || sheet.getLastColumn() < 1) return false;
  const rows = sheet.getRange(1, 1, Math.min(sheet.getLastRow(), APP.MAX_ROWS),
    Math.min(sheet.getLastColumn(), APP.MAX_COLUMNS)).getDisplayValues();
  const h = findHeader_(rows);
  if (!h) return false;
  for (let i = h.row + 1; i < rows.length; i++) {
    if (parseTargetClass_(rows[i][h.grade], h.room >= 0 ? rows[i][h.room] : '')) return true;
  }
  return false;
}

function ensureStudentSheet_(ss) {
  const p = PropertiesService.getScriptProperties();
  const chosenName = p.getProperty('DATA_SHEET_NAME');
  const chosen = chosenName ? ss.getSheetByName(chosenName) : null;
  if (chosen && sheetHasTargetClasses_(chosen)) {
    p.setProperty('DATA_SOURCE_VALIDATED_V4', '1');
    return chosen;
  }

  const candidates = ss.getSheets().slice().filter(sh =>
    sh.getName() !== APP.QUEUE_SHEET && !/^__BKP_/i.test(sh.getName()) && !/^__/.test(sh.getName()))
    .sort((a,b) => studentSheetPriority_(b.getName()) - studentSheetPriority_(a.getName()));
  const targetSheet = candidates.find(sheetHasTargetClasses_);
  if (targetSheet) {
    p.setProperty('DATA_SHEET_NAME', targetSheet.getName());
    p.setProperty('DATA_SOURCE_VALIDATED_V4', '1');
    return targetSheet;
  }

  // ถ้ายังไม่มีแถวห้องเรียนเป้าหมาย ให้คงชีทเดิมที่มีหัวตารางไว้ ไม่ลบข้อมูลใด ๆ
  const own = ss.getSheetByName(APP.DATA_SHEET);
  if (own && findHeader_(own.getRange(1,1,Math.min(own.getLastRow() || 1,20),
      Math.min(own.getLastColumn() || 1,APP.MAX_COLUMNS)).getDisplayValues())) {
    p.setProperty('DATA_SHEET_NAME', own.getName());
    p.setProperty('DATA_SOURCE_VALIDATED_V4', '1');
    return own;
  }
  if (chosen && findHeader_(chosen.getRange(1,1,Math.min(chosen.getLastRow() || 1,20),
      Math.min(chosen.getLastColumn() || 1,APP.MAX_COLUMNS)).getDisplayValues())) {
    p.setProperty('DATA_SOURCE_VALIDATED_V4', '1');
    return chosen;
  }

  const sheet = own || ss.insertSheet(APP.DATA_SHEET);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1,1,1,5).setValues([['ลำดับ','รหัสนักเรียน','ระดับชั้น','ห้อง','เพศ']]);
    sheet.getRange(1,1,1,5).setBackground('#2563eb').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  p.setProperty('DATA_SHEET_NAME', sheet.getName());
  p.setProperty('DATA_SOURCE_VALIDATED_V4', '1');
  return sheet;
}

function getStudentSheet_(ss) {
  const p = PropertiesService.getScriptProperties();
  const name = p.getProperty('DATA_SHEET_NAME');
  const sheet = name ? ss.getSheetByName(name) : null;
  if (!sheet || p.getProperty('DATA_SOURCE_VALIDATED_V4') !== '1') return ensureStudentSheet_(ss);
  return sheet;
}

function ensureQueueSheet_(ss) {
  let sheet = ss.getSheetByName(APP.QUEUE_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(APP.QUEUE_SHEET);
    sheet.appendRow(['สร้างเมื่อ','ประเภท','ข้อความ','Telegram','LINE','ความพยายาม','สถานะ','รายละเอียด']);
    sheet.getRange(1,1,1,8).setBackground('#7c3aed').setFontColor('#fff').setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.hideSheet();
  }
  return sheet;
}

/** ปิดการนำเข้าไฟล์นักเรียน: ระบบอ่านข้อมูลจากแท็บห้องเรียนจริงเท่านั้น */
function uploadStudentData(token) {
  requireAdmin_(token);
  throw new Error('ปิดการนำเข้าไฟล์นักเรียนแล้ว ระบบใช้ข้อมูลจากแท็บ ม.1/1–ม.6/5 และชีทแสดงผล 0 ร มส โดยตรง');
}

function uploadBranding(token, kind, base64Data, mimeType) {
  try {
    requireAdmin_(token);
    if (kind !== 'logo' && kind !== 'background') throw new Error('ชนิดรูปภาพไม่ถูกต้อง');
    if (!/^image\/(png|jpeg|webp)$/.test(String(mimeType))) throw new Error('รองรับ PNG / JPEG / WebP เท่านั้น');
    const bytes = Utilities.base64Decode(String(base64Data || ''));
    if (!bytes.length || bytes.length > APP.MAX_IMAGE_BYTES) throw new Error('รูปภาพต้องไม่เกิน 2 MB');
    const ext = mimeType === 'image/png' ? 'png' : mimeType === 'image/webp' ? 'webp' : 'jpg';
    const filename = APP.SCHOOL + '_' + kind + '_' + Date.now() + '.' + ext;
    const file = DriveApp.getFolderById(APP.FOLDER_ID)
      .createFile(Utilities.newBlob(bytes,mimeType,filename));
    PropertiesService.getScriptProperties().setProperty(kind === 'logo' ? 'BRAND_LOGO_ID' : 'BRAND_BACKGROUND_ID',file.getId());
    CacheService.getScriptCache().remove('BRANDING_CACHE_V2');
    return {success:true,message:kind === 'logo' ? 'บันทึกโลโก้แล้ว' : 'บันทึกภาพพื้นหลังแล้ว'};
  } catch(e) { return {success:false,error:e.message || String(e)}; }
}

function getBranding() {
  let cache = null;
  try { cache = CacheService.getScriptCache(); } catch (e) {}
  if (cache) {
    try {
      const cached = cache.get('BRANDING_CACHE_V2');
      if (cached) return JSON.parse(cached);
    } catch (e) {
      try { cache.remove('BRANDING_CACHE_V2'); } catch (removeError) {}
    }
  }

  const p = PropertiesService.getScriptProperties();
  const result = {
    success: true,
    logo: readBrand_(p.getProperty('BRAND_LOGO_ID')),
    background: readBrand_(p.getProperty('BRAND_BACKGROUND_ID'))
  };
  if (cache) {
    try { cache.put('BRANDING_CACHE_V2', JSON.stringify(result), 3600); }
    catch (cacheError) {
      console.warn('Branding cache skipped: ' + (cacheError.message || cacheError));
    }
  }
  return result;
}
function readBrand_(id) {
  if (!id) return '';
  try {
    const blob = DriveApp.getFileById(id).getBlob();
    const bytes = blob.getBytes();
    if (bytes.length > APP.MAX_IMAGE_BYTES) return '';
    return 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(bytes);
  } catch (e) { return ''; }
}

function enqueueNotification_(ss,type,message) {
  const p = PropertiesService.getScriptProperties();
  const tg = Boolean(p.getProperty('TELEGRAM_BOT_TOKEN') && p.getProperty('TELEGRAM_CHAT_ID'));
  const line = Boolean(p.getProperty('LINE_CHANNEL_ACCESS_TOKEN') && p.getProperty('LINE_TO'));
  ensureQueueSheet_(ss).appendRow([new Date(),type,message,
    tg ? 'PENDING':'SKIP',line ? 'PENDING':'SKIP',0,(tg||line)?'PENDING':'SKIPPED','']);
}
function ensureQueueTrigger_() {
  if (!ScriptApp.getProjectTriggers().some(t =>
    t.getHandlerFunction() === 'processNotificationQueue_')) {
    ScriptApp.newTrigger('processNotificationQueue_').timeBased().everyMinutes(1).create();
  }
}

/** Trigger ราย 1 นาที: ทำให้หน้าเว็บไม่ต้องรอ API แจ้งเตือน */
function processNotificationQueue_() {
  const lock=LockService.getScriptLock();
  if (!lock.tryLock(2000)) return;
  try {
    const ss=SpreadsheetApp.openById(APP.SHEET_ID);
    const sh=ensureQueueSheet_(ss);
    const last=sh.getLastRow();
    if(last<2)return;
    const rows=sh.getRange(2,1,last-1,8).getValues();
    let processed=0;
    for(let i=0;i<rows.length && processed<8;i++){
      const r=rows[i], rowNum=i+2;
      if(r[6]!=='PENDING')continue;
      processed++;
      let tg=r[3],line=r[4],attempts=Number(r[5])||0;
      const report=[];
      if(tg==='PENDING'){
        const result=sendTelegram_(String(r[2]));
        tg=result.ok?'SENT':(attempts>=2?'FAILED':'PENDING');
        report.push('Telegram: '+result.message);
      }
      if(line==='PENDING'){
        const result=sendLine_(String(r[2]));
        line=result.ok?'SENT':(attempts>=2?'FAILED':'PENDING');
        report.push('LINE: '+result.message);
      }
      const nextStatus=(tg==='PENDING'||line==='PENDING')?'PENDING':
        ((tg==='FAILED'||line==='FAILED')?'PARTIAL/FAILED':'DONE');
      sh.getRange(rowNum,4,1,5).setValues([[tg,line,attempts+1,nextStatus,report.join(' | ').slice(0,500)]]);
    }
  }finally{lock.releaseLock();}
}

function testNotifications(token) {
  try {
    requireAdmin_(token);
    const msg='ทดสอบระบบแจ้งเตือน\n' + APP.DEPARTMENT + '\n' + APP.SCHOOL +
      '\nเวลาทดสอบ: ' + Utilities.formatDate(new Date(),'Asia/Bangkok','dd/MM/yyyy HH:mm:ss');
    const tg=sendTelegram_(msg),line=sendLine_(msg);
    return {success:tg.ok||line.ok,telegram:tg.message,line:line.message};
  }catch(e){return{success:false,error:e.message||String(e)};}
}
function sendTelegram_(message) {
  const p=PropertiesService.getScriptProperties();
  const bot=p.getProperty('TELEGRAM_BOT_TOKEN'), chat=p.getProperty('TELEGRAM_CHAT_ID');
  if(!bot||!chat)return{ok:false,message:'ยังไม่ตั้งค่า'};
  try{
    const r=UrlFetchApp.fetch('https://api.telegram.org/bot'+bot+'/sendMessage',{
      method:'post',contentType:'application/json',
      payload:JSON.stringify({chat_id:chat,text:message}),muteHttpExceptions:true
    });
    const obj=JSON.parse(r.getContentText()||'{}');
    return{ok:r.getResponseCode()===200&&obj.ok===true,
      message:r.getResponseCode()===200&&obj.ok?'ส่งสำเร็จ':'HTTP '+r.getResponseCode()+' '+String(obj.description||'').slice(0,100)};
  }catch(e){return{ok:false,message:String(e.message||e).slice(0,120)};}
}
function sendLine_(message) {
  const p=PropertiesService.getScriptProperties();
  const token=p.getProperty('LINE_CHANNEL_ACCESS_TOKEN'), to=p.getProperty('LINE_TO');
  if(!token||!to)return{ok:false,message:'ยังไม่ตั้งค่า'};
  try{
    const r=UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push',{
      method:'post',contentType:'application/json',
      headers:{Authorization:'Bearer '+token},
      payload:JSON.stringify({to,messages:[{type:'text',text:message.slice(0,4900)}]}),
      muteHttpExceptions:true
    });
    return{ok:r.getResponseCode()===200,
      message:r.getResponseCode()===200?'ส่งคำขอสำเร็จ':'HTTP '+r.getResponseCode()+' '+r.getContentText().slice(0,100)};
  }catch(e){return{ok:false,message:String(e.message||e).slice(0,120)};}
}


/**
 * STATUS LOOKUP (V6)
 * ใช้ข้อมูลจากแท็บ ม.1/1–ม.6/5 และชีทแสดงผล 0 ร มส เท่านั้น
 * นักเรียนใช้เลขบัตรประชาชน 13 หลักเป็นทั้ง Username และ Password และเห็นข้อมูลตนเองเท่านั้น
 * ผู้ดูแลระบบ admin ค้นหานักเรียนได้ด้วยเลขบัตรประชาชนหรือชื่อ–นามสกุล
 */
function studentStatusLoginFailure_(username) {
  const cache = CacheService.getScriptCache();
  const key = digest_(normalizeThaiDigits_(String(username || 'unknown'))).slice(0, 28);
  const lockKey = 'STU_STATUS_LOCK_' + key;
  if (cache.get(lockKey) === '1') throw new Error('เข้าสู่ระบบผิดหลายครั้ง กรุณาลองใหม่หลังจาก 5 นาที');
  const failKey = 'STU_STATUS_FAIL_' + key;
  const attempts = Number(cache.get(failKey) || 0) + 1;
  if (attempts >= 5) {
    cache.put(lockKey, '1', 300);
    cache.remove(failKey);
    throw new Error('เข้าสู่ระบบผิดครบ 5 ครั้ง กรุณาลองใหม่หลังจาก 5 นาที');
  }
  cache.put(failKey, String(attempts), 300);
  throw new Error('เลขบัตรประชาชนหรือรหัสผ่านไม่ถูกต้อง หรือไม่พบเลขบัตรในรายชื่อ ม.1/1–ม.6/5');
}

/** นักเรียนใช้เลขบัตรประชาชน 13 หลักเดียวกันทั้ง Username และ Password */
function loginStudentStatus(username, password) {
  const user = normalizeCitizen_(username);
  const pass = normalizeCitizen_(password);
  const cache = CacheService.getScriptCache();
  const rateKey = digest_(normalizeThaiDigits_(String(username || 'unknown'))).slice(0, 28);
  if (cache.get('STU_STATUS_LOCK_' + rateKey) === '1') {
    throw new Error('เข้าสู่ระบบผิดหลายครั้ง กรุณาลองใหม่หลังจาก 5 นาที');
  }
  if (!/^\d{13}$/.test(user) || user !== pass) return studentStatusLoginFailure_(username);
  const ss = ensureAcademicSetup_();
  const source = getStatusLookupData_(ss);
  const citizenHash = digest_('CITIZEN:' + user);
  const index = source.studentIndexByCitizenHash && source.studentIndexByCitizenHash[citizenHash];
  if (index === undefined || index === -1 || !source.students[index]) return studentStatusLoginFailure_(username);
  cache.remove('STU_STATUS_FAIL_' + rateKey);
  const token = Utilities.getUuid() + Utilities.getUuid();
  // เก็บค่า hash ของเลขบัตรใน session cache แทนเลขบัตรจริง
  cache.put('STUDENT_STATUS_AUTH_' + digest_(token), digest_('CITIZEN:' + user), APP.SESSION_SECONDS);
  return { success: true, token: token, expiresIn: APP.SESSION_SECONDS };
}

function requireStatusAccess_(token) {
  try {
    requireAdmin_(token);
    return { role: 'admin' };
  } catch (adminError) {
    const citizenHash = typeof token === 'string' && token.length === 72
      ? CacheService.getScriptCache().get('STUDENT_STATUS_AUTH_' + digest_(token)) : null;
    if (citizenHash && /^[a-f0-9]{64}$/.test(citizenHash)) return { role: 'student', citizenHash };
    throw new Error('เซสชันหมดอายุ กรุณาเข้าสู่ระบบด้วยเลขบัตรประชาชน');
  }
}

function logoutStudentStatus(token) {
  if (typeof token === 'string') {
    CacheService.getScriptCache().remove('STUDENT_STATUS_AUTH_' + digest_(token));
  }
  return { success: true };
}

function matchStatusRecordsForStudent_(student, records) {
  return (records || []).filter(rec => {
    if (rec.classLabel !== student.classLabel) return false;
    const sameNo = student.studentNo && rec.studentNo && student.studentNo === rec.studentNo;
    if (sameNo) return true;
    const sameName = student.normName && rec.normName && student.normName === rec.normName;
    return Boolean(sameName && (!student.studentNo || !rec.studentNo));
  });
}

function getStudentStatus(sessionToken, searchTerm) {
  try {
    const access = requireStatusAccess_(sessionToken);
    const ss = ensureAcademicSetup_();
    const source = getStatusLookupData_(ss);
    const students = source.students || [];
    let student = null;

    if (access.role === 'student') {
      const index = source.studentIndexByCitizenHash && source.studentIndexByCitizenHash[access.citizenHash];
      if (index === undefined || index === -1 || !students[index]) {
        return { success: true, found: false, message: 'ไม่พบข้อมูลบัญชีของนักเรียน กรุณาติดต่อกลุ่มบริหารงานวิชาการ' };
      }
      student = students[index];
    } else {
      const search = normalizeThaiDigits_(String(searchTerm || '')).trim();
      if (/^\d{13}$/.test(search)) {
        const searchHash = digest_('CITIZEN:' + search);
        const index = source.studentIndexByCitizenHash && source.studentIndexByCitizenHash[searchHash];
        if (index === -1) {
          return { success: true, found: false, ambiguous: true, message: 'พบเลขบัตรประชาชนซ้ำในข้อมูล กรุณาตรวจสอบทะเบียนนักเรียน' };
        }
        student = index === undefined ? null : (students[index] || null);
      } else {
        const nameQuery = normalizeStudentName_(search);
        if (nameQuery.length < 3 || !/[^0-9]/.test(nameQuery) || search.length > 120) {
          throw new Error('กรุณากรอกเลขบัตรประชาชน 13 หลัก หรือชื่อ–นามสกุลนักเรียน');
        }
        const exactIndex = source.studentIndexByName && source.studentIndexByName[nameQuery];
        if (exactIndex === -1) {
          return { success: true, found: false, ambiguous: true, message: 'พบชื่อที่ตรงกันมากกว่าหนึ่งราย กรุณาค้นหาด้วยเลขบัตรประชาชน 13 หลัก' };
        }
        if (exactIndex !== undefined) {
          student = students[exactIndex] || null;
        } else {
          const candidates = students.filter(s => s.normName && s.normName.includes(nameQuery));
          if (candidates.length > 1) {
            return { success: true, found: false, ambiguous: true, message: 'พบชื่อที่ตรงกันมากกว่าหนึ่งราย กรุณาค้นหาด้วยเลขบัตรประชาชน 13 หลัก' };
          }
          student = candidates[0] || null;
        }
      }
      if (!student) return { success: true, found: false, message: 'ไม่พบข้อมูลในรายชื่อนักเรียน ม.1/1–ม.6/5' };
    }

    const studentKey = student.classLabel + '|' + (student.studentNo || student.normName);
    const matches = source.statusByStudent && Object.prototype.hasOwnProperty.call(source.statusByStudent, studentKey)
      ? source.statusByStudent[studentKey]
      : [];
    const counts = { zero: 0, r: 0, ms: 0 };
    const details = matches.map(rec => {
      counts[rec.type]++;
      return {
        type: rec.type,
        subject: rec.subject || 'ไม่ระบุรายวิชา',
        subjectCode: rec.subjectCode,
        teacher: rec.teacher,
        year: rec.year,
        term: rec.term,
        editStatus: rec.editStatus || 'ยังไม่ระบุสถานะ'
      };
    });
    return {
      success: true,
      found: true,
      student: {
        name: student.name,
        grade: student.grade,
        room: student.room,
        studentNo: student.studentNo || '',
        citizenId: student.citizenId || '',
        classLabel: student.classLabel
      },
      academic: {
        available: true,
        zero: counts.zero,
        r: counts.r,
        ms: counts.ms,
        total: counts.zero + counts.r + counts.ms,
        details: details
      },
      behavior: {
        available: student.hasBehaviorScore,
        score: student.behaviorScore,
        source: student.classLabel
      },
      note: access.role === 'student'
        ? 'แสดงข้อมูลเฉพาะเจ้าของเลขบัตรที่เข้าสู่ระบบ'
        : 'ผู้ดูแลระบบตรวจสอบได้เฉพาะข้อมูลจากห้อง ม.1/1–ม.6/5'
    };
  } catch (e) {
    return { success: false, error: e.message || String(e) };
  }
}

function normalizeStudentName_(value) {
  return String(value || '').normalize('NFKC')
    .replace(/^(?:เด็กชาย|เด็กหญิง|ด\.?ช\.?|ด\.?ญ\.?|นาย|นางสาว|นาง)\s*/, '')
    .replace(/[.\s\-–_]+/g, '').toLowerCase();
}


/**
 * สร้าง PDF รายงานจาก Google Docs ต้นฉบับ
 * - อ่านข้อมูลซ้ำจาก session ฝั่งเซิร์ฟเวอร์ ไม่เชื่อข้อมูลรายงานที่ส่งมาจาก browser
 * - ทำสำเนาเอกสารต้นแบบ เติมข้อมูลลงในตารางเดิม แล้วส่ง PDF กลับให้หน้าเว็บ
 * - ลบสำเนาชั่วคราวหลังส่งออก เพื่อไม่เปิดเผยข้อมูลนักเรียนผ่านไฟล์สาธารณะ
 */
function generateStudentStatusReportPdf(sessionToken, searchTerm) {
  const started = Date.now();
  const result = getStudentStatus(sessionToken, searchTerm || '');

  if (!result || result.success !== true) {
    return {
      success: false,
      error: result && result.error ? result.error : 'ตรวจสอบข้อมูลนักเรียนไม่สำเร็จ',
      elapsedMs: Date.now() - started
    };
  }

  // ส่งผลการค้นหากลับไปด้วย เพื่อให้หน้าเว็บไม่ต้องเรียก getStudentStatus ซ้ำอีกครั้ง
  if (result.found !== true || !result.student) {
    return {
      success: true,
      found: false,
      studentResult: result,
      elapsedMs: Date.now() - started
    };
  }

  const student = result.student;
  const safeName = String(student.name || 'นักเรียน')
    .replace(/[\\\/:*?"<>|]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 80) || 'นักเรียน';
  const fileName = 'ใบแสดงผลสถานะผลการเรียน_' + safeName + '.pdf';
  const pageCount = Math.max(1, Math.ceil(
    Math.max(1, (result.academic && result.academic.details || []).length) /
    APP.REPORT_ROWS_PER_PAGE
  ));

  // Cache ผูกกับ session + นักเรียน + DATA_REVISION จึงไม่ส่ง PDF ข้ามบัญชี
  // และ cache เก่าจะไม่ถูกใช้เมื่อมีการแก้ข้อมูลต้นทาง
  const revision = PropertiesService.getScriptProperties().getProperty('DATA_REVISION') || '0';
  const studentKey = [
    student.classLabel || '',
    student.studentNo || '',
    student.citizenId || student.name || ''
  ].join('|');
  const cacheKey = 'STU_REPORT_PDF_V2_' +
    digest_(String(sessionToken || '') + '|' + revision + '|' + studentKey).slice(0, 40);

  let base64 = '';
  let cached = false;
  let cache = null;
  try {
    cache = CacheService.getScriptCache();
    base64 = readStudentReportPdfCache_(cache, cacheKey);
    cached = Boolean(base64);
  } catch (cacheReadError) {
    console.warn('Report PDF cache read skipped: ' +
      (cacheReadError.message || cacheReadError));
  }

  if (!base64) {
    try {
      const pdfBytes = createStudentStatusReportPdfBytes_(result, fileName);
      base64 = Utilities.base64Encode(pdfBytes);

      // เก็บเฉพาะไฟล์ขนาดเล็กถึงปานกลาง 3 นาที ลดเวลาสร้างซ้ำโดยไม่ค้างข้อมูลนาน
      if (cache && base64.length <= 180000) {
        try {
          writeStudentReportPdfCache_(cache, cacheKey, base64, 180);
        } catch (cacheWriteError) {
          console.warn('Report PDF cache write skipped: ' +
            (cacheWriteError.message || cacheWriteError));
        }
      }
    } catch (reportError) {
      return {
        success: true,
        found: true,
        studentResult: result,
        reportError: reportError && reportError.message
          ? reportError.message : String(reportError),
        elapsedMs: Date.now() - started
      };
    }
  }

  return {
    success: true,
    found: true,
    studentResult: result,
    fileName: fileName,
    mimeType: 'application/pdf',
    base64: base64,
    pageCount: pageCount,
    cached: cached,
    elapsedMs: Date.now() - started
  };
}

/** อ่าน PDF จาก CacheService แบบแบ่งชิ้นเพื่อไม่เกินขนาดสูงสุดต่อ cache item */
function readStudentReportPdfCache_(cache, baseKey) {
  const countText = cache.get(baseKey + '_N');
  const count = Number(countText);
  if (!Number.isInteger(count) || count < 1 || count > 3) return '';

  const keys = [];
  for (let i = 0; i < count; i++) keys.push(baseKey + '_' + i);
  const values = cache.getAll(keys);
  const chunks = keys.map(key => values[key] || '');
  if (chunks.some(chunk => !chunk)) {
    try { cache.removeAll([baseKey + '_N'].concat(keys)); } catch (e) {}
    return '';
  }
  return chunks.join('');
}

/** บันทึก PDF base64 เป็นชิ้นละ 60,000 ตัวอักษร อายุ cache ตามที่กำหนด */
function writeStudentReportPdfCache_(cache, baseKey, base64, seconds) {
  const chunkSize = 60000;
  const chunks = [];
  for (let i = 0; i < base64.length; i += chunkSize) {
    chunks.push(base64.slice(i, i + chunkSize));
  }
  if (!chunks.length || chunks.length > 3) return false;

  const values = {};
  values[baseKey + '_N'] = String(chunks.length);
  chunks.forEach((chunk, index) => {
    values[baseKey + '_' + index] = chunk;
  });
  cache.putAll(values, seconds || 180);
  return true;
}

/**
 * สร้างสำเนาเอกสารและเติมข้อมูลลงในตารางเดิม โดยคงรูปแบบจาก Google Docs
 */
function createStudentStatusReportPdfBytes_(result, fileName) {
  let temporaryFile = null;
  try {
    const templateFile = DriveApp.getFileById(APP.REPORT_TEMPLATE_DOC_ID);
    if (templateFile.getMimeType() !== 'application/vnd.google-apps.document') {
      throw new Error('ไฟล์ต้นแบบรายงานไม่ใช่ Google Docs');
    }

    temporaryFile = templateFile.makeCopy(
      'TEMP_REPORT_' + Utilities.getUuid().replace(/-/g, '')
    );

    const doc = DocumentApp.openById(temporaryFile.getId());
    const body = doc.getBody();
    const originalTables = body.getTables();

    if (originalTables.length !== 2 ||
        originalTables[0].getNumRows() < 3 ||
        originalTables[1].getNumRows() !== APP.REPORT_ROWS_PER_PAGE + 1 ||
        originalTables[1].getRow(0).getNumCells() !== 10) {
      throw new Error('โครงสร้าง Google Docs ต้นแบบเปลี่ยนไป กรุณาตรวจสอบตารางข้อมูลนักเรียนและตารางผลการเรียน');
    }

    const academic = result.academic || {};
    const behavior = result.behavior || {};
    const details = Array.isArray(academic.details) ? academic.details : [];
    const rowsPerPage = APP.REPORT_ROWS_PER_PAGE;
    const pageCount = Math.max(1, Math.ceil(Math.max(1, details.length) / rowsPerPage));

    // เก็บสำเนาองค์ประกอบทั้งหน้าไว้ก่อน เพื่อทำหน้าถัดไปด้วยรูปแบบเดียวกันทุกประการ
    const pageElements = [];
    for (let i = 0; i < body.getNumChildren(); i++) {
      pageElements.push(body.getChild(i).copy());
    }

    for (let page = 1; page < pageCount; page++) {
      body.appendPageBreak();
      pageElements.forEach(function(element) {
        const type = element.getType();
        if (type === DocumentApp.ElementType.PARAGRAPH) {
          body.appendParagraph(element.copy().asParagraph());
        } else if (type === DocumentApp.ElementType.TABLE) {
          body.appendTable(element.copy().asTable());
        } else {
          throw new Error('พบองค์ประกอบที่ไม่รองรับในหน้าแม่แบบ: ' + type);
        }
      });
    }

    const allTables = body.getTables();
    if (allTables.length !== pageCount * 2) {
      throw new Error('ไม่สามารถทำสำเนาหน้ารายงานจากต้นแบบ Google Docs ได้ครบ');
    }

    const student = result.student || {};
    const nameParts = splitStudentNameForReport_(student.name || '');
    const classText = student.classLabel ||
      ((student.grade || '') + (student.room ? '/' + student.room : ''));
    const behaviorText = behavior.available &&
      behavior.score !== undefined && behavior.score !== null
        ? String(behavior.score)
        : 'ไม่มีข้อมูล';

    for (let page = 0; page < pageCount; page++) {
      const studentTable = allTables[page * 2];
      const resultTable = allTables[page * 2 + 1];

      // ตารางข้อมูลนักเรียน 3 แถว 2 คอลัมน์: เติมต่อจากป้ายกำกับเดิม
      appendReportCellText_(studentTable.getCell(0, 0), nameParts.first);
      appendReportCellText_(studentTable.getCell(0, 1), nameParts.last);
      replaceReportClassCell_(studentTable.getCell(1, 0), classText);
      appendReportCellText_(studentTable.getCell(1, 1), student.studentNo || '');
      appendReportCellText_(studentTable.getCell(2, 0), student.citizenId || '');

      // ตารางผลการเรียน: แถว 0 เป็นหัวตาราง และแถว 1–12 เป็นข้อมูล
      for (let rowIndex = 1; rowIndex <= rowsPerPage; rowIndex++) {
        const absoluteIndex = page * rowsPerPage + rowIndex - 1;
        const item = details[absoluteIndex] || null;
        const row = resultTable.getRow(rowIndex);

        if (page === 0 && rowIndex === 1) {
          appendReportCellText_(row.getCell(1), behaviorText);
        }

        if (!item) continue;

        const type = String(item.type || '').toLowerCase();
        const marks = {
          zero: type === 'zero' ? '✓' : '',
          r: type === 'r' ? '✓' : '',
          ms: type === 'ms' ? '✓' : ''
        };
        const subject = (item.subject || 'ไม่ระบุรายวิชา') +
          (item.subjectCode ? ' (' + item.subjectCode + ')' : '');

        appendReportCellText_(row.getCell(0), String(absoluteIndex + 1));
        appendReportCellText_(row.getCell(2), marks.zero);
        appendReportCellText_(row.getCell(3), marks.r);
        appendReportCellText_(row.getCell(4), marks.ms);
        appendReportCellText_(row.getCell(5), item.editStatus || 'ยังไม่ระบุสถานะ');
        appendReportCellText_(row.getCell(6), subject);
        appendReportCellText_(row.getCell(7), item.teacher || '');
        appendReportCellText_(row.getCell(8), item.year || '');
        appendReportCellText_(row.getCell(9), item.term || '');
      }
    }

    doc.saveAndClose();

    const pdfBlob = DriveApp.getFileById(temporaryFile.getId())
      .getAs(MimeType.PDF)
      .setName(fileName);
    return pdfBlob.getBytes();
  } catch (error) {
    throw new Error('สร้างรายงานจาก Google Docs ไม่สำเร็จ: ' +
      (error && error.message ? error.message : String(error)));
  } finally {
    if (temporaryFile) {
      try {
        temporaryFile.setTrashed(true);
      } catch (cleanupError) {
        console.warn('ลบสำเนาเอกสารรายงานชั่วคราวไม่สำเร็จ: ' +
          (cleanupError.message || cleanupError));
      }
    }
  }
}

/** เติมข้อความในเซลล์โดยต่อท้ายข้อความเดิม เพื่อรักษารูปแบบของต้นแบบ */
function appendReportCellText_(cell, value) {
  const text = String(value === undefined || value === null ? '' : value);
  if (!text) return;
  const paragraph = cell.getChild(0).asParagraph();
  paragraph.appendText(text);
}

/** เปลี่ยนเฉพาะค่าชั้นเรียนในช่องที่มีป้ายกำกับ "ชั้น" อยู่แล้ว */
function replaceReportClassCell_(cell, classText) {
  const paragraph = cell.getChild(0).asParagraph();
  const text = paragraph.editAsText();
  const current = text.getText();
  if (!/ม\.\d+\/\d+/.test(current)) {
    throw new Error('ไม่พบตำแหน่งชั้นเรียนในช่องข้อมูลของ Google Docs ต้นแบบ');
  }
  text.replaceText('ม\\.\\d+\\/\\d+', String(classText || ''));
}

/** แยกคำนำหน้าและชื่อออกจากนามสกุลตามรูปแบบรายงานเดิม */
function splitStudentNameForReport_(fullName) {
  const value = String(fullName || '').trim().replace(/\s+/g, ' ');
  const prefixMatch = value.match(/^(เด็กชาย|เด็กหญิง|ด\.?\s*ช\.?|ด\.?\s*ญ\.?|นาย|นางสาว|นาง)\s*/);
  const prefix = prefixMatch ? prefixMatch[0].trim() : '';
  const body = prefix ? value.slice(prefixMatch[0].length).trim() : value;
  const parts = body.split(/\s+/).filter(Boolean);
  const first = parts.shift() || '';
  return {
    first: (prefix ? prefix : '') + first,
    last: parts.join(' ')
  };
}



/**
 * เรียกใช้จาก Apps Script Editor เพียงครั้งเดียวหลังเพิ่ม oauthScopes ใน appsscript.json
 * ฟังก์ชันนี้อ่าน Google Docs ต้นแบบเท่านั้น ไม่แก้ไขเอกสารและไม่สร้างไฟล์ใหม่
 */
function authorizeGoogleDocsReport() {
  const file = DriveApp.getFileById(APP.REPORT_TEMPLATE_DOC_ID);
  if (file.getMimeType() !== 'application/vnd.google-apps.document') {
    throw new Error('ไฟล์ต้นแบบที่กำหนดไม่ใช่ Google Docs');
  }

  const doc = DocumentApp.openById(APP.REPORT_TEMPLATE_DOC_ID);
  const tables = doc.getBody().getTables();

  if (tables.length !== 2 ||
      tables[0].getNumRows() < 3 ||
      tables[1].getNumRows() !== APP.REPORT_ROWS_PER_PAGE + 1 ||
      tables[1].getRow(0).getNumCells() !== 10) {
    throw new Error('เปิด Google Docs ได้แล้ว แต่โครงสร้างตารางต้นแบบไม่ตรงกับที่ระบบต้องใช้');
  }

  const result = {
    success: true,
    message: 'อนุญาตและเปิด Google Docs ต้นแบบได้แล้ว',
    documentName: doc.getName(),
    tableCount: tables.length
  };
  Logger.log(JSON.stringify(result));
  return result;
}