const DB_SHEET = 'Database';
const TRASH_SHEET = 'Trash';
const BACKUP_SHEET = 'Backup';
const INCOMING_SHEET = 'Incoming';

function normalizeDateKey_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  if (!value) return '';
  const text = value.toString().trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return text;

  let match = text.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (match) {
    let year = Number(match[3]);
    if (year > 2400) year -= 543;
    return `${String(year).padStart(4, '0')}-${match[2].padStart(2, '0')}-${match[1].padStart(2, '0')}`;
  }

  match = text.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
  if (match) {
    let year = Number(match[1]);
    if (year > 2400) year -= 543;
    return `${String(year).padStart(4, '0')}-${match[2].padStart(2, '0')}-${match[3].padStart(2, '0')}`;
  }

  const parsed = new Date(text);
  return isNaN(parsed.getTime()) ? text : Utilities.formatDate(parsed, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function formatThaiDate_(dateKey) {
  const normalized = normalizeDateKey_(dateKey);
  const match = normalized.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return normalized || '-';
  return `${match[3]}/${match[2]}/${Number(match[1]) + 543}`;
}
function doGet() {
  return HtmlService
    .createHtmlOutputFromFile('Index')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const raw = e && e.postData && e.postData.contents ? e.postData.contents : '{}';
    const payload = JSON.parse(raw);

    // เดิม: OR Queue / Chrome Extension ส่งข้อมูลเข้ามาโดยตรง
    // ใหม่: GitHub frontend ส่ง { action, data, token }
    if (payload && payload.action) {
      return jsonResponse_(handleApiRequest_(payload));
    }

    const result = importOrQueuePending(payload);
    return jsonResponse_(result);
  } catch (err) {
    return jsonResponse_({
      success: false,
      message: err && err.message ? err.message : 'รับข้อมูลไม่สำเร็จ'
    });
  }
}

/* =========================================================
   API สำหรับ GitHub Pages
   - ไม่กระทบ google.script.run ของเว็บเดิม
   - ใช้ POST แบบ text/plain เพื่อลดปัญหา CORS preflight
   ========================================================= */
function handleApiRequest_(payload) {
  const action = cleanText_(payload.action);
  const data = payload.data || {};

  if (action === 'ping') {
    return {
      success: true,
      service: 'OR Appointment API',
      serverTime: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')
    };
  }

  if (action === 'loginWithToken') {
    const result = loginWithToken(data.username, data.password);
    return result;
  }

  // import จาก OR Queue เดิมยังใช้ doPost แบบไม่มี action ได้เหมือนเดิม
  if (!isValidApiToken_(payload.token)) {
    return { success: false, message: 'UNAUTHORIZED' };
  }

  try {
    switch (action) {
      case 'getAppointmentsFast':
        return getAppointmentsFast();
      case 'saveToSheet':
        return saveToSheet(data);
      case 'getPatientHistory':
        return getPatientHistory(data.name);
      case 'getPendingImports':
        return getPendingImports();
      case 'approvePendingImport':
        return approvePendingImport(data.incomingId, data.edited || {});
      case 'rejectPendingImport':
        return rejectPendingImport(data.incomingId, data.note || '');
      case 'moveToTrash':
        return moveToTrash(data.id, data.deleteNote || '');
      case 'getTrashData':
        return getTrashData();
      case 'restoreTrashData':
        return restoreTrashData(data.id);
      case 'deleteTrashForever':
        return deleteTrashForever(data.id);
      case 'deleteAllTrashForever':
        return deleteAllTrashForever();
      default:
        return { success: false, message: 'ไม่รู้จัก API action: ' + action };
    }
  } catch (err) {
    return {
      success: false,
      message: err && err.message ? err.message : 'ดำเนินการไม่สำเร็จ'
    };
  }
}

function isValidApiToken_(token) {
  const cleanToken = cleanText_(token);
  if (!cleanToken) return false;
  const cache = CacheService.getScriptCache();
  return cache.get('OR_API_SESSION_' + cleanToken) === '1';
}

function cleanText_(value) {
  return value === null || value === undefined ? '' : value.toString().replace(/\s+/g, ' ').trim();
}

function normalizePhone_(value) {
  let phone = cleanText_(value).replace(/[^\d,]/g, '');
  if (phone.length === 9 && !phone.startsWith('0')) phone = '0' + phone;
  return phone;
}

function isInvalidProcedure_(value) {
  const text = cleanText_(value);
  if (!text) return true;
  if (/^\d{4}-\d{1,2}-\d{1,2}$/.test(text)) return true;
  if (/^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{4}$/.test(text)) return true;
  if (/^\d{1,2}[.:]\d{2}\s*[-–]\s*\d{1,2}[.:]\d{2}$/.test(text)) return true;
  if (/^[0-9\s,()\-]+$/.test(text)) return true;
  return false;
}

function ensureIncomingSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(INCOMING_SHEET);
  if (!sh) sh = ss.insertSheet(INCOMING_SHEET);

  const headers = [
    'IncomingId',
    'Status',
    'ReceivedAt',
    'AppointmentDate',
    'OperationTime',
    'Name',
    'HN',
    'CID',
    'Phone',
    'Coverage',
    'Doctor',
    'Procedure',
    'Source',
    'SourceKey',
    'RawJson',
    'ReviewedAt',
    'ReviewedBy',
    'DatabaseRowId',
    'RejectNote'
  ];

  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  } else {
    const existing = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), headers.length)).getValues()[0];
    headers.forEach((header, index) => {
      if (existing[index] !== header) sh.getRange(1, index + 1).setValue(header);
    });
  }
  return sh;
}

function normalizeIncomingItem_(item) {
  const normalized = {
    appointmentDate: normalizeDateKey_(item.appointmentDate || item.date),
    operationTime: cleanText_(item.operationTime || item.time),
    name: cleanText_(item.name),
    hn: cleanText_(item.hn),
    cid: cleanText_(item.cid),
    phone: normalizePhone_(item.phone),
    coverage: cleanText_(item.coverage),
    doctor: cleanText_(item.doctor),
    procedure: cleanText_(item.procedure),
    source: cleanText_(item.source || 'OR Queue Chrome Extension'),
    raw: item
  };

  const missing = [];
  if (!normalized.appointmentDate) missing.push('วันที่นัด');
  if (!normalized.name) missing.push('ชื่อ-สกุล');
  if (isInvalidProcedure_(normalized.procedure)) missing.push('หัตถการ');
  if (!normalized.hn && !normalized.cid) missing.push('HN หรือ CID');

  normalized.valid = missing.length === 0;
  normalized.missing = missing;
  normalized.sourceKey = buildIncomingSourceKey_(normalized);
  return normalized;
}

function buildIncomingSourceKey_(item) {
  const patientKey = item.cid || item.hn || item.name;
  return [
    normalizeDateKey_(item.appointmentDate),
    cleanText_(patientKey).toLowerCase(),
    cleanText_(item.procedure).toLowerCase()
  ].join('|');
}

function importOrQueuePending(payload) {
  const items = Array.isArray(payload) ? payload : (Array.isArray(payload.items) ? payload.items : [payload]);
  if (!items.length) return { success: false, message: 'ไม่มีข้อมูลสำหรับนำเข้า' };

  const sh = ensureIncomingSheet_();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const lastRow = sh.getLastRow();
    const existingRows = lastRow > 1 ? sh.getRange(2, 1, lastRow - 1, 19).getValues() : [];
    const existingKeys = {};
    existingRows.forEach(row => {
      const status = cleanText_(row[1]);
      const key = cleanText_(row[13]);
      if (key && status !== 'REJECTED') existingKeys[key] = true;
    });

    const databaseKeys = {};
    getAppointmentsFast().forEach(a => {
      const key = buildIncomingSourceKey_({
        appointmentDate: a.date,
        name: a.name,
        hn: '',
        cid: a.cid,
        procedure: a.procedure
      });
      if (key) databaseKeys[key] = true;
    });

    const rows = [];
    const skipped = [];
    const invalid = [];
    const now = new Date();

    items.forEach(item => {
      const normalized = normalizeIncomingItem_(item || {});
      if (!normalized.valid) {
        invalid.push({
          name: normalized.name,
          hn: normalized.hn,
          cid: normalized.cid,
          procedure: normalized.procedure,
          missing: normalized.missing
        });
        return;
      }
      if (existingKeys[normalized.sourceKey] || databaseKeys[normalized.sourceKey]) {
        skipped.push({
          name: normalized.name,
          hn: normalized.hn,
          cid: normalized.cid,
          procedure: normalized.procedure
        });
        return;
      }
      existingKeys[normalized.sourceKey] = true;
      rows.push([
        'INC-' + Date.now() + '-' + Math.floor(Math.random() * 100000),
        'PENDING',
        now,
        normalized.appointmentDate,
        normalized.operationTime,
        normalized.name,
        normalized.hn,
        normalized.cid,
        normalized.phone ? "'" + normalized.phone : '',
        normalized.coverage,
        normalized.doctor,
        normalized.procedure,
        normalized.source,
        normalized.sourceKey,
        JSON.stringify(normalized.raw),
        '',
        '',
        '',
        ''
      ]);
    });

    if (rows.length) {
      sh.getRange(sh.getLastRow() + 1, 1, rows.length, 19).setValues(rows);
    }

    return {
      success: true,
      imported: rows.length,
      skipped: skipped.length,
      invalid: invalid.length,
      skippedItems: skipped,
      invalidItems: invalid
    };
  } finally {
    lock.releaseLock();
  }
}

function getPendingImports() {
  const sh = ensureIncomingSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const values = sh.getRange(2, 1, lastRow - 1, 19).getValues();
  return values
    .filter(row => cleanText_(row[1]) === 'PENDING')
    .map(row => ({
      incomingId: row[0],
      status: row[1],
      receivedAt: formatDateTime_(row[2]),
      date: normalizeDateKey_(row[3]),
      operationTime: cleanText_(row[4]),
      name: cleanText_(row[5]),
      hn: cleanText_(row[6]),
      cid: cleanText_(row[7]),
      phone: cleanText_(row[8]).replace(/^'/, ''),
      coverage: cleanText_(row[9]),
      doctor: cleanText_(row[10]),
      procedure: cleanText_(row[11]),
      source: cleanText_(row[12])
    }));
}

function findIncomingRow_(incomingId) {
  const sh = ensureIncomingSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return null;
  const ids = sh.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (cleanText_(ids[i][0]) === cleanText_(incomingId)) {
      return { sheet: sh, rowNumber: i + 2 };
    }
  }
  return null;
}

function approvePendingImport(incomingId, edited) {
  const found = findIncomingRow_(incomingId);
  if (!found) return { success: false, message: 'ไม่พบรายการรอตรวจสอบ' };

  const row = found.sheet.getRange(found.rowNumber, 1, 1, 19).getValues()[0];
  if (cleanText_(row[1]) !== 'PENDING') {
    return { success: false, message: 'รายการนี้ไม่ได้อยู่ในสถานะรอตรวจสอบ' };
  }

  const obj = {
    rowId: -1,
    date: edited && edited.date ? edited.date : row[3],
    name: edited && edited.name ? edited.name : row[5],
    cid: edited && edited.cid !== undefined ? edited.cid : row[7],
    phone: edited && edited.phone !== undefined ? edited.phone : row[8],
    doctor: edited && edited.doctor !== undefined ? edited.doctor : row[10],
    procedure: edited && edited.procedure ? edited.procedure : row[11]
  };

  const result = saveToSheet(obj);
  if (!result.success) return result;

  found.sheet.getRange(found.rowNumber, 2).setValue('APPROVED');
  found.sheet.getRange(found.rowNumber, 16).setValue(new Date());
  found.sheet.getRange(found.rowNumber, 17).setValue(Session.getActiveUser().getEmail() || 'user');
  found.sheet.getRange(found.rowNumber, 18).setValue(result.rowId || '');
  return { success: true };
}

function rejectPendingImport(incomingId, note) {
  const found = findIncomingRow_(incomingId);
  if (!found) return { success: false, message: 'ไม่พบรายการรอตรวจสอบ' };
  found.sheet.getRange(found.rowNumber, 2).setValue('REJECTED');
  found.sheet.getRange(found.rowNumber, 16).setValue(new Date());
  found.sheet.getRange(found.rowNumber, 17).setValue(Session.getActiveUser().getEmail() || 'user');
  found.sheet.getRange(found.rowNumber, 19).setValue(cleanText_(note));
  return { success: true };
}

/* LOGIN */
function loginWithToken(username,password){
  const USER = 'admin';
  const PASS = '1234';
  if(username === USER && password === PASS){
    // สร้าง session token ใหม่ทุกครั้งที่ login
    const token = Utilities.getUuid().replace(/-/g, '');
    CacheService.getScriptCache().put('OR_API_SESSION_' + token, '1', 21600); // 6 ชั่วโมง
    return { success:true, token:token };
  }
  return { success:false };
}

/* GET DATA - ปรับปรุงเพื่อรองรับเคสข้อมูลเดิมที่ช่องคอลัมน์ A เป็นค่าว่าง */
function getAppointmentsFast() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(DB_SHEET);
  const lastRow = sh.getLastRow();

  if(lastRow < 2){
    return [];
  }

  const values = sh.getRange(2, 1, lastRow - 1, 9).getValues();
  return values.map((r, index) => {
    let dateStr = normalizeDateKey_(r[2] || r[1]);

    let phoneStr = r[7] ? r[7].toString().trim() : '';
    if (phoneStr.length === 9 && !phoneStr.startsWith('0')) {
      phoneStr = '0' + phoneStr;
    }

    // หากแถวข้อมูลเดิมไม่มี rowID ให้จัดรหัสอ้างอิงสำรองชั่วคราวให้ตรงลำดับแถว
    let uniqueId = r[0] ? r[0].toString().trim() : (index + 2).toString();

    return {
      rowId: uniqueId, 
      date: dateStr,
      name: r[3] || '',
      cid: r[4] ? r[4].toString().trim() : '',
      doctor: r[5] ? r[5].toString().trim() : '',
      procedure: r[6] ? r[6].toString().trim() : '',
      phone: phoneStr
    };
  });
}

/* SAVE */
function saveToSheet(obj) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(DB_SHEET);
  const data = sh.getDataRange().getValues();
  const targetDateStr = normalizeDateKey_(obj.date);
  
  const duplicate = data.find((r, i) => {
    if (i === 0) return false;
    let rDate = normalizeDateKey_(r[2] || r[1]);
    return (rDate == targetDateStr && r[3] == obj.name && r[6] == obj.procedure);
  });
  
  if (duplicate && obj.rowId == -1) {
    return { success: false, message: 'ข้อมูลซ้ำ' };
  }
  
  let phoneToSave = obj.phone ? obj.phone.toString().trim() : '';
  if (phoneToSave && !phoneToSave.startsWith("'")) {
    if (phoneToSave.length === 9 && !phoneToSave.startsWith('0')) {
      phoneToSave = '0' + phoneToSave;
    }
    phoneToSave = "'" + phoneToSave; 
  }

  let savedRowId = obj.rowId;
  if (obj.rowId == -1) {
    savedRowId = 'REC-' + Date.now();
    sh.appendRow([
      savedRowId,
      targetDateStr, 
      targetDateStr, 
      obj.name,
      obj.cid,
      obj.doctor,
      obj.procedure,
      phoneToSave,
      new Date()
    ]);
  } else {
    for (let i = 1; i < data.length; i++) {
      let currentId = data[i][0] ? data[i][0].toString().trim() : (i + 1).toString();
      if (currentId == obj.rowId.toString()) {
        sh.getRange(i + 1, 2, 1, 8).setValues([[
            targetDateStr,
            targetDateStr,
            obj.name,
            obj.cid,
            obj.doctor,
            obj.procedure,
            phoneToSave,
            new Date()
        ]]);
        break;
      }
    }
  }

  CacheService.getScriptCache().remove('appointments');
  return { success: true, rowId: savedRowId };
}

/* DELETE TO TRASH - ปรับปรุงระบบค้นหาให้สามารถลบแถวที่ไม่มีรหัส rowID ได้แม่นยำ */
function formatDateTime_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  }
  return value ? value.toString() : '';
}

function ensureTrashMetadataHeaders_(trashSheet) {
  const maxColumns = trashSheet.getMaxColumns();
  if (maxColumns < 11) {
    trashSheet.insertColumnsAfter(maxColumns, 11 - maxColumns);
  }
  const headers = trashSheet.getRange(1, 10, 1, 2).getValues()[0];
  if (!headers[0]) trashSheet.getRange(1, 10).setValue('DeletedAt');
  if (!headers[1]) trashSheet.getRange(1, 11).setValue('DeleteNote');
}

function moveToTrash(id, deleteNote) {
  const ss = SpreadsheetApp.getActive();
  const db = ss.getSheetByName(DB_SHEET);
  const trash = ss.getSheetByName(TRASH_SHEET);
  ensureTrashMetadataHeaders_(trash);
  const data = db.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    let currentId = data[i][0] ? data[i][0].toString().trim() : (i + 1).toString();
    if (currentId == id.toString()) {
      trash.appendRow([
        ...data[i].slice(0, 9),
        new Date(),
        deleteNote ? deleteNote.toString().trim() : ''
      ]);
      db.deleteRow(i + 1);
      break;
    }
  }
  CacheService.getScriptCache().remove('appointments');
  return { success: true };
}

/* GET TRASH - แก้ไขคอลัมน์วันที่นัดหมายจากดัชนี B (r[1]) ให้เป็น ดัชนี C (r[2]) ตามจริงเพื่อให้ข้อมูลแสดงผล */
function getTrashData() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(TRASH_SHEET);
  const lastRow = sh.getLastRow();

  if (lastRow < 2) return [];
  ensureTrashMetadataHeaders_(sh);
  const data = sh.getRange(2, 1, lastRow - 1, Math.max(11, sh.getLastColumn())).getValues();

  return data.map((r, index) => {
    let dateStr = normalizeDateKey_(r[2] || r[1]);

    let timeStr = formatDateTime_(r[8]);
    let deletedAtStr = formatDateTime_(r[9]);

    // ล็อกรหัสผ่านแถวกรณีเคสเก่าย้ายมาแล้วไม่มีรหัส
    let uniqueId = r[0] ? r[0].toString().trim() : "TRASH-" + (index + 2).toString();

    return {
      rowId: uniqueId,
      date: dateStr,
      name: r[3] || '',
      cid: r[4] ? r[4].toString().trim() : '',
      doctor: r[5] ? r[5].toString().trim() : '',
      procedure: r[6] ? r[6].toString().trim() : '',
      phone: r[7] ? r[7].toString().trim() : '',
      timestamp: timeStr,
      deletedAt: deletedAtStr,
      deleteNote: r[10] ? r[10].toString().trim() : ''
    };
  });
}

/* RESTORE - ปรับปรุงให้รองรับเงื่อนไขการตรวจสอบกู้คืนข้อมูลแบบใหม่ ไร้รอยต่อ */
function restoreTrashData(id) {
  const ss = SpreadsheetApp.getActive();
  const db = ss.getSheetByName(DB_SHEET);
  const trash = ss.getSheetByName(TRASH_SHEET);
  const data = trash.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    let currentId = data[i][0] ? data[i][0].toString().trim() : "TRASH-" + (i + 1).toString();
    if (currentId == id.toString()) {
      let rowToRestore = data[i];
      // ก่อนส่งกลับ หากไม่มีรหัสคอลัมน์ A ให้ติดรหัสถาวรชุดใหม่ให้ทันทีเพื่อเปิดแสดงผลในหน้าหลักได้ทันที
      if (!rowToRestore[0] || rowToRestore[0].toString().trim() === '') {
        rowToRestore[0] = 'REC-' + new Date().getTime();
      }
      db.appendRow(rowToRestore.slice(0, 9));
      trash.deleteRow(i + 1);
      break;
    }
  }
  CacheService.getScriptCache().remove('appointments');
  return { success: true };
}

/* DELETE FOREVER */
function deleteTrashForever(id) {
  const ss = SpreadsheetApp.getActive();
  const trash = ss.getSheetByName(TRASH_SHEET);
  const data = trash.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    let currentId = data[i][0] ? data[i][0].toString().trim() : "TRASH-" + (i + 1).toString();
    if (currentId == id.toString()) {
      trash.deleteRow(i + 1);
      break;
    }
  }
  return { success: true };
}

function deleteAllTrashForever() {
  const ss = SpreadsheetApp.getActive();
  const trash = ss.getSheetByName(TRASH_SHEET);
  const lastRow = trash.getLastRow();
  if (lastRow > 1) {
    trash.deleteRows(2, lastRow - 1);
  }
  return { success: true };
}

/* AUTO HISTORY */
function getPatientHistory(name) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(DB_SHEET);
  const data = sh.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][3] == name) {
      return {
        cid: data[i][4],     
        phone: data[i][7]    
      };
    }
  }
  return null;
}

/* BACKUP EVERY 4 MONTHS */
function parseAppointmentDate_(value) {
  const dateKey = normalizeDateKey_(value);
  const match = dateKey.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return null;
  return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
}

function getOrCreateBackupSheet_(ss, dbSheet) {
  let backup = ss.getSheetByName(BACKUP_SHEET);
  if (!backup) {
    backup = ss.insertSheet(BACKUP_SHEET);
  }
  if (backup.getLastRow() === 0) {
    const headers = dbSheet.getRange(1, 1, 1, 9).getValues()[0];
    backup.getRange(1, 1, 1, 10).setValues([[...headers, 'BackupAt']]);
  }
  return backup;
}

function shouldRunBackupNow_(now) {
  const props = PropertiesService.getScriptProperties();
  const lastRunText = props.getProperty('lastBackupRunAt');
  if (!lastRunText) return true;
  const lastRun = new Date(lastRunText);
  if (isNaN(lastRun.getTime())) return true;
  const nextRun = new Date(lastRun.getFullYear(), lastRun.getMonth() + 4, lastRun.getDate());
  return now >= nextRun;
}

function backupAppointmentsEvery4Months(forceRun) {
  const now = new Date();
  if (!forceRun && !shouldRunBackupNow_(now)) {
    return { success: true, moved: 0, skipped: true };
  }
  const ss = SpreadsheetApp.getActive();
  const db = ss.getSheetByName(DB_SHEET);
  const backup = getOrCreateBackupSheet_(ss, db);
  const lastRow = db.getLastRow();
  if (lastRow < 2) {
    PropertiesService.getScriptProperties().setProperty('lastBackupRunAt', now.toISOString());
    return { success: true, moved: 0 };
  }
  const cutoff = new Date(now.getFullYear(), now.getMonth() - 4, now.getDate());
  const data = db.getRange(2, 1, lastRow - 1, 9).getValues();
  const rowsToBackup = [];
  const rowsToDelete = [];

  data.forEach((row, index) => {
    const appDate = parseAppointmentDate_(row[2]);
    if (appDate && appDate < cutoff) {
      rowsToBackup.push([...row, now]);
      rowsToDelete.push(index + 2);
    }
  });
  if (rowsToBackup.length > 0) {
    backup.getRange(backup.getLastRow() + 1, 1, rowsToBackup.length, 10).setValues(rowsToBackup);
    rowsToDelete.reverse().forEach(rowNumber => db.deleteRow(rowNumber));
    CacheService.getScriptCache().remove('appointments');
  }
  PropertiesService.getScriptProperties().setProperty('lastBackupRunAt', now.toISOString());
  return { success: true, moved: rowsToBackup.length };
}

function setupBackupTriggerEvery4Months() {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'backupAppointmentsEvery4Months') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  ScriptApp.newTrigger('backupAppointmentsEvery4Months').timeBased().onMonthDay(1).atHour(2).create();
  return { success: true, schedule: 'ตรวจทุกวันที่ 1 เวลา 02:00 และสำรองจริงเมื่อครบ 4 เดือน' };
}

function runBackupNow() {
  return backupAppointmentsEvery4Months(true);
}

/* LINE OA DAILY REPORT */
function formatTodayKey_() {
  return normalizeDateKey_(new Date());
}

function getReportDateKey_(targetDate) {
  if (targetDate instanceof Date) return normalizeDateKey_(targetDate);
  if (typeof targetDate === 'string' || typeof targetDate === 'number') return normalizeDateKey_(targetDate);
  return formatTodayKey_();
}
function buildDailyLineReport_(targetDate) {
  const reportDate = getReportDateKey_(targetDate);
  const data = getAppointmentsFast().filter(a => normalizeDateKey_(a.date) === reportDate);
  const groups = {};
  let total = 0;

  data.forEach(a => {
    const doctor = (a.doctor || '').trim() || 'ไม่ระบุแพทย์';
    const proc = (a.procedure || '').trim() || 'ไม่ระบุหัตถการ';
    if (!groups[doctor]) groups[doctor] = { total: 0, procedures: {}, patients: [] };
    groups[doctor].total += 1;
    groups[doctor].procedures[proc] = (groups[doctor].procedures[proc] || 0) + 1;
    groups[doctor].patients.push(a);
    total += 1;
  });

  let message = `สรุปเคส OR Appointment\nวันที่: ${formatThaiDate_(reportDate)}\nรวมทั้งหมด: ${total} ราย`;

  if (total === 0) {
    return `${message}\n\nวันนี้ไม่มีเคสนัดหมาย`;
  }

  Object.keys(groups).sort().forEach(doctor => {
    const group = groups[doctor];
    message += `\n\n${doctor} (${group.total} ราย)`;
    Object.keys(group.procedures).sort().forEach(proc => {
      message += `\n- ${proc}: ${group.procedures[proc]} ราย`;
    });
  });
  return message;
}

function sendLineOaText_(text) {
  const token = PropertiesService.getScriptProperties().getProperty('LINE_TOKEN');
  if (!token) {
    throw new Error('ไม่พบ LINE_TOKEN ใน Script Properties');
  }
  const response = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/broadcast', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: `Bearer ${token}` },
    payload: JSON.stringify({ messages: [{ type: 'text', text }] }),
    muteHttpExceptions: true
  });
  const code = response.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error(`ส่ง LINE OA ไม่สำเร็จ (${code}): ${response.getContentText()}`);
  }
  return { success: true, statusCode: code };
}

function sendDailyLineOaReport(targetDate) {
  const isManualDate =
    targetDate instanceof Date ||
    typeof targetDate === 'string' ||
    typeof targetDate === 'number';

  const checkDate = isManualDate ? new Date(targetDate) : new Date();
  const reportDate = getReportDateKey_(checkDate);

  Logger.log('SCRIPT TIMEZONE = ' + Session.getScriptTimeZone());
  Logger.log('REPORT DATE = ' + reportDate);
  Logger.log(
    'RUN TIME = ' +
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')
  );

  const holiday = isHoliday_(checkDate);
  Logger.log('IS HOLIDAY = ' + holiday);

  if (holiday) {
    Logger.log('SKIP LINE OA');
    return {
      success: true,
      skipped: true,
      reportDate: reportDate
    };
  }

  Logger.log('SEND LINE OA');
  return sendLineOaFlexReport_(reportDate);
}

function setupDailyLineOaReportTrigger() {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'sendDailyLineOaReport') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  ScriptApp.newTrigger('sendDailyLineOaReport').timeBased().everyDays(1).atHour(7).nearMinute(0).create();
  return { success: true, schedule: 'ส่ง LINE OA ทุกวันเวลา 07:00 น.' };
}

function clearCache(){
  CacheService.getScriptCache().removeAll([]);
}

// ฟังก์ชันติดตั้งและซิงก์ล็อกรหัสแถวให้กรณีเคสเก่าไม่มี rowID ทั้งหน้าหลักและถังขยะ
function SAFE_fixAndSyncRowIDs() {
  const ss = SpreadsheetApp.getActive();
  const dbSheet = ss.getSheetByName(DB_SHEET);
  if (dbSheet) {
    const lastRow = dbSheet.getLastRow();
    if (lastRow >= 2) {
      const range = dbSheet.getRange(2, 1, lastRow - 1, 1);
      const values = range.getValues();
      for (let i = 0; i < values.length; i++) {
        if (!values[i][0] || values[i][0].toString().trim() === '') {
          values[i][0] = (i + 2).toString(); 
        }
      }
      range.setValues(values);
    }
  }
  const trashSheet = ss.getSheetByName(TRASH_SHEET);
  if (trashSheet) {
    const lastRowTrash = trashSheet.getLastRow();
    if (lastRowTrash >= 2) {
      const rangeTrash = trashSheet.getRange(2, 1, lastRowTrash - 1, 1);
      const valuesTrash = rangeTrash.getValues();
      for (let i = 0; i < valuesTrash.length; i++) {
        if (!valuesTrash[i][0] || valuesTrash[i][0].toString().trim() === '') {
          valuesTrash[i][0] = "TRASH-" + (i + 2).toString();
        }
      }
      rangeTrash.setValues(valuesTrash);
    }
  }
  CacheService.getScriptCache().remove('appointments');
  return "ซิงก์ข้อมูลสำเร็จ";
}

function syncThaiHolidays() {

  const CALENDAR_ID =
    'th.th#holiday@group.v.calendar.google.com';

  const ss = SpreadsheetApp.getActive();

  let sh = ss.getSheetByName('Holiday');

  if (!sh) {
    sh = ss.insertSheet('Holiday');
  }

  sh.clear();

  sh.getRange(1,1).setValue('Date');

  const calendar =
    CalendarApp.getCalendarById(CALENDAR_ID);

  const startDate = new Date(2025,0,1);
  const endDate   = new Date(2030,11,31);

  const events =
    calendar.getEvents(startDate,endDate);

  const rows = [];

  events.forEach(event => {

    rows.push([
      Utilities.formatDate(
        event.getStartTime(),
        Session.getScriptTimeZone(),
        'yyyy-MM-dd'
      )
    ]);

  });

  rows.sort();

  if(rows.length){
    sh.getRange(2,1,rows.length,1)
      .setValues(rows);
  }

  Logger.log('โหลดวันหยุด ' + rows.length + ' วัน');
}

function isHoliday_(dateObj) {

  const day = dateObj.getDay();

  // อาทิตย์ = 0
  // เสาร์ = 6
  if (day === 0 || day === 6) {
    return true;
  }

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Holiday');

  if (!sh) {
    return false;
  }

  const targetDate = normalizeDateKey_(dateObj);

  const lastRow = sh.getLastRow();

  if (lastRow < 2) {
    return false;
  }

  const holidayList = sh
    .getRange(2, 1, lastRow - 1, 1)
    .getValues()
    .flat()
    .map(d => normalizeDateKey_(d));

  return holidayList.includes(targetDate);
}

function testHolidayCheck() {

  Logger.log(
    isHoliday_(new Date())
  );

}

function testWorkingDay() {
  Logger.log(
    isHoliday_(new Date('2026-06-01'))
  );
}
function sendLineOaFlexReport_(targetDate) {

  const token = PropertiesService
    .getScriptProperties()
    .getProperty('LINE_TOKEN');

  if (!token) {
    throw new Error('ไม่พบ LINE_TOKEN');
  }

  const reportDate = getReportDateKey_(targetDate);

  const data = getAppointmentsFast().filter(function(a) {
    return normalizeDateKey_(a.date) === reportDate;
  });

  const groups = {};
  let total = 0;

  data.forEach(function(a) {

    const doctor =
      (a.doctor || '').trim() || 'ไม่ระบุแพทย์';

    const proc =
      (a.procedure || '').trim() || 'ไม่ระบุหัตถการ';

    if (!groups[doctor]) {
      groups[doctor] = {
        total: 0,
        procedures: {}
      };
    }

    groups[doctor].total++;

    groups[doctor].procedures[proc] =
      (groups[doctor].procedures[proc] || 0) + 1;

    total++;

  });

  const doctorBoxes = [];

  Object.keys(groups).sort().forEach(function(doctor) {

    const contents = [];

    contents.push({
      type: "text",
      text: doctor + " (" + groups[doctor].total + " ราย)",
      weight: "bold",
      size: "md",
      color: "#1B5E20"
    });

    Object.keys(groups[doctor].procedures)
      .sort()
      .forEach(function(proc) {

        contents.push({
          type: "box",
          layout: "baseline",
          spacing: "sm",
          contents: [
            {
              type: "text",
              text: proc,
              size: "sm",
              color: "#555555",
              flex: 4,
              wrap: true
            },
            {
              type: "text",
              text: String(
                groups[doctor].procedures[proc]
              ) + " ราย",
              size: "sm",
              weight: "bold",
              align: "end",
              flex: 2
            }
          ]
        });

      });

    contents.push({
      type: "separator",
      margin: "md"
    });

    doctorBoxes.push({
      type: "box",
      layout: "vertical",
      margin: "lg",
      spacing: "sm",
      contents: contents
    });

  });

  const bodyContents = [];

  bodyContents.push({
    type: "text",
    text: "วันที่ " + formatThaiDate_(reportDate),
    size: "md",
    weight: "bold"
  });

  bodyContents.push({
    type: "box",
    layout: "vertical",
    backgroundColor: "#E8F5E9",
    cornerRadius: "8px",
    paddingAll: "10px",
    contents: [
      {
        type: "text",
        text: "TOTAL CASES",
        size: "xs",
        color: "#666666"
      },
      {
        type: "text",
        text: String(total),
        size: "xxl",
        weight: "bold",
        color: "#1B5E20"
      }
    ]
  });

  bodyContents.push({
    type: "separator"
  });

  if (doctorBoxes.length > 0) {

    doctorBoxes.forEach(function(box) {
      bodyContents.push(box);
    });

  } else {

    bodyContents.push({
      type: "text",
      text: "ไม่มีเคสนัดหมาย",
      color: "#999999",
      align: "center"
    });

  }

  const flexMessage = {
    type: "flex",
    altText:
      "Appointment Report " +
      formatThaiDate_(reportDate),

    contents: {
      type: "bubble",

      header: {
        type: "box",
        layout: "vertical",
        backgroundColor: "#1B5E20",
        paddingAll: "15px",

        contents: [
          {
  type: "text",
  text: "SUNGMEN HOSPITAL",
  color: "#FFFFFF",
  weight: "bold",
  size: "xl",
  align: "center"
},
{
  type: "text",
  text: "Operating Room Daily Report",
  color: "#FFFFFF",
  size: "sm",
  align: "center"
}
        ]
      },

      body: {
        type: "box",
        layout: "vertical",
        spacing: "md",
        contents: bodyContents
      }
    }
  };

  const response = UrlFetchApp.fetch(
    'https://api.line.me/v2/bot/message/broadcast',
    {
      method: 'post',
      contentType: 'application/json',
      headers: {
        Authorization: 'Bearer ' + token
      },
      payload: JSON.stringify({
        messages: [flexMessage]
      }),
      muteHttpExceptions: true
    }
  );

  const code = response.getResponseCode();
  const body = response.getContentText();

  Logger.log(body);

  if (code < 200 || code >= 300) {
    throw new Error(
      'LINE Error (' +
      code +
      '): ' +
      body
    );
  }

  return {
    success: true,
    statusCode: code
  };

}

function testFlexNow() {
  return sendLineOaFlexReport_(new Date());
}

function debugHolidayToday() {

  const today = new Date();

  Logger.log("Today = " + today);

  Logger.log(
    "DateKey = " +
    normalizeDateKey_(today)
  );

  Logger.log(
    "isHoliday = " +
    isHoliday_(today)
  );

}

function testHolidayTrigger() {

  return sendDailyLineOaReport(
    '2026-06-01'
  );

}

