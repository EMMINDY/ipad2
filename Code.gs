// ==========================================
// 1. ส่วนตั้งค่าระบบ (CONFIGURATION)
// ==========================================

const SPREADSHEET_ID = '1Kr1GOn5F8rBNJGA7_Sqp4be7JirrRvVci0AGhtkA5hQ'; // ID ของ Google Sheet
const FOLDER_ID = '1pPtZlI8XYBle02byB5lthAhtLX8012Pa'; // ID ของ Google Drive Folder
const FOLDER_RETURN_ID = '16Rn35Lv0gC3HRt2ohUWmi_DdQShRcUfn'; // โฟลเดอร์เก็บเอกสารการคืน iPad

const SHEET_NAMES = {
  STUDENTS: [
    'รายชื่อนักเรียน ม.3', 
    'รายชื่อนักเรียน ม.4', 
    'รายชื่อนักเรียน ม.5', 
    'รายชื่อนักเรียน ม.6'
  ],
  TEACHERS: 'รายชื่อครู',
  ASSETS: 'รายงานทะเบียนทรัพย์สิน',
  ALL_NAMES: 'รายชื่อทั้งหมด',
  DATA_DB: 'ข้อมูล',
  LOGS: 'Log',
  ADMIN: 'แอดมิน',
  ADVISOR: 'ครูที่ปรึกษา'
};

// ==========================================
// 2. ฟังก์ชันพื้นฐาน & Helper
// ==========================================

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('ระบบจัดเก็บเอกสารการยืม-คืน ไอแพด โรงเรียนอรัญประเทศ')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) { 
  return HtmlService.createHtmlOutputFromFile(filename).getContent(); 
}

/** เคลียร์ Cache */
function invalidateSystemDataCache() {
  try {
    const cache = CacheService.getScriptCache();
    const keys = ['systemData_' + SPREADSHEET_ID,
                  'systemData_' + SPREADSHEET_ID + '_meta',
                  'quickStats_'  + SPREADSHEET_ID];
    for (let i = 1; i <= 20; i++) keys.push('systemData_' + SPREADSHEET_ID + '_' + i);
    // removeAll รองรับ array สูงสุด 100 keys
    cache.removeAll(keys);
  } catch (_) {}
}

/**
 * Warm-up: โหลดข้อมูลล่วงหน้าเข้า Cache
 * รันผ่าน Time Trigger ทุก 5 นาที เพื่อให้ผู้ใช้ได้ข้อมูลจาก cache เสมอ
 */
function warmUpCache() {
  try {
    invalidateSystemDataCache();
    const data = getAllSystemData();
    Logger.log('warmUpCache OK: ' + data.length + ' records');
  } catch (e) {
    Logger.log('warmUpCache error: ' + e.toString());
  }
}

/**
 * สร้าง Time Trigger สำหรับ warm-up cache ทุก 5 นาที
 * รัน createWarmUpTrigger() ครั้งเดียวจาก Apps Script Editor
 */
function createWarmUpTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'warmUpCache' ||
        t.getHandlerFunction() === 'refreshAndSyncData') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('warmUpCache').timeBased().everyMinutes(5).create();
  Logger.log('✅ warmUpCache Trigger created (every 5 min)');
}

/** Helper: อ่าน cache แบบ chunk */
function _readCacheChunked(cache, cacheKey) {
  try {
    const metaRaw = cache.get(cacheKey + '_meta');
    if (metaRaw) {
      const meta = JSON.parse(metaRaw);
      const keys = [];
      for (let c = 1; c <= meta.chunks; c++) keys.push(cacheKey + '_' + c);
      const parts = cache.getAll(keys);
      let combined = '';
      for (let c = 1; c <= meta.chunks; c++) {
        const p = parts[cacheKey + '_' + c];
        if (!p) return null;
        combined += p;
      }
      return JSON.parse(combined);
    }
    const cached = cache.get(cacheKey);
    if (cached) return JSON.parse(cached);
  } catch (_) {}
  return null;
}

/** Helper: เขียน cache แบบ chunk + putAll */
function _writeCacheChunked(cache, cacheKey, data, ttl) {
  try {
    const serialized = JSON.stringify(data);
    const CHUNK = 90000;
    if (serialized.length <= CHUNK) {
      cache.put(cacheKey, serialized, ttl);
    } else {
      const total = Math.ceil(serialized.length / CHUNK);
      const toStore = {};
      for (let c = 0; c < total; c++) {
        toStore[cacheKey + '_' + (c+1)] = serialized.slice(c*CHUNK, (c+1)*CHUNK);
      }
      toStore[cacheKey + '_meta'] = JSON.stringify({ chunks: total });
      cache.putAll(toStore, ttl);
    }
  } catch (_) {}
}

/**
 * สถิติเบื้องต้นสำหรับ Dashboard (เร็วมาก — cache แยก 5 นาที)
 */
function getQuickStats() {
  const cache = CacheService.getScriptCache();
  const key   = 'quickStats_' + SPREADSHEET_ID;
  try {
    const cached = cache.get(key);
    if (cached) return JSON.parse(cached);
  } catch (_) {}

  const data = getAllSystemData();
  const stats = { total: data.length, borrowed: 0, notBorrowed: 0, returned: 0, levels: {} };
  data.forEach(function(p) {
    const s = p.borrowStatus || '';
    if (s === 'ยืมอยู่' || s === 'อยู่ระหว่างการส่งคืน') stats.borrowed++;
    else if (s === 'คืนแล้ว') stats.returned++;
    else stats.notBorrowed++;
    const src = p.source_sheet || '-';
    stats.levels[src] = (stats.levels[src] || 0) + 1;
  });
  try { cache.put(key, JSON.stringify(stats), 300); } catch (_) {}
  return stats;
}

/** ฟังก์ชันปรับชื่อให้สะอาดที่สุดเพื่อการจับคู่ที่แม่นยำ */
function normalizeName(name) {
  if (!name) return "";
  let n = name.toString().normalize('NFC');
  const titleRegex = /^(?:ว่าที่\s*ร(?:้อย)?\.?\s*[ตทพ]\.?|จ(?:่า)?\.?ส(?:ิบ)?\.?[อทต]\.?|ส(?:ิบ)?\.?[อทต]\.?|พล(?:ทหาร)?\.?|ส\.อ\.|จ\.ส\.อ\.|ร\.ต\.|ดร\.?|ผศ\.?|รศ\.?|ศ\.?|เด็กชาย|เด็กหญิง|นางสาว|ด\.?\s*ช\.?|ด\.?\s*ญ\.?|น\.?\s*ส\.?|นาย|นาง|ครู|อ\.?|mr\.?|mrs\.?|ms\.?|miss)[\s\.]*/gi;
  n = n.replace(titleRegex, ''); 
  n = n.replace(/[^ก-๙a-zA-Z]/g, ''); 
  return n.toLowerCase();
}

/** Levenshtein Distance สำหรับ Fuzzy Match */
function getEditDistance(a, b) {
  if (a.length === 0) return b.length; 
  if (b.length === 0) return a.length; 
  var matrix = [];
  for (var i = 0; i <= b.length; i++) { matrix[i] = [i]; }
  for (var j = 0; j <= a.length; j++) { matrix[0][j] = j; }
  for (var i = 1; i <= b.length; i++) {
    for (var j = 1; j <= a.length; j++) {
      if (b.charAt(i - 1) == a.charAt(j - 1)) {
        matrix[i][j] = matrix[i - 1][j - 1];
      } else {
        matrix[i][j] = Math.min(
          matrix[i - 1][j - 1] + 1,
          Math.min(matrix[i][j - 1] + 1, matrix[i - 1][j] + 1)
        );
      }
    }
  }
  return matrix[b.length][a.length];
}

// ==========================================
// 3. ฟังก์ชันดึงข้อมูลหลัก (MAIN DATA ENGINE)
// ==========================================

function getAllSystemData() {
  const cache    = CacheService.getScriptCache();
  const cacheKey = 'systemData_' + SPREADSHEET_ID;

  // ── อ่าน Cache ก่อนเสมอ (ไม่ต้องรอ lock) ──
  const cached1 = _readCacheChunked(cache, cacheKey);
  if (cached1 && cached1.length > 0) return cached1;

  // ── พยายามขอ Lock เพื่อกัน request ซ้ำซ้อน ──
  // ใช้ getScriptLock (รองรับ Web App — getDocumentLock ใช้ได้เฉพาะ bound script)
  const lock = LockService.getScriptLock();
  let lockAcquired = false;
  try {
    // รอสูงสุด 10 วินาที — ถ้าไม่ได้ก็ไม่เป็นไร ทำงานต่อได้
    lock.waitLock(10000);
    lockAcquired = true;
  } catch (_) {
    // lock timeout: ตรวจ cache อีกครั้งก่อน proceed
    // (request ก่อนหน้าอาจเขียน cache เสร็จแล้วระหว่างที่เรารอ)
    const cachedRetry = _readCacheChunked(cache, cacheKey);
    if (cachedRetry && cachedRetry.length > 0) return cachedRetry;
    // ถ้ายังไม่มี: ดำเนินการต่อโดยไม่มี lock
    // (อาจเกิดการคำนวณซ้ำ แต่ดีกว่า throw error ใส่ผู้ใช้)
  }

  // ── ถ้าได้ lock: ตรวจ cache อีกครั้ง (request ก่อนหน้าอาจสร้างไว้แล้ว) ──
  if (lockAcquired) {
    try {
      const cached2 = _readCacheChunked(cache, cacheKey);
      if (cached2 && cached2.length > 0) {
        lock.releaseLock();
        return cached2;
      }
    } catch (_) {}
  }

  let allPeople = [];
  try {
    // ── เปิด SpreadsheetApp ครั้งเดียว แล้วส่งต่อทุกฟังก์ชัน ──
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
    // --- A. ดึงข้อมูล Asset (ทะเบียนทรัพย์สิน) ---
    const assetSheet = ss.getSheetByName(SHEET_NAMES.ASSETS);
    let assetMap = {};
    let assetKeys = [];

    if (assetSheet) {
      const lastRow = assetSheet.getLastRow();
      if (lastRow > 1) {
        const assetData = assetSheet.getRange(1, 1, lastRow, 10).getValues();
        for (let i = 1; i < assetData.length; i++) {
          let rawName = assetData[i][4]; 
          let serial = assetData[i][2] || assetData[i][3]; 
          if (rawName) {
            let cName = normalizeName(rawName);
            if (cName.length > 0) {
              assetMap[cName] = { 
                serial: serial ? serial.toString() : '-', 
                status: 'ยืมอยู่' 
              };
              assetKeys.push(cName);
            }
          }
        }
      }
    }

    // --- B. ดึงข้อมูล Log (Database) ---
    const dbSheet = ss.getSheetByName(SHEET_NAMES.DATA_DB);
    let dbMap = {}; 
    
    if (dbSheet) {
      const dbLastRow = dbSheet.getLastRow();
      if (dbLastRow > 1) {
        const dbData = dbSheet.getRange(1, 1, dbLastRow, 14).getValues();
        for (let i = 1; i < dbData.length; i++) {
          let id = dbData[i][1];
          if (!id) continue;
          id = id.toString();

          let rowSerial = dbData[i][5] ? dbData[i][5].toString() : '';
          let rowStatus = dbData[i][8] ? dbData[i][8].toString() : '';
          let rowNote   = dbData[i][7] ? dbData[i][7].toString() : '';
          let rowAction = dbData[i][6] ? dbData[i][6].toString() : '';
          let hasFiles  = (dbData[i][9] || dbData[i][10] || dbData[i][11] || dbData[i][12]);
          
          if (!dbMap[id]) dbMap[id] = {
            borrowStatus: 'ยังไม่ยืม', docStatus: 'ยังไม่ส่ง', returnDocStatus: 'ยังไม่ส่ง',
            serial: '-', files: {}, returnDocUrl: '',
            docFinal: false, returnDocFinal: false
          };
          
          if (rowSerial && rowSerial !== '-' && rowSerial !== '') dbMap[id].serial = rowSerial;
          
          if (hasFiles) {
            dbMap[id].files = { 
              agreement: dbData[i][9], card_std: dbData[i][10], 
              card_parent: dbData[i][11], house: dbData[i][12], phone: dbData[i][13] 
            };
            if (!dbMap[id].docFinal && !rowStatus.includes('ADMIN') && !rowStatus.includes('ADVISOR')) {
              dbMap[id].docStatus = 'รอตรวจสอบ';
            }
          }

          if (rowStatus === 'เอกสารผ่าน' || rowStatus.includes('เอกสารผ่าน') || rowAction === 'ADVISOR_APPROVE') {
            dbMap[id].docStatus = 'เอกสารผ่าน';
            dbMap[id].docFinal = true;
          } else if (rowStatus === 'เอกสารไม่ผ่าน' || rowStatus.includes('ไม่ผ่าน')) {
            dbMap[id].docStatus = 'เอกสารไม่ผ่าน';
            dbMap[id].docFinal = true;
          } else if (!dbMap[id].docFinal) {
            if (rowStatus.includes('รอตรวจสอบเอกสาร') || rowStatus.includes('รอตรวจสอบ')) {
              dbMap[id].docStatus = 'รอตรวจสอบ';
            } else if (rowStatus === 'ยังไม่ส่ง') {
              dbMap[id].docStatus = 'ยังไม่ส่ง';
            }
          }

        if ((rowAction === 'ADVISOR_RETURN' || rowAction === 'USER_RETURN') && rowNote.indexOf('[หลักฐานการคืน]:') >= 0) {
            const match = rowNote.match(/\[หลักฐานการคืน\]:\s*(https?:\/\/[^\s\n]+)/);
            if (match) {
              dbMap[id].returnDocUrl = match[1].trim();
              // ✅ อัปเดต returnDocStatus (แยกออกจาก docStatus ของการยืม)
              if (!dbMap[id].returnDocFinal) {
                dbMap[id].returnDocStatus = 'รอตรวจสอบ';
              }
            }
          }

          // ✅ แอดมินกำหนดสถานะเอกสารคืนโดยตรง (action = "ADMIN...|RETURN_DOC")
          if (rowAction.includes('|RETURN_DOC')) {
            if (rowStatus === 'เอกสารผ่าน' || rowStatus.includes('เอกสารผ่าน')) {
              dbMap[id].returnDocStatus = 'เอกสารผ่าน';
              dbMap[id].returnDocFinal = true;
            } else if (rowStatus === 'เอกสารไม่ผ่าน' || rowStatus.includes('ไม่ผ่าน')) {
              dbMap[id].returnDocStatus = 'เอกสารไม่ผ่าน';
              dbMap[id].returnDocFinal = true;
            } else if (rowStatus === 'รอตรวจสอบ') {
              if (!dbMap[id].returnDocFinal) dbMap[id].returnDocStatus = 'รอตรวจสอบ';
            }
          }

          if (rowStatus.indexOf('อยู่ระหว่างการส่งคืน') >= 0) {
            dbMap[id].borrowStatus = 'อยู่ระหว่างการส่งคืน';
          } else if (rowStatus === 'คืนแล้ว') {
            dbMap[id].borrowStatus = 'คืนแล้ว';
            // ✅ อนุมัติคืน → เอกสารคืนผ่านอัตโนมัติ (ถ้ายังรอตรวจสอบอยู่)
            if (!dbMap[id].returnDocFinal && dbMap[id].returnDocStatus === 'รอตรวจสอบ') {
              dbMap[id].returnDocStatus = 'เอกสารผ่าน';
              dbMap[id].returnDocFinal = true;
            }
          } else if (rowStatus.includes('ยืมอยู่') || rowStatus === 'ยืมได้') {
            dbMap[id].borrowStatus = 'ยืมอยู่';
          } else if (rowStatus.includes('ซ่อม')) {
            dbMap[id].borrowStatus = 'ส่งซ่อม';
          } else if (rowStatus.includes('สละ')) {
            dbMap[id].borrowStatus = 'สละสิทธิ์';
          } else if (rowStatus === 'ยังไม่ยืม') {
            dbMap[id].borrowStatus = 'ยังไม่ยืม';
          }
        }
      }
    }

    // --- C. รวมข้อมูล (Merge + Fuzzy Match) ---
    const processPerson = (type, no, id, name, room, source) => {
      if (!name) return;
      id = id.toString();
      let cleanedName = normalizeName(name);
      
      let finalBorrow = 'ยังไม่ยืม';
      let finalDoc = 'ยังไม่ส่ง';
      let finalReturnDocStatus = 'ยังไม่ส่ง';
      let finalSerial = '-';
      let finalFiles = {};
      let isInAssetSheet = false;

      // 1. Exact Match — เร็วสุด
      if (assetMap[cleanedName]) { 
        finalBorrow = assetMap[cleanedName].status; 
        finalSerial = assetMap[cleanedName].serial;
        isInAssetSheet = true;
      } else {
        // 2. Fuzzy Match — ── FIX #2: กรองความยาวก่อน ลด loop ~70% ──
        const lenA = cleanedName.length;
        const allowedErrors = lenA > 5 ? 2 : 1;

        for (let i = 0; i < assetKeys.length; i++) {
          let assetKey = assetKeys[i];

          // ✅ ถ้าความยาวต่างกันเกิน threshold ข้ามเลย ไม่ต้องคำนวณ
          if (Math.abs(assetKey.length - lenA) > allowedErrors) continue;

          let dist = getEditDistance(cleanedName, assetKey);
          if (dist <= allowedErrors) {
            finalBorrow = assetMap[assetKey].status;
            finalSerial = assetMap[assetKey].serial;
            isInAssetSheet = true;
            break;
          }
        }
      }

      if (dbMap[id]) {
        if (dbMap[id].borrowStatus !== 'ยังไม่ยืม') {
          finalBorrow = dbMap[id].borrowStatus;
        } else if (finalBorrow === 'ยังไม่ยืม' && isInAssetSheet) {
          finalBorrow = 'ยืมอยู่'; 
        }
        finalDoc    = dbMap[id].docStatus;
        finalReturnDocStatus = dbMap[id].returnDocStatus;
        if (dbMap[id].serial !== '-') finalSerial = dbMap[id].serial;
        finalFiles  = dbMap[id].files;
      }

      let finalReturnDocUrl = (dbMap[id] && dbMap[id].returnDocUrl) ? dbMap[id].returnDocUrl : '';

      allPeople.push({ 
        type: type, no: no, id: id, name: name, room: room, source_sheet: source, 
        serial: finalSerial, borrowStatus: finalBorrow, docStatus: finalDoc, 
        returnDocStatus: finalReturnDocStatus,
        files: finalFiles, returnDocUrl: finalReturnDocUrl, inAsset: isInAssetSheet 
      });
    };

    // วนลูปอ่านรายชื่อนักเรียน
    SHEET_NAMES.STUDENTS.forEach(sheetName => {
      let sheet = ss.getSheetByName(sheetName);
      if (sheet) { 
        let lastRow = sheet.getLastRow();
        if (lastRow > 1) {
          let data = sheet.getRange(1, 1, lastRow, 4).getValues(); 
          for (let i = 1; i < data.length; i++) { 
            processPerson('student', data[i][0], data[i][1], data[i][2], data[i][3], sheetName); 
          } 
        }
      }
    });

    // วนลูปอ่านรายชื่อครู
    let teacherSheet = ss.getSheetByName(SHEET_NAMES.TEACHERS);
    if (teacherSheet) { 
      let lastRow = teacherSheet.getLastRow();
      if (lastRow > 1) {
        let tData = teacherSheet.getRange(1, 1, lastRow, 2).getValues(); 
        for (let i = 1; i < tData.length; i++) { 
          processPerson('teacher', tData[i][0], 'T-'+tData[i][0], tData[i][1], 'ห้องพักครู', SHEET_NAMES.TEACHERS); 
        } 
      }
    }

    // ── Cache 600 วินาที + chunk batch ──
    _writeCacheChunked(cache, cacheKey, allPeople, 600);

    return allPeople;

  } catch (e) {
    invalidateSystemDataCache();
    throw new Error("ไม่สามารถโหลดข้อมูลได้: " + e.toString());
  } finally {
    try { if (lockAcquired) lock.releaseLock(); } catch (_) {}
  }
}

// ==========================================
// 4. ฟังก์ชันสำหรับระบบจับคู่ชื่อ (AUDIT & FIX)
// ==========================================

function getAssetAuditData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let peopleList = [];
  
  const fetchPeople = (sheetNames, isTeacherSheet) => {
    if(!Array.isArray(sheetNames)) sheetNames = [sheetNames];
    sheetNames.forEach(sheetName => {
      let sheet = ss.getSheetByName(sheetName);
      if(sheet && sheet.getLastRow() > 1) {
        let lastR = Math.max(2, sheet.getLastRow() - 1);
        let cols = isTeacherSheet ? 2 : 4;
        let data = sheet.getRange(2, 1, lastR, cols).getValues();
        data.forEach(r => {
          let originalName = isTeacherSheet ? r[1] : r[2];
          if(originalName) { 
            let norm = normalizeName(originalName);
            if (norm) {
              peopleList.push({
                id: isTeacherSheet ? ('T-'+r[0]) : r[1], 
                name: originalName, 
                room: isTeacherSheet ? 'ห้องพักครู' : (r[3]||''), 
                sheet: sheetName, 
                norm: norm
              });
            }
          }
        });
      }
    });
  };

  fetchPeople(SHEET_NAMES.STUDENTS, false);
  fetchPeople(SHEET_NAMES.TEACHERS, true);

  let peopleNormMap = new Map();
  peopleList.forEach(p => peopleNormMap.set(p.norm, p.name));

  let assetSheet = ss.getSheetByName(SHEET_NAMES.ASSETS);
  let orphans = [];

  if(assetSheet && assetSheet.getLastRow() > 1) {
    let assetData = assetSheet.getRange(2, 1, assetSheet.getLastRow() - 1, 10).getValues(); 
    
    assetData.forEach((r, index) => {
      let assetName = r[4];
      let serial = r[2] || r[3] || 'ไม่มี Serial'; 
      
      if(assetName) {
        let assetNorm = normalizeName(assetName);
        if(!assetNorm) return;
        if(peopleNormMap.has(assetNorm)) return;

        let suggestions = [];
        for(let p of peopleList) {
          let dist = getEditDistance(assetNorm, p.norm);
          let isPartial = assetNorm.includes(p.norm) || p.norm.includes(assetNorm);
          let threshold = Math.ceil(Math.max(assetNorm.length, p.norm.length) * 0.3);
          if(isPartial || dist <= threshold) {
            suggestions.push({
              name: p.name, sheet: p.sheet, room: p.room || '-',
              diff: isPartial ? 0 : dist, isPartial: isPartial
            });
          }
        }
        suggestions.sort((a,b) => (a.isPartial === b.isPartial) ? a.diff - b.diff : (a.isPartial ? -1 : 1));
        orphans.push({
          row: index + 2, assetName: assetName,
          serial: serial.toString(), suggestions: suggestions.slice(0, 5) 
        });
      }
    });
  }
  return orphans;
}

function adminFixAssetName(data) {
  if (!data || data.oldAssetName == null || data.correctName == null) {
    return { success: false, message: "ข้อมูลไม่ครบถ้วน" };
  }
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.ASSETS);
  if (!sheet) return { success: false, message: "ไม่พบแผ่นงานรายงานทะเบียนทรัพย์สิน" };
  
  try {
    const rows = sheet.getDataRange().getDisplayValues();
    let rowToUpdate = -1;
    for(let i=0; i<rows.length; i++) {
      let currentName   = rows[i][4];
      let currentSerial = rows[i][2] || rows[i][3];
      if(currentName == data.oldAssetName && String(currentSerial) == String(data.serial)) {
        rowToUpdate = i + 1;
        break;
      }
    }
    if(rowToUpdate > -1) {
      sheet.getRange(rowToUpdate, 5).setValue(data.correctName);
      invalidateSystemDataCache();
      // ── FIX #5: ไม่เรียก syncAllNamesToSheet ทันที (ให้ Trigger จัดการ) ──
      return { success: true, message: "แก้ไขชื่อในทะเบียนทรัพย์สินเรียบร้อยแล้ว" };
    } else {
      return { success: false, message: "ไม่พบข้อมูลเดิมในทะเบียนทรัพย์สิน (อาจถูกแก้ไขไปแล้ว)" };
    }
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function adminFixAssetNameBulk(updateList) {
  if (!updateList || !Array.isArray(updateList) || updateList.length === 0) {
    return { success: false, message: "ไม่มีรายการที่จะแก้ไข" };
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.ASSETS);
  if (!sheet) return { success: false, message: "ไม่พบแผ่นงานรายงานทะเบียนทรัพย์สิน" };

  const lock = LockService.getScriptLock();
  let lockAcquired = false;

  try {
    lockAcquired = lock.tryLock(10000);
    if (!lockAcquired) {
      return { success: false, message: "ระบบกำลังดำเนินการอื่นอยู่ กรุณาลองใหม่อีกครั้ง" };
    }

    const rows = sheet.getDataRange().getDisplayValues();
    // ── FIX #4: สร้าง lookup map แทนการวนลูปซ้อน แล้ว batch write ──
    const pendingUpdates = {}; // { rowNumber(1-based): newName }

    updateList.forEach(item => {
      if (!item || item.oldAssetName == null || item.correctName == null) return;
      const targetName   = String(item.oldAssetName).trim();
      const targetSerial = String(item.serial || '').trim();

      for (let i = 1; i < rows.length; i++) {
        const currentName   = String(rows[i][4] || '').trim();
        const currentSerial = String(rows[i][2] || rows[i][3] || '').trim();
        if (currentName === targetName && currentSerial === targetSerial) {
          pendingUpdates[i + 1] = item.correctName;
          rows[i][4] = item.correctName; // อัปเดต local copy ป้องกัน duplicate match
          break;
        }
      }
    });

    // เขียนลง Sheet — ทีละ row แต่รวมไว้ใน batch loop เดียว
    const updateRows = Object.keys(pendingUpdates);
    updateRows.forEach(rowNum => {
      sheet.getRange(Number(rowNum), 5).setValue(pendingUpdates[rowNum]);
    });

    SpreadsheetApp.flush();
    invalidateSystemDataCache();
    // ── FIX #5: ไม่เรียก syncAllNamesToSheet ทันที (ให้ Trigger จัดการ) ──
    return { success: true, count: updateRows.length };

  } catch (e) {
    return { success: false, message: "เกิดข้อผิดพลาด: " + e.toString() };
  } finally {
    if (lockAcquired) lock.releaseLock();
  }
}

// ==========================================
// 5. STANDARD HELPERS (Form, Auth, etc.)
// ==========================================

/**
 * แปลง base64 เป็น Blob สำหรับอัปโหลดไฟล์ผ่าน google.script.run
 */
function base64ToBlob(base64Data, fileName) {
  if (!base64Data || typeof base64Data !== 'string') return null;
  try {
    let base64  = base64Data;
    let mimeType = 'application/octet-stream';
    if (base64Data.indexOf('base64,') >= 0) {
      const parts = base64Data.split(';base64,');
      mimeType = parts[0].replace('data:', '') || mimeType;
      base64   = parts[1] || base64;
    }
    const bytes = Utilities.base64Decode(base64);
    return Utilities.newBlob(bytes, mimeType, fileName);
  } catch (e) {
    return null;
  }
}

/**
 * อัปโหลดเอกสารจากหน้าแรก (การยืม หรือ การคืน) - นักเรียน/ครู
 */
function processUploadDoc(obj) {
  if (!obj || !obj.userId || !obj.userName) {
    return { success: false, message: "ข้อมูลไม่ครบถ้วน" };
  }
  const ss          = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheetData   = ss.getSheetByName(SHEET_NAMES.DATA_DB);
  const folder      = DriveApp.getFolderById(FOLDER_ID);
  const folderReturn = DriveApp.getFolderById(FOLDER_RETURN_ID);
  
  try {
    const timestamp = new Date();
    let url = "";
    let action       = "USER_UPDATE";
    let statusToSave = ""; 
    let note         = obj.note || "";
    
    if (obj.uploadType === 'borrow') {
      if (!obj.file_b64) return { success: false, message: "กรุณาเลือกไฟล์ PDF" };
      const fullName = "AGREEMENT_" + obj.userName + "_" + timestamp.getTime();
      const blob = base64ToBlob(obj.file_b64, fullName);
      if (!blob) return { success: false, message: "ไม่สามารถอ่านไฟล์ได้" };
      url          = folder.createFile(blob).setName(fullName).getUrl();
      
      statusToSave = "อัปโหลดเอกสาร | รอตรวจสอบ"; 
      action       = "UPLOAD_DOC";
    } else {
      if (!obj.file_b64) return { success: false, message: "กรุณาเลือกไฟล์ PDF" };
      const fullName = "RETURN_ห้อง" + (obj.userRoom || "") + "_" + obj.userName + "_" + timestamp.getTime();
      const blob = base64ToBlob(obj.file_b64, fullName);
      if (!blob) return { success: false, message: "ไม่สามารถอ่านไฟล์ได้" };
      url          = folderReturn.createFile(blob).setName(fullName).getUrl();
      action       = "USER_RETURN";
      statusToSave = "อยู่ระหว่างการส่งคืน";
      note        += (note ? "\n" : "") + "[หลักฐานการคืน]: " + url;
    }
    
    sheetData.appendRow([
      timestamp, obj.userId, obj.userName, obj.userType, obj.userRoom || "", obj.userSerial || "",
      action, note, statusToSave,
      obj.uploadType === 'borrow' ? url : "", "", "", "", ""
    ]);

    // อัปเดตสถานะลงชีตรายชื่อหลัก
    if (obj.source_sheet) {
      const targetSheet = ss.getSheetByName(obj.source_sheet);
      if (targetSheet) {
        const sheetValues = targetSheet.getDataRange().getValues();
        const isTeacher   = obj.source_sheet === SHEET_NAMES.TEACHERS;
        
        for (let i = 1; i < sheetValues.length; i++) {
          const rowId = isTeacher ? ('T-' + String(sheetValues[i][0])) : String(sheetValues[i][1]);
          
          if (rowId === String(obj.userId)) {
            const currentBorrow = sheetValues[i][10] || 'ยังไม่ยืม';
            const newBorrow = obj.uploadType === 'return' ? 'อยู่ระหว่างการส่งคืน' : currentBorrow;
            const newDoc    = obj.uploadType === 'borrow' ? 'รอตรวจสอบ' : (sheetValues[i][11] || 'รอตรวจสอบ');
            
            targetSheet.getRange(i + 1, 11).setValue(newBorrow);
            if (obj.uploadType === 'borrow') {
              targetSheet.getRange(i + 1, 12).setValue(newDoc);
              // 👇 เพิ่มบรรทัดนี้ เพื่อบันทึกลิงก์ลงคอลัมน์ M (คอลัมน์ที่ 13)
              targetSheet.getRange(i + 1, 13).setValue(url); 
            }
            break;
          }
        }
      }
    }

    invalidateSystemDataCache();
    // 👇 แก้ไขบรรทัดนี้ ให้ส่ง url กลับไปที่หน้าเว็บด้วย
    return { success: true, message: "อัปโหลดเอกสารสำเร็จ", url: url }; 
  } catch (e) {
    return { success: false, message: "เกิดข้อผิดพลาด: " + e.toString() };
  }
}

function processForm(formObject) {
  if (!formObject || !formObject.userId || !formObject.userName) {
    return { success: false, message: "ข้อมูลไม่ครบถ้วน (userId, userName)" };
  }

  const ss        = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheetData = ss.getSheetByName(SHEET_NAMES.DATA_DB);
  const folder    = DriveApp.getFolderById(FOLDER_ID);
  
  try {
    const timestamp = new Date();
    const uploadFromBase64 = (b64Data, prefix) => {
      if (!b64Data) return "";
      const fullName = prefix + "_" + formObject.userName + "_" + timestamp.getTime();
      const blob = base64ToBlob(b64Data, fullName);
      if (!blob) return "";
      return folder.createFile(blob).setName(fullName).getUrl();
    };

    let url_agreement = uploadFromBase64(formObject.file_agreement_b64, "AGREEMENT");
    let url_return    = "";
    const folderReturn = DriveApp.getFolderById(FOLDER_RETURN_ID);
    
    if (formObject.file_return_b64) {
      const fullName = "RETURN_ห้อง" + (formObject.userRoom || "") + "_" + formObject.userName + "_" + timestamp.getTime();
      const blob = base64ToBlob(formObject.file_return_b64, fullName);
      if (blob) url_return = folderReturn.createFile(blob).setName(fullName).getUrl();
    }

    const hasAgreement = url_agreement !== "";
    const hasReturn    = url_return    !== "";
    const serialVal    = (formObject.userSerial && formObject.userSerial !== '-') ? String(formObject.userSerial).trim() : '';

    let statusToSave = formObject.statusSelect || 'ยืมอยู่';
    let action       = "USER_UPDATE";
    let note         = formObject.note || "";

    if (hasAgreement) statusToSave = statusToSave + " | รอตรวจสอบเอกสาร";

    if (hasReturn) {
      statusToSave = "อยู่ระหว่างการส่งคืน";
      action       = "USER_RETURN";
      note        += (note ? "\n" : "") + "[หลักฐานการคืน]: " + url_return;
    } else if (!hasAgreement && serialVal) {
      // Serial-only submit (ไม่มีไฟล์) → ระบุว่ายืมแล้ว
      statusToSave = "ยืมอยู่";
      action       = "USER_SERIAL_UPDATE";
    }

    sheetData.appendRow([
      timestamp, formObject.userId, formObject.userName, formObject.userType,
      formObject.userRoom, serialVal || formObject.userSerial || '-', action, note, statusToSave,
      url_agreement, "", "", "", ""
    ]);
    invalidateSystemDataCache();
    // ── FIX #5: ลบ syncAllNamesToSheet ออก ──
    return { success: true, message: "บันทึกข้อมูลเรียบร้อย" };
  } catch (error) { 
    return { success: false, message: "เกิดข้อผิดพลาด: " + error.toString() };
  }
}

function verifyAdmin(u, p) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.ADMIN);
  if (!sheet) return { success: false, message: "No Admin Sheet" };
  const data = sheet.getDataRange().getDisplayValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(u).trim() && String(data[i][1]).trim() === String(p).trim()) {
      return { success: true, role: 'admin', name: data[i][2] ? String(data[i][2]).trim() : 'Admin' };
    }
  }
  return { success: false, message: "Login Failed" };
}

function verifyAdvisor(u, p) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.ADVISOR);
  if (!sheet) return { success: false, message: "ไม่พบชีตครูที่ปรึกษา" };
  const data = sheet.getDataRange().getDisplayValues();
  for (let i = 1; i < data.length; i++) {
    let rowUser = String(data[i][0] || '').trim();
    let rowPass = String(data[i][1] || '').trim();
    if (rowUser === String(u).trim() && rowPass === String(p).trim()) {
      let level = data[i][2] || '';
      let room  = data[i][3] ? String(data[i][3]).trim() : '';
      let name  = data[i][4] ? String(data[i][4]).trim() : 'คุณครูที่ปรึกษา';
      return { success: true, role: 'advisor', level: level, room: room, name: name };
    }
  }
  return { success: false, message: "ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง" };
}

function adminUpdateData(data) {
  const ss       = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheetLog = ss.getSheetByName(SHEET_NAMES.DATA_DB); 
  
  try {
    const timestamp = new Date();
    let editor = data.editorRole === 'advisor' 
      ? ("ครูที่ปรึกษา: " + data.editorName) 
      : ("ADMIN: " + (data.editorName || "ผู้ดูแลระบบ"));
    
    const targetSheetName = data.source_sheet; 
    let finalSerialToSave = data.userSerial; // เตรียมตัวแปรไว้เก็บ Serial ที่จะเซฟจริง

    if (targetSheetName) {
      const targetSheet = ss.getSheetByName(targetSheetName);
      if (targetSheet) {
        const sheetValues  = targetSheet.getDataRange().getValues();
        const isTeacherSheet = targetSheetName === SHEET_NAMES.TEACHERS;
        
        for (let i = 1; i < sheetValues.length; i++) {
          const rowMatches = isTeacherSheet
            ? ("T-" + String(sheetValues[i][0])) === String(data.userId)
            : String(sheetValues[i][1]) === String(data.userId);
            
          if (rowMatches) {
            // ==========================================
            // ✅ ระบบจัดการเครื่องว่าง (Inventory System)
            // ==========================================
            const oldSerial = sheetValues[i][6]; // คอลัมน์ G (Index 6) คือ Serial เดิม
            
            // เรียกใช้ฟังก์ชันย้ายเครื่องว่าง (ต้องมีฟังก์ชัน processInventoryChanges อยู่ในไฟล์ด้วยนะ)
            finalSerialToSave = processInventoryChanges(oldSerial, data.userSerial, data.borrowStatusSelect, data.userName);
            
            // เซฟค่า Serial ใหม่ลงไปในคอลัมน์ G (7)
            targetSheet.getRange(i + 1, 7).setValue(finalSerialToSave);
            // ==========================================

            if (data.note !== undefined) targetSheet.getRange(i + 1, 8).setValue(data.note);
            if (data.borrowStatusSelect)  targetSheet.getRange(i + 1, 11).setValue(data.borrowStatusSelect);
            if (data.docStatusSelect && data.docStatusSelect !== "") targetSheet.getRange(i + 1, 12).setValue(data.docStatusSelect);
            break;
          }
        }
      } else {
        throw new Error("หาแผ่นงาน '" + targetSheetName + "' ไม่เจอในไฟล์นี้");
      }
    }

    // เซฟ Log การทำงาน (ย้ายมาไว้ด้านล่าง เพื่อให้ได้ finalSerialToSave ที่อัปเดตแล้ว)
    if (sheetLog) {
      if (data.borrowStatusSelect) {
        sheetLog.appendRow([
          timestamp, data.userId, data.userName, data.userType, data.userRoom, finalSerialToSave,
          editor, data.note || "-", data.borrowStatusSelect, "", "", "", "", ""
        ]);
      }
      if (data.docStatusSelect && data.docStatusSelect !== "") {
        sheetLog.appendRow([
          timestamp, data.userId, data.userName, data.userType, data.userRoom, finalSerialToSave,
          editor + "|DOC", data.note || "-", data.docStatusSelect, "", "", "", "", ""
        ]);
      }
      // ✅ แยก log สำหรับสถานะเอกสารการคืน
      if (data.returnDocStatusSelect && data.returnDocStatusSelect !== "") {
        sheetLog.appendRow([
          timestamp, data.userId, data.userName, data.userType, data.userRoom, finalSerialToSave,
          editor + "|RETURN_DOC", data.note || "-", data.returnDocStatusSelect, "", "", "", "", ""
        ]);
      }
    }

    invalidateSystemDataCache();
    // ── FIX #5: ลบ syncAllNamesToSheet ออก ──
    return { success: true, message: "อัปเดตข้อมูลและสถานะใน Sheet สำเร็จ" };
  } catch (e) { 
    return { success: false, message: "เกิดข้อผิดพลาด: " + e.toString() }; 
  }
}

function getAdminSummary() {
  try {
    const data = getAllSystemData();
    let total = 0, borrowed = 0, notBorrow = 0, pending = 0, returned = 0, repair = 0;
    data.forEach(function(p) {
      total++;
      const s = p.borrowStatus || '';
      if (s === 'ยืมอยู่' || s === 'อยู่ระหว่างการส่งคืน') borrowed++;
      else if (s === 'ยังไม่ยืม') notBorrow++;
      else if (s === 'คืนแล้ว') returned++;
      else if (s === 'ส่งซ่อม' || s === 'สละสิทธิ์') repair++;
      if ((p.docStatus || '') === 'รอตรวจสอบ') pending++;
    });
    return { success: true, total, borrowed, notBorrow, pending, returned, repair };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function advisorApproveDoc(data) {
  if (!data || !data.userId || !data.source_sheet) {
    return { success: false, message: "ข้อมูลไม่ครบถ้วน" };
  }
  const ss              = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheetLog        = ss.getSheetByName(SHEET_NAMES.DATA_DB);
  const targetSheetName = data.source_sheet;
  const targetSheet     = ss.getSheetByName(targetSheetName);
  
  if (!targetSheet) return { success: false, message: "ไม่พบแผ่นงาน" };
  
  try {
    const timestamp = new Date();
    const editor    = "ครูที่ปรึกษา: " + (data.editorName || "คุณครู");
    
    if (sheetLog) {
      sheetLog.appendRow([
        timestamp, data.userId, data.userName || "-", data.userType || "student", 
        data.userRoom || "-", data.userSerial || "-",
        editor, "อนุมัติเอกสาร", "เอกสารผ่าน",
        "", "", "", "", ""
      ]);
    }
    
    const sheetValues    = targetSheet.getDataRange().getValues();
    const isTeacherSheet = targetSheetName === SHEET_NAMES.TEACHERS;
    
    for (let i = 1; i < sheetValues.length; i++) {
      const rowMatches = isTeacherSheet
        ? ("T-" + String(sheetValues[i][0])) === String(data.userId)
        : String(sheetValues[i][1]) === String(data.userId);
      if (rowMatches) {
        targetSheet.getRange(i + 1, 12).setValue("เอกสารผ่าน");
        break;
      }
    }
    
    invalidateSystemDataCache();
    // ── FIX #5: ลบ syncAllNamesToSheet ออก ──
    return { success: true, message: "อนุมัติเอกสารเรียบร้อย (สถานะเครื่องยังอยู่ระหว่างการส่งคืน จนกว่าแอดมินจะอนุมัติ)" };
  } catch (e) {
    return { success: false, message: "เกิดข้อผิดพลาด: " + e.toString() };
  }
}

function advisorRejectDoc(data) {
  if (!data || !data.userId || !data.source_sheet) {
    return { success: false, message: "ข้อมูลไม่ครบถ้วน" };
  }
  const ss              = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheetLog        = ss.getSheetByName(SHEET_NAMES.DATA_DB);
  const targetSheetName = data.source_sheet;
  const targetSheet     = ss.getSheetByName(targetSheetName);
  
  if (!targetSheet) return { success: false, message: "ไม่พบแผ่นงาน" };
  
  try {
    const timestamp  = new Date();
    const editor     = "ครูที่ปรึกษา: " + (data.editorName || "คุณครู");
    const rejectNote = "[แจ้งแก้ไข] " + (data.note || "เอกสารไม่ผ่านการตรวจสอบ");
    
    if (sheetLog) {
      sheetLog.appendRow([
        timestamp, data.userId, data.userName || "-", data.userType || "student",
        data.userRoom || "-", data.userSerial || "-",
        editor, rejectNote, "เอกสารไม่ผ่าน",
        "", "", "", "", ""
      ]);
    }
    
    const sheetValues    = targetSheet.getDataRange().getValues();
    const isTeacherSheet = targetSheetName === SHEET_NAMES.TEACHERS;
    
    for (let i = 1; i < sheetValues.length; i++) {
      const rowMatches = isTeacherSheet
        ? ("T-" + String(sheetValues[i][0])) === String(data.userId)
        : String(sheetValues[i][1]) === String(data.userId);
      if (rowMatches) {
        targetSheet.getRange(i + 1, 12).setValue("เอกสารไม่ผ่าน");
        break;
      }
    }
    
    invalidateSystemDataCache();
    return { success: true };
  } catch (e) {
    return { success: false, message: "เกิดข้อผิดพลาด: " + e.toString() };
  }
}

function adminDeleteUser(data) {
  if (data.editorRole === 'advisor') return { success: false, message: "ครูที่ปรึกษาไม่ได้รับอนุญาตให้ลบข้อมูล" };
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(data.source_sheet);
  if (!sheet) return { success: false, message: "ไม่พบแผ่นงาน" };
  try {
    const rows = sheet.getDataRange().getDisplayValues();
    let rowToDelete = -1;
    for (let i = 0; i < rows.length; i++) {
      if (data.source_sheet === SHEET_NAMES.TEACHERS) { 
        if (String(rows[i][1]).trim() === String(data.name).trim()) { rowToDelete = i + 1; break; } 
      } else { 
        if (String(rows[i][1]).trim() === String(data.id).trim()) { rowToDelete = i + 1; break; } 
      }
    }
    if (rowToDelete > -1) { 
      sheet.deleteRow(rowToDelete);
      invalidateSystemDataCache();
      // ── FIX #5: ลบ syncAllNamesToSheet ออก ──
      return { success: true, message: "ลบข้อมูลเรียบร้อยแล้ว" }; 
    } else { return { success: false, message: "ไม่พบข้อมูล" }; }
  } catch (e) { return { success: false, message: "Error: " + e.toString() }; }
}

function adminAddUser(data) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(data.targetSheet); 
  if (!sheet) return { success: false, message: "ไม่พบแผ่นงาน" };
  try {
    const nextNo = sheet.getLastRow(); 
    if (data.targetSheet === SHEET_NAMES.TEACHERS) {
      sheet.appendRow([nextNo, data.name]); 
    } else {
      sheet.appendRow([nextNo, data.id, data.name, data.room]);
    }
    invalidateSystemDataCache();
    // ── FIX #5: ลบ syncAllNamesToSheet ออก ──
    return { success: true, message: "เพิ่มรายชื่อเรียบร้อยแล้ว" };
  } catch (e) { return { success: false, message: "Error: " + e.toString() }; }
}

// ==========================================
// 6. DASHBOARD STATS SYSTEM
// ==========================================

/**
 * ── FIX #3: ใช้ getAllSystemData() แทนการเปิด Sheet ซ้ำ ──
 * ประหยัด Quota และเร็วกว่ามากเพราะใช้ Cache
 */
function getDashboardStats() {
  try {
    const data = getAllSystemData();
    let borrowed = 0;
    for (let i = 0; i < data.length; i++) {
      const s = data[i].borrowStatus || '';
      if (s === 'ยืมอยู่' || s === 'อยู่ระหว่างการส่งคืน') borrowed++;
    }
    return { total: 2085, borrowed: borrowed, available: 2085 - borrowed };
  } catch(e) {
    return { total: 2085, borrowed: 0, available: 2085 };
  }
}

// ==========================================
// ฟังก์ชันสำหรับครูที่ปรึกษาทำเรื่องคืน iPad
// ==========================================
function processAdvisorReturn(formObject) {
  if (!formObject || !formObject.userId || !formObject.userName) {
    return { success: false, message: "ข้อมูลไม่ครบถ้วน (userId, userName)" };
  }

  const ss           = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheetData    = ss.getSheetByName(SHEET_NAMES.DATA_DB);
  const folderReturn = DriveApp.getFolderById(FOLDER_RETURN_ID);
  
  try {
    const timestamp = new Date();
    let url_return  = "";
    
    if (formObject.file_return_b64) {
      const fileName = "RETURN_ห้อง" + (formObject.userRoom || "") + "_" + formObject.userName + "_" + timestamp.getTime();
      const blob = base64ToBlob(formObject.file_return_b64, fileName);
      if (blob) url_return = folderReturn.createFile(blob).setName(fileName).getUrl();
    }
    
    let statusToSave = "อยู่ระหว่างการส่งคืน";
    let newNote = formObject.note || "";
    if (url_return !== "") {
      newNote += "\n[หลักฐานการคืน]: " + url_return;
    }

    sheetData.appendRow([
      timestamp, formObject.userId, formObject.userName, formObject.userType, 
      formObject.userRoom, formObject.userSerial, "ADVISOR_RETURN", newNote, statusToSave, 
      "", "", "", "", ""
    ]);

    if (formObject.source_sheet) {
      const targetSheet = ss.getSheetByName(formObject.source_sheet);
      if (targetSheet) {
        const sheetValues  = targetSheet.getDataRange().getValues();
        const isTeacher    = formObject.source_sheet === SHEET_NAMES.TEACHERS;
        for (let i = 1; i < sheetValues.length; i++) {
          const rowId = isTeacher ? ('T-' + String(sheetValues[i][0])) : String(sheetValues[i][1]);
          if (rowId === String(formObject.userId)) {
            targetSheet.getRange(i + 1, 11).setValue('อยู่ระหว่างการส่งคืน');
            break;
          }
        }
      }
    }

    invalidateSystemDataCache();
    // ── FIX #5: ลบ syncAllNamesToSheet ออก ──
    return { success: true, message: "ส่งเรื่องคืนเรียบร้อย รอแอดมินตรวจสอบ" };
  } catch (error) { 
    return { success: false, message: "เกิดข้อผิดพลาด: " + error.toString() };
  }
}

// ==========================================
// 7. EXPORT & REPORT
// ==========================================

function getExportData(options) {
  try {
    let data = getAllSystemData();
    if (!options || options.scope !== 'filtered') return data;
    let filtered = data;
    if (options.levelFilter) {
      filtered = filtered.filter(function(r) { return String(r.source_sheet || '').indexOf(options.levelFilter) >= 0; });
    }
    if (options.roomFilter) {
      filtered = filtered.filter(function(r) { return String(r.room || '') === String(options.roomFilter); });
    }
    return filtered;
  } catch (e) {
    throw new Error("ไม่สามารถดึงข้อมูล Export ได้: " + e.toString());
  }
}

function getReportData() {
  try {
    var data  = getAllSystemData();
    var byKey = {};

    for (var i = 0; i < data.length; i++) {
      var p     = data[i];
      var level = p.source_sheet || '-';
      var room  = String(p.room || '-');
      var key   = level + '|' + room;

      if (!byKey[key]) {
        byKey[key] = { level: level, room: room, total: 0, borrowed: 0, docPassed: 0, docSubmitted: 0, notBorrowed: 0, returned: 0, repair: 0 };
      }

      byKey[key].total++;
      var status  = p.borrowStatus || '';
      var docStat = p.docStatus    || '';

      if (status === 'ยืมอยู่' || status === 'อยู่ระหว่างการส่งคืน') byKey[key].borrowed++;
      else if (status === 'ยังไม่ยืม') byKey[key].notBorrowed++;
      else if (status === 'คืนแล้ว') byKey[key].returned++;
      else if (status === 'ส่งซ่อม' || status === 'สละสิทธิ์') byKey[key].repair++;

      if (docStat === 'เอกสารผ่าน') { byKey[key].docPassed++; byKey[key].docSubmitted++; }
      else if (docStat === 'รอตรวจสอบ' || docStat === 'เอกสารไม่ผ่าน') byKey[key].docSubmitted++;
    }

    return Object.keys(byKey).map(function(k) { return byKey[k]; }).sort(function(a, b) {
      if (a.level !== b.level) return String(a.level).localeCompare(b.level);
      return String(a.room).localeCompare(b.room);
    });
  } catch (e) {
    throw new Error("ไม่สามารถดึงข้อมูลรายงานได้: " + e.toString());
  }
}

// ==========================================
// 8. ซิงค์ชีต "รายชื่อทั้งหมด" (SYNC ALL NAMES SHEET)
// ==========================================

function syncAllNamesToSheet(data) {
  if (!data || !Array.isArray(data)) return;
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEET_NAMES.ALL_NAMES);
    if (!sheet) sheet = ss.insertSheet(SHEET_NAMES.ALL_NAMES);
    
    const headers = ['No', 'รหัส', 'ชื่อ-นามสกุล', 'ประเภท', 'ระดับชั้น', 'ห้อง', 'Serial', 'สถานะเครื่อง', 'สถานะเอกสาร'];
    
    // จำกัดจำนวนข้อมูลที่จะ sync เพื่อป้องกัน timeout
    const MAX_ROWS = 1000;
    const limitedData = data.slice(0, MAX_ROWS);
    
    const rows = limitedData.map(function(p, idx) {
      return [
        idx + 1, p.id || '', p.name || '',
        (p.type === 'teacher' ? 'ครู' : 'นักเรียน'),
        p.source_sheet || '', p.room || '',
        p.serial || '-', p.borrowStatus || 'ยังไม่ยืม', p.docStatus || 'ยังไม่ส่ง'
      ];
    });
    
    // ใช้ batch operations เพื่อความเร็ว
    sheet.clearContents();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    }
    SpreadsheetApp.flush();
    
    // คืนค่าสำเร็จ
    return { success: true, synced: rows.length, total: data.length };
  } catch (e) {
    Logger.log('syncAllNamesToSheet error: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

/**
 * อัปเดตข้อมูล + ซิงค์ชีตรายชื่อทั้งหมด (สำหรับปุ่ม "อัปเดต" และ Time Trigger)
 */
function refreshAndSyncData() {
  invalidateSystemDataCache();
  const data = getAllSystemData();
  syncAllNamesToSheet(data);
  return data;
}

// ==========================================
// 9. TIME-BASED TRIGGER SETUP
// ── FIX #5: สร้าง Trigger ให้ syncAllNamesToSheet ทำงานอัตโนมัติ
//    ทุก 5 นาที แทนการเรียกหลัง action ทุกครั้ง ──
// วิธีใช้: รัน createSyncTrigger() ครั้งเดียวจาก Apps Script Editor
// ==========================================

/**
 * สร้าง Time-Based Trigger สำหรับ sync ชีต "รายชื่อทั้งหมด" ทุก 5 นาที
 * รัน refreshAndSyncData ผ่าน Trigger แทนการเรียกหลัง action ทุกครั้ง
 */
function createSyncTrigger() {
  // ลบ Trigger เก่าออกก่อน (ป้องกัน duplicate)
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'refreshAndSyncData') {
      ScriptApp.deleteTrigger(t);
    }
  });
  // สร้างใหม่ทุก 5 นาที
  ScriptApp.newTrigger('refreshAndSyncData')
    .timeBased()
    .everyMinutes(5)
    .create();
  Logger.log('✅ Sync Trigger created: refreshAndSyncData ทุก 5 นาที');
}

/**
 * ลบ Trigger ทั้งหมด (ใช้เมื่อต้องการหยุด auto-sync)
 */
function deleteSyncTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'refreshAndSyncData') {
      ScriptApp.deleteTrigger(t);
    }
  });
  Logger.log('🗑️ Sync Trigger ถูกลบแล้ว');
}

// ==========================================
// ✅ ฟังก์ชัน Activity Log (ประวัติการยืม-คืน)
// ==========================================

/**
 * ดึงประวัติกิจกรรมล่าสุดจาก DATA_DB (cache 60 วินาที)
 * @param {number} limit จำนวนแถวที่ต้องการ (default 60)
 */
function getActivityLog(limit) {
  limit = limit || 60;
  const cache    = CacheService.getScriptCache();
  const cacheKey = 'activityLog_' + SPREADSHEET_ID + '_' + limit;
  try {
    const cached = cache.get(cacheKey);
    if (cached) return JSON.parse(cached);
  } catch (_) {}

  try {
    const ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet   = ss.getSheetByName(SHEET_NAMES.DATA_DB);
    if (!sheet) return [];

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    // ดึงแค่แถวสุดท้าย limit แถว (reverse order = ใหม่สุดขึ้นก่อน)
    const startRow = Math.max(2, lastRow - limit + 1);
    const numRows  = lastRow - startRow + 1;
    const data     = sheet.getRange(startRow, 1, numRows, 9).getValues();

    const logs = [];
    for (let i = data.length - 1; i >= 0; i--) {
      const row = data[i];
      const ts      = row[0];
      const id      = row[1] ? String(row[1]) : '';
      const name    = row[2] ? String(row[2]) : '';
      const type    = row[3] ? String(row[3]) : '';
      const room    = row[4] ? String(row[4]) : '';
      const serial  = row[5] ? String(row[5]) : '';
      const action  = row[6] ? String(row[6]) : '';
      const note    = row[7] ? String(row[7]) : '';
      const status  = row[8] ? String(row[8]) : '';
      if (!name && !id) continue;
      logs.push({
        ts:     ts instanceof Date ? Utilities.formatDate(ts, 'Asia/Bangkok', 'dd/MM/yy HH:mm') : String(ts),
        id:     id,
        name:   name,
        type:   type,
        room:   room,
        serial: serial,
        action: action,
        note:   note,
        status: status
      });
    }

    try { cache.put(cacheKey, JSON.stringify(logs), 60); } catch (_) {}
    return logs;

  } catch (e) {
    Logger.log('getActivityLog error: ' + e.toString());
    return [];
  }
}

/**
 * จำนวน iPad ที่ว่าง (จากชีต "เครื่องว่าง") พร้อม cache 5 นาที
 */
function getInventoryAvailableCount() {
  const cache    = CacheService.getScriptCache();
  const cacheKey = 'invCount_' + SPREADSHEET_ID;
  try {
    const cached = cache.get(cacheKey);
    if (cached) return parseInt(cached, 10);
  } catch (_) {}
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('เครื่องว่าง');
    if (!sheet) return 0;
    const lastRow = sheet.getLastRow();
    const count   = Math.max(0, lastRow - 1);
    try { cache.put(cacheKey, String(count), 300); } catch (_) {}
    return count;
  } catch (_) {
    return 0;
  }
}

/**
 * Batch update: อัปเดตหลายคนพร้อมกัน (Admin เท่านั้น)
 * items: [{ userId, source_sheet, borrowStatusSelect, docStatusSelect, note }]
 */
function adminBatchUpdate(items, editorName) {
  if (!items || !Array.isArray(items) || !items.length) {
    return { success: false, message: 'ไม่มีรายการที่จะอัปเดต' };
  }
  const lock = LockService.getScriptLock();
  let lockAcquired = false;
  try {
    lockAcquired = lock.tryLock(12000);
    if (!lockAcquired) return { success: false, message: 'ระบบยุ่งอยู่ กรุณาลองใหม่' };

    const ss         = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetLog   = ss.getSheetByName(SHEET_NAMES.DATA_DB);
    const timestamp  = new Date();
    const editor     = 'ADMIN: ' + (editorName || 'ผู้ดูแลระบบ');
    let   updated    = 0;

    const sheetCache = {}; // cache sheet values per sheetName
    const logRows    = [];

    items.forEach(function(item) {
      if (!item.userId || !item.source_sheet) return;
      const sheetName = item.source_sheet;
      if (!sheetCache[sheetName]) {
        const s = ss.getSheetByName(sheetName);
        if (s) sheetCache[sheetName] = { sheet: s, values: s.getDataRange().getValues() };
      }
      const sc = sheetCache[sheetName];
      if (!sc) return;
      const isTeacher = sheetName === SHEET_NAMES.TEACHERS;
      for (let i = 1; i < sc.values.length; i++) {
        const rowId = isTeacher ? ('T-' + String(sc.values[i][0])) : String(sc.values[i][1]);
        if (rowId === String(item.userId)) {
          if (item.borrowStatusSelect) {
            sc.sheet.getRange(i + 1, 11).setValue(item.borrowStatusSelect);
            sc.values[i][10] = item.borrowStatusSelect;
          }
          if (item.docStatusSelect) {
            sc.sheet.getRange(i + 1, 12).setValue(item.docStatusSelect);
            sc.values[i][11] = item.docStatusSelect;
          }
          logRows.push([
            timestamp, item.userId, item.userName || '-', item.userType || 'student',
            item.userRoom || '-', item.userSerial || '-',
            editor, item.note || 'Batch update', item.borrowStatusSelect || '-',
            '', '', '', '', ''
          ]);
          updated++;
          break;
        }
      }
    });

    if (sheetLog && logRows.length) {
      sheetLog.getRange(sheetLog.getLastRow() + 1, 1, logRows.length, 14).setValues(logRows);
    }
    SpreadsheetApp.flush();
    invalidateSystemDataCache();
    return { success: true, updated: updated, message: 'อัปเดต ' + updated + ' รายการเรียบร้อย' };
  } catch (e) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + e.toString() };
  } finally {
    if (lockAcquired) try { lock.releaseLock(); } catch(_) {}
  }
}

// ==========================================
// ระบบจัดการคลังพัสดุ (Inventory System)
// ==========================================

// 1. ฟังก์ชันส่งรายชื่อเครื่องว่างให้หน้าเว็บ (อัปเดตใหม่ ป้องกัน Error ชีตว่าง)
function getAvailableSerials() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('เครื่องว่าง');
    
    // ถ้าหาชีตไม่เจอ ให้คืนค่าว่าง
    if (!sheet) return []; 
    
    var lastRow = sheet.getLastRow();
    // ถ้าชีตมีแค่หัวตาราง (1 บรรทัด) หรือไม่มีข้อมูลเลย ให้คืนค่าว่าง
    if (lastRow < 2) return []; 
    
    // ดึงข้อมูลตั้งแต่บรรทัดที่ 2 ถึงบรรทัดสุดท้าย คอลัมน์ที่ 1 (A)
    var data = sheet.getRange(2, 1, lastRow - 1, 1).getValues(); 
    var serials = [];
    
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] && String(data[i][0]).trim() !== "") {
        serials.push(String(data[i][0]).trim());
      }
    }
    return serials;
  } catch (e) {
    // ปริ้น Error ลง Logs หลังบ้านเผื่อไว้เช็ค
    console.error("Error in getAvailableSerials: " + e.message);
    return [];
  }
}

// 2. ฟังก์ชันย้าย Serial ไป-มา อัตโนมัติ
function processInventoryChanges(oldSerial, newSerial, newStatus, userName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var emptySheet = ss.getSheetByName('เครื่องว่าง');
  if (!emptySheet) return newSerial;

  var dateStr = Utilities.formatDate(new Date(), "Asia/Bangkok", "dd/MM/yyyy HH:mm");

  // กรณีที่ 1: แอดมินกด "คืนแล้ว" -> เอา Serial ไปเก็บในชีต "เครื่องว่าง"
  if (newStatus === 'คืนแล้ว' && oldSerial && oldSerial !== '-') {
    var emptyData = emptySheet.getRange("A:A").getValues().flat();
    if (emptyData.indexOf(oldSerial) === -1) { 
      // โครงสร้าง: [Serial Number, ผู้ใช้งานล่าสุด, วันที่รับคืน, สภาพเครื่อง]
      emptySheet.appendRow([oldSerial, userName, dateStr, "ปกติ"]);
    }
    return '-'; // บังคับลบ Serial ออกจากชื่อเด็กเมื่อคืนแล้ว
  }

  // กรณีที่ 2: แอดมินจ่ายเครื่องใหม่ (เปลี่ยน Serial ไปจากเดิม) -> ลบเลขนั้นออกจาก "เครื่องว่าง"
  if (newSerial !== oldSerial && newSerial !== '-') {
    var data = emptySheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(newSerial).trim()) {
        emptySheet.deleteRow(i + 1);
        break;
      }
    }
  }
  
  return newSerial; // คืนค่า Serial ล่าสุดกลับไปบันทึก
}
// ==========================================
// ฟังก์ชันจัดการรายชื่อนักเรียนโดยครูที่ปรึกษา
// เพิ่มต่อท้ายไฟล์ Code.gs (ก่อน closing บรรทัดสุดท้าย)
// ==========================================

/**
 * ครูที่ปรึกษา: เพิ่มนักเรียนใหม่ในห้องตัวเอง
 * data: { targetSheet, id, name, room, advisorLevel, advisorRoom, editorName }
 */

// ฟังก์ชันสำหรับครูที่ปรึกษากดเพิ่มรายชื่อนักเรียน
// ฟังก์ชันสำหรับครูที่ปรึกษากดเพิ่มรายชื่อนักเรียน
function advisorAddStudent(data) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // 1. ตรวจสอบชื่อชีต
    const sheetNameRaw = data.targetSheet || data.source_sheet || '';
    const targetSheetName = sheetNameRaw.trim();
    
    if (!targetSheetName) {
       return { success: false, message: 'ไม่พบข้อมูลชื่อหน้าชีตที่จะเพิ่ม' };
    }
    
    if (SHEET_NAMES.STUDENTS.indexOf(targetSheetName) < 0) {
      return { success: false, message: 'ไม่มีสิทธิ์แก้ไขในชีตนี้ หรือชื่อชีตไม่ถูกต้อง' };
    }

    const sheet = ss.getSheetByName(targetSheetName);
    if (!sheet) {
      return { success: false, message: 'ไม่พบหน้าชีต: ' + targetSheetName };
    }

    // ==========================================
    // 2. วิธีแก้ปัญหาข้อมูลทับกัน (เช็คบรรทัดสุดท้ายจากทุกคอลัมน์)
    // ==========================================
    // ดึงข้อมูลทั้งหมดในชีตมาตรวจสอบ
    const vals = sheet.getDataRange().getValues();
    let realLastRow = 1;
    
    // เช็คย้อนกลับจากบรรทัดล่างสุดขึ้นมาบนสุด
    for (let i = vals.length - 1; i >= 0; i--) {
      // เอาค่าทุกช่องในบรรทัดนั้นมารวมกัน ถ้าไม่ได้ว่างเปล่า แสดงว่าเป็นบรรทัดที่มีข้อมูลล่าสุด
      if (vals[i].join('').trim() !== '') {
        realLastRow = i + 2; // +1 เพื่อเป็นเลขแถวปกติ และ +1 เพื่อแทรกลงบรรทัดว่างถัดไป
        break;
      }
    }
    
    // นำข้อมูลไปวางต่อท้ายบรรทัดที่ว่างจริงๆ (ลำดับ, รหัสประจำตัว, ชื่อ-สกุล, ห้อง)
    sheet.getRange(realLastRow, 1, 1, 4).setValues([[data.no, data.id, data.name, data.room]]);

    // 3. เคลียร์ Cache เพื่อให้หน้าเว็บอัปเดตใหม่ทันที
    invalidateSystemDataCache();

    return { success: true, message: 'เพิ่มรายชื่อเรียบร้อยแล้ว' };

  } catch (e) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + e.toString() };
  }
}

/**
 * ครูที่ปรึกษา: แก้ไขข้อมูลนักเรียน (ชื่อ, รหัส, ห้อง) ในห้องตัวเอง
 * data: { id, newName, newId, newRoom, source_sheet, editorName }
 */
function advisorEditStudent(data) {
  if (!data || !data.id || !data.source_sheet) {
    return { success: false, message: 'ข้อมูลไม่ครบถ้วน' };
  }
  if (SHEET_NAMES.STUDENTS.indexOf(data.source_sheet) < 0) {
    return { success: false, message: 'ไม่มีสิทธิ์แก้ไขในชีตนี้' };
  }
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(data.source_sheet);
  if (!sheet) return { success: false, message: 'ไม่พบแผ่นงาน: ' + data.source_sheet };

  try {
    const values = sheet.getDataRange().getValues();
    let rowFound = -1;
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][1]).trim() === String(data.id).trim()) {
        rowFound = i + 1; // 1-based row number
        break;
      }
    }
    if (rowFound < 0) return { success: false, message: 'ไม่พบนักเรียนรหัส ' + data.id };

    // ตรวจสอบรหัสใหม่ไม่ซ้ำ (ถ้ามีการเปลี่ยน)
    if (data.newId && String(data.newId).trim() !== String(data.id).trim()) {
      for (let i = 1; i < values.length; i++) {
        if (String(values[i][1]).trim() === String(data.newId).trim() && (i + 1) !== rowFound) {
          return { success: false, message: 'รหัสนักเรียน ' + data.newId + ' มีอยู่ในระบบแล้ว' };
        }
      }
      sheet.getRange(rowFound, 2).setValue(String(data.newId).trim());
    }
    if (data.newName) sheet.getRange(rowFound, 3).setValue(data.newName.trim());
    if (data.newRoom !== undefined && data.newRoom !== null) sheet.getRange(rowFound, 4).setValue(data.newRoom);

    // บันทึก Log
    const logSheet = ss.getSheetByName(SHEET_NAMES.DATA_DB);
    if (logSheet) {
      logSheet.appendRow([
        new Date(), data.newId || data.id, data.newName || values[rowFound-1][2], 'student', data.newRoom || values[rowFound-1][3], '-',
        'ADVISOR_EDIT_STUDENT', 'แก้ไขข้อมูลนักเรียนโดย: ' + (data.editorName || 'ครูที่ปรึกษา') +
          ' | เดิม: รหัส=' + data.id + ' ชื่อ=' + values[rowFound-1][2],
        '-', '', '', '', '', ''
      ]);
    }

    invalidateSystemDataCache();
    
    // อัปเดตชีตรายชื่อทั้งหมด
    try {
      const updatedData = getAllSystemData();
      const syncResult = syncAllNamesToSheet(updatedData);
      if (!syncResult.success) {
        Logger.log('syncAllNamesToSheet failed in advisorEditStudent: ' + syncResult.error);
        return { success: true, message: 'แก้ไขข้อมูลนักเรียนเรียบร้อยแล้ว (แต่ sync ชีตรายชื่อไม่สำเร็จ)' };
      }
    } catch (e) {
      Logger.log('syncAllNamesToSheet error in advisorEditStudent: ' + e.toString());
      return { success: true, message: 'แก้ไขข้อมูลนักเรียนเรียบร้อยแล้ว (แต่ sync ชีตรายชื่อไม่สำเร็จ)' };
    }
    
    return { success: true, message: 'แก้ไขข้อมูลนักเรียนเรียบร้อยแล้ว' };
  } catch (e) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + e.toString() };
  }
}

/**
 * ครูที่ปรึกษา: ลบนักเรียนออกจากห้องตัวเอง
 * data: { id, name, source_sheet, editorName }
 */
function advisorDeleteStudent(data) {
  if (!data || !data.id || !data.source_sheet) {
    return { success: false, message: 'ข้อมูลไม่ครบถ้วน' };
  }
  if (SHEET_NAMES.STUDENTS.indexOf(data.source_sheet) < 0) {
    return { success: false, message: 'ไม่มีสิทธิ์ลบในชีตนี้ (สามารถลบได้เฉพาะนักเรียนเท่านั้น)' };
  }
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(data.source_sheet);
  if (!sheet) return { success: false, message: 'ไม่พบแผ่นงาน: ' + data.source_sheet };

  try {
    const values = sheet.getDataRange().getDisplayValues();
    let rowToDelete = -1;
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][1]).trim() === String(data.id).trim()) {
        rowToDelete = i + 1;
        break;
      }
    }
    if (rowToDelete < 0) return { success: false, message: 'ไม่พบนักเรียนรหัส ' + data.id };

    // บันทึก Log ก่อนลบ
    const logSheet = ss.getSheetByName(SHEET_NAMES.DATA_DB);
    if (logSheet) {
      logSheet.appendRow([
        new Date(), data.id, data.name || '-', 'student', '-', '-',
        'ADVISOR_DELETE_STUDENT', 'ลบรายชื่อนักเรียนโดย: ' + (data.editorName || 'ครูที่ปรึกษา'),
        'ลบออกจากระบบ', '', '', '', '', ''
      ]);
    }

    sheet.deleteRow(rowToDelete);
    invalidateSystemDataCache();
    
    // อัปเดตชีตรายชื่อทั้งหมด
    try {
      const updatedData = getAllSystemData();
      const syncResult = syncAllNamesToSheet(updatedData);
      if (!syncResult.success) {
        Logger.log('syncAllNamesToSheet failed in advisorDeleteStudent: ' + syncResult.error);
        return { success: true, message: 'ลบรายชื่อ "' + (data.name || data.id) + '" เรียบร้อยแล้ว (แต่ sync ชีตรายชื่อไม่สำเร็จ)' };
      }
    } catch (e) {
      Logger.log('syncAllNamesToSheet error in advisorDeleteStudent: ' + e.toString());
      return { success: true, message: 'ลบรายชื่อ "' + (data.name || data.id) + '" เรียบร้อยแล้ว (แต่ sync ชีตรายชื่อไม่สำเร็จ)' };
    }
    
    return { success: true, message: 'ลบรายชื่อ "' + (data.name || data.id) + '" เรียบร้อยแล้ว' };
  } catch (e) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + e.toString() };
  }
}

