// ==========================================
// 1. ส่วนตั้งค่าระบบ (CONFIGURATION)
// ==========================================

const SPREADSHEET_ID = '1Kr1GOn5F8rBNJGA7_Sqp4be7JirrRvVci0AGhtkA5hQ'; // ID ของ Google Sheet
const FOLDER_ID = '1pPtZlI8XYBle02byB5lthAhtLX8012Pa'; // ID ของ Google Drive Folder

const SHEET_NAMES = {
  STUDENTS: [
    'รายชื่อนักเรียน ม.3', 
    'รายชื่อนักเรียน ม.4', 
    'รายชื่อนักเรียน ม.5', 
    'รายชื่อนักเรียน ม.6'
  ],
  TEACHERS: 'รายชื่อครู',
  ASSETS: 'รายงานทะเบียนทรัพย์สิน', // ชื่อชีตต้องตรงเป๊ะ
  DATA_DB: 'ข้อมูล',
  LOGS: 'Log',
  ADMIN: 'แอดมิน',
  ADVISOR: 'ครูที่ปรึกษา'
};

// ==========================================
// 2. ฟังก์ชันพื้นฐาน & AI Helper
// ==========================================

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('ระบบตรวจสอบสถานะ iPad โรงเรียนอรัญประเทศ')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) { 
  return HtmlService.createHtmlOutputFromFile(filename).getContent(); 
}

// *** ฟังก์ชันล้างชื่อขั้นสูง (Normalize) ***
// ตัดคำนำหน้า ยศทหาร และอักขระพิเศษออก เพื่อให้เทียบชื่อได้แม่นยำ
function normalizeName(name) {
  if (!name) return "";
  
  // แปลงเป็น String, ลบอักขระล่องหน (Zero-width space), และ Trim
  let n = name.toString().replace(/[\u200b\u00a0\u180e\u2000-\u200a\u202f\u205f\u3000]/g, '').trim();

  // Regex รวมยศทหาร, ตำแหน่งวิชาการ, คำนำหน้า, และคำที่มักพิมพ์ผิด
  const titleRegex = /^(?:ว่าที่\s*ร(?:้อย)?\.?\s*ต(?:รี)?\.?|จ(?:่า)?\.?ส(?:ิบ)?\.?[อทต]\.?|ส(?:ิบ)?\.?[อทต]\.?|พล(?:ทหาร)?\.?|ส\.อ\.|จ\.ส\.อ\.|เด็กชาย|เด็กหญิง|นางสาว|ด\.?\s*ช\.?|ด\.?\s*ญ\.?|น\.?\s*ส\.?|นาย|นาง|ครู|อ\.|Mr\.?|Mrs\.?|Ms\.?|Miss|Dr\.?)[\s\.]*/gi;
  
  n = n.replace(titleRegex, ''); // ลบคำนำหน้าทิ้ง
  n = n.replace(/\s/g, '');      // ลบช่องว่างทิ้งทั้งหมด
  
  return n;
}

// *** ฟังก์ชันคำนวณความต่างของคำ (Levenshtein Distance) ***
// ใช้สำหรับหาชื่อที่สะกดผิดเล็กน้อย
function getEditDistance(a, b) {
  if (a.length === 0) return b.length; 
  if (b.length === 0) return a.length; 

  var matrix = [];

  // สร้างตารางเมทริกซ์
  for (var i = 0; i <= b.length; i++) { matrix[i] = [i]; }
  for (var j = 0; j <= a.length; j++) { matrix[0][j] = j; }

  // คำนวณระยะห่าง
  for (var i = 1; i <= b.length; i++) {
    for (var j = 1; j <= a.length; j++) {
      if (b.charAt(i - 1) == a.charAt(j - 1)) {
        matrix[i][j] = matrix[i - 1][j - 1];
      } else {
        matrix[i][j] = Math.min(
          matrix[i - 1][j - 1] + 1, // substitution
          Math.min(
            matrix[i][j - 1] + 1, // insertion
            matrix[i - 1][j] + 1  // deletion
          )
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
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // --- A. ดึงข้อมูล Asset (ทะเบียนทรัพย์สิน) ---
  const assetSheet = ss.getSheetByName(SHEET_NAMES.ASSETS);
  let assetMap = {};
  let assetKeys = []; // เก็บรายชื่อที่ล้างแล้วไว้ทำ Fuzzy Match

  if (assetSheet) {
    const lastRow = assetSheet.getLastRow();
    if (lastRow > 1) {
      // อ่านข้อมูล Fast Mode (อ่านถึงคอลัมน์ J = 10 คอลัมน์)
      const assetData = assetSheet.getRange(1, 1, lastRow, 10).getValues();
      
      for (let i = 1; i < assetData.length; i++) {
        // ชื่ออยู่ที่ Col E (Index 4)
        let rawName = assetData[i][4]; 
        // Serial อยู่ที่ Col C (Index 2) หรือ D (Index 3)
        let serial = assetData[i][2] || assetData[i][3]; 

        if (rawName) {
          let cName = normalizeName(rawName); // ล้างชื่อ
          if (cName.length > 0) {
            assetMap[cName] = { 
              serial: serial ? serial.toString() : '-', 
              status: 'ยืมอยู่' 
            };
            assetKeys.push(cName); // เก็บเข้าลิสต์
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
      // อ่านข้อมูลถึงคอลัมน์ N (14 คอลัมน์)
      const dbData = dbSheet.getRange(1, 1, dbLastRow, 14).getValues();
      
      for (let i = 1; i < dbData.length; i++) {
        let id = dbData[i][1]; // รหัสประจำตัว
        if (!id) continue;
        id = id.toString();

        let rowSerial = dbData[i][5] ? dbData[i][5].toString() : '';
        let rowStatus = dbData[i][8] ? dbData[i][8].toString() : '';
        let hasFiles = (dbData[i][9] || dbData[i][10] || dbData[i][11] || dbData[i][12]);
        
        if (!dbMap[id]) dbMap[id] = { borrowStatus: 'ยังไม่ยืม', docStatus: 'ยังไม่ส่ง', serial: '-', files: {} };
        
        if (rowSerial && rowSerial !== '-' && rowSerial !== '') dbMap[id].serial = rowSerial;
        
        if (hasFiles) {
          dbMap[id].files = { 
            agreement: dbData[i][9], card_std: dbData[i][10], 
            card_parent: dbData[i][11], house: dbData[i][12], phone: dbData[i][13] 
          };
          if (!rowStatus.includes('ADMIN') && !rowStatus.includes('ADVISOR')) {
            dbMap[id].docStatus = 'รอตรวจสอบ';
          }
        }

        if (rowStatus.includes('เอกสารผ่าน')) dbMap[id].docStatus = 'เอกสารผ่าน';
        else if (rowStatus.includes('ไม่ผ่าน')) dbMap[id].docStatus = 'เอกสารไม่ผ่าน';
        else if (rowStatus.includes('รอตรวจสอบ')) dbMap[id].docStatus = 'รอตรวจสอบ';

        if (rowStatus.includes('ยืมอยู่') || rowStatus === 'ยืมได้') dbMap[id].borrowStatus = 'ยืมอยู่';
        else if (rowStatus.includes('คืน')) dbMap[id].borrowStatus = 'คืนแล้ว';
        else if (rowStatus.includes('ซ่อม')) dbMap[id].borrowStatus = 'ส่งซ่อม';
        else if (rowStatus.includes('สละ')) dbMap[id].borrowStatus = 'สละสิทธิ์';
        else if (rowStatus === 'ยังไม่ยืม') dbMap[id].borrowStatus = 'ยังไม่ยืม';
      }
    }
  }

  // --- C. รวมข้อมูล (Merge + Fuzzy Match) ---
  let allPeople = [];
  
  const processPerson = (type, no, id, name, room, source) => {
    if (!name) return;
    id = id.toString();
    
    let cleanedName = normalizeName(name); // ล้างชื่อต้นทาง
    
    let finalBorrow = 'ยังไม่ยืม';
    let finalDoc = 'ยังไม่ส่ง';
    let finalSerial = '-';
    let finalFiles = {};
    let isInAssetSheet = false;

    // 1. ลองหาแบบ "ตรงเป๊ะ" (Exact Match) - เร็วสุด
    if (assetMap[cleanedName]) { 
      finalBorrow = assetMap[cleanedName].status; 
      finalSerial = assetMap[cleanedName].serial;
      isInAssetSheet = true;
    } else {
      // 2. ถ้าไม่เจอ -> ใช้ระบบ "Fuzzy Match" (หาความใกล้เคียง)
      for (let i = 0; i < assetKeys.length; i++) {
        let assetKey = assetKeys[i];
        
        // คำนวณความต่าง
        let dist = getEditDistance(cleanedName, assetKey);
        
        // กติกา: ชื่อยาวเกิน 5 ยอมให้ผิด 2 จุด, ถ้าน้อยกว่ายอมให้ผิด 1 จุด
        let allowedErrors = cleanedName.length > 5 ? 2 : 1;

        if (dist <= allowedErrors) {
          // เจอคู่ที่ใกล้เคียงแล้ว!
          finalBorrow = assetMap[assetKey].status;
          finalSerial = assetMap[assetKey].serial;
          isInAssetSheet = true;
          break; // เจอแล้วหยุดหา
        }
      }
    }

    // อัปเดตข้อมูลจาก Log (Database)
    if (dbMap[id]) {
      if (dbMap[id].borrowStatus !== 'ยังไม่ยืม') {
        finalBorrow = dbMap[id].borrowStatus;
      } else if (finalBorrow === 'ยังไม่ยืม' && isInAssetSheet) {
        // ถ้า DB ยังไม่ยืม แต่ Asset มีชื่อ -> ยึด Asset
        finalBorrow = 'ยืมอยู่'; 
      }
      
      finalDoc = dbMap[id].docStatus;
      if (dbMap[id].serial !== '-') finalSerial = dbMap[id].serial;
      finalFiles = dbMap[id].files;
    }

    allPeople.push({ 
      type: type, 
      no: no, 
      id: id, 
      name: name, 
      room: room, 
      source_sheet: source, 
      serial: finalSerial, 
      borrowStatus: finalBorrow, 
      docStatus: finalDoc, 
      files: finalFiles, 
      inAsset: isInAssetSheet 
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

  return allPeople;
}

// ==========================================
// 4. ฟังก์ชันสำหรับระบบจับคู่ชื่อ (AUDIT & FIX)
// ==========================================

function getAssetAuditData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // 1. ดึงข้อมูล "คนในระบบ" (Students/Teachers) มาเก็บไว้เทียบ
  let peopleList = [];
  
  const fetchPeople = (sheets) => {
    if(!Array.isArray(sheets)) sheets = [sheets];
    sheets.forEach(sheetName => {
      let sheet = ss.getSheetByName(sheetName);
      if(sheet && sheet.getLastRow() > 1) {
        // อ่านถึง Col D (No, ID, Name, Room)
        let data = sheet.getRange(2, 1, sheet.getLastRow()-1, 4).getValues();
        data.forEach(r => {
          if(r[2]) { // มีชื่อ
             peopleList.push({
               id: r[1],
               name: r[2],
               room: r[3],
               sheet: sheetName,
               norm: normalizeName(r[2]) // ชื่อที่ล้างแล้ว
             });
          }
        });
      }
    });
  };

  fetchPeople(SHEET_NAMES.STUDENTS);
  fetchPeople(SHEET_NAMES.TEACHERS);

  // สร้าง Set ของชื่อคนในระบบ (เอาไว้เช็คว่ามีตัวตนไหม)
  let peopleNormSet = new Set(peopleList.map(p => p.norm));

  // 2. ดึงข้อมูล "Asset" มาตรวจสอบหาคนตกหล่น
  let assetSheet = ss.getSheetByName(SHEET_NAMES.ASSETS);
  let orphans = [];

  if(assetSheet && assetSheet.getLastRow() > 1) {
    let assetData = assetSheet.getRange(2, 1, assetSheet.getLastRow()-1, 10).getValues(); 
    
    assetData.forEach(r => {
      let assetName = r[4]; // ชื่อใน Asset (Col E)
      let serial = r[2] || r[3]; // Serial
      
      if(assetName) {
        let assetNorm = normalizeName(assetName);
        
        // ถ้าชื่อใน Asset ไม่มีในระบบรายชื่อ (และไม่ใช่ค่าว่าง)
        if(!peopleNormSet.has(assetNorm) && assetNorm !== "") {
          
          // หาคู่ที่หน้าตาคล้ายๆ กัน (Fuzzy Suggestions)
          let suggestions = [];
          
          for(let p of peopleList) {
             let dist = getEditDistance(assetNorm, p.norm);
             // เกณฑ์ความเหมือน: ผิดได้ไม่เกิน 2-3 จุด
             let threshold = assetNorm.length > 6 ? 3 : 2; 
             
             if(dist <= threshold) {
               suggestions.push({
                 name: p.name, // ชื่อที่ถูกต้องในระบบ
                 sheet: p.sheet,
                 diff: dist
               });
             }
          }
          
          // เรียงเอาคนที่เหมือนที่สุดขึ้นก่อน
          suggestions.sort((a,b) => a.diff - b.diff);

          orphans.push({
            assetName: assetName, // ชื่อที่ผิด (จาก Asset)
            serial: serial ? serial.toString() : '-',
            suggestions: suggestions.slice(0, 5) // เอามาแค่ 5 อันดับแรก
          });
        }
      }
    });
  }
  
  return orphans;
}

// ฟังก์ชันแก้ชื่อใน "ทะเบียนทรัพย์สิน" (ตามคำสั่ง)
function adminFixAssetName(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.ASSETS);
  
  if(!sheet) return { success: false, message: "ไม่พบแผ่นงานรายงานทะเบียนทรัพย์สิน" };
  
  try {
    const rows = sheet.getDataRange().getDisplayValues();
    let rowToUpdate = -1;
    
    // วนหาแถวที่ต้องแก้ในชีตทรัพย์สิน
    for(let i=0; i<rows.length; i++) {
       // เงื่อนไข: ชื่อตรงกับชื่อเดิม (ที่ผิด) AND Serial ตรงกัน
       let currentName = rows[i][4];
       let currentSerial = rows[i][2] || rows[i][3];
       
       if(currentName == data.oldAssetName && String(currentSerial) == String(data.serial)) {
          rowToUpdate = i + 1;
          break;
       }
    }

    if(rowToUpdate > -1) {
      // แก้ไขชื่อใน Col E (คอลัมน์ที่ 5) เป็นชื่อที่ถูกต้อง
      sheet.getRange(rowToUpdate, 5).setValue(data.correctName);
      return { success: true, message: "แก้ไขชื่อในทะเบียนทรัพย์สินเรียบร้อยแล้ว" };
    } else {
      return { success: false, message: "ไม่พบข้อมูลเดิมในทะเบียนทรัพย์สิน (อาจถูกแก้ไขไปแล้ว)" };
    }
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}
// ฟังก์ชันใหม่: แก้ไขชื่อทีละหลายรายการ (Bulk Update) - ปรับปรุงใหม่
function adminFixAssetNameBulk(updateList) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.ASSETS);
  if(!sheet) return { success: false, message: "ไม่พบแผ่นงานรายงานทะเบียนทรัพย์สิน" };

  // ใช้ LockService ป้องกันการชนกันของข้อมูล
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // รอคิวได้สูงสุด 10 วินาที

    const dataRange = sheet.getDataRange();
    const rows = dataRange.getDisplayValues(); 
    let updateCount = 0;

    // วนลูปรายการที่จะแก้
    updateList.forEach(item => {
      // เตรียมข้อมูลเปรียบเทียบ (แปลงเป็น String และตัดช่องว่างหน้าหลังออก)
      let targetName = String(item.oldAssetName).trim();
      let targetSerial = String(item.serial).trim();

      for(let i=1; i<rows.length; i++) {
        let currentName = String(rows[i][4]).trim();           // Col E
        let currentSerial = String(rows[i][2] || rows[i][3]).trim(); // Col C or D

        // เปรียบเทียบแบบแม่นยำขึ้น
        if(currentName === targetName && currentSerial === targetSerial) {
          // อัปเดตข้อมูล (i + 1 คือแถว, 5 คือคอลัมน์ E)
          sheet.getRange(i + 1, 5).setValue(item.correctName);
          updateCount++;
          // อัปเดตค่าในตัวแปร rows ด้วย เพื่อกันพลาดกรณีชื่อซ้ำในรอบถัดไป
          rows[i][4] = item.correctName; 
          break; 
        }
      }
    });

    // *** หัวใจสำคัญ: บังคับบันทึกข้อมูลลง Sheet ทันที ***
    SpreadsheetApp.flush(); 

    return { success: true, count: updateCount };

  } catch(e) {
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock(); // ปล่อยล็อกเสมอ
  }
}

// ==========================================
// 5. STANDARD HELPERS (Form, Auth, etc.)
// ==========================================

function processForm(formObject) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheetData = ss.getSheetByName(SHEET_NAMES.DATA_DB);
  const folder = DriveApp.getFolderById(FOLDER_ID);
  
  try {
    const timestamp = new Date();
    
    const uploadFile = (fileBlob, prefix) => {
      if (!fileBlob || fileBlob.name == "") return "";
      let fileName = prefix + "_" + formObject.userName + "_" + timestamp.getTime();
      return folder.createFile(fileBlob).setName(fileName).getUrl();
    };

    let url_agreement = uploadFile(formObject.file_agreement, "AGREEMENT");
    let url_card_std = "", url_card_parent = "", url_house = "", parent_phone = "";

    if (formObject.userType === 'student') {
      url_card_std = uploadFile(formObject.file_card_std, "CARD_STD");
      url_card_parent = uploadFile(formObject.file_card_parent, "CARD_PARENT");
      url_house = uploadFile(formObject.file_house, "HOUSE");
      parent_phone = "'" + formObject.parent_phone;
    }

    let statusToSave = formObject.statusSelect;
    if (url_agreement !== "") {
      statusToSave = statusToSave + " | รอตรวจสอบเอกสาร"; 
    }

    sheetData.appendRow([
      timestamp, 
      formObject.userId, 
      formObject.userName, 
      formObject.userType, 
      formObject.userRoom, 
      formObject.userSerial, 
      "USER_UPDATE", 
      formObject.note || "", 
      statusToSave, 
      url_agreement, url_card_std, url_card_parent, url_house, parent_phone
    ]);
    
    return { success: true, message: "บันทึกข้อมูลเรียบร้อย" };
  } catch (error) { 
    return { success: false, message: "เกิดข้อผิดพลาด: " + error.toString() }; 
  }
}

function verifyAdmin(u, p) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.ADMIN);
  if (!sheet) return { success: false, message: "No Admin Sheet" };
  const data = sheet.getDataRange().getDisplayValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(u).trim() && String(data[i][1]).trim() === String(p).trim()) {
      return { success: true, role: 'admin' };
    }
  }
  return { success: false, message: "Login Failed" };
}

function verifyAdvisor(u, p) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.ADVISOR);
  if (!sheet) return { success: false, message: "ไม่พบชีตครูที่ปรึกษา" };
  const data = sheet.getDataRange().getDisplayValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(u).trim() && String(data[i][1]).trim() === String(p).trim()) {
      return { success: true, role: 'advisor', level: data[i][2], room: data[i][3], name: data[i][4] || "คุณครูที่ปรึกษา" };
    }
  }
  return { success: false, message: "ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง" };
}

function adminUpdateData(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheetData = ss.getSheetByName(SHEET_NAMES.DATA_DB);
  try {
    const timestamp = new Date();
    let editor = data.editorRole === 'advisor' ? ("ADVISOR: " + data.editorName) : "ADMIN_EDIT";
    
    if (data.borrowStatusSelect) {
      sheetData.appendRow([
        timestamp, data.userId, data.userName, data.userType, data.userRoom, 
        data.userSerial, editor, data.note, data.borrowStatusSelect, "", "", "", "", ""
      ]);
    }
    
    if (data.docStatusSelect && data.docStatusSelect !== "") {
       Utilities.sleep(100); 
       sheetData.appendRow([
         new Date(), data.userId, data.userName, data.userType, data.userRoom, 
         data.userSerial, editor, data.note, data.docStatusSelect, "", "", "", "", ""
       ]);
    }
    return { success: true, message: "อัปเดตข้อมูลสำเร็จ" };
  } catch (e) { return { success: false, message: e.toString() }; }
}

function adminDeleteUser(data) {
  if (data.editorRole === 'advisor') return { success: false, message: "ครูที่ปรึกษาไม่ได้รับอนุญาตให้ลบข้อมูล" };
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(data.source_sheet);
  if (!sheet) return { success: false, message: "ไม่พบแผ่นงาน" };
  try {
    const rows = sheet.getDataRange().getDisplayValues();
    let rowToDelete = -1;
    for (let i = 0; i < rows.length; i++) {
      if (data.source_sheet === SHEET_NAMES.TEACHERS) { 
        if (rows[i][1] == data.name) { rowToDelete = i + 1; break; } 
      } else { 
        if (rows[i][1] == data.id) { rowToDelete = i + 1; break; } 
      }
    }
    if (rowToDelete > -1) { 
      sheet.deleteRow(rowToDelete); 
      return { success: true, message: "ลบข้อมูลเรียบร้อยแล้ว" }; 
    } else { return { success: false, message: "ไม่พบข้อมูล" }; }
  } catch (e) { return { success: false, message: "Error: " + e.toString() }; }
}

function adminAddUser(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(data.targetSheet); 
  if (!sheet) return { success: false, message: "ไม่พบแผ่นงาน" };
  try {
    const nextNo = sheet.getLastRow(); 
    if (data.targetSheet === SHEET_NAMES.TEACHERS) {
      sheet.appendRow([nextNo, data.name]); 
    } else {
      sheet.appendRow([nextNo, data.id, data.name, data.room]);
    }
    return { success: true, message: "เพิ่มรายชื่อเรียบร้อยแล้ว" };
  } catch (e) { return { success: false, message: "Error: " + e.toString() }; }
}

// ==========================================
// 6. DASHBOARD STATS SYSTEM (เพิ่มใหม่)
// ==========================================

function getDashboardStats() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const assetSheet = ss.getSheetByName(SHEET_NAMES.ASSETS);
  
  // ค่าเริ่มต้น
  let stats = {
    total: 0,      // ทั้งหมด
    borrowed: 0,   // ยืมแล้ว
    available: 0   // ว่าง
  };

  if (assetSheet) {
    const lastRow = assetSheet.getLastRow();
    if (lastRow > 1) {
      // 1. ยอดทั้งหมด = จำนวนแถวทั้งหมด ลบ 1 (หัวตาราง)
      stats.total = lastRow - 1;

      // 2. หายอดยืม = นับแถวที่มี "ชื่อ-สกุล" (Col E)
      // อ่านข้อมูลเฉพาะคอลัมน์ E (Column index 5)
      const data = assetSheet.getRange(2, 5, stats.total, 1).getValues();
      
      // นับจำนวนช่องที่ไม่ว่าง
      let count = 0;
      for (let i = 0; i < data.length; i++) {
        // ตรวจสอบว่ามีชื่อไหม (ตัดช่องว่างหน้าหลังออกแล้วเช็คว่าไม่ว่าง)
        if (String(data[i][0]).trim() !== "") {
          count++;
        }
      }
      stats.borrowed = count;
      
      // 3. ยอดว่าง
      stats.available = stats.total - stats.borrowed;
    }
  }
  
  return stats;
}
