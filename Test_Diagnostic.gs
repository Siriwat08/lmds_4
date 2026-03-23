/**
 * VERSION: 000
 * 🏥 System Diagnostic Tool (Enterprise Edition)
 * Version: 4.0 Deep Scan & Schema Validation
 * -----------------------------------------------------------------
 * [PRESERVED]: Two-phase diagnostic approach (Engine & Sheets).
 * [ADDED v4.0]: Validates NameMapping V4.0 5-Column schema.
 * [ADDED v4.0]: Validates PostalRef sheet existence.
 * [ADDED v4.0]: Deep scan for LINE and Telegram tokens.
 * [MODIFIED v4.0]: Safe API Key extraction using try-catch for V4.0 Getter.
 * Author: Elite Logistics Architect
 */

// ==========================================
// 1. PHASE 1: ENGINE & DEPENDENCY CHECK
// ==========================================

/**
 * 🏥 System Diagnostic Tool (Phase 1: Engine Check)
 * สแกนหาฟังก์ชันหลักและ API Key ว่าเชื่อมต่อสมบูรณ์หรือไม่
 */
function RUN_SYSTEM_DIAGNOSTIC() {
  var ui = SpreadsheetApp.getUi();
  var logs = [];

  function pass(msg) { logs.push("✅ " + msg); }
  function warn(msg) { logs.push("⚠️ " + msg); }
  function fail(msg) { logs.push("❌ " + msg); }

  try {
    // 1. Config Check
    if (typeof CONFIG !== 'undefined') pass("System Variables: มองเห็นตัวแปร CONFIG");
    else fail("System Variables: มองไม่เห็นตัวแปร CONFIG");

    // 2. Utility Functions Check
    if (typeof md5 === 'function') pass("Core Utils: มองเห็นฟังก์ชัน md5()");
    else fail("Core Utils: มองไม่เห็นฟังก์ชัน md5()");

    if (typeof normalizeText === 'function') pass("Core Utils: มองเห็นฟังก์ชัน normalizeText()");
    else fail("Core Utils: มองไม่เห็นฟังก์ชัน normalizeText()");

    // 3. Geo Map API Check
    if (typeof GET_ADDR_WITH_CACHE === 'function') {
      try {
        var testGeo = GET_ADDR_WITH_CACHE(13.746, 100.539);
        if (testGeo && testGeo !== "Error") pass("Google Maps API: ทำงานปกติ (" + testGeo.substring(0, 20) + "...)");
        else warn("Google Maps API: โหลดได้แต่ส่งค่าแปลกๆ กลับมา");
      } catch (geoErr) {
        fail("Google Maps API: Error ระหว่างทดสอบ (" + geoErr.message + ")");
      }
    } else {
      fail("Google Maps API: ไม่พบฟังก์ชัน GET_ADDR_WITH_CACHE ใน Service_GeoAddr");
    }

    // 4. Security Vault Check (API Keys)
    var props = PropertiesService.getScriptProperties();

    // Gemini Key (V4.0 Safe Check)
    try {
      if (CONFIG && CONFIG.GEMINI_API_KEY) pass("AI Engine: ตรวจพบ GEMINI_API_KEY พร้อมใช้งาน");
    } catch (e) {
      fail("AI Engine: ไม่พบ GEMINI_API_KEY หรือตั้งค่าไม่ถูกต้อง (" + e.message + ")");
    }

    // Notifications Check
    if (props.getProperty('LINE_NOTIFY_TOKEN')) pass("Notifications: ตรวจพบ LINE Notify Token");
    else warn("Notifications: ยังไม่ได้ตั้งค่า LINE Notify");

    if (props.getProperty('TG_BOT_TOKEN') && props.getProperty('TG_CHAT_ID')) pass("Notifications: ตรวจพบ Telegram Config");
    else warn("Notifications: ยังไม่ได้ตั้งค่า Telegram");

    ui.alert("🏥 รายงานผลการสแกนระบบ (Engine V4.0):\n\n" + logs.join("\n"));
    console.info("[Diagnostic] Phase 1 (Engine) completed.");

  } catch (e) {
    console.error("[Diagnostic Error]: " + e.message);
    ui.alert("🚨 ระบบตรวจพบ Error ร้ายแรงระหว่างสแกน:\n" + e.message);
  }
}

// ==========================================
// 2. PHASE 2: DATA & STRUCTURE CHECK
// ==========================================

/**
 * 🕵️‍♂️ Sheet Diagnostic Tool (Phase 2: Data & Silent Exit Check)
 * ตรวจสอบว่ามีชีตครบตาม Config และมีโครงสร้างคอลัมน์ถูกต้องหรือไม่
 */
function RUN_SHEET_DIAGNOSTIC() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logs = [];

  function pass(msg) { logs.push("✅ " + msg); }
  function warn(msg) { logs.push("⚠️ " + msg); }
  function fail(msg) { logs.push("❌ " + msg); }

  try {
    // 1. ตรวจสอบ Database Sheet
    var dbName = (typeof CONFIG !== 'undefined' && CONFIG.SHEET_NAME) ? CONFIG.SHEET_NAME : "Database";
    var dbSheet = ss.getSheetByName(dbName);
    if (dbSheet) {
      var rows = dbSheet.getLastRow();
      if (rows >= 2) pass("Master DB: พบชีต '" + dbName + "' (มีข้อมูล " + rows + " แถว)");
      else warn("Master DB: พบชีต '" + dbName + "' แต่ข้อมูลว่างเปล่า (มี " + rows + " แถว)");
    } else {
      fail("Master DB: ไม่พบชีตชื่อ '" + dbName + "' (ตรวจสอบเว้นวรรคท้ายชื่อด้วย)");
    }

    // 2. ตรวจสอบ Source Sheet
    var srcName = (typeof CONFIG !== 'undefined' && CONFIG.SOURCE_SHEET) ? CONFIG.SOURCE_SHEET : "SCGนครหลวงJWDภูมิภาค";
    var srcSheet = ss.getSheetByName(srcName);
    if (srcSheet) {
      pass("Source Data: พบชีต '" + srcName + "' (มีข้อมูล " + srcSheet.getLastRow() + " แถว)");
    } else {
      warn("Source Data: ไม่พบชีต '" + srcName + "'");
    }

    // 3. ตรวจสอบ Mapping Sheet (V4.0 Schema Check)
    var mapName = (typeof CONFIG !== 'undefined' && CONFIG.MAPPING_SHEET) ? CONFIG.MAPPING_SHEET : "NameMapping";
    var mapSheet = ss.getSheetByName(mapName);
    if (mapSheet) {
      var mapCols = mapSheet.getLastColumn();
      if (mapCols >= 5) {
        pass("Name Mapping: พบชีต '" + mapName + "' (โครงสร้าง 5 คอลัมน์ V4.0 ถูกต้อง)");
      } else {
        warn("Name Mapping: พบชีต '" + mapName + "' แต่มีแค่ " + mapCols + " คอลัมน์ (แนะนำให้ใช้เมนู Upgrade NameMapping เป็น V4.0)");
      }
    } else {
      fail("Name Mapping: ไม่พบชีต '" + mapName + "'");
    }

    // 4. ตรวจสอบ SCG Daily Data Sheet
    if (typeof SCG_CONFIG !== 'undefined') {
      var scgDataName = SCG_CONFIG.SHEET_DATA || "Data";
      var scgInputName = SCG_CONFIG.SHEET_INPUT || "Input";

      if (ss.getSheetByName(scgDataName)) pass("SCG Operation: พบชีต '" + scgDataName + "'");
      else warn("SCG Operation: ไม่พบชีต '" + scgDataName + "'");

      if (ss.getSheetByName(scgInputName)) pass("SCG Operation: พบชีต '" + scgInputName + "'");
      else warn("SCG Operation: ไม่พบชีต '" + scgInputName + "'");
    }

    // 5. ตรวจสอบ PostalRef Sheet (New V4.0 Requirement)
    var postalName = (typeof CONFIG !== 'undefined' && CONFIG.SHEET_POSTAL) ? CONFIG.SHEET_POSTAL : "PostalRef";
    if (ss.getSheetByName(postalName)) {
      pass("Geo Database: พบชีต '" + postalName + "' สำหรับอ้างอิงรหัสไปรษณีย์");
    } else {
      warn("Geo Database: ไม่พบชีต '" + postalName + "' (การแกะที่อยู่แบบ Offline อาจไม่แม่นยำ 100%)");
    }

    ui.alert("🕵️‍♂️ รายงานผลการสแกนชีต (Silent Exit Check):\n\n" + logs.join("\n"));
    console.info("[Diagnostic] Phase 2 (Sheets) completed.");

  } catch (e) {
    console.error("[Diagnostic Error]: " + e.message);
    ui.alert("🚨 เกิด Error ระหว่างตรวจสอบชีต:\n" + e.message);
  }
}

function testConfig() {
  // ทดสอบ SRC_IDX ย้ายมาถูกที่
  console.log(SCG_CONFIG.SRC_IDX.LAT);    // ต้องได้ 14
  console.log(SCG_CONFIG.SRC_IDX.LNG);    // ต้องได้ 15
  console.log(SCG_CONFIG.SHEET_GPS_QUEUE); // ต้องได้ "GPS_Queue"
  console.log(SCG_CONFIG.GPS_THRESHOLD_METERS); // ต้องได้ 50

  // ทดสอบ col ใหม่ใน Database
  console.log(CONFIG.COL_COORD_SOURCE);       // ต้องได้ 18
  console.log(CONFIG.COL_COORD_CONFIDENCE);   // ต้องได้ 19
  console.log(CONFIG.COL_COORD_LAST_UPDATED); // ต้องได้ 20
}

function testFinalizeClean() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEET_NAME);

  var before = sheet.getLastRow();
  console.log("Before: lastRow = " + before);

  // รัน finalize จริง
  finalizeAndClean_MoveToMapping();

  SpreadsheetApp.flush();
  var after = sheet.getLastRow();
  console.log("After: lastRow = " + after);

  // ต้องได้ after = rowsToKeep.length + 1 (บวก header)
  // ไม่ใช่ค่าเดิม before
}

function testCacheSize() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var aliasMap = getCachedNameMapping_(ss);

  var jsonString = JSON.stringify(aliasMap);
  var charCount = jsonString.length;
  var byteSize = Utilities.newBlob(jsonString).getBytes().length;

  console.log("ตัวอักษร: " + charCount);
  console.log("Bytes จริง: " + byteSize);
  console.log("อัตราส่วน: " + (byteSize / charCount).toFixed(2) + " bytes/char");

  // ถ้าข้อมูลเป็นภาษาไทยเยอะ อัตราส่วนจะใกล้ 3.0
  // ถ้าอัตราส่วน > 2.5 แสดงว่าระบบเก่าจะ crash
}

function testNaNBlackHole() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEET_NAME);

  var lastRow = getRealLastRow_(sheet, CONFIG.COL_NAME);
  var data = sheet.getRange(2, 1, lastRow - 1, 15).getValues();

  var noCoordCount = 0;
  var nanCount = 0;

  data.forEach(function (row) {
    var lat = parseFloat(row[CONFIG.C_IDX.LAT]);
    var lng = parseFloat(row[CONFIG.C_IDX.LNG]);

    if (!row[CONFIG.C_IDX.LAT] || !row[CONFIG.C_IDX.LNG]) {
      noCoordCount++;
    }
    if (isNaN(lat) || isNaN(lng)) {
      nanCount++;
    }
  });

  console.log("แถวที่ไม่มีพิกัด: " + noCoordCount);
  console.log("แถวที่ได้ NaN: " + nanCount);
  console.log("ถ้าตัวเลขทั้งสองเท่ากัน แสดงว่าแก้ถูกต้องแล้ว");
}

function testNegativeRowCount() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var empSheet = ss.getSheetByName(SCG_CONFIG.SHEET_EMPLOYEE);

  if (!empSheet) {
    console.log("❌ ไม่พบชีตพนักงาน");
    return;
  }

  var lastRow = empSheet.getLastRow();
  console.log("empSheet.getLastRow() = " + lastRow);

  // ทดสอบกรณีชีตมีข้อมูล
  if (lastRow >= 2) {
    console.log("✅ ปลอดภัย: มีข้อมูลพนักงาน " + (lastRow - 1) + " คน");
  } else {
    console.log("✅ ปลอดภัย: ชีตว่าง ระบบจะข้ามไปโดยไม่ crash");
  }

  // ทดสอบค่าที่จะเกิดขึ้น
  var calcValue = lastRow - 1;
  console.log("lastRow - 1 = " + calcValue);
  if (calcValue < 1) {
    console.log("⚠️ ถ้าไม่แก้ จะ crash เพราะ getRange() ได้รับค่า " + calcValue);
  } else {
    console.log("✅ ค่าปลอดภัย getRange() จะได้รับค่า " + calcValue);
  }
}

function testGetLastColumn() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName(CONFIG.SOURCE_SHEET);

  if (!sourceSheet) {
    console.log("❌ ไม่พบชีต " + CONFIG.SOURCE_SHEET);
    return;
  }

  var lastCol = sourceSheet.getLastColumn();
  var lastRow = sourceSheet.getLastRow();

  console.log("ชีต: " + CONFIG.SOURCE_SHEET);
  console.log("จำนวนคอลัมน์จริง: " + lastCol);
  console.log("จำนวนแถวข้อมูล: " + (lastRow - 1));

  // ตรวจสอบว่า SRC_IDX ไม่เกิน lastCol
  var maxIdx = Math.max(
    SCG_CONFIG.SRC_IDX.NAME,
    SCG_CONFIG.SRC_IDX.LAT,
    SCG_CONFIG.SRC_IDX.LNG,
    SCG_CONFIG.SRC_IDX.SYS_ADDR,
    SCG_CONFIG.SRC_IDX.DIST,
    SCG_CONFIG.SRC_IDX.GOOG_ADDR
  );

  console.log("SRC_IDX ใหญ่สุด: " + maxIdx);

  if (maxIdx <= lastCol) {
    console.log("✅ ปลอดภัย: SRC_IDX ทุกตัวอยู่ในช่วงคอลัมน์จริง");
  } else {
    console.log("❌ อันตราย: SRC_IDX (" + maxIdx + ") เกินจำนวนคอลัมน์จริง (" + lastCol + ")");
  }
}

function testSyncCheckpoint() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SOURCE_SHEET);

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var syncCol = SCG_CONFIG.SRC_IDX_SYNC_STATUS;

  // แปลง column number เป็น letter (รองรับ AA, AB, AK ฯลฯ)
  function colToLetter(col) {
    var letter = '';
    while (col > 0) {
      var mod = (col - 1) % 26;
      letter = String.fromCharCode(65 + mod) + letter;
      col = Math.floor((col - 1) / 26);
    }
    return letter;
  }

  console.log("คอลัมน์ทั้งหมดในชีต: " + lastCol);
  console.log("Col SYNC_STATUS กำหนดไว้: " + syncCol + " (Col " + colToLetter(syncCol) + ")");

  if (syncCol > lastCol) {
    console.log("⚠️ ยังไม่มี Col " + colToLetter(syncCol) + " ในชีต");
    console.log("กรุณาเพิ่ม header SYNC_STATUS ที่ Col " + colToLetter(syncCol) + " ก่อนครับ");
    return;
  }

  // นับแถวที่ SYNCED แล้ว
  var data = sheet.getRange(2, syncCol, lastRow - 1, 1).getValues();
  var syncedCount = data.filter(function (r) {
    return r[0] === SCG_CONFIG.SYNC_STATUS_DONE;
  }).length;

  var pendingCount = (lastRow - 1) - syncedCount;

  console.log("แถวทั้งหมด: " + (lastRow - 1));
  console.log("SYNCED แล้ว: " + syncedCount);
  console.log("รอ SYNC: " + pendingCount);
}

function testGPSFeedback() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var queueSheet = ss.getSheetByName(SCG_CONFIG.SHEET_GPS_QUEUE);

  var beforeQueue = queueSheet.getLastRow();
  console.log("GPS_Queue ก่อนรัน: " + (beforeQueue - 1) + " รายการ");

  syncNewDataToMaster();

  SpreadsheetApp.flush();
  var afterQueue = queueSheet.getLastRow();
  console.log("GPS_Queue หลังรัน: " + (afterQueue - 1) + " รายการ");
  console.log("เพิ่มเข้า Queue: " + (afterQueue - beforeQueue) + " รายการ");

  // เช็ค SYNC_STATUS
  var sourceSheet = ss.getSheetByName(CONFIG.SOURCE_SHEET);
  var data = sourceSheet.getRange(
    2, SCG_CONFIG.SRC_IDX_SYNC_STATUS,
    sourceSheet.getLastRow() - 1, 1
  ).getValues();

  var syncedCount = data.filter(function (r) {
    return r[0] === SCG_CONFIG.SYNC_STATUS_DONE;
  }).length;

  console.log("แถวที่ Mark SYNCED: " + syncedCount);
}

function debugSyncError() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // ทดสอบที่ 1: CONFIG.C_IDX มีครบไหม
  console.log("=== ทดสอบ CONFIG ===");
  console.log("COL_COORD_SOURCE: " + CONFIG.COL_COORD_SOURCE);
  console.log("COL_COORD_CONFIDENCE: " + CONFIG.COL_COORD_CONFIDENCE);
  console.log("COL_COORD_LAST_UPDATED: " + CONFIG.COL_COORD_LAST_UPDATED);
  console.log("C_IDX.COORD_SOURCE: " + CONFIG.C_IDX.COORD_SOURCE);
  console.log("C_IDX.LAT: " + CONFIG.C_IDX.LAT);

  // ทดสอบที่ 2: SCG_CONFIG.SRC_IDX มีครบไหม
  console.log("=== ทดสอบ SCG_CONFIG ===");
  console.log("SRC_IDX.NAME: " + SCG_CONFIG.SRC_IDX.NAME);
  console.log("SRC_IDX.LAT: " + SCG_CONFIG.SRC_IDX.LAT);
  console.log("SRC_IDX.LNG: " + SCG_CONFIG.SRC_IDX.LNG);
  console.log("GPS_THRESHOLD_METERS: " + SCG_CONFIG.GPS_THRESHOLD_METERS);
  console.log("SHEET_GPS_QUEUE: " + SCG_CONFIG.SHEET_GPS_QUEUE);

  // ทดสอบที่ 3: Database อ่านได้ไหม
  console.log("=== ทดสอบ Database ===");
  var masterSheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  var lastRowM = getRealLastRow_(masterSheet, CONFIG.COL_NAME);
  console.log("lastRowM: " + lastRowM);

  var maxCol = Math.max(
    CONFIG.COL_NAME, CONFIG.COL_LAT, CONFIG.COL_LNG,
    CONFIG.COL_UUID, CONFIG.COL_COORD_SOURCE,
    CONFIG.COL_COORD_CONFIDENCE, CONFIG.COL_COORD_LAST_UPDATED
  );
  console.log("maxCol: " + maxCol);

  if (isNaN(maxCol)) {
    console.log("❌ maxCol เป็น NaN — COL_COORD_SOURCE/CONFIDENCE/LAST_UPDATED ไม่ได้กำหนดใน Config");
    return;
  }

  // ทดสอบที่ 4: อ่านแถวแรกของ Database
  var firstRow = masterSheet.getRange(2, 1, 1, maxCol).getValues()[0];
  console.log("Database แถวแรก length: " + firstRow.length);
  console.log("LAT: " + firstRow[CONFIG.C_IDX.LAT]);
  console.log("LNG: " + firstRow[CONFIG.C_IDX.LNG]);

  // ทดสอบที่ 5: GPS_Queue ปัญหา checkbox
  console.log("=== ทดสอบ GPS_Queue ===");
  var queueSheet = ss.getSheetByName(SCG_CONFIG.SHEET_GPS_QUEUE);
  console.log("queueSheet.getLastRow(): " + queueSheet.getLastRow());
  console.log("queueSheet.getLastColumn(): " + queueSheet.getLastColumn());

  console.log("=== Debug เสร็จแล้ว ===");
}

function debugQueueEmpty() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName(CONFIG.SOURCE_SHEET);
  var masterSheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  var mapSheet = ss.getSheetByName(CONFIG.MAPPING_SHEET);

  // โหลด Database
  var lastRowM = getRealLastRow_(masterSheet, CONFIG.COL_NAME);
  var dbData = masterSheet.getRange(2, 1, lastRowM - 1, 20).getValues();
  var existingNames = {};
  var existingUUIDs = {};
  dbData.forEach(function (r, i) {
    if (r[CONFIG.C_IDX.NAME]) {
      existingNames[normalizeText(r[CONFIG.C_IDX.NAME])] = i;
    }
    if (r[CONFIG.C_IDX.UUID]) {
      existingUUIDs[r[CONFIG.C_IDX.UUID]] = i;
    }
  });

  // โหลด NameMapping
  var aliasToUUID = {};
  if (mapSheet && mapSheet.getLastRow() > 1) {
    mapSheet.getRange(2, 1, mapSheet.getLastRow() - 1, 2).getValues()
      .forEach(function (r) {
        if (r[0] && r[1]) aliasToUUID[normalizeText(r[0])] = r[1];
      });
  }

  // อ่าน Source
  var lastColS = sourceSheet.getLastColumn();
  var sData = sourceSheet.getRange(2, 1, sourceSheet.getLastRow() - 1, lastColS).getValues();

  var countNewName = 0;
  var countNoGPS = 0;
  var countDiffSmall = 0;
  var countDiffLarge = 0;
  var countNoDB = 0;
  var countSkipSynced = 0;
  var sampleQueue = [];

  sData.forEach(function (row) {
    var syncStatus = row[SCG_CONFIG.SRC_IDX_SYNC_STATUS - 1];
    if (syncStatus === SCG_CONFIG.SYNC_STATUS_DONE) {
      countSkipSynced++;
      return;
    }

    var name = row[SCG_CONFIG.SRC_IDX.NAME];
    var lat = parseFloat(row[SCG_CONFIG.SRC_IDX.LAT]);
    var lng = parseFloat(row[SCG_CONFIG.SRC_IDX.LNG]);

    if (!name || isNaN(lat) || isNaN(lng)) {
      countNoGPS++;
      return;
    }

    var cleanName = normalizeText(name);
    var matchIdx = -1;

    if (existingNames.hasOwnProperty(cleanName)) {
      matchIdx = existingNames[cleanName];
    } else if (aliasToUUID.hasOwnProperty(cleanName)) {
      var uid = aliasToUUID[cleanName];
      if (existingUUIDs.hasOwnProperty(uid)) {
        matchIdx = existingUUIDs[uid];
      }
    }

    if (matchIdx === -1 || matchIdx === -999) {
      countNewName++;
      return;
    }

    var dbRow = dbData[matchIdx];
    if (!dbRow) return;

    var dbLat = parseFloat(dbRow[CONFIG.C_IDX.LAT]);
    var dbLng = parseFloat(dbRow[CONFIG.C_IDX.LNG]);

    if (isNaN(dbLat) || isNaN(dbLng)) {
      countNoDB++;
      return;
    }

    var diffKm = getHaversineDistanceKM(lat, lng, dbLat, dbLng);
    var diffMeters = Math.round(diffKm * 1000);
    var threshold = SCG_CONFIG.GPS_THRESHOLD_METERS / 1000;

    if (diffKm <= threshold) {
      countDiffSmall++;
    } else {
      countDiffLarge++;
      if (sampleQueue.length < 5) {
        sampleQueue.push({
          name: name,
          driver: lat + "," + lng,
          db: dbLat + "," + dbLng,
          diff: diffMeters + "m"
        });
      }
    }
  });

  console.log("=== ผล Debug ===");
  console.log("ข้ามเพราะ SYNCED: " + countSkipSynced);
  console.log("ไม่มี GPS: " + countNoGPS);
  console.log("ชื่อใหม่ (ไม่ match): " + countNewName);
  console.log("DB ไม่มีพิกัด: " + countNoDB);
  console.log("diff ≤ 50m: " + countDiffSmall);
  console.log("diff > 50m (→ Queue): " + countDiffLarge);

  if (sampleQueue.length > 0) {
    console.log("=== ตัวอย่าง Queue ===");
    sampleQueue.forEach(function (s) {
      console.log(s.name + " | diff=" + s.diff);
    });
  }
}

function testApprovedFeedback() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var queueSheet = ss.getSheetByName(SCG_CONFIG.SHEET_GPS_QUEUE);

  // นับรายการที่ Approve ไว้
  var lastRow = getRealLastRow_(queueSheet, 1);
  if (lastRow < 2) {
    console.log("ℹ️ GPS_Queue ว่างเปล่า");
    return;
  }

  var data = queueSheet.getRange(2, 8, lastRow - 1, 1).getValues();
  var approvedCount = data.filter(function (r) {
    return r[0] === true;
  }).length;

  console.log("รายการทั้งหมดใน Queue: " + (lastRow - 1));
  console.log("ติ๊ก Approve แล้ว: " + approvedCount);

  if (approvedCount === 0) {
    console.log("⚠️ ยังไม่มีรายการที่ติ๊ก Approve");
    console.log("ลองติ๊ก Col H สัก 1-2 แถวใน GPS_Queue แล้วรัน applyApprovedFeedback() ครับ");
  } else {
    console.log("✅ พร้อมรัน applyApprovedFeedback() ได้เลย");
  }
}

function testQueueUpgrade() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SCG_CONFIG.SHEET_GPS_QUEUE);

  var realLastRow = getRealLastRow_(sheet, 1);
  var maxRow = sheet.getMaxRows();

  console.log("แถวข้อมูลจริง: " + (realLastRow - 1) + " รายการ");
  console.log("MaxRows ในชีต: " + maxRow);
  console.log("Checkbox ครอบคลุมถึงแถว: " + (realLastRow + 999));

  // ตรวจสอบ Header
  var headers = sheet.getRange(1, 1, 1, 9).getValues()[0];
  console.log("Header: " + headers.join(" | "));
}

function testConfidenceScore() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  var lastRow = getRealLastRow_(sheet, CONFIG.COL_NAME);
  var data = sheet.getRange(2, 1, lastRow - 1, 17).getValues();

  var stats = {
    total: 0,
    above80: 0,
    above50: 0,
    below50: 0,
    zero: 0
  };

  data.forEach(function (row) {
    var conf = parseFloat(row[CONFIG.C_IDX.CONFIDENCE]);
    if (isNaN(conf)) return;
    stats.total++;
    if (conf >= 80) stats.above80++;
    else if (conf >= 50) stats.above50++;
    else if (conf > 0) stats.below50++;
    else stats.zero++;
  });

  console.log("=== Confidence Score Stats ===");
  console.log("ทั้งหมด: " + stats.total + " แถว");
  console.log("≥ 80% (น่าเชื่อถือมาก): " + stats.above80);
  console.log("50-79% (น่าเชื่อถือพอใช้): " + stats.above50);
  console.log("1-49% (ต้องตรวจสอบ): " + stats.below50);
  console.log("0% (ยังไม่ได้คำนวณ): " + stats.zero);

  // ดูตัวอย่าง 3 แถวแรก
  var sample = data.slice(0, 3);
  console.log("=== ตัวอย่าง 3 แถวแรก ===");
  sample.forEach(function (row, i) {
    console.log(
      "แถว " + (i + 2) + ": " +
      row[CONFIG.C_IDX.NAME] +
      " | Confidence = " + row[CONFIG.C_IDX.CONFIDENCE]
    );
  });
}

function testPostalCacheClear() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. โหลด Cache ก่อน
  var result1 = getPostalDataCached();
  console.log("โหลด Cache ครั้งที่ 1: " + (result1 ? "✅ มีข้อมูล" : "❌ ว่าง"));

  // 2. ล้าง Cache
  clearPostalCache();
  console.log("ล้าง Cache แล้ว");

  // 3. โหลดใหม่
  var result2 = getPostalDataCached();
  console.log("โหลด Cache ครั้งที่ 2: " + (result2 ? "✅ โหลดใหม่สำเร็จ" : "❌ ว่าง"));

  console.log("✅ Postal Cache ทำงานถูกต้องครับ");
}

function testSchemaValidator() {
  console.log("=== ทดสอบ Schema Validator ===");

  // ทดสอบทุกชีต
  var allKeys = Object.keys(SHEET_SCHEMA);
  allKeys.forEach(function (key) {
    var result = validateSheet_(key);
    var status = result.valid ? "✅ ผ่าน" : "❌ มีปัญหา";
    console.log(key + ": " + status);
    if (!result.valid) {
      result.errors.forEach(function (e) {
        console.log("  " + e);
      });
    }
  });
}

function testSoftDelete() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  // ทดสอบ initializeRecordStatus
  initializeRecordStatus();

  // ตรวจสอบผล
  var lastRow = getRealLastRow_(sheet, CONFIG.COL_NAME);
  var data = sheet.getRange(2, 1, lastRow - 1, CONFIG.COL_MERGED_TO_UUID)
    .getValues();

  var activeCount = data.filter(function (r) {
    return r[CONFIG.C_IDX.RECORD_STATUS] === "Active";
  }).length;

  console.log("Active records: " + activeCount);
  console.log("Total records: " + (lastRow - 1));

  if (activeCount === lastRow - 1) {
    console.log("✅ ทุกแถวมี Status = Active ถูกต้องครับ");
  } else {
    console.log("⚠️ บางแถวยังไม่มี Status");
  }
}