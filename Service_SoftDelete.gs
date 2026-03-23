/**
 * VERSION: 4.2 — Phase A
 * [Phase A] เพิ่ม resolveRowUUIDOrNull_() และ isActiveUUID_()
 * เพื่อให้ flows อื่นเรียกใช้ตรวจสอบ UUID ก่อน consume
 */

// ==========================================
// 1. INITIALIZE STATUS
// ==========================================

function initializeRecordStatus() {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var ui      = SpreadsheetApp.getUi();
  var sheet   = ss.getSheetByName(CONFIG.SHEET_NAME);
  var lastRow = getRealLastRow_(sheet, CONFIG.COL_NAME);
  if (lastRow < 2) return;

  var maxCol = CONFIG.DB_TOTAL_COLS;
  var data   = sheet.getRange(2, 1, lastRow - 1, maxCol).getValues();
  var count  = 0;

  data.forEach(function(row, i) {
    if (!row[CONFIG.C_IDX.NAME]) return;
    if (!row[CONFIG.C_IDX.RECORD_STATUS]) {
      data[i][CONFIG.C_IDX.RECORD_STATUS] = "Active";
      count++;
    }
  });

  if (count > 0) {
    sheet.getRange(2, 1, data.length, maxCol).setValues(data);
    SpreadsheetApp.flush();
  }

  ui.alert(
    "✅ Initialize สำเร็จ!\n\n" +
    "ตั้งค่า Record_Status = Active: " + count + " แถว\n\n" +
    "สถานะในระบบ:\n" +
    "Active   = ใช้งานปกติ\n" +
    "Inactive = ปิดการใช้งาน\n" +
    "Merged   = รวมเข้ากับ UUID อื่นแล้ว"
  );
}

// ==========================================
// 2. SOFT DELETE
// ==========================================

function softDeleteRecord(uuid) {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var sheet   = ss.getSheetByName(CONFIG.SHEET_NAME);
  var lastRow = getRealLastRow_(sheet, CONFIG.COL_NAME);
  var data    = sheet.getRange(2, 1, lastRow - 1, CONFIG.DB_TOTAL_COLS).getValues();

  for (var i = 0; i < data.length; i++) {
    if (data[i][CONFIG.C_IDX.UUID] === uuid) {
      var rowNum = i + 2;
      sheet.getRange(rowNum, CONFIG.COL_RECORD_STATUS).setValue("Inactive");
      sheet.getRange(rowNum, CONFIG.COL_UPDATED).setValue(new Date());
      console.log("[softDeleteRecord] UUID " + uuid + " → Inactive");
      return true;
    }
  }
  return false;
}

// ==========================================
// 3. MERGE UUIDs
// ==========================================

function mergeUUIDs(masterUUID, duplicateUUID) {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var sheet   = ss.getSheetByName(CONFIG.SHEET_NAME);
  var lastRow = getRealLastRow_(sheet, CONFIG.COL_NAME);
  var data    = sheet.getRange(2, 1, lastRow - 1, CONFIG.DB_TOTAL_COLS).getValues();

  var masterFound = false, duplicateFound = false;

  for (var i = 0; i < data.length; i++) {
    var rowUUID = data[i][CONFIG.C_IDX.UUID];
    var rowNum  = i + 2;
    if (rowUUID === masterUUID)    masterFound = true;
    if (rowUUID === duplicateUUID) {
      sheet.getRange(rowNum, CONFIG.COL_RECORD_STATUS).setValue("Merged");
      sheet.getRange(rowNum, CONFIG.COL_MERGED_TO_UUID).setValue(masterUUID);
      sheet.getRange(rowNum, CONFIG.COL_UPDATED).setValue(new Date());
      duplicateFound = true;
      console.log("[mergeUUIDs] " + duplicateUUID + " → " + masterUUID);
    }
  }
  return { masterFound: masterFound, duplicateFound: duplicateFound };
}

// ==========================================
// 4. RESOLVE UUID (ติดตาม Merge chain)
// ==========================================

function resolveUUID(uuid) {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var sheet   = ss.getSheetByName(CONFIG.SHEET_NAME);
  var lastRow = getRealLastRow_(sheet, CONFIG.COL_NAME);
  var data    = sheet.getRange(2, 1, lastRow - 1, CONFIG.DB_TOTAL_COLS).getValues();

  var uuidMap = {};
  data.forEach(function(row) {
    var u = row[CONFIG.C_IDX.UUID];
    if (u) uuidMap[u] = {
      status:   row[CONFIG.C_IDX.RECORD_STATUS],
      mergedTo: row[CONFIG.C_IDX.MERGED_TO_UUID]
    };
  });

  var current  = uuid;
  var maxHops  = 10;
  var hopCount = 0;
  while (hopCount < maxHops) {
    var info = uuidMap[current];
    if (!info || info.status !== "Merged" || !info.mergedTo) break;
    current = info.mergedTo;
    hopCount++;
  }
  return current;
}

// ==========================================
// 5. [Phase A NEW] HELPER FUNCTIONS
// ==========================================

/**
 * resolveRowUUIDOrNull_()
 * ใช้ก่อน consume UUID จาก mapping หรือ search
 * คืน canonical UUID ถ้า active, คืน null ถ้า inactive หรือหาไม่เจอ
 */
function resolveRowUUIDOrNull_(uuid) {
  if (!uuid) return null;
  var resolved = resolveUUID(uuid);
  if (!resolved) return null;
  if (!isActiveUUID_(resolved)) {
    console.warn("[resolveRowUUIDOrNull_] UUID '" + resolved + "' ไม่ active");
    return null;
  }
  return resolved;
}

/**
 * isActiveUUID_()
 * ตรวจว่า UUID นี้ยัง active ใช้งานได้อยู่หรือไม่
 */
function isActiveUUID_(uuid) {
  if (!uuid) return false;
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var sheet   = ss.getSheetByName(CONFIG.SHEET_NAME);
  var lastRow = getRealLastRow_(sheet, CONFIG.COL_NAME);
  var data    = sheet.getRange(2, 1, lastRow - 1, CONFIG.DB_TOTAL_COLS).getValues();

  for (var i = 0; i < data.length; i++) {
    if (data[i][CONFIG.C_IDX.UUID] === uuid) {
      var status = data[i][CONFIG.C_IDX.RECORD_STATUS];
      return (status === "Active" || status === "");
    }
  }
  return false;
}

// ==========================================
// 6. UI
// ==========================================

function mergeDuplicates_UI() {
  var ui = SpreadsheetApp.getUi();

  var resMaster = ui.prompt("🔀 Merge UUID (1/2)", "กรุณาใส่ Master UUID:", ui.ButtonSet.OK_CANCEL);
  if (resMaster.getSelectedButton() !== ui.Button.OK) return;
  var masterUUID = resMaster.getResponseText().trim();

  var resDup = ui.prompt("🔀 Merge UUID (2/2)", "กรุณาใส่ Duplicate UUID:", ui.ButtonSet.OK_CANCEL);
  if (resDup.getSelectedButton() !== ui.Button.OK) return;
  var duplicateUUID = resDup.getResponseText().trim();

  if (!masterUUID || !duplicateUUID) { ui.alert("❌ UUID ไม่ครบ"); return; }
  if (masterUUID === duplicateUUID)  { ui.alert("❌ UUID เดียวกัน"); return; }

  var result = mergeUUIDs(masterUUID, duplicateUUID);
  if (!result.masterFound)    { ui.alert("❌ ไม่พบ Master UUID"); return; }
  if (!result.duplicateFound) { ui.alert("❌ ไม่พบ Duplicate UUID"); return; }

  if (typeof clearSearchCache === 'function') clearSearchCache();

  ui.alert(
    "✅ Merge สำเร็จ!\n\n" +
    "Master: "    + masterUUID    + "\n" +
    "Duplicate: " + duplicateUUID + "\n\n" +
    "Duplicate ถูก Mark เป็น 'Merged' แล้วครับ\nข้อมูลเดิมยังอยู่ครบ"
  );
}

function showRecordStatusReport() {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var ui      = SpreadsheetApp.getUi();
  var sheet   = ss.getSheetByName(CONFIG.SHEET_NAME);
  var lastRow = getRealLastRow_(sheet, CONFIG.COL_NAME);
  if (lastRow < 2) { ui.alert("ℹ️ Database ว่างเปล่าครับ"); return; }

  var data  = sheet.getRange(2, 1, lastRow - 1, CONFIG.DB_TOTAL_COLS).getValues();
  var stats = { active: 0, inactive: 0, merged: 0, noStatus: 0 };

  data.forEach(function(row) {
    if (!row[CONFIG.C_IDX.NAME]) return;
    var s = row[CONFIG.C_IDX.RECORD_STATUS];
    if (s === "Active")        stats.active++;
    else if (s === "Inactive") stats.inactive++;
    else if (s === "Merged")   stats.merged++;
    else                       stats.noStatus++;
  });

  ui.alert(
    "📊 Record Status Report\n━━━━━━━━━━━━━━━━━━━━━━━\n" +
    "✅ Active:    " + stats.active   + " แถว\n" +
    "⚫ Inactive:  " + stats.inactive + " แถว\n" +
    "🔀 Merged:    " + stats.merged   + " แถว\n" +
    "❓ ไม่มีสถานะ: " + stats.noStatus + " แถว\n" +
    "━━━━━━━━━━━━━━━━━━━━━━━\n" +
    "รวม: " + (stats.active + stats.inactive + stats.merged + stats.noStatus) + " แถว"
  );
}

/**
 * [Phase C NEW] buildUUIDStateMap_()
 * โหลด UUID state ทั้งหมดจาก Database ครั้งเดียว
 * ใช้ใน resolveUUID(), isActiveUUID_() เพื่อลด Sheets API calls
 * เรียกจาก flow ที่ต้องตรวจ UUID หลายตัวพร้อมกัน
 */
function buildUUIDStateMap_() {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var sheet   = ss.getSheetByName(CONFIG.SHEET_NAME);
  var lastRow = getRealLastRow_(sheet, CONFIG.COL_NAME);
  var map     = {};
  if (lastRow < 2) return map;

  var data = sheet.getRange(2, 1, lastRow - 1, CONFIG.DB_TOTAL_COLS).getValues();
  data.forEach(function(row) {
    var u = row[CONFIG.C_IDX.UUID];
    if (u) {
      map[u] = {
        status:   row[CONFIG.C_IDX.RECORD_STATUS]  || "Active",
        mergedTo: row[CONFIG.C_IDX.MERGED_TO_UUID] || ""
      };
    }
  });
  return map;
}

/**
 * [Phase C NEW] resolveUUIDFromMap_()
 * เหมือน resolveUUID() แต่รับ stateMap ที่โหลดไว้แล้ว
 * ใช้เมื่อต้อง resolve UUID หลายตัวในรอบเดียวกัน ไม่ต้องอ่าน Sheet ซ้ำ
 */
function resolveUUIDFromMap_(uuid, stateMap) {
  if (!uuid || !stateMap) return uuid;
  var current  = uuid;
  var maxHops  = 10;
  var hopCount = 0;

  while (hopCount < maxHops) {
    var info = stateMap[current];
    if (!info || info.status !== "Merged" || !info.mergedTo) break;
    current = info.mergedTo;
    hopCount++;
  }
  return current;
}

/**
 * [Phase C NEW] isActiveFromMap_()
 * ตรวจสถานะจาก stateMap ที่โหลดไว้แล้ว
 */
function isActiveFromMap_(uuid, stateMap) {
  if (!uuid || !stateMap) return false;
  var info = stateMap[uuid];
  if (!info) return false;
  return (info.status === "Active" || info.status === "");
}
