/**
 * 🗂️ Service: Soft Delete & UUID Merging
 * Version: 4.1 Enterprise Edition
 * หน้าที่: จัดการการรวม UUID และ Soft Delete
 * ไม่ลบข้อมูลจริง แต่ Mark สถานะแทน
 */

// ==========================================
// 1. INITIALIZE STATUS (รันครั้งแรกครั้งเดียว)
// ==========================================

/**
 * ตั้งค่า Record_Status = "Active" ให้ทุกแถวที่ยังว่างอยู่
 */
function initializeRecordStatus() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  var lastRow = getRealLastRow_(sheet, CONFIG.COL_NAME);
  if (lastRow < 2) return;

  var maxCol = Math.max(CONFIG.COL_RECORD_STATUS, CONFIG.COL_MERGED_TO_UUID);
  var data = sheet.getRange(2, 1, lastRow - 1, maxCol).getValues();
  var count = 0;

  data.forEach(function (row, i) {
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
    "สถานะที่ใช้ในระบบ:\n" +
    "Active   = ใช้งานปกติ\n" +
    "Inactive = ปิดการใช้งาน\n" +
    "Merged   = รวมเข้ากับ UUID อื่นแล้ว"
  );
}

// ==========================================
// 2. SOFT DELETE (แทนการลบจริง)
// ==========================================

/**
 * เปลี่ยน Status เป็น Inactive แทนการลบ
 */
function softDeleteRecord(uuid) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  var lastRow = getRealLastRow_(sheet, CONFIG.COL_NAME);
  var data = sheet.getRange(2, 1, lastRow - 1, CONFIG.COL_MERGED_TO_UUID)
    .getValues();

  for (var i = 0; i < data.length; i++) {
    if (data[i][CONFIG.C_IDX.UUID] === uuid) {
      var rowNum = i + 2;
      sheet.getRange(rowNum, CONFIG.COL_RECORD_STATUS)
        .setValue("Inactive");
      sheet.getRange(rowNum, CONFIG.COL_UPDATED)
        .setValue(new Date());
      console.log("Soft Delete: UUID " + uuid + " → Inactive");
      return true;
    }
  }
  return false;
}

// ==========================================
// 3. MERGE UUIDs (รวม 2 UUID เป็น 1)
// ==========================================

/**
 * รวม duplicateUUID เข้ากับ masterUUID
 * - duplicateUUID จะถูก Mark เป็น Merged
 * - ชี้ MERGED_TO_UUID → masterUUID
 * - ไม่ลบข้อมูลจริง
 */
function mergeUUIDs(masterUUID, duplicateUUID) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  var lastRow = getRealLastRow_(sheet, CONFIG.COL_NAME);
  var data = sheet.getRange(2, 1, lastRow - 1, CONFIG.COL_MERGED_TO_UUID)
    .getValues();

  var masterFound = false;
  var duplicateFound = false;

  for (var i = 0; i < data.length; i++) {
    var rowUUID = data[i][CONFIG.C_IDX.UUID];
    var rowNum = i + 2;

    if (rowUUID === masterUUID) {
      masterFound = true;
    }

    if (rowUUID === duplicateUUID) {
      // Mark เป็น Merged และชี้ไป masterUUID
      sheet.getRange(rowNum, CONFIG.COL_RECORD_STATUS)
        .setValue("Merged");
      sheet.getRange(rowNum, CONFIG.COL_MERGED_TO_UUID)
        .setValue(masterUUID);
      sheet.getRange(rowNum, CONFIG.COL_UPDATED)
        .setValue(new Date());

      duplicateFound = true;
      console.log("Merged: " + duplicateUUID + " → " + masterUUID);
    }
  }

  return { masterFound: masterFound, duplicateFound: duplicateFound };
}

// ==========================================
// 4. RESOLVE UUID (ติดตาม Merge chain)
// ==========================================

/**
 * ถ้า UUID ถูก Merge แล้ว ให้ไปดึง UUID ปลายทางแทน
 * รองรับ chain: A → B → C (คืนค่า C)
 */
function resolveUUID(uuid) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  var lastRow = getRealLastRow_(sheet, CONFIG.COL_NAME);
  var data = sheet.getRange(2, 1, lastRow - 1, CONFIG.COL_MERGED_TO_UUID)
    .getValues();

  // สร้าง map: uuid → {status, mergedTo}
  var uuidMap = {};
  data.forEach(function (row) {
    var u = row[CONFIG.C_IDX.UUID];
    if (u) {
      uuidMap[u] = {
        status: row[CONFIG.C_IDX.RECORD_STATUS],
        mergedTo: row[CONFIG.C_IDX.MERGED_TO_UUID]
      };
    }
  });

  // ติดตาม chain (ป้องกัน infinite loop)
  var current = uuid;
  var maxHops = 10;
  var hopCount = 0;

  while (hopCount < maxHops) {
    var info = uuidMap[current];
    if (!info) break;
    if (info.status !== "Merged" || !info.mergedTo) break;
    current = info.mergedTo;
    hopCount++;
  }

  return current;
}

// ==========================================
// 5. UI — Merge จาก Dialog
// ==========================================

function mergeDuplicates_UI() {
  var ui = SpreadsheetApp.getUi();

  var resMaster = ui.prompt(
    "🔀 Merge UUID (ขั้นที่ 1/2)",
    "กรุณาใส่ Master UUID (UUID ที่จะเก็บไว้):",
    ui.ButtonSet.OK_CANCEL
  );
  if (resMaster.getSelectedButton() !== ui.Button.OK) return;
  var masterUUID = resMaster.getResponseText().trim();

  var resDup = ui.prompt(
    "🔀 Merge UUID (ขั้นที่ 2/2)",
    "กรุณาใส่ Duplicate UUID (UUID ที่จะรวมเข้า Master):",
    ui.ButtonSet.OK_CANCEL
  );
  if (resDup.getSelectedButton() !== ui.Button.OK) return;
  var duplicateUUID = resDup.getResponseText().trim();

  if (!masterUUID || !duplicateUUID) {
    ui.alert("❌ UUID ไม่ครบ กรุณาลองใหม่ครับ");
    return;
  }

  if (masterUUID === duplicateUUID) {
    ui.alert("❌ Master และ Duplicate เป็น UUID เดียวกัน");
    return;
  }

  var result = mergeUUIDs(masterUUID, duplicateUUID);

  if (!result.masterFound) {
    ui.alert("❌ ไม่พบ Master UUID: " + masterUUID);
    return;
  }
  if (!result.duplicateFound) {
    ui.alert("❌ ไม่พบ Duplicate UUID: " + duplicateUUID);
    return;
  }

  // ล้าง Search Cache เพราะข้อมูลเปลี่ยน
  if (typeof clearSearchCache === 'function') clearSearchCache();

  ui.alert(
    "✅ Merge สำเร็จ!\n\n" +
    "Master UUID:    " + masterUUID + "\n" +
    "Duplicate UUID: " + duplicateUUID + "\n\n" +
    "Duplicate ถูก Mark เป็น 'Merged'\n" +
    "และชี้ไปที่ Master UUID แล้วครับ\n\n" +
    "ข้อมูลเดิมยังอยู่ครบ ไม่มีอะไรถูกลบ"
  );
}

// ==========================================
// 6. REPORT — สรุปสถานะ
// ==========================================

function showRecordStatusReport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  var lastRow = getRealLastRow_(sheet, CONFIG.COL_NAME);
  if (lastRow < 2) {
    ui.alert("ℹ️ Database ว่างเปล่าครับ");
    return;
  }

  var data = sheet.getRange(2, 1, lastRow - 1, CONFIG.COL_MERGED_TO_UUID)
    .getValues();

  var stats = {
    active: 0,
    inactive: 0,
    merged: 0,
    noStatus: 0
  };

  data.forEach(function (row) {
    if (!row[CONFIG.C_IDX.NAME]) return;
    var status = row[CONFIG.C_IDX.RECORD_STATUS];
    if (status === "Active") stats.active++;
    else if (status === "Inactive") stats.inactive++;
    else if (status === "Merged") stats.merged++;
    else stats.noStatus++;
  });

  ui.alert(
    "📊 Record Status Report\n" +
    "━━━━━━━━━━━━━━━━━━━━━━━\n" +
    "✅ Active:   " + stats.active + " แถว\n" +
    "⚫ Inactive: " + stats.inactive + " แถว\n" +
    "🔀 Merged:   " + stats.merged + " แถว\n" +
    "❓ ไม่มีสถานะ: " + stats.noStatus + " แถว\n" +
    "━━━━━━━━━━━━━━━━━━━━━━━\n" +
    "รวม: " + (stats.active + stats.inactive + stats.merged + stats.noStatus) + " แถว"
  );
}