function createGPSQueueSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  // เช็คว่ามีอยู่แล้วไหม
  if (ss.getSheetByName(SCG_CONFIG.SHEET_GPS_QUEUE)) {
    ui.alert("ℹ️ ชีต GPS_Queue มีอยู่แล้วครับ");
    return;
  }

  // สร้างชีตใหม่
  var sheet = ss.insertSheet(SCG_CONFIG.SHEET_GPS_QUEUE);

  // สร้าง Header
  var headers = [
    "Timestamp",       // A
    "ShipToName",      // B
    "UUID_DB",         // C
    "LatLng_Driver",   // D
    "LatLng_DB",       // E
    "Diff_Meters",     // F
    "Reason",          // G
    "Approve",         // H
    "Reject"           // I
  ];

  // ตั้งค่า Header
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight("bold");
  headerRange.setBackground("#4f46e5");
  headerRange.setFontColor("white");

  // เพิ่ม Checkbox ใน Col H และ I สำหรับ 500 แถว
  sheet.getRange(2, 8, 500, 1).insertCheckboxes();
  sheet.getRange(2, 9, 500, 1).insertCheckboxes();

  // ปรับความกว้างคอลัมน์
  sheet.setColumnWidth(1, 160);  // Timestamp
  sheet.setColumnWidth(2, 250);  // ShipToName
  sheet.setColumnWidth(3, 280);  // UUID_DB
  sheet.setColumnWidth(4, 160);  // LatLng_Driver
  sheet.setColumnWidth(5, 160);  // LatLng_DB
  sheet.setColumnWidth(6, 100);  // Diff_Meters
  sheet.setColumnWidth(7, 120);  // Reason
  sheet.setColumnWidth(8, 80);   // Approve
  sheet.setColumnWidth(9, 80);   // Reject

  // Freeze header
  sheet.setFrozenRows(1);

  SpreadsheetApp.flush();

  ui.alert("✅ สร้างชีต GPS_Queue สำเร็จแล้วครับ\nพร้อมใช้งาน");
}

function resetSyncStatus() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var sheet = ss.getSheetByName(CONFIG.SOURCE_SHEET);

  var lastRow = sheet.getLastRow();
  var syncCol = SCG_CONFIG.SRC_IDX_SYNC_STATUS;

  var result = ui.alert(
    "⚠️ Reset SYNC_STATUS?",
    "จะล้าง SYNCED ทั้งหมด " + (lastRow - 1) + " แถว\n" +
    "เพื่อให้ระบบประมวลผลใหม่ทั้งหมด\n\n" +
    "ใช้สำหรับทดสอบเท่านั้น ยืนยันหรือไม่?",
    ui.ButtonSet.YES_NO
  );

  if (result !== ui.Button.YES) return;

  sheet.getRange(2, syncCol, lastRow - 1, 1).clearContent();
  SpreadsheetApp.flush();

  console.log("✅ Reset เรียบร้อย พร้อมรัน syncNewDataToMaster() ใหม่");
  ui.alert("✅ Reset เรียบร้อยแล้วครับ");
}

function applyApprovedFeedback() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var queueSheet = ss.getSheetByName(SCG_CONFIG.SHEET_GPS_QUEUE);
  var masterSheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!queueSheet || !masterSheet) {
    ui.alert("❌ ไม่พบชีต GPS_Queue หรือ Database");
    return;
  }

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    ui.alert("⚠️ ระบบคิวทำงาน", "มีผู้ใช้งานอื่นกำลังใช้งานอยู่ กรุณารอสักครู่", ui.ButtonSet.OK);
    return;
  }

  // [NEW v4.1] ตรวจสอบ Schema ก่อนทำงาน
  try { preCheck_Approve(); } catch (e) {
    ui.alert("❌ Schema Error", e.message, ui.ButtonSet.OK);
    return;
  }

  try {
    // --- อ่าน GPS_Queue ---
    var lastQueueRow = getRealLastRow_(queueSheet, 1);
    if (lastQueueRow < 2) {
      ui.alert("ℹ️ ไม่มีรายการใน GPS_Queue");
      return;
    }

    var queueData = queueSheet.getRange(2, 1, lastQueueRow - 1, 9).getValues();

    // --- โหลด Database เข้า Memory ---
    var lastRowM = getRealLastRow_(masterSheet, CONFIG.COL_NAME);
    var dbData = masterSheet.getRange(2, 1, lastRowM - 1, 20).getValues();

    // สร้าง UUID → rowIndex map
    var uuidMap = {};
    dbData.forEach(function (r, i) {
      if (r[CONFIG.C_IDX.UUID]) {
        uuidMap[r[CONFIG.C_IDX.UUID]] = i;
      }
    });

    // --- ประมวลผลแถวที่ Approve = true ---
    var approvedCount = 0;
    var skippedCount = 0;
    var ts = new Date();

    queueData.forEach(function (row, i) {
      var isApproved = row[7]; // Col H: Approve
      var isRejected = row[8]; // Col I: Reject
      var reason = row[6]; // Col G: Reason

      // ข้ามแถวที่ไม่ได้ติ๊ก หรือ Reject แล้ว
      if (!isApproved || isRejected) {
        skippedCount++;
        return;
      }

      // ข้ามแถวที่ดำเนินการแล้ว
      if (reason === "APPROVED" || reason === "REJECTED") {
        return;
      }

      var uuid = row[2]; // Col C: UUID_DB
      var latLngDriver = row[3]; // Col D: LatLng_Driver

      if (!uuid || !latLngDriver) {
        skippedCount++;
        return;
      }

      // แปลง "lat, lng" → ตัวเลข
      var parts = latLngDriver.toString().split(",");
      if (parts.length !== 2) {
        skippedCount++;
        return;
      }

      var newLat = parseFloat(parts[0].trim());
      var newLng = parseFloat(parts[1].trim());

      if (isNaN(newLat) || isNaN(newLng)) {
        skippedCount++;
        return;
      }

      // หาแถวใน Database
      if (!uuidMap.hasOwnProperty(uuid)) {
        skippedCount++;
        console.warn("UUID ไม่พบใน Database: " + uuid);
        return;
      }

      var dbRowIndex = uuidMap[uuid];
      var dbRowNum = dbRowIndex + 2; // +1 header +1 เพราะ 0-based

      // อัปเดต Database
      masterSheet.getRange(dbRowNum, CONFIG.COL_LAT)
        .setValue(newLat);
      masterSheet.getRange(dbRowNum, CONFIG.COL_LNG)
        .setValue(newLng);
      masterSheet.getRange(dbRowNum, CONFIG.COL_COORD_SOURCE)
        .setValue("Driver_GPS");
      masterSheet.getRange(dbRowNum, CONFIG.COL_COORD_CONFIDENCE)
        .setValue(95);
      masterSheet.getRange(dbRowNum, CONFIG.COL_COORD_LAST_UPDATED)
        .setValue(ts);
      masterSheet.getRange(dbRowNum, CONFIG.COL_UPDATED)
        .setValue(ts);

      // Mark Queue ว่า APPROVED
      queueSheet.getRange(i + 2, 7).setValue("APPROVED");

      approvedCount++;
      console.log("✅ Approved: " + row[1] + " | UUID: " + uuid);
    });

    // ล้าง Search Cache เพราะพิกัดเปลี่ยน
    if (typeof clearSearchCache === 'function') {
      clearSearchCache();
    }

    SpreadsheetApp.flush();

    var msg = "✅ อนุมัติเรียบร้อย!\n\n" +
      "📍 อัปเดตพิกัดใน Database: " + approvedCount + " ราย\n" +
      "⏭️ ข้ามไป: " + skippedCount + " ราย\n\n";

    if (approvedCount > 0) {
      msg += "Database ได้รับการอัปเดตพิกัดจาก Driver GPS แล้วครับ";
    } else {
      msg += "ไม่มีรายการที่ติ๊ก Approve\nกรุณาติ๊ก Col H ก่อนรันครับ";
    }

    ui.alert(msg);

  } catch (e) {
    console.error("applyApprovedFeedback Error: " + e.message);
    ui.alert("❌ เกิดข้อผิดพลาด: " + e.message);
  } finally {
    lock.releaseLock();
  }
}


function showGPSQueueStats() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var queueSheet = ss.getSheetByName(SCG_CONFIG.SHEET_GPS_QUEUE);

  if (!queueSheet) {
    ui.alert("❌ ไม่พบชีต GPS_Queue\nกรุณารัน 'สร้างชีต GPS_Queue ใหม่' ก่อนครับ");
    return;
  }

  var lastRow = getRealLastRow_(queueSheet, 1);
  if (lastRow < 2) {
    ui.alert("ℹ️ GPS_Queue ว่างเปล่า ยังไม่มีรายการครับ");
    return;
  }

  var data = queueSheet.getRange(2, 1, lastRow - 1, 9).getValues();

  var stats = {
    total: 0,
    pending: 0,
    approved: 0,
    rejected: 0,
    gpsDiff: 0,
    noGPS: 0
  };

  data.forEach(function (row) {
    var reason = row[6]; // Col G
    var approve = row[7]; // Col H
    var reject = row[8]; // Col I

    // นับเฉพาะแถวที่มีข้อมูลจริง
    if (!row[0]) return;
    stats.total++;

    if (reason === "APPROVED") stats.approved++;
    else if (reason === "REJECTED") stats.rejected++;
    else if (approve) stats.approved++;
    else if (reject) stats.rejected++;
    else stats.pending++;

    if (reason === "GPS_DIFF") stats.gpsDiff++;
    else if (reason === "DB_NO_GPS") stats.noGPS++;
  });

  var msg =
    "📊 GPS Queue สถิติ\n" +
    "━━━━━━━━━━━━━━━━━━━━━━━\n" +
    "📋 รายการทั้งหมด: " + stats.total + " ราย\n" +
    "⏳ รอตรวจสอบ: " + stats.pending + " ราย\n" +
    "✅ อนุมัติแล้ว: " + stats.approved + " ราย\n" +
    "❌ ปฏิเสธแล้ว: " + stats.rejected + " ราย\n" +
    "━━━━━━━━━━━━━━━━━━━━━━━\n" +
    "📍 พิกัดต่างกัน > 50m: " + stats.gpsDiff + " ราย\n" +
    "🔍 DB ไม่มีพิกัด: " + stats.noGPS + " ราย";

  ui.alert(msg);
}

function upgradeGPSQueueSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var sheet = ss.getSheetByName(SCG_CONFIG.SHEET_GPS_QUEUE);

  if (!sheet) {
    ui.alert("❌ ไม่พบชีต GPS_Queue ครับ");
    return;
  }

  // 1. ตรวจสอบ Header
  var headers = [
    "Timestamp",
    "ShipToName",
    "UUID_DB",
    "LatLng_Driver",
    "LatLng_DB",
    "Diff_Meters",
    "Reason",
    "Approve",
    "Reject"
  ];

  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight("bold");
  headerRange.setBackground("#4f46e5");
  headerRange.setFontColor("white");
  headerRange.setFontSize(11);

  // 2. หาแถวข้อมูลจริง
  var realLastRow = getRealLastRow_(sheet, 1);
  console.log("แถวข้อมูลจริง: " + realLastRow);

  // 3. ลบ Checkbox เก่าออกทั้งหมดก่อน
  var maxRow = sheet.getMaxRows();
  if (maxRow > 1) {
    sheet.getRange(2, 8, maxRow - 1, 2).clearContent();
  }

  // 4. เพิ่ม Checkbox ใหม่ให้ครอบคลุมข้อมูล + เผื่อไว้ 1000 แถว
  var checkboxEnd = Math.max(realLastRow, 1) + 1000;
  sheet.getRange(2, 8, checkboxEnd, 1).insertCheckboxes(); // Approve
  sheet.getRange(2, 9, checkboxEnd, 1).insertCheckboxes(); // Reject

  // 5. Format คอลัมน์
  sheet.setColumnWidth(1, 160);  // Timestamp
  sheet.setColumnWidth(2, 250);  // ShipToName
  sheet.setColumnWidth(3, 280);  // UUID_DB
  sheet.setColumnWidth(4, 160);  // LatLng_Driver
  sheet.setColumnWidth(5, 160);  // LatLng_DB
  sheet.setColumnWidth(6, 110);  // Diff_Meters
  sheet.setColumnWidth(7, 120);  // Reason
  sheet.setColumnWidth(8, 80);   // Approve
  sheet.setColumnWidth(9, 80);   // Reject

  // 6. Format Diff_Meters ให้แสดงเป็นตัวเลข
  if (realLastRow > 1) {
    sheet.getRange(2, 6, realLastRow - 1, 1)
      .setNumberFormat("#,##0");
  }

  // 7. Freeze Header
  sheet.setFrozenRows(1);

  // 8. Color แถวตาม Reason
  if (realLastRow > 1) {
    var reasonData = sheet.getRange(2, 7, realLastRow - 1, 1).getValues();
    reasonData.forEach(function (row, i) {
      var rowNum = i + 2;
      var reason = row[0];
      var bg = "#ffffff";

      if (reason === "GPS_DIFF") bg = "#fff3cd"; // เหลือง
      if (reason === "DB_NO_GPS") bg = "#f8d7da"; // แดง
      if (reason === "NO_MATCH") bg = "#d1ecf1"; // ฟ้า
      if (reason === "APPROVED") bg = "#d4edda"; // เขียว
      if (reason === "REJECTED") bg = "#e2e3e5"; // เทา

      sheet.getRange(rowNum, 1, 1, 9).setBackground(bg);
    });
  }

  SpreadsheetApp.flush();

  ui.alert(
    "✅ อัปเกรด GPS_Queue สำเร็จ!\n\n" +
    "📋 แถวข้อมูลจริง: " + (realLastRow - 1) + " รายการ\n" +
    "☑️ Checkbox: ครอบคลุมถึงแถว " + (realLastRow + 999) + "\n\n" +
    "สีในชีต:\n" +
    "🟡 เหลือง = GPS ต่างกัน > 50m\n" +
    "🔴 แดง = DB ไม่มีพิกัด\n" +
    "🔵 ฟ้า = Match ชื่อไม่เจอ\n" +
    "🟢 เขียว = อนุมัติแล้ว\n" +
    "⚫ เทา = ปฏิเสธแล้ว"
  );
}