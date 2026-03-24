function createGPSQueueSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  
  //  เช็คว่ามีอยู่แล้วไหม
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

/**
 * [Phase A FIXED] applyApprovedFeedback()
 * เพิ่ม REJECTED path, CONFLICT detection, batch write DB
 */
function applyApprovedFeedback() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var ui  = SpreadsheetApp.getUi();

  var queueSheet  = ss.getSheetByName(SCG_CONFIG.SHEET_GPS_QUEUE);
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

  try { preCheck_Approve(); } catch(e) {
    lock.releaseLock();
    ui.alert("❌ Schema Error", e.message, ui.ButtonSet.OK);
    return;
  }

  try {
    console.log("[applyApprovedFeedback] START — Phase A Fixed");

    var lastQueueRow = getRealLastRow_(queueSheet, 1);
    if (lastQueueRow < 2) { ui.alert("ℹ️ ไม่มีรายการใน GPS_Queue"); return; }

    var queueData = queueSheet.getRange(2, 1, lastQueueRow - 1, CONFIG.GPS_QUEUE_TOTAL_COLS).getValues();

    // โหลด DB UUID map
    var lastRowM = getRealLastRow_(masterSheet, CONFIG.COL_NAME);
    var dbData   = masterSheet.getRange(2, 1, lastRowM - 1, CONFIG.DB_TOTAL_COLS).getValues();
    var uuidMap  = {};
    dbData.forEach(function(r, i) {
      if (r[CONFIG.C_IDX.UUID]) uuidMap[r[CONFIG.C_IDX.UUID]] = i;
    });

    // Counters
    var counts = { approved: 0, rejected: 0, conflict: 0, skipped: 0, alreadyDone: 0 };
    var ts     = new Date();

    // รวม updates ไว้เขียน batch
    var dbRowUpdates       = {};
    var queueReasonUpdates = [];

    queueData.forEach(function(row, i) {
      var isApproved = row[7];
      var isRejected = row[8];
      var reason     = (row[6] || "").toString();

      // ข้ามที่ดำเนินการแล้ว
      if (reason === "APPROVED" || reason === "REJECTED" || reason === "CONFLICT") {
        counts.alreadyDone++;
        return;
      }

      // [Phase A NEW] Conflict Detection
      if (isApproved === true && isRejected === true) {
        counts.conflict++;
        queueReasonUpdates.push({ rowNum: i + 2, reason: "CONFLICT" });
        console.warn("[applyApprovedFeedback] CONFLICT แถว " + (i + 2));
        return;
      }

      // [Phase A NEW] Reject Path
      if (isRejected === true && isApproved !== true) {
        counts.rejected++;
        queueReasonUpdates.push({ rowNum: i + 2, reason: "REJECTED" });
        console.log("[applyApprovedFeedback] REJECTED แถว " + (i + 2) + ": " + row[1]);
        return;
      }

      // Approve Path
      if (isApproved !== true) { counts.skipped++; return; }

      var uuid         = (row[2] || "").toString();
      var latLngDriver = (row[3] || "").toString();
      if (!uuid || !latLngDriver) { counts.skipped++; return; }

      var parts  = latLngDriver.split(",");
      if (parts.length !== 2) { counts.skipped++; return; }

      var newLat = parseFloat(parts[0].trim());
      var newLng = parseFloat(parts[1].trim());
      if (isNaN(newLat) || isNaN(newLng)) { counts.skipped++; return; }

      if (!uuidMap.hasOwnProperty(uuid)) {
        console.warn("[applyApprovedFeedback] UUID ไม่พบ: " + uuid);
        counts.skipped++;
        return;
      }

      dbRowUpdates[uuidMap[uuid]] = { lat: newLat, lng: newLng, ts: ts };
      queueReasonUpdates.push({ rowNum: i + 2, reason: "APPROVED" });
      counts.approved++;
    });

    // [Phase A NEW] Batch write DB
    if (Object.keys(dbRowUpdates).length > 0) {
      var fullDb = masterSheet.getRange(2, 1, lastRowM - 1, CONFIG.DB_TOTAL_COLS).getValues();
      Object.keys(dbRowUpdates).forEach(function(idx) {
        var upd = dbRowUpdates[idx];
        var i   = parseInt(idx);
        fullDb[i][CONFIG.C_IDX.LAT]                = upd.lat;
        fullDb[i][CONFIG.C_IDX.LNG]                = upd.lng;
        fullDb[i][CONFIG.C_IDX.COORD_SOURCE]       = "Driver_GPS";
        fullDb[i][CONFIG.C_IDX.COORD_CONFIDENCE]   = 95;
        fullDb[i][CONFIG.C_IDX.COORD_LAST_UPDATED] = upd.ts;
        fullDb[i][CONFIG.C_IDX.UPDATED]            = upd.ts;
      });
      masterSheet.getRange(2, 1, lastRowM - 1, CONFIG.DB_TOTAL_COLS).setValues(fullDb);
      console.log("[applyApprovedFeedback] Batch write: " + Object.keys(dbRowUpdates).length + " rows");
    }

    // Write queue reasons
    queueReasonUpdates.forEach(function(upd) {
      queueSheet.getRange(upd.rowNum, 7).setValue(upd.reason);
    });

    if (typeof clearSearchCache === 'function') clearSearchCache();
    SpreadsheetApp.flush();

    var msg = "✅ ดำเนินการเรียบร้อย!\n\n" +
              "📍 APPROVED: "  + counts.approved   + " ราย\n" +
              "❌ REJECTED: "  + counts.rejected   + " ราย\n" +
              "⚠️ CONFLICT: "  + counts.conflict   + " ราย\n" +
              "⏭️ ข้ามไป: "   + counts.skipped    + " ราย\n" +
              "✅ ทำแล้ว: "    + counts.alreadyDone + " ราย";

    if (counts.conflict > 0) {
      msg += "\n\n⚠️ พบ " + counts.conflict + " แถว CONFLICT\n" +
             "กรุณาตรวจสอบและ reset checkbox ก่อนตัดสินใจใหม่ครับ";
    }

    console.log("[applyApprovedFeedback] DONE — " + JSON.stringify(counts));
    ui.alert(msg);

  } catch(e) {
    console.error("[applyApprovedFeedback] Error: " + e.message);
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
    total:    0,
    pending:  0,
    approved: 0,
    rejected: 0,
    gpsDiff:  0,
    noGPS:    0
  };
  
  data.forEach(function(row) {
    var reason   = row[6]; // Col G
    var approve  = row[7]; // Col H
    var reject   = row[8]; // Col I
    
    // นับเฉพาะแถวที่มีข้อมูลจริง
    if (!row[0]) return;
    stats.total++;
    
    if (reason === "APPROVED")       stats.approved++;
    else if (reason === "REJECTED")  stats.rejected++;
    else if (approve)                stats.approved++;
    else if (reject)                 stats.rejected++;
    else                             stats.pending++;
    
    if (reason === "GPS_DIFF")       stats.gpsDiff++;
    else if (reason === "DB_NO_GPS") stats.noGPS++;
  });
  
  var msg = 
    "📊 GPS Queue สถิติ\n" +
    "━━━━━━━━━━━━━━━━━━━━━━━\n" +
    "📋 รายการทั้งหมด: "  + stats.total    + " ราย\n" +
    "⏳ รอตรวจสอบ: "      + stats.pending  + " ราย\n" +
    "✅ อนุมัติแล้ว: "    + stats.approved + " ราย\n" +
    "❌ ปฏิเสธแล้ว: "     + stats.rejected + " ราย\n" +
    "━━━━━━━━━━━━━━━━━━━━━━━\n" +
    "📍 พิกัดต่างกัน > 50m: " + stats.gpsDiff + " ราย\n" +
    "🔍 DB ไม่มีพิกัด: "      + stats.noGPS   + " ราย";
  
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
    reasonData.forEach(function(row, i) {
      var rowNum = i + 2;
      var reason = row[0];
      var bg = "#ffffff";
      
      if (reason === "GPS_DIFF")   bg = "#fff3cd"; // เหลือง
      if (reason === "DB_NO_GPS")  bg = "#f8d7da"; // แดง
      if (reason === "NO_MATCH")   bg = "#d1ecf1"; // ฟ้า
      if (reason === "APPROVED")   bg = "#d4edda"; // เขียว
      if (reason === "REJECTED")   bg = "#e2e3e5"; // เทา
      
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
