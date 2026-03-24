/**
 * VERSION : 4.2 — Phase E
 * [Phase E] แยก diagnostics core คืน object
 * [Phase E] เพิ่ม runDryRunMappingConflicts(), runDryRunUUIDIntegrity()
 */

// ==========================================
// 1. CORE DIAGNOSTICS (คืน object — ไม่ผูกกับ UI)
// ==========================================

function collectSystemDiagnostics_() {
  var results = [];

  function pass(msg) { results.push({ status: "pass", message: msg }); }
  function warn(msg) { results.push({ status: "warn", message: msg }); }
  function fail(msg) { results.push({ status: "fail", message: msg }); }

  // Config
  if (typeof CONFIG !== 'undefined') pass("CONFIG: พร้อมใช้งาน");
  else fail("CONFIG: ไม่พบ");

  // Utils
  if (typeof md5           === 'function') pass("md5(): พร้อม");
  else fail("md5(): ไม่พบ");

  if (typeof normalizeText === 'function') pass("normalizeText(): พร้อม");
  else fail("normalizeText(): ไม่พบ");

  if (typeof dbRowToObject === 'function') pass("dbRowToObject(): พร้อม [Phase B]");
  else fail("dbRowToObject(): ไม่พบ — Phase B ยังไม่ได้ deploy");

  if (typeof buildUUIDStateMap_ === 'function') pass("buildUUIDStateMap_(): พร้อม [Phase C]");
  else fail("buildUUIDStateMap_(): ไม่พบ — Phase C ยังไม่ได้ deploy");

  if (typeof retrieveCandidateMasters_ === 'function') pass("retrieveCandidateMasters_(): พร้อม [Phase D]");
  else warn("retrieveCandidateMasters_(): ไม่พบ — Phase D ยังไม่ได้ deploy");

  // Google Maps API
  if (typeof GET_ADDR_WITH_CACHE === 'function') {
    try {
      var testGeo = GET_ADDR_WITH_CACHE(13.746, 100.539);
      if (testGeo && testGeo !== "Error") pass("Google Maps API: ทำงานปกติ");
      else warn("Google Maps API: ได้ค่าแปลก");
    } catch(e) { fail("Google Maps API: Error — " + e.message); }
  } else {
    fail("Google Maps API: ไม่พบ GET_ADDR_WITH_CACHE");
  }

  // Gemini API Key
  try {
    if (CONFIG && CONFIG.GEMINI_API_KEY) pass("Gemini API Key: พร้อมใช้งาน");
  } catch(e) { fail("Gemini API Key: ไม่พบ — " + e.message); }

  // Notifications
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('LINE_NOTIFY_TOKEN')) pass("LINE Notify: ตั้งค่าแล้ว");
  else warn("LINE Notify: ยังไม่ได้ตั้งค่า");

  if (props.getProperty('TG_BOT_TOKEN') && props.getProperty('TG_CHAT_ID')) {
    pass("Telegram: ตั้งค่าแล้ว");
  } else {
    warn("Telegram: ยังไม่ได้ตั้งค่า");
  }

  return results;
}

function collectSheetDiagnostics_() {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var results = [];

  function pass(msg) { results.push({ status: "pass", message: msg }); }
  function warn(msg) { results.push({ status: "warn", message: msg }); }
  function fail(msg) { results.push({ status: "fail", message: msg }); }

  // Database
  var dbSheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (dbSheet) {
    var rows = typeof getRealLastRow_ === 'function'
      ? getRealLastRow_(dbSheet, CONFIG.COL_NAME)
      : dbSheet.getLastRow();
    var cols = dbSheet.getLastColumn();
    if (rows >= 2) pass("Database: " + (rows-1) + " records | " + cols + " columns");
    else warn("Database: ว่างเปล่า");
    if (cols < CONFIG.DB_TOTAL_COLS) {
      warn("Database: columns " + cols + " < " + CONFIG.DB_TOTAL_COLS + " (ควรรัน upgradeDatabaseStructure)");
    }
  } else { fail("Database: ไม่พบชีต"); }

  // NameMapping
  var mapSheet = ss.getSheetByName(CONFIG.MAPPING_SHEET);
  if (mapSheet) {
    var mapCols = mapSheet.getLastColumn();
    if (mapCols >= CONFIG.MAP_TOTAL_COLS) pass("NameMapping: schema V4.2 ถูกต้อง");
    else warn("NameMapping: " + mapCols + " columns (ควร " + CONFIG.MAP_TOTAL_COLS + ")");
  } else { fail("NameMapping: ไม่พบชีต"); }

  // GPS Queue
  var queueSheet = ss.getSheetByName(SCG_CONFIG.SHEET_GPS_QUEUE);
  if (queueSheet) {
    var qRows = queueSheet.getLastRow() - 1;
    pass("GPS_Queue: " + qRows + " records");
  } else { warn("GPS_Queue: ไม่พบชีต"); }

  // Source Sheet
  var srcSheet = ss.getSheetByName(CONFIG.SOURCE_SHEET);
  if (srcSheet) pass("Source Sheet: พบชีต '" + CONFIG.SOURCE_SHEET + "'");
  else warn("Source Sheet: ไม่พบชีต '" + CONFIG.SOURCE_SHEET + "'");

  // Data, Input, PostalRef
  [SCG_CONFIG.SHEET_DATA, SCG_CONFIG.SHEET_INPUT, CONFIG.SHEET_POSTAL].forEach(function(name) {
    if (ss.getSheetByName(name)) pass(name + ": พบชีต");
    else warn(name + ": ไม่พบชีต");
  });

  return results;
}

// ==========================================
// 2. UI WRAPPERS
// ==========================================

function RUN_SYSTEM_DIAGNOSTIC() {
  var ui      = SpreadsheetApp.getUi();
  var results = collectSystemDiagnostics_();

  var msg = "🏥 System Diagnostic Report (V4.2)\n━━━━━━━━━━━━━━━━━━━━━━━\n\n";
  results.forEach(function(r) {
    var icon = r.status === "pass" ? "✅" : r.status === "warn" ? "⚠️" : "❌";
    msg += icon + " " + r.message + "\n";
  });
  msg += "\n━━━━━━━━━━━━━━━━━━━━━━━";
  ui.alert(msg);
}

function RUN_SHEET_DIAGNOSTIC() {
  var ui      = SpreadsheetApp.getUi();
  var results = collectSheetDiagnostics_();

  var msg = "🕵️ Sheet Diagnostic Report (V4.2)\n━━━━━━━━━━━━━━━━━━━━━━━\n\n";
  results.forEach(function(r) {
    var icon = r.status === "pass" ? "✅" : r.status === "warn" ? "⚠️" : "❌";
    msg += icon + " " + r.message + "\n";
  });
  msg += "\n━━━━━━━━━━━━━━━━━━━━━━━";
  ui.alert(msg);
}

// ==========================================
// 3. [Phase E NEW] DRY RUN FUNCTIONS
// ==========================================

/**
 * runDryRunMappingConflicts()
 * ตรวจ NameMapping conflicts โดยไม่แก้ไขข้อมูลจริง
 */
function runDryRunMappingConflicts() {
  var ui       = SpreadsheetApp.getUi();
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var mapSheet = ss.getSheetByName(CONFIG.MAPPING_SHEET);

  if (!mapSheet || mapSheet.getLastRow() < 2) {
    ui.alert("ℹ️ NameMapping ว่างเปล่า");
    return;
  }

  var mapRows   = mapSheet.getRange(2, 1, mapSheet.getLastRow() - 1, CONFIG.MAP_TOTAL_COLS).getValues();
  var variantMap = {};
  var conflicts  = [];
  var duplicates = [];

  // โหลด DB UUID
  var dbSheet   = ss.getSheetByName(CONFIG.SHEET_NAME);
  var lastDbRow = getRealLastRow_(dbSheet, CONFIG.COL_NAME);
  var dbData    = (lastDbRow > 1)
    ? dbSheet.getRange(2, 1, lastDbRow - 1, CONFIG.DB_TOTAL_COLS).getValues()
    : [];
  var validUUIDs = new Set();
  dbData.forEach(function(r) {
    if (r[CONFIG.C_IDX.UUID]) validUUIDs.add(r[CONFIG.C_IDX.UUID]);
  });

  mapRows.forEach(function(r, i) {
    var obj       = mapRowToObject(r);
    var normVar   = normalizeText(obj.variant || "");
    if (!normVar) return;

    // ตรวจ variant ซ้ำ
    if (variantMap[normVar]) {
      duplicates.push("แถว " + (i+2) + ": variant '" + obj.variant + "' ซ้ำกับแถว " + variantMap[normVar]);
    } else {
      variantMap[normVar] = i + 2;
    }

    // ตรวจ UUID ไม่มีใน DB
    if (obj.uid && !validUUIDs.has(obj.uid)) {
      conflicts.push("แถว " + (i+2) + ": UUID '" + obj.uid.substring(0,12) + "...' ไม่พบใน Database");
    }

    // ตรวจ UUID ว่าง
    if (!obj.uid) {
      conflicts.push("แถว " + (i+2) + ": variant '" + obj.variant + "' ไม่มี UUID");
    }
  });

  var msg = "🔵 Dry Run: Mapping Conflicts\n━━━━━━━━━━━━━━━━━━━━━━━\n\n" +
            "📋 ตรวจ: " + mapRows.length + " mappings\n\n";

  if (duplicates.length === 0 && conflicts.length === 0) {
    msg += "✅ ไม่พบ conflicts ทั้งหมด\n";
  } else {
    if (duplicates.length > 0) {
      msg += "⚠️ Variant ซ้ำ: " + duplicates.length + " รายการ\n";
      duplicates.slice(0, 5).forEach(function(d) { msg += "  " + d + "\n"; });
      if (duplicates.length > 5) msg += "  ...และอีก " + (duplicates.length - 5) + " รายการ\n";
      msg += "\n";
    }
    if (conflicts.length > 0) {
      msg += "❌ UUID conflicts: " + conflicts.length + " รายการ\n";
      conflicts.slice(0, 5).forEach(function(c) { msg += "  " + c + "\n"; });
      if (conflicts.length > 5) msg += "  ...และอีก " + (conflicts.length - 5) + " รายการ\n";
    }
  }

  msg += "\n💡 Dry Run — ไม่มีการแก้ไขข้อมูลจริงครับ";
  ui.alert(msg);
}

/**
 * runDryRunUUIDIntegrity()
 * ตรวจ UUID integrity ใน Database โดยไม่แก้ไขข้อมูลจริง
 */
function runDryRunUUIDIntegrity() {
  var ui      = SpreadsheetApp.getUi();
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var sheet   = ss.getSheetByName(CONFIG.SHEET_NAME);
  var lastRow = getRealLastRow_(sheet, CONFIG.COL_NAME);

  if (lastRow < 2) { ui.alert("ℹ️ Database ว่างเปล่า"); return; }

  var data        = sheet.getRange(2, 1, lastRow - 1, CONFIG.DB_TOTAL_COLS).getValues();
  var uuidSet     = new Set();
  var issues      = [];
  var stats       = { total: 0, noUUID: 0, dupUUID: 0, merged: 0, inactive: 0, noStatus: 0 };

  data.forEach(function(row, i) {
    var obj = dbRowToObject(row);
    if (!obj.name) return;
    stats.total++;

    if (!obj.uuid) {
      stats.noUUID++;
      issues.push("แถว " + (i+2) + ": '" + obj.name + "' ไม่มี UUID");
    } else if (uuidSet.has(obj.uuid)) {
      stats.dupUUID++;
      issues.push("แถว " + (i+2) + ": UUID ซ้ำ — " + obj.uuid.substring(0,12) + "...");
    } else {
      uuidSet.add(obj.uuid);
    }

    var status = obj.recordStatus || "";
    if (status === "Merged")   stats.merged++;
    if (status === "Inactive") stats.inactive++;
    if (!status)               stats.noStatus++;
  });

  var msg = "🔵 Dry Run: UUID Integrity\n━━━━━━━━━━━━━━━━━━━━━━━\n\n" +
            "📋 ตรวจ: " + stats.total + " records\n" +
            "❌ ไม่มี UUID: "    + stats.noUUID   + " records\n" +
            "🔁 UUID ซ้ำ: "      + stats.dupUUID  + " records\n" +
            "🔀 Merged: "         + stats.merged   + " records\n" +
            "⚫ Inactive: "        + stats.inactive + " records\n" +
            "❓ ไม่มี Status: "   + stats.noStatus + " records\n\n";

  if (issues.length > 0) {
    msg += "⚠️ Issues:\n";
    issues.slice(0, 8).forEach(function(iss) { msg += "  " + iss + "\n"; });
    if (issues.length > 8) msg += "  ...และอีก " + (issues.length - 8) + " รายการ\n";
  } else {
    msg += "✅ UUID integrity ปกติทั้งหมด\n";
  }

  msg += "\n💡 Dry Run — ไม่มีการแก้ไขข้อมูลจริงครับ";
  ui.alert(msg);
}
