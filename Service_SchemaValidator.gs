/**
 * VERSION: 4.2 — Phase A
 * [Phase A] minColumns = 22, central headers, queue conflict validation
 */

var SHEET_SCHEMA = {

  DATABASE: {
    sheetName:       function() { return CONFIG.SHEET_NAME; },
    minColumns:      22,
    requiredHeaders: CONFIG.DB_REQUIRED_HEADERS
  },

  NAMEMAPPING: {
    sheetName:       function() { return CONFIG.MAPPING_SHEET; },
    minColumns:      CONFIG.MAP_TOTAL_COLS,
    requiredHeaders: CONFIG.MAP_REQUIRED_HEADERS
  },

  SCG_SOURCE: {
    sheetName:    function() { return CONFIG.SOURCE_SHEET; },
    minColumns:   37,
    requiredHeaders: {
      13: "ชื่อปลายทาง", 15: "LAT", 16: "LONG",
      19: "ที่อยู่ปลายทาง", 24: "ระยะทางจากคลัง_Km",
      25: "ชื่อที่อยู่จาก_LatLong", 37: "SYNC_STATUS"
    }
  },

  GPS_QUEUE: {
    sheetName:       function() { return SCG_CONFIG.SHEET_GPS_QUEUE; },
    minColumns:      CONFIG.GPS_QUEUE_TOTAL_COLS,
    requiredHeaders: CONFIG.GPS_QUEUE_REQUIRED_HEADERS
  },

  DATA: {
    sheetName:    function() { return SCG_CONFIG.SHEET_DATA; },
    minColumns:   27,
    requiredHeaders: {
      1: "ID_งานประจำวัน", 4: "ShipmentNo",
      11: "ShipToName", 27: "LatLong_Actual"
    }
  }
};

// ==========================================
// CORE VALIDATOR
// ==========================================

function validateSheet_(schemaKey) {
  var schema = SHEET_SCHEMA[schemaKey];
  if (!schema) return { valid: false, errors: ["Schema '" + schemaKey + "' ไม่พบ"] };

  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = schema.sheetName();
  var errors    = [];
  var sheet     = ss.getSheetByName(sheetName);

  if (!sheet) return { valid: false, errors: ["❌ ไม่พบชีต '" + sheetName + "'"] };

  var lastCol = sheet.getLastColumn();
  if (lastCol < schema.minColumns) {
    errors.push("❌ ชีต '" + sheetName + "' มีแค่ " + lastCol +
                " คอลัมน์ (ต้องการ ≥ " + schema.minColumns + ")");
  }

  if (lastCol > 0 && sheet.getLastRow() > 0) {
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    Object.keys(schema.requiredHeaders).forEach(function(colNum) {
      var idx      = parseInt(colNum) - 1;
      var expected = schema.requiredHeaders[colNum];
      var actual   = headers[idx] || "";
      if (actual.toString().trim() !== expected.toString().trim()) {
        errors.push("⚠️ Col " + colNum + ": คาดว่า '" + expected +
                    "' แต่เจอ '" + actual + "'");
      }
    });
  }

  return { valid: errors.length === 0, errors: errors, sheetName: sheetName };
}

// [Phase A NEW] ตรวจ GPS Queue conflict
function validateGPSQueueIntegrity_(sheet) {
  var issues  = [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return issues;
  var data = sheet.getRange(2, 8, lastRow - 1, 2).getValues();
  data.forEach(function(row, i) {
    if (row[0] === true && row[1] === true) {
      issues.push("แถว " + (i + 2) + ": Approve และ Reject ถูกติ๊กพร้อมกัน");
    }
  });
  return issues;
}

function validateSchemas(schemaKeys) {
  var results  = {};
  var allValid = true;
  schemaKeys.forEach(function(key) {
    var result   = validateSheet_(key);
    results[key] = result;
    if (!result.valid) allValid = false;
  });
  return { allValid: allValid, results: results };
}

function preCheck_Sync() {
  var check = validateSchemas(["DATABASE","NAMEMAPPING","SCG_SOURCE","GPS_QUEUE"]);
  if (!check.allValid) throwSchemaError_(check.results, "Sync GPS Feedback");
  return true;
}

function preCheck_Apply() {
  var check = validateSchemas(["DATABASE","NAMEMAPPING","DATA"]);
  if (!check.allValid) throwSchemaError_(check.results, "Apply Master Coordinates");
  return true;
}

function preCheck_Approve() {
  var check = validateSchemas(["DATABASE","GPS_QUEUE"]);
  if (!check.allValid) throwSchemaError_(check.results, "Apply Approved Feedback");
  return true;
}

function throwSchemaError_(results, flowName) {
  var msg = "❌ Schema Validation Failed\nFlow: " + flowName +
            "\n━━━━━━━━━━━━━━━━━━━━━━━\n\n";
  Object.keys(results).forEach(function(key) {
    var r = results[key];
    if (!r.valid) {
      msg += "📋 ชีต: " + r.sheetName + "\n";
      r.errors.forEach(function(e) { msg += "  " + e + "\n"; });
      msg += "\n";
    }
  });
  msg += "━━━━━━━━━━━━━━━━━━━━━━━\n💡 กรุณาตรวจสอบโครงสร้างชีตก่อนรันใหม่";
  SpreadsheetApp.getActiveSpreadsheet().toast("❌ Schema Error: " + flowName, "Alert", 10);
  throw new Error(msg);
}

function runFullSchemaValidation() {
  var ui      = SpreadsheetApp.getUi();
  var allKeys = Object.keys(SHEET_SCHEMA);
  var check   = validateSchemas(allKeys);

  // [Phase A NEW] ตรวจ GPS Queue conflict ด้วย
  var msg = "🛡️ Schema Validation Report\n━━━━━━━━━━━━━━━━━━━━━━━━━\n\n";
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var qs  = ss.getSheetByName(SCG_CONFIG.SHEET_GPS_QUEUE);
  if (qs) {
    var qi = validateGPSQueueIntegrity_(qs);
    if (qi.length > 0) {
      msg += "⚠️ GPS_Queue Conflicts:\n";
      qi.forEach(function(i) { msg += "  " + i + "\n"; });
      msg += "\n";
    }
  }

  allKeys.forEach(function(key) {
    var r = check.results[key];
    msg += (r.valid ? "✅ " : "❌ ") + r.sheetName + "\n";
    if (!r.valid) r.errors.forEach(function(e) { msg += "   " + e + "\n"; });
  });

  msg += "\n━━━━━━━━━━━━━━━━━━━━━━━━━\n";
  msg += check.allValid ? "✅ ทุกชีตผ่านการตรวจสอบ" : "❌ พบปัญหา กรุณาแก้ไขก่อนใช้งาน";
  ui.alert(msg);
}

function fixNameMappingHeaders() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var ui    = SpreadsheetApp.getUi();
  var sheet = ss.getSheetByName(CONFIG.MAPPING_SHEET);
  if (!sheet) { ui.alert("❌ ไม่พบชีต NameMapping"); return; }
  var headers = Object.values(CONFIG.MAP_REQUIRED_HEADERS);
  var r = sheet.getRange(1, 1, 1, 5);
  r.setValues([headers]);
  r.setFontWeight("bold").setBackground("#7c3aed").setFontColor("white");
  sheet.setFrozenRows(1);
  SpreadsheetApp.flush();
  ui.alert("✅ อัปเกรด NameMapping Header สำเร็จ!");
}
