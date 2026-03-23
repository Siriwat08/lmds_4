/**
 * 🛡️ Service: Schema Validator (Enterprise Edition)
 * Version: 4.1
 * หน้าที่: ตรวจสอบโครงสร้างชีตก่อนทำงานทุก Flow สำคัญ
 * ป้องกันระบบพังเงียบๆ เมื่อมีการแก้ไข Header หรือขยับคอลัมน์
 */

// ==========================================
// 1. SCHEMA DEFINITIONS
// ==========================================

var SHEET_SCHEMA = {

  // Database Sheet
  DATABASE: {
    sheetName: function () { return CONFIG.SHEET_NAME; },
    minColumns: 20,
    requiredHeaders: {
      1: "NAME",
      2: "LAT",
      3: "LNG",
      11: "UUID",
      15: "QUALITY",
      16: "CREATED",
      17: "UPDATED",
      18: "Coord_Source",
      19: "Coord_Confidence",
      20: "Coord_Last_Updated"
    }
  },

  // NameMapping Sheet
  NAMEMAPPING: {
    sheetName: function () { return CONFIG.MAPPING_SHEET; },
    minColumns: 5,
    requiredHeaders: {
      1: "Variant_Name",
      2: "Master_UID",
      3: "Confidence_Score",
      4: "Mapped_By",
      5: "Timestamp"
    }
  },

  SCG_SOURCE: {
    sheetName: function () { return CONFIG.SOURCE_SHEET; },
    minColumns: 37,
    requiredHeaders: {
      13: "ชื่อปลายทาง",
      15: "LAT",
      16: "LONG",
      19: "ที่อยู่ปลายทาง",
      24: "ระยะทางจากคลัง_Km",
      25: "ชื่อที่อยู่จาก_LatLong",
      37: "SYNC_STATUS"
    }
  },

  // GPS_Queue Sheet
  GPS_QUEUE: {
    sheetName: function () { return SCG_CONFIG.SHEET_GPS_QUEUE; },
    minColumns: 9,
    requiredHeaders: {
      1: "Timestamp",
      2: "ShipToName",
      3: "UUID_DB",
      4: "LatLng_Driver",
      5: "LatLng_DB",
      6: "Diff_Meters",
      7: "Reason",
      8: "Approve",
      9: "Reject"
    }
  },

  // Data Sheet
  DATA: {
    sheetName: function () { return SCG_CONFIG.SHEET_DATA; },
    minColumns: 27,
    requiredHeaders: {
      1: "ID_งานประจำวัน",
      4: "ShipmentNo",
      11: "ShipToName",
      27: "LatLong_Actual"
    }
  }
};

// ==========================================
// 2. CORE VALIDATOR
// ==========================================

/**
 * ตรวจสอบชีตเดียว
 * @returns {object} { valid: bool, errors: [] }
 */
function validateSheet_(schemaKey) {
  var schema = SHEET_SCHEMA[schemaKey];
  if (!schema) return { valid: false, errors: ["Schema '" + schemaKey + "' ไม่พบ"] };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = schema.sheetName();
  var errors = [];

  // 1. ตรวจว่าชีตมีอยู่
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    return {
      valid: false,
      errors: ["❌ ไม่พบชีต '" + sheetName + "'"]
    };
  }

  // 2. ตรวจจำนวนคอลัมน์
  var lastCol = sheet.getLastColumn();
  if (lastCol < schema.minColumns) {
    errors.push(
      "❌ ชีต '" + sheetName + "' มีแค่ " + lastCol +
      " คอลัมน์ (ต้องการ ≥ " + schema.minColumns + ")"
    );
  }

  // 3. ตรวจ Header
  if (lastCol > 0 && sheet.getLastRow() > 0) {
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

    Object.keys(schema.requiredHeaders).forEach(function (colNum) {
      var idx = parseInt(colNum) - 1;
      var expected = schema.requiredHeaders[colNum];
      var actual = headers[idx] || "";

      if (actual.toString().trim() !== expected.toString().trim()) {
        errors.push(
          "⚠️ Col " + colNum + ": คาดว่า '" + expected +
          "' แต่เจอ '" + actual + "'"
        );
      }
    });
  }

  return {
    valid: errors.length === 0,
    errors: errors,
    sheetName: sheetName
  };
}

/**
 * ตรวจสอบหลายชีตพร้อมกัน
 * @param {string[]} schemaKeys - รายชื่อ schema ที่จะตรวจ
 * @returns {object} { allValid: bool, results: {} }
 */
function validateSchemas(schemaKeys) {
  var results = {};
  var allValid = true;

  schemaKeys.forEach(function (key) {
    var result = validateSheet_(key);
    results[key] = result;
    if (!result.valid) allValid = false;
  });

  return { allValid: allValid, results: results };
}

// ==========================================
// 3. PRE-FLIGHT CHECKS
// ==========================================

/**
 * ตรวจสอบก่อนรัน syncNewDataToMaster
 */
function preCheck_Sync() {
  var check = validateSchemas(["DATABASE", "NAMEMAPPING", "SCG_SOURCE", "GPS_QUEUE"]);
  if (!check.allValid) {
    throwSchemaError_(check.results, "Sync GPS Feedback");
  }
  return true;
}

/**
 * ตรวจสอบก่อนรัน applyMasterCoordinatesToDailyJob
 */
function preCheck_Apply() {
  var check = validateSchemas(["DATABASE", "NAMEMAPPING", "DATA"]);
  if (!check.allValid) {
    throwSchemaError_(check.results, "Apply Master Coordinates");
  }
  return true;
}

/**
 * ตรวจสอบก่อนรัน applyApprovedFeedback
 */
function preCheck_Approve() {
  var check = validateSchemas(["DATABASE", "GPS_QUEUE"]);
  if (!check.allValid) {
    throwSchemaError_(check.results, "Apply Approved Feedback");
  }
  return true;
}

// ==========================================
// 4. ERROR REPORTER
// ==========================================

function throwSchemaError_(results, flowName) {
  var msg = "❌ Schema Validation Failed\n";
  msg += "Flow: " + flowName + "\n";
  msg += "━━━━━━━━━━━━━━━━━━━━━━━\n\n";

  Object.keys(results).forEach(function (key) {
    var r = results[key];
    if (!r.valid) {
      msg += "📋 ชีต: " + r.sheetName + "\n";
      r.errors.forEach(function (e) {
        msg += "  " + e + "\n";
      });
      msg += "\n";
    }
  });

  msg += "━━━━━━━━━━━━━━━━━━━━━━━\n";
  msg += "💡 กรุณาตรวจสอบโครงสร้างชีตก่อนรันใหม่\n";
  msg += "หรือกดเมนู ⚙️ System Admin → 🔬 Diagnostic";

  SpreadsheetApp.getActiveSpreadsheet().toast(
    "❌ Schema Error: " + flowName, "System Alert", 10
  );
  throw new Error(msg);
}

// ==========================================
// 5. FULL DIAGNOSTIC (เรียกจากเมนู)
// ==========================================

function runFullSchemaValidation() {
  var ui = SpreadsheetApp.getUi();
  var allKeys = Object.keys(SHEET_SCHEMA);
  var check = validateSchemas(allKeys);

  var msg = "🛡️ Schema Validation Report\n";
  msg += "━━━━━━━━━━━━━━━━━━━━━━━━━\n\n";

  allKeys.forEach(function (key) {
    var r = check.results[key];
    if (r.valid) {
      msg += "✅ " + r.sheetName + "\n";
    } else {
      msg += "❌ " + r.sheetName + "\n";
      r.errors.forEach(function (e) {
        msg += "   " + e + "\n";
      });
    }
  });

  msg += "\n━━━━━━━━━━━━━━━━━━━━━━━━━\n";
  msg += check.allValid
    ? "✅ ทุกชีตผ่านการตรวจสอบ พร้อมใช้งาน"
    : "❌ พบปัญหา กรุณาแก้ไขก่อนใช้งาน";

  ui.alert(msg);
}