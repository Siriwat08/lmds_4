/**
 * VERSION: 000
 * 🚛 Logistics Master Data System - Configuration V4.0 (Enterprise Edition)
 * ------------------------------------------------------------------
 * [PRESERVED]: โครงสร้างเดิมทั้งหมดได้รับการรักษาไว้ (Preservation Protocol)
 * [ADDED v4.0]: กำหนดคอลัมน์ NameMapping สำหรับ 4-Tier Smart Resolution
 * [ADDED v4.0]: ตัวแปรควบคุม AI Batch Size และ Cache Expiration
 * [MODIFIED]: อัปเกรดระบบ Logging เป็น console.log/error สำหรับ GCP Monitoring
 * Author: Elite Logistics Architect
 */

var CONFIG = {
  // --- SHEET NAMES ---
  SHEET_NAME: "Database",
  MAPPING_SHEET: "NameMapping",
  SOURCE_SHEET: "SCGนครหลวงJWDภูมิภาค",
  SHEET_POSTAL: "PostalRef", // รองรับ Service_GeoAddr

  // --- 🧠 AI CONFIGURATION (SECURED) ---
  // วิธีตั้งค่า: รันฟังก์ชัน setupEnvironment() ในไฟล์ Setup_Security.gs
  get GEMINI_API_KEY() {
    var key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!key) throw new Error("CRITICAL ERROR: GEMINI_API_KEY is not set. Please run setupEnvironment() first.");
    return key;
  },
  USE_AI_AUTO_FIX: true,
  AI_MODEL: "gemini-1.5-flash", 
  AI_BATCH_SIZE: 20, // [ADDED v4.0]: จำกัดจำนวนส่งให้ AI ครั้งละ 20 รายการเพื่อไม่ให้เกิน 6 นาที

  // --- 🔴 DEPOT LOCATION ---
  DEPOT_LAT: 14.164688, 
  DEPOT_LNG: 100.625354,

  // --- SYSTEM THRESHOLDS & LIMITS ---
  DISTANCE_THRESHOLD_KM: 0.05, 
  BATCH_LIMIT: 50,  
  DEEP_CLEAN_LIMIT: 100,
  API_MAX_RETRIES: 3,       // จำนวนครั้งที่จะลองใหม่ถ้า API SCG ล่ม
  API_TIMEOUT_MS: 30000,    // เวลา Timeout (30 วิ)
  CACHE_EXPIRATION: 21600,  // [ADDED v4.0]: เวลา Cache (วินาที) -> 6 ชั่วโมง (สำหรับ Geo Maps)

  // --- DATABASE COLUMNS INDEX (1-BASED) ---
  COL_NAME: 1,       // A: ชื่อลูกค้า
  COL_LAT: 2,        // B: Latitude
  COL_LNG: 3,        // C: Longitude
  COL_SUGGESTED: 4,  // D: ชื่อที่ระบบแนะนำ
  COL_CONFIDENCE: 5, // E: ความมั่นใจ
  COL_NORMALIZED: 6, // F: ชื่อที่ Clean แล้ว
  COL_VERIFIED: 7,   // G: สถานะตรวจสอบ (Checkbox)
  COL_SYS_ADDR: 8,   // H: ที่อยู่จากระบบต้นทาง
  COL_ADDR_GOOG: 9,  // I: ที่อยู่จาก Google Maps
  COL_DIST_KM: 10,   // J: ระยะทางจากคลัง
  COL_UUID: 11,      // K: Unique ID
  COL_PROVINCE: 12,  // L: จังหวัด
  COL_DISTRICT: 13,  // M: อำเภอ
  COL_POSTCODE: 14,  // N: รหัสไปรษณีย์
  COL_QUALITY: 15,   // O: Quality Score
  COL_CREATED: 16,   // P: วันที่สร้าง (Created)
  COL_UPDATED: 17,   // Q: วันที่แก้ไขล่าสุด (Updated)

    // --- [NEW] GPS TRACKING COLUMNS ---
  COL_COORD_SOURCE:       18, // R: พิกัดมาจากไหน (SCG_System / Driver_GPS)
  COL_COORD_CONFIDENCE:   19, // S: ความน่าเชื่อถือ 0-100
  COL_COORD_LAST_UPDATED: 20, // T: อัปเดตพิกัดล่าสุดเมื่อไร
  COL_RECORD_STATUS:  21, // U: Active / Inactive / Merged
COL_MERGED_TO_UUID: 22, // V: UUID ที่ชี้ไปหลัง Merge


  // --- [NEW v4.0] NAMEMAPPING COLUMNS INDEX (1-BASED) ---
  // เตรียมโครงสร้างให้ AI ทำการ Map ชื่อสกปรกเข้ากับชื่อจริง
  MAP_COL_VARIANT: 1,    // A: Variant_Name (ชื่อแปลกๆ เช่น บจก. เอบีซี, เอบีซี จำกัด)
  MAP_COL_UID: 2,        // B: Master_UID (รหัสอ้างอิง Database หรือชื่อจริง)
  MAP_COL_CONFIDENCE: 3, // C: Confidence_Score (ความมั่นใจ AI 0-100)
  MAP_COL_MAPPED_BY: 4,  // D: Mapped_By (Human / AI)
  MAP_COL_TIMESTAMP: 5,  // E: Timestamp (เวลาที่อัปเดต)

  // --- DATABASE ARRAY INDEX MAPPING (0-BASED) ---
  get C_IDX() {
    return {
      NAME: this.COL_NAME - 1,
      LAT: this.COL_LAT - 1,
      LNG: this.COL_LNG - 1,
      SUGGESTED: this.COL_SUGGESTED - 1,
      CONFIDENCE: this.COL_CONFIDENCE - 1,
      NORMALIZED: this.COL_NORMALIZED - 1,
      VERIFIED: this.COL_VERIFIED - 1,
      SYS_ADDR: this.COL_SYS_ADDR - 1,
      GOOGLE_ADDR: this.COL_ADDR_GOOG - 1,
      DIST_KM: this.COL_DIST_KM - 1,
      UUID: this.COL_UUID - 1,
      PROVINCE: this.COL_PROVINCE - 1,
      DISTRICT: this.COL_DISTRICT - 1,
      POSTCODE: this.COL_POSTCODE - 1,
      QUALITY: this.COL_QUALITY - 1,
      CREATED: this.COL_CREATED - 1,
      UPDATED: this.COL_UPDATED - 1,
      COORD_SOURCE:       this.COL_COORD_SOURCE - 1,
      COORD_CONFIDENCE:   this.COL_COORD_CONFIDENCE - 1,
      COORD_LAST_UPDATED: this.COL_COORD_LAST_UPDATED - 1,
      RECORD_STATUS:  this.COL_RECORD_STATUS - 1,
      MERGED_TO_UUID: this.COL_MERGED_TO_UUID - 1
    };
  },

  // --- [NEW v4.0] NAMEMAPPING ARRAY INDEX (0-BASED) ---
  get MAP_IDX() {
    return {
      VARIANT: this.MAP_COL_VARIANT - 1,
      UID: this.MAP_COL_UID - 1,
      CONFIDENCE: this.MAP_COL_CONFIDENCE - 1,
      MAPPED_BY: this.MAP_COL_MAPPED_BY - 1,
      TIMESTAMP: this.MAP_COL_TIMESTAMP - 1
    };
  }
};

// --- SCG SPECIFIC CONFIG ---
const SCG_CONFIG = {
  SHEET_DATA: 'Data',
  SHEET_INPUT: 'Input',
  SHEET_EMPLOYEE: 'ข้อมูลพนักงาน',
  API_URL: 'https://fsm.scgjwd.com/Monitor/SearchDelivery',
  INPUT_START_ROW: 4,
  COOKIE_CELL: 'B1',
  SHIPMENT_STRING_CELL: 'B3',
  SHEET_MASTER_DB: 'Database',
  SHEET_MAPPING: 'NameMapping',

  // --- [NEW] GPS FEEDBACK ---
  SHEET_GPS_QUEUE: 'GPS_Queue',
  GPS_THRESHOLD_METERS: 50,

  // --- [NEW] SCGนครหลวงJWDภูมิภาค COLUMN INDEX (0-based) ---
  // ย้ายมาจาก Service_Master.gs เพื่อให้แก้ที่เดียว
  SRC_IDX: {
    NAME:     12,  // Col M: ชื่อปลายทาง
    LAT:      14,  // Col O: LAT (GPS จริงจากคนขับ)
    LNG:      15,  // Col P: LONG (GPS จริงจากคนขับ)
    SYS_ADDR: 18,  // Col S: ที่อยู่ปลายทาง
    DIST:     23,  // Col X: ระยะทางจากคลัง_Km
    GOOG_ADDR:24   // Col Y: ชื่อที่อยู่จาก_LatLong
  },

  // --- [NEW v4.1] SYNC TRACKING ---
  SRC_IDX_SYNC_STATUS: 37, // Col AK: สถานะการ Sync (0-based = 35)
  SYNC_STATUS_DONE: "SYNCED",

  JSON_MAP: {
    SHIPMENT_NO: 'shipmentNo',
    CUSTOMER_NAME: 'customerName',
    DELIVERY_DATE: 'deliveryDate'
  }
};

/**
 * [ENHANCED v4.0] System Health Check
 * ตรวจสอบความพร้อมของ Sheet และ Config ก่อนเริ่มงาน
 */
CONFIG.validateSystemIntegrity = function() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var errors = [];

  // 1. Check Sheets Existence (เพิ่มการตรวจสอบ SHEET_POSTAL)
  var requiredSheets = [this.SHEET_NAME, this.MAPPING_SHEET, SCG_CONFIG.SHEET_INPUT, this.SHEET_POSTAL];
  requiredSheets.forEach(function(name) {
    if (!ss.getSheetByName(name)) errors.push("Missing Sheet: " + name);
  });

  // 2. Check API Key
  try {
    var key = this.GEMINI_API_KEY; 
    if (!key || key.length < 20) errors.push("Invalid Gemini API Key format");
  } catch (e) {
    errors.push("Gemini API Key is not set in ScriptProperties. Please run setupEnvironment() first.");
  }

  // 3. Report
  if (errors.length > 0) {
    var msg = "⚠️ SYSTEM INTEGRITY FAILED:\n" + errors.join("\n");
    console.error(msg); // [MODIFIED]: ใช้ console.error สำหรับ Enterprise Monitoring
    throw new Error(msg);
  } else {
    console.log("✅ System Integrity: OK"); // [MODIFIED]: ใช้ console.log
    return true;
  }
};


