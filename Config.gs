/**
 * VERSION: 4.2 — Phase A
 * [Phase A] เพิ่ม DB_TOTAL_COLS, header constants
 */

var CONFIG = {
  SHEET_NAME:    "Database",
  MAPPING_SHEET: "NameMapping",
  SOURCE_SHEET:  "SCGนครหลวงJWDภูมิภาค",
  SHEET_POSTAL:  "PostalRef",

  // [Phase A NEW] Schema Width Constants
  DB_TOTAL_COLS:        22,
  DB_LEGACY_COLS:       17,
  MAP_TOTAL_COLS:       5,
  GPS_QUEUE_TOTAL_COLS: 9,
  DATA_TOTAL_COLS:      29,

  // [Phase A NEW] Header Arrays กลาง
  DB_REQUIRED_HEADERS: {
    1: "NAME", 2: "LAT", 3: "LNG", 11: "UUID",
    15: "QUALITY", 16: "CREATED", 17: "UPDATED",
    18: "Coord_Source", 19: "Coord_Confidence",
    20: "Coord_Last_Updated",
    21: "Record_Status",
    22: "Merged_To_UUID"
  },

  MAP_REQUIRED_HEADERS: {
    1: "Variant_Name", 2: "Master_UID",
    3: "Confidence_Score", 4: "Mapped_By", 5: "Timestamp"
  },

  GPS_QUEUE_REQUIRED_HEADERS: {
    1: "Timestamp", 2: "ShipToName", 3: "UUID_DB",
    4: "LatLng_Driver", 5: "LatLng_DB", 6: "Diff_Meters",
    7: "Reason", 8: "Approve", 9: "Reject"
  },

  get GEMINI_API_KEY() {
    var key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!key) throw new Error(
      "CRITICAL ERROR: GEMINI_API_KEY is not set. Please run setupEnvironment() first."
    );
    return key;
  },
  USE_AI_AUTO_FIX: true,
  AI_MODEL:       "gemini-1.5-flash",
  AI_BATCH_SIZE:  20,

  DEPOT_LAT: 14.164688,
  DEPOT_LNG: 100.625354,

  DISTANCE_THRESHOLD_KM: 0.05,
  BATCH_LIMIT:            50,
  DEEP_CLEAN_LIMIT:       100,
  API_MAX_RETRIES:        3,
  API_TIMEOUT_MS:         30000,
  CACHE_EXPIRATION:       21600,

  COL_NAME: 1,       COL_LAT: 2,        COL_LNG: 3,
  COL_SUGGESTED: 4,  COL_CONFIDENCE: 5, COL_NORMALIZED: 6,
  COL_VERIFIED: 7,   COL_SYS_ADDR: 8,   COL_ADDR_GOOG: 9,
  COL_DIST_KM: 10,   COL_UUID: 11,      COL_PROVINCE: 12,
  COL_DISTRICT: 13,  COL_POSTCODE: 14,  COL_QUALITY: 15,
  COL_CREATED: 16,   COL_UPDATED: 17,
  COL_COORD_SOURCE:       18,
  COL_COORD_CONFIDENCE:   19,
  COL_COORD_LAST_UPDATED: 20,
  COL_RECORD_STATUS:      21,
  COL_MERGED_TO_UUID:     22,

  MAP_COL_VARIANT: 1, MAP_COL_UID: 2,   MAP_COL_CONFIDENCE: 3,
  MAP_COL_MAPPED_BY: 4, MAP_COL_TIMESTAMP: 5,

  get C_IDX() {
    return {
      NAME: this.COL_NAME - 1,           LAT: this.COL_LAT - 1,
      LNG: this.COL_LNG - 1,             SUGGESTED: this.COL_SUGGESTED - 1,
      CONFIDENCE: this.COL_CONFIDENCE - 1, NORMALIZED: this.COL_NORMALIZED - 1,
      VERIFIED: this.COL_VERIFIED - 1,   SYS_ADDR: this.COL_SYS_ADDR - 1,
      GOOGLE_ADDR: this.COL_ADDR_GOOG - 1, DIST_KM: this.COL_DIST_KM - 1,
      UUID: this.COL_UUID - 1,           PROVINCE: this.COL_PROVINCE - 1,
      DISTRICT: this.COL_DISTRICT - 1,   POSTCODE: this.COL_POSTCODE - 1,
      QUALITY: this.COL_QUALITY - 1,     CREATED: this.COL_CREATED - 1,
      UPDATED: this.COL_UPDATED - 1,
      COORD_SOURCE:       this.COL_COORD_SOURCE - 1,
      COORD_CONFIDENCE:   this.COL_COORD_CONFIDENCE - 1,
      COORD_LAST_UPDATED: this.COL_COORD_LAST_UPDATED - 1,
      RECORD_STATUS:      this.COL_RECORD_STATUS - 1,
      MERGED_TO_UUID:     this.COL_MERGED_TO_UUID - 1
    };
  },

  get MAP_IDX() {
    return {
      VARIANT:    this.MAP_COL_VARIANT - 1,
      UID:        this.MAP_COL_UID - 1,
      CONFIDENCE: this.MAP_COL_CONFIDENCE - 1,
      MAPPED_BY:  this.MAP_COL_MAPPED_BY - 1,
      TIMESTAMP:  this.MAP_COL_TIMESTAMP - 1
    };
  }
};

const SCG_CONFIG = {
  SHEET_DATA:     'Data',
  SHEET_INPUT:    'Input',
  SHEET_EMPLOYEE: 'ข้อมูลพนักงาน',
  API_URL:        'https://fsm.scgjwd.com/Monitor/SearchDelivery',
  INPUT_START_ROW: 4,
  COOKIE_CELL:    'B1',
  SHIPMENT_STRING_CELL: 'B3',
  SHEET_MASTER_DB: 'Database',
  SHEET_MAPPING:   'NameMapping',
  SHEET_GPS_QUEUE: 'GPS_Queue',
  GPS_THRESHOLD_METERS: 50,
  SRC_IDX: {
    NAME: 12, LAT: 14, LNG: 15,
    SYS_ADDR: 18, DIST: 23, GOOG_ADDR: 24
  },
  SRC_IDX_SYNC_STATUS: 37,
  SYNC_STATUS_DONE: "SYNCED",
  JSON_MAP: {
    SHIPMENT_NO:   'shipmentNo',
    CUSTOMER_NAME: 'customerName',
    DELIVERY_DATE: 'deliveryDate'
  }
};

CONFIG.validateSystemIntegrity = function() {
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var errors = [];
  [this.SHEET_NAME, this.MAPPING_SHEET,
   SCG_CONFIG.SHEET_INPUT, this.SHEET_POSTAL].forEach(function(name) {
    if (!ss.getSheetByName(name)) errors.push("Missing Sheet: " + name);
  });
  try {
    var key = this.GEMINI_API_KEY;
    if (!key || key.length < 20) errors.push("Invalid Gemini API Key format");
  } catch(e) {
    errors.push("Gemini API Key not set. Run setupEnvironment() first.");
  }
  if (errors.length > 0) {
    var msg = "⚠️ SYSTEM INTEGRITY FAILED:\n" + errors.join("\n");
    console.error(msg);
    throw new Error(msg);
  }
  console.log("✅ System Integrity: OK");
  return true;
};
