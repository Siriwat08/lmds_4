/**
 * VERSION: 000
 * 🖥️ MODULE: Menu UI Interface
 * Version: 4.1 Enterprise Edition (UI Text Fix)
 * ---------------------------------------------------
 * [FIXED v4.1]: Dynamic UI Alert pulling exact sheet names from CONFIG.
 * Author: Elite Logistics Architect
 */

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  
  // =================================================================
  // 🚛 เมนูชุดที่ 1: ระบบจัดการ Master Data (Operation)
  // =================================================================
  ui.createMenu('🚛 1. ระบบจัดการ Master Data')
      .addItem('1️⃣ ดึงลูกค้าใหม่ (Sync New Data)', 'syncNewDataToMaster_UI')
      .addItem('2️⃣ เติมข้อมูลพิกัด/ที่อยู่ (ทีละ 50)', 'updateGeoData_SmartCache')
      .addItem('3️⃣ จัดกลุ่มชื่อซ้ำ (Clustering)', 'autoGenerateMasterList_Smart')
      .addItem('🧠 4️⃣ ส่งชื่อแปลกให้ AI วิเคราะห์ (Smart Resolution)', 'runAIBatchResolver_UI')
      .addSeparator()
      .addItem('🚀 5️⃣ Deep Clean (ตรวจสอบความสมบูรณ์)', 'runDeepCleanBatch_100')
      .addItem('🔄 รีเซ็ตความจำปุ่ม 5 (เริ่มแถว 2 ใหม่)', 'resetDeepCleanMemory_UI')
      .addSeparator()
      .addItem('✅ 6️⃣ จบงาน (Finalize & Move to Mapping)', 'finalizeAndClean_UI')
      .addSeparator()
      .addSubMenu(ui.createMenu('🛠️ Admin & Repair Tools')
    .addItem('🔑 สร้าง UUID ให้ครบทุกแถว', 'assignMissingUUIDs')
    .addItem('🚑 ซ่อมแซม NameMapping (L3)', 'repairNameMapping_UI')
    .addSeparator()
    .addItem('🔍 ค้นหาพิกัดซ้ำซ้อน (Hidden Duplicates)', 'findHiddenDuplicates')
    .addItem('📊 ตรวจสอบคุณภาพข้อมูล (Quality Report)', 'showQualityReport_UI')
    .addItem('🔄 คำนวณ Quality ใหม่ทั้งหมด', 'recalculateAllQuality')
    .addItem('🎯 คำนวณ Confidence ใหม่ทั้งหมด', 'recalculateAllConfidence')
    .addSeparator()
    .addItem('🗂️ Initialize Record Status', 'initializeRecordStatus')  // ← เพิ่ม
    .addItem('🔀 Merge UUID ซ้ำซ้อน', 'mergeDuplicates_UI')            // ← เพิ่ม
    .addItem('📋 ดูสถานะ Record ทั้งหมด', 'showRecordStatusReport')    // ← เพิ่ม
)
.addToUi();
  // =================================================================
  // 📦 เมนูชุดที่ 2: เมนูพิเศษ SCG (Daily Operation)
  // =================================================================
 ui.createMenu('📦 2. เมนูพิเศษ SCG') 
  .addItem('📥 1. โหลดข้อมูล Shipment (+E-POD)', 'fetchDataFromSCGJWD')
  .addItem('🟢 2. อัปเดตพิกัด + อีเมลพนักงาน', 'applyMasterCoordinatesToDailyJob')
  .addSeparator()
  .addSubMenu(ui.createMenu('📍 GPS Queue Management')
    .addItem('🔄 1. Sync GPS จากคนขับ → Queue', 'syncNewDataToMaster_UI')
    .addItem('✅ 2. อนุมัติรายการที่ติ๊กแล้ว', 'applyApprovedFeedback')
    .addItem('📊 3. ดูสถิติ Queue', 'showGPSQueueStats')
    .addSeparator()
    .addItem('🛠️ สร้างชีต GPS_Queue ใหม่', 'createGPSQueueSheet')
  )
  .addSeparator()
  .addSubMenu(ui.createMenu('🧹 เมนูล้างข้อมูล (Dangerous Zone)')
    .addItem('⚠️ ล้างเฉพาะชีต Data', 'clearDataSheet_UI')
    .addItem('⚠️ ล้างเฉพาะชีต Input', 'clearInputSheet_UI')
    .addItem('⚠️ ล้างเฉพาะชีต สรุป_เจ้าของสินค้า', 'clearSummarySheet_UI')
    .addItem('🔥 ล้างทั้งหมด (Input + Data + สรุป)', 'clearAllSCGSheets_UI')
  )
  .addToUi();
  // =================================================================
  // 🤖 เมนูชุดที่ 3: ระบบอัตโนมัติ (Automation)
  // =================================================================
  ui.createMenu('🤖 3. ระบบอัตโนมัติ')
  .addItem('▶️ เปิดระบบช่วยเหลืองาน (Auto-Pilot)', 'START_AUTO_PILOT')
  .addItem('⏹️ ปิดระบบช่วยเหลือ', 'STOP_AUTO_PILOT')
  .addItem('👋 ปลุก AI Agent ทำงานทันที', 'WAKE_UP_AGENT')
  .addSeparator()
  .addSubMenu(ui.createMenu('🧪 Debug & Test Tools')
    .addItem('🚀 รัน AI Indexing ทันที', 'forceRunAI_Now')
    .addItem('🧠 ทดสอบ Tier 4 AI Resolution', 'debug_TestTier4SmartResolution')
    .addItem('📡 ทดสอบ Gemini Connection', 'debugGeminiConnection')
    .addItem('🔄 ล้าง AI Tags (แถวที่เลือก)', 'debug_ResetSelectedRowsAI')
    .addSeparator()
    .addItem('🔁 Reset SYNC_STATUS (ทดสอบ)', 'resetSyncStatus')
  )
  .addToUi();
  // =================================================================
  // ⚙️ เมนูชุดที่ 4: System Admin
  // =================================================================
  ui.createMenu('⚙️ System Admin')
  .addItem('🏥 ตรวจสอบสถานะระบบ (Health Check)', 'runSystemHealthCheck')
  .addItem('🧹 ล้าง Backup เก่า (>30 วัน)', 'cleanupOldBackups')
  .addItem('📊 เช็คปริมาณข้อมูล (Cell Usage)', 'checkSpreadsheetHealth')
  .addSeparator()
  .addSubMenu(ui.createMenu('🔬 System Diagnostic')
  .addItem('🛡️ ตรวจสอบ Schema ทุกชีต', 'runFullSchemaValidation') // ← เพิ่มบรรทัดนี้
  .addItem('🔍 ตรวจสอบ Engine (Phase 1)', 'RUN_SYSTEM_DIAGNOSTIC')
  .addItem('🕵️ ตรวจสอบชีต (Phase 2)', 'RUN_SHEET_DIAGNOSTIC')
  .addItem('🧹 ล้าง Postal Cache', 'clearPostalCache_UI')
  .addItem('🧹 ล้าง Search Cache', 'clearSearchCache_UI')
  )
  .addSeparator()
  .addItem('🔔 ตั้งค่า LINE Notify', 'setupLineToken')
  .addItem('✈️ ตั้งค่า Telegram Notify', 'setupTelegramConfig')
  .addItem('🔐 ตั้งค่า API Key (Setup)', 'setupEnvironment')
  .addToUi();
}

// =================================================================
// 🛡️ SAFETY WRAPPERS
// =================================================================

/**
 * Wrapper: ยืนยันก่อนดึงข้อมูลลูกค้าใหม่
 * [FIXED v4.1]: ปรับข้อความให้ดึงชื่อจากตัวแปร Config จริงๆ
 */
function syncNewDataToMaster_UI() {
  var ui = SpreadsheetApp.getUi();
  var sourceName = (typeof CONFIG !== 'undefined' && CONFIG.SOURCE_SHEET) ? CONFIG.SOURCE_SHEET : 'ชีตนำเข้า';
  var dbName = (typeof CONFIG !== 'undefined' && CONFIG.SHEET_NAME) ? CONFIG.SHEET_NAME : 'Database';
  
  var result = ui.alert(
    'ยืนยันการดึงข้อมูลใหม่?',
    'ระบบจะดึงรายชื่อลูกค้าจากชีต "' + sourceName + '"\nมาเพิ่มต่อท้ายในชีต "' + dbName + '"\n(เฉพาะรายชื่อที่ยังไม่เคยมีในระบบ)\n\nคุณต้องการดำเนินการต่อหรือไม่?',
    ui.ButtonSet.YES_NO
  );
  if (result == ui.Button.YES) {
    syncNewDataToMaster();
  }
}

function runAIBatchResolver_UI() {
  var ui = SpreadsheetApp.getUi();
  var batchSize = (typeof CONFIG !== 'undefined' && CONFIG.AI_BATCH_SIZE) ? CONFIG.AI_BATCH_SIZE : 20;
  
  var result = ui.alert(
    '🧠 ยืนยันการรัน AI Smart Resolution?',
    'ระบบจะรวบรวมชื่อที่ยังหาพิกัดไม่เจอ/ไม่รู้จัก (สูงสุด ' + batchSize + ' รายการ)\nส่งให้ Gemini AI วิเคราะห์และจับคู่กับ Database อัตโนมัติ\n\nต้องการเริ่มเลยหรือไม่?',
    ui.ButtonSet.YES_NO
  );
  
  if (result == ui.Button.YES) {
    if (typeof resolveUnknownNamesWithAI === 'function') {
       resolveUnknownNamesWithAI();
    } else {
       ui.alert(
         '⚠️ System Note', 
         'ฟังก์ชัน AI (Service_Agent.gs) กำลังอยู่ระหว่างการติดตั้ง (Coming soon!)\nกรุณารออัปเดตโมดูลถัดไปครับ', 
         ui.ButtonSet.OK
       );
    }
  }
}

function finalizeAndClean_UI() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert(
    '⚠️ ยืนยันการจบงาน (Finalize)?',
    'รายการที่ติ๊กถูก "Verified" จะถูกย้ายไปยัง NameMapping และลบออกจาก Database\nข้อมูลต้นฉบับจะถูก Backup ไว้\n\nยืนยันหรือไม่?',
    ui.ButtonSet.OK_CANCEL
  );
  if (result == ui.Button.OK) {
    finalizeAndClean_MoveToMapping();
  }
}

function resetDeepCleanMemory_UI() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert(
    'ยืนยันการรีเซ็ต?',
    'ระบบจะเริ่มตรวจสอบ Deep Clean ตั้งแต่แถวแรกใหม่\nใช้ในกรณีที่ต้องการ Re-check ข้อมูลทั้งหมด',
    ui.ButtonSet.YES_NO
  );
  if (result == ui.Button.YES) {
    resetDeepCleanMemory();
  }
}

function clearDataSheet_UI() {
  confirmAction('ล้างชีต Data', 'ข้อมูลผลลัพธ์ทั้งหมดจะหายไป', clearDataSheet);
}

function clearInputSheet_UI() {
  confirmAction('ล้างชีต Input', 'ข้อมูลนำเข้า (Shipment) ทั้งหมดจะหายไป', clearInputSheet);
}

function repairNameMapping_UI() {
  confirmAction('ซ่อมแซม NameMapping', 'ระบบจะลบแถวซ้ำและเติม UUID ให้ครบ', repairNameMapping_Full);
}

function confirmAction(title, message, callbackFunction) {
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert(title, message, ui.ButtonSet.YES_NO);
  if (result == ui.Button.YES) {
    callbackFunction();
  }
}

function runSystemHealthCheck() {
  var ui = SpreadsheetApp.getUi();
  try {
    if (typeof CONFIG !== 'undefined' && CONFIG.validateSystemIntegrity) {
      CONFIG.validateSystemIntegrity(); 
      ui.alert(
        "✅ System Health: Excellent\n",
        "ระบบพร้อมทำงานสมบูรณ์ครับ!\n- โครงสร้างชีตครบถ้วน\n- เชื่อมต่อ API (Gemini) พร้อมใช้งาน",
        ui.ButtonSet.OK
      );
    } else {
      ui.alert("⚠️ System Warning", "Config check skipped (CONFIG.validateSystemIntegrity ไม่ทำงาน)", ui.ButtonSet.OK);
    }
  } catch (e) {
    ui.alert("❌ System Health: FAILED", e.message, ui.ButtonSet.OK);
  }
}

function showQualityReport_UI() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var ui  = SpreadsheetApp.getUi();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  
  var lastRow = getRealLastRow_(sheet, CONFIG.COL_NAME);
  if (lastRow < 2) {
    ui.alert("ℹ️ Database ว่างเปล่าครับ");
    return;
  }
  
  var data = sheet.getRange(2, 1, lastRow - 1, 17).getValues();
  
  var stats = {
    total:      0,
    noCoord:    0,
    noProvince: 0,
    noUUID:     0,
    noAddr:     0,
    notVerified:0,
    highQ:      0,
    midQ:       0,
    lowQ:       0
  };
  
  data.forEach(function(row) {
    if (!row[CONFIG.C_IDX.NAME]) return;
    stats.total++;
    
    var lat = parseFloat(row[CONFIG.C_IDX.LAT]);
    var lng = parseFloat(row[CONFIG.C_IDX.LNG]);
    var q   = parseFloat(row[CONFIG.C_IDX.QUALITY]);
    
    if (isNaN(lat) || isNaN(lng))          stats.noCoord++;
    if (!row[CONFIG.C_IDX.PROVINCE])       stats.noProvince++;
    if (!row[CONFIG.C_IDX.UUID])           stats.noUUID++;
    if (!row[CONFIG.C_IDX.GOOGLE_ADDR])    stats.noAddr++;
    if (row[CONFIG.C_IDX.VERIFIED] !== true) stats.notVerified++;
    
    if (q >= 80)      stats.highQ++;
    else if (q >= 50) stats.midQ++;
    else              stats.lowQ++;
  });
  
  var msg =
    "📊 Database Quality Report\n" +
    "━━━━━━━━━━━━━━━━━━━━━━━━━\n" +
    "📝 ทั้งหมด: " + stats.total + " แถว\n\n" +
    
    "🎯 Quality Score:\n" +
    "🟢 ≥ 80% (ดีมาก):      " + stats.highQ + " แถว\n" +
    "🟡 50-79% (ดีพอใช้):   " + stats.midQ  + " แถว\n" +
    "🔴 < 50% (ต้องปรับปรุง): " + stats.lowQ  + " แถว\n\n" +
    
    "⚠️ ข้อมูลที่ขาดหาย:\n" +
    "📍 ไม่มีพิกัด:     " + stats.noCoord    + " แถว\n" +
    "🏙️ ไม่มี Province: " + stats.noProvince + " แถว\n" +
    "🗺️ ไม่มีที่อยู่:   " + stats.noAddr     + " แถว\n" +
    "🔑 ไม่มี UUID:     " + stats.noUUID     + " แถว\n" +
    "✅ ยังไม่ Verified: " + stats.notVerified + " แถว\n\n" +
    
    "━━━━━━━━━━━━━━━━━━━━━━━━━\n" +
    "💡 แนะนำ:\n";
    
  if (stats.noCoord > 0) {
    msg += "• รัน 'เติมข้อมูลพิกัด (ทีละ 50)' เพื่อเติมพิกัด\n";
  }
  if (stats.noProvince > 0) {
    msg += "• รัน 'Deep Clean' เพื่อเติม Province/District\n";
  }
  if (stats.noUUID > 0) {
    msg += "• รัน 'สร้าง UUID ให้ครบทุกแถว'\n";
  }
  if (stats.lowQ > 0) {
    msg += "• ตรวจสอบ " + stats.lowQ + " แถวที่ Quality ต่ำ\n";
  }
    
  ui.alert(msg);
}

function clearPostalCache_UI() {
  var ui = SpreadsheetApp.getUi();
  try {
    // [FIXED v4.1] เรียกผ่านฟังก์ชันใน Service_GeoAddr.gs
    // เพราะ _POSTAL_CACHE อยู่ในไฟล์นั้น แก้ตรงๆ จากไฟล์อื่นไม่ได้
    clearPostalCache();
    ui.alert(
      "✅ ล้าง Postal Cache เรียบร้อย!\n\n" +
      "ครั้งถัดไปที่ระบบค้นหารหัสไปรษณีย์\n" +
      "จะโหลดข้อมูลใหม่จากชีต PostalRef ครับ"
    );
    console.log("[Cache] Postal Cache cleared by user.");
  } catch(e) {
    ui.alert("❌ Error: " + e.message);
  }
}

function clearSearchCache_UI() {
  var ui = SpreadsheetApp.getUi();
  try {
    clearSearchCache();
    ui.alert(
      "✅ ล้าง Search Cache เรียบร้อย!\n\n" +
      "ครั้งถัดไปที่ค้นหาผ่าน WebApp\n" +
      "จะโหลด NameMapping ใหม่จากชีตครับ"
    );
    console.log("[Cache] Search Cache cleared by user.");
  } catch(e) {
    ui.alert("❌ Error: " + e.message);
  }
}

