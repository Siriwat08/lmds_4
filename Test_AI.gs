/**
 * VERSION: 000
 * 🧪 Test & Debug: AI Capabilities (Enterprise Debugging Suite)
 * Version: 4.0 Compatible with System V4.0
 * ---------------------------------------------
 * [PRESERVED]: Manual triggers, Connection test, and Row Reset logic.
 * [MODIFIED v4.0]: Upgraded debug_ResetSelectedRowsAI to clear both [AI] and [Agent_V4] tags.
 * [MODIFIED v4.0]: Replaced legacy Browser.msgBox with SpreadsheetApp.getUi() for stability.
 * [ADDED v4.0]: debug_TestTier4SmartResolution() to manually trigger the new Tier 4 AI.
 * Author: Elite Logistics Architect
 */

// ==========================================
// 1. MANUAL TRIGGERS (AI BATCH RUNNERS)
// ==========================================

/**
 * 🚀 Manual Trigger: สั่งรัน AI ทันที (AutoPilot Batch - 20 แถว)
 * ใช้สำหรับทดสอบการทำงาน หรือเร่งด่วนเก็บตกข้อมูล (สร้าง Index)
 */
function forceRunAI_Now() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    // 1. Dependency Check
    if (typeof processAIIndexing_Batch !== 'function') {
      throw new Error("Critical: ไม่พบฟังก์ชัน 'processAIIndexing_Batch' ใน Service_AutoPilot.gs");
    }

    // 2. Execution
    ss.toast("🚀 กำลังเริ่มระบบ AI Indexing (Batch Mode)...", "Debug System", 10);
    console.info("[Debug] Manual Trigger: processAIIndexing_Batch");

    // เรียกฟังก์ชันจาก Service_AutoPilot
    processAIIndexing_Batch();

    ui.alert(
      "✅ สั่งงานเรียบร้อย!\n" +
      "ระบบได้ประมวลผลข้อมูลชุดล่าสุดเสร็จสิ้น\n" +
      "กรุณาตรวจสอบคอลัมน์ Normalized ใน Database ว่ามี Tag '[AI]' หรือไม่"
    );

  } catch (e) {
    console.error("[Debug Error] forceRunAI_Now: " + e.message);
    ui.alert("❌ Error: " + e.message);
  }
}

/**
 * 🧠 [NEW v4.0] Manual Trigger: ทดสอบ Tier 4 Smart Resolution ทันที
 */
function debug_TestTier4SmartResolution() {
  var ui = SpreadsheetApp.getUi();
  try {
    if (typeof resolveUnknownNamesWithAI !== 'function') {
      throw new Error("Critical: ไม่พบฟังก์ชัน 'resolveUnknownNamesWithAI' ใน Service_Agent.gs");
    }

    var response = ui.alert("🧠 ยืนยันรันทดสอบ Tier 4", "ต้องการดึงรายชื่อที่ไม่มีพิกัดจากหน้า SCG Data\nไปให้ Gemini วิเคราะห์จับคู่กับ Master Database เลยหรือไม่?", ui.ButtonSet.YES_NO);

    if (response == ui.Button.YES) {
      console.info("[Debug] Manual Trigger: resolveUnknownNamesWithAI");
      resolveUnknownNamesWithAI();
    }
  } catch (e) {
    console.error("[Debug Error] Tier 4 Test: " + e.message);
    ui.alert("❌ Error: " + e.message);
  }
}

// ==========================================
// 2. API CONNECTION TESTING
// ==========================================

/**
 * 📡 Connection Test: ทดสอบคุยกับ Gemini (ไม่ยุ่งกับ Database)
 * ใช้เช็คว่า API Key ใช้งานได้จริงหรือไม่
 */
function debugGeminiConnection() {
  var ui = SpreadsheetApp.getUi();
  var apiKey;

  try {
    // [MODIFIED v4.0] Safe Getter Extraction
    apiKey = CONFIG.GEMINI_API_KEY;
  } catch (e) {
    ui.alert("❌ API Key Error", "กรุณาตั้งค่า API Key ผ่าน Setup_Security.gs ก่อนครับ\n(" + e.message + ")", ui.ButtonSet.OK);
    return;
  }

  var testWord = "SCG (Bang Sue Branch)";
  ui.alert("📡 กำลังทดสอบส่งข้อความหา Gemini...\nInput: " + testWord);

  try {
    console.info("[Debug] Pinging Gemini API...");

    // Fallback: ยิง API เองเพื่อ Isolate ปัญหา (จะได้รู้ว่าผิดที่ฟังก์ชันหรือ API)
    var model = (typeof CONFIG !== 'undefined' && CONFIG.AI_MODEL) ? CONFIG.AI_MODEL : "gemini-1.5-flash";
    var url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;
    var payload = {
      "contents": [{ "parts": [{ "text": `Hello Gemini, test connection. Say "Connection Success" and reply with Thai translation of ${testWord}` }] }]
    };
    var options = {
      "method": "post", "contentType": "application/json",
      "payload": JSON.stringify(payload), "muteHttpExceptions": true
    };

    var res = UrlFetchApp.fetch(url, options);

    if (res.getResponseCode() === 200) {
      var json = JSON.parse(res.getContentText());
      var text = (json.candidates && json.candidates[0].content) ? json.candidates[0].content.parts[0].text : "No Text Data";
      ui.alert("✅ API Ping Success!\n\nResponse:\n" + text);
      console.log("[Debug] Gemini API Connection: OK");
    } else {
      ui.alert("❌ API Error: " + res.getContentText());
      console.error("[Debug] Gemini API Error: " + res.getContentText());
    }

  } catch (e) {
    ui.alert("❌ Connection Failed: " + e.message);
    console.error("[Debug] Connection Failed: " + e.message);
  }
}

// ==========================================
// 3. ROW MANIPULATION (FOR RE-RUNNING AI)
// ==========================================

/**
 * 🔄 Reset AI Tags: ล้าง Tag ระบบ AI เพื่อให้รันใหม่ (เฉพาะแถวที่เลือก)
 * [MODIFIED v4.0]: ล้างทั้ง [AI] และ [Agent_V4]
 */
function debug_ResetSelectedRowsAI() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var sheet = ss.getActiveSheet();

  if (sheet.getName() !== CONFIG.SHEET_NAME) {
    ui.alert("⚠️ System Note", "กรุณาไฮไลต์เลือก Cell ในชีต Database เท่านั้นครับ", ui.ButtonSet.OK);
    return;
  }

  var range = sheet.getActiveRange();
  var startRow = range.getRow();
  var numRows = range.getNumRows();

  // ใช้ C_IDX ถ้ามี หรือ Fallback
  var colIndex = (typeof CONFIG !== 'undefined' && CONFIG.COL_NORMALIZED) ? CONFIG.COL_NORMALIZED : 6;

  var targetRange = sheet.getRange(startRow, colIndex, numRows, 1);
  var values = targetRange.getValues();

  var resetCount = 0;
  for (var i = 0; i < values.length; i++) {
    var val = values[i][0] ? values[i][0].toString() : "";

    // ตรวจหา Tag ของ AI (ทั้งระบบเก่าและใหม่)
    if (val.indexOf("[AI]") !== -1 || val.indexOf("[Agent_") !== -1) {

      // ลบ Tags ออก (ทิ้งคำที่ AI เติมไว้ได้ หรือจะลบให้ว่างเลยก็ได้)
      // V4.0 เราเลือกลบแค่ตัว Tag ออกเพื่อให้ AI เข้ามาประมวลผลซ้ำ
      var cleanedVal = val
        .replace(" [AI]", "").replace("[AI]", "")
        .replace(/\[Agent_.*?\]/g, "") // ลบ Tag รูปแบบ [Agent_xxx] ทั้งหมด
        .trim();

      values[i][0] = cleanedVal;
      resetCount++;
    }
  }

  if (resetCount > 0) {
    targetRange.setValues(values);
    ss.toast("🔄 Reset AI Status เรียบร้อย " + resetCount + " แถว", "Debug", 5);
    console.log(`[Debug] Reset AI tags for ${resetCount} rows.`);
  } else {
    ss.toast("ℹ️ ไม่พบรายการที่มี Tag AI ในส่วนที่คุณไฮไลต์เลือกไว้", "Debug", 5);
  }
}


