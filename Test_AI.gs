/**
 * VERSION : 4.2 — Phase D
 * [Phase D] เพิ่ม testRetrieveCandidates_(), testAIResponseValidation_()
 */

function forceRunAI_Now() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    if (typeof processAIIndexing_Batch !== 'function') {
      throw new Error("ไม่พบฟังก์ชัน 'processAIIndexing_Batch'");
    }
    ss.toast("🚀 กำลังเริ่ม AI Indexing...", "Debug", 10);
    processAIIndexing_Batch();
    ui.alert("✅ สั่งงานเรียบร้อย!\nตรวจ Column Normalized ว่ามี Tag '[AI]' หรือไม่");
  } catch(e) {
    ui.alert("❌ Error: " + e.message);
  }
}

function debug_TestTier4SmartResolution() {
  var ui = SpreadsheetApp.getUi();
  try {
    if (typeof resolveUnknownNamesWithAI !== 'function') {
      throw new Error("ไม่พบฟังก์ชัน 'resolveUnknownNamesWithAI'");
    }
    var response = ui.alert("🧠 ยืนยันรัน Tier 4?",
      "ดึงรายชื่อที่ไม่มีพิกัดจาก SCG Data ส่งให้ AI วิเคราะห์",
      ui.ButtonSet.YES_NO);
    if (response === ui.Button.YES) resolveUnknownNamesWithAI();
  } catch(e) {
    ui.alert("❌ Error: " + e.message);
  }
}

function debugGeminiConnection() {
  var ui = SpreadsheetApp.getUi();
  var apiKey;
  try { apiKey = CONFIG.GEMINI_API_KEY; }
  catch(e) { ui.alert("❌ API Key Error", e.message, ui.ButtonSet.OK); return; }

  try {
    var model    = CONFIG.AI_MODEL || "gemini-1.5-flash";
    var url      = "https://generativelanguage.googleapis.com/v1beta/models/" +
                   model + ":generateContent?key=" + apiKey;
    var payload  = { "contents": [{ "parts": [{ "text": "Say: Connection OK. Prompt-Version: " + AI_CONFIG.PROMPT_VERSION }] }] };
    var options  = { "method": "post", "contentType": "application/json",
                     "payload": JSON.stringify(payload), "muteHttpExceptions": true };
    var res      = UrlFetchApp.fetch(url, options);

    if (res.getResponseCode() === 200) {
      var json = JSON.parse(res.getContentText());
      var text = (json.candidates && json.candidates[0].content)
        ? json.candidates[0].content.parts[0].text
        : "No Text";
      ui.alert("✅ API Connection OK!\n\nResponse:\n" + text);
    } else {
      ui.alert("❌ API Error: " + res.getContentText());
    }
  } catch(e) {
    ui.alert("❌ Connection Failed: " + e.message);
  }
}

function debug_ResetSelectedRowsAI() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var ui    = SpreadsheetApp.getUi();
  var sheet = ss.getActiveSheet();

  if (sheet.getName() !== CONFIG.SHEET_NAME) {
    ui.alert("⚠️ กรุณาไฮไลต์ Cell ในชีต Database เท่านั้นครับ");
    return;
  }

  var range      = sheet.getActiveRange();
  var startRow   = range.getRow();
  var numRows    = range.getNumRows();
  var colIndex   = CONFIG.COL_NORMALIZED;
  var targetRange = sheet.getRange(startRow, colIndex, numRows, 1);
  var values     = targetRange.getValues();
  var resetCount = 0;

  for (var i = 0; i < values.length; i++) {
    var val = values[i][0] ? values[i][0].toString() : "";
    // [Phase D] ลบทั้ง [AI] และ prompt version tag
    if (val.indexOf(AI_CONFIG.TAG_AI) !== -1 || val.indexOf("[Agent_") !== -1) {
      val = val
        .replace(/\s*\[AI\]/g, "")
        .replace(/\s*\[Agent_.*?\]/g, "")
        .replace(/\s*\[v\d+\.\d+\]/g, "")
        .trim();
      values[i][0] = val;
      resetCount++;
    }
  }

  if (resetCount > 0) {
    targetRange.setValues(values);
    ss.toast("🔄 Reset AI tags: " + resetCount + " แถว", "Debug", 5);
  } else {
    ss.toast("ℹ️ ไม่พบ AI tags ในส่วนที่เลือก", "Debug", 5);
  }
}

// ==========================================
// [Phase D NEW] TEST HELPERS
// ==========================================

/**
 * testRetrieveCandidates_()
 * ทดสอบว่า retrieval คัด candidates ได้ถูกต้องก่อนส่ง AI
 */
function testRetrieveCandidates() {
  var ui = SpreadsheetApp.getUi();

  var response = ui.prompt(
    "🔍 Test Retrieval",
    "ใส่ชื่อที่ต้องการทดสอบ:",
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  var testName = response.getResponseText().trim();
  if (!testName) return;

  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var dbSheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  var lastRow = dbSheet.getLastRow();
  if (lastRow < 2) { ui.alert("ℹ️ Database ว่างเปล่า"); return; }

  var dbRows   = dbSheet.getRange(2, 1, lastRow - 1, CONFIG.DB_TOTAL_COLS).getValues();
  var mapRows  = loadNameMappingRows_();
  var results  = retrieveCandidateMasters_(testName, dbRows, mapRows, 10);

  if (results.length === 0) {
    ui.alert("ℹ️ ไม่พบ candidates สำหรับ: '" + testName + "'");
    return;
  }

  var msg = "🔍 Retrieval Results สำหรับ: '" + testName + "'\n" +
            "━━━━━━━━━━━━━━━━━━━━━━━\n";
  results.forEach(function(r, i) {
    msg += (i + 1) + ". " + r.name + "\n   UUID: " + r.uid.substring(0, 12) + "...\n";
  });
  msg += "━━━━━━━━━━━━━━━━━━━━━━━\n";
  msg += "พบ " + results.length + " candidates\n";
  msg += "💡 นี่คือสิ่งที่จะส่งให้ AI วิเคราะห์แทน slice(0,500)";

  ui.alert(msg);
}

/**
 * testAIResponseValidation_()
 * ทดสอบ parse guard กับ response จริงจาก Gemini
 */
function testAIResponseValidation() {
  var ui = SpreadsheetApp.getUi();
  var apiKey;
  try { apiKey = CONFIG.GEMINI_API_KEY; }
  catch(e) { ui.alert("❌ API Key Error: " + e.message); return; }

  var testName = "โลตัส สาขาบางนา";
  var result   = callGeminiThinking_JSON(testName, apiKey);

  if (result) {
    ui.alert(
      "✅ Parse Guard: ผ่าน!\n\n" +
      "Input: " + testName + "\n" +
      "Output: " + result + "\n" +
      "Prompt Version: " + AI_CONFIG.PROMPT_VERSION
    );
  } else {
    ui.alert(
      "⚠️ Parse Guard: ไม่ได้ keywords\n\n" +
      "Input: " + testName + "\n" +
      "อาจเป็นเพราะ API key ปัญหา หรือ rate limit\n" +
      "ตรวจสอบ Execution Log ครับ"
    );
  }
}
