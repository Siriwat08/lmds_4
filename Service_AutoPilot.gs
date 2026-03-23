/**
 * VERSION: 4.2 — Phase D
 * [Phase D] เพิ่ม parse guard, prompt versioning, AI response logging
 * [Phase D] แยก concern: normalization (deterministic) vs AI enrichment (probabilistic)
 */

function START_AUTO_PILOT() {
  STOP_AUTO_PILOT();
  ScriptApp.newTrigger("autoPilotRoutine").timeBased().everyMinutes(10).create();
  var ui = SpreadsheetApp.getUi();
  if (ui) ui.alert("▶️ AI Auto-Pilot: ACTIVATE\nระบบสมองกลจะทำงานเบื้องหลังทุกๆ 10 นาทีครับ");
}

function STOP_AUTO_PILOT() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "autoPilotRoutine") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

function autoPilotRoutine() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    console.warn("[AutoPilot] Skipped: มี instance อื่นกำลังรันอยู่");
    return;
  }
  try {
    console.time("AutoPilot_Duration");
    console.info("[AutoPilot] 🚀 Starting routine...");

    try {
      if (typeof applyMasterCoordinatesToDailyJob === 'function') {
        var ss        = SpreadsheetApp.getActiveSpreadsheet();
        var dataSheet = ss.getSheetByName(SCG_CONFIG.SHEET_DATA);
        if (dataSheet && dataSheet.getLastRow() > 1) {
          applyMasterCoordinatesToDailyJob();
          console.log("✅ AutoPilot: SCG Sync Completed");
        }
      }
    } catch(e) { console.error("[AutoPilot] SCG Sync Error: " + e.message); }

    try { processAIIndexing_Batch(); }
    catch(e) { console.error("[AutoPilot] AI Indexing Error: " + e.message); }

    console.timeEnd("AutoPilot_Duration");
    console.info("[AutoPilot] 🏁 Routine finished.");
  } catch(e) {
    console.error("[AutoPilot] CRITICAL Error: " + e.message);
  } finally {
    lock.releaseLock();
  }
}

function processAIIndexing_Batch() {
  var apiKey;
  try { apiKey = CONFIG.GEMINI_API_KEY; }
  catch(e) { console.warn("⚠️ SKIPPED AI: " + e.message); return; }

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) return;

  var lastRow = typeof getRealLastRow_ === 'function'
    ? getRealLastRow_(sheet, CONFIG.COL_NAME)
    : sheet.getLastRow();
  if (lastRow < 2) return;

  var rangeName = sheet.getRange(2, CONFIG.COL_NAME,       lastRow - 1, 1);
  var rangeNorm = sheet.getRange(2, CONFIG.COL_NORMALIZED, lastRow - 1, 1);

  var nameValues = rangeName.getValues();
  var normValues = rangeNorm.getValues();

  var aiCount  = 0;
  var AI_LIMIT = CONFIG.AI_BATCH_SIZE || 20;
  var updated  = false;

  for (var i = 0; i < nameValues.length; i++) {
    if (aiCount >= AI_LIMIT) break;

    var name        = nameValues[i][0];
    var currentNorm = normValues[i][0];

    if (name && typeof name === 'string' &&
        (!currentNorm || currentNorm.toString().indexOf(AI_CONFIG.TAG_AI) === -1)) {

      // [Phase D] deterministic part — ไม่ขึ้นกับ AI
      var basicKey = createBasicSmartKey(name);

      // [Phase D] probabilistic part — ขึ้นกับ AI
      var aiKeywords = "";
      if (name.length > 3) {
        aiKeywords = genericRetry(function() {
          return callGeminiThinking_JSON(name, apiKey);
        }, 2);
      }

      // [Phase D] เพิ่ม prompt version tag
      var finalString = basicKey +
                        (aiKeywords ? " " + aiKeywords : "") +
                        " " + AI_CONFIG.TAG_AI +
                        " [" + AI_CONFIG.PROMPT_VERSION + "]";

      normValues[i][0] = finalString.trim();

      console.log("[AI Indexing] (" + (aiCount + 1) + "/" + AI_LIMIT + ") " +
                  name + " → " + (aiKeywords || "basic only"));
      aiCount++;
      updated = true;
    }
  }

  if (updated) {
    rangeNorm.setValues(normValues);
    console.log("✅ [AI Indexing] Batch write: " + aiCount + " records updated");
  } else {
    console.log("ℹ️ [AI Indexing] ไม่มีข้อมูลใหม่ที่ต้องประมวลผล");
  }
}

/**
 * [Phase D] callGeminiThinking_JSON()
 * เพิ่ม parse guard ป้องกัน silent fail
 */
function callGeminiThinking_JSON(customerName, apiKey) {
  try {
    var model  = CONFIG.AI_MODEL || "gemini-1.5-flash";
    var apiUrl = "https://generativelanguage.googleapis.com/v1beta/models/" +
                 model + ":generateContent?key=" + apiKey;

    var prompt =
      "Task: Analyze this Thai logistics customer name: \"" + customerName + "\"\n" +
      "Goal: Return a JSON list of search keywords, abbreviations, and common typos.\n" +
      "Requirements:\n" +
      "1. If English, provide Thai phonetics.\n" +
      "2. If Thai abbreviation, provide full text.\n" +
      "3. No generic words: Company, Limited, จำกัด, บริษัท\n" +
      "4. Max 5 keywords.\n" +
      "Prompt-Version: " + AI_CONFIG.PROMPT_VERSION + "\n" +
      "Output Format: JSON Array of Strings ONLY.";

    var payload = {
      "contents": [{ "parts": [{ "text": prompt }] }],
      "generationConfig": { "responseMimeType": "application/json" }
    };

    var response   = UrlFetchApp.fetch(apiUrl, {
      "method": "post",
      "contentType": "application/json",
      "payload": JSON.stringify(payload),
      "muteHttpExceptions": true
    });

    var statusCode = response.getResponseCode();

    // [Phase D] ตรวจ HTTP status ก่อน parse
    if (statusCode !== 200) {
      console.warn("[callGeminiThinking_JSON] HTTP " + statusCode +
                   " for '" + customerName + "': " + response.getContentText().substring(0, 100));
      return "";
    }

    var json = JSON.parse(response.getContentText());

    // [Phase D] parse guard — ตรวจ structure ก่อนใช้
    if (!json.candidates || !Array.isArray(json.candidates) || json.candidates.length === 0) {
      console.warn("[callGeminiThinking_JSON] No candidates for '" + customerName + "'");
      return "";
    }

    var content = json.candidates[0].content;
    if (!content || !content.parts || !content.parts[0] || !content.parts[0].text) {
      console.warn("[callGeminiThinking_JSON] Invalid response structure for '" + customerName + "'");
      return "";
    }

    var text     = content.parts[0].text;
    var keywords = JSON.parse(text);

    // [Phase D] ตรวจว่าเป็น array จริงๆ
    if (!Array.isArray(keywords)) {
      console.warn("[callGeminiThinking_JSON] Response is not array for '" + customerName + "'");
      return "";
    }

    return keywords.join(" ");

  } catch(e) {
    console.warn("[callGeminiThinking_JSON] Error (" + customerName + "): " + e.message);
    return "";
  }
}

function createBasicSmartKey(text) {
  if (!text) return "";
  return typeof normalizeText === 'function'
    ? normalizeText(text)
    : text.toString().toLowerCase().replace(/\s/g, "");
}
