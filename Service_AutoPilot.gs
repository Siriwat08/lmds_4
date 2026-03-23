/**
 * VERSION: 000
 * 🤖 Service: Auto Pilot (Enterprise AI Edition)
 * Version: 4.2 Clean SmartKey & Stable Fallback
 * --------------------------------------------
 * [FIXED v4.2]: Removed duplicate "tone-less" basic key to keep data clean. 
 * AI will handle typos and phonetic variations instead.
 * [FIXED v4.2]: Enforced 'gemini-1.5-flash-latest' to resolve v1beta 404 errors.
 * [PRESERVED]: Trigger management, LockService, and JSON output parsing.
 * Author: Elite Logistics Architect
 */

function START_AUTO_PILOT() {
  STOP_AUTO_PILOT(); 
  ScriptApp.newTrigger("autoPilotRoutine")
    .timeBased()
    .everyMinutes(10)
    .create();
    
  var ui = SpreadsheetApp.getUi();
  if (ui) {
    ui.alert("▶️ AI Auto-Pilot: ACTIVATE\nระบบสมองกลจะทำงานเบื้องหลังทุกๆ 10 นาทีครับ");
  }
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
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var dataSheet = ss.getSheetByName(typeof SCG_CONFIG !== 'undefined' ? SCG_CONFIG.SHEET_DATA : 'Data');
        if (dataSheet && dataSheet.getLastRow() > 1) {
           applyMasterCoordinatesToDailyJob();
           console.log("✅ AutoPilot: SCG Sync Completed");
        }
      }
    } catch(e) { console.error("[AutoPilot] SCG Sync Error: " + e.message); }

    try {
      processAIIndexing_Batch(); 
    } catch(e) { console.error("[AutoPilot] AI Indexing Error: " + e.message); }

    console.timeEnd("AutoPilot_Duration");
    console.info("[AutoPilot] 🏁 Routine finished successfully.");

  } catch (e) {
    console.error("CRITICAL AutoPilot Error: " + e.message);
  } finally {
    lock.releaseLock();
  }
}

function processAIIndexing_Batch() {
  var apiKey;
  try {
    apiKey = CONFIG.GEMINI_API_KEY;
  } catch (e) {
    console.warn("⚠️ SKIPPED AI: " + e.message); 
    return;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) return;

  var lastRow = typeof getRealLastRow_ === 'function' ? getRealLastRow_(sheet, CONFIG.COL_NAME) : sheet.getLastRow();
  if (lastRow < 2) return;

  var rangeName = sheet.getRange(2, CONFIG.COL_NAME, lastRow - 1, 1);
  var rangeNorm = sheet.getRange(2, CONFIG.COL_NORMALIZED, lastRow - 1, 1);
  
  var nameValues = rangeName.getValues();
  var normValues = rangeNorm.getValues(); 
  
  var aiCount = 0;
  var AI_LIMIT = (typeof CONFIG !== 'undefined' && CONFIG.AI_BATCH_SIZE) ? CONFIG.AI_BATCH_SIZE : 20; 
  var updated = false;

  for (var i = 0; i < nameValues.length; i++) {
    if (aiCount >= AI_LIMIT) break;

    var name = nameValues[i][0];
    var currentNorm = normValues[i][0];

    if (name && typeof name === 'string' && (!currentNorm || currentNorm.toString().indexOf("[AI]") === -1)) {
      
      var basicKey = createBasicSmartKey(name);
      var aiKeywords = "";
      
      if (name.length > 3) {
        aiKeywords = genericRetry(function() { 
          return callGeminiThinking_JSON(name, apiKey); 
        }, 2); 
      }
      
      var finalString = basicKey + (aiKeywords ? " " + aiKeywords : "") + " [AI]";
      normValues[i][0] = finalString.trim();
      
      console.log(`🤖 AI Processed (${aiCount+1}/${AI_LIMIT}): [${name}] -> ${aiKeywords}`);
      aiCount++;
      updated = true;
    }
  }

  if (updated) {
    rangeNorm.setValues(normValues);
    console.log(`✅ AI Batch Write: อัปเดตฐานข้อมูล ${aiCount} รายการ.`);
  } else {
    console.log("ℹ️ AI Standby: ไม่มีข้อมูลใหม่ที่ต้องให้ AI วิเคราะห์.");
  }
}

function callGeminiThinking_JSON(customerName, apiKey) {
  try {
    // [FIXED v4.2] Enforce latest model to prevent v1beta 404 NOT_FOUND API Errors
    var model = (typeof CONFIG !== 'undefined' && CONFIG.AI_MODEL) ? CONFIG.AI_MODEL : "gemini-1.5-flash-latest";
    var apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;
    
    var prompt = `
      Task: Analyze this Thai logistics customer name: "${customerName}"
      Goal: Return a JSON list of search keywords, abbreviations, and common typos.
      Requirements:
      1. If English, provide Thai phonetics.
      2. If Thai abbreviation (e.g., บจก, รพ), provide full text.
      3. No generic words like "Company", "Limited", "จำกัด", "บริษัท".
      4. Max 5 keywords.
      
      Output Format: JSON Array of Strings ONLY.
      Example: ["Keyword1", "Keyword2"]
    `;

    var payload = {
      "contents": [{ "parts": [{ "text": prompt }] }],
      "generationConfig": { "responseMimeType": "application/json" } 
    };

    var options = {
      "method": "post",
      "contentType": "application/json",
      "payload": JSON.stringify(payload),
      "muteHttpExceptions": true
    };

    var response = UrlFetchApp.fetch(apiUrl, options);
    var statusCode = response.getResponseCode();
    
    if (statusCode !== 200) {
      throw new Error(`API Error ${statusCode}: ${response.getContentText()}`);
    }

    var json = JSON.parse(response.getContentText());

    if (json.candidates && json.candidates.length > 0) {
      var text = json.candidates[0].content.parts[0].text;
      var keywords = JSON.parse(text); 
      
      if (Array.isArray(keywords)) {
        return keywords.join(" "); 
      }
    }
  } catch (e) {
    console.warn("Gemini Error (" + customerName + "): " + e.message);
    return ""; 
  }
  return "";
}

/**
 * 🔨 Helper: สร้าง Index แบบพื้นฐาน (Regex)
 * [FIXED v4.2]: ยกเลิกการเติมคำซ้ำ (ตัดวรรณยุกต์) เพื่อให้ข้อมูลดูสะอาดตาที่สุด
 */
function createBasicSmartKey(text) {
  if (!text) return "";
  var clean = typeof normalizeText === 'function' ? normalizeText(text) : text.toString().toLowerCase().replace(/\s/g, ""); 
  return clean; // คืนค่าเฉพาะคำที่ตัด Stop Words ออกแล้ว โดยไม่ Duplicate ให้รกช่อง
}