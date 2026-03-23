/**
 * VERSION: 000
 * 🕵️ Service: Logistics AI Agent (Enterprise Edition)
 * Codename: "The Steward"
 * Version: 4.0 Smart Resolution & Safe Concurrency
 * -------------------------------------------
 * [PRESERVED]: Manual/Scheduled Triggers and basic typo prediction logic.
 * [FIXED v4.0]: Changed Full-Sheet write to Specific-Column write to prevent data collision.
 * [ADDED v4.0]: resolveUnknownNamesWithAI() - The Tier 4 Smart Resolution engine 
 * that maps unknown names to Master_UIDs and auto-updates NameMapping.
 * [MODIFIED v4.0]: AI Calls now enforce application/json for system stability.
 * Author: Elite Logistics Architect
 */

var AGENT_CONFIG = {
  NAME: "Logistics_Agent_01",
  MODEL: (typeof CONFIG !== 'undefined' && CONFIG.AI_MODEL) ? CONFIG.AI_MODEL : "gemini-1.5-flash",
  BATCH_SIZE: (typeof CONFIG !== 'undefined' && CONFIG.AI_BATCH_SIZE) ? CONFIG.AI_BATCH_SIZE : 20,
  TAG: "[Agent_V4]" // Tag ประจำตัว Agent รุ่นใหม่
};

// ==========================================
// 1. AGENT TRIGGERS & CONTROLS
// ==========================================

/**
 * 👋 สั่ง Agent ให้ตื่นมาทำงานเดี๋ยวนี้ (Manual Trigger)
 */
function WAKE_UP_AGENT() {
  SpreadsheetApp.getUi().toast("🕵️ Agent: ผมตื่นแล้วครับ กำลังเริ่มวิเคราะห์ข้อมูล...", "AI Agent Started");

  try {
    processAIIndexing_Batch();  // ← เปลี่ยนตรงนี้
    SpreadsheetApp.getUi().alert("✅ Agent รายงานผล:\nวิเคราะห์ข้อมูลชุดล่าสุดเสร็จสิ้น (Batch Mode)");
  } catch (e) {
    SpreadsheetApp.getUi().alert("❌ Agent Error: " + e.message);
  }
}

/**
 * ⏰ ตั้งเวลาให้ Agent ตื่นมาทำงานเองทุก 10 นาที
 */
function SCHEDULE_AGENT_WORK() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "runAgentLoop") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  ScriptApp.newTrigger("runAgentLoop")
    .timeBased()
    .everyMinutes(10)
    .create();

  SpreadsheetApp.getUi().alert("✅ ตั้งค่าเรียบร้อย!\nThe Steward จะทำงานทุก 10 นาที");
}

// ==========================================
// 2. TIER 4: SMART RESOLUTION (NEW v4.0)
// ==========================================

/**
 * 🧠 [NEW v4.0] ฟังก์ชันส่งชื่อแปลกๆ ให้ AI วิเคราะห์จับคู่กับ Database
 * ถูกเรียกใช้โดยเมนู: 🧠 4️⃣ ส่งชื่อแปลกให้ AI วิเคราะห์ (Smart Resolution)
 */
function resolveUnknownNamesWithAI() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = ss.getSheetByName(typeof SCG_CONFIG !== 'undefined' ? SCG_CONFIG.SHEET_DATA : 'Data');
  var dbSheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  var mapSheet = ss.getSheetByName(CONFIG.MAPPING_SHEET);

  if (!dataSheet || !dbSheet || !mapSheet) return;

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    SpreadsheetApp.getUi().alert("⚠️ ระบบคิวทำงาน", "มีระบบอื่นกำลังใช้งานอยู่ กรุณารอสักครู่", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  try {
    console.time("SmartResolution_Time");

    // 1. หาชื่อที่ยังจับคู่ไม่ได้จากชีต Data (ดูจากคอลัมน์พิกัดว่าว่างไหม)
    var dLastRow = dataSheet.getLastRow();
    if (dLastRow < 2) return;

    var dataValues = dataSheet.getRange(2, 1, dLastRow - 1, 29).getValues();
    var unknownNames = new Set();

    dataValues.forEach(function (r) {
      var shipToName = r[10]; // Col K: ShipToName
      var actualGeo = r[26];  // Col AA: LatLong_Actual (พิกัดที่ระบบหาได้)
      if (shipToName && !actualGeo) {
        unknownNames.add(normalizeText(shipToName));
      }
    });

    var unknownsArray = Array.from(unknownNames).slice(0, AGENT_CONFIG.BATCH_SIZE);
    if (unknownsArray.length === 0) {
      SpreadsheetApp.getUi().alert("ℹ️ AI Standby: ไม่มีรายชื่อตกหล่นที่ต้องให้ AI วิเคราะห์ครับ");
      return;
    }

    // 2. ดึง Master Data มาเป็นตัวเลือกให้ AI
    var mLastRow = dbSheet.getLastRow();
    var dbValues = dbSheet.getRange(2, 1, mLastRow - 1, Math.max(CONFIG.COL_NAME, CONFIG.COL_UUID)).getValues();
    var masterOptions = [];

    dbValues.forEach(function (r) {
      var name = r[CONFIG.C_IDX.NAME];
      var uid = r[CONFIG.C_IDX.UUID];
      if (name && uid) {
        masterOptions.push({ "uid": uid, "name": name });
      }
    });

    // Limit master options to 500 to save context window (Optional, Gemini 1.5 handles big context well)
    var masterSubset = masterOptions.slice(0, 500);

    SpreadsheetApp.getActiveSpreadsheet().toast(`กำลังส่ง ${unknownsArray.length} รายชื่อให้ AI วิเคราะห์...`, "🤖 Tier 4 AI", 10);

    // 3. ส่งข้อมูลให้ Gemini คิด (Prompt Engineering)
    var apiKey = CONFIG.GEMINI_API_KEY;
    var prompt = `
      You are an expert Thai Logistics Data Analyst.
      I have a list of 'unknown_names' from a daily delivery sheet. They contain typos, abbreviations, or missing branches.
      I also have a 'master_database' of valid delivery locations with their UIDs.
      
      Task: Match each unknown name to the most likely master database entry.
      If confidence is less than 60%, do not match it (skip it).
      
      Unknown Names: ${JSON.stringify(unknownsArray)}
      Master Database: ${JSON.stringify(masterSubset)}
      
      Output ONLY a JSON array of objects with this format:
      [ { "variant": "Unknown Name", "uid": "Matched UID", "confidence": 95 } ]
    `;

    var payload = {
      "contents": [{ "parts": [{ "text": prompt }] }],
      "generationConfig": { "responseMimeType": "application/json", "temperature": 0.1 }
    };

    var response = UrlFetchApp.fetch(`https://generativelanguage.googleapis.com/v1beta/models/${AGENT_CONFIG.MODEL}:generateContent?key=${apiKey}`, {
      "method": "post", "contentType": "application/json", "payload": JSON.stringify(payload), "muteHttpExceptions": true
    });

    var json = JSON.parse(response.getContentText());
    if (!json.candidates || json.candidates.length === 0) throw new Error("AI returned no results.");

    var aiResultText = json.candidates[0].content.parts[0].text;
    var matchedResults = JSON.parse(aiResultText);

    // 4. บันทึกผลลง NameMapping (5-Column Schema V4.0)
    var mapRows = [];
    var ts = new Date();

    if (Array.isArray(matchedResults) && matchedResults.length > 0) {
      matchedResults.forEach(function (match) {
        if (match.uid && match.confidence >= 60) {
          mapRows.push([
            match.variant,       // Variant_Name
            match.uid,           // Master_UID
            match.confidence,    // Confidence_Score
            "AI_Agent_V4",       // Mapped_By
            ts                   // Timestamp
          ]);
        }
      });
    }

    if (mapRows.length > 0) {
      mapSheet.getRange(mapSheet.getLastRow() + 1, 1, mapRows.length, 5).setValues(mapRows);

      // สั่งเคลียร์ Cache ค้นหา และ รันจับคู่พิกัดซ้ำทันที
      if (typeof clearSearchCache === 'function') clearSearchCache();
      if (typeof applyMasterCoordinatesToDailyJob === 'function') applyMasterCoordinatesToDailyJob();

      SpreadsheetApp.getUi().alert(`✅ AI ทำงานสำเร็จ!\nจับคู่รายชื่อสำเร็จ ${mapRows.length} รายการ และบันทึกลง NameMapping อัตโนมัติแล้ว`);
    } else {
      SpreadsheetApp.getUi().alert("ℹ️ AI ทำงานเสร็จสิ้น แต่ไม่สามารถจับคู่รายชื่อด้วยความมั่นใจเกิน 60% ได้ (ต้องตรวจสอบมือ)");
    }

  } catch (e) {
    console.error("[AI Smart Resolution Error]: " + e.message);
    SpreadsheetApp.getUi().alert("❌ เกิดข้อผิดพลาดในระบบ AI: " + e.message);
  } finally {
    lock.releaseLock();
    console.timeEnd("SmartResolution_Time");
  }
}



/**
 * 📡 Skill: การคาดเดาคำผิด (Typos Prediction)
 * [MODIFIED v4.0]: Enforced JSON output for stability
 */
function askGeminiToPredictTypos(originalName) {
  var prompt = `
    Task: You are a Thai Logistics Search Agent.
    Input Name: "${originalName}"
    Goal: Generate search keywords including common typos, phonetic spellings, and abbreviations.
    Constraint: Output ONLY a JSON array of strings.
    Example Input: "บี-ควิก (สาขาลาดพร้าว)"
    Example Output: ["บีควิก", "บีขวิก", "บีวิก", "BeQuik", "BQuik", "B-Quik", "ลาดพร้าว", "BQuick"]
  `;

  var payload = {
    "contents": [{ "parts": [{ "text": prompt }] }],
    "generationConfig": { "responseMimeType": "application/json", "temperature": 0.4 }
  };

  var options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  var url = `https://generativelanguage.googleapis.com/v1beta/models/${AGENT_CONFIG.MODEL}:generateContent?key=${CONFIG.GEMINI_API_KEY}`;

  var response = UrlFetchApp.fetch(url, options);

  if (response.getResponseCode() !== 200) {
    throw new Error("Gemini API Error: " + response.getContentText());
  }

  var json = JSON.parse(response.getContentText());

  if (json.candidates && json.candidates[0].content) {
    var text = json.candidates[0].content.parts[0].text;
    var keywordsArray = JSON.parse(text);
    if (Array.isArray(keywordsArray)) {
      return keywordsArray.join(" "); // รวมเป็น String เพื่อเก็บลงช่อง Normalized
    }
  }

  return "";
}


