/**
 * VERSION: 4.2 — Phase D
 * [Phase D] เพิ่ม retrieveCandidateMasters_() ก่อนส่ง AI
 * [Phase D] เพิ่ม confidence bands: auto-map / review / ignore
 * [Phase D] เพิ่ม AI audit logging
 */

var AGENT_CONFIG = {
  NAME:       "Logistics_Agent_01",
  MODEL:      (typeof CONFIG !== 'undefined' && CONFIG.AI_MODEL) ? CONFIG.AI_MODEL : "gemini-1.5-flash",
  BATCH_SIZE: (typeof CONFIG !== 'undefined' && CONFIG.AI_BATCH_SIZE) ? CONFIG.AI_BATCH_SIZE : 20,
  TAG:        "[Agent_V4]"
};

// ==========================================
// 1. AGENT TRIGGERS
// ==========================================

function WAKE_UP_AGENT() {
  SpreadsheetApp.getUi().toast("🕵️ Agent: กำลังเริ่มวิเคราะห์...", "AI Agent Started");
  try {
    processAIIndexing_Batch();
    SpreadsheetApp.getUi().alert("✅ Agent รายงานผล:\nวิเคราะห์ข้อมูลชุดล่าสุดเสร็จสิ้น");
  } catch(e) {
    SpreadsheetApp.getUi().alert("❌ Agent Error: " + e.message);
  }
}

function SCHEDULE_AGENT_WORK() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "runAgentLoop") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger("autoPilotRoutine").timeBased().everyMinutes(10).create();
  SpreadsheetApp.getUi().alert("✅ ตั้งค่าเรียบร้อย!\nระบบจะทำงานทุก 10 นาที");
}

// ==========================================
// 2. [Phase D NEW] RETRIEVAL HELPER
// ==========================================

/**
 * retrieveCandidateMasters_()
 * คัด top-N candidates ที่เกี่ยวข้องก่อนส่ง AI
 * แทน slice(0, 500) แบบตัดตรง
 */
function retrieveCandidateMasters_(unknownName, dbRows, mapRows, limit) {
  var normUnknown = normalizeText(unknownName);
  var tokens      = normUnknown.split(/\s+/).filter(function(t) { return t.length > 1; });
  var scored      = [];

  dbRows.forEach(function(r) {
    var obj = dbRowToObject(r);
    if (!obj.name || !obj.uuid) return;
    if (obj.recordStatus === "Inactive" || obj.recordStatus === "Merged") return;

    var normName = normalizeText(obj.name);
    var score    = 0;

    // exact match
    if (normName === normUnknown) score += 100;

    // token overlap
    tokens.forEach(function(token) {
      if (normName.indexOf(token) > -1) score += 10;
    });

    // partial prefix match
    if (normName.indexOf(normUnknown.substring(0, 3)) > -1) score += 5;

    if (score > 0) scored.push({ uid: obj.uuid, name: obj.name, score: score });
  });

  // เพิ่ม alias matches
  if (mapRows && mapRows.length > 0) {
    mapRows.forEach(function(r) {
      var mapObj = mapRowToObject(r);
      if (!mapObj || !mapObj.uid) return;
      var normVariant = normalizeText(mapObj.variant || "");
      var score       = 0;
      if (normVariant === normUnknown) score += 80;
      tokens.forEach(function(token) {
        if (normVariant.indexOf(token) > -1) score += 8;
      });
      if (score > 0) scored.push({ uid: mapObj.uid, name: mapObj.variant, score: score });
    });
  }

  // sort by score desc แล้วตัด limit
  scored.sort(function(a, b) { return b.score - a.score; });

  // dedupe by uid
  var seen   = new Set();
  var result = [];
  scored.forEach(function(item) {
    if (!seen.has(item.uid) && result.length < (limit || AI_CONFIG.RETRIEVAL_LIMIT)) {
      seen.add(item.uid);
      result.push({ uid: item.uid, name: item.name });
    }
  });

  return result;
}

// ==========================================
// 3. TIER 4: SMART RESOLUTION
// ==========================================

function resolveUnknownNamesWithAI() {
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = ss.getSheetByName(SCG_CONFIG.SHEET_DATA);
  var dbSheet   = ss.getSheetByName(CONFIG.SHEET_NAME);
  var mapSheet  = ss.getSheetByName(CONFIG.MAPPING_SHEET);

  if (!dataSheet || !dbSheet || !mapSheet) return;

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    SpreadsheetApp.getUi().alert("⚠️ ระบบคิวทำงาน", "กรุณารอสักครู่", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  try {
    console.time("SmartResolution_Time");

    // หาชื่อที่ยังไม่มีพิกัด
    var dLastRow = dataSheet.getLastRow();
    if (dLastRow < 2) return;

    var dataValues   = dataSheet.getRange(2, 1, dLastRow - 1, CONFIG.DATA_TOTAL_COLS).getValues();
    var unknownNames = new Set();

    dataValues.forEach(function(r) {
      var job = dailyJobRowToObject(r);
      if (job.shipToName && !job.latLngActual) {
        unknownNames.add(job.shipToName);
      }
    });

    var unknownsArray = Array.from(unknownNames).slice(0, AGENT_CONFIG.BATCH_SIZE);
    if (unknownsArray.length === 0) {
      SpreadsheetApp.getUi().alert("ℹ️ AI Standby: ไม่มีรายชื่อตกหล่น");
      return;
    }

    // [Phase D] โหลด DB + Map rows สำหรับ retrieval
    var mLastRow = dbSheet.getLastRow();
    var dbRows   = dbSheet.getRange(2, 1, mLastRow - 1, CONFIG.DB_TOTAL_COLS).getValues();
    var mapRows  = loadNameMappingRows_();

    // [Phase D] โหลด UUID state map
    var uuidStateMap = buildUUIDStateMap_();

    var apiKey = CONFIG.GEMINI_API_KEY;

    var autoMapped  = [];  // confidence >= 90
    var reviewItems = [];  // confidence 70-89
    var ts          = new Date();
    var auditLog    = [];  // [Phase D] audit trail

    unknownsArray.forEach(function(unknownName) {
      // [Phase D] retrieval ก่อน AI
      var candidates = retrieveCandidateMasters_(unknownName, dbRows, mapRows, AI_CONFIG.RETRIEVAL_LIMIT);

      if (candidates.length === 0) {
        console.log("[resolveUnknownNamesWithAI] '" + unknownName + "': no candidates found");
        return;
      }

      SpreadsheetApp.getActiveSpreadsheet().toast(
        "กำลัง AI วิเคราะห์: " + unknownName, "🤖 Tier 4 AI", 5
      );

      var prompt =
        "You are an expert Thai Logistics Data Analyst.\n" +
        "Match this unknown delivery name to the most likely entry in the candidate list.\n" +
        "If confidence < 60%, do not match.\n" +
        "Prompt-Version: " + AI_CONFIG.PROMPT_VERSION + "\n\n" +
        "Unknown Name: " + JSON.stringify(normalizeText(unknownName)) + "\n" +
        "Candidates: " + JSON.stringify(candidates) + "\n\n" +
        "Output ONLY a JSON object: { \"uid\": \"matched UID\", \"confidence\": 95 }\n" +
        "Or if no match: { \"uid\": null, \"confidence\": 0 }";

      var payload = {
        "contents": [{ "parts": [{ "text": prompt }] }],
        "generationConfig": { "responseMimeType": "application/json", "temperature": 0.1 }
      };

      try {
        var response = UrlFetchApp.fetch(
          "https://generativelanguage.googleapis.com/v1beta/models/" +
          AGENT_CONFIG.MODEL + ":generateContent?key=" + apiKey,
          { "method": "post", "contentType": "application/json",
            "payload": JSON.stringify(payload), "muteHttpExceptions": true }
        );

        var statusCode = response.getResponseCode();

        // [Phase D] parse guard
        if (statusCode !== 200) {
          console.warn("[resolveUnknownNamesWithAI] HTTP " + statusCode + " for '" + unknownName + "'");
          return;
        }

        var json = JSON.parse(response.getContentText());

        if (!json.candidates || !json.candidates[0] ||
            !json.candidates[0].content || !json.candidates[0].content.parts) {
          console.warn("[resolveUnknownNamesWithAI] Invalid structure for '" + unknownName + "'");
          return;
        }

        var result     = JSON.parse(json.candidates[0].content.parts[0].text);
        var matchedUid = result.uid;
        var confidence = result.confidence || 0;

        // [Phase D] audit log
        auditLog.push({
          unknownName:     unknownName,
          candidateCount:  candidates.length,
          chosenUid:       matchedUid,
          confidence:      confidence,
          promptVersion:   AI_CONFIG.PROMPT_VERSION,
          model:           AGENT_CONFIG.MODEL,
          timestamp:       ts
        });

        if (!matchedUid || confidence < AI_CONFIG.THRESHOLD_IGNORE) {
          console.log("[resolveUnknownNamesWithAI] '" + unknownName + "': confidence " + confidence + " < threshold → ignore");
          return;
        }

        // [Phase D] resolve canonical UUID ก่อน write
        var canonicalUid = resolveUUIDFromMap_(matchedUid, uuidStateMap);
        if (!canonicalUid || !isActiveFromMap_(canonicalUid, uuidStateMap)) {
          console.warn("[resolveUnknownNamesWithAI] '" + unknownName + "': canonical UUID inactive → skip");
          return;
        }

        // [Phase D] Confidence bands
        if (confidence >= AI_CONFIG.THRESHOLD_AUTO_MAP) {
          // auto-map ทันที
          autoMapped.push(mapObjectToRow({
            variant:    unknownName,
            uid:        canonicalUid,
            confidence: confidence,
            mappedBy:   "AI_Agent_" + AI_CONFIG.PROMPT_VERSION,
            timestamp:  ts
          }));
          console.log("[resolveUnknownNamesWithAI] AUTO-MAP: '" + unknownName +
                      "' → " + canonicalUid.substring(0, 8) + "... (conf: " + confidence + ")");

        } else if (confidence >= AI_CONFIG.THRESHOLD_REVIEW) {
          // ส่งเข้า review
          reviewItems.push(mapObjectToRow({
            variant:    unknownName,
            uid:        canonicalUid,
            confidence: confidence,
            mappedBy:   "AI_REVIEW_PENDING",
            timestamp:  ts
          }));
          console.log("[resolveUnknownNamesWithAI] REVIEW: '" + unknownName +
                      "' → " + canonicalUid.substring(0, 8) + "... (conf: " + confidence + ")");
        }

      } catch(e) {
        console.error("[resolveUnknownNamesWithAI] Error for '" + unknownName + "': " + e.message);
      }
    });

    // [Phase D] audit log summary
    console.log("[resolveUnknownNamesWithAI] Audit log: " + JSON.stringify(auditLog));

    // เขียน auto-mapped
    if (autoMapped.length > 0) {
      appendNameMappings_(autoMapped);
      if (typeof clearSearchCache === 'function') clearSearchCache();
      if (typeof applyMasterCoordinatesToDailyJob === 'function') applyMasterCoordinatesToDailyJob();
    }

    // เขียน review items (confidence band 70-89)
    if (reviewItems.length > 0) {
      appendNameMappings_(reviewItems);
      console.log("[resolveUnknownNamesWithAI] Review items written: " + reviewItems.length);
    }

    var msg = "✅ AI ทำงานเสร็จสิ้น!\n\n" +
              "🔍 วิเคราะห์: "       + unknownsArray.length + " รายชื่อ\n" +
              "✅ Auto-mapped (≥90): " + autoMapped.length   + " รายการ\n" +
              "👀 Review (70-89): "   + reviewItems.length   + " รายการ\n" +
              "❌ Ignored (<70): "    +
              (unknownsArray.length - autoMapped.length - reviewItems.length) + " รายการ";

    SpreadsheetApp.getUi().alert(msg);

  } catch(e) {
    console.error("[resolveUnknownNamesWithAI] CRITICAL: " + e.message);
    SpreadsheetApp.getUi().alert("❌ เกิดข้อผิดพลาด: " + e.message);
  } finally {
    lock.releaseLock();
    console.timeEnd("SmartResolution_Time");
  }
}

function askGeminiToPredictTypos(originalName) {
  var prompt =
    "Task: Thai Logistics Search Agent.\n" +
    "Input: \"" + originalName + "\"\n" +
    "Goal: Generate keywords including typos, phonetics, abbreviations.\n" +
    "Prompt-Version: " + AI_CONFIG.PROMPT_VERSION + "\n" +
    "Output: JSON Array of Strings ONLY.";

  var payload = {
    "contents": [{ "parts": [{ "text": prompt }] }],
    "generationConfig": { "responseMimeType": "application/json", "temperature": 0.4 }
  };

  var response = UrlFetchApp.fetch(
    "https://generativelanguage.googleapis.com/v1beta/models/" +
    AGENT_CONFIG.MODEL + ":generateContent?key=" + CONFIG.GEMINI_API_KEY,
    { "method": "post", "contentType": "application/json",
      "payload": JSON.stringify(payload), "muteHttpExceptions": true }
  );

  if (response.getResponseCode() !== 200) return "";

  var json = JSON.parse(response.getContentText());
  if (!json.candidates || !json.candidates[0] ||
      !json.candidates[0].content) return "";

  var text     = json.candidates[0].content.parts[0].text;
  var keywords = JSON.parse(text);
  return Array.isArray(keywords) ? keywords.join(" ") : "";
}
