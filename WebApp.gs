/**
 * VERSION : 000
 * 🌐 WebApp Controller (Enterprise Edition)
 * Version: 4.0 Omni-Channel Interface
 * ------------------------------------------
 * [PRESERVED]: URL Parameter handling, Safe Include, Version Control.
 * [ADDED v4.0]: doPost() for API/Webhook readiness (AppSheet/External Triggers).
 * [ADDED v4.0]: Page routing logic (e.parameter.page) for multi-view support.
 * [MODIFIED v4.0]: Enterprise logging tracking for web accesses.
 * [MODIFIED v4.0]: Safe user context extraction.
 * Author: Elite Logistics Architect
 */

/**
 * 🖥️ ฟังก์ชันแสดงผลหน้าเว็บ (HTTP GET)
 * รองรับ: https://script.google.com/.../exec?q=ค้นหา&page=Index
 */
/**
 * VERSION: 4.2 — Phase E
 * [Phase E] แยก doPost() routing เป็น switch/handler map
 */

function doGet(e) {
  try {
    console.info("[WebApp] GET Request: " + JSON.stringify(e.parameter));
    var page     = (e && e.parameter && e.parameter.page) ? e.parameter.page : 'Index';
    var template = HtmlService.createTemplateFromFile(page);

    template.initialQuery = (e && e.parameter && e.parameter.q) ? e.parameter.q : "";
    template.appVersion   = new Date().getTime();
    template.isEnterprise = true;

    var output = template.evaluate()
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=0')
      .setTitle('🔍 Logistics Master Search (V4.2)')
      .setFaviconUrl('https://img.icons8.com/color/48/truck--v1.png');

    output.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    return output;

  } catch(err) {
    console.error("[WebApp] GET Error: " + err.message);
    return HtmlService.createHtmlOutput(
      '<div style="font-family:sans-serif;padding:20px;text-align:center;background:#ffebee;">' +
      '<h3 style="color:#d32f2f;">❌ System Error (V4.2)</h3>' +
      '<p>' + err.message + '</p></div>'
    );
  }
}

/**
 * [Phase E] doPost() — handler map แทน if/else
 */
function doPost(e) {
  try {
    console.info("[WebApp] POST Request received.");
    if (!e || !e.postData) throw new Error("No payload found.");

    var payload = JSON.parse(e.postData.contents);
    var action  = payload.action || "";

    // [Phase E] Handler map — เพิ่ม action ใหม่ได้ง่าย
    var handlers = {
      "triggerAIBatch": function() {
        if (typeof processAIIndexing_Batch === 'function') {
          processAIIndexing_Batch();
          return { status: "success", message: "AI Batch triggered" };
        }
        return { status: "error", message: "processAIIndexing_Batch not found" };
      },
      "triggerSync": function() {
        if (typeof syncNewDataToMaster === 'function') {
          return { status: "success", message: "Sync triggered" };
        }
        return { status: "error", message: "syncNewDataToMaster not found" };
      },
      "healthCheck": function() {
        try {
          CONFIG.validateSystemIntegrity();
          return { status: "success", message: "System healthy" };
        } catch(e) {
          return { status: "error", message: e.message };
        }
      }
    };

    if (handlers[action]) {
      return createJsonResponse_(handlers[action]());
    }

    return createJsonResponse_({ status: "success", message: "Webhook received", action: action });

  } catch(err) {
    console.error("[WebApp] POST Error: " + err.message);
    return createJsonResponse_({ status: "error", message: err.message });
  }
}

function createJsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
                       .setMimeType(ContentService.MimeType.JSON);
}

function include(filename) {
  try { return HtmlService.createHtmlOutputFromFile(filename).getContent(); }
  catch(e) { return "<!-- Error: File '" + filename + "' not found. -->"; }
}

function getUserContext() {
  try {
    return {
      email:  Session.getActiveUser().getEmail() || "anonymous",
      locale: Session.getActiveUserLocale()      || "th"
    };
  } catch(e) {
    return { email: "unknown", locale: "th" };
  }
}
