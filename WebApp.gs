/**
 * VERSION: 000
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
function doGet(e) {
  try {
    // บันทึก Log การเข้าใช้งาน
    console.info(`[WebApp] GET Request received. Params: ${JSON.stringify(e.parameter)}`);

    // 1. Page Routing (เตรียมพร้อมสำหรับหน้าจออื่นๆ เช่น Admin, Dashboard)
    var page = (e && e.parameter && e.parameter.page) ? e.parameter.page : 'Index';
    
    // 2. สร้าง Template จากไฟล์ HTML
    var template = HtmlService.createTemplateFromFile(page);
    
    // 3. รับค่าจาก URL Parameter (Deep Linking)
    var paramQuery = (e && e.parameter && e.parameter.q) ? e.parameter.q : "";
    template.initialQuery = paramQuery;
    
    // 4. ส่งค่า Config/Version ไปหน้าบ้าน (แก้ปัญหา Browser Cache)
    template.appVersion = new Date().getTime(); // บังคับโหลดใหม่เสมอ
    template.isEnterprise = true;
    
    // 5. Evaluate & Render
    var output = template.evaluate()
        .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=0') 
        .setTitle('🔍 Logistics Master Search (V4.0)')
        .setFaviconUrl('https://img.icons8.com/color/48/truck--v1.png');

    // 6. X-Frame Options 
    // ALLOWALL: จำเป็นสำหรับการ Embed ใน SharePoint, Google Sites หรือ AppSheet
    output.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    
    return output;

  } catch (err) {
    console.error(`[WebApp] GET Error: ${err.message}`);
    // Fallback กรณีระบบล่ม หรือหาไฟล์ HTML ไม่เจอ
    return HtmlService.createHtmlOutput(`
      <div style="font-family: sans-serif; padding: 20px; text-align: center; background-color: #ffebee;">
        <h3 style="color: #d32f2f;">❌ System Error (V4.0)</h3>
        <p>${err.message}</p>
        <p style="color: #666; font-size: 12px;">กรุณาตรวจสอบชื่อไฟล์ HTML หรือติดต่อ System Administrator</p>
      </div>
    `);
  }
}

/**
 * 📡 [ADDED v4.0] ฟังก์ชันรับข้อมูลผ่าน Webhook/API (HTTP POST)
 * รองรับการเชื่อมต่อจาก AppSheet หรือระบบภายนอกเพื่อสั่งงานเบื้องหลัง
 */
function doPost(e) {
  try {
    console.info("[WebApp] POST Request received.");
    if (!e || !e.postData) throw new Error("No payload found in POST request.");
    
    var payload = JSON.parse(e.postData.contents);
    var action = payload.action;

    // ตัวอย่างการทำ Routing API เบื้องต้น
    if (action === "triggerAIBatch") {
       // สั่งให้ AI ทำงานจากภายนอก
       if (typeof processAIIndexing_Batch === 'function') {
         processAIIndexing_Batch();
         return createJsonResponse_({ status: "success", message: "AI Batch Processing Triggered" });
       }
    }

    return createJsonResponse_({ status: "success", message: "Webhook received", data: payload });

  } catch (err) {
    console.error("[WebApp] POST Error: " + err.message);
    return createJsonResponse_({ status: "error", message: err.message });
  }
}

/**
 * Helper: สร้าง JSON Response ให้ doPost
 */
function createJsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/**
 * 📦 ฟังก์ชันสำหรับดึง CSS/JS เข้ามาใน HTML (Server-Side Include)
 */
function include(filename) {
  try {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch (e) {
    console.warn("[WebApp] Missing include file: " + filename);
    return "<!-- Error: File '" + filename + "' not found. -->";
  }
}

/**
 * 🔐 ฟังก์ชันสำหรับตรวจสอบ User (Safe Context)
 * เอาไว้เรียกจากฝั่ง Client เพื่อดูว่าใครใช้งานอยู่
 */
function getUserContext() {
  try {
    return {
      email: Session.getActiveUser().getEmail() || "anonymous",
      locale: Session.getActiveUserLocale() || "th"
    };
  } catch (e) {
    console.warn("[WebApp] Failed to get user context: " + e.message);
    return { email: "unknown", locale: "th" };
  }
}


