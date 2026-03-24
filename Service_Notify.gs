/**
 * VERSION : 000
 * 🔔 Service: Omni-Channel Notification Hub (Enterprise Edition)
 * Version: 4.0 Centralized Broadcaster
 * หน้าที่: ศูนย์กลางส่งแจ้งเตือนสถานะระบบและ Error เข้า LINE และ Telegram
 * ------------------------------------------------
 * [PRESERVED]: Dual-channel architecture and HTML escaping.
 * [REMOVED v4.0]: Setup functions removed (Delegated to Setup_Security.gs V4.0).
 * [MODIFIED v4.0]: Overrides basic notifiers in Module 14 with robust Try-Catch logic.
 * [MODIFIED v4.0]: Prevents API limits/errors from crashing main business flows.
 * Author: Elite Logistics Architect
 */

// ==========================================
// 1. CORE SENDING LOGIC (Unified Broadcaster)
// ==========================================

/**
 * 📤 ฟังก์ชันส่งข้อความรวม (Broadcast V4.0)
 * ส่งเข้าทุกช่องทางที่ตั้งค่าไว้ (LINE และ/หรือ Telegram)
 * @param {string} message - ข้อความ
 * @param {boolean} isUrgent - เป็น Error หรือเรื่องด่วนหรือไม่
 */
function sendSystemNotify(message, isUrgent) {
  console.info(`[Notification Hub] Broadcasting message (Urgent: ${!!isUrgent})`);
  
  // รันแบบขนาน (จำลองใน GAS โดยใช้ Try-Catch แยกกัน)
  // ป้องกันกรณีช่องทางใดช่องทางหนึ่งตาย แล้วพาลให้อีกช่องทางไม่ส่ง
  
  try {
    sendLineNotify_Internal_(message, isUrgent);
  } catch (e) {
    console.error("[Notify Hub] LINE Broadcast Failed: " + e.message);
  }

  try {
    sendTelegramNotify_Internal_(message, isUrgent);
  } catch (e) {
    console.error("[Notify Hub] Telegram Broadcast Failed: " + e.message);
  }
}

// ==========================================
// 2. PUBLIC WRAPPERS (Overrides Module 14)
// ==========================================

/**
 * [MODIFIED v4.0] Wrapper สำหรับเขียนทับ (Override) ฟังก์ชันใน Service_Maintenance.gs
 * ทำให้ทุกการเรียกใช้ sendLineNotify ในระบบ วิ่งมาใช้ Logic ระดับ Enterprise ตัวนี้แทน
 */
function sendLineNotify(message, isUrgent) {
  sendLineNotify_Internal_(message, isUrgent);
}

/**
 * [MODIFIED v4.0] Wrapper สำหรับเขียนทับ (Override) ฟังก์ชันใน Service_Maintenance.gs
 */
function sendTelegramNotify(message, isUrgent) {
  sendTelegramNotify_Internal_(message, isUrgent);
}

// ==========================================
// 3. INTERNAL CHANNEL HANDLERS
// ==========================================

/**
 * Internal: ยิง API เข้า LINE Notify อย่างปลอดภัย
 */
function sendLineNotify_Internal_(message, isUrgent) {
  var token = PropertiesService.getScriptProperties().getProperty('LINE_NOTIFY_TOKEN');
  if (!token) return; // Silent skip if not configured

  var prefix = isUrgent ? "🚨 URGENT ALERT:\n" : "🤖 SYSTEM REPORT:\n";
  var fullMsg = prefix + message;

  try {
    var response = UrlFetchApp.fetch("https://notify-api.line.me/api/notify", {
      "method": "post",
      "headers": { "Authorization": "Bearer " + token },
      "payload": { "message": fullMsg },
      "muteHttpExceptions": true
    });
    
    if (response.getResponseCode() !== 200) {
      console.warn("[LINE API Error] " + response.getContentText());
    }
  } catch (e) {
    console.warn("[LINE Exception] " + e.message);
  }
}

/**
 * Internal: ยิง API เข้า Telegram อย่างปลอดภัย
 */
function sendTelegramNotify_Internal_(message, isUrgent) {
  var token = PropertiesService.getScriptProperties().getProperty('TG_BOT_TOKEN'); // ใช้ Key ตาม Setup_Security V4.0
  var chatId = PropertiesService.getScriptProperties().getProperty('TG_CHAT_ID');  // ใช้ Key ตาม Setup_Security V4.0
  
  // Fallback for V2.0 keys if still present
  if (!token) token = PropertiesService.getScriptProperties().getProperty('TELEGRAM_BOT_TOKEN');
  if (!chatId) chatId = PropertiesService.getScriptProperties().getProperty('TELEGRAM_CHAT_ID');

  if (!token || !chatId) return; // Silent skip if not configured

  // Format Message (HTML Style)
  var icon = isUrgent ? "🚨" : "🤖";
  var title = isUrgent ? "<b>SYSTEM ALERT</b>" : "<b>SYSTEM REPORT</b>";
  var htmlMsg = `${icon} ${title}\n\n${escapeHtml_(message)}`;

  try {
    var url = "https://api.telegram.org/bot" + token + "/sendMessage";
    var payload = {
      "chat_id": chatId,
      "text": htmlMsg,
      "parse_mode": "HTML"
    };

    var response = UrlFetchApp.fetch(url, {
      "method": "post",
      "contentType": "application/json",
      "payload": JSON.stringify(payload),
      "muteHttpExceptions": true
    });
    
    if (response.getResponseCode() !== 200) {
      console.warn("[Telegram API Error] " + response.getContentText());
    }
  } catch (e) {
    console.warn("[Telegram Exception] " + e.message);
  }
}

/**
 * Helper: Escape HTML special chars for Telegram to prevent formatting errors
 */
function escapeHtml_(text) {
  if (!text) return "";
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

// ==========================================
// 4. SPECIFIC EVENT NOTIFIERS
// ==========================================

/**
 * [UPGRADED v4.0] Wrapper สำหรับ AutoPilot
 * สรุปยอดการทำงานส่งให้ผู้ดูแลระบบ
 */
function notifyAutoPilotStatus(scgStatus, aiCount, aiMappedCount) {
  // รองรับพารามิเตอร์ 3 ตัวเพื่อโชว์ผลลัพธ์ของ Tier 4 AI ด้วย
  var mappedMsg = aiMappedCount !== undefined ? `\n🎯 AI Tier-4 จับคู่สำเร็จ: ${aiMappedCount} ร้าน` : "";
  
  var msg = "------------------\n" +
            "✅ AutoPilot V4.0 รอบล่าสุด:\n" +
            "📦 ดึงงาน SCG: " + scgStatus + "\n" +
            "🧠 AI Indexing: " + aiCount + " รายการ" + 
            mappedMsg;
            
  sendSystemNotify(msg, false); 
}


