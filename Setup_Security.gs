/**
 * VERSION : 000
 * 🔐 Security Setup Utility (Enterprise Edition)
 * Version: 4.0 Omni-Vault (Safe Storage & Validation)
 * -----------------------------------------------------------------
 * [PRESERVED]: PropertiesService for secure credential storage.
 * [MODIFIED v4.0]: Upgraded validation to check for "AIza" prefix for Gemini.
 * [MODIFIED v4.0]: Changed resetEnvironment to selectively delete keys (preventing full wipe).
 * [ADDED v4.0]: setupLineToken() & setupTelegramConfig() to support Menu V4.0.
 * [MODIFIED v4.0]: Switched to console.info for GCP Audit Logging.
 * Author: Elite Logistics Architect
 */

// ==========================================
// 1. GEMINI AI (CORE SECURITY)
// ==========================================

/**
 * 🔐 ตั้งค่า Gemini API Key อย่างปลอดภัย
 * ห้ามแก้ Config.gs เพื่อใส่ Key โดยตรงเด็ดขาด!
 */
function setupEnvironment() {
  var ui = SpreadsheetApp.getUi();
  
  var response = ui.prompt(
    '🔐 Security Setup: Gemini API', 
    'กรุณากรอก Gemini API Key (ต้องขึ้นต้นด้วย AIza...):\nสามารถรับฟรีได้ที่ Google AI Studio', 
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() == ui.Button.OK) {
    var key = response.getResponseText().trim();
    
    // [MODIFIED v4.0]: ตรวจสอบความถูกต้องของ Key ขั้นสูง
    if (key.length > 30 && key.startsWith("AIza")) {
      // Save to Script Properties (Hidden & Secure)
      PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', key);
      
      ui.alert('✅ บันทึก API Key สำเร็จ!\nระบบ AI พร้อมใช้งานแล้วครับ');
      console.info("[Security Audit] User updated GEMINI_API_KEY.");
    } else {
      ui.alert('❌ API Key ไม่ถูกต้อง', 'Key ของ Gemini ต้องขึ้นต้นด้วย "AIza" และมีความยาวที่ถูกต้อง กรุณาลองใหม่ครับ', ui.ButtonSet.OK);
      console.warn("[Security Audit] Failed attempt to update GEMINI_API_KEY (Invalid format).");
    }
  } else {
    console.info("[Security Audit] Setup cancelled by user.");
  }
}

// ==========================================
// 2. NOTIFICATION TOKENS (NEW v4.0)
// ==========================================

/**
 * 🔔 [ADDED v4.0] ตั้งค่า LINE Notify Token
 * รองรับเมนู V4.0 ที่ประกาศไว้ใน Menu.gs
 */
function setupLineToken() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt(
    '🔔 Setup: LINE Notify', 
    'กรุณากรอก LINE Notify Token ของกลุ่มที่ต้องการให้ระบบแจ้งเตือน:', 
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() == ui.Button.OK) {
    var token = response.getResponseText().trim();
    if (token.length > 20) {
      PropertiesService.getScriptProperties().setProperty('LINE_NOTIFY_TOKEN', token);
      ui.alert('✅ บันทึก LINE Token สำเร็จ!');
      console.info("[Security Audit] User updated LINE_NOTIFY_TOKEN.");
    } else {
      ui.alert('❌ Token สั้นเกินไป กรุณาตรวจสอบอีกครั้ง');
    }
  }
}

/**
 * ✈️ [ADDED v4.0] ตั้งค่า Telegram Config (Bot Token & Chat ID)
 */
function setupTelegramConfig() {
  var ui = SpreadsheetApp.getUi();
  var props = PropertiesService.getScriptProperties();
  
  var resBot = ui.prompt('✈️ Setup: Telegram', '1. กรุณากรอก Bot Token (เช่น 123456:ABC-DEF...):', ui.ButtonSet.OK_CANCEL);
  if (resBot.getSelectedButton() !== ui.Button.OK) return;
  var botToken = resBot.getResponseText().trim();

  var resChat = ui.prompt('✈️ Setup: Telegram', '2. กรุณากรอก Chat ID (เช่น -100123456789):', ui.ButtonSet.OK_CANCEL);
  if (resChat.getSelectedButton() !== ui.Button.OK) return;
  var chatId = resChat.getResponseText().trim();

  if (botToken && chatId) {
    props.setProperty('TG_BOT_TOKEN', botToken);
    props.setProperty('TG_CHAT_ID', chatId);
    ui.alert('✅ บันทึก Telegram Config สำเร็จ!');
    console.info("[Security Audit] User updated Telegram configurations.");
  } else {
    ui.alert('❌ ข้อมูลไม่ครบถ้วน ยกเลิกการบันทึก');
  }
}

// ==========================================
// 3. MAINTENANCE & AUDIT
// ==========================================

/**
 * 🗑️ [MODIFIED v4.0] ล้างค่าเฉพาะระบบที่ต้องการ (Safe Reset)
 * ป้องกันการเผลอลบ Token สำคัญอื่นๆ ที่ไม่ได้เกี่ยวข้อง
 */
function resetEnvironment() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    '⚠️ Danger Zone', 
    'คุณต้องการล้างรหัส API Key ของ Gemini ใช่หรือไม่?\n(ระบบจะลบเฉพาะ GEMINI_API_KEY เท่านั้น)', 
    ui.ButtonSet.YES_NO
  );

  if (response == ui.Button.YES) {
    PropertiesService.getScriptProperties().deleteProperty('GEMINI_API_KEY');
    ui.alert('🗑️ ล้างการตั้งค่า Gemini API Key เรียบร้อยแล้ว');
    console.info("[Security Audit] User DELETED GEMINI_API_KEY.");
  }
}

/**
 * 🏥 ตรวจสอบสถานะการเชื่อมต่อ (System Secrets Status)
 * ใช้ตรวจเช็คว่าเราลืมใส่ Key ไหนไปบ้าง โดยไม่เปิดเผย Key จริง
 */
function checkCurrentKeyStatus() {
  var props = PropertiesService.getScriptProperties();
  var geminiKey = props.getProperty('GEMINI_API_KEY');
  var lineToken = props.getProperty('LINE_NOTIFY_TOKEN');
  var tgBot = props.getProperty('TG_BOT_TOKEN');
  var ui = SpreadsheetApp.getUi();
  
  var statusMsg = "📊 **System Secrets Status**\n\n";
  
  if (geminiKey) {
    statusMsg += "🟢 Gemini AI: READY (Ends with ..." + geminiKey.slice(-4) + ")\n";
  } else {
    statusMsg += "🔴 Gemini AI: NOT SET\n";
  }

  if (lineToken) {
    statusMsg += "🟢 LINE Notify: READY\n";
  } else {
    statusMsg += "⚪ LINE Notify: NOT SET\n";
  }

  if (tgBot) {
    statusMsg += "🟢 Telegram: READY\n";
  } else {
    statusMsg += "⚪ Telegram: NOT SET\n";
  }

  ui.alert("System Health Check", statusMsg, ui.ButtonSet.OK);
  console.info("[Security Audit] Secrets status checked by user.");
}


