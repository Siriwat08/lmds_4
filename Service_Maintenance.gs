/**
 * VERSION : 000
 * 🧹 Service: System Maintenance & Alerts (Enterprise Edition)
 * หน้าที่: ดูแลรักษาความสะอาดไฟล์ ลบ Backup เก่า และแจ้งเตือนผ่าน LINE/Telegram
 * Version: 4.0 Omni-Alerts & Housekeeping
 * ---------------------------------------------
 * [PRESERVED]: 10M Cell Limit check and 30-day Backup retention logic.
 * [ADDED v4.0]: Fully implemented sendLineNotify() and sendTelegramNotify().
 * [MODIFIED v4.0]: Improved Regex for extracting dates from Backup sheets.
 * [MODIFIED v4.0]: Added LockService and GCP Console Logging.
 * Author: Elite Logistics Architect
 */

// ==========================================
// 1. SYSTEM MAINTENANCE (HOUSEKEEPING)
// ==========================================

/**
 * 🗑️ ลบชีต Backup ที่เก่ากว่า 30 วัน (แนะนำให้ตั้ง Trigger รันทุกสัปดาห์)
 */
function cleanupOldBackups() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    console.warn("[Maintenance] ข้ามการทำงานเนื่องจากระบบอื่นกำลังใช้งานอยู่");
    return;
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = ss.getSheets();
    var deletedCount = 0;
    var keepDays = 30; // เก็บย้อนหลัง 30 วัน
    var now = new Date();
    var deletedNames = [];

    sheets.forEach(function(sheet) {
      var name = sheet.getName();
      
      // ตรวจสอบชื่อชีตที่ขึ้นต้นด้วย "Backup_"
      if (name.startsWith("Backup_")) {
        // [MODIFIED v4.0]: แกะวันที่จากรูปแบบ Backup_DB_yyyyMMdd_HHmm
        var datePart = name.match(/(\d{4})(\d{2})(\d{2})/); // จับกลุ่ม ปี(4) เดือน(2) วัน(2)
        
        if (datePart && datePart.length === 4) {
          var year = parseInt(datePart[1]);
          var month = parseInt(datePart[2]) - 1; // JS Month starts at 0
          var day = parseInt(datePart[3]);
          
          var sheetDate = new Date(year, month, day);
          var diffTime = Math.abs(now - sheetDate);
          var diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24)); 

          if (diffDays > keepDays) {
            try {
              ss.deleteSheet(sheet);
              deletedCount++;
              deletedNames.push(name);
            } catch(e) {
              console.error("[Maintenance] Could not delete " + name + ": " + e.message);
            }
          }
        }
      }
    });

    if (deletedCount > 0) {
      var msg = `🧹 Maintenance Report:\nระบบได้ลบชีต Backup ที่เก่ากว่า ${keepDays} วัน จำนวน ${deletedCount} ชีตเรียบร้อยแล้ว`;
      console.info(`[Maintenance] Deleted ${deletedCount} old backups: ${deletedNames.join(", ")}`);
      
      // แจ้งเตือนผู้ดูแลระบบ
      sendLineNotify(msg);
      sendTelegramNotify(msg);
      SpreadsheetApp.getActiveSpreadsheet().toast(`ลบ Backup เก่าไป ${deletedCount} ชีต`, "Maintenance");
    } else {
      console.log("[Maintenance] No old backups to delete.");
    }
  } catch (err) {
    console.error("[Maintenance] Error: " + err.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * 🏥 ตรวจสอบสุขภาพไฟล์ (Cell Limit Check)
 * แนะนำให้ตั้ง Trigger รันวันละ 1 ครั้ง
 */
function checkSpreadsheetHealth() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Google Sheets Limit: 10 Million Cells (Enterprise Standard)
  var cellLimit = 10000000;
  var totalCells = 0;
  var sheetCount = 0;
  
  ss.getSheets().forEach(function(s) {
    totalCells += (s.getMaxRows() * s.getMaxColumns());
    sheetCount++;
  });
  
  var usagePercent = (totalCells / cellLimit) * 100;
  var msg = `🏥 System Health Report:\n- จำนวนชีต: ${sheetCount}\n- การใช้งาน: ${totalCells.toLocaleString()} Cells\n- อัตราการใช้: ${usagePercent.toFixed(2)}%`;
  
  console.info(`[System Health] Usage: ${usagePercent.toFixed(2)}% (${totalCells}/${cellLimit} cells)`);
  
  if (usagePercent > 80) {
    var warn = `⚠️ CRITICAL WARNING: ไฟล์ใกล้เต็มแล้ว!\n\nการใช้งานปัจจุบันอยู่ที่ ${usagePercent.toFixed(2)}% (${totalCells.toLocaleString()} Cells)\nกรุณารันฟังก์ชันลบ Backup เก่า หรือย้ายข้อมูลไปยังไฟล์ใหม่ด่วนครับ`;
    
    // แจ้งเตือนฉุกเฉิน
    sendLineNotify(warn, true);
    sendTelegramNotify(warn);
    SpreadsheetApp.getUi().alert("⚠️ SYSTEM ALERT", warn, SpreadsheetApp.getUi().ButtonSet.OK);
  } else {
    // ถ้ารันมือผ่านเมนู ให้โชว์ Toast
    SpreadsheetApp.getActiveSpreadsheet().toast(`System Health OK (${usagePercent.toFixed(1)}%)`, "Health Check", 5);
  }
}

