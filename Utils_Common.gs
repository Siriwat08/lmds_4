/**
 * VERSION: 000
 * 🛠️ Utilities: Common Helper Functions
 * Version: 4.0 Enterprise Edition (AI & Batch Preparedness)
 * ------------------------------------------------------
 * [PRESERVED]: Hashing, Haversine Math, Fuzzy Matching, and Smart Naming.
 * [ADDED v4.0]: chunkArray() helper for AI Batch Processing.
 * [MODIFIED v4.0]: Enhanced normalizeText() with more logistics-specific stop words.
 * [MODIFIED v4.0]: genericRetry() upgraded with Enterprise-grade console logging.
 * Author: Elite Logistics Architect
 */

// ====================================================
// 1. Hashing & ID Generation
// ====================================================

function md5(key) {
  if (!key) return "empty_hash";
  var code = key.toString().toLowerCase().replace(/\s/g, "");
  return Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, code)
    .map(function(char) { return (char + 256).toString(16).slice(-2); })
    .join("");
}

function generateUUID() {
  return Utilities.getUuid();
}

// ====================================================
// 2. Text Processing & Normalization
// ====================================================

/**
 * [MODIFIED v4.0]: เพิ่ม Stop words สำหรับงาน Logistics (โกดัง, คลังสินค้า, อาคาร ฯลฯ)
 * ทำหน้าที่เป็น Tier 2 Resolution (Clean Text)
 */
function normalizeText(text) {
  if (!text) return "";
  var clean = text.toString().toLowerCase();
  
  var stopWordsPattern = /บริษัท|บจก\.?|บมจ\.?|หจก\.?|ห้างหุ้นส่วน|จำกัด|มหาชน|ส่วนบุคคล|ร้าน|ห้าง|สาขา|สำนักงานใหญ่|store|shop|company|co\.?|ltd\.?|inc\.?|จังหวัด|อำเภอ|ตำบล|เขต|แขวง|ถนน|ซอย|นาย|นาง|นางสาว|โกดัง|คลังสินค้า|หมู่ที่|หมู่|อาคาร|ชั้น/g;
  clean = clean.replace(stopWordsPattern, "");

  return clean.replace(/[^a-z0-9\u0E00-\u0E7F]/g, "");
}

function cleanDistance(val) {
  if (!val && val !== 0) return "";
  var str = val.toString().replace(/[^0-9.]/g, ""); 
  var num = parseFloat(str);
  return isNaN(num) ? "" : num.toFixed(2);
}

// ====================================================
// 3. 🧠 Smart Naming Logic
// ====================================================

function getBestName_Smart(names) {
  if (!names || names.length === 0) return "";
  
  var nameScores = {};
  var bestName = names[0];
  var maxScore = -9999;
  
  names.forEach(function(n) {
    if (!n) return;
    var original = n.toString().trim();
    if (original === "") return;

    if (!nameScores[original]) {
       nameScores[original] = { count: 0, score: 0 };
    }
    nameScores[original].count += 1;
  });

  for (var n in nameScores) {
    var s = nameScores[n].count * 10; 
    
    if (/(บริษัท|บจก|หจก|บมจ)/.test(n)) s += 5; 
    if (/(จำกัด|มหาชน)/.test(n)) s += 5;        
    if (/(สาขา)/.test(n)) s += 5;               
    
    var openBrackets = (n.match(/\(/g) || []).length;
    var closeBrackets = (n.match(/\)/g) || []).length;
    
    if (openBrackets > 0 && openBrackets === closeBrackets) {
      s += 5; 
    } else if (openBrackets !== closeBrackets) {
      s -= 30; 
    }
    
    if (/[0-9]{9,10}/.test(n) || /โทร/.test(n)) s -= 30; 
    if (/ส่ง|รับ|ติดต่อ/.test(n)) s -= 10;                
    
    var len = n.length;
    if (len > 70) {
      s -= (len - 70); 
    } else if (len < 5) {
      s -= 10;         
    } else {
      s += (len * 0.1);
    }

    nameScores[n].score = s;
    
    if (s > maxScore) {
      maxScore = s;
      bestName = n;
    }
  }
  
  return cleanDisplayName(bestName);
}

function cleanDisplayName(name) {
  var clean = name.toString();
  clean = clean.replace(/\s*โทร\.?\s*[0-9-]{9,12}/g, '');
  clean = clean.replace(/\s*0[0-9]{1,2}-[0-9]{3}-[0-9]{4}/g, '');
  clean = clean.replace(/\s+/g, ' ').trim();
  return clean;
}

// ====================================================
// 4. Geo Math & Fuzzy Matching
// ====================================================

function getHaversineDistanceKM(lat1, lon1, lat2, lon2) {
  if (!lat1 || !lon1 || !lat2 || !lon2) return null;
  var R = 6371; 
  var dLat = (lat2 - lat1) * Math.PI / 180;
  var dLon = (lon2 - lon1) * Math.PI / 180;
  var a = Math.sin(dLat/2) * Math.sin(dLat/2) +
          Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) *
          Math.sin(dLon/2) * Math.sin(dLon/2);
  var c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));
  return parseFloat((R * c).toFixed(3)); 
}


// ====================================================
// 5. System Utilities (Logging, Retry & Array Ops)
// ====================================================

/**
 * [MODIFIED v4.0]: Enterprise Logging
 */
function genericRetry(func, maxRetries) {
  for (var i = 0; i < maxRetries; i++) {
    try { return func(); } 
    catch (e) {
      if (i === maxRetries - 1) {
        console.error("[GenericRetry] FATAL ERROR after " + maxRetries + " attempts: " + e.message);
        throw e;
      }
      Utilities.sleep(1000 * Math.pow(2, i)); 
      console.warn("[GenericRetry] Attempt " + (i + 1) + " failed: " + e.message + ". Retrying...");
    }
  }
}

function safeJsonParse(str) {
  try { return JSON.parse(str); } catch (e) { return null; }
}


function checkUnusedFunctions() {
  var ui = SpreadsheetApp.getUi();
  
  var funcs = [
    'calculateSimilarity',
    'editDistance', 
    'cleanPhoneNumber',
    'parseThaiDate',
    'chunkArray'
  ];
  
  console.log("=== ตรวจสอบฟังก์ชันที่ไม่ได้ใช้ ===");
  funcs.forEach(function(name) {
    var exists = typeof eval(name) === 'function';
    console.log(name + ": " + (exists ? "✅ มีอยู่" : "❌ ไม่พบ"));
  });
  
  console.log("\nถ้าทุกตัวแสดง ✅ มีอยู่ แสดงว่าพร้อมลบได้ครับ");
}

function verifyFunctionsRemoved() {
  var funcs = [
    'calculateSimilarity',
    'editDistance',
    'cleanPhoneNumber', 
    'parseThaiDate',
    'chunkArray'
  ];
  
  var allRemoved = true;
  
  funcs.forEach(function(name) {
    try {
      var result = eval('typeof ' + name);
      if (result === 'function') {
        console.log("⚠️ " + name + " ยังอยู่ → ลบไม่สำเร็จ");
        allRemoved = false;
      } else {
        console.log("✅ " + name + " ลบออกแล้ว");
      }
    } catch(e) {
      console.log("✅ " + name + " ลบออกแล้ว");
    }
  });
  
  if (allRemoved) {
    console.log("\n✅ ลบครบทุกฟังก์ชันแล้วครับ");
  } else {
    console.log("\n⚠️ ยังมีฟังก์ชันที่ลบไม่สำเร็จ ตรวจสอบอีกครั้งครับ");
  }
}