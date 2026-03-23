/**
 * VERSION: 000
 * 🔍 Service: Search Engine (Enterprise Edition)
 * Version: 4.0 Omni-Search (UUID & AI Integrated)
 * ----------------------------------------------
 * [PRESERVED]: Multi-Token search logic and Pagination structure.
 * [MODIFIED v4.0]: Upgraded NameMapping cache to use Master_UID instead of Name.
 * [MODIFIED v4.0]: Added try-catch around CacheService to prevent 100KB limit crash.
 * [MODIFIED v4.0]: Added Enterprise Performance Logging (console.time).
 * Author: Elite Logistics Architect
 */

function searchMasterData(keyword, page) {
  console.time("SearchLatency");
  try {
    // 1. Input Validation & Setup
    var pageNum = parseInt(page) || 1;
    var pageSize = 20;

    if (!keyword || keyword.toString().trim() === "") {
      return { items: [], total: 0, totalPages: 0, currentPage: 1 };
    }
    
    // Prepare Keywords (Split by space for multi-token match)
    // Example: "SCG Rayong" -> ["scg", "rayong"]
    var rawKey = keyword.toString().toLowerCase().trim();
    var searchTokens = rawKey.split(/\s+/).filter(function(k) { return k.length > 0; });
    
    if (searchTokens.length === 0) return { items: [], total: 0, totalPages: 0, currentPage: 1 };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 2. [UPGRADED v4.0] Load NameMapping (With Smart Cache via UUID)
    var aliasMap = getCachedNameMapping_(ss);

    // 3. Load Database
    var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    if (!sheet) return { items: [], total: 0, totalPages: 0, currentPage: 1 };

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { items: [], total: 0, totalPages: 0, currentPage: 1 };

    // Read Data
    var data = sheet.getRange(2, 1, lastRow - 1, CONFIG.COL_RECORD_STATUS).getValues();
    var matches = []; 

    // 4. Search Algorithm (Linear Scan with Token Logic)
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      
      var name = row[CONFIG.C_IDX.NAME];
      if (!name) continue;

      var address = row[CONFIG.C_IDX.GOOGLE_ADDR] || row[CONFIG.C_IDX.SYS_ADDR] || "";
      var lat = row[CONFIG.C_IDX.LAT];
      var lng = row[CONFIG.C_IDX.LNG];
      var uuid = row[CONFIG.C_IDX.UUID]; // [ADDED v4.0]: Use UUID for relational link
      
      // AI Brain: ดึงข้อมูลที่ Agent คิดไว้มาช่วยค้นหา (Tag [AI])
      var aiKeywords = row[CONFIG.C_IDX.NORMALIZED] ? row[CONFIG.C_IDX.NORMALIZED].toString().toLowerCase() : "";
      var normName = typeof normalizeText === 'function' ? normalizeText(name) : name.toString().toLowerCase();
      var rawName = name.toString().toLowerCase();
      
      // [UPGRADED v4.0]: Alias Lookup using UUID instead of Name
      var aliases = uuid ? (aliasMap[uuid] || "") : "";
      
      // Combine all searchable text into one "Haystack"
      var haystack = (rawName + " " + normName + " " + aliases + " " + aiKeywords + " " + address.toString().toLowerCase());
      
      // Multi-Token Check: ต้องเจอ "ทุกคำ" ที่พิมพ์มา (AND Logic)
     var recordStatus = row[CONFIG.C_IDX.RECORD_STATUS] || "Active";

// [v4.1] กรอง Inactive และ Merged ออกจากผลค้นหา
if (recordStatus === "Inactive" || recordStatus === "Merged") return;

var isMatch = searchTokens.every(function(token) { return haystack.indexOf(token) > -1; });

if (isMatch) {
  matches.push({
    name: name, address: address, lat: lat, lng: lng,
    mapLink: (lat && lng) ? "https://www.google.com/maps/dir/?api=1&destination=" + lat + "," + lng : "",
    uuid: uuid,
    score: aiKeywords.includes(rawKey) ? 10 : 1,
    status: recordStatus  // [v4.1] ส่งสถานะมาด้วย
  });
}
    }

    // [Optional] Sort by score (AI exact matches first)
    matches.sort(function(a, b) { return b.score - a.score; });

    // 5. Pagination Logic
    var totalItems = matches.length;
    var totalPages = Math.ceil(totalItems / pageSize);
    
    if (pageNum > totalPages && totalPages > 0) pageNum = 1;
    
    var startIndex = (pageNum - 1) * pageSize;
    var endIndex = startIndex + pageSize;
    var pagedItems = matches.slice(startIndex, endIndex);

    console.log(`[Search] Query: "${rawKey}" | Found: ${totalItems} | Page: ${pageNum}/${totalPages}`);
    return {
      items: pagedItems,
      total: totalItems,
      totalPages: totalPages,
      currentPage: pageNum
    };

  } catch (error) {
    console.error("[Search Error]: " + error.message);
    return { items: [], total: 0, totalPages: 0, currentPage: 1, error: error.message };
  } finally {
    console.timeEnd("SearchLatency");
  }
}

/**
 * 🛠️ Internal Helper: Get NameMapping with Caching
 * [UPGRADED v4.0]: Relational mapping using Variant -> UID
 */
function getCachedNameMapping_(ss) {
  var cache = CacheService.getScriptCache();
  var cachedMap = cache.get("NAME_MAPPING_JSON_V4");
  
  if (cachedMap) {
    return JSON.parse(cachedMap);
  }
  
  // ถ้าไม่มีใน Cache ให้โหลดจาก Sheet
  var mapSheet = ss.getSheetByName(CONFIG.MAPPING_SHEET);
  var aliasMap = {}; 
  
  if (mapSheet && mapSheet.getLastRow() > 1) {
    // โหลด 2 คอลัมน์แรก (Col A: Variant, Col B: UID) ตามโครงสร้าง V4.0
    var mapData = mapSheet.getRange(2, 1, mapSheet.getLastRow() - 1, 2).getValues();
    
    mapData.forEach(function(row) {
      var variant = row[0]; // Variant_Name
      var uid = row[1];     // Master_UID
      
      if (variant && uid) {
        if (!aliasMap[uid]) aliasMap[uid] = "";
        
        // ต่อ String Variant Name เก็บไว้ใน Key ของ UID
        var normVariant = typeof normalizeText === 'function' ? normalizeText(variant) : variant.toString().toLowerCase();
        aliasMap[uid] += " " + normVariant + " " + variant.toString().toLowerCase();
      }
    });
    
    // Save to Cache (Duration: 1 hour)
    // ป้องกัน Error 100KB Limit ของ Google Cache
    try {
  var jsonString = JSON.stringify(aliasMap);

  // [FIXED v4.1] วัด Bytes จริงแทน .length เพราะภาษาไทย 1 ตัว = 3 Bytes
  var byteSize = Utilities.newBlob(jsonString).getBytes().length;

  if (byteSize < 100000) {
    cache.put("NAME_MAPPING_JSON_V4", jsonString, 3600);
    console.log("[Cache] NameMapping cached successfully (" + byteSize + " bytes)");
  } else {
    console.warn("[Cache] NameMapping too large (" + byteSize + " bytes), skipping cache.");
  }
} catch (e) {
  console.warn("[Cache Error]: " + e.message);
}
  }
  
  return aliasMap;
}

/**
 * [Optional] Function to clear cache if Mapping is updated
 * Call this when running 'finalizeAndClean'
 */
function clearSearchCache() {
  CacheService.getScriptCache().remove("NAME_MAPPING_JSON_V4");
  console.log("[Cache] Search Cache Cleared.");
}


