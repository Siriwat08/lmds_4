/**
 * VERSION : 4.2 — Phase B
 * [Phase B] อ่าน full DB width (DB_TOTAL_COLS)
 * [Phase B FIXED] แก้ bug return → continue ใน for loop
 * [Phase B] exclude Inactive/Merged อย่างถูกต้อง
 */

/**
 * [Phase C FIXED] searchMasterData()
 * ใช้ resolveUUIDFromMap_() ก่อน return canonical UUID
 * โหลด UUID state ครั้งเดียว ไม่อ่าน Sheet ซ้ำต่อ record
 */
/**
 * [Phase E] searchMasterData()
 * เพิ่ม verified, coordSource, coordConfidence ใน return items
 * เพื่อให้ Index.html แสดง badge ได้
 */
function searchMasterData(keyword, page) {
  console.time("SearchLatency");
  try {
    var pageNum  = parseInt(page) || 1;
    var pageSize = 20;

    if (!keyword || keyword.toString().trim() === "") {
      return { items: [], total: 0, totalPages: 0, currentPage: 1 };
    }

    var rawKey       = keyword.toString().toLowerCase().trim();
    var searchTokens = rawKey.split(/\s+/).filter(function(k) { return k.length > 0; });
    if (searchTokens.length === 0) return { items: [], total: 0, totalPages: 0, currentPage: 1 };

    var ss       = SpreadsheetApp.getActiveSpreadsheet();
    var aliasMap = getCachedNameMapping_(ss);

    var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    if (!sheet) return { items: [], total: 0, totalPages: 0, currentPage: 1 };

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { items: [], total: 0, totalPages: 0, currentPage: 1 };

    var data         = sheet.getRange(2, 1, lastRow - 1, CONFIG.DB_TOTAL_COLS).getValues();
    var uuidStateMap = buildUUIDStateMap_();
    var matches      = [];

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var obj = dbRowToObject(row);

      if (!obj.name) continue;

      var recordStatus = obj.recordStatus || "Active";
      if (recordStatus === "Inactive" || recordStatus === "Merged") continue;

      var aiKeywords = obj.normalized ? obj.normalized.toString().toLowerCase() : "";
      var normName   = normalizeText(obj.name);
      var rawName    = obj.name.toString().toLowerCase();
      var aliases    = obj.uuid ? (aliasMap[obj.uuid] || "") : "";
      var address    = (obj.googleAddr || obj.sysAddr || "").toString().toLowerCase();
      var haystack   = rawName + " " + normName + " " + aliases + " " + aiKeywords + " " + address;

      var isMatch = searchTokens.every(function(token) {
        return haystack.indexOf(token) > -1;
      });

      if (isMatch) {
        var canonicalUuid = obj.uuid
          ? resolveUUIDFromMap_(obj.uuid, uuidStateMap)
          : obj.uuid;

        matches.push({
          name:            obj.name,
          address:         obj.googleAddr || obj.sysAddr || "",
          lat:             obj.lat,
          lng:             obj.lng,
          mapLink:         (obj.lat && obj.lng)
            ? "https://www.google.com/maps/dir/?api=1&destination=" + obj.lat + "," + obj.lng
            : "",
          uuid:            canonicalUuid,
          score:           aiKeywords.includes(rawKey) ? 10 : 1,
          status:          recordStatus,
          // [Phase E NEW] metadata สำหรับ badges
          verified:        obj.verified,
          coordSource:     obj.coordSource     || "",
          coordConfidence: obj.coordConfidence || 0
        });
      }
    }

    matches.sort(function(a, b) { return b.score - a.score; });

    var totalItems = matches.length;
    var totalPages = Math.ceil(totalItems / pageSize);
    if (pageNum > totalPages && totalPages > 0) pageNum = 1;

    var pagedItems = matches.slice((pageNum - 1) * pageSize, pageNum * pageSize);

    console.log("[Search] '" + rawKey + "' | Found: " + totalItems + " | Page: " + pageNum + "/" + totalPages);
    return { items: pagedItems, total: totalItems, totalPages: totalPages, currentPage: pageNum };

  } catch(error) {
    console.error("[Search Error]: " + error.message);
    return { items: [], total: 0, totalPages: 0, currentPage: 1, error: error.message };
  } finally {
    console.timeEnd("SearchLatency");
  }
}

function getCachedNameMapping_(ss) {
  var cache     = CacheService.getScriptCache();
  var cachedMap = cache.get("NAME_MAPPING_JSON_V4");
  if (cachedMap) return JSON.parse(cachedMap);

  var mapSheet = ss.getSheetByName(CONFIG.MAPPING_SHEET);
  var aliasMap = {};

  if (mapSheet && mapSheet.getLastRow() > 1) {
    var mapData = mapSheet.getRange(2, 1, mapSheet.getLastRow() - 1, 2).getValues();
    mapData.forEach(function(row) {
      var variant = row[0];
      var uid     = row[1];
      if (variant && uid) {
        if (!aliasMap[uid]) aliasMap[uid] = "";
        var normVariant = normalizeText(variant);
        aliasMap[uid] += " " + normVariant + " " + variant.toString().toLowerCase();
      }
    });

    try {
      var jsonString = JSON.stringify(aliasMap);
      var byteSize   = Utilities.newBlob(jsonString).getBytes().length;
      if (byteSize < 100000) {
        cache.put("NAME_MAPPING_JSON_V4", jsonString, 3600);
        console.log("[Cache] NameMapping cached (" + byteSize + " bytes)");
      } else {
        console.warn("[Cache] NameMapping too large (" + byteSize + " bytes), skipping");
      }
    } catch(e) {
      console.warn("[Cache Error]: " + e.message);
    }
  }

  return aliasMap;
}

function clearSearchCache() {
  CacheService.getScriptCache().remove("NAME_MAPPING_JSON_V4");
  console.log("[Cache] Search Cache Cleared.");
}
