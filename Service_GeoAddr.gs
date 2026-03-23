/**
 * VERSION: 000
 * 🌍 Service: Geo Address & Google Maps Formulas (Enterprise Edition)
 * Version: 4.0 Omni-Geo Engine & API Hardening
 * -------------------------------------------------------
 * [PRESERVED]: Fully Integrated Google Maps Formulas by Amit Agarwal.
 * [PRESERVED]: 7 Custom Functions for Spreadsheet directly.
 * [MODIFIED v4.0]: Added Try-Catch to _mapsSetCache to prevent 100KB limits crash.
 * [MODIFIED v4.0]: Enterprise Audit Logging (GCP Console) for API Failures.
 * [MODIFIED v4.0]: Enhanced regex in parseAddressFromText for better Tier 2 parsing.
 * [FINAL POLISH]: Bulletproof distance calculation (handling commas in API response).
 * Author: Elite Logistics Architect
 */

// ==========================================
// 1. CONFIGURATION (Internal)
// ==========================================

const POSTAL_COL = {
  ZIP: 0,       // Col A (postcode)
  DISTRICT: 2,  // Col C (district)
  PROVINCE: 3   // Col D (province)
};

var _POSTAL_CACHE = null;

// ==========================================
// 2. 🧠 SMART ADDRESS PARSING LOGIC (Tier 2 Resolution)
// ==========================================

/**
 * แกะรหัสไปรษณีย์ จังหวัด และอำเภอ จากที่อยู่ดิบ
 */
function parseAddressFromText(fullAddress) {
  var result = { province: "", district: "", postcode: "" };
  if (!fullAddress) return result;
  
  var addrStr = fullAddress.toString().trim();
  
  // 1. หารหัสไปรษณีย์ก่อน (ตัวเลข 5 หลักติดกัน)
  var zipMatch = addrStr.match(/(\d{5})/);
  if (zipMatch && zipMatch[1]) {
    result.postcode = zipMatch[1];
  }
  
  // 2. ลองหาจาก Database PostalRef (ถ้ามี)
  var postalDB = getPostalDataCached();
  if (postalDB && result.postcode && postalDB.byZip[result.postcode]) {
    var infoList = postalDB.byZip[result.postcode];
    if (infoList.length > 0) {
       result.province = infoList[0].province;
       result.district = infoList[0].district;
       return result; // ถ้าเจอใน DB จบเลย แม่นยำสุด
    }
  }

  // 3. FALLBACK: ถ้าไม่มี DB หรือหาไม่เจอ ให้ใช้ Regex แกะจาก Text ทันที (อัปเกรด Regex V4.0)
  var provMatch = addrStr.match(/(?:จ\.|จังหวัด)\s*([ก-๙a-zA-Z0-9]+)/i);
  if (provMatch && provMatch[1]) {
    result.province = provMatch[1].trim();
  }
  
  var distMatch = addrStr.match(/(?:อ\.|อำเภอ|เขต)\s*([ก-๙a-zA-Z0-9]+)/i);
  if (distMatch && distMatch[1]) {
    result.district = distMatch[1].trim();
  }

  // Fallback พิเศษสำหรับ กทม.
  if (!result.province && (addrStr.includes("กรุงเทพ") || addrStr.includes("Bangkok") || addrStr.includes("กทม"))) {
    result.province = "กรุงเทพมหานคร";
  }

  return result;
}

function getPostalDataCached() {
  if (_POSTAL_CACHE) return _POSTAL_CACHE;
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = (typeof CONFIG !== 'undefined' && CONFIG.SHEET_POSTAL) ? CONFIG.SHEET_POSTAL : "PostalRef";
  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) return null; 
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  
  var data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  var db = { byZip: {} };
  
  data.forEach(function(row) {
    if (row.length <= POSTAL_COL.PROVINCE) return;
    
    var pc = String(row[POSTAL_COL.ZIP]).trim(); 
    if (!pc) return;

    if (!db.byZip[pc]) db.byZip[pc] = [];
    db.byZip[pc].push({ 
      postcode: pc, 
      district: row[POSTAL_COL.DISTRICT], 
      province: row[POSTAL_COL.PROVINCE] 
    });
  });
  
  _POSTAL_CACHE = db;
  return db;
}
/**
 * [NEW v4.1] ล้าง Postal Cache อย่างปลอดภัย
 * เรียกจาก Menu.gs ผ่าน clearPostalCache_UI()
 */
function clearPostalCache() {
  _POSTAL_CACHE = null;
  console.log("[Cache] Postal Cache cleared.");
}

// ==========================================
// 3. 🗺️ GOOGLE MAPS FORMULAS (Amit Agarwal)
// ==========================================

const _mapsMd5 = (key = "") => {
  const code = key.toLowerCase().replace(/\s/g, "");
  return Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, key)
    .map((char) => (char + 256).toString(16).slice(-2))
    .join("");
};

const _mapsGetCache = (key) => {
  try {
    return CacheService.getDocumentCache().get(_mapsMd5(key));
  } catch(e) {
    return null;
  }
};

/**
 * [MODIFIED v4.0]: ป้องกัน Error กรณี String ของ Maps Directions เกิน 100KB
 */
const _mapsSetCache = (key, value) => {
  try {
    const expirationInSeconds = (typeof CONFIG !== 'undefined' && CONFIG.CACHE_EXPIRATION) ? CONFIG.CACHE_EXPIRATION : 21600; // 6 hours
    if (value && value.toString().length < 90000) { 
       CacheService.getDocumentCache().put(_mapsMd5(key), value, expirationInSeconds);
    }
  } catch (e) {
    console.warn("[Geo Cache Warn]: Could not cache key " + key + " - " + e.message);
  }
};

/**
 * 2.3 Calculate the travel time between two locations on Google Maps.
 * @customFunction
 */
const GOOGLEMAPS_DURATION = (origin, destination, mode = "driving") => {
  if (!origin || !destination) throw new Error("No address specified!");
  if (origin.map) return origin.map(o => GOOGLEMAPS_DURATION(o, destination, mode));
  
  const key = ["duration", origin, destination, mode].join(",");
  const value = _mapsGetCache(key);
  if (value !== null) return value;

  Utilities.sleep(150); // API Throttling protection
  const { routes: [data] = [] } = Maps.newDirectionFinder()
    .setOrigin(origin)
    .setDestination(destination)
    .setMode(mode)
    .getDirections();
  
  if (!data) throw new Error("No route found!");
  
  const { legs: [{ duration: { text: time } } = {}] = [] } = data;
  _mapsSetCache(key, time);
  return time;
};

/**
 * 2.1 Calculate the distance between two locations on Google Maps.
 * @customFunction
 */
const GOOGLEMAPS_DISTANCE = (origin, destination, mode = "driving") => {
  if (!origin || !destination) throw new Error("No address specified!");
  if (origin.map) return origin.map(o => GOOGLEMAPS_DISTANCE(o, destination, mode));
  
  const key = ["distance", origin, destination, mode].join(",");
  const value = _mapsGetCache(key);
  if (value !== null) return value;

  Utilities.sleep(150);
  const { routes: [data] = [] } = Maps.newDirectionFinder()
    .setOrigin(origin)
    .setDestination(destination)
    .setMode(mode)
    .getDirections();
    
  if (!data) throw new Error("No route found!");
  
  const { legs: [{ distance: { text: distance } } = {}] = [] } = data;
  _mapsSetCache(key, distance);
  return distance;
};

/**
 * 2.4 Get the latitude and longitude of any address on Google Maps.
 * @customFunction
 */
const GOOGLEMAPS_LATLONG = (address) => {
  if (!address) throw new Error("No address specified!");
  if (address.map) return address.map(a => GOOGLEMAPS_LATLONG(a));
  
  const key = ["latlong", address].join(",");
  const value = _mapsGetCache(key);
  if (value !== null) return value;

  Utilities.sleep(150);
  const { results: [data = null] = [] } = Maps.newGeocoder().geocode(address);
  if (data === null) throw new Error("Address not found!");
  
  const { geometry: { location: { lat, lng } } = {} } = data;
  const answer = `${lat}, ${lng}`;
  _mapsSetCache(key, answer);
  return answer;
};

/**
 * 2.5 Get the full address of any zip code or partial address on Google Maps.
 * @customFunction
 */
const GOOGLEMAPS_ADDRESS = (address) => {
  if (!address) throw new Error("No address specified!");
  if (address.map) return address.map(a => GOOGLEMAPS_ADDRESS(a));
  
  const key = ["address", address].join(",");
  const value = _mapsGetCache(key);
  if (value !== null) return value;

  Utilities.sleep(150);
  const { results: [data = null] = [] } = Maps.newGeocoder().geocode(address);
  if (data === null) throw new Error("Address not found!");
  
  const { formatted_address } = data;
  _mapsSetCache(key, formatted_address);
  return formatted_address;
};

/**
 * 2.2 Use Reverse Geocoding to get the address of a point location.
 * @customFunction
 */
const GOOGLEMAPS_REVERSEGEOCODE = (latitude, longitude) => {
  if (!latitude || !longitude) throw new Error("Lat/Lng not specified!");
  
  const key = ["reverse", latitude, longitude].join(",");
  const value = _mapsGetCache(key);
  if (value !== null) return value;

  Utilities.sleep(150);
  const { results: [data = {}] = [] } = Maps.newGeocoder().reverseGeocode(latitude, longitude);
  const { formatted_address } = data;
  if (!formatted_address) return "Address not found";
  
  _mapsSetCache(key, formatted_address);
  return formatted_address;
};

/**
 * 2.6 Get the country name of an address on Google Maps.
 * @customFunction
 */
const GOOGLEMAPS_COUNTRY = (address) => {
  if (!address) throw new Error("No address specified!");
  if (address.map) return address.map(a => GOOGLEMAPS_COUNTRY(a));

  const key = ["country", address].join(",");
  const value = _mapsGetCache(key);
  if (value !== null) return value;

  Utilities.sleep(150);
  const { results: [data = null] = [] } = Maps.newGeocoder().geocode(address);
  if (data === null) throw new Error("Address not found!");
  
  const [{ short_name, long_name } = {}] = data.address_components.filter(
    ({ types: [level] }) => level === "country"
  );
  if (!short_name) throw new Error("Country not found!");
  
  const answer = `${long_name} (${short_name})`;
  _mapsSetCache(key, answer);
  return answer;
};

/**
 * 2.7 Find the driving direction between two locations on Google Maps.
 * @customFunction
 */
const GOOGLEMAPS_DIRECTIONS = (origin, destination, mode = "driving") => {
  if (!origin || !destination) throw new Error("No address specified!");
  
  const key = ["directions", origin, destination, mode].join(",");
  const value = _mapsGetCache(key);
  if (value !== null) return value;

  Utilities.sleep(150);
  const { routes = [] } = Maps.newDirectionFinder()
    .setOrigin(origin)
    .setDestination(destination)
    .setMode(mode)
    .getDirections();
    
  if (!routes.length) throw new Error("No route found!");
  
  const directions = routes
    .map(({ legs }) => {
      return legs.map(({ steps }) => {
        return steps.map((step) => {
          return step.html_instructions
            .replace("><", "> <")
            .replace(/<[^>]+>/g, "");
        });
      });
    })
    .join(", ");
    
  _mapsSetCache(key, directions);
  return directions;
};

// ==========================================
// 4. 🔗 BACKEND INTEGRATION (System Calls V4.0)
// ==========================================

/**
 * Wrapper for Backend System: Reverse Geocode
 * ดึงพิกัด Lat, Lng มาแปลเป็นที่อยู่
 */
function GET_ADDR_WITH_CACHE(lat, lng) {
  try {
    return GOOGLEMAPS_REVERSEGEOCODE(lat, lng);
  } catch (e) {
    console.error(`[GeoAddr API] Reverse Geocode Error (${lat}, ${lng}): ${e.message}`);
    return "";
  }
}

/**
 * Wrapper for Backend System: Calculate Distance
 * ดึงระยะทางจากสูตรคุณ Amit แล้วแปลง "1,250.5 km" ให้เหลือแค่ "1250.50" (ตัวเลขล้วน)
 */
function CALCULATE_DISTANCE_KM(origin, destination) {
  try {
    var distanceText = GOOGLEMAPS_DISTANCE(origin, destination, "driving");
    if (!distanceText) return "";
    
    // [FINAL POLISH] กำจัดลูกน้ำ (,) ออกก่อน แล้วค่อยกรองเฉพาะตัวเลขและจุดทศนิยม
    var cleanStr = String(distanceText).replace(/,/g, "").replace(/[^0-9.]/g, "");
    var val = parseFloat(cleanStr);
    
    return isNaN(val) ? "" : val.toFixed(2);
  } catch (e) {
    console.error(`[GeoAddr API] Distance Error (${origin} -> ${destination}): ${e.message}`);
    return "";
  }
}



