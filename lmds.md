# วิเคราะห์ระบบแบบละเอียด

# **วิเคราะห์ระบบแบบละเอียด**

## **1\) ภาพรวมระบบ**

repo นี้คือ Google Apps Script \+ Google Sheets based Logistics Master Data System สำหรับเก็บ “ฐานข้อมูลกลาง” ของปลายทางลูกค้า, รับข้อมูล shipment จาก SCG, ทำการ match พิกัด, ทำ normalized name / mapping, ใช้ AI ช่วย match ชื่อที่ตกหล่น, และเปิด WebApp สำหรับค้นหาพิกัดลูกค้าได้จากมือถือหรือ browser. ภาพรวมนี้อธิบายไว้ชัดเจนใน README และแบ่งไฟล์เป็น Configuration, Utilities, Core Services, AI/Automation, Notifications, UI/Menu, Setup และ Testing. 

หัวใจของระบบมี 4 ส่วนใหญ่:

1. Master Database (Database) เป็น golden record ของลูกค้า.   
2. NameMapping (NameMapping) ใช้ map alias/variant name ไป UUID กลาง.   
3. SCG daily operation (Input, Data, SCGนครหลวงJWDภูมิภาค) ใช้ดึง shipment และ GPS จริงจาก operation.   
4. Web search/UI ผ่าน WebApp.gs \+ Index.html \+ Service\_Search.gs. 

---

## **2\) โครงสร้างสถาปัตยกรรมเชิงโมดูล**

### **2.1 ชั้น Configuration**

Config.gs เป็น Single Source of Truth ของชื่อชีต, index คอลัมน์, AI config, threshold ต่าง ๆ, และยังนิยาม schema ฝั่ง SCG ด้วย. จุดนี้สำคัญมาก เพราะแทบทุก service อ้าง index จากที่นี่. 

### **2.2 ชั้น Utility**

Utils\_Common.gs และ Service\_GeoAddr.gs เป็น utility layer:

* Utils\_Common.gs ดูแล text normalization, UUID, Haversine, retry, JSON parse.   
* Service\_GeoAddr.gs ดูแล postal cache, parse address, reverse geocode, distance, maps cache. 

### **2.3 ชั้น Core Domain**

* Service\_Master.gs \= lifecycle ของฐานข้อมูลกลาง: sync ข้อมูลใหม่, deep clean, clustering, finalize, repair, recalculate.   
* Service\_SCG.gs \= ฝั่ง daily logistics operation: fetch shipment จาก API, map master coordinates, สร้าง summary และ clear sheets.   
* Service\_Search.gs \= search engine สำหรับ WebApp. 

### **2.4 ชั้น Workflow Safety**

* Service\_SchemaValidator.gs ป้องกัน run flow ผิด schema.   
* Service\_SoftDelete.gs ทำ soft delete / merge UUID. 

### **2.5 ชั้น AI & Automation**

* Service\_AutoPilot.gs \= scheduler \+ AI indexing background.   
* Service\_Agent.gs \= AI tier-4 matching สำหรับชื่อที่หา geo ไม่เจอใน daily data. 

### **2.6 ชั้น UI / Entry Points**

* Menu.gs \= menu ทั้งหมดใน Google Sheets.   
* WebApp.gs \= HTTP entry points (doGet, doPost).   
* Index.html \= frontend search page. 

---

# **3\) วิเคราะห์ทีละไฟล์ / ทีละโมดูล / ทีละฟังก์ชัน**

---

## **README.md**

README ทำหน้ที่เป็นทั้ง onboarding document และ release note ของ V4.1:

* ระบุรายการไฟล์ทั้งหมด 21 ไฟล์ พร้อมหน้าที่แต่ละไฟล์.   
* ระบุชีตที่ระบบต้องมี เช่น Database, NameMapping, GPS\_Queue, PostalRef, Data, Input.   
* ระบุขั้นตอนติดตั้ง เช่น setupEnvironment(), createGPSQueueSheet(), initializeRecordStatus(), runFullSchemaValidation().   
* สรุป bug fix / feature ใหม่ของ v4.1 เช่น SYNC checkpoint, soft delete, schema validation, queue management. 

มุมมองเชิงวิศวกรรม: README นี้สะท้อนว่า V4.1 เน้น “ความปลอดภัยของ flow” มากกว่า V เดิม เช่น guard schema, soft delete, anti-duplicate, checkpoint. 

---

## **Config.gs**

### CONFIG

เป็น object กลางที่เก็บ:

* ชื่อชีตหลัก.   
* secure getter สำหรับ GEMINI\_API\_KEY ที่ throw error ถ้ายังไม่ได้ setup.   
* AI settings เช่น model และ batch size.   
* depot location, threshold, limit, timeout, cache expiration.   
* ดัชนีคอลัมน์ Database แบบ 1-based.   
* computed index maps แบบ 0-based ผ่าน C\_IDX และ MAP\_IDX. 

ความสำคัญระดับบรรทัด:  
บรรทัด 77-113 เป็นแกนของ codebase เพราะทำให้ทุก service ใช้ symbolic field name แทน magic number เวลา access array จากชีต. ถ้าจะ refactor ระบบนี้ จุดนี้คือหัวใจที่สุด. 

### SCG\_CONFIG

นิยาม:

* ชื่อชีต daily operation.  
* API URL ของ SCG.  
* input positions เช่น B1, B3.  
* GPS queue sheet.  
* threshold 50 เมตร.  
* source column index ฝั่ง SCGนครหลวงJWDภูมิภาค.  
* SYNC status column AK. 

### CONFIG.validateSystemIntegrity

เช็ก:

* sheet จำเป็นต้องมี.  
* API key มีและ format ดูสมเหตุผล.  
* ถ้า fail จะ throw แบบรวมข้อความทั้งหมด. 

จุดเด่น: เป็น health check ระดับ config/bootstrap ที่เหมาะกับเมนู admin. 

---

## **Utils\_Common.gs**

### md5(key)

hash helper สำหรับ key normalization. ถ้าไม่มี key จะคืน "empty\_hash". ใช้ digest ของ GAS. 

### generateUUID()

wrapper ตรง ๆ ของ Utilities.getUuid(). เรียบง่ายและถูกใช้หลายที่. 

### normalizeText(text)

นี่คือฟังก์ชันที่มีผลต่อทั้งระบบมากที่สุดตัวหนึ่ง:

* lowercase  
* ลบ stop words ไทย/อังกฤษ เช่น บริษัท, จำกัด, store, shop, จังหวัด, อำเภอ ฯลฯ  
* ลบอักขระที่ไม่ใช่ไทย/อังกฤษ/ตัวเลข. 

ผลเชิงพฤติกรรม:  
มันทำให้ชื่อ “บริษัท เอ บี ซ จำกัด สาขาเชียงใหม่” กับ “ABC” มีโอกาสใกล้กันขึ้น แต่ก็มี trade-off คือคำว่า “สาขา” ถูกตัดออก ทำให้บางชื่อ branch สูญเสีย context ได้. 

### cleanDistance(val)

ทำความสะอาดค่า distance ให้อยู่ในรูปเลขทศนิยม 2 ตำแหน่ง. 

### getBestName\_Smart(names)

heuristic scorer สำหรับเลือกชื่อที่ “น่าจะเป็นชื่อมาตรฐานที่สุด”:

* นับ frequency.  
* โบนัสสำหรับ pattern ทางกฎหมายธุรกิจ.  
* โบนัสถ้ามีวงเล็บสมบูรณ์.  
* หักคะแนนถ้ามีเบอร์โทร / คำสั่งงาน / สั้นเกิน / ยาวเกิน. 

นี่คือ logic clustering winner selection ของระบบ. มันไม่ใช่ ML แต่เป็น rules-based ranking ที่ practical มากกับข้อมูลโลจิสติกส์ไทย. 

### cleanDisplayName(name)

ลบเบอร์โทรและ spacing เกินจากชื่อที่จะแสดงผล. 

### getHaversineDistanceKM(...)

คำนวณระยะตรงโลกกลมแบบ Haversine และปัด 3 ตำแหน่ง. ใช้ใน duplicate detection, sync diff, clustering. 

### genericRetry(func, maxRetries)

exponential backoff generic. ใช้ใน AI/API path. 

### safeJsonParse(str)

JSON parse แบบไม่ throw. utility เล็กแต่ดี. 

---

## **Service\_Master.gs**

นี่คือแกนของ “master data governance”.

### getRealLastRow\_(sheet, columnIndex)

วนจากท้ายขึ้นมาเพื่อหาแถวสุดท้ายที่มี data จริง โดยไม่โดน checkbox/boolean หลอก. แก้ปัญหา getLastRow() ใน Google Sheets ได้ดีมาก. 

### syncNewDataToMaster()

flow ใหญ่ที่สุดตัวหนึ่ง:

1. lock script ก่อน.   
2. run preCheck\_Sync() ตรวจ schema.   
3. โหลด Database \+ GPS\_Queue \+ NameMapping เข้า memory.   
4. อ่าน source sheet.   
5. วนทุกแถว:  
   * ข้ามถ้า SYNCED.   
   * ถ้าไม่มีชื่อหรือ lat/lng ไม่สมบูรณ์ ข้าม.   
   * หา match จาก normalized name หรือ alias→UUID.   
   * ถ้าไม่เจอ \= สร้างลูกค้าใหม่เข้า Database.   
   * ถ้าเจอแต่ DB ไม่มีพิกัด \= ส่งเข้า GPS\_Queue ด้วย reason DB\_NO\_GPS.   
   * ถ้าระยะต่าง ≤ 50m \= อัปเดตแค่ timestamp.   
   * ถ้าต่าง \> 50m \= ส่งเข้า queue ด้วย reason GPS\_DIFF.   
6. เขียนผลกลับ DB/Queue.   
7. mark source row เป็น SYNCED. 

เชิงออกแบบ: ฟังก์ชันนี้ทำหน้าที่เป็น “ingestion pipeline \+ triage engine” มากกว่าจะเป็นแค่ sync. มันแบ่งข้อมูลเข้า 3 lane: new record, timestamp refresh, human approval queue. 

จุดแข็ง:

* มี lock.   
* มี schema precheck.   
* ใช้ UUID \+ alias map.   
* มี checkpoint กันประมวลผลซ้ำ. 

### cleanDistance\_Helper(val)

helper เล็กสำหรับ parse distance text. 

### updateGeoData\_SmartCache()

alias เรียก runDeepCleanBatch\_100(). เป็น API façade มากกว่า function จริง. 

### autoGenerateMasterList\_Smart()

alias เรียก clustering. 

### runDeepCleanBatch\_100()

ทำ batch cleanup ทีละช่วง:

* ใช้ property DEEP\_CLEAN\_POINTER จำตำแหน่งค้าง.   
* อ่านข้อมูล 22 คอลัมน์.   
* เติม reverse-geocoded address ถ้ายังไม่มี.   
* เติม distance จาก depot ถ้ายังไม่มี.   
* เติม UUID / created date ถ้ายังไม่มี.   
* parse province/district/postcode จาก Google address.   
* คำนวณ QUALITY score ตาม heuristic.   
* กำหนด Record\_Status \= Active ถ้ายังไม่มี.   
* เขียนกลับและอัปเดต pointer. 

นี่คือ data enrichment engine ของ master DB. 

### resetDeepCleanMemory()

ล้าง pointer. 

### finalizeAndClean\_MoveToMapping()

flow “จบงาน”:

* lock script.   
* อ่านข้อมูล master จริงด้วย getRealLastRow\_.   
* สร้าง uuidMap จากชื่อและ suggested name.   
* backup sheet ก่อนแก้จริง.   
* เก็บ rows ที่ verified ไว้.   
* สำหรับแถว unverified แต่มี suggested ต่างจาก raw name จะ upload mapping เข้า NameMapping.   
* clear ข้อมูลเดิม แล้วเขียนเฉพาะ rows ที่ keep.   
* delete ghost rows. 

มุมมอง business logic: verified rows อยู่ในฐานหลักต่อ, ชื่อแปรผันถูกย้ายไปทำ alias mapping แทน. นี่คือแนวคิด “golden record \+ alias resolution”. 

### assignMissingUUIDs()

เติม UUID เฉพาะคอลัมน์ UUID ที่ว่าง. simple แต่จำเป็น. 

### repairNameMapping\_Full()

อ่าน NameMapping, dedupe ตาม normalized variant, เติม default metadata, และเขียน clean list กลับ. 

ข้อสังเกตสำคัญ: บรรทัด 496 ใช้ uuidMap\[normalizeText(r\[1\])\] โดย r\[1\] คือ UID column ไม่ใช่ชื่อ จึงดูมีความเป็นไปได้ว่าเป็น logic ที่สับสน/ไม่ค่อยมีผลตามเจตนา. ถ้าตั้งใจ fallback จากชื่อ variant น่าจะควร lookup ด้วย oldN มากกว่า. 

### processClustering\_GridOptimized()

จับกลุ่มชื่อด้วยพิกัด:

* bucket ตาม grid 0.1 degree.   
* verified row ทำตัวเป็น locked cluster center.   
* แถว unverified จะพยายามเข้าคลัสเตอร์ถ้าพิกัดใกล้ภายใน threshold.   
* ใช้ getBestName\_Smart() หา winner.   
* คำนวณ confidence เป็นเปอร์เซ็นต์จริง.   
* เขียน suggested/confidence/normalized กลับเฉพาะที่เปลี่ยน. 

นี่คือ duplicate-name harmonization engine ของระบบ. 

### recalculateAllConfidence()

re-score confidence ใหม่จาก:

* verified  
* มีพิกัด  
* มี address  
* มี province/district  
* มี UUID  
* ถ้า source เป็น Driver\_GPS ได้ bonus. 

### recalculateAllQuality()

re-score quality ใหม่ และสรุป high/mid/low quality. 

---

## **Service\_SCG.gs**

### fetchDataFromSCGJWD()

ทำหน้าที่ ETL จาก SCG API:

* lock ก่อน.   
* อ่าน cookie จาก Input\!B1.   
* อ่าน shipment list จาก Input.   
* ยิง API แบบ retry.   
* flatten shipment → note → item ให้กลายเป็น row ของ Data.   
* aggregate ระดับร้าน (ShopKey) เพื่อคำนวณ qty/weight/invoices/E-POD.   
* set headers และเขียนลงชีต Data.   
* ต่อท้ายด้วย applyMasterCoordinatesToDailyJob(), buildOwnerSummary(), buildShipmentSummary(). 

เชิงระบบ: นี่คือ daily operational pipeline ซึ่งต่อเข้ากับ master DB ทันทีหลัง fetch. 

### applyMasterCoordinatesToDailyJob()

อ่าน Data, Database, NameMapping, ข้อมูลพนักงาน:

* build map ชื่อ→coords และ UUID→coords.   
* build alias map จาก NameMapping.   
* build employee name → email map.   
* วนแต่ละ row ใน Data:  
  * match จาก alias UUID ก่อน  
  * ถัดมา match จาก normalized name  
  * ถ้าไม่เจอ ลอง branch heuristic (tryMatchBranch\_).   
  * update LatLong\_Actual \+ background color \+ employee email. 

สีพื้นหลัง:

* เขียว \= exact/UUID mapping  
* เหลือง \= branch heuristic match. 

### fetchWithRetry\_()

retry แบบ exponential backoff สำหรับ API. 

### tryMatchBranch\_()

ช่วยเดา parent name ของชื่อที่มีคำพวก “สาขา”, “branch”, “store”, “shop”. 

### checkIsEPOD(ownerName, invoiceNo)

กฎแยก invoice ที่เป็น E-POD:

* owner บางรายการถือว่า E-POD ทั้งหมด.  
* DENSO มี rule เฉพาะ invoice format. 

### buildOwnerSummary()

สรุปจำนวน invoice ปกติและ E-POD ต่อ owner ลงชีต สรุป\_เจ้าของสินค้า. 

### buildShipmentSummary()

สรุปจำนวน invoice ปกติและ E-POD ต่อ shipment+truck ลงชีต สรุป\_Shipment. 

### **clear functions**

clearDataSheet, clearInputSheet, clearSummarySheet, clearShipmentSummarySheet, clearSummarySheet\_UI, clearAllSCGSheets\_UI เป็น housekeeping utilities สำหรับ reset daily operation sheets. 

---

## **Service\_Search.gs**

### searchMasterData(keyword, page)

search engine ของ WebApp:

* parse page และ tokenize keyword.   
* โหลด alias map จาก cache/helper.   
* ใช้ getRealLastRow\_ เพื่ออ่านเฉพาะ data จริง.   
* สร้าง haystack จาก:  
  * raw name  
  * normalized name  
  * aliases  
  * normalized/AI keywords  
  * address.   
* ทุก token ต้อง match (every).   
* score 10 ถ้า aiKeywords contains rawKey, ไม่งั้น 1\.   
* sort, paginate, return object. 

ข้อดี: search รองรับทั้งชื่อจริง, alias, normalized token, และ address. 

### getCachedNameMapping\_(ss)

อ่าน NameMapping แล้วสร้าง map uid \-\> all aliases text.  
จุดแก้สำคัญ v4.1 คือวัดขนาด cache ด้วย byte size จริง แทน .length เพื่อกันปัญหาอักษรไทย. 

### clearSearchCache()

ลบ cache key เดียว. 

---

## **Service\_GeoAddr.gs**

### parseAddressFromText(fullAddress)

พยายามแยก province, district, postcode จาก text:

1. หา zip code 5 หลัก.   
2. ถ้ามี postal DB match ด้วย zip ก็คืนค่าเลย.   
3. ถ้าไม่เจอ ใช้ regex หา จังหวัด/จ. และ อ./อำเภอ/เขต.   
4. special case กรุงเทพมหานคร. 

### getPostalDataCached()

อ่าน PostalRef ครั้งเดียวแล้วเก็บไว้ใน \_POSTAL\_CACHE. 

### clearPostalCache()

reset cache ใน memory. 

### \_mapsMd5**,** \_mapsGetCache**,** \_mapsSetCache

เป็น document cache helper ของ Google Maps wrapper. 

### GOOGLEMAPS\_DURATION**,** GOOGLEMAPS\_DISTANCE**,** GOOGLEMAPS\_LATLONG**,** GOOGLEMAPS\_ADDRESS**,** GOOGLEMAPS\_REVERSEGEOCODE**,** GOOGLEMAPS\_COUNTRY

ทั้งหมดเป็น GAS wrapper ที่:

* validate input  
* รองรับ array input แบบ map  
* cache key  
* sleep 150 ms  
* call Maps service  
* cache response. 

### GET\_ADDR\_WITH\_CACHE(lat, lng)

wrapper safe สำหรับ reverse geocode. ถ้า error จะ log และคืน empty string. 

### CALCULATE\_DISTANCE\_KM(origin, destination)

เรียก maps distance แล้ว strip text ให้เป็นตัวเลข km. 

---

## **Service\_GPSFeedback.gs**

### createGPSQueueSheet()

สร้างชีต queue พร้อม header, สี, checkbox, width, frozen row. 

### applyApprovedFeedback()

flow อนุมัติ GPS driver feedback:

* lock script.   
* schema precheck.   
* โหลด queue \+ DB \+ uuid map.   
* วนเฉพาะรายการที่ Approve=true และไม่ rejected.   
* parse driver lat/lng.   
* หา DB row ตาม UUID แล้ว update lat/lng \+ coord metadata เป็น Driver\_GPS, confidence 95\.   
* เปลี่ยน queue reason เป็น APPROVED. 

ข้อสังเกต:  
ฟังก์ชันนี้มี path อนุมัติชัดเจน แต่ยังไม่มี path ที่ stamp REJECTED ให้กับรายการที่ถูกติ๊ก reject; ปัจจุบัน reject แค่ทำให้ skip ไม่เขียนกลับ. ถ้าต้องการ audit trail ฝั่ง reject จริง ๆ ยังไม่ครบ. 

### showGPSQueueStats()

สรุป queue stats เป็น total/pending/approved/rejected และแยกเหตุผล GPS\_DIFF / DB\_NO\_GPS. 

### upgradeGPSQueueSheet()

repair/upgrade header, checkbox zone, color coding ของ row ตาม reason/status. 

### resetSyncStatus()

ล้าง SYNC\_STATUS ทั้ง source sheet เพื่อ rerun sync. ใช้สำหรับทดสอบเท่านั้น. 

---

## **Service\_AutoPilot.gs**

### START\_AUTO\_PILOT()

ลบ trigger เก่าก่อน แล้วสร้าง trigger autoPilotRoutine ทุก 10 นาที. 

### STOP\_AUTO\_PILOT()

ลบ trigger handler ชื่อ autoPilotRoutine. 

### autoPilotRoutine()

background orchestrator:

* lock script.   
* ถ้ามี daily data ให้ apply master coordinates.   
* รัน processAIIndexing\_Batch(). 

### processAIIndexing\_Batch()

AI enrichment สำหรับ master DB:

* ตรวจ API key ก่อน; ถ้าไม่มีให้ skip.   
* อ่าน name และ normalized column.   
* วนทีละแถวจนถึง limit:  
  * ถ้ายังไม่มี \[AI\] tag ให้สร้าง basic smart key  
  * ถ้าชื่อยาวพอ ใช้ Gemini generate keywords/typos  
  * รวมเป็น string แล้ว append \[AI\].   
* เขียนกลับ normalized column ทีเดียว. 

### callGeminiThinking\_JSON(customerName, apiKey)

ยิง Gemini เพื่อขอ keyword/abbreviation/typo สูงสุด 5 คำ และ parse JSON array กลับ. 

### createBasicSmartKey(text)

fallback normalize ธรรมดา. 

---

## **Service\_Agent.gs**

### AGENT\_CONFIG

ใช้ config กลางหรือ fallback model/batch size. 

### WAKE\_UP\_AGENT()

manual trigger ให้คนกดใน menu แล้วรัน processAIIndexing\_Batch(). 

### SCHEDULE\_AGENT\_WORK()

ลบ trigger เก่าที่ชื่อ runAgentLoop แล้วสร้าง autoPilotRoutine every 10 minutes. แสดงว่านี่คือ legacy compatibility helper มากกว่าจะเป็น schedule หลัก. 

### resolveUnknownNamesWithAI()

flow AI tier-4:

1. lock script.   
2. อ่าน Data และหา ShipToName ที่ยังไม่มี LatLong\_Actual.   
3. อ่าน Database เพื่อสร้าง candidate master options.   
4. จำกัด subset 500 รายการ.   
5. prompt Gemini ให้ map unknown names → matched UID เฉพาะที่ confidence \>= 60\.   
6. parse result แล้ว append ลง NameMapping.   
7. clear search cache \+ apply coordinates ใหม่ใน daily data. 

ข้อจำกัดเชิงคุณภาพ: master subset ถูก slice(0, 500\) ตรง ๆ ไม่ได้ rank ตาม relevance จึงถ้า DB โตมาก AI จะเห็นเพียง 500 ชื่อแรก อาจพลาด candidate สำคัญ. 

### askGeminiToPredictTypos(originalName)

ฟังก์ชัน utility สำหรับ generate typo keywords. ใน codebase นี้ยังดูเป็น helper เสริมมากกว่าหัว flow หลัก. 

---

## **Service\_SchemaValidator.gs**

### SHEET\_SCHEMA

ประกาศ expected schema ของ DATABASE, NAMEMAPPING, SCG\_SOURCE, GPS\_QUEUE, DATA. 

### validateSheet\_(schemaKey)

เช็ก:

* sheet มีจริงไหม  
* คอลัมน์ถึงขั้นต่ำไหม  
* header สำคัญตรงตาม expected หรือไม่. 

### validateSchemas(schemaKeys)

รวมผลหลาย schema เป็น object เดียว. 

### preCheck\_Sync**,** preCheck\_Apply**,** preCheck\_Approve

เป็น guard function สำหรับ flow สำคัญ. 

### throwSchemaError\_(results, flowName)

รวม error ทั้งหมดเป็นข้อความเดียวแล้ว throw. UX ดีสำหรับ admin. 

### runFullSchemaValidation()

แสดง report ทุก schema. 

### fixNameMappingHeaders()

เขียน header มาตรฐานของ NameMapping. 

---

## **Service\_SoftDelete.gs**

### initializeRecordStatus()

เติม Record\_Status \= Active ให้ทุก row ที่ยังไม่มี. 

### softDeleteRecord(uuid)

ไม่ลบ row จริง แต่ mark Inactive \+ update timestamp. 

### mergeUUIDs(masterUUID, duplicateUUID)

mark duplicate row เป็น Merged, set MERGED\_TO\_UUID, update timestamp. 

### resolveUUID(uuid)

ตาม chain ของ merged UUID สูงสุด 10 hop เพื่อหา UUID ปลายทางจริง. นี่สำคัญมากสำหรับระบบที่มี merge ซ้อนหลายครั้ง. 

### mergeDuplicates\_UI()

UI prompt 2 จังหวะสำหรับ merge UUID. clear cache หลัง merge. 

### showRecordStatusReport()

สรุป active/inactive/merged/noStatus. 

---

## **Service\_Notify.gs**

### sendSystemNotify(message, isUrgent)

fan-out ไปทั้ง LINE และ Telegram แบบ best effort. 

### sendLineNotify **/** sendTelegramNotify

public wrappers. 

### sendLineNotify\_Internal\_()

ส่งผ่าน LINE Notify API ด้วย bearer token. 

### sendTelegramNotify\_Internal\_()

รองรับทั้ง property ชุด TG\_\* และ legacy TELEGRAM\_\*, ใช้ HTML parse mode และ escape HTML ก่อนส่ง. 

### notifyAutoPilotStatus(...)

format สถานะ autopilot รอบล่าสุดแล้วกระจายแจ้งเตือน. 

---

## **Service\_Maintenance.gs**

### cleanupOldBackups()

ค้น sheet ที่ขึ้นต้น Backup\_, parse วันที่จากชื่อ, ถ้าเก่ากว่า 30 วันให้ลบ, แล้วแจ้งเตือนผ่าน LINE/Telegram. 

### checkSpreadsheetHealth()

รวมจำนวนเซลล์จากทุกชีต เทียบกับ limit 10M cells:

* ถ้าเกิน 80% \= แจ้งเตือนรุนแรง  
* ไม่งั้น toast ว่า OK. 

---

## **Setup\_Security.gs**

### setupEnvironment()

prompt ใส่ Gemini API key และ validate ว่าขึ้นต้น AIza. 

### setupLineToken()

prompt และบันทึก LINE token. 

### setupTelegramConfig()

prompt bot token \+ chat ID แล้วเก็บใน properties. 

### resetEnvironment()

ลบ Gemini API key. 

---

## **Setup\_Upgrade.gs**

### upgradeDatabaseStructure()

เพิ่ม headers ของ GPS tracking columns ถ้าขาด. 

### upgradeNameMappingStructure\_V4()

format NameMapping ให้เป็น schema 5 คอลัมน์มาตรฐาน. 

### findHiddenDuplicates()

ใช้ grid bucket \+ Haversine เพื่อตรวจพิกัดต่างชื่อที่อยู่ใกล้ ≤ 50 m. เหมาะเป็น QA tool. 

### verifyDatabaseStructure()

log ตรวจคอลัมน์ 18-20. เป็น developer utility. 

---

## **Menu.gs**

### onOpen()

สร้าง 4 menu ใหญ่:

1. Master Data  
2. SCG Operation  
3. Automation  
4. System Admin. 

นี่บอกชัดว่า UX หลักของระบบคือ “ใช้จาก Google Sheets menu” ไม่ใช่ command/API-first. 

### **safety wrappers**

* syncNewDataToMaster\_UI() \= confirm ก่อน sync GPS feedback.   
* runAIBatchResolver\_UI() \= confirm ก่อน AI smart resolution.   
* finalizeAndClean\_UI() \= confirm ก่อน finalize.   
* resetDeepCleanMemory\_UI() \= confirm ก่อน reset pointer.   
* clearDataSheet\_UI(), clearInputSheet\_UI(), repairNameMapping\_UI() \= reuse confirmAction.   
* runSystemHealthCheck() \= เรียก CONFIG.validateSystemIntegrity().   
* clearPostalCache\_UI(), clearSearchCache\_UI() \= UI wrappers ของ cache management.   
* showQualityReport\_UI() \= คำนวณสถิติคุณภาพระดับ dashboard. 

---

## **WebApp.gs**

### doGet(e)

* อ่าน page parameter, default Index.  
* inject initialQuery.  
* set title, favicon, viewport.  
* allow iframe. 

### doPost(e)

รองรับ webhook/action ง่าย ๆ:

* parse payload JSON  
* ถ้า action \= triggerAIBatch ให้รัน AI indexing  
* ไม่งั้น echo webhook response. 

### createJsonResponse\_(obj)

JSON output helper. 

### include(filename)

อ่าน partial HTML file ถ้ามี. 

### getUserContext()

ดึง email/locale จาก session. 

---

## **Index.html**

### **โครงสร้าง HTML/CSS**

หน้า UI ถูกออกแบบ mobile-first:

* header gradient  
* search box  
* result cards  
* pagination bar  
* loading overlay  
* toast. 

### **JavaScript functions**

* window.onload โหลด initial query จาก server template และ bind event.   
* toggleClearBtn() แสดง/ซ่อนปุ่ม clear.   
* clearSearch() เคลียร์ input.   
* triggerSearch() validate keyword แล้วเริ่ม search.   
* fetchData(page) เรียก GAS server function searchMasterData.   
* renderResults(response) สร้าง DOM card:  
  * name  
  * AI badge ถ้า score สูง  
  * UUID  
  * address  
  * copy coords  
  * ปุ่ม Google Maps / Waze.   
* renderPagination(total, current) ทำ pagination dots แบบย่อ.   
* changePage, showLoading, handleError \= helper UI.   
* copyCoord, fallbackCopy \= copy to clipboard.   
* showToast, escapeHtml \= UX/security helpers. 

จุดดี: client side มี escapeHtml ก่อน inject content จึงป้องกัน XSS ในผลค้นหาได้พอสมควร. 

---

## **Test\_AI.gs**

* forceRunAI\_Now() \= manual run AI indexing.   
* debug\_TestTier4SmartResolution() \= ยิง AI tier-4 resolver แบบ confirm.   
* debugGeminiConnection() \= ping Gemini ด้วยข้อความทดสอบ.   
* debug\_ResetSelectedRowsAI() \= ลบ \[AI\] tags ใน row ที่เลือก. 

ชุดนี้คือ operational debug tools มากกว่าจะเป็น unit tests. 

---

## **Test\_Diagnostic.gs**

* RUN\_SYSTEM\_DIAGNOSTIC() \= เช็ก symbol/function visibility, Maps API, GEMINI\_API\_KEY, LINE/TG config.   
* RUN\_SHEET\_DIAGNOSTIC() \= เช็กว่าชีตสำคัญมีหรือไม่ และบางส่วนเช็ก schema แบบหยาบ. 

ชุดนี้คือ smoke-test layer ที่ใช้จากเมนู admin ได้ทันที. 

---

# **4\) วิเคราะห์ flow หลักแบบ end-to-end**

## **Flow A: รับ GPS จริงจาก SCG source เข้าฐานกลาง**

1. เมนู syncNewDataToMaster\_UI() ให้ user confirm.   
2. syncNewDataToMaster() อ่าน source, DB, NameMapping, GPS\_Queue.   
3. ชื่อใหม่ → insert DB.   
4. ชื่อเดิมแต่พิกัดต่าง \>50m → queue รออนุมัติ.   
5. admin อนุมัติผ่าน applyApprovedFeedback(). 

นี่คือ feedback loop สำคัญที่สุดของ v4.1. 

## **Flow B: ดึง shipment ประจำวัน**

1. ใส่ cookie \+ shipment list ใน Input.   
2. รัน fetchDataFromSCGJWD().   
3. ระบบ flatten ข้อมูลง Data.   
4. ระบบ apply master coordinates \+ employee emails.   
5. ระบบสร้าง owner/shipment summaries. 

## **Flow C: Clean \+ cluster \+ finalize**

1. runDeepCleanBatch\_100() เติม geocode/uuid/admin fields.   
2. processClustering\_GridOptimized() สร้าง suggested names/confidence.   
3. human tick verified ในชีต.  
4. finalizeAndClean\_MoveToMapping() ย้าย alias เข้า NameMapping และเก็บ verified rows เป็นฐาน active. 

## **Flow D: Web search**

1. ผู้ใช้เปิด webapp ผ่าน doGet().   
2. ใส่ keyword ใน Index.html.   
3. server searchMasterData() ค้นจาก DB \+ alias \+ address \+ AI tokens.   
4. client render result cards \+ navigation links. 

---

# **5\) จุดแข็งเชิงวิศวกรรม**

## **จุดแข็งหลัก**

* ใช้ CONFIG/C\_IDX เป็นศูนย์กลาง ลด hardcode.   
* ใช้ getRealLastRow\_() แก้ปัญหา checkbox/ghost rows ได้ดี.   
* มี lock ใน flow สำคัญหลายตัว.   
* มี schema validator ก่อน run flow สำคัญ.   
* ใช้ soft delete/merge แทน hard delete.   
* แยก UI, business logic, utility, setup, diagnostics ค่อนขางชัด. 

---

# **6\) จุดเสี่ยง / code smell / ข้อสังเกตเชิงลึก**

## **6.1** repairNameMapping\_Full() **มี logic fallback UUID ที่น่าสงสัย**

บรรทัด 496 lookup uuidMap\[normalizeText(r\[1\])\] แต่ r\[1\] เป็น UID ไม่ใช่ชื่อ. โอกาสสูงว่าบรรทัดนี้ไม่ได้ช่วยอะไรจริงตามเจตนา. 

## **6.2 queue reject path ยังไม่สมบูรณ์**

applyApprovedFeedback() สนใจเฉพาะ approve; ถ้าติ๊ก reject จะ skip แต่ไม่ได้ stamp REJECTED ลงคอลัมน์ reason. ทำให้ audit trail ฝั่ง reject ไม่ชัด. 

## **6.3 AI tier-4 จำกัด candidate DB แบบตัดตรง 500 รายการแรก**

ถ้า DB ใหญ่มาก accuracy จะขึ้นกับลำดับข้อมูลใน sheet มากเกินไป. ควร prefilter candidates ตาม token ก่อนค่อยส่ง AI. 

## **6.4 search score ยังหยาบ**

score มีแค่ 10 หรือ 1 ตามการ match AI keyword; ยังไม่มี ranking จาก exact name match, alias match, address proximity, verified status ฯลฯ. 

## **6.5 search ยังอ่านแค่ 17 คอลัมน์**

จึงไม่สามารถใช้ข้อมูลใหม่ของ v4.1 เช่น Coord\_Source, Coord\_Confidence, Record\_Status มา influence result ได้. 

## **6.6 finalize ยังทำงานบน 17 คอลัมน์หลัก**

logic finalizeAndClean\_MoveToMapping() อ่าน/เขียน 17 คอลัมน์ ไม่รวม GPS tracking และ soft delete columns โดยตรง จึงหมายความว่า flow นี้ยัง anchored กับ schema pre-v4.1 บางส่วน แม้ระบบรวมโดยรวมจะมีคอลัมน์ถึง 22\. 

---

# **7\) สรุปความสัมพันธ์ “ทุกโมดูล”**

## **Master governance**

* Config.gs  
* Utils\_Common.gs  
* Service\_Master.gs  
* Service\_SoftDelete.gs  
* Service\_SchemaValidator.gs. 

## **SCG daily operations**

* Service\_SCG.gs  
* Service\_GPSFeedback.gs  
* Service\_GeoAddr.gs. 

## **AI & automation**

* Service\_AutoPilot.gs  
* Service\_Agent.gs  
* Test\_AI.gs. 

## **UI / Access layer**

* Menu.gs  
* WebApp.gs  
* Index.html  
* Service\_Search.gs. 

## **Ops & admin**

* Setup\_Security.gs  
* Setup\_Upgrade.gs  
* Service\_Maintenance.gs  
* Service\_Notify.gs  
* Test\_Diagnostic.gs. 

---

# **8\) สรุปแบบตรงไปตรงมา**

ถ้าสรุปสั้นที่สุด:

* repo นี้เป็น ระบบ MDM \+ daily logistics ops \+ AI-assisted name resolution \+ mobile search UI ที่ส้างบน Google Sheets/Apps Script.   
* ไฟล์ที่ “สำคัญที่สุด” เชิง business logic คือ Service\_Master.gs, Service\_SCG.gs, Service\_Search.gs, Service\_GPSFeedback.gs.   
* V4.1 ชัดเจนว่าถูกออกแบบเพื่อเพิ่ม ความปลอดภัยของข้อมูล, human approval loop, และ diagnostic visibility.   
* โค้ดโดยรวม practical มากและใช้ได้จริงในบริบท Google Sheets operation แต่ยังมีจุดที่ควร refactor เรื่อง ranking, reject handling, candidate selection ของ AI, และ consistency ระหว่าง schema 17 vs 22 columns. 

# ข้อสรุปเชิงอนุมาน

ด้านล่างนี้คือ Dependency Map \+ Sequence Diagram \+ Data Dictionary ทุกชีต/ทุกคอลัมน์ที่ source code บอกได้ โดยผมจะระบุชัด ๆ ว่าอะไรคือ ข้อมูลที่ยืนยันได้จาก code และอะไรคือ ข้อสรุปเชิงอนุมานจากตำแหน่งคอลัมน์/ลำดับการเขียนข้อมูล. 

---

# **1\) Dependency Map**

## **1.1 Dependency Map ระดับไฟล์/โมดูล**

### **คำอธิบาย dependency สำคัญ**

* Config.gs เป็นศูนย์กลางของทุกโมดูล เพราะทุก service อ้างชื่อชีต, index คอลัมน์, threshold, AI model และ config ผ่าน CONFIG/SCG\_CONFIG.   
* Utils\_Common.gs เป็นฐาน utility สำหรับ name normalization, UUID, Haversine, retry และถูกใช้ข้ามโมดูลเกือบทั้งระบบ.   
* Service\_Master.gs พึ่งพา Utils\_Common.gs, Service\_GeoAddr.gs, และ Service\_SchemaValidator.gs โดยตรงในการ sync / enrich / cluster / finalize.   
* Service\_SCG.gs พึ่ง Config, Utils, และผลลัพธ์จาก Master/NameMapping เพื่อ map พิกัดลง daily job sheet.   
* Service\_Search.gs พึ่ง NameMapping cache \+ Database \+ normalizeText() เพื่อทำ search.   
* Service\_GPSFeedback.gs พึ่ง Service\_Master.gs ทางอ้อมผ่าน UUID/DB structure และใช้ precheck จาก Service\_SchemaValidator.gs.   
* Service\_AutoPilot.gs เรียก applyMasterCoordinatesToDailyJob() และ processAIIndexing\_Batch() จึงทำตัวเป็น scheduler/orchestrator.   
* Service\_Agent.gs พึ่งทั้ง Data, Database, NameMapping, และหลัง AI match แล้วย้อนกลับไปเรียก applyMasterCoordinatesToDailyJob().   
* WebApp.gs เสิร์ฟ Index.html, ส่วน Index.html เรียก server function searchMasterData() จาก Service\_Search.gs. 

---

## **1.2 Dependency Map ระดับชีตข้อมูล**

### **สรุปความสัมพันธ์เชิงข้อมูล**

* SCGนครหลวงJWDภูมิภาค คือ source ของ GPS จริงที่ feed เข้าทั้ง Database และ GPS\_Queue.   
* Database คือ canonical source ของพิกัดและ customer identity.   
* NameMapping คือ layer ของ alias/variant → UUID.   
* Data คือ operational working sheet ที่ได้จาก SCG API แล้วถูก enrich ด้วย master coordinates และ employee email.   
* PostalRef ใช้ enrich จังหวัด/อำเภอ/รหัสไปรษณีย์ให้ Database.   
* GPS\_Queue คือ human approval buffer ก่อนอัปเดตพิกัดจริงกลับเข้า Database.   
* สรุป\_เจ้าของสินค้า และ สรุป\_Shipment ได้ข้อมูลจาก Data เท่านั้น. 

---

# **2\) Sequence Diagram**

---

## **2.1 Sequence: Sync GPS Feedback → Database / Queue**

อ้างอิง flow นี้มาจาก syncNewDataToMaster\_UI() ในเมนู, preCheck\_Sync(), และ logic หลักของ syncNewDataToMaster(). 

---

## **2.2 Sequence: Approve GPS Queue → Update Master**

อ้างอิงจาก applyApprovedFeedback() และ clearSearchCache(). 

---

## **2.3 Sequence: Fetch Shipment → Enrich Daily Jobs → Build Summary**

อ้างอิงจาก fetchDataFromSCGJWD(), applyMasterCoordinatesToDailyJob(), buildOwnerSummary(), และ buildShipmentSummary(). 

---

## **2.4 Sequence: AI Tier-4 Resolve Unknown Names**

อ้างอิงจาก runAIBatchResolver\_UI() และ resolveUnknownNamesWithAI(). 

---

## **2.5 Sequence: Web Search**

อ้างอิงจาก doGet(), searchMasterData(), getCachedNameMapping\_(), และ client JS ใน Index.html. 

---

# **3\) Data Dictionary ทุกชีต / ทุกคอลัมน์**

หมายเหตุสำคัญ:  
บางชีตมี header ชัดจาก code, บางชีต code อ้างเพียงบางคอลัมน์.  
ผมจะแบ่งเป็น 3 ระดับ:

* \[ยืนยัน\] \= code ระบุชัดเจน  
* \[อนุมานสูง\] \= ชื่อตาม semantic จาก CONFIG/ลำดับ array  
* \[ไม่ระบุใน code\] \= source มีคอลัมน์อื่นแต่ repo นี้ไม่ใช้

---

## **3.1 ชีต** Database

README ระบุว่าชีตนี้เป็น Golden Record ลูกค้า และใช้คอลัมน์ A-V รวม 22 คอลัมน์. Config.gs นิยามความหมายครบทุกตำแหน่ง. 

### **Data Dictionary:** Database

| Col | Field | สถานะ | ความหมาย | ใช้โดย |
| ----- | ----- | ----- | ----- | ----- |
| A | NAME | ยืนยัน | ชื่อลูกค้า/ปลายทางหลัก | sync, search, cluster, finalize, quality.  |
| B | LAT | ยืนยัน | latitude หลักของลูกค้า | sync, apply, search, quality.  |
| C | LNG | ยืนยัน | longitude หลัก | sync, apply, search, quality.  |
| D | SUGGESTED | อนุมานสูง | suggested master name จาก clustering/human curation | clustering, finalize.  |
| E | CONFIDENCE | อนุมานสูง | คะแนนความมั่นใจ 0-100 ของชื่อ/record | clustering, recalc.  |
| F | NORMALIZED | อนุมานสูง | normalized / AI keywords / search index text | search, autopilot AI.  |
| G | VERIFIED | อนุมานสูง | checkbox/flag ว่าระเบียนนี้ human verified แล้ว | finalize, clustering, quality.  |
| H | SYS\_ADDR | อนุมานสูง | address จากระบบต้นทาง/SCG | sync, search fallback.  |
| I | ADDR\_GOOG / GOOGLE\_ADDR | อนุมานสูง | address จาก reverse geocode | deep clean, search, quality.  |
| J | DIST\_KM | อนุมานสูง | ระยะทางจาก depot เป็นกิโลเมตร | deep clean.  |
| K | UUID | ยืนยัน | รหัสประจำลูกค้า canonical identity | mapping, search, merge, queue update.  |
| L | PROVINCE | อนุมานสูง | จังหวัด | deep clean, quality.  |
| M | DISTRICT | อนุมานสูง | อำเภอ/เขต | deep clean, quality.  |
| N | POSTCODE | อนุมานสูง | รหัสไปรษณีย์ | deep clean, quality.  |
| O | QUALITY | ยืนยัน | quality score 0-100 ของ data completeness | report, recalc.  |
| P | CREATED | ยืนยัน | วันเวลาสร้าง record | sync, deep clean.  |
| Q | UPDATED | ยืนยัน | วันเวลาอัปเดตล่าสุด | sync, deep clean, feedback, merge.  |
| R | Coord\_Source | ยืนยัน | แหล่งที่มาของพิกัด เช่น SCG\_System, Driver\_GPS | sync, feedback, confidence.  |
| S | Coord\_Confidence | ยืนยัน | ความเชื่อมั่นของพิกัด 0-100 | sync, feedback.  |
| T | Coord\_Last\_Updated | ยืนยัน | timestamp พิกัดล่าสุด | sync, feedback.  |
| U | Record\_Status | อนุมานสูง | Active / Inactive / Merged | soft delete/report/deep clean.  |
| V | Merged\_To\_UUID | อนุมานสูง | UUID ปลายทางหลัง merge | merge/resolve UUID.  |

### **ข้อสังเกตเชิงใช้งาน**

* Schema validator ปัจจุบันเช็กขั้นต่ำถึงคอลัมน์ 20 และ header สำคัญบางคอลัมน์เท่านั้น ยังไม่ได้ assert U/V ใน validator.   
* หลาย flow ยังอ่าน/เขียน 17 คอลัมน์แรกอยู่ จึงทำให้คอลัมน์ R-V เป็น “extension layer” ของ v4.1 มากกว่าจะถูกรวมเต็มในทุกฟังก์ชัน. 

---

## **3.2 ชีต** NameMapping

Config.gs ระบุ index ครบ, และ schema validator ระบุชื่อ header ชัดเจน. 

### **Data Dictionary:** NameMapping

| Col | Header | สถานะ | ความหมาย | ใช้โดย |
| ----- | ----- | ----- | ----- | ----- |
| A | Variant\_Name | ยืนยัน | ชื่อแปรผัน/ชื่อเดิม/alias | search, alias resolution, finalize.  |
| B | Master\_UID | ยืนยัน | UUID ของ master record ใน Database | search, sync, apply coordinates.  |
| C | Confidence\_Score | ยืนยัน | ความมั่นใจของ mapping | AI match, finalize, repair.  |
| D | Mapped\_By | ยืนยัน | แหล่งที่มาของ mapping เช่น Human, AI\_Agent\_V4, System\_Repair | audit.  |
| E | Timestamp | ยืนยัน | วันเวลาที่สร้าง/แก้ mapping | audit.  |

---

## **3.3 ชีต** SCGนครหลวงJWDภูมิภาค

repo นี้ไม่ได้ให้ header ครบทุกคอลัมน์ของชีตนี้ แต่ SCG\_CONFIG.SRC\_IDX และ schema validator บอกคอลัมน์สำคัญที่ระบบใช้จริง. 

### **Data Dictionary:** SCGนครหลวงJWDภูมิภาค

| Col | Header/Field | สถานะ | ความหมาย | ใช้โดย |
| ----- | ----- | ----- | ----- | ----- |
| M (13) | ชื่อปลายทาง | ยืนยัน | ชื่อปลายทางจาก source | sync ชื่อลูกค้าใหม่/หา match DB.  |
| O (15) | LAT | ยืนยัน | GPS จริง latitude จากคนขับ | sync diff / queue.  |
| P (16) | LONG | ยืนยัน | GPS จริง longitude จากคนขับ | sync diff / queue.  |
| S (19) | ที่อยู่ปลายทาง | ยืนยัน | address ปลายทางจาก source system | ใช้เติม SYS\_ADDR ใน DB แถวใหม่.  |
| X (24) | ระยะทางจากคลัง\_Km | ยืนยันแต่ยังไม่ใช้ตรงใน flow หลักนี้ | ระยะจากคลัง | ระบุใน schema/config แต่ไม่ได้ถูกอ่านใน syncNewDataToMaster().  |
| Y (25) | ชื่อที่อยู่จาก\_LatLong | ยืนยันแต่ยังไม่ใช้ตรงใน flow หลักนี้ | address จาก lat/lng | ระบุใน schema/config.  |
| AK (37) | SYNC\_STATUS | ยืนยัน | checkpoint ว่าประมวลผลแล้วหรือยัง | SYNCED กันซ้ำ.  |

### **คอลัมน์อื่น**

* ไม่ระบุใน code: คอลัมน์ A-L, N, Q-R, T-W, Z-AJ และอื่น ๆ repo นี้ไม่ได้ใช้โดยตรง จึงไม่สามารถทำ dictionary แบบแม่นยำได้จาก source เพียงอย่างเดียว. 

---

## **3.4 ชีต** Input

SCG\_CONFIG ระบุ cell สำคัญและ code ใช้จริงชัดเจน. 

### **Data Dictionary:** Input

| ตำแหน | Field | สถานะ | ความหมาย | ใช้โดย |
| ----- | ----- | ----- | ----- | ----- |
| B1 | Cookie | ยืนยัน | cookie สำหรับยิง SCG API | fetchDataFromSCGJWD().  |
| B3 | Shipment String | ยืนยัน | comma-separated shipment string ที่ระบบ compose ให้ | บันทึกเพื่อ reference/debug.  |
| A4:A... | Shipment Numbers | ยืนยัน | รายการ shipment number ที่จะยิง API | fetchDataFromSCGJWD().  |

### **คอลัมน์/เซลล์อื่น**

* ไม่ระบุใน code ว่ามีบทบาทอะไร. repo นี้สนใจเพียง B1, B3 และ A ตั้งแต่ row 4 ลงไป. 

---

## **3.5 ชีต** Data

header ถูกสร้างแบบ explicit ใน fetchDataFromSCGJWD(). อันนี้ทำ data dictionary ได้ครบที่สุด. 

### **Data Dictionary:** Data

| Col | Header | สถานะ | ความหมาย | ใช้โดย |
| ----- | ----- | ----- | ----- | ----- |
| A | ID\_งานประจำวัน | ยืนยัน | daily job id \= PurchaseOrder \+ running row | operational identity.  |
| B | PlanDelivery | ยืนยัน | วันที่ plan delivery | แสดง/ใช้งาน operational.  |
| C | InvoiceNo | ยืนยัน | invoice / purchase order string | summary / ePOD logic.  |
| D | ShipmentNo | ยืนยัน | shipment number | shipment summary.  |
| E | DriverName | ยืนยัน | ชื่อคนขับ | ใช้ map email.  |
| F | TruckLicense | ยืนยัน | ทะเบียนรถ | shipment summary key.  |
| G | CarrierCode | ยืนยัน | รหัส carrier | operational display.  |
| H | CarrierName | ยืนยัน | ชื่อ carrier | operational display.  |
| I | SoldToCode | ยืนยัน | owner/customer code | operational.  |
| J | SoldToName | ยืนยัน | ชื่อเจ้าของสินค้า | owner summary, ePOD logic.  |
| K | ShipToName | ยืนยัน | ชื่อปลายทาง | coordinate matching, AI tier-4.  |
| L | ShipToAddress | ยืนยัน | ที่อยู่ปลายทางจาก API | operational display.  |
| M | LatLong\_SCG | ยืนยัน | พิกัดจาก SCG API | reference เทียบกับ actual.  |
| N | MaterialName | ยืนยัน | ชื่อสินค้า | operational detail.  |
| O | ItemQuantity | ยืนยัน | จำนวนต่อ item | ใช้ aggregate ระดับร้าน.  |
| P | QuantityUnit | ยืนยัน | หน่วยสินค้า | operational detail.  |
| Q | ItemWeight | ยืนยัน | น้ำหนักต่อ item | ใช้ aggregate ระดับร้าน. {line\_range\_start=77 line\_range\_end=77 path=Service\_SCG.gsL91-L92】 |
| R | DeliveryNo | ยืนยัน | delivery no | operational/detail. 【F:Service\_SCG.gs git\_url="https://github.com/kamonwantanakun-svg/claude\_lmds\_v4.1/blob/main/Service\_SCG.gsL91-L92】 |
| R | DeliveryNo | ยืนยัน | delivery no | operational/detail. 【F:Service\_SCG.gs\#L77-L77"} |
| S | จำนวนปลายทาง\_System | ยืนยัน | จำนวน destination unique ใน shipment นั้น | generated metric.  |
| T | รายชื่อปลายทาง\_System | ยืนยัน | รายชื่อปลายทางรวมของ shipment | generated metric.  |
| U | ScanStatus | ยืนยัน | สถานะสแกน | default รอสแกน.  |
| V | DeliveryStatus | ยืนยัน | สถานะส่งของ | default ยังไม่ได้ส่ง.  |
| W | Email พนักงาน | ยืนยัน | email คนขับ/พนักงาน | ถูกเติมจาก ข้อมูลพนักงาน.  |
| X | จำนวนสินค้ารวมของร้านนี้ | ยืนยัน | sum quantity per shop key | generated aggregate.  |
| Y | น้ำหนักสินค้ารวมของร้านนี้ | ยืนยัน | sum weight per shop key | generated aggregate.  |
| Z | จำนวน\_Invoice\_ที่ต้องสแกน | ยืนยัน | จำนวน invoice ที่ต้อง scan หลังหัก E-POD | generated aggregate.  |
| AA | LatLong\_Actual | ยืนยัน | พิกัดจริงที่ match จาก master DB | search by ops / AI unknown detection.  |
| AB | ชื่อเจ้าของสินค้า\_Invoice\_ที่ต้องสแกน | ยืนยัน | label owner \+ scan invoice count | generated text.  |
| AC | ShopKey | ยืนยัน | composite key \`ShipmentNo | ShipToName\` |

---

## **3.6 ชีต** GPS\_Queue

header ถูกสร้าง explicit ใน createGPSQueueSheet() และ validator ก็ยืนยันตรงกัน. 

### **Data Dictionary:** GPS\_Queue

| Col | Header | สถานะ | ความหมาย | ใช้โดย |
| ----- | ----- | ----- | ----- | ----- |
| A | Timestamp | ยืนยัน | เวลาที่ queue item ถูกสร้าง | queue audit.  |
| B | ShipToName | ยืนยัน | ชื่อปลายทางที่มี GPS feedback | admin review.  |
| C | UUID\_DB | ยืนยัน | UUID ของ record ใน DB ที่เกี่ยวข้อง | apply approved feedback.  |
| D | LatLng\_Driver | ยืนยัน | พิกัดจากคนขับ | apply approved feedback.  |
| E | LatLng\_DB | ยืนยัน | พิกัดเดิมใน DB | compare/admin review.  |
| F | Diff\_Meters | ยืนยัน | ความต่างระยะเมตร | queue stats/review.  |
| G | Reason | ยืนยัน | GPS\_DIFF, DB\_NO\_GPS, APPROVED, และเชิงออกแบบน่าจะรองรับ REJECTED | workflow state.  |
| H | Approve | ยืนยัน | checkbox อนุมัติ | admin action.  |
| I | Reject | ยืนยัน | checkbox ปฏิเสธ | admin action แต่ยังไม่ stamp state กลับชัดเจน.  |

---

## **3.7 ชีต** PostalRef

README ระบุว่ามี postcode, district, province. โค้ดใช้เพียง column 1, 3, 4 โดย 0-based index. 

### **Data Dictionary:** PostalRef

| Col | Field | สถานะ | ความหมาย | ใช้โดย |
| ----- | ----- | ----- | ----- | ----- |
| A | postcode / ZIP | ยืนยัน | รหัสไปรษณีย์ | parseAddressFromText().  |
| B | ไม่ใช้ใน code | ไม่ระบุ | repo นี้ไม่ได้อ่าน | \- |
| C | district | ยืนยัน | อำเภอ/เขต | parseAddressFromText().  |
| D | province | ยืนยัน | จังหวัด | parseAddressFromText().  |
| E+ | ไม่ใช้ใน code | ไม่ระบุ | ถ้ามี ไม่ถูกใช้งาน | \- |

---

## **3.8 ชีต** ข้อมูลพนักงาน

README และ applyMasterCoordinatesToDailyJob() ระบุชัดว่าต้องใช้:

* คอลัมน์ B \= ชื่อ  
* คอลัมน์ G \= Email. 

### **Data Dictionary:** ข้อมูลพนักงาน

| Col | Field | สถานะ | ความหมาย | ใช้โดย |
| ----- | ----- | ----- | ----- | ----- |
| B | ชื่อพนักงาน/คนขับ | ยืนยัน | ใช้ normalize แล้ว match กับ DriverName ใน Data | applyMasterCoordinatesToDailyJob().  |
| G | Email | ยืนยัน | email ของคนขับ/พนักงาน | เติมลง Data\!W.  |
| คอลัมน์อื่น | \- | ไม่ระบุ | repo นี้ไม่ใช้โดยตรง | \- |

---

## **3.9 ชีต** สรุป\_เจ้าของสินค้า

repo นี้ ไม่สร้าง header ให้ชีตนี้ แต่ buildOwnerSummary() เขียนข้อมูล 6 คอลัมน์ตามลำดับแน่นอน. เพราะฉะนั้น dictionary นี้เป็น \[อนุมานสูงจาก write order\]. 

### **Data Dictionary:** สรุป\_เจ้าของสินค้า

| Col | Field | สถานะ | ความหมาย |
| ----- | ----- | ----- | ----- |
| A | Placeholder / Manual flag | อนุมานสูง | code ใส่ค่าว่างเสมอ อาจเผื่อ user ใช้เอง |
| B | Owner Name | อนุมานสูง | ชื่อเจ้าของสินค้า (SoldToName) |
| C | Placeholder / Manual note | อนุมานสูง | code ใส่ค่าว่างเสมอ |
| D | Non-EPOD Invoice Count | อนุมานสูง | จำนวน invoice ปกติ |
| E | EPOD Invoice Count | อนุมานสูง | จำนวน invoice แบบ E-POD |
| F | Generated At | อนุมานสูง | เวลาที่ระบบสร้าง summary |

อ้างอิงจาก rows.push(\["", owner, "", o.all.size, o.epod.size, new Date()\]). 

---

## **3.10 ชีต** สรุป\_Shipment

เช่นเดียวกัน ไม่มี header ถูกสร้างโดย code แต่ buildShipmentSummary() เขียน 7 คอลัมน์ตามลำดับแน่นอน. จึงเป็น \[อนุมานสูงจาก write order\]. 

### **Data Dictionary:** สรุป\_Shipment

| Col | Field | สถานะ | ความหมาย |
| ----- | ----- | ----- | ----- |
| A | ShipmentTruckKey | อนุมานสูง | key \= ShipmentNo\_TruckLicense |
| B | ShipmentNo | อนุมานสูง | shipment number |
| C | TruckLicense | อนุมานสูง | ทะเบียนรถ |
| D | Placeholder / Manual note | อนุมานสูง | code ใส่ค่าว่าง |
| E | Non-EPOD Invoice Count | อนุมานสูง | จำนวน invoice ปกติ |
| F | EPOD Invoice Count | อนุมานสูง | จำนวน invoice แบบ E-POD |
| G | Generated At | อนุมานสูง | เวลาที่สร้าง summary |

อ้างอิงจาก rows.push(\[key, s.shipmentNo, s.truck, "", s.all.size, s.epod.size, new Date()\]). 

---

# **4\) Cross-Sheet Field Mapping**

## **4.1 การ map identity**

| From | To | Rule |
| ----- | ----- | ----- |
| SCGนครหลวงJWDภูมิภาค\!M | Database\!A | normalize name แล้ว match ถ้าไม่มีจึงสร้างใหม่.  |
| NameMapping\!A | NameMapping\!B | alias → UUID.  |
| NameMapping\!B | Database\!K | UUID link ระหว่าง alias กับ master.  |
| Data\!K (ShipToName) | NameMapping\!A / Database\!A | apply master coordinates และ AI resolution.  |
| GPS\_Queue\!C | Database\!K | อัปเดตพิกัดจาก queue กลับสู่ master.  |

---

## **4.2 การ map พิกัด**

| From | To | Rule |
| ----- | ----- | ----- |
| SCG source O/P | Database B/C | ถ้าเป็น record ใหม่ หรืออนุมัติ feedback ภายหลัง.  |
| Database B/C | Data AA | เมื่อชื่อ match หรือ alias match.  |
| Database R/S/T | ภายใน DB | metadata ของแหล่งพิกัดและความเชื่อมั่น.  |

---

## **4.3 การ map คุณภาพ/สถานะ**

| Field | คำนวณจาก |
| ----- | ----- |
| Database\!O QUALITY | name, coords, geo area validity, address, province/district, postcode, UUID, verified.  |
| Database\!E CONFIDENCE | verified, coords, address, geo fields, UUID, coord source.  |
| Database\!U Record\_Status | initialize / soft delete / merge flow.  |

---

# **5\) ถ้าจะวาด “ภาพสถาปัตยกรรมข้อมูล” แบบสั้นที่สุด**

อ้างอิงจาก flow ของ Service\_Master.gs, Service\_SCG.gs, Service\_GPSFeedback.gs, Service\_Agent.gs, Service\_Search.gs, และ WebApp.gs. 

---

# **6\) ข้อสรุปเชิงใช้งานสำหรับคุณ**

ถ้าคุณจะดู repo นี้ในมุม “ระบบข้อมูล” ให้จำ 3 แกนนี้:

1. Database \= ความจริงกลางของลูกค้า.   
2. NameMapping \= ชั้นเชื่อมชื่อหลายแบบให้เข้าหา UUID เดียว.   
3. GPS\_Queue \= safety buffer ก่อนเอา GPS จากคนขับกลับเข้า master. 

ส่วน Data เป็นเพียง working sheet สำหรับงานรายวัน ที่ถูก enrich มาจาก master อีกชั้นหนึ่ง. 

---

ถ้าคุณต้องการ ผมทำต่อให้ได้อีก 3 แบบ:

1. ER Diagram ของทุกชีตแบบ relational view  
2. CRUD Matrix ว่าแต่ละฟังก์ชันอ่าน/เขียนชีตไหนบ้าง  
3. Risk Review ว่าควร refactor schema / index / AI flow ตรงไหนก่อนหลัง

ต้องการ ข้อ3 ช่วยทำ Risk Review ว่าควร refactor schema / index / AI flow ตรงไหนก่อนหลัง แบบละเอียดให้ผมด้วยนะครับ  
คัดลอก

## **ได้ครับ — ด้านล่างคือ Risk Review \+ ลำดับ refactor ก่อน/หลัง แบบลงรายละเอียด**

ผมจะประเมินใน 4 มิติ:

1. ความเสี่ยงด้านข้อมูลเสียหาย / data integrity  
2. ความเสี่ยงด้าน schema/index drift  
3. ความเสี่ยงด้าน AI flow / quality drift  
4. ความเสี่ยงด้าน scalability / maintainability

และจะสรุปเป็น ลำดับ refactor ที่ควรทำก่อน-หลัง ให้ชัดเจนครับ.  
อ้างอิงทั้งหมดมาจาก source repo ที่ผมอ่านในรอบก่อน เช่น Config.gs, Service\_Master.gs, Service\_SCG.gs, Service\_GPSFeedback.gs, Service\_Agent.gs, Service\_Search.gs, Service\_SchemaValidator.gs, Utils\_Common.gs, และไฟล์ admin/setup/test ต่าง ๆ. 

---

# **1\) Executive Summary — ถ้าจะ refactor จริง ควรทำอะไรก่อน**

## **ลำดับที่ควรทำก่อนที่สุด**

### **Priority 0 — กันข้อมูลเพี้ยน / กันข้อมูลพัง**

1. ทำให้ schema 22 คอลัมน์ของ Database ถูกใช้อย่างสม่ำเสมอทุก flow.   
2. แก้ queue reject workflow ให้มีสถานะ REJECTED จริง.   
3. แก้ logic ที่น่าสงสัยใน repairNameMapping\_Full().   
4. เพิ่ม validation/invariant checks ระดับ row ก่อนเขียนกลับชีต. 

### **Priority 1 — ลดความเสี่ยง schema/index drift**

5. สร้าง header constants และ schema contract ครบทุกชีต แทนการอ้างตำแหน่ง array ล้วน ๆ.   
6. รวม 17-column world กับ 22-column world ให้เหลือชุดเดียว.   
7. สร้าง accessor/helper สำหรับ row objects แทน raw index. 

### **Priority 2 — ปรับ AI flow ให้แม่นขึ้นและเสถียรกว่าเดิม**

8. เปลี่ยน AI candidate selection จาก “slice 500 แถวแรก” เป็น retrieval ก่อน AI.   
9. แยก normalized field ออกจาก AI keywords field.   
10. เพิ่ม audit log ของ AI mapping. 

### **Priority 3 — scalability / maintainability**

11. ลดการ setValue ทีละ cell ใน loop ไปเป็น batch write.   
12. สร้าง service layer ชัดเจน: Repository / Domain / Workflow.   
13. เพิ่ม test harness ที่ไม่ผูกกับ UI alert อย่างเดียว. 

---

# **2\) Risk Matrix แบบละเอียด**

---

## **R1. Schema Drift ระหว่าง “17 คอลัมน์เดิม” กับ “22 คอลัมน์ v4.1”**

### **ระดับความเสี่ยง: สูงมาก**

### **อาการใน code**

แม้ Config.gs กำหนด Database ถึงคอลัมน์ 22 (Coord\_\*, Record\_Status, Merged\_To\_UUID) แต่หลายฟังก์ชันยังอ่าน/เขียนเพียง 17 คอลัมน์. 

ตัวอย่าง:

* finalizeAndClean\_MoveToMapping() อ่าน 17 คอลัมน์, clear 17 คอลัมน์, set 17 คอลัมน์.   
* recalculateAllQuality() ก็อ่าน 17 คอลัมน์.   
* searchMasterData() อ่าน 17 คอลัมน์ ทำให้คอลัมน์ใหม่ v4.1 ไม่มีผลกับ search. 

### **ความเสี่ยงจริง**

* flow ใหม่ของ v4.1 อาจ “มีอยู่ใน schema แต่ไม่ถูกใช้จริงสม่ำเสมอ”  
* ถ้ามีการย้ายลำดับคอลัมน์ในอนาคต ความเสียหายจะกระจายเร็วมาก  
* คนดูแลระบบอาจเข้าใจผิดว่าคอลัมน์ R-V ถูกใช้อยู่ทุก flow ทั้งที่ยังไม่ครบ. 

### **ควร refactor อย่างไร**

#### **Phase 1**

* สร้าง constant DB\_TOTAL\_COLS \= 22  
* เปลี่ยนทุก getRange(..., 17\) ที่เกี่ยวกับ Database ให้ใช้ CONFIG.COL\_MERGED\_TO\_UUID หรือ helper เดียว. 

#### **Phase 2**

* แยก helper:  
  * readDatabaseRows\_()  
  * writeDatabaseRows\_()  
  * getDatabaseRowObject\_(rowArray)  
  * toDatabaseRowArray\_(rowObject)

#### **ทำก่อนหรือหลัง?**

ทำก่อนที่สุด เพราะเป็นรากของ data integrity ทั้งระบบ. 

---

## **R2.** repairNameMapping\_Full() **มี logic fallback UUID ที่ผิดเจตนาได้**

### **ระดับความเสี่ยง: สูง**

### **อาการใน code**

บรรทัดนี้น่าสงสัยมาก:

if (\!uid) uid \= uuidMap\[normalizeText(r\[1\])\] || generateUUID();

โดย r\[1\] คือคอลัมน์ UID ของ NameMapping เอง ไม่ใช่ชื่อ. จึงแทบไม่มีเหตุผลที่จะ normalize UID แล้วไป lookup uuidMap ที่ key มาจากชื่อ normalized. 

### **ความเสี่ยงจริง**

* อาจ generate UUID ใหม่ให้ mapping ที่ควรจะชี้ไป UUID เดิม  
* ทำให้ alias หลุดจาก master จริง  
* search / apply coordinates / AI post-processing อาจผิดปลายทาง. 

### **ควร refactor อย่างไร**

* แก้ fallback logic ให้ใช้ oldN หรือ lookup จากชื่อ variant ที่ถูกต้อง  
* ถ้า uid หายจริง ควร:  
  1. หา UUID จาก Database.NAME  
  2. ถ้าไม่เจอ ให้ mark invalid row  
  3. อย่า generate UUID ใหม่ทันที เว้นแต่ตั้งใจสร้าง master ใหม่จริง

### **ทำก่อนหรือหลัง?**

ทำทันทีในกลุ่ม Priority 0 เพราะเสี่ยงทำ identity map เสีย. 

---

## **R3. GPS Queue มี approve path ชัด แต่ reject path ไม่ครบ**

### **ระดับความเสี่ยง: สูง**

### **อาการใน code**

applyApprovedFeedback() ประมวลผลเฉพาะ Approve=true && Reject=false แล้ว set reason เป็น APPROVED. แต่ถ้า reject ถูกติ๊ก จะถูกนับเป็น skipped และไม่ stamp REJECTED กลับลงชีต. 

### **ความเสี่ยงจริง**

* admin audit ไม่ชัดว่า reject ไปแล้วหรือยัง  
* showGPSQueueStats() รองรับ REJECTED ทางตรรกะ แต่ workflow ปัจจุบันแทบไม่สร้างสถานะนี้จาก flow หลัก.   
* queue อาจสะสมรายการ ambiguous ที่จริงถูก reject ไปแล้วแต่ยังดู pending/checked ค้าง

### **ควร refactor อย่างไร**

* เพิ่ม path:  
  * ถ้า Reject=true && Approve=false → set Reason="REJECTED"  
  * clear checkbox ฝั่งตรงข้าม  
  * อาจเขียน Reviewed\_At, Reviewed\_By, Review\_Note  
* ป้องกันกรณี Approve+Reject พร้อมกันด้วย validation. 

### **ทำก่อนหรือหลัง?**

ทำเร็วมาก เพราะเกี่ยวกับ human approval governance โดยตรง. 

---

## **R4. AI Tier-4 ใช้ candidate subset แบบ “500 แถวแรก”**

### **ระดับความเสี่ยง: สูง**

### **อาการใน code**

resolveUnknownNamesWithAI() สร้าง masterOptions ทั้งหมดแล้ว slice(0, 500\) ตัดตรง ๆ ก่อนส่งให้ Gemini. 

### **ความเสี่ยงจริง**

* ถ้า DB โตขึ้น accuracy จะขึ้นอยู่กับ “ลำดับในชีต” ไม่ใช่ความเกี่ยวข้อง  
* candidate ที่จริงควร match อาจไม่ถูกส่งเข้า AI เลย  
* user จะเข้าใจว่า AI “จับคู่ไม่เก่ง” ทั้งที่ retrieval ผิดตั้งแต่ก่อน AI

### **ควร refactor อย่างไร**

#### **ขั้นที่แนะนำ**

1. retrieval ก่อน AI:  
   * tokenize unknown name  
   * หา top-N candidate จาก Database.NAME, NameMapping.Variant, NORMALIZED  
2. ส่งเฉพาะ top 20-50 ที่เกี่ยวข้องที่สุดเข้า Gemini  
3. เก็บ score retrieval \+ score AI แยกกัน. 

### **ทำก่อนหรือหลัง?**

ทำหลัง schema/data integrity แต่ก่อน optimization ระยะยาว. 

---

## **R5. AI keywords ถูกเก็บรวมกับ** NORMALIZED

### **ระดับความเสี่ยง: กลางถึงสูง**

### **อาการใน code**

processAIIndexing\_Batch() เขียน basicKey \+ aiKeywords \+ \[AI\] ลง COL\_NORMALIZED.   
ขณะที่ searchMasterData() ก็ใช้คอลัมน์นี้เป็น aiKeywords. 

### **ความเสี่ยงจริง**

* field เดียวทำหน้าที่ 3 อย่างพร้อมกัน:  
  1. normalized text  
  2. AI-generated keyword bag  
  3. tag state \[AI\]  
* แยก debug ยาก  
* ranking/search quality จะมั่วง่าย  
* AI reset logic ก็ต้องมาลบ tag จาก field เดียวกัน. 

### **ควร refactor อย่างไร**

แยกเป็น:

* NORMALIZED\_NAME  
* SEARCH\_KEYWORDS  
* AI\_STATUS  
* AI\_UPDATED\_AT

### **ทำก่อนหรือหลัง?**

ทำใน Priority 2 หลังแก้ schema/identity risk แล้ว. 

---

## **R6. Search ranking ยังหยาบและไม่ใช้ metadata ใหม่ของ v4.1**

### **ระดับความเสี่ยง: กลาง**

### **อาการใน code**

search ใช้ all-token match และ score แค่ 10/1. ไม่ใช้ verified, coord source, confidence, quality, record status. 

### **ความเสี่ยงจริง**

* result ที่ “match text” แต่คุณภาพต่ำ อาจขึ้นเท่ากับ result คุณภาพสูง  
* merged/inactive records อาจยังหลุดขึ้นถ้าอยู่ใน DB  
* user ฝั่ง webapp อาจเจอผลค้นหาที่ไม่ใช่ record ดีสุด. 

### **ควร refactor อย่างไร**

เพิ่ม weighted ranking:

* exact name match \+50  
* alias match \+40  
* normalized token match \+20  
* verified \+20  
* coord source \= Driver\_GPS \+10  
* quality score / confidence score มีน้ำหนัก  
* inactive / merged \= ลดหนักหรือ exclude

### **ทำก่อนหรือหลัง?**

ทำหลังแยก field AI/search และหลัง unify schema. 

---

## **R7. ใช้** setValue() **ทีละ cell ใน loop หลายจุด**

### **ระดับความเสี่ยง: กลาง**

### **อาการใน code**

* syncNewDataToMaster() อัปเดต timestamp ทีละแถว.   
* applyApprovedFeedback() อัปเดต DB ทีละ cell ต่อหนึ่ง queue row. 

### **ความเสี่ยงจริง**

* ช้าเมื่อข้อมูลโต  
* เสี่ยง partial write ถ้ารันกลางทางแล้ว error  
* quota consumption สูงกว่าจำเป็น

### **ควร refactor อย่างไร**

* ดึง row ทั้งก้อนมาแก้ใน memory แล้ว setValues() ทีเดียว  
* ถ้าต้องเขียนเฉพาะบางคอลัมน์ ให้ batch range ที่ต่อเนื่อง  
* เพิ่ม SpreadsheetApp.flush() เฉพาะจุดจำเป็นจริง

### **ทำก่อนหรือหลัง?**

Priority 3 ยกเว้นถ้าตอนนี้ performance เริ่มช้าแล้ว ให้เลื่อนขึ้นได้. 

---

## **R8.** normalizeText() **รุนแรงและอาจลบ semantic สำคัญเกินไป**

### **ระดับความเสี่ยง: กลาง**

### **อาการใน code**

stop words มีทั้ง “ร้าน”, “สาขา”, “โกดัง”, “คลังสินค้า”, “จังหวัด”, “อำเภอ”, “store”, “shop” ฯลฯ. 

### **ความเสี่ยงจริง**

* branch/store distinctions หาย  
* ชื่อสั้น ๆ หลายร้านอาจ collapse มาหากัน  
* AI/search/cluster อาจ over-merge

### **ควร refactor อย่างไร**

* แยก normalization เป็น 2 ระดับ:  
  * normalizeForIdentity()  
  * normalizeForSearch()  
* อย่าลบคำว่า สาขา เสมอไปในทุก use case  
* เก็บ normalized variants หลายแบบ

### **ทำก่อนหรือหลัง?**

Priority 2-3 แต่ควรเริ่ม design เร็ว เพราะไปแตะทุก flow. 

---

## **R9. Soft delete มีแล้ว แต่ search / apply flows ยังไม่ใช้** resolveUUID() **อย่างทั่วถึง**

### **ระดับความเสี่ยง: กลาง**

### **อาการใน code**

ระบบมี mergeUUIDs() และ resolveUUID() ดีแล้ว.   
แต่ flow search/apply/sync ยังไม่ได้ใช้ resolveUUID() อย่าง explicit ในทุกจุด. 

### **ความเสี่ยงจริง**

* alias อาจชี้ไป UUID ที่ถูก merge แล้ว  
* search ยังอาจแสดง record ที่ไม่ใช่ปลายทาง canonical  
* coordinate apply อาจยังยึด UUID เดิมที่ควรถูกรวมแล้ว

### **ควร refactor อย่างไร**

* ทุกจุดที่อ่าน UUID จาก NameMapping ควรผ่าน resolveUUID() ก่อน  
* search ควร exclude Record\_Status \= Merged/Inactive หรือ redirect ไป active UUID

### **ทำก่อนหรือหลัง?**

Priority 1-2 หลัง unify schema. 

---

## **R10. Diagnostic/Test ยังเป็น manual smoke test มากกว่า automated contract test**

### **ระดับความเสี่ยง: กลาง**

### **อาการใน code**

Test\_AI.gs และ Test\_Diagnostic.gs เป็น UI-triggered manual tests ผ่าน alert/toast. 

### **ความเสี่ยงจริง**

* regression detect ยาก  
* change index/schema แล้วพังเงียบได้  
* AI/network-dependent logic test ซ้ำยาก

### **ควร refactor อย่างไร**

* เพิ่ม test helpers ที่คืน object/json แทน alert  
* เพิ่ม assertSchema\_, assertDatabaseRowShape\_, assertQueueRowShape\_  
* ทำ dry-run mode สำหรับ sync/apply/finalize

### **ทำก่อนหรือหลัง?**

Priority 3 แต่มี ROI สูงมากถ้าจะเริ่ม refactor ใหญ่. 

---

# **3\) ลำดับ Refactor Roadmap แบบแนะนำ**

---

## **Phase 0 — “กันพังก่อน” (ควรทำก่อนสุด)**

### **เป้าหมาย**

หยุดความเสี่ยงที่ทำให้ identity หรือ state ของข้อมูลผิด

### **งาน**

1. แก้ repairNameMapping\_Full() fallback UUID bug.   
2. เพิ่ม reject workflow จริงใน GPS\_Queue.   
3. เพิ่ม invariant checks:  
   * UUID\_DB ต้องมีใน DB ก่อนอัปเดต  
   * LAT/LNG ต้อง parse ได้  
   * Approve และ Reject ห้าม true พร้อมกัน  
4. เพิ่ม dry-run summary ใน flow เสี่ยง:  
   * sync  
   * apply approved  
   * finalize

### **ผลลัพธ์**

ลดโอกาสข้อมูลผิดเส้นทางและทำให้ human review audit ได้ชัด. 

---

## **Phase 1 — “รวม schema ให้เป็นหนึ่งเดียว”**

### **เป้าหมาย**

ไม่ให้ระบบครึ่งหนึ่งอยู่โลก 17 คอลัมน์ อีกครึ่งอยู่โลก 22 คอลัมน์

### **งาน**

1. สร้าง helper กลางสำหรับ Database read/write.   
2. แทนค่าพวก 17, 22 ด้วย constant/function เดียว.   
3. อัปเดต:  
   * finalizeAndClean\_MoveToMapping()  
   * recalculateAllQuality()  
   * searchMasterData()  
   * flow อื่นที่ยังใช้ 17 คอลัมน์.   
4. ขยาย schema validator ให้ assert U/V ด้วย. 

### **ผลลัพธ์**

refactor ครั้งต่อ ๆ ไปปลอดภัยขึ้นมาก. 

---

## **Phase 2 — “จัดระเบียบ index และ row model”**

### **เป้าหมาย**

ลด bug จาก raw array index

### **งาน**

1. สร้าง row adapters:  
   * dbRowToObject  
   * dbObjectToRow  
   * queueRowToObject  
   * dataRowToObject  
2. สร้าง field enums ต่อชีต ไม่ใช้ r\[10\], r\[26\], r\[28\] กระจัดกระจาย.   
3. รวม header definitions กับ schema validator ไว้จุดเดียว

### **ผลลัพธ์**

อ่านโค้ดง่ายขึ้น, แก้ schema ภายหลังง่ายขึ้น, ลด off-by-one risk. 

---

## **Phase 3 — “ยกเครื่อง AI flow”**

### **เป้าหมาย**

ให้ AI flow มีคุณภาพและ auditability จริง

### **งาน**

1. retrieval ก่อน AI แทน slice(0,500).   
2. แยก fields:  
   * normalized  
   * search keywords  
   * AI status  
   * AI updated time.   
3. เก็บ AI decision log:  
   * unknown input  
   * candidate list  
   * chosen UUID  
   * AI confidence  
   * prompt version / model  
4. เพิ่ม threshold strategy:  
   * 90+ auto-map  
   * 70-89 queue for review  
   * \<70 ignore

### **ผลลัพธ์**

AI จะไม่ใช่ black box มากเกินไป และปรับปรุงทีหลังง่าย. 

---

## **Phase 4 — “ปรับ Search และ consumer flow”**

### **เป้าหมาย**

ให้ระบบปลายทางใช้ metadata ใหม่จาก v4.1 จริง

### **งาน**

1. ranking ใหม่ใน searchMasterData().   
2. exclude / de-prioritize merged/inactive record.   
3. ใช้ Coord\_Source, Coord\_Confidence, Quality, Verified ใน score.   
4. UI แสดง badge เช่น:  
   * Driver GPS  
   * Verified  
   * Low confidence warning

### **ผลลัพธ์**

end-user จะได้ผลค้นหาที่ “ใช้งานจริง” มากขึ้น ไม่ใช่แค่ textual match. 

---

## **Phase 5 — “Performance / Ops / Testing”**

### **เป้าหมาย**

รองรับ scale และเปลี่ยนแปลงได้ปลอดภัย

### **งาน**

1. batch write ลด setValue() ใน loop.   
2. เพิ่ม repository/service separation.  
3. เพิ่ม function แบบ non-UI return object เพื่อทดสอบง่าย.  
4. เพิ่ม regression diagnostics:  
   * schema contract test  
   * search relevance test  
   * queue apply dry-run test  
   * AI prompt parse test

### **ผลลัพธ์**

ระบบ maintainable ระยะยาวและพร้อมโต. 

---

# **4\) ถ้าผมต้องจัดอันดับ “Top 10 จุดที่ควรแก้ก่อน” แบบชัดที่สุด**

## **1\. แก้** repairNameMapping\_Full() **ก่อน**

เพราะมันแตะ identity map โดยตรง. 

## **2\. ทำ reject workflow ให้สมบูรณ์**

เพราะตอนนี้ human governance ยังไม่ complete. 

## **3\. รวม** Database **schema ให้ทุก flow ใช้ 22 คอลัมน์ชุดเดียว**

นี่คือ structural debt ใหญ่สุด. 

## **4\. เพิ่ม helper สำหรับ database row model**

ลด index bug ระยะยาว. 

## **5\. ให้ทุก UUID ผ่าน** resolveUUID() **ก่อนใช้จริง**

เพื่อรองรับ merge correctly. 

## **6\. เปลี่ยน AI candidate selection เป็น retrieval-based**

ช่วย accuracy มากที่สุดต่อ effort ใน AI flow. 

## **7\. แยก** NORMALIZED **ออกจาก AI/search metadata**

ตอนนี้ field เดียวแบกหลายหน้าที่เกินไป. 

## **8\. ยกระดับ search ranking**

เพื่อให้ user ได้ผลลัพธ์ที่เชื่อถือได้มากขึ้น. 

## **9\. batch write ทุกจุดที่ update loop หนัก**

เพื่อ performance และลด partial write risk. 

## **10\. เพิ่ม automated contract tests**

เพื่อให้ refactor ใหญ่ทำได้อย่างปลอดภัย. 

---

# **5\) Recommendation เชิงกลยุทธ์: ถ้าคุณมีเวลาแค่ 1 สัปดาห์**

## **Day 1-2**

* แก้ NameMapping repair bug  
* เพิ่ม reject path  
* เพิ่ม invariant checks ใน queue/apply/sync. 

## **Day 3-4**

* unify Database read/write เป็น 22 columns  
* refactor functions ให้อ่านผ่าน helper กลาง. 

## **Day 5**

* ใช้ resolveUUID() กับ search/apply/alias flows. 

## **Day 6**

* เปลี่ยน AI retrieval  
* แยก search keywords จาก normalized field. 

## **Day 7**

* เพิ่ม regression diagnostics / smoke tests / dry runs. 

---

# **6\) สรุปสุดท้ายแบบสั้นที่สุด**

ถ้าพูดตรงที่สุด:

* ความเสี่ยงใหญ่สุดตอนนี้ไม่ใช่ AI แต่คือ schema/index inconsistency ระหว่างโลกเก่า 17 คอลัมน์กับโลกใหม่ 22 คอลัมน์.   
* ความเสี่ยงอันดับ 2 คือ identity governance โดยเฉพาะ NameMapping repair และ UUID merge flow ที่ยังไม่ได้ถูก consume ทั่วระบบ.   
* ความเสี่ยงอันดับ 3 คือ human review workflow ของ GPS queue ที่ approve ได้แต่ reject ยังไม่สมบูรณ์.   
* หลังจากนั้นค่อยไปยกเครื่อง AI และ search ranking จะคุ้มที่สุด. 

---

## **Commands ที่ผมใช้ในการตรวจ source รอบนี้และรอบก่อน**

* rg \--files  
* nl \-ba Config.gs  
* nl \-ba Service\_Master.gs | sed \-n '1,240p'  
* nl \-ba Service\_Master.gs | sed \-n '241,420p'  
* nl \-ba Service\_Master.gs | sed \-n '421,760p'  
* nl \-ba Service\_SCG.gs | sed \-n '1,220p'  
* nl \-ba Service\_SCG.gs | sed \-n '221,450p'  
* nl \-ba Service\_GPSFeedback.gs  
* nl \-ba Service\_Agent.gs  
* nl \-ba Service\_AutoPilot.gs  
* nl \-ba Service\_Search.gs  
* nl \-ba Service\_SchemaValidator.gs  
* nl \-ba Service\_SoftDelete.gs  
* nl \-ba Utils\_Common.gs  
* nl \-ba Test\_AI.gs  
* nl \-ba Test\_Diagnostic.gs

# Refactor Blueprint

## **ด้านล่างคือ Refactor Blueprint แบบลงถึงระดับ file / function / ลำดับการแก้ อย่างละเอียด**

ผมจะทำเป็น “แผนปฏิบัติการ” ที่ตอบ 4 อย่างพร้อมกัน:

1. แก้ไฟล์ไหนก่อน  
2. แก้ฟังก์ชันไหนก่อน  
3. จะแก้อย่างไรแบบ concrete  
4. ทำแล้วจะลดความเสี่ยงอะไร

ทั้งหมดนี้อ้างอิงจาก source จริงใน repo เช่น Config.gs, Service\_Master.gs, Service\_GPSFeedback.gs, Service\_Agent.gs, Service\_Search.gs, Service\_SchemaValidator.gs, Utils\_Common.gs, Service\_SCG.gs, Service\_SoftDelete.gs, Menu.gs, Test\_AI.gs, และ Test\_Diagnostic.gs. 

---

# **1\) เป้าหมายของ Blueprint นี้**

Blueprint นี้ไม่ได้เป็นแค่ “ข้อเสนอทั่วไป” แต่เป็น ลำดับ refactor ที่เอาไปทำงานจริงได้ โดยยึดหลัก:

* ห้ามแก้ใหญ่พร้อมกันทั้ง repo  
* ต้องลดความเสี่ยงข้อมูลก่อน  
* ต้องสร้างจุดรวมของ schema/index ก่อนค่อยขยายไป AI/search  
* ทุก phase ควร deploy ได้ทีละก้อน. 

---

# **2\) ภาพรวมลำดบทำงาน**

## **ลำดับที่แนะนำ**

### **Phase A — Stabilize Data Integrity**

* Config.gs  
* Service\_SchemaValidator.gs  
* Service\_Master.gs  
* Service\_GPSFeedback.gs  
* Service\_SoftDelete.gs

### **Phase B — Normalize Schema Access Layer**

* Config.gs  
* Utils\_Common.gs  
* Service\_Master.gs  
* Service\_SCG.gs  
* Service\_Search.gs

### **Phase C — Fix Identity & Mapping Pipeline**

* Service\_Master.gs  
* Service\_SoftDelete.gs  
* Service\_Search.gs  
* Service\_SCG.gs

### **Phase D — Refactor AI Flow**

* Service\_AutoPilot.gs  
* Service\_Agent.gs  
* Service\_Search.gs  
* Test\_AI.gs

### **Phase E — Refactor UI / Search / Diagnostics**

* WebApp.gs  
* Index.html  
* Menu.gs  
* Test\_Diagnostic.gs

---

# **3\) Refactor Blueprint ราย Phase**

---

## **Phase A — Stabilize Data Integrity**

เป้าหมาย: ทำให้ระบบ “ไม่ทำข้อมูลเพี้ยน / ไม่ทำ state เพี้ยน” ก่อนอย่างอื่น

---

### **A1) แก้** Config.gs **ให้เป็น Source of Truth ที่แข็งแรงกว่าเดิม**

## **ไฟล์**

* Config.gs. 

## **ฟังก์ชัน/ส่วนที่เกี่ยวข้อง**

* CONFIG  
* SCG\_CONFIG  
* C\_IDX  
* MAP\_IDX  
* validateSystemIntegrity(). 

## **สิ่งที่ควรแก้**

### **1\. เพิ่ม constants สำหรับจำนวนคอลัมน์ต่อชีต**

ตอนนี้หลายไฟล์ hardcode 17, 22, 29, 9, 5. ควรย้ายมาไว้ใน config เช่น:

* DB\_TOTAL\_COLS \= 22  
* DB\_LEGACY\_COLS \= 17 ถ้ายังจำเป็นต้องมีช่วง transition  
* MAP\_TOTAL\_COLS \= 5  
* GPS\_QUEUE\_TOTAL\_COLS \= 9  
* DATA\_TOTAL\_COLS \= 29. 

### **2\. เพิ่ม header constants ของ** Database

ตอนนี้ Database มีความหมายผ่าน COL\_\* แต่ไม่มี header list กลางเหมือน Data หรือ GPS\_Queue. ควรสร้าง:

* DB\_HEADERS  
* MAP\_HEADERS  
* GPS\_QUEUE\_HEADERS  
* DATA\_HEADERS  
  เพื่อให้ schema validator, upgrade, setup, tests ใช้ชุดเดียวกัน. 

### **3\. เพิ่ม field groups**

เช่น:

* DB\_REQUIRED\_FIELDS\_CORE  
* DB\_REQUIRED\_FIELDS\_COORD  
* DB\_REQUIRED\_FIELDS\_STATUS  
* DB\_REQUIRED\_FIELDS\_SEARCH  
  เพื่อให้ validator/test/report reuse ได้. 

## **ผลลัพธ์ที่ได้**

* ลด magic number ทั่ว repo  
* ทำให้ refactor phase ถัดไปง่ายขึ้นมาก  
* ทำให้ schema changes ในอนาคตกระทบจุดเดียว. 

---

### **A2) ขยาย** Service\_SchemaValidator.gs **ให้เช็ก schema “จริง” ของ v4.1**

## **ไฟล์**

* Service\_SchemaValidator.gs. 

## **ฟังก์ชันที่ควรแก้**

* SHEET\_SCHEMA  
* validateSheet\_()  
* preCheck\_Sync()  
* preCheck\_Apply()  
* preCheck\_Approve()  
* runFullSchemaValidation(). 

## **สิ่งที่ควรแก้**

### **1\. แก้** DATABASE.minColumns

ตอนนี้ DATABASE.minColumns เป็น 20 ทั้งที่ config ใช้ถึง 22\. ควรอัปเดตให้สะท้อน schema v4.1 เต็มจริง. 

### **2\. เพิ่ม required headers สำหรับ U/V**

* Record\_Status  
* Merged\_To\_UUID.  
  ตอนนี้ validator ยังไม่ assert 2 คอลัมน์นี้. 

### **3\. เพิ่ม validation สำหรับ** Data

ตอนนี้ DATA เช็กขั้นต่ำแค่บางคอลัมน์. ควรเช็ก header สำคัญมากขึ้น เช่น:

* DriverName  
* Email พนักงาน  
* LatLong\_Actual  
* ShopKey. 

### **4\. เพิ่ม validation สำหรับ queue state**

GPS\_Queue ควรมี validation rule:

* ห้าม Approve และ Reject เป็น true พร้อมกัน  
* UUID\_DB ต้องไม่ว่างในแถวที่ pending review  
* LatLng\_Driver parse ได้. 

## **ผลลัพธ์ที่ได้**

* ปัญหา schema drift จะถูกเจอก่อน run flow  
* ลดโอกาส data corruption แบบ silent. 

---

### **A3) แก้** Service\_Master.gs **จุดเสี่ยงก่อน**

## **ไฟล์**

* Service\_Master.gs. 

## **ฟังก์ชันที่ควรแก้ก่อนสุด**

1. repairNameMapping\_Full()  
2. finalizeAndClean\_MoveToMapping()  
3. syncNewDataToMaster()  
4. recalculateAllQuality()  
5. recalculateAllConfidence(). 

### **จุดแก้ 1 —** repairNameMapping\_Full()

#### **ปัญหา**

fallback UUID logic ผิดเจตนาได้สูง. 

#### **วิีแก้**

* แยก helper ชัด:  
  1. resolveMappingTargetUUID\_(variantName, providedUid, dbIndexByNormName)  
* ถ้า providedUid ว่าง:  
  1. lookup จาก variantName  
  2. ถ้าไม่เจอ ให้ mark invalid  
  3. อย่า generate UUID ใหม่มั่ว เว้นแต่เป็น flow สร้าง master ใหม่จริง

#### **ควรเพิ่ม**

* report จำนวน rows ที่ถูก fix  
* report จำนวน rows invalid  
* dry-run mode

---

### **จุดแก้ 2 —** finalizeAndClean\_MoveToMapping()

#### **ปัญหา**

ทำงานแค่ 17 คอลัมน์ ทั้งที่ schema ปัจจุบัน 22 คอลัมน์. 

#### **วิธีแก้**

* เปลี่ยนทุก 17 เป็น CONFIG.COL\_MERGED\_TO\_UUID  
* ใช้ row object แทน row\[CONFIG.C\_IDX.\*\] กระจายทั้งฟังก์ชัน  
* เวลาสร้าง backup ให้บันทึก metadata เช่น:  
  * backup reason  
  * initiated by  
  * timestamp

#### **สิ่งที่ควรเพิ่ม**

* ก่อน finalize ให้ validate:  
  * UUID ไม่ซ้ำ  
  * ชื่อที่ verified ไม่มี duplicate UUID conflict  
  * suggested names ถ้าจะ map ต้องมี target UUID ชัด  
* หลัง finalize ให้ clear search cache ด้วย เพราะ alias map เปลี่ยน. 

---

### **จุดแก้ 3 —** syncNewDataToMaster()

#### **ปัญหา**

* ยังเขียนบางอย่างทีละ row  
* queue payload ยังไม่ละเอียดพอสำหรับ review  
* ยังไม่มี rejected/duplicate protection ระดับ queue business rule. 

#### **วิธีแก้**

* สร้าง helper:  
  * buildNewDatabaseEntryFromSourceRow\_(row, ts)  
  * buildGPSQueueEntry\_(reason, name, uuid, driverLatLng, dbLatLng, diffMeters, ts)  
* batch write ให้เป็นชุด  
* ก่อน append queue ควรเช็กว่ามี pending item เดิมสำหรับ UUID เดียวกัน/พิกัดเดียวกันหรือไม่

#### **ควรเพิ่ม field ใน queue หรือไม่**

ถ้าจะ refactor ลึก แนะนำเพิ่ม:

* Reviewed\_At  
* Reviewed\_By  
* Review\_Note  
* Queue\_Status  
  แต่ถ้ายังไม่อยากเปลี่ยน schema ใหญ่ ให้ใช้ Reason \+ metadata columns ใหม่ภายหลัง. 

---

### **จุดแก้ 4 —** recalculateAllQuality() **/** recalculateAllConfidence()

#### **ปัญหา**

2 ฟังก์ชันนี้มี scoring logic แยกต่างหาก ทำให้เสี่ยง divergence กับ deep clean logic. 

#### **วิธีแก้**

* แยก helper กลาง:  
  * computeQualityScore\_(rowObj)  
  * computeConfidenceScore\_(rowObj)  
* ให้ runDeepCleanBatch\_100(), recalculateAllQuality(), recalculateAllConfidence() ใช้ helper เดียวกัน

#### **ผลลัพธ์**

* ลด scoring drift  
* ปรับกติกาในอนาคตง่าย. 

---

### **A4) แก้** Service\_GPSFeedback.gs **ให้ workflow สมบูรณ์**

## **ไฟล์**

* Service\_GPSFeedback.gs. 

## **ฟังก์ชันที่ควรแก้**

* applyApprovedFeedback()  
* showGPSQueueStats()  
* upgradeGPSQueueSheet()  
* createGPSQueueSheet(). 

### **จุดแก้ 1 —** applyApprovedFeedback()

#### **วิธีแก้**

* รองรับ 3 state ชัดเจน:  
  * PENDING  
  * APPROVED  
  * REJECTED  
* rule:  
  * approve only → update DB \+ stamp approved  
  * reject only → stamp rejected  
  * both true → error  
  * both false → pending

#### **ถ้าจะเปลี่ยน schema**

เพิ่ม columns:

* Review\_Status  
* Reviewed\_At  
* Reviewed\_By  
* Review\_Note

#### **ถ้ายังไม่เปลี่ยน schema**

ใช้ Reason เป็น state machine ชั่วคราว แต่ต้อง write REJECTED จริง. 

### **จุดแก้ 2 — batch update**

ตอนนี้ update LAT, LNG, Coord\_Source, Coord\_Confidence, Coord\_Last\_Updated, UPDATED ทีละ cell ต่อ row. ควร:

* โหลด DB ทั้งก้อน  
* mutate row ใน memory  
* setValues กลับทีเดียว. 

### **จุดแก้ 3 —** showGPSQueueStats()

ให้ย้ายไปนับจาก explicit state แทน infer จาก checkbox \+ reason ปนกัน. 

---

### **A5) ทำให้** Service\_SoftDelete.gs **ถูก consume จริงทั่วระบบ**

## **ไฟล์**

* Service\_SoftDelete.gs. 

## **ฟังก์ชันที่ต้องเป็นฐานให้ไฟล์อื่นเรียก**

* resolveUUID(uuid)  
* mergeUUIDs(masterUUID, duplicateUUID)  
* softDeleteRecord(uuid). 

## **สิ่งที่ควรทำ**

* ประกาศให้ชัดว่า ทุกที่ที่อ่าน UUID จาก mapping ต้องผ่าน resolveUUID()  
* เพิ่ม helper:  
  * resolveRowUUIDOrNull\_(uuid)  
  * isActiveUUID\_(uuid)  
* Search / apply / sync ควร ignore record ที่ Inactive และ redirect record ที่ Merged

## **ไฟล์ถัดไปที่จะต้องแก้ตาม**

* Service\_SCG.gs  
* Service\_Search.gs  
* Service\_Master.gs. 

---

## **Phase B — Normalize Schema Access Layer**

เป้าหมาย: ทำให้ทุกไฟล์เลิกอ้าง raw array index แบบกระจัดกระจาย

---

### **B1) สร้าง Row Access Layer**

## **ไฟล์หลักที่จะเริ่ม**

* Utils\_Common.gs  
* หรือสร้างไฟล์ใหม่เชิงโครงสร้าง เช่น Utils\_RowModel.gs ถ้าจะจัด repo ระยะยาว  
  แต่ถ้าคุณอยาก refactor แบบ minimally invasive ให้เริ่มใน Utils\_Common.gs ก่อน. 

## **สิ่งที่ควรเพิ่ม**

### **สำหรับ Database**

* dbRowToObject(row)  
* dbObjectToRow(obj)

### **สำหรับ NameMapping**

* mapRowToObject(row)  
* mapObjectToRow(obj)

### **สำหรับ Data**

* dailyJobRowToObject(row)  
* dailyJobObjectToRow(obj)

### **สำหรับ GPS Queue**

* queueRowToObject(row)  
* queueObjectToRow(obj)

## **ทำไมต้องทำ**

ตอนนี้โค้ดเต็มไปด้วย:

* r\[10\]  
* r\[22\]  
* r\[26\]  
* r\[28\]  
  ซึ่งอ่านยากและเสี่ยงผิดทันทีเมื่อ schema ขยับ. 

---

### **B2) Refactor** Service\_SCG.gs **ให้ใช้ accessor layer**

## **ไฟล์**

* Service\_SCG.gs. 

## **ฟังก์ชันที่ควรแก้**

1. applyMasterCoordinatesToDailyJob()  
2. buildOwnerSummary()  
3. buildShipmentSummary()  
4. fetchDataFromSCGJWD(). 

### **จุดแก้ 1 —** applyMasterCoordinatesToDailyJob()

#### **ปัญหา**

ใช้ raw positions เช่น r\[10\], r\[22\], r\[4\]. 

#### **วิธีแก้**

* สร้าง DATA\_IDX ใน config หรือใช้ row object  
* แยก helper:  
  * resolveDailyJobCoordinates\_(dailyRowObj, aliasMap, masterMaps)  
  * resolveDriverEmail\_(driverName, empMap)  
* call preCheck\_Apply() ก่อนจริง ๆ ถ้าต้องการใช้ schema validator ให้ครบ flow; ตอนนี้มีฟังก์ชันแต่ applyMasterCoordinatesToDailyJob() ไม่ได้เรียก precheck เอง. 【F:Service\_SchemaValidator.gs†L101-L105

### **จุดแก้ 2 —** fetchDataFromSCGJWD()

#### **วิธีแก้**

* แยก transformation ออก:  
  * flattenShipmentResponse\_(shipments)  
  * aggregateShopMetrics\_(rows)  
  * writeDailyData\_(dataSheet, rows)  
* ย้าย headers ไป config กลาง

### **จุดแก้ 3 — summary builders**

* ใช้ row object แทน r\[9\], r\[2\], r\[3\], r\[5\]  
* แยก isEpodInvoice\_() ออกจาก summary logic ถ้าจะ reuse มากขึ้น. 

---

### **B3) Refactor** Service\_Search.gs **ให้ใช้ field model ใหม่**

## **ไฟล์**

* Service\_Search.gs. 

## **ฟังก์ชันที่ควรแก้**

* searchMasterData()  
* getCachedNameMapping\_(). 

### **สิ่งที่ควรแก้**

1. เปลี่ยน read range เป็น full database width  
2. exclude record ที่ Record\_Status \!== Active  
3. ถ้า UUID ถูก merged ให้ resolve ก่อนแสดง  
4. แยก field:  
   * normalizedName  
   * searchKeywords  
   * aliasText  
5. สร้าง scoring function:  
   * scoreSearchMatch\_(record, queryTokens, context)

### **ถ้าทำต่อเนื่องจาก Phase B**

search จะเป็น consumer ตัวแรกที่ได้ประโยชน์จาก row object เต็ม ๆ. 

---

## **Phase C — Fix Identity & Mapping Pipeline**

เป้าหมาย: ทำให้ UUID, Variant, Merged, Resolved UUID ทำงานเป็นระบบเดียว

---

### **C1) Refactor** Service\_Master.gs **ในส่วน finalize/mapping**

## **ฟังก์ชันที่ควรแก้**

* finalizeAndClean\_MoveToMapping()  
* repairNameMapping\_Full()  
* assignMissingUUIDs(). 

### **แผนแก้**

#### **1\. แยก mapping repository helpers**

* loadDatabaseIndexByUUID\_()  
* loadDatabaseIndexByNormalizedName\_()  
* loadNameMappingRows\_()  
* appendNameMappings\_(rows)

#### **2\. ทำให้ finalize เป็น 3 step ชัดเจน**

* collectFinalizeCandidates\_()  
* buildFinalizeMappingRows\_()  
* rewriteMasterSheetAfterFinalize\_()

#### **3\. เพิ่ม conflict report**

เช่น:

* variant เดียว map ไปหลาย UUID  
* suggested name ไม่มี target UUID  
* verified row ไม่มี UUID

### **ผลลัพธ์**

identity pipeline จะ trace ได้ง่ายกว่าเดิมมาก. 

---

### **C2) ทำให้** Service\_SCG.gs **ใช้** resolveUUID()

## **ฟังก์ชันที่ควรแก้**

* applyMasterCoordinatesToDailyJob(). 

### **วิธีแก้**

* เวลาหา targetUID \= aliasMap\[rawName\]  
* อย่าใช้ทันที  
* เรียก resolveUUID(targetUID) ก่อน  
* ถ้า UUID สุดท้ายเป็น inactive/หาไม่เจอ ให้ fallback ไป name match หรือ mark unresolved

### **เหตุผล**

เพื่อให้ merge UUID แล้ว daily job ยังชี้ไปพิกัด canonical ได้. 

---

### **C3) ทำให้** Service\_Search.gs **และ AI consume canonical UUID**

## **ฟังก์ชัน**

* searchMasterData()  
* resolveUnknownNamesWithAI(). 

### **วิธีแก้**

* candidate ที่ได้จาก mapping ต้องผ่าน resolveUUID()  
* search results ควร return canonical UUID  
* AI write-back ควรเขียน canonical UUID ไม่ใช่ stale UUID

---

## **Phase D — Refactor AI Flow**

เป้าหมาย: ให้ AI มี retrieval, audit, และไม่ปะปนกับ normalized field เดิม

---

### **D1) Refactor** Service\_AutoPilot.gs

## **ฟังก์ชันที่ควรแก้**

* processAIIndexing\_Batch()  
* callGeminiThinking\_JSON()  
* createBasicSmartKey(). 

### **สิ่งที่ควรเปลี่ยน**

#### **1\. เปลี่ยน output target**

จากเขียนลง COL\_NORMALIZED โดยตรง  
เป็นเขียนลง field ใหม่ เช่น:

* SEARCH\_KEYWORDS  
* AI\_STATUS  
* AI\_UPDATED\_AT

#### **2\. แยก 2 concern**

* normalization \= deterministic  
* AI keyword enrichment \= probabilistic

#### **3\. ทำ prompt versioning**

เพิ่ม PROMPT\_VERSION ใน output metadata/log

#### **4\. เพิ่ม parse guard**

ตอนนี้ parse JSON จาก Gemini ถ้า structure เปลี่ยนจะล้มแบบเงียบเป็นบางกรณี ควรมี validator ของ AI response. 

---

### **D2) Refactor** Service\_Agent.gs

## **ฟังก์ชันที่ควรแก้**

* resolveUnknownNamesWithAI()  
* askGeminiToPredictTypos(). 

### **สิ่งที่ควรเปลี่ยน**

#### **1\. Retrieval ก่อน AI**

สร้าง helper:

* retrieveCandidateMasters\_(unknownName, dbRows, mapRows, limit)  
* ใช้ normalize \+ alias \+ token overlap ก่อนส่ง AI

#### **2\. Structured response validation**

ก่อน append mapping:

* variant ต้องไม่ว่าง  
* uid ต้องมีอยู่จริง  
* confidence ต้องเป็นเลข  
* uid ต้องถูก canonicalize ผ่าน resolveUUID()

#### **3\. เพิ่ม review band**

* confidence \>= 90 → append ได้  
* 70-89 → ส่งเข้า AI review queue / pending mapping sheet  
* \<70 → ignore

#### **4\. บันทึก AI audit**

อย่างน้อยเก็บ:

* unknown name  
* candidate subset  
* chosen uid  
* confidence  
* timestamp  
* model

### **ผลลัพธ์**

AI จะเป็น “assistive workflow” ที่ตรวจสอบย้อนหลังได้ ไม่ใช่ black box. 

---

### **D3) Refactor** Service\_Search.gs **ให้ใช้ AI output ใหม่**

เมื่อ AI fields แยกแล้ว search ควร:

* ใช้ normalized\_name กับ search\_keywords แยกกัน  
* ให้น้ำหนักต่างกัน  
* ไม่ต้องอาศัย \[AI\] tag ใน field text เดิม. 

---

## **Phase E — UI / Diagnostics / DevEx**

---

### **E1) Refactor** Menu.gs

## **ฟังก์ชันที่ควรแก้**

* wrappers ทั้งหมด  
* runSystemHealthCheck()  
* showQualityReport\_UI(). 

### **สิ่งที่ควรทำ**

* ย้ายข้อความ confirm/review ไป constants หรือ helper กลาง  
* แยก UI wrappers ออกจาก business logic ชัดขึ้น  
* เพิ่มเมนู diagnostic ใหม่เช่น:  
  * “Dry Run: Sync GPS”  
  * “Dry Run: Finalize”  
  * “Show Mapping Conflicts”

### **เหตุผล**

ตอนนี้ menu ทำหน้าที่ทั้ง UI routing และมี logic ย่อยปนอยู่บ้าง. 

---

### **E2) Refactor** Test\_Diagnostic.gs

## **ฟังก์ชันที่ควรแก้**

* RUN\_SYSTEM\_DIAGNOSTIC()  
* RUN\_SHEET\_DIAGNOSTIC(). 

### **วิธีแก้**

* แยก diagnostics core เป็นฟังก์ชันคืน object:  
  * collectSystemDiagnostics\_()  
  * collectSheetDiagnostics\_()  
* แล้วให้ UI wrapper แค่เอา object ไป format alert  
* เพิ่ม diagnostics ใหม่:  
  * duplicate UUID  
  * inactive/merged leak in search  
  * stale mappings  
  * queue conflicts

### **ผลลัพธ์**

ใช้ test ได้ทั้งใน UI และ automation ภายหลัง. 

---

### **E3) Refactor** Test\_AI.gs

## **ฟังก์ชันที่ควรแก้**

* debugGeminiConnection()  
* debug\_ResetSelectedRowsAI()  
* forceRunAI\_Now(). 

### **วิธีแก้**

* ให้ debugGeminiConnection() return raw response object/log ด้วย ไม่ใช่แค่ alert  
* debug\_ResetSelectedRowsAI() ต้องปรับตาม field model ใหม่เมื่อแยก AI status ออกจาก normalized  
* เพิม test helper เช่น:  
  * testRetrieveCandidates\_()  
  * testAIResponseValidation\_()

---

### **E4) Refactor** WebApp.gs **\+** Index.html

## **ไฟล์**

* WebApp.gs  
* Index.html. 

## **สิ่งที่ควรทำ**

### **1\.** WebApp.gs

* เพิ่ม API action ที่ชัดกว่าเดิมถ้าจะมีหลาย action  
* แยก doPost() routing เป็น switch/handler map. 

### **2\.** Index.html

* เตรียมรองรับ metadata ใหม่จาก search result:  
  * verified badge  
  * coord source badge  
  * low confidence warning  
* ถ้าจะ refactor search ranking ควร render explanation เช่น “AI Match”, “Verified”, “Driver GPS”

### **ทำเมื่อไร**

หลัง Service\_Search.gs refactor เสร็จจึงค่อยอัปเดต UI เพราะจะได้ไม่แก้ซ้ำสองรอบ. 

---

# **4\) File-by-File Action Plan แบบสั้นแต่ actionable**

---

## Config.gs

### **ทำอะไร**

* เพิ่ม total column constants  
* เพิ่ม header arrays กลาง  
* เพิ่ม field groups / schema groups  
* เพิ่ม helper สำหรับ canonical widths. 

### **ทำก่อน/หลัง**

ก่อนสุด

---

## Service\_SchemaValidator.gs

### **ทำอะไร**

* อัปเดต schema Database ให้ครบ 22 cols  
* เช็ก U/V  
* เช็ก queue conflict rule  
* ขยาย Data required headers. 

### **ทำก่อน/หลัง**

พร้อมกับ Config.gs

---

## Service\_Master.gs

### **ทำอะไร**

* แก้ repairNameMapping\_Full()  
* แก้ finalizeAndClean\_MoveToMapping() ให้ใช้ full schema  
* แยก scoring helpers  
* แยก repository/helper functions  
* batch write จุดที่สำคัญ. 

### **ทำก่อน/หลัง**

ทันทีหลัง config/schema validator

---

## Service\_GPSFeedback.gs

### **ทำอะไร**

* เติม reject workflow  
* batch update  
* state model ที่ชัดเจน  
* stats อิง state ใหม่. 

### **ทำก่อน/หลัง**

หลัง Master phase แรก หรือทำคู่กันได้

---

## Utils\_Common.gs

### **ทำอะไร**

* เพิ่ม row model helpers  
* อาจแยก normalization แบบหลายระดับ  
* เก็บ utility scoring/parsing กลาง. 

### **ทำก่อน/หลัง**

หลัง Stabilize data integrity

---

## Service\_SCG.gs

### **ทำอะไร**

* ใช้ row objects  
* ใช้ resolveUUID()  
* ใช้ preCheck\_Apply() จริง  
* แยก transform helpers. 

### **ทำก่อน/หลัง**

หลัง row access layer พร้อม

---

## Service\_SoftDelete.gs

### **ทำอะไร**

* ทำให้ resolveUUID() ถูกเรียกจริงใน consumer flows  
* อาจเพิ่ม helper resolveCanonicalUUID\_ / isActiveUUID\_. 

### **ทำก่อน/หลัง**

ระหว่าง Phase B-C

---

## Service\_AutoPilot.gs

### **ทำอะไร**

* แยก AI keywords field ออกจาก normalized field  
* เพิ่ม response validation  
* เพิ่ม metadata/audit output. 

### **ทำก่อน/หลัง**

หลัง schema/index refactor

---

## Service\_Agent.gs

### **ทำะไร**

* retrieval ก่อน AI  
* candidate selection ใหม่  
* AI audit  
* canonical UUID enforcement. 

### **ทำก่อน/หลัง**

หลัง Service\_AutoPilot.gs design ใหม่

---

## Service\_Search.gs

### **ทำอะไร**

* read full schema  
* filter inactive/merged  
* ranking ใหม่  
* ใช้ separated search fields  
* canonical UUID. 

### **ทำก่อน/หลัง**

หลัง identity \+ AI field separation

---

## Menu.gs

### **ทำอะไร**

* ทำเมนู dry-run / conflict report / advanced diagnostics  
* keep UI wrapper บางลง. 

### **ทำก่อน/หลัง**

ช่วงท้าย

---

## Test\_Diagnostic.gs

### **ทำอะไร**

* แยก data-returning diagnostics helpers  
* เพิ่ม contract checks ใหม่. 

### **ทำก่อน/หลัง**

ควรเริ่มเร็วหลัง Phase A เพื่อช่วย refactor phase ต่อไป

---

## Test\_AI.gs

### **ทำอะไร**

* ปรับตาม AI field model ใหม่  
* เพิ่ม retrieval/validation tests. 

### **ทำก่อน/หลัง**

หลัง Phase D

---

## WebApp.gs **/** Index.html

### **ทำอะไร**

* ใช้ metadata ใหม่จาก search  
* แสดงผลคุณภาพ/สถานะมากขึ้น. 

### **ทำก่อน/หลัง**

ช่วงท้าย

---

# **5\) แผนแบบ Sprint ที่เอาไปใช้กับทีมได้**

## **Sprint 1 — Data Safety**

* Config.gs  
* Service\_SchemaValidator.gs  
* Service\_Master.gs  
* Service\_GPSFeedback.gs  
* Test\_Diagnostic.gs

Definition of Done

* DB schema validator รองรับ 22 cols  
* queue reject ทำงานจริง  
* mapping repair ไม่มี fallback UUID ผิด  
* diagnostics รันแล้วรายงาน invariants สำคัญได้. 

---

## **Sprint 2 — Schema Access Unification**

* Utils\_Common.gs  
* Service\_Master.gs  
* Service\_SCG.gs  
* Service\_Search.gs

Definition of Done

* ไม่มี hardcoded 17/22/29/9/5 ในจุดสำคัญ  
* row object helpers ถูกใช้ใน master/scg/search  
* search อ่าน full DB schema. 

---

## **Sprint 3 — Identity & Canonical UUID**

* Service\_SoftDelete.gs  
* Service\_SCG.gs  
* Service\_Search.gs  
* Service\_Agent.gs

Definition of Done

* UUID จาก mapping ถูก resolve เป็น canonical เสมอ  
* merged/inactive ไม่รั่วไป search/apply flow. 

---

## **Sprint 4 — AI Refactor**

* Service\_AutoPilot.gs  
* Service\_Agent.gs  
* Test\_AI.gs  
* Service\_Search.gs

Definition of Done

* retrieval-before-AI ใช้งาน  
* AI fields แยกจาก normalized  
* มี AI audit trail. 

---

## **Sprint 5 — UX / Search / Admin Tooling**

* Menu.gs  
* WebApp.gs  
* Index.html  
* Test\_Diagnostic.gs

Definition of Done

* search แสดง metadata quality  
* admin มี dry-run/conflict report  
* diagnostics ครบกว่าเดิม. 

---

# **6\) จุดที่ “ห้ามแก้พร้อมกัน” เพื่อไม่ให้เกิด regression หนัก**

## **อย่าแก้ 3 เรื่องนี้พร้อมกันใน commit เดียว**

1. schema/index refactor  
2. AI field separation  
3. search ranking overhaul

### **เหตุผล**

ทั้ง 3 เรื่องแตะ data model คนละชั้น แต่มีผลกระทบต่อกัน ถ้าทำพร้อมกันจะ trace regression ยากมาก. 

---

# **7\) ถ้าจะเริ่มลงมือจริง “ฟังก์ชันแรก” ที่ผมแนะนำให้แก้มี 5 ตัวนี้**

1. Service\_Master.gs::repairNameMapping\_Full()   
2. Service\_GPSFeedback.gs::applyApprovedFeedback()   
3. Service\_SchemaValidator.gs::SHEET\_SCHEMA.DATABASE \+ validateSheet\_()   
4. Service\_Master.gs::finalizeAndClean\_MoveToMapping()   
5. Service\_Search.gs::searchMasterData() 

นี่คือ 5 จุดที่ให้ผลคุ้มสุดต่อความเสี่ยงทั้งระบบครับ.

---

# **8\) สรุปสุดท้ายแบบ practical มากที่สุด**

ถ้าผมต้องสรุป Blueprint นี้ให้เป็นคำสั่งเดียวสำหรับทีม dev:

เริ่มจาก “ทำ schema/index ให้เป็นระบบเดียว \+ ทำ UUID/mapping ให้ canonical \+ ทำ queue review ให้ครบ state” ก่อน แล้วค่อยไปแตะ AI และ search ranking. 

นี่คือเส้นทางที่ปลอดภัยที่สุดและคุ้มที่สุดสำหรับ repo นี้ครับ.

---

## **Commands ที่ผมใช้อ้างอิง source สำหรับทำ blueprint นี้**

* rg \--files  
* nl \-ba Config.gs  
* nl \-ba Service\_Master.gs | sed \-n '1,240p'  
* nl \-ba Service\_Master.gs | sed \-n '241,420p'  
* nl \-ba Service\_Master.gs | sed \-n '421,760p'  
* nl \-ba Service\_GPSFeedback.gs  
* nl \-ba Service\_Agent.gs  
* nl \-ba Service\_AutoPilot.gs  
* nl \-ba Service\_Search.gs  
* nl \-ba Service\_SchemaValidator.gs  
* nl \-ba Service\_SoftDelete.gs  
* nl \-ba Service\_SCG.gs | sed \-n '1,220p'  
* nl \-ba Service\_SCG.gs | sed \-n '221,450p'  
* nl \-ba Utils\_Common.gs  
* nl \-ba Menu.gs | sed \-n '1,260p'  
* nl \-ba Test\_AI.gs  
* nl \-ba Test\_Diagnostic.gs

ถ้าคุณต้องการ ผมทำต่อให้ได้อีก 2 แบบที่ต่อจาก Blueprint นี้โดยตรง:

1. แตก Blueprint นี้เป็น TODO list รายไฟล์ รายฟังก์ชัน แบบพร้อมลง Jira/Trello  
2. จัดลำดับ commit plan ว่าควร commit/refactor เป็นกี่ PR และแต่ละ PR ควรมี scope อะไร

ผมต้องการข้อนี้"แตก Blueprint นี้เป็น TODO list รายไฟล์ รายฟังก์ชัน แบบพร้อมลง Jira/Trello"\\nแต่ผมไม่มี Jira/Trello จะทำยังไงดีครับ คุณใช้ GitHub Issues แทนได้มั้ยครับ  
แต่ผมไม่รู้ว่าจะต้องทำยยังไงต่อคับ ช่วยบอกผมหน่อยนะ  
คัดลอก

## **ได้ครับ ใช้ GitHub Issues แทน Jira/Trello ได้สบายมาก**

สำหรับ repo นี้ ผมแนะนำแบบง่ายที่สุดเลย:

* ใช้ GitHub Issues เป็นรายการงาน  
* ใช้ Labels แทนประเภทงาน  
* ใช้ Milestones แทน sprint/phase  
* ถ้าอยากได้ board แบบ Trello ให้ใช้ GitHub Projects เพิ่มทีหลังได้

คุณ ไม่จำเป็นต้องมี Jira/Trello ก็จัดการงานได้ครบครับ

---

# **1\) คุณควรทำยังไงต่อ แบบง่ายที่สุด**

## **ทางเลือกที่ง่ายสุด**

### **ใช้แค่ 3 อย่างนี้พอ**

1. Issues \= งานแต่ละชิ้น  
2. Labels \= ประเภทงาน  
3. Milestones \= กลุ่มงานตาม phase

สำหรับ repo นี้ ผมมองว่าเพียงเท่านี้ก็เริ่มทำงานได้แล้ว เพราะ backlog ที่เราคุยกันสามารถแตกเป็น issue รายไฟล์/รายฟังก์ชันได้ชัดเจนจาก source ที่มีอยู่แล้ว. 

---

# **2\) ถ้าคุณไม่รู้จะเริ่มกดตรงไหนบน GitHub ให้ทำตามนี้**

## **Step 1: เข้า repo ของคุณบน GitHub**

จากที่คุณให้ลิงก์ repo มาก่อนหน้านี้ ให้เปิดหน้า repo นั้น

## **Step 2: เปิดแท็บ** Issues

* ด้านบนของ repo จะมีแท็บ Issues  
* กดเข้าไป  
* กดปุ่ม New issue

## **Step 3: สร้าง Labels ก่อน**

ผมแนะนำ labels แบบนี้:

* phase:stabilize  
* phase:schema  
* phase:identity  
* phase:ai  
* phase:ui  
* type:bug  
* type:refactor  
* type:tech-debt  
* type:test  
* priority:P0  
* priority:P1  
* priority:P2  
* priority:P3

เหตุผลคือ backlog ของ repo นี้แบ่งชัดตาม phase และ priority ตามที่เราวิเคราะห์จาก Config.gs, Service\_Master.gs, Service\_GPSFeedback.gs, Service\_SchemaValidator.gs, Service\_Agent.gs, และ Service\_Search.gs. 

## **Step 4: สร้าง Milestones**

ตั้ง milestones แบบนี้:

* Phase A \- Stabilize Data Integrity  
* Phase B \- Normalize Schema Access  
* Phase C \- Identity & Mapping  
* Phase D \- AI Refactor  
* Phase E \- UI / Search / Diagnostics

นี่สอดคล้องกับ blueprint ที่แยกตามกลุ่มความเสี่ยงและลำดับ refactor. 

## **Step 5: สร้าง Issues จาก TODO list ด้านล่าง**

คุณสามารถ copy-paste ทีละ issue ได้เลย

---

# **3\) TODO List แบบพร้อมลง GitHub Issues**

ด้านล่างนี้คือ issue backlog พร้อมใช้  
ผมจะให้ในรูปแบบ:

* Title  
* Why  
* Files  
* Tasks  
* Acceptance Criteria  
* Labels  
* Milestone

---

## **ISSUE 1**

### **Title**

\[P0\] Fix NameMapping UUID fallback logic in repairNameMapping\_Full

### **Why**

repairNameMapping\_Full() มี logic fallback UUID ที่เสี่ยงใช้ค่าผิดเจตนา เพราะ lookup จาก r\[1\] ซึ่งเป็น UID field แทนที่จะ lookup จากชื่อ variant ทำให้ alias อาจชี้ UUID ผิดหรือเกิด UUID ใหม่โดยไม่ตั้งใจ. 

### **Files**

* Service\_Master.gs 

### **Tasks**

* ตรวจ logic ใน repairNameMapping\_Full()  
* แยก helper สำหรับ resolve UUID ของ mapping row  
* เปลี่ยน fallback ให้หา UUID จาก variant/master record อย่างถูกต้อง  
* หยุดการ generate UUID ใหม่แบบอัตโนมัติถ้ายังไม่แน่ใจว่าเป็น master ใหม่จริง  
* เพิ่มรายงาน rows ที่ invalid หรือ fix ไม่ได้

### **Acceptance Criteria**

* ไม่ใช้ uuidMap\[normalizeText(r\[1\])\] แบบเดิม  
* mapping ที่ไม่มี UID จะไม่สร้าง UUID ใหม่แบบเงียบ ๆ  
* repair run แล้วไม่ทำ alias หลุดจาก master record เดิม

### **Labels**

* priority:P0  
* type:bug  
* phase:stabilize

### **Milestone**

Phase A \- Stabilize Data Integrity

---

## **ISSUE 2**

### **Title**

\[P0\] Complete GPS Queue review workflow with explicit REJECTED state

### **Why**

applyApprovedFeedback() มี approve path ชัดเจน แต่ reject path ยังไม่สมบูรณ์ เพราะ reject ถูกติ๊กแล้วจะถูก skip โดยไม่ stamp สถานะ REJECTED กลับลง queue ทำให้ audit trail ไม่ครบ. 

### **Files**

* Service\_GPSFeedback.gs 

### **Tasks**

* เพิ่ม reject flow ใน applyApprovedFeedback()  
* ถ้า Reject=true && Approve=false ให้เขียน REJECTED  
* ถ้า Approve=true && Reject=true ให้แจ้ง error หรือ skip พร้อม log  
* ปรับ showGPSQueueStats() ให้นับจาก state จริง  
* พิจารณาเพิ่ม metadata review fields ในอนาคต

### **Acceptance Criteria**

* queue row ที่ reject แล้วต้องมีสถานะชัดเจน  
* stats ต้องนับ approved/rejected/pending ได้ถูกต้อง  
* ห้ามมี row ที่ approve+reject พร้อมกันโดยไม่ถูกจับ

### **Labels**

* priority:P0  
* type:bug  
* type:refactor  
* phase:stabilize

### **Milestone**

Phase A \- Stabilize Data Integrity

---

## **ISSUE 3**

### **Title**

\[P0\] Upgrade Database schema validation from 20 to 22 columns

### **Why**

Config.gs กำหนด Database ถึงคอลัมน์ 22 แต่ Service\_SchemaValidator.gs ยัง validate แค่ขั้นต่ำ 20 คอลัมน์ ทำให้คอลัมน์ Record\_Status และ Merged\_To\_UUID ยังไม่ถูกบังคับเชิง schema อย่างเต็มรูปแบบ. 

### **Files**

* Config.gs  
* Service\_SchemaValidator.gs 

### **Tasks**

* ปรับ DATABASE.minColumns เป็น 22  
* เพิ่ม required headers สำหรับ U/V  
* อัปเดต report ใน runFullSchemaValidation()  
* ทบทวน validateSystemIntegrity() ให้สอดคล้อง schema ใหม่

### **Acceptance Criteria**

* schema validator ต้อง fail ถ้า Database ไม่มี U/V  
* report ต้องแสดง header mismatch ของคอลัมน์ใหม่ได้

### **Labels**

* priority:P0  
* type:refactor  
* phase:stabilize  
* type:test

### **Milestone**

Phase A \- Stabilize Data Integrity

---

## **ISSUE 4**

### **Title**

\[P0\] Unify Database read/write width across master flows

### **Why**

หลาย flow ยังใช้ 17 คอลัมน์เดิม แม้ schema ปัจจุบันของ Database มี 22 คอลัมน์แล้ว ทำให้เกิด schema drift ระหว่างโลกเก่ากับโลกใหม่. 

### **Files**

* Config.gs  
* Service\_Master.gs  
* Service\_Search.gs 

### **Tasks**

* สร้าง constant สำหรับ total DB columns  
* เปลี่ยนจุดที่ใช้ 17 ให้ใช้ width กลางเดียว  
* ทบทวน finalizeAndClean\_MoveToMapping()  
* ทบทวน recalculateAllQuality()  
* ทบทวน searchMasterData()

### **Acceptance Criteria**

* ไม่มีจุดสำคัญที่อ่าน/เขียน Database แค่ 17 คอลัมน์โดยไม่ตั้งใจ  
* DB flows ทั้งหมดใช้ width เดียวกันอย่างสม่ำเสมอ

### **Labels**

* priority:P0  
* type:tech-debt  
* type:refactor  
* phase:schema

### **Milestone**

Phase B \- Normalize Schema Access

---

## **ISSUE 5**

### **Title**

\[P1\] Introduce shared row model helpers for Database, Data, Queue, and NameMapping

### **Why**

หลายไฟล์ยังใช้ raw array index เช่น r\[10\], r\[22\], r\[26\], r\[28\] ทำให้อ่านยากและเสี่ยงต่อ off-by-one bugs เมื่อ schema ขยับ. 

### **Files**

* Utils\_Common.gs  
* Config.gs  
* Service\_SCG.gs  
* Service\_Master.gs  
* Service\_Search.gs  
* Service\_GPSFeedback.gs 

### **Tasks**

* เพิ่ม helper แปลง row array → object  
* เพิ่ม helper แปลง object → row array  
* สร้าง constants/idx สำหรับ Data และ GPS\_Queue  
* ค่อย ๆ migrate flow ที่อ่านยากที่สุดก่อน

### **Acceptance Criteria**

* มี shared helper อย่างน้อยสำหรับ Database และ GPS\_Queue  
* ฟังก์ชันสำคัญ 2-3 ตัวแรกย้ายไปใช้ row object แล้ว

### **Labels**

* priority:P1  
* type:refactor  
* type:tech-debt  
* phase:schema

### **Milestone**

Phase B \- Normalize Schema Access

---

## **ISSUE 6**

### **Title**

\[P1\] Refactor finalizeAndClean\_MoveToMapping to full-schema safe workflow

### **Why**

finalizeAndClean\_MoveToMapping() ยังทำงานบน 17 คอลัมน์หลัก และมี logic ที่ควรแยกเป็นขั้นตอนเพื่อให้ตรวจสอบ conflict และ rollback ได้ง่ายขึ้น. 

### **Files**

* Service\_Master.gs 

### **Tasks**

* แยก logic เป็น collect/build/rewrite steps  
* ใช้ DB width เต็มตาม schema ใหม่  
* เพิ่ม conflict checks:  
  * duplicate UUID  
  * missing target UUID  
  * invalid suggested mapping  
* เคลียร์ search cache หลัง finalize

### **Acceptance Criteria**

* finalize flow อ่านง่ายขึ้น  
* มี conflict reporting  
* ไม่ทิ้ง metadata columns ของ v4.1

### **Labels**

* priority:P1  
* type:refactor  
* phase:schema  
* phase:identity

### **Milestone**

Phase B \- Normalize Schema Access

---

## **ISSUE 7**

### **Title**

\[P1\] Make all UUID consumers resolve canonical UUID via resolveUUID()

### **Why**

ระบบมี resolveUUID() แล้ว แต่ consumer flows เช่น search, apply coordinates, และ AI mapping ยังไม่ได้ใช้ canonical UUID อย่างทั่วถึง ทำให้ merged UUID อาจยังรั่วไปใช้งานต่อ. 

### **Files**

* Service\_SoftDelete.gs  
* Service\_SCG.gs  
* Service\_Search.gs  
* Service\_Agent.gs 

### **Tasks**

* define canonical UUID usage rule  
* apply resolveUUID() ในทุกจุดที่อ่าน UUID จาก mapping  
* exclude merged/inactive record จาก search/apply  
* ตรวจผลกระทบกับ AI write-back

### **Acceptance Criteria**

* alias/mapping/search/apply flows ใช้ canonical UUID เสมอ  
* merged UUID ไม่ถูกแสดงหรือใช้งานเป็นปลายทางหลัก

### **Labels**

* priority:P1  
* type:refactor  
* phase:identity

### **Milestone**

Phase C \- Identity & Mapping

---

## **ISSUE 8**

### **Title**

\[P1\] Add retrieval step before AI matching in resolveUnknownNamesWithAI

### **Why**

AI tier-4 ตอนนี้ใช้ slice(0, 500\) จาก database ตรง ๆ แทน retrieval-by-relevance ทำให้ผลลัพธ์ขึ้นกับลำดับแถวในชีตมากเกินไป. 

### **Files**

* Service\_Agent.gs  
* Service\_Search.gs  
* Utils\_Common.gs 

### **Tasks**

* สร้าง candidate retrieval helper  
* ใช้ normalize/token overlap/alias text คัด top-N  
* ส่งเฉพาะ candidate ที่เกี่ยวข้องเข้า AI  
* เปรียบเทียบจำนวน match ก่อน/หลัง refactor

### **Acceptance Criteria**

* เลิกใช้ slice(0, 500\) แบบตัดตรง  
* AI prompt ได้ candidate list ที่เกี่ยวข้องกว่าเดิม  
* โค้ดแยก retrieval กับ AI decision ชัดเจน

### **Labels**

* priority:P1  
* type:refactor  
* phase:ai

### **Milestone**

Phase D \- AI Refactor

---

## **ISSUE 9**

### **Title**

\[P1\] Separate NORMALIZED field from AI-generated search keywords

### **Why**

ตอนนี้ COL\_NORMALIZED ถูกใช้ทั้งเป็น normalized text, AI keyword bag, และสถานะ \[AI\] tag ทำให้ field เดียวแบกหลายหน้าที่เกินไปและ debug/search ranking ยาก. 

### **Files**

* Config.gs  
* Service\_AutoPilot.gs  
* Service\_Search.gs  
* Test\_AI.gs 

### **Tasks**

* ออกแบบ field model ใหม่:  
  * normalized\_name  
  * search\_keywords  
  * ai\_status  
  * ai\_updated\_at  
* migrate autopilot indexing  
* migrate search consumer  
* migrate debug reset function

### **Acceptance Criteria**

* ไม่มีการใช้ \[AI\] tag ปนใน normalized field เดิม  
* search อ่าน AI keywords จาก field ที่แยกชัด  
* debug/reset AI ทำงานกับ field ใหม่ได้

### **Labels**

* priority:P1  
* type:refactor  
* phase:ai  
* type:tech-debt

### **Milestone**

Phase D \- AI Refactor

---

## **ISSUE 10**

### **Title**

\[P2\] Improve search ranking using verified, quality, coord source, and canonical status

### **Why**

search ปัจจุบันให้คะแนนแบบหยาบมากและยังไม่ใช้ metadata ใหม่ของ v4.1 เช่น Verified, Quality, Coord\_Source, Record\_Status. 

### **Files**

* Service\_Search.gs  
* Service\_SoftDelete.gs  
* Index.html 

### **Tasks**

* สร้าง scoring function แยกต่างหาก  
* เพิ่ม weighted score ตาม metadata  
* exclude inactive/merged rows  
* ถ้าพร้อม ให้ UI แสดง badge metadata

### **Acceptance Criteria**

* search result คุณภาพสูงขึ้น  
* inactive/merged records ไม่ขึ้นเป็นผลหลัก  
* score logic อ่านง่ายและปรับแต่งได้

### **Labels**

* priority:P2  
* type:refactor  
* phase:ui  
* phase:ai

### **Milestone**

Phase E \- UI / Search / Diagnostics

---

## **ISSUE 11**

### **Title**

\[P2\] Batch-write master and queue updates to reduce partial-write risk

### **Why**

บาง flow ยังอัปเดตชีตทีละ cell ใน loop ซึ่งเสี่ยงช้าและ partial write เมื่อข้อมูลเริ่มโต. 

### **Files**

* Service\_Master.gs  
* Service\_GPSFeedback.gs 

### **Tasks**

* เปลี่ยน update timestamp ใน master เป็น batch write  
* เปลี่ยน apply approved feedback เป็น mutate in-memory แล้ว setValues กลับ  
* วัดความต่างด้าน performance

### **Acceptance Criteria**

* ไม่มี cell-by-cell write loop ใน 2 flow นี้  
* flow ยังได้ผลลัพธ์เท่าเดิมแต่เร็ว/เสถียรกว่าเดิม

### **Labels**

* priority:P2  
* type:refactor  
* type:tech-debt

### **Milestone**

Phase E \- UI / Search / Diagnostics

---

## **ISSUE 12**

### **Title**

\[P2\] Refactor system diagnostics into reusable non-UI test helpers

### **Why**

diagnostic ปัจจุบันเน้น alert()/toast() มากกว่า structured result ทำให้ reuse สำหรับ regression checks หรือ automation ยาก. 

### **Files**

* Test\_Diagnostic.gs  
* Menu.gs  
* Service\_SchemaValidator.gs 

### **Tasks**

* แยก collectSystemDiagnostics\_()  
* แยก collectSheetDiagnostics\_()  
* ให้ UI wrapper เป็นชั้นแสดงผลเท่านั้น  
* เพิ่ม checks เช่น:  
  * duplicate UUID  
  * canonical UUID leak  
  * queue conflicts  
  * schema coverage 22 cols

### **Acceptance Criteria**

* diagnostic core คืน object ได้  
* UI ยังแสดงผลได้เหมือนเดิม  
* เพิ่ม regression checks สำคัญได้ง่าย

### **Labels**

* priority:P2  
* type:test  
* type:refactor

### **Milestone**

Phase E \- UI / Search / Diagnostics

---

# **4\) ถ้าคุณอยากใช้งานแบบ “ง่ายมาก” โดยไม่ต้องสร้างเยอะทีเดียว**

## **วิธีเริ่มที่ง่ายที่สุด**

สร้างแค่ 5 issues ก่อนนี้ก็พอ:

1. Fix NameMapping UUID fallback  
2. Complete GPS Queue reject workflow  
3. Upgrade Database schema validation to 22 columns  
4. Unify Database read/write width  
5. Add retrieval step before AI matching

แค่นี้ก็เพียงพอให้คุณเริ่มเดินงานจริงได้แล้ว เพราะครอบคลุม:

* data safety  
* schema safety  
* AI quality. 

---

# **5\) แล้วถ้าคุณ “ไม่รู้ว่าจะทำ issue ยังไง” ผมแนะนำ workflow นี้**

## **แบบง่ายสุด**

### **ขั้นตอน**

1. สร้าง issue ตามรายการด้านบน  
2. ตั้ง label ให้ครบ  
3. เริ่มจาก issue ที่เป็น priority:P0  
4. ทำทีละ issue  
5. เวลาแก้ code ให้ใส่ใน commit message ว่าเกี่ยวกับ issue ไหน เช่น:  
   * fix: complete GPS queue rejected state (\#2)  
   * refactor: unify database width constants (\#4)

### **เวลา issue เสร็จ**

* คอมเมนต์ใน issue ว่า:  
  * ทำอะไรไปแล้ว  
  * แก้ไฟล์ไหน  
  * ทดสอบอะไรบ้าง  
* แล้วค่อยปิด issue

---

# **6\) ถ้าคุณไม่มี board แบบ Trello แล้วอยากเห็นงานเป็นคอลัมน์**

ใช้ GitHub Projects ได้ครับ  
แต่มันไม่จำเป็นในวันแรก

ถ้าจะใช้ ให้ตั้งคอลัมน์แบบนี้:

* Backlog  
* Ready  
* In Progress  
* Review  
* Done

แล้วลาก issues เข้าแต่ละคอลัมน์ได้เลย

แต่ถ้าคุณเพิ่งเริ่ม ผมแนะนำ:

เริ่มด้วย Issues \+ Labels \+ Milestones ก่อน  
พอคุ้นแล้วค่อยเพิ่ม Projects ทีหลัง

---

# **7\) ผมแนะนำให้คุณทำต่อแบบไหนดีที่สุด**

## **ถ้าคุณทำคนเดียว**

* ใช้ Issues อย่างเดียวพอ  
* สร้าง 5 issue แรก  
* ทำทีละ issue  
* ไม่ต้องรีบสร้างทุก issue พร้อมกัน

## **ถ้าคุณมีทีม 2-5 คน**

* ใช้ Issues \+ Labels \+ Milestones  
* สร้างทุก issue ที่ผมเขียนให้  
* ใช้ assignee กำหนดคนรับผิดชอบ  
* ถ้าต้องการค่อยเปิด GitHub Projects เพิ่ม

---

# **8\) ผมช่วยคุณต่อได้ยังไงบ้างตอนนี้**

ผมช่วยได้ 3 แบบ:

### **แบบ A — ผมแปลง issue ทั้งหมดเป็น Markdown พร้อม copy-paste ลง GitHub Issues**

เช่น issue ละบล็อกสวย ๆ พร้อมหัวข้อ/checkbox

### **แบบ B — ผมทำเป็น GitHub Issue Template**

เช่น:

* bug template  
* refactor template  
* tech-debt template

### **แบบ C — ผมช่วยคุณเลือกก่อนว่า**

“5 issue แรกที่ควรเปิดทันที”  
พร้อมข้อความ copy-paste แบบสั้นที่สุด

---

# **9\) คำตอบตรง ๆ สำหรับคำถามคุณ**

## **ใช้ GitHub Issues แทน Jira/Trello ได้ไหม?**

ได้ครับ และเหมาะมากด้วย สำหรับ repo นี้

## **ถ้าคุณไม่รู้ต้องทำยังไงต่อ?**

ให้เริ่มแบบนี้:

1. เปิด GitHub repo  
2. เข้า Issues  
3. สร้าง labels  
4. สร้าง milestone Phase A \- Stabilize Data Integrity  
5. เปิด 3 issue แรกนี้ก่อน:  
   * Fix NameMapping UUID fallback  
   * Complete GPS Queue reject workflow  
   * Upgrade Database schema validation to 22 columns

นี่คือจุดเริ่มที่ดีที่สุดครับ เพราะมันแก้ risk หลักของระบบจาก mapping, queue governance, และ schema consistency โดยตรง. 

---

## **Commands ที่ผมใช้อ้างอิง source สำหรับแตก backlog นี้**

* rg \--files  
* nl \-ba Config.gs  
* nl \-ba Service\_Master.gs | sed \-n '421,760p'  
* nl \-ba Service\_GPSFeedback.gs  
* nl \-ba Service\_Agent.gs  
* nl \-ba Service\_AutoPilot.gs  
* nl \-ba Service\_Search.gs  
* nl \-ba Service\_SchemaValidator.gs  
* nl \-ba Service\_SoftDelete.gs  
* nl \-ba Service\_SCG.gs | sed \-n '1,220p'  
* nl \-ba Utils\_Common.gs  
* nl \-ba Menu.gs | sed \-n '1,260p'  
* nl \-ba Test\_AI.gs  
* nl \-ba Test\_Diagnostic.gs

# Markdown GitHub Issues

## **ด้านล่างคือ Markdown พร้อม copy-paste ลง GitHub Issues ได้ทันที**

ผมจัดให้เป็น issue ละ 1 บล็อก โดยคุณสามารถ:

1. เปิด GitHub repo  
2. เข้าแท็บ Issues  
3. กด New issue  
4. copy เนื้อหาจากแต่ละบล็อกไปวางได้เลย

รายการ issue เหล่านี้อิงจาก risk review และ blueprint ที่แตกจากโค้ดจริงใน Config.gs, Service\_Master.gs, Service\_GPSFeedback.gs, Service\_SchemaValidator.gs, Service\_SoftDelete.gs, Service\_Agent.gs, Service\_AutoPilot.gs, Service\_Search.gs, Service\_SCG.gs, Utils\_Common.gs, Menu.gs, Test\_AI.gs, และ Test\_Diagnostic.gs. 

---

# **แนะนำก่อนเริ่ม**

## **Labels ที่ควรสร้างก่อน**

คัดลอกไปใช้เป็นรายการได้เลย:

* priority:P0  
* priority:P1  
* priority:P2  
* priority:P3  
* type:bug  
* type:refactor  
* type:tech-debt  
* type:test  
* phase:stabilize  
* phase:schema  
* phase:identity  
* phase:ai  
* phase:ui

## **Milestones ที่ควรสร้างก่อน**

* Phase A \- Stabilize Data Integrity  
* Phase B \- Normalize Schema Access  
* Phase C \- Identity & Mapping  
* Phase D \- AI Refactor  
* Phase E \- UI / Search / Diagnostics

---

---

## **ISSUE 1**

### **Title**

\[P0\] Fix NameMapping UUID fallback logic in repairNameMapping\_Full

### **Body**

\#\# Summary

Fix incorrect UUID fallback behavior in \`repairNameMapping\_Full()\` to prevent alias rows from being linked to the wrong UUID or silently generating new UUIDs.

\#\# Problem

\`repairNameMapping\_Full()\` currently has fallback logic that can use the wrong lookup key when \`uid\` is missing. This creates a risk that NameMapping rows may become detached from the real master record.

\#\# Why this matters

This is an identity integrity issue:

\- alias \-\> UUID relationships may become invalid

\- search may point to the wrong master record

\- coordinate matching may use the wrong destination

\- future merge/soft-delete flows become less reliable

\#\# Affected files

\- \`Service\_Master.gs\`

\#\# Relevant functions

\- \`repairNameMapping\_Full()\`

\#\# Tasks

\- \[ \] Review current fallback UUID logic in \`repairNameMapping\_Full()\`

\- \[ \] Replace incorrect fallback lookup with logic based on variant/master name resolution

\- \[ \] Prevent silent UUID generation unless a new master record is intentionally being created

\- \[ \] Add reporting for invalid or unresolved mapping rows

\- \[ \] Add a dry-run-friendly internal helper if needed

\#\# Acceptance Criteria

\- \[ \] Missing mapping UUIDs are resolved from the correct source

\- \[ \] The function no longer silently creates incorrect new UUIDs

\- \[ \] Repair output clearly reports unresolved rows

\- \[ \] Existing valid mappings remain stable after repair

\#\# Suggested Labels

\- \`priority:P0\`

\- \`type:bug\`

\- \`phase:stabilize\`

\#\# Suggested Milestone

\`Phase A \- Stabilize Data Integrity\`

---

## **ISSUE 2**

### **Title**

\[P0\] Complete GPS Queue review workflow with explicit REJECTED state

### **Body**

\#\# Summary

Complete the GPS review workflow so rejected queue rows are explicitly marked as \`REJECTED\` instead of being silently skipped.

\#\# Problem

\`applyApprovedFeedback()\` handles approved rows but does not fully process rejected rows. If a row is checked as rejected, it is currently skipped without writing a clear final state.

\#\# Why this matters

This causes review/audit ambiguity:

\- admins cannot clearly see which rows were rejected

\- queue stats become less reliable

\- rejected rows may look pending or partially processed

\#\# Affected files

\- \`Service\_GPSFeedback.gs\`

\#\# Relevant functions

\- \`applyApprovedFeedback()\`

\- \`showGPSQueueStats()\`

\- \`upgradeGPSQueueSheet()\`

\#\# Tasks

\- \[ \] Add explicit reject flow in \`applyApprovedFeedback()\`

\- \[ \] If \`Reject=true\` and \`Approve=false\`, write \`REJECTED\`

\- \[ \] If \`Approve=true\` and \`Reject=true\`, fail validation or skip with clear log/message

\- \[ \] Update queue stats to rely on explicit state when possible

\- \[ \] Consider cleanup behavior for checkbox values after approval/rejection

\#\# Acceptance Criteria

\- \[ \] Rejected rows are clearly marked \`REJECTED\`

\- \[ \] Approved rows are clearly marked \`APPROVED\`

\- \[ \] Conflicting approve/reject states are detected

\- \[ \] Stats correctly count approved/rejected/pending rows

\#\# Suggested Labels

\- \`priority:P0\`

\- \`type:bug\`

\- \`type:refactor\`

\- \`phase:stabilize\`

\#\# Suggested Milestone

\`Phase A \- Stabilize Data Integrity\`

---

## **ISSUE 3**

### **Title**

\[P0\] Upgrade Database schema validation from 20 to 22 columns

### **Body**

\#\# Summary

Update schema validation so \`Database\` is validated against the full v4.1 schema (22 columns), including soft delete columns.

\#\# Problem

The config defines \`Database\` fields through column 22, but schema validation still treats the database as requiring only 20 columns.

\#\# Why this matters

This creates a schema drift risk:

\- v4.1 status fields may be missing without validation failing

\- flows can appear healthy while part of the schema is absent

\- future refactors become more fragile

\#\# Affected files

\- \`Config.gs\`

\- \`Service\_SchemaValidator.gs\`

\#\# Relevant functions / sections

\- \`CONFIG\`

\- \`SHEET\_SCHEMA.DATABASE\`

\- \`validateSheet\_()\`

\- \`runFullSchemaValidation()\`

\#\# Tasks

\- \[ \] Change \`DATABASE.minColumns\` from 20 to 22

\- \[ \] Add required header validation for \`Record\_Status\`

\- \[ \] Add required header validation for \`Merged\_To\_UUID\`

\- \[ \] Review whether \`validateSystemIntegrity()\` should also include stronger DB schema checks

\- \[ \] Update any related diagnostic messages

\#\# Acceptance Criteria

\- \[ \] Validation fails when DB columns U/V are missing

\- \[ \] Full schema report includes new columns

\- \[ \] Health/admin diagnostics reflect full v4.1 expectations

\#\# Suggested Labels

\- \`priority:P0\`

\- \`type:refactor\`

\- \`type:test\`

\- \`phase:stabilize\`

\#\# Suggested Milestone

\`Phase A \- Stabilize Data Integrity\`

---

## **ISSUE 4**

### **Title**

\[P0\] Unify Database read/write width across master flows

### **Body**

\#\# Summary

Unify all major \`Database\` read/write operations so they use one consistent schema width instead of mixing legacy 17-column behavior with v4.1 22-column behavior.

\#\# Problem

Some flows still read/write only 17 columns, while the current database model has 22 columns.

\#\# Why this matters

This is a major structural risk:

\- newer metadata may be ignored

\- schema changes are harder to reason about

\- data integrity depends on which function touched a row last

\#\# Affected files

\- \`Config.gs\`

\- \`Service\_Master.gs\`

\- \`Service\_Search.gs\`

\#\# Relevant functions

\- \`finalizeAndClean\_MoveToMapping()\`

\- \`recalculateAllQuality()\`

\- \`searchMasterData()\`

\#\# Tasks

\- \[ \] Introduce a single DB width constant

\- \[ \] Replace raw \`17\` reads/writes in critical flows

\- \[ \] Update finalize flow to preserve full row shape

\- \[ \] Update search flow to use full DB width where appropriate

\- \[ \] Review quality/confidence functions for width consistency

\#\# Acceptance Criteria

\- \[ \] Critical DB flows no longer use legacy width by accident

\- \[ \] DB read/write width is centrally defined

\- \[ \] Full schema metadata survives key workflows

\#\# Suggested Labels

\- \`priority:P0\`

\- \`type:refactor\`

\- \`type:tech-debt\`

\- \`phase:schema\`

\#\# Suggested Milestone

\`Phase B \- Normalize Schema Access\`

---

## **ISSUE 5**

### **Title**

\[P1\] Introduce shared row model helpers for Database, Data, Queue, and NameMapping

### **Body**

\#\# Summary

Create shared helpers to convert raw sheet rows into structured objects and back, reducing index-based bugs and making future schema changes safer.

\#\# Problem

Many flows still use raw array indexes directly (for example \`r\[10\]\`, \`r\[22\]\`, \`r\[26\]\`, \`r\[28\]\`), which makes the code hard to read and fragile.

\#\# Why this matters

This is a maintainability and correctness issue:

\- off-by-one bugs are easy to introduce

\- schema changes require touching too many call sites

\- business logic is hard to understand during debugging

\#\# Affected files

\- \`Utils\_Common.gs\`

\- \`Config.gs\`

\- \`Service\_Master.gs\`

\- \`Service\_SCG.gs\`

\- \`Service\_Search.gs\`

\- \`Service\_GPSFeedback.gs\`

\#\# Tasks

\- \[ \] Add helper(s) for \`Database\` row \<-\> object conversion

\- \[ \] Add helper(s) for \`NameMapping\` row \<-\> object conversion

\- \[ \] Add helper(s) for \`Data\` row \<-\> object conversion

\- \[ \] Add helper(s) for \`GPS\_Queue\` row \<-\> object conversion

\- \[ \] Migrate at least 2 critical flows to the new helpers first

\#\# Acceptance Criteria

\- \[ \] Shared row helpers exist for core sheets

\- \[ \] At least one master flow and one queue/search/data flow use row objects

\- \[ \] New code avoids raw position-based access where practical

\#\# Suggested Labels

\- \`priority:P1\`

\- \`type:refactor\`

\- \`type:tech-debt\`

\- \`phase:schema\`

\#\# Suggested Milestone

\`Phase B \- Normalize Schema Access\`

---

## **ISSUE 6**

### **Title**

\[P1\] Refactor finalizeAndClean\_MoveToMapping to full-schema safe workflow

### **Body**

\#\# Summary

Refactor \`finalizeAndClean\_MoveToMapping()\` into a clearer multi-step workflow that preserves full schema safety and reports mapping conflicts more explicitly.

\#\# Problem

The finalize flow is still anchored to a narrower legacy shape and bundles too much logic into one function.

\#\# Why this matters

Finalize is one of the highest-risk workflows:

\- it rewrites master data

\- it creates mappings

\- it deletes/cleans rows

\- it can create long-term identity problems if something goes wrong

\#\# Affected files

\- \`Service\_Master.gs\`

\#\# Relevant functions

\- \`finalizeAndClean\_MoveToMapping()\`

\#\# Tasks

\- \[ \] Split finalize flow into collection/build/write steps

\- \[ \] Preserve full DB schema width

\- \[ \] Add conflict checks for suggested-name target resolution

\- \[ \] Add validation for missing/duplicate UUID cases

\- \[ \] Clear related caches after finalize if needed

\#\# Acceptance Criteria

\- \[ \] Finalize flow is separated into smaller internal helpers

\- \[ \] Finalize preserves v4.1 metadata columns

\- \[ \] Mapping conflicts are reported clearly

\- \[ \] Finalize remains recoverable/auditable through backup sheet behavior

\#\# Suggested Labels

\- \`priority:P1\`

\- \`type:refactor\`

\- \`phase:schema\`

\- \`phase:identity\`

\#\# Suggested Milestone

\`Phase B \- Normalize Schema Access\`

---

## **ISSUE 7**

### **Title**

\[P1\] Make all UUID consumers resolve canonical UUID via resolveUUID()

### **Body**

\#\# Summary

Ensure every consumer that uses a UUID from NameMapping or related flows resolves it through \`resolveUUID()\` before use.

\#\# Problem

The system already supports merged UUID chains, but not all consumer flows appear to resolve to the canonical UUID before acting on it.

\#\# Why this matters

This can cause stale or merged UUIDs to continue leaking into:

\- search results

\- coordinate matching

\- AI write-back

\- alias resolution

\#\# Affected files

\- \`Service\_SoftDelete.gs\`

\- \`Service\_SCG.gs\`

\- \`Service\_Search.gs\`

\- \`Service\_Agent.gs\`

\#\# Relevant functions

\- \`resolveUUID()\`

\- \`applyMasterCoordinatesToDailyJob()\`

\- \`searchMasterData()\`

\- \`resolveUnknownNamesWithAI()\`

\#\# Tasks

\- \[ \] Define canonical UUID usage rule for the codebase

\- \[ \] Use \`resolveUUID()\` wherever UUIDs are read from mapping-like sources

\- \[ \] Exclude or redirect merged/inactive records in consumer flows

\- \[ \] Review AI write-back to ensure stored UUIDs are canonical

\#\# Acceptance Criteria

\- \[ \] Search uses canonical UUIDs

\- \[ \] Coordinate application uses canonical UUIDs

\- \[ \] AI mapping writes canonical UUIDs

\- \[ \] Merged UUIDs no longer act like active primary records

\#\# Suggested Labels

\- \`priority:P1\`

\- \`type:refactor\`

\- \`phase:identity\`

\#\# Suggested Milestone

\`Phase C \- Identity & Mapping\`

---

## **ISSUE 8**

### **Title**

\[P1\] Add retrieval step before AI matching in resolveUnknownNamesWithAI

### **Body**

\#\# Summary

Introduce candidate retrieval before sending unknown names to Gemini, instead of slicing the first 500 database rows.

\#\# Problem

\`resolveUnknownNamesWithAI()\` currently uses a fixed subset of master options rather than selecting the most relevant candidates first.

\#\# Why this matters

This weakens AI accuracy:

\- results depend on row order in the sheet

\- relevant candidates may be excluded

\- poor retrieval makes the AI look worse than it is

\#\# Affected files

\- \`Service\_Agent.gs\`

\- \`Utils\_Common.gs\`

\- \`Service\_Search.gs\` (if helper reuse makes sense)

\#\# Relevant functions

\- \`resolveUnknownNamesWithAI()\`

\#\# Tasks

\- \[ \] Add retrieval helper to select relevant candidates per unknown name

\- \[ \] Use normalized text / alias overlap / token similarity before AI call

\- \[ \] Replace first-500 slicing approach

\- \[ \] Keep prompt payload smaller and more relevant

\- \[ \] Add comparison notes for before/after behavior if practical

\#\# Acceptance Criteria

\- \[ \] Candidate subset is relevance-based, not row-order-based

\- \[ \] AI receives smaller but better candidate lists

\- \[ \] Retrieval and AI decision logic are separated

\#\# Suggested Labels

\- \`priority:P1\`

\- \`type:refactor\`

\- \`phase:ai\`

\#\# Suggested Milestone

\`Phase D \- AI Refactor\`

---

## **ISSUE 9**

### **Title**

\[P1\] Separate NORMALIZED field from AI-generated search keywords

### **Body**

\#\# Summary

Split deterministic normalization from AI-generated keyword enrichment so one field does not serve multiple unrelated roles.

\#\# Problem

The current normalized field mixes:

\- normalized base text

\- AI-generated search terms

\- AI status marker/tagging

\#\# Why this matters

This makes the system harder to debug and evolve:

\- search relevance tuning becomes harder

\- AI reset/debug functions become brittle

\- field semantics are unclear

\#\# Affected files

\- \`Config.gs\`

\- \`Service\_AutoPilot.gs\`

\- \`Service\_Search.gs\`

\- \`Test\_AI.gs\`

\#\# Relevant functions

\- \`processAIIndexing\_Batch()\`

\- \`searchMasterData()\`

\- \`debug\_ResetSelectedRowsAI()\`

\#\# Tasks

\- \[ \] Define new field model for normalized/search/AI state

\- \[ \] Update autopilot indexing to write to the new field(s)

\- \[ \] Update search to read the correct field(s)

\- \[ \] Update debug/reset helper logic

\- \[ \] Review whether schema changes are needed

\#\# Acceptance Criteria

\- \[ \] Normalized base text is no longer overloaded with AI tags/keywords

\- \[ \] Search reads explicit search-keyword field(s)

\- \[ \] AI reset/debug logic works with the new model

\#\# Suggested Labels

\- \`priority:P1\`

\- \`type:refactor\`

\- \`type:tech-debt\`

\- \`phase:ai\`

\#\# Suggested Milestone

\`Phase D \- AI Refactor\`

---

## **ISSUE 10**

### **Title**

\[P2\] Improve search ranking using verified, quality, coord source, and canonical status

### **Body**

\#\# Summary

Upgrade search ranking so it uses richer metadata instead of relying mainly on simple token match \+ basic score.

\#\# Problem

Current search scoring is too simple and does not make use of useful metadata already present in the system.

\#\# Why this matters

Users may not get the best result first:

\- active verified records may not rank above weaker matches

\- merged/inactive records may still leak through

\- driver-confirmed GPS records are not prioritized

\#\# Affected files

\- \`Service\_Search.gs\`

\- \`Service\_SoftDelete.gs\`

\- \`Index.html\` (optional UI metadata display)

\#\# Relevant functions

\- \`searchMasterData()\`

\- \`resolveUUID()\`

\#\# Tasks

\- \[ \] Add weighted ranking function

\- \[ \] Include exact/alias/normalized match signals

\- \[ \] Include verified/quality/confidence/coord source signals

\- \[ \] Exclude or de-prioritize inactive/merged records

\- \[ \] Optionally expose ranking-related metadata in UI

\#\# Acceptance Criteria

\- \[ \] Search ranking uses meaningful metadata

\- \[ \] Inactive/merged records are suppressed or strongly de-prioritized

\- \[ \] Best operational record appears earlier in typical searches

\#\# Suggested Labels

\- \`priority:P2\`

\- \`type:refactor\`

\- \`phase:ui\`

\- \`phase:ai\`

\#\# Suggested Milestone

\`Phase E \- UI / Search / Diagnostics\`

---

## **ISSUE 11**

### **Title**

\[P2\] Batch-write master and queue updates to reduce partial-write risk

### **Body**

\#\# Summary

Refactor cell-by-cell writes in major flows into batch updates to improve performance and reduce partial-write risk.

\#\# Problem

Some workflows still update rows/cells inside loops, which is slower and more fragile than batch operations.

\#\# Why this matters

This affects:

\- performance on larger sheets

\- quota efficiency

\- consistency if a run fails mid-loop

\#\# Affected files

\- \`Service\_Master.gs\`

\- \`Service\_GPSFeedback.gs\`

\#\# Relevant functions

\- \`syncNewDataToMaster()\`

\- \`applyApprovedFeedback()\`

\#\# Tasks

\- \[ \] Identify row-level updates currently written cell-by-cell

\- \[ \] Load target rows into memory

\- \[ \] Apply mutations in memory

\- \[ \] Write back in batches

\- \[ \] Validate behavior remains equivalent after refactor

\#\# Acceptance Criteria

\- \[ \] Critical flows avoid repeated cell-by-cell writes

\- \[ \] Functional output remains unchanged

\- \[ \] Large-run performance is improved or at least safer

\#\# Suggested Labels

\- \`priority:P2\`

\- \`type:refactor\`

\- \`type:tech-debt\`

\#\# Suggested Milestone

\`Phase E \- UI / Search / Diagnostics\`

---

## **ISSUE 12**

### **Title**

\[P2\] Refactor system diagnostics into reusable non-UI test helpers

### **Body**

\#\# Summary

Refactor diagnostics so the core logic returns structured data and UI wrappers only handle presentation.

\#\# Problem

Current diagnostics are tightly coupled to alerts/toasts, which makes reuse, automation, and regression checks harder.

\#\# Why this matters

Better diagnostics will support future refactors:

\- schema checks can be reused

\- regression checks become easier

\- automation becomes possible later

\#\# Affected files

\- \`Test\_Diagnostic.gs\`

\- \`Menu.gs\`

\- \`Service\_SchemaValidator.gs\`

\#\# Relevant functions

\- \`RUN\_SYSTEM\_DIAGNOSTIC()\`

\- \`RUN\_SHEET\_DIAGNOSTIC()\`

\#\# Tasks

\- \[ \] Extract diagnostics collectors that return objects/results

\- \[ \] Keep UI wrapper functions for alerts

\- \[ \] Add checks for duplicate UUIDs and canonical UUID leaks

\- \[ \] Add checks for queue state conflicts

\- \[ \] Reuse schema validator output where possible

\#\# Acceptance Criteria

\- \[ \] Diagnostic core logic can run without direct UI dependency

\- \[ \] UI output still works for spreadsheet users

\- \[ \] Additional regression checks are easy to add

\#\# Suggested Labels

\- \`priority:P2\`

\- \`type:test\`

\- \`type:refactor\`

\#\# Suggested Milestone

\`Phase E \- UI / Search / Diagnostics\`

---

# **4\) ถ้าคุณอยากเริ่มแบบ “น้อยที่สุดแต่ถูกทาง”**

ผมแนะนำให้คุณเปิดแค่ 5 issue แรกนี้ก่อน:

1. Fix NameMapping UUID fallback logic  
2. Complete GPS Queue REJECTED state  
3. Upgrade Database schema validation to 22 columns  
4. Unify Database read/write width  
5. Add retrieval step before AI matching

เพราะ 5 ตัวนี้ครอบคลุม risk หลักของระบบจาก:

* identity mapping  
* queue governance  
* schema drift  
* AI quality baseline. 

---

# **5\) วิธีใช้จริงแบบง่ายมาก**

## **วิธีสร้าง issue ทีละตัว**

1. เข้า Issues  
2. กด New issue  
3. วาง Title  
4. วาง Body  
5. ใส่ labels  
6. ใส่ milestone  
7. กด Create

## **วิธีจัดลำดับทำงาน**

* เริ่มจากทุก issue ที่เป็น priority:P0  
* ทำให้เสร็จทีละ issue  
* ค่อยขึ้น P1  
* P2 เอาไว้เป็นรอบถัดไป

---

# **6\) ถ้าคุณอยากให้ผมช่วยต่อจากตรงนี้**

ผมช่วยได้อีก 3 แบบ:

### **แบบ 1**

ผมช่วยทำ issue template มาตรฐาน สำหรับ:

* bug  
* refactor  
* tech debt  
* test

### **แบบ 2**

ผมช่วยทำ GitHub Projects board structure ให้คุณแบบพร้อมใช้งาน

### **แบบ 3**

ผมช่วยทำ “5 issue แรกแบบย่อที่สุด” ให้คุณ copy-paste เร็ว ๆ ภายในไม่กี่นาที

---

## **Commands ที่ผมใช้อ้างอิง source สำหรับจัด issue เหล่านี้**

* rg \--files  
* nl \-ba Config.gs  
* nl \-ba Service\_Master.gs | sed \-n '421,760p'  
* nl \-ba Service\_GPSFeedback.gs  
* nl \-ba Service\_Agent.gs  
* nl \-ba Service\_AutoPilot.gs  
* nl \-ba Service\_Search.gs  
* nl \-ba Service\_SchemaValidator.gs  
* nl \-ba Service\_SoftDelete.gs  
* nl \-ba Service\_SCG.gs | sed \-n '1,220p'  
* nl \-ba Utils\_Common.gs  
* nl \-ba Menu.gs | sed \-n '1,260p'  
* nl \-ba Test\_AI.gs  
* nl \-ba Test\_Diagnostic.gs

# ถ้าระบบนี้เป็นของคุณ

ครับ

สมมุตินะ ถ้าระบบนี้เป็นของคุณ และคุณได้ทำการแก้ไข ตามที่แนะนำเรียบร้อยแล้ว  
คุณจะสร้างระบบนี้ยังไงให้ดีขึ้นกว่านี้ กับปัญหาด้านล่างนี้ครับ  
1.เรื่องชื่อซ้ำ 2.เรื่องที่อย่ซ้ำ 3.เรื่องLatLong 4.เรื่องชื่อเขียนไม่เหมือนกัน 5.เรื่องคนละชื่อแต่ที่อยู่เดียวกัน 6.เรื่องชื่อเดียวกันแต่ที่อยู่ไม่เหมือนกัน 7.เรื่องชื่อเดียวกันแต่LatLongคนละที่ 8.เรื่องคนละชื่อแต่LatLongเดียวกัน  
คัดลอก

## **ถ้าระบบนี้เป็นของผม และผมแก้ตามที่แนะนำไปแล้วทั้งหมด**

## **ผมจะ “ออกแบบใหม่ให้ฉลาดขึ้น” สำหรับ 8 ปัญหานี้แบบนี้ครับ**

ก่อนตอบตรง ๆ ผมขอสรุปสั้นก่อนว่า ระบบปัจจุบันมีฐานที่ดีอยู่แล้ว เพราะมี:

* master database  
* alias mapping  
* clustering จากพิกัด  
* quality/confidence  
* GPS feedback queue  
* soft delete / merge UUID  
* AI ช่วยจับคู่ชื่อ. 

แต่ถ้าจะทำให้ “ดีกว่ามาก” สำหรับปัญหา dedup / identity / geospatial ambiguity ผมจะไม่ให้ระบบคิดแบบ “ชื่ออย่างเดียว” หรือ “LatLong อย่างเดียว” อีกต่อไป  
ผมจะทำให้มันคิดแบบ Entity Resolution Engine หรือเรียกง่าย ๆ ว่า:

1 record \= 1 สถานที่จริง \+ 1 ตัวตนธุรกิจจริง \+ หลักฐานหลายด้าน

กล่าวคือ ทุกการตัดสินใจจะไม่ดูแค่ชื่อ แต่จะดูพร้อมกัน 4 มิติ:

1. Name — ชื่อ  
2. Address — ที่อยู่  
3. Geo — Lat/Long  
4. Behavior / Source Evidence — มาจากไหน, ใครยืนยัน, ถูกพบกี่ครั้ง, คนขับยืนยันไหม, AI คิดังไง. 

---

# **1\) ผมจะเปลี่ยน “โมเดลข้อมูล” ก่อน**

ระบบเดิมมี Database, NameMapping, GPS\_Queue ซึ่งดีแล้ว แต่ถ้าเป็นของผม ผมจะเพิ่มแนวคิดข้อมูลอีก 3 ชั้น:

---

## **1.1 แยก “ชื่อที่แสดง” ออกจาก “ชื่อมาตรฐาน” ออกจาก “ชื่อค้นหา”**

ตอนนี้ระบบมี NAME, SUGGESTED, NORMALIZED, UUID, NameMapping อยู่แล้ว. 

แต่ผมจะขยายเป็น:

### **ใน master record**

* DISPLAY\_NAME \= ชื่อที่อยากให้ user เห็น  
* CANONICAL\_NAME \= ชื่อมาตรฐานกลาง  
* SEARCH\_NAME\_KEY \= normalized สำหรับค้นหา  
* BRANCH\_NAME  
* COMPANY\_NAME  
* SITE\_TYPE เช่น store/warehouse/branch/head office

### **ใน alias table**

* RAW\_NAME  
* NORMALIZED\_NAME  
* PHONETIC\_NAME  
* ALIAS\_TYPE เช่น human/AI/source/system  
* CANONICAL\_UUID

### **เหตุผล**

เพราะปัญหาหลายข้อที่คุณถาม เช่น

* ชื่อเดียวกันแต่คนละที่  
* คนละชื่อแต่ที่เดียวกัน  
* ชื่อไม่เหมือนกันแต่เป็นที่เดียวกัน  
  ทั้งหมดนี้เกิดจาก “ชื่อ” ถูกใช้เกิน 1 บทบาทพร้อมกัน. 

---

## **1.2 แยก “ที่อยู่ดิบ” ออกจาก “ที่อยู่แยกส่วน”**

ปัจจุบันระบบมี SYS\_ADDR, GOOGLE\_ADDR, PROVINCE, DISTRICT, POSTCODE. 

ถ้าเป็นของผม ผมจะเพิ่มโครงสร้างที่อยู่แบบละเอียด:

* RAW\_ADDRESS  
* NORMALIZED\_ADDRESS  
* HOUSE\_NO  
* MOO  
* ROAD  
* SOI  
* SUBDISTRICT  
* DISTRICT  
* PROVINCE  
* POSTCODE  
* ADDRESS\_HASH  
* ADDRESS\_CONFIDENCE

### **ทำไมต้องทำ**

เพราะ “ที่อยู่ซ้ำ” และ “คนละชื่อแต่ที่อยู่เดียวกัน” จะจับได้ง่ายกว่ามากถ้าไม่เก็บเป็น string ก้อนเดียว. 

---

## **1.3 แยก “พิกัดหลายแหล่ง” ออกจาก “พิกัดหลัก”**

ตอนนี้ระบบมี LAT, LNG, Coord\_Source, Coord\_Confidence, Coord\_Last\_Updated. ดีมาก แต่ยังเป็น single current coordinate มากกว่า coordinate history. 

ถ้าเป็นของผม ผมจะเพิ่ม “ตารางพิกัดหลักฐาน” แยกออกมา เช่น concept ใหม่ชื่อ:

### CoordinateEvidence

* UUID  
* LAT  
* LNG  
* SOURCE\_TYPE เช่น SCG\_System / Driver\_GPS / Google / Human / AI  
* SOURCE\_REF  
* REPORTED\_AT  
* CONFIDENCE  
* IS\_APPROVED  
* DISTANCE\_FROM\_CURRENT  
* CLUSTER\_ID

แล้ว Database จะเก็บแค่:

* PRIMARY\_LAT  
* PRIMARY\_LNG  
* PRIMARY\_COORD\_SOURCE  
* PRIMARY\_COORD\_CONFIDENCE

### **ทำไมต้องทำ**

เพราะปัญหา:

* ชื่อเดีวกันแต่ LatLong คนละที่  
* คนละชื่อแต่ LatLong เดียวกัน  
* GPS จาก driver หลายครั้งไม่ตรงกัน  
  ควรแก้ด้วย “หลักฐานหลายจุด” ไม่ใช่เขียนทับค่าเดิมอย่างเดียว. 

---

# **2\) ผมจะสร้าง “Matching Engine” ใหม่ให้คิดเป็นคะแนนหลายมิติ**

ระบบปัจจุบันมี:

* normalizeText()  
* clustering จากระยะทาง  
* quality / confidence  
* alias mapping  
* AI matching. 

ถ้าเป็นของผม ผมจะสร้าง engine กลางชื่อประมาณนี้:

* scoreNameMatch()  
* scoreAddressMatch()  
* scoreGeoMatch()  
* scoreSourceTrust()  
* scoreEntityMatch()

แล้วผลลัพธ์สุดท้ายจะเป็น:

* MATCH\_SCORE  
* MATCH\_REASON  
* MATCH\_CLASS

โดย MATCH\_CLASS อาจแบ่งเป็น:

* AUTO\_MERGE  
* LIKELY\_SAME\_ENTITY  
* REVIEW\_REQUIRED  
* LIKELY\_DIFFERENT  
* SAME\_NAME\_DIFFERENT\_SITE  
* DIFFERENT\_NAME\_SAME\_SITE

---

# **3\) คำตอบทีละข้อสำหรับ 8 ปัญหาของคุณ**

---

## **3.1 ปัญหา 1 — เรื่องชื่อซ้ำ**

### **สิ่งที่ผมจะทำ**

ผมจะไม่มอง “ชื่อซ้ำ” ว่าเป็นปัญหาเดียว  
แต่จะแยกเป็น 3 กรณี:

1. ชื่อซ้ำ \+ ที่อยู่เหมือน \+ พิกัดใกล้ → น่าจะ record ซ้ำ  
2. ชื่อซ้ำ \+ ที่อยู่ต่าง \+ พิกัดต่าง → น่าจะคนละสาขา  
3. ชื่อซ้ำ \+ ที่อยู่คล้าย \+ พิกัดกลาง ๆ → ต้อง review

### **วิธีทำ**

* normalize ชื่อหลายแบบ  
* แยก legal words / branch words / locality words  
* เก็บ COMPANY\_NAME กับ BRANCH\_NAME แยก  
* ใช้ scoring:  
  * exact normalized name  
  * phonetic similarity  
  * token overlap  
  * address similarity  
  * geo distance. 

### **ถ้าผมออกแบบ rule**

* ชื่อซ้ำอย่างเดียว ห้าม merge ทันที  
* ต้องมี address หรือ geo หนุนด้วยเสมอ

---

## **3.2 ปัญหา 2 — เรื่องที่อยู่ซ้ำ**

### **สิ่งที่ผมจะทำ**

ผมจะสร้าง ADDRESS\_HASH จากที่อยู่ที่ normalize แล้ว เช่น:

* ลบคำฟุ่มเฟือย  
* map คำย่อ  
* แยก province/district/postcode  
* normalize บ้านเลขที่/หมู่/ซอย/ถนน

### **Rule**

* ถ้า ADDRESS\_HASH เหมือนกัน และพิกัดใกล้ → สถานที่เดียวกันสูง  
* ถ้า ADDRESS\_HASH เหมือน แต่ชื่อคนละชุด → candidate “different name same site”  
* ถ้า address raw ต่าง แต่ parsed address เหมือน → same site ได้เหมือนกัน. 

### **สิ่งที่เพิ่มจากระบบเดิม**

ระบบเดิม parse address ได้บางส่วนแล้ว แต่ยังไม่ได้ใช้ address เป็น first-class matching entity เท่าชื่อกับพิกัด. 

---

## **3.3 ปัญหา 3 — เรื่อง LatLong**

### **สิ่งที่ผมจะทำ**

ผมจะเลิกคิดว่า Lat/Long เป็นค่าศักดิ์สิทธิ์ค่าเดียว  
แต่จะมองว่าเป็น “หลักฐานที่มีความน่าเชื่อถือหลายระดับ”

### **ผมจะมี 3 ชั้น**

1. Raw coordinates — ทุก source ที่เข้ามา  
2. Coordinate clusters — กลุ่มพิกัดที่อยู่ใกล้กัน  
3. Primary coordinate — พิกัดหลักที่ถูกเลือก

### **วิธีตัดสิน primary**

* Driver\_GPS approved \> Human verified \> SCG\_System \> Google geocode  
* ถ้ามีหลายหลักฐานใน cluster เดียว → confidence สูงขึ้น  
* ถ้ามี 2 cluster ห่างกันมาก → บอกว่า entity นี้ ambiguous ต้อง review. 

### **สิ่งที่ระบบใหม่จะดีกว่าเดิม**

ตอนนี้ queue ใช้ threshold 50 เมตร ซึ่งดีมาก แต่ยังคิดเป็น “ต่าง/ไม่ต่าง”  
ผมจะขยายให้คิดเป็น:

* same point  
* same compound  
* same site cluster  
* suspicious mismatch  
* definitely different site. 

---

## **3.4 ปัญหา 4 — เรื่องชื่อเขียนไม่เหมือนกัน**

### **สิ่งที่ผมจะทำ**

นี่คือเรื่อง alias / typo / variant ซึ่งระบบเดิมเริ่มทำแล้วผ่าน normalizeText(), NameMapping, และ AI keyword generation. 

แต่ถ้าเป็นของผม ผมจะเพิ่ม 5 เทคนิค:

### **1\. Multi-normalization**

* legal normalize  
* search normalize  
* phonetic normalize  
* branch-insensitive normalize  
* whitespace/punctuation normalize

### **2\. Thai phonetic / typo rules**

เช่น:

* เซ็นทรัล / เซนทรัล  
* เด็นโซ่ / DENSO  
* โลตัส / Lotus  
* เจดับบลิวดี / JWD

### **3\. Alias lifecycle**

ทุก alias จะมี:

* source  
* confidence  
* created by  
* last seen  
* canonical UUID

### **4\. Human-confirmed alias wins**

ถ้า human confirm alias แล้ว AI/heuristic ต้องไม่ override ง่าย ๆ

### **5\. Search expansion**

เวลาค้นหา ให้ query expand ด้วย alias/phonetic/AI keyword แต่ผลลัพธ์ยังต้องชี้ canonical UUID เดียว. 

---

## **3.5 ปัญหา 5 — คนละชื่อแต่ที่อยู่เดียวกัน**

### **สิ่งที่ผมจะทำ**

ผมจะถือว่านี่คือ candidate group สำคัญมาก เพราะมีได้ 3 ความจริง:

1. เป็นร้านเดียวกันแต่ชื่อเขียนคนละแบบ  
2. เป็นบริษัทแม่กับชื่อสาขา  
3. เป็นหลาย business unit ในสถานที่เดียวกันจริง

### **วิธีตัดสิน**

ใช้ matrix แบบนี้:

* Address same  
* Geo same/near  
* Name similar?  
* Owner/Shipment behavior same?  
* Source trust same?

### **การ classify**

* Address same \+ Geo same \+ Name similar → same entity likely  
* Address same \+ Geo same \+ Name very different → same site, different tenant/business  
* Address same \+ Geo near \+ multiple delivery patterns → review as site group. 

### **ผมจะเพิ่ม concept ใหม่**

* SITE\_UUID \= ตัวแทน “สถานที่”  
* ENTITY\_UUID \= ตัวแทน “ลูกค้าหรือร้าน”

เพราะบางครั้งสถานที่เดียวกันมีหลาย entity จริง  
นี่คือจุดที่ระบบปัจจุบันยังไม่แยกระหว่าง “site” กับ “business entity” ชัดพอ. 

---

## **3.6 ปัญหา 6 — ชื่อเดียวกันแต่ที่อยู่ไม่เหมือนกัน**

### **สิ่งที่ผมจะทำ**

ผมจะไม่ merge ทันทีเด็ดขาด  
ผมจะตั้ง rule ว่า:

same name \!= same entity

ถ้า:

* ชื่อเหมือน  
* แต่ address ต่างกัน  
* และ geo ต่างกันเกิน threshold  
  ให้ถือเป็น same brand / different site ก่อน

### **ผมจะเก็บเพิ่ม**

* BRAND\_KEY  
* SITE\_KEY  
* ENTITY\_KEY

เช่น:

* ชื่อเดียวกัน \= brand เดียว  
* แต่ address/geo คนละที่ \= site คนละแห่ง

### **ตัวอย่าง logic**

* "โลตัส" อาจมี 100 สาขา  
* "SCG EXPRESS" อาจมีหลายจุด  
* "บริษัท ABC" อาจมี warehouse \+ office คนละที่

ดังนั้น same name ต้องไปแยกต่อด้วย address \+ geo เสมอ. 

---

## **3.7 ปัญหา 7 — ชื่อเดียวกันแต่ LatLong คนละที่**

### **สิ่งที่ผมจะทำ**

ผมจะสร้าง geo conflict detector

### **Rule**

ถ้า normalized name เดียวกัน แต่มี coordinate clusters คนละกลุ่ม:

* \< 50m \= same  
* 50m–300m \= อาจ same compound / mapping error  
* 300m–2km \= suspicious  
* 2km \= likely different site unless address says otherwise

### **วิธีจัดการ**

* สร้าง queue พิเศษชื่อประมาณ EntityConflictQueue  
* ให้ admin เห็น conflict type เช่น:  
  * same\_name\_diff\_geo  
  * same\_name\_diff\_address  
  * diff\_name\_same\_geo  
* ระบบไม่ merge เองถ้า conflict สูง. 

### **ถ้าเป็นของผม**

ผมจะให้หน้า admin เห็นว่า:

* ชื่อเดียวกันนี้ถูกพบกี่ครั้ง  
* source ไหนบ้าง  
* พิกัดกระจายเป็นกี่ cluster  
* ที่อยู่สรุปเป็นกี่แบบ

---

## **3.8 ปัญหา 8 — คนละชื่อแต่ LatLong เดียวกัน**

### **สิ่งที่ผมจะทำ**

ผมจะถือว่านี่เป็นกรณี “สถานที่เดียวกัน แต่อาจคนละ entity  
ไม่ใช่ merge ทันที

### **ผมจะ classify เป็น 4 แบบ**

1. alias same entity  
2. same branch/site  
3. co-located different businesses  
4. bad geocode / copied GPS

### **วิธีแยก**

ใช้:

* address similarity  
* shipment pattern  
* owner/customer relationship  
* source trust  
* human review history. 

### **ตัวอย่าง rule**

* ถ้า LatLong เดียวกัน, address ใกล้มาก, name คล้าย → same entity likely  
* ถ้า LatLong เดียวกัน, address เหมือน, name ต่างมาก → same site different name  
* ถ้า LatLong เดียวกันแต่มาจาก geocode source เท่านั้น → confidence ต่ำ  
* ถ้า LatLong เดียวกันจาก driver-approved หลายรอบ → confidence สูง. 

---

# **4\) ถ้าเป็นของผม ผมจะเพิ่ม “Conflict Engine” โดยเฉพาะ**

ระบบเดิมมี:

* hidden duplicate scan  
* clustering  
* queue GPS feedback. 

แต่ถ้าจะให้ดีขึ้นจริง ผมจะเพิ่มชีต/โมดูลใหม่ เช่น:

## Conflict\_Queue

เก็บ conflict ที่ระบบตรวจพบอัตโนมัติ เช่น

* same\_name\_diff\_address  
* same\_name\_diff\_geo  
* diff\_name\_same\_address  
* diff\_name\_same\_geo  
* same\_address\_multi\_entity  
* geo\_outlier  
* alias\_conflict

### **แต่ละ row จะมี**

* conflict\_id  
* conflict\_type  
* candidate\_uuid\_1  
* candidate\_uuid\_2  
* name\_score  
* address\_score  
* geo\_score  
* final\_score  
* recommended\_action  
* reviewer  
* reviewed\_at  
* resolution

### **ทำไมต้องมี**

เพราะ 8 ปัญหาที่คุณถาม จริง ๆ คือ “conflict types” คนละแบบ  
ถ้าไม่มี conflict engine แยก สุดท้าย logic จะกระจายไปทั่วทุกฟังก์ชันและดูแลยาก. 

---

# **5\) ผมจะเปลี่ยนจาก “merge-centric” เป็น “site \+ entity-centric”**

นี่คือหัวใจสำคัญที่สุดของคำตอบนี้ครับ

ปัจจุบันระบบมองไปที่ “1 record \= 1 customer row” ซึ่งใช้งานได้ดีระดับหนึ่ง. 

แต่ถ้าเป็นของผม ผมจะปรับ model ให้เป็น:

## **5.1** Site

ตัวแทนสถานที่จริง

* address  
* coordinate clusters  
* postcode/district/province  
* place confidence

## **5.2** Entity

ตัวแทนลูกค้า/ร้าน/บริษัท

* canonical name  
* brand  
* alias set  
* linked site(s)

## **5.3** Evidence

หลักฐานที่สนับสนุน entity/site

* raw source row  
* SCG system  
* driver GPS  
* geocoding result  
* human review  
* AI suggestion

### **ทำไม model นี้ดีกว่า**

เพราะ 8 ปัญหาของคุณ หลายข้อไม่ใช่ “duplicate record” แต่คือ:

* same site, different entity  
* same entity, different name  
* same brand, different site  
* same name, bad geocode  
* same coordinate, shared compound

model แบบ site/entity/evidence จะตอบโจทย์ทั้งหมดดีกว่า “ชื่อกับพิกัดใน row เดียว”. 

---

# **6\) ถ้าจะทำให้ระบบตัดสินใจได้แม่นขึ้น ผมจะกำหนด “Decision Policy” ชัด ๆ**

ผมจะให้ระบบตัดสินใจแบบนี้:

---

## **Policy A — Auto Merge**

เมื่อ:

* name score สูงมาก  
* address score สูงมาก  
* geo score สูงมาก  
* ไม่มี conflict history  
* source เชื่อถือได้

## **Policy B — Auto Alias**

เมื่อ:

* ต่างชื่อเล็กน้อย  
* address/geo/supporting evidence ชี้ว่า entity เดียวกัน  
* แค่ต้องเพิ่ม variant name เข้า mapping

## **Policy C — Same Site / Different Entity**

เมื่อ:

* address/geo เหมือน  
* แต่ชื่อ/behavior ต่างกันมาก

## **Policy D — Same Brand / Different Site**

เมื่อ:

* canonical/brand name ใกล้กัน  
* แต่ address/geo ต่างชัด

## **Policy E — Review Required**

เมื่อ:

* score ขัดกัน  
* source หลายแหล่งให้ข้อมูลไม่ตรงกัน  
* AI กับ heuristic ขัดกัน  
* driver GPS conflict กับ system GPS. 

---

# **7\) ถ้าผมจะปรับ UX/Admin ให้ดีขึ้นด้วย**

ผมจะไม่ให้ admin ดูแต่ row ใน sheet แล้วเดาเอง  
ผมจะเพิ่มหน้า review ที่เห็นเป็น “คู่ conflict” เช่น:

### **card ซ้าย / card ขวา**

* ชื่อ  
* alias ทั้งหมด  
* ที่อยู่  
* province/district/postcode  
* lat/long  
* source history  
* shipment frequency  
* driver confirmations  
* confidence/quality  
* recommended action

แล้วให้กดได้:

* Merge Entity  
* Link as Alias  
* Mark Same Site Different Entity  
* Mark Same Brand Different Site  
* Reject Match. 

---

# **8\) ถ้าจะตอบสั้น ๆ ทีละข้อแบบเป็น checklist**

## **1\. ชื่อซ้ำ**

* แยก brand / branch / display name  
* อย่า merge เพราะชื่อเหมือนอย่างเดียว  
* ใช้ address \+ geo \+ source evidence ร่วมกัน

## **2\. ที่อยู่ซ้ำ**

* parse address เป็นโครงสร้าง  
* สร้าง address hash  
* ใช้เป็นตัวชี้ site identity

## **3\. LatLong**

* เก็บ coordinate evidence หลายแหล่ง  
* ใช้ clustering  
* เลือก primary coordinate จากความเชื่อถือ

## **4\. ชื่อเขียนไม่เหมือนกัน**

* สร้าง alias engine \+ phonetic \+ typo rules  
* แยก normalized ออกจาก search keyword  
* human-confirmed alias ต้องมีน้ำหนักสูง

## **5\. คนละชื่อแต่ที่อยู่เดียวกัน**

* classify เป็น same site candidate  
* อย่า merge อัตโนมัติทันที

## **6\. ชื่อเดียวกันแต่ที่อยู่ไม่เหมือนกัน**

* จัดเป็น same brand / different site candidate  
* แยก site key จาก entity key

## **7\. ชื่อเดียวกันแต่ LatLong คนละที่**

* สร้าง geo conflict queue  
* แยก coordinate cluster

## **8\. คนละชื่อแต่ LatLong เดียวกัน**

* classify เป็น same site / co-located entities / geocode issue  
* ใช้ address \+ source history ช่วยตัดสิน

---

# **9\) ถ้าสรุปให้เหลือ “3 ของที่ควรสร้างเพิ่ม” ถ้าจะให้ระบบนี้เก่งขึ้นจริง**

## **อย่างที่ 1:** Site **model**

เพื่อแก้ปัญหา address/latlong/compound/shared location

## **อย่างที่ 2:** Evidence **model**

เพื่อเก็บหลักฐานจากหลาย source แทนการ overwrite ค่าเดียว

## **อย่างที่ 3:** Conflict Engine

เพื่อจัดการ 8 ปัญหาที่คุณถามอย่างเป็นระบบและ review ได้

นี่คือ 3 อย่างที่ถ้าทำเพิ่ม ระบบจะก้าวจาก “sheet-based master list” ไปเป็น “entity resolution platform” ได้เลยครับ

---

# **10\) สรุปสุดท้ายแบบตรงที่สุด**

ถ้าระบบนี้เป็นของผม ผมจะไม่พยายามแก้ 8 ปัญหาด้วย “สูตรเดียว”  
เพราะจริง ๆ แล้วทั้ง 8 ข้อคือ คนละประเภทของความคลุมเครือ:

* บางข้อคือ same entity  
* บางข้อคือ same site  
* บางข้อคือ same brand  
* บางข้อคือ bad source data  
* บางข้อคือ human review case

ดังนั้นระบบที่ดีขึ้นต้อง:

1. แยก ชื่อ / ที่อยู่ / พิกัด / หลักฐาน ออกจากกัน  
2. คิดเป็น score หลายมิติ  
3. เก็บ history \+ evidence  
4. มี conflict queue  
5. แยก Site ออกจาก Entity

ซึ่งแนวทางนี้ต่อยอดได้ตรงจากสิ่งที่ repo ปัจจุบันมีอยู่แล้ว เช่น alias mapping, clustering, quality/confidence, GPS queue, และ soft delete/merge UUID. 

---

## **Commands ที่ผมใช้อ้างอิง source เพื่อสรุปแนวออกแบบนี้**

* rg \--files  
* nl \-ba Config.gs  
* nl \-ba Service\_Master.gs | sed \-n '1,240p'  
* nl \-ba Service\_Master.gs | sed \-n '241,420p'  
* nl \-ba Service\_Master.gs | sed \-n '421,760p'  
* nl \-ba Service\_GPSFeedback.gs  
* nl \-ba Service\_Agent.gs  
* nl \-ba Service\_AutoPilot.gs  
* nl \-ba Service\_Search.gs  
* nl \-ba Service\_SoftDelete.gs  
* nl \-ba Service\_SCG.gs | sed \-n '1,220p'  
* nl \-ba Service\_SCG.gs | sed \-n '221,450p'  
* nl \-ba Service\_GeoAddr.gs  
* nl \-ba Utils\_Common.gs  
* nl \-ba Menu.gs | sed \-n '1,260p'

ถ้าคุณต้องการ ผมช่วยต่อได้ 2 แบบที่ “ใช้งานต่อได้จริง” มาก:

1. ออกแบบ Schema ใหม่ของระบบเวอร์ชันถัดไป ว่าต้องเพิ่มชีตอะไร คอลัมน์อะไรบ้าง  
2. วาด Flow การตัดสินใจของ Conflict Engine แบบ decision tree สำหรับทั้ง 8 กรณี

ถ้าคุณอยากเอาไปทำต่อจริง ผมแนะนำให้พิมพ์ต่อว่า:

“ช่วยออกแบบ Schema ใหม่ของระบบเวอร์ชันถัดไปให้ผม”

