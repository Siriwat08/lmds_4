# Next-Gen Entity Resolution Design

เอกสารนี้สรุปแนวทางยกระดับระบบจาก master list + alias mapping + GPS feedback queue ไปเป็น **Entity Resolution Platform** ที่รองรับปัญหา dedup, identity และ geospatial ambiguity ได้เป็นระบบมากขึ้น โดยต่อยอดจากโครงสร้างปัจจุบันใน `Database`, `NameMapping`, `GPS_Queue` และ service ที่มีอยู่แล้ว.

---

## 1) Design Goal

หลักคิดใหม่คือ:

> 1 record = 1 สถานที่จริง + 1 ตัวตนธุรกิจจริง + หลักฐานหลายด้าน

ระบบเวอร์ชันถัดไปไม่ควรตัดสินจากแค่ชื่อ หรือพิกัดอย่างเดียว แต่ควรประเมินพร้อมกัน 4 มิติ:

1. **Name** — ชื่อจริง, ชื่อแสดงผล, alias, phonetic
2. **Address** — ที่อยู่ดิบและที่อยู่แบบแยกส่วน
3. **Geo** — พิกัดจากหลายแหล่ง + clustering
4. **Behavior / Source Evidence** — source ไหนส่งมา, driver ยืนยันหรือไม่, พบซ้ำกี่ครั้ง, AI ให้ความเห็นอย่างไร

---

## 2) Why the current system is a strong base

ระบบปัจจุบันมี foundation ที่ดีอยู่แล้ว:

- master database พร้อม UUID และ quality score
- alias mapping ใน `NameMapping`
- clustering จากระยะทาง
- confidence / quality field
- GPS feedback queue จากคนขับ
- soft delete + merge UUID
- AI assistance สำหรับ name matching

ดังนั้นการ redesign ที่แนะนำ **ไม่ใช่การรื้อระบบทั้งหมด** แต่เป็นการเพิ่มชั้นข้อมูลและ decision engine ให้ชัดขึ้น.

---

## 3) Proposed data model

## 3.1 Core separation

ให้แยก model ออกเป็น 3 แกนหลัก:

### A. Site
ตัวแทน “สถานที่จริง”

ตัวอย่าง field:

- `SITE_UUID`
- `DISPLAY_ADDRESS`
- `NORMALIZED_ADDRESS`
- `HOUSE_NO`
- `MOO`
- `ROAD`
- `SOI`
- `SUBDISTRICT`
- `DISTRICT`
- `PROVINCE`
- `POSTCODE`
- `ADDRESS_HASH`
- `PRIMARY_LAT`
- `PRIMARY_LNG`
- `PRIMARY_COORD_SOURCE`
- `PRIMARY_COORD_CONFIDENCE`
- `SITE_CONFIDENCE`
- `SITE_STATUS`

### B. Entity
ตัวแทน “ลูกค้า / ร้าน / บริษัท / business identity”

ตัวอย่าง field:

- `ENTITY_UUID`
- `DISPLAY_NAME`
- `CANONICAL_NAME`
- `SEARCH_NAME_KEY`
- `PHONETIC_NAME`
- `BRAND_KEY`
- `COMPANY_NAME`
- `BRANCH_NAME`
- `SITE_TYPE` เช่น `store`, `warehouse`, `branch`, `head_office`
- `ENTITY_STATUS`
- `MERGED_TO_ENTITY_UUID`

### C. Evidence
ตัวแทน “หลักฐาน” ที่ใช้สนับสนุน site หรือ entity

ตัวอย่าง field:

- `EVIDENCE_UUID`
- `EVIDENCE_TYPE` เช่น `SCG_System`, `Driver_GPS`, `Google`, `Human_Review`, `AI_Suggestion`
- `SOURCE_REF`
- `RAW_NAME`
- `RAW_ADDRESS`
- `RAW_LAT`
- `RAW_LNG`
- `REPORTED_AT`
- `SOURCE_TRUST_SCORE`
- `IS_APPROVED`
- `LINKED_SITE_UUID`
- `LINKED_ENTITY_UUID`

แนวคิดนี้ช่วยให้ระบบเก็บ history และหลายมุมมองของข้อมูลได้ โดยไม่ต้อง overwrite ค่าหลักทุกครั้ง.

---

## 3.2 Name model

ให้แยกบทบาทของชื่อออกจากกันให้ชัด:

### ใน master entity

- `DISPLAY_NAME` — ชื่อที่ต้องการให้ user เห็น
- `CANONICAL_NAME` — ชื่อมาตรฐานกลาง
- `SEARCH_NAME_KEY` — ชื่อ normalized สำหรับค้นหา
- `PHONETIC_NAME` — ช่วยจับคำอ่านใกล้เคียง
- `COMPANY_NAME` — ชื่อบริษัทแม่
- `BRANCH_NAME` — ชื่อสาขา
- `SITE_TYPE` — ประเภทสถานที่

### ใน alias table

- `RAW_NAME`
- `NORMALIZED_NAME`
- `PHONETIC_NAME`
- `ALIAS_TYPE` เช่น `human`, `AI`, `source`, `system`
- `CANONICAL_ENTITY_UUID`
- `ALIAS_CONFIDENCE`
- `CREATED_BY`
- `LAST_SEEN_AT`

เหตุผลสำคัญคือ “ชื่อเดียวกัน”, “คนละชื่อแต่ที่เดียวกัน”, และ “ชื่อไม่เหมือนกันแต่ entity เดียวกัน” เป็นคนละปัญหากัน ถ้าชื่อทุกแบบถูกรวมอยู่ใน field เดียว ระบบจะตัดสินพลาดง่าย.

---

## 3.3 Address model

ให้แยก “ที่อยู่ดิบ” ออกจาก “ที่อยู่ parsed แล้ว” เสมอ:

- `RAW_ADDRESS`
- `NORMALIZED_ADDRESS`
- `HOUSE_NO`
- `MOO`
- `ROAD`
- `SOI`
- `SUBDISTRICT`
- `DISTRICT`
- `PROVINCE`
- `POSTCODE`
- `ADDRESS_HASH`
- `ADDRESS_CONFIDENCE`

### Address policy

- ใช้ `ADDRESS_HASH` เป็น site identity signal ระดับต้น
- ถ้า `RAW_ADDRESS` ต่าง แต่ parsed address เหมือนกัน ยังถือว่า same site candidate ได้
- ถ้า address เหมือน + geo ใกล้มาก ให้คะแนน same site สูง

---

## 3.4 Coordinate evidence model

ปัจจุบันระบบมี primary coordinate เพียงชุดเดียวใน `Database` ซึ่งดีสำหรับ operational lookup แต่ไม่พอสำหรับ conflict analysis.

แนะนำตารางใหม่ `CoordinateEvidence`:

- `COORD_EVIDENCE_UUID`
- `SITE_UUID`
- `ENTITY_UUID` (optional)
- `LAT`
- `LNG`
- `SOURCE_TYPE` เช่น `SCG_System`, `Driver_GPS`, `Google`, `Human`, `AI`
- `SOURCE_REF`
- `REPORTED_AT`
- `CONFIDENCE`
- `IS_APPROVED`
- `DISTANCE_FROM_PRIMARY`
- `CLUSTER_ID`

ส่วน table หลักคงเก็บเพียง:

- `PRIMARY_LAT`
- `PRIMARY_LNG`
- `PRIMARY_COORD_SOURCE`
- `PRIMARY_COORD_CONFIDENCE`

### Coordinate selection policy

ลำดับความน่าเชื่อถือที่แนะนำ:

1. `Driver_GPS` ที่ approved แล้ว
2. `Human verified`
3. `SCG_System`
4. `Google geocode`
5. `AI inferred`

ถ้ามีหลายหลักฐานใน cluster เดียวกัน confidence ควรเพิ่มขึ้น แต่ถ้ามีหลาย cluster ห่างกันมากควรสร้าง conflict แทนการเขียนทับทันที.

---

## 4) Matching engine

แนะนำให้มี scoring engine กลางแทน rule กระจายหลายจุด:

- `scoreNameMatch()`
- `scoreAddressMatch()`
- `scoreGeoMatch()`
- `scoreSourceTrust()`
- `scoreBehaviorMatch()`
- `scoreEntityMatch()`

### Output มาตรฐาน

- `MATCH_SCORE`
- `MATCH_REASON`
- `MATCH_CLASS`

### Match class ที่แนะนำ

- `AUTO_MERGE`
- `LIKELY_SAME_ENTITY`
- `REVIEW_REQUIRED`
- `LIKELY_DIFFERENT`
- `SAME_NAME_DIFFERENT_SITE`
- `DIFFERENT_NAME_SAME_SITE`
- `SAME_BRAND_DIFFERENT_SITE`
- `CO_LOCATED_DIFFERENT_ENTITIES`

---

## 5) Conflict engine

8 ปัญหาหลักของระบบ dedup จริง ๆ คือ “คนละประเภทของ conflict” จึงควรมี queue กลางสำหรับ review โดยเฉพาะ.

แนะนำตาราง `Conflict_Queue`:

- `CONFLICT_ID`
- `CONFLICT_TYPE`
- `CANDIDATE_1_TYPE` (`site` / `entity` / `evidence`)
- `CANDIDATE_1_UUID`
- `CANDIDATE_2_TYPE`
- `CANDIDATE_2_UUID`
- `NAME_SCORE`
- `ADDRESS_SCORE`
- `GEO_SCORE`
- `SOURCE_TRUST_SCORE`
- `FINAL_SCORE`
- `RECOMMENDED_ACTION`
- `REVIEWER`
- `REVIEWED_AT`
- `RESOLUTION`
- `NOTES`

### Conflict types ที่ควรมีอย่างน้อย

- `same_name_diff_address`
- `same_name_diff_geo`
- `diff_name_same_address`
- `diff_name_same_geo`
- `same_address_multi_entity`
- `geo_outlier`
- `alias_conflict`
- `same_brand_diff_site`

ข้อดีคือ logic จะถูกรวมไว้ที่เดียว ดูแลง่าย และออกแบบ UX review ได้ง่ายขึ้น.

---

## 6) Decision policy

### Policy A — Auto Merge
ใช้เมื่อ:

- name score สูงมาก
- address score สูงมาก
- geo score สูงมาก
- source น่าเชื่อถือ
- ไม่มี conflict history

### Policy B — Auto Alias
ใช้เมื่อ:

- ชื่อต่างกันเล็กน้อย
- address/geo บ่งชี้ว่าเป็น entity เดียวกัน
- ควรเพิ่ม alias มากกว่ารวม row ทันที

### Policy C — Same Site / Different Entity
ใช้เมื่อ:

- address เหมือนหรือใกล้มาก
- geo เหมือนหรืออยู่ cluster เดียวกัน
- ชื่อและพฤติกรรมต่างกันชัด

### Policy D — Same Brand / Different Site
ใช้เมื่อ:

- brand ใกล้กันมาก
- แต่ address/geo ต่างชัดเจน

### Policy E — Review Required
ใช้เมื่อ:

- score ขัดกันเอง
- source หลายแหล่งให้ข้อมูลไม่ตรงกัน
- AI กับ heuristic ให้ผลคนละทาง
- driver GPS conflict กับพิกัดเดิม

---

## 7) Resolution guidance for the 8 common problems

## 7.1 ชื่อซ้ำ

ให้แยก 3 กรณี:

1. ชื่อซ้ำ + ที่อยู่เหมือน + geo ใกล้ → duplicate candidate
2. ชื่อซ้ำ + ที่อยู่ต่าง + geo ต่าง → คนละ site/branch
3. ชื่อซ้ำ + หลักฐานก้ำกึ่ง → review

**Rule:** ห้าม merge เพราะชื่อเหมือนอย่างเดียว.

## 7.2 ที่อยู่ซ้ำ

- parse address เป็นส่วนย่อย
- สร้าง `ADDRESS_HASH`
- ใช้ address เป็น first-class signal

**Rule:** ถ้า `ADDRESS_HASH` เหมือน + geo ใกล้ → same site likelihood สูง

## 7.3 Lat/Long conflict

มองพิกัดเป็นหลักฐานหลายจุด ไม่ใช่ค่าศักดิ์สิทธิ์ค่าเดียว.

แยกเป็น 3 ชั้น:

1. raw coordinates
2. coordinate clusters
3. primary coordinate

## 7.4 ชื่อเขียนไม่เหมือนกัน

เพิ่ม 5 เทคนิค:

1. multi-normalization
2. phonetic / typo rules
3. alias lifecycle
4. human-confirmed alias มีน้ำหนักสูง
5. search expansion แต่ชี้กลับ canonical UUID เดียว

## 7.5 คนละชื่อแต่ที่อยู่เดียวกัน

จัดเป็น **same-site candidate** ก่อน ไม่ใช่ merge ทันที เพราะอาจเป็น:

- ร้านเดียวกันคนละชื่อ
- บริษัทแม่กับชื่อสาขา
- หลาย business unit ในสถานที่เดียวกัน

## 7.6 ชื่อเดียวกันแต่ที่อยู่ไม่เหมือนกัน

จัดเป็น **same brand / different site** โดย default.

**Rule:** `same name != same entity`

## 7.7 ชื่อเดียวกันแต่ Lat/Long คนละที่

แนะนำ geo conflict detector:

- `< 50m` = same point
- `50–300m` = same compound / mapping error ได้
- `300m–2km` = suspicious
- `> 2km` = likely different site unless address strongly matches

## 7.8 คนละชื่อแต่ Lat/Long เดียวกัน

จัดเป็น 4 แบบให้เลือก:

- alias same entity
- same branch/site
- co-located different businesses
- bad geocode / copied GPS

---

## 8) Suggested admin UX

หน้า review ควรเปลี่ยนจากดูแค่ row เป็นการดู “candidate pair” หรือ “candidate group” พร้อม context:

- card ซ้าย / card ขวา
- display name + canonical name
- alias ทั้งหมด
- ที่อยู่ parsed
- lat/lng และ cluster info
- source history
- shipment frequency / behavior summary
- confidence breakdown
- recommended action

### Action buttons ที่ควรมี

- `Merge Entity`
- `Link as Alias`
- `Mark Same Site Different Entity`
- `Mark Same Brand Different Site`
- `Reject Match`

---

## 9) Migration path from the current repo

เพื่อให้ implement ได้จริง แนะนำทำเป็น 3 phase:

### Phase 1 — Extend current schema

เพิ่ม sheet/table ใหม่โดยไม่ทำลายของเดิม:

- `Site_Master`
- `Entity_Master`
- `Entity_Alias`
- `CoordinateEvidence`
- `Conflict_Queue`

พร้อมเก็บ `SITE_UUID` และ `ENTITY_UUID` อ้างกลับจาก `Database` เดิมในช่วง transitional period.

### Phase 2 — Introduce scoring services

สร้าง service ใหม่ เช่น:

- `Service_MatchEngine.gs`
- `Service_ConflictEngine.gs`
- `Service_SiteModel.gs`

โดยเริ่มจากอ่านข้อมูลเดิมใน `Database`, `NameMapping`, `GPS_Queue` และใช้ AI เป็นเพียงหนึ่งใน evidence sources.

### Phase 3 — Move write-paths to evidence-first

เปลี่ยน flow จาก “overwrite row หลัก” เป็น:

1. รับ source row
2. สร้าง evidence
3. run matching engine
4. update site/entity primary record ตาม policy
5. ถ้ามี conflict ให้ push เข้า queue

---

## 10) Implementation priority: the 3 highest-value additions

ถ้าต้องเลือกทำเพิ่มเพียง 3 อย่างก่อน ให้เริ่มจาก:

1. **Site model** — เพื่อแยก “สถานที่” ออกจาก “ตัวตนธุรกิจ”
2. **Evidence model** — เพื่อเก็บข้อมูลหลายแหล่งโดยไม่ overwrite ทันที
3. **Conflict engine** — เพื่อจัดการทั้ง 8 ปัญหาแบบเป็นระบบ

3 ส่วนนี้จะพาระบบจาก spreadsheet-based master list ไปสู่ entity resolution platform ได้จริง.

---

## 11) Expected impact

เมื่อเทียบกับแนวคิดปัจจุบัน ระบบใหม่จะตอบคำถามได้ชัดขึ้นว่า record ที่กำลังพิจารณาเป็นกรณีใด:

- same entity
- same site
- same brand
- bad source data
- human review case

และนั่นทำให้การ merge, aliasing, geo correction และ admin review มีเหตุผลรองรับมากขึ้น.
