# 🚛 Logistics Master Data System — V4.1

ระบบจัดการฐานข้อมูล Logistics สำหรับ SCG JWD

---

## 📁 โครงสร้างไฟล์ทั้งหมด (21 ไฟล์)

### Configuration
| ไฟล์ | หน้าที่ |
|---|---|
| `Config.gs` | ค่าคงที่ทุกอย่าง, column index, SCG config, GPS config |

### Utilities
| ไฟล์ | หน้าที่ |
|---|---|
| `Utils_Common.gs` | normalizeText, generateUUID, getHaversineDistanceKM, getBestName_Smart |

### Services (Core)
| ไฟล์ | หน้าที่ |
|---|---|
| `Service_Master.gs` | จัดการ Database, Sync GPS, Clustering, Quality/Confidence |
| `Service_SCG.gs` | ดึงข้อมูล SCG API, จับคู่พิกัด, สรุปงาน |
| `Service_Search.gs` | Search Engine สำหรับ WebApp |
| `Service_GeoAddr.gs` | Google Maps functions, Postal cache |

### Services (New v4.1)
| ไฟล์ | หน้าที่ |
|---|---|
| `Service_GPSFeedback.gs` | GPS Queue Management, อนุมัติพิกัดจากคนขับ |
| `Service_SchemaValidator.gs` | ตรวจสอบโครงสร้างชีตก่อนทำงาน |
| `Service_SoftDelete.gs` | Soft Delete, Merge UUID ซ้ำซ้อน |

### Services (AI & Automation)
| ไฟล์ | หน้าที่ |
|---|---|
| `Service_Agent.gs` | AI Tier 4 Smart Resolution |
| `Service_AutoPilot.gs` | Background AI Indexing |

### Services (Notifications & Maintenance)
| ไฟล์ | หน้าที่ |
|---|---|
| `Service_Notify.gs` | LINE + Telegram Notifications (authoritative) |
| `Service_Maintenance.gs` | ล้าง Backup เก่า, เช็ค Cell Usage |

### UI & Menu
| ไฟล์ | หน้าที่ |
|---|---|
| `Menu.gs` | เมนูทั้งหมดใน Google Sheets |
| `WebApp.gs` | WebApp Controller (doGet, doPost) |
| `Index.html` | หน้าเว็บค้นหาพิกัดลูกค้า |

### Setup
| ไฟล์ | หน้าที่ |
|---|---|
| `Setup_Security.gs` | ตั้งค่า API Keys อย่างปลอดภัย |
| `Setup_Upgrade.gs` | อัปเกรด Schema, ค้นหา Duplicates |

### Testing
| ไฟล์ | หน้าที่ |
|---|---|
| `Test_AI.gs` | Debug AI functions |
| `Test_Diagnostic.gs` | ตรวจสอบระบบ Engine + Sheets |

---

## 🗂️ ชีตที่ต้องมีใน Google Sheets

| ชีต | หน้าที่ | คอลัมน์สำคัญ |
|---|---|---|
| `Database` | Golden Record ลูกค้า | A-V (22 col) |
| `NameMapping` | Variant → Master UUID | A-E (5 col) |
| `SCGนครหลวงJWDภูมิภาค` | GPS จริงจากคนขับ | O=LAT, P=LONG, AK=SYNC_STATUS |
| `Data` | งานประจำวัน | AA=LatLong_Actual |
| `Input` | Shipment + Cookie | B1=Cookie |
| `GPS_Queue` | รอ Admin อนุมัติพิกัด | H=Approve, I=Reject |
| `PostalRef` | รหัสไปรษณีย์ | postcode, district, province |
| `ข้อมูลพนักงาน` | Email คนขับ | col B=ชื่อ, col G=Email |
| `สรุป_เจ้าของสินค้า` | สรุปรายงาน | - |
| `สรุป_Shipment` | สรุปรายงาน | - |

---

## 🔧 การติดตั้งครั้งแรก

1. Copy ไฟล์ .gs ทั้งหมดเข้า Google Apps Script
2. Copy Index.html เข้า Google Apps Script
3. รัน `setupEnvironment()` เพื่อใส่ Gemini API Key
4. รัน `createGPSQueueSheet()` เพื่อสร้างชีต GPS_Queue
5. รัน `initializeRecordStatus()` เพื่อตั้งค่าสถานะเริ่มต้น
6. รัน `runFullSchemaValidation()` เพื่อตรวจสอบโครงสร้าง
7. Deploy WebApp (Execute as: Me, Access: Anyone)

---

## ✅ การเปลี่ยนแปลงใน V4.1

### Bug Fixes
- แก้ `checklsEPOD` → `checkIsEPOD` (typo ที่ทำให้สรุปงานผิด)
- แก้ `lastRow` ghost rows ใน `finalizeAndClean` ด้วย `deleteRows()`
- แก้ UTF-8 Cache Bomb: วัดขนาดด้วย Bytes แทน `.length`
- แก้ NaN_NaN Black Hole ใน Clustering
- แก้ Negative Row Count ใน `applyMasterCoordinatesToDailyJob`
- แก้ Hardcode 25 คอลัมน์ → `getLastColumn()`

### New Features
- GPS Feedback Loop: รับพิกัดจริงจากคนขับเข้า GPS_Queue
- Schema Validator: ตรวจสอบโครงสร้างชีตก่อนทำงาน
- Soft Delete + MERGED_TO_UUID: ไม่ลบข้อมูลจริง
- SYNC_STATUS Checkpoint: ป้องกันการประมวลผลซ้ำ
- COL_CONFIDENCE และ COL_QUALITY คำนวณเป็น % จริงๆ

### Code Cleanup
- ลบ `normalizeText()` ซ้ำออกจาก Service_SCG.gs
- ลบ `runAgentLoop()` ซ้ำออกจาก Service_Agent.gs
- ลบ `sendLineNotify/Telegram` ซ้ำออกจาก Service_Maintenance.gs
- ลบ `clearAllSCGSheets_UI()` ซ้ำออกจาก Menu.gs
- ลบฟังก์ชันที่ไม่ใช้ 5 ตัวออกจาก Utils_Common.gs
- ลบ Haversine Fallback ออกจาก Setup_Upgrade.gs

---

## 📞 ติดต่อ

SCG JWD Logistics — Logistics Architect Team
