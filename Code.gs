/**
 * ═══════════════════════════════════════════════════════════════════════
 * Trip Logger Pro — Google Apps Script Backend
 * ═══════════════════════════════════════════════════════════════════════
 * โครงสร้างข้อมูล (4 ชีตใน Spreadsheet เดียวกับที่ผูก Apps Script นี้ไว้):
 *
 *   Sites   : SiteId | SiteName | Status | CreatedAt | ConfirmedAt |
 *             StartMile | EndMile | CarImgUrl
 *   Hotels  : SiteId | HotelNo | ItemNo | ImgUrl | Desc
 *   Fuel    : SiteId | FuelNo | BillUrl | PreUrl | PostUrl | Date | Province
 *   (Counter): เก็บเลขลำดับ SiteId ล่าสุด ป้องกัน id ซ้ำ
 *
 * วิธีติดตั้ง:
 * 1. สร้าง Google Sheet เปล่า 1 ไฟล์ (ไม่ต้องสร้างชีต/หัวตารางเอง
 *    โค้ดนี้จะสร้างให้อัตโนมัติตอนเปิดเว็บแอปครั้งแรก)
 * 2. คัดลอก Spreadsheet ID จาก URL ของ Sheet นั้น (เลขชุดยาวๆ ระหว่าง
 *    /d/ กับ /edit) มาใส่ใน SPREADSHEET_ID ด้านล่าง
 * 3. เมนู ส่วนขยาย (Extensions) → Apps Script (หรือสร้างโปรเจกต์ใหม่
 *    จาก script.google.com ก็ได้เช่นกัน เพราะตอนนี้ไม่ต้องพึ่ง
 *    getActiveSpreadsheet() ที่ใช้ได้แค่กับ container-bound project แล้ว)
 * 4. ลบโค้ดเริ่มต้นออก แล้ววางไฟล์นี้แทน (ตั้งชื่อไฟล์ Code.gs)
 * 5. สร้างไฟล์ HTML ใหม่ชื่อ Index แล้ววางโค้ดที่ให้แยกไว้
 * 6. แก้ค่า FOLDER_ID ด้านล่างให้เป็น Id ของโฟลเดอร์ Google Drive ที่จะเก็บรูป
 *    (สร้างโฟลเดอร์ใหม่ใน Drive ก่อน แล้วคัดลอกเลขจาก URL)
 * 7. กด Deploy → New deployment → เลือกประเภท "Web app"
 *    - Execute as: Me
 *    - Who has access: เลือกตามที่ต้องการ (Anyone หรือ Anyone in org)
 * 8. คัดลอกลิงก์เว็บแอปที่ได้ไปแชร์ให้ทีมใช้งาน
 * ═══════════════════════════════════════════════════════════════════════
 */

// ── ตั้งค่า ──────────────────────────────────────────────────────────────
const FOLDER_ID = "1hHN-2eMcDVyoD4KK_bc_S6YgTnqePQfT"; // เปลี่ยนเป็น Id โฟลเดอร์ Drive ของคุณ

// อีเมลที่จะรับรายงาน Word ทุกครั้งที่มีคนกด "ยืนยัน (จบทริป)"
// ใช้ MailApp ของ Apps Script — ส่งจากบัญชี Google ที่ deploy เว็บแอปนี้ไว้
// โดยอัตโนมัติ ไม่ต้องตั้งค่า App Password หรือ SMTP เพิ่มเติมใดๆ เลย
const EMAIL_RECIPIENT = "sawitreephi@cpall.co.th";

// Spreadsheet ID ที่จะใช้เก็บข้อมูล — คัดลอกจาก URL ของ Google Sheet
// เช่น https://docs.google.com/spreadsheets/d/XXXXXXXXXXXXXXXX/edit
//                                              ^^^^^^^^^^^^^^^^ เลขส่วนนี้
// ถ้าเว้นว่างไว้เป็น "" ระบบจะลองใช้ SpreadsheetApp.getActiveSpreadsheet()
// แทน (ใช้ได้เฉพาะกรณีสร้างโปรเจกต์ผ่านเมนู Extensions ของ Sheet เท่านั้น —
// ถ้าสร้างโปรเจกต์จาก script.google.com ตรงๆ (standalone) ต้องใส่ ID ตรงนี้เสมอ
// ไม่งั้นจะเจอ error "Cannot read properties of null (reading 'getSheetByName')")
const SPREADSHEET_ID = "";

/** เปิด Spreadsheet ที่ถูกต้องเสมอ ไม่ว่าโปรเจกต์จะเป็น standalone หรือ container-bound */
function _getSpreadsheet() {
  if (SPREADSHEET_ID) {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error(
      "ไม่พบ Spreadsheet ที่ผูกไว้ — เนื่องจากนี่เป็น standalone project " +
      "กรุณาใส่ค่า SPREADSHEET_ID ที่ด้านบนของไฟล์ Code.gs " +
      "(คัดลอกจาก URL ของ Google Sheet ที่ต้องการใช้เก็บข้อมูล)"
    );
  }
  return ss;
}

const SHEET_SITES  = "Sites";
const SHEET_HOTELS = "Hotels";
const SHEET_FUEL   = "Fuel";
const SHEET_COUNTER = "Counter";

const HOTEL_COUNT = 3;       // จำนวนโรงแรมต่อไซต์
const HOTEL_ITEM_COUNT = 3;  // จำนวนรูปต่อโรงแรม
const FUEL_COUNT = 20;       // จำนวนครั้งเติมน้ำมันต่อไซต์

const SITES_HEADER  = ["SiteId", "SiteName", "Status", "CreatedAt", "ConfirmedAt",
                        "StartMile", "EndMile", "CarImgUrl", "ReportUrl"];
const HOTELS_HEADER = ["SiteId", "HotelNo", "ItemNo", "ImgUrl", "Desc"];
const FUEL_HEADER   = ["SiteId", "FuelNo", "BillUrl", "PreUrl", "PostUrl", "Date", "Province"];


// ─────────────────────────────────────────────────────────────────────────
// 1. Entry point — เสิร์ฟหน้าเว็บ
// ─────────────────────────────────────────────────────────────────────────
function doGet(e) {
  ensureSheetsExist();
  // หมายเหตุ: เดิมเคยมี .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
  // ต่อท้ายตรงนี้ แต่พบว่าทำให้เกิด iframe sandbox warning
  // ("An iframe which has both allow-scripts and allow-same-origin...")
  // และทำให้ google.script.run ผูกกับปุ่ม onclick ไม่สำเร็จ (ReferenceError: xxx is not defined)
  // ค่าเริ่มต้นของ Apps Script (ไม่ต้องตั้ง XFrameOptionsMode เลย) ทำงานถูกต้องและปลอดภัยกว่า
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Trip Logger Pro')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** ใช้สำหรับ include ไฟล์ HTML ย่อยเข้าไฟล์หลัก (เผื่อแยกไฟล์ในอนาคต) */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


// ─────────────────────────────────────────────────────────────────────────
// 2. ตั้งค่าชีตอัตโนมัติ — สร้างชีต+หัวตารางให้ถ้ายังไม่มี (กันปัญหา
//    "บันทึกแล้วไม่ขึ้นในชีต" ที่เกิดจาก schema ไม่ตรงกัน)
// ─────────────────────────────────────────────────────────────────────────
function ensureSheetsExist() {
  const ss = _getSpreadsheet();

  function ensure(name, header) {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.appendRow(header);
      sheet.setFrozenRows(1);
    } else if (sheet.getLastRow() === 0) {
      sheet.appendRow(header);
      sheet.setFrozenRows(1);
    }
    return sheet;
  }

  ensure(SHEET_SITES, SITES_HEADER);
  ensure(SHEET_HOTELS, HOTELS_HEADER);
  ensure(SHEET_FUEL, FUEL_HEADER);

  let counterSheet = ss.getSheetByName(SHEET_COUNTER);
  if (!counterSheet) {
    counterSheet = ss.insertSheet(SHEET_COUNTER);
    counterSheet.appendRow(["LastId"]);
    counterSheet.appendRow([0]);
  }

  // ลบชีตเริ่มต้น "Sheet1" ถ้ายังไม่มีข้อมูลและมีชีตอื่นแล้ว (ทำความสะอาด)
  const sheet1 = ss.getSheetByName("Sheet1") || ss.getSheetByName("แผ่นงาน1");
  if (sheet1 && sheet1.getLastRow() === 0 && ss.getSheets().length > 1) {
    ss.deleteSheet(sheet1);
  }
}

/** สร้าง SiteId ใหม่แบบ thread-safe (กัน id ชนกันถ้ามีคนกดพร้อมกัน) */
function getNextSiteId() {
  // เดิมใช้ LockService.getScriptLock() เพื่อกัน SiteId ชนกัน แต่พบว่าเป็นสาเหตุ
  // ที่ทำให้ createSite() ค้างเงียบในบาง deployment (โดยเฉพาะเมื่อมีหลาย user
  // เรียกพร้อมกัน หรือ user ไม่ใช่เจ้าของสคริปต์) — LockService อาจค้างจน
  // execution timeout โดยไม่ trigger withFailureHandler ฝั่ง frontend เลย
  //
  // เปลี่ยนมาใช้ Utilities.getUuid() ของ Apps Script แทน ซึ่งรับประกันว่าไม่ซ้ำ
  // กัน 100% (มาตรฐาน UUID v4) และทำงานทันทีไม่ต้องรอ lock จากที่ไหนเลย
  // ตัด prefix ให้สั้นลงเหลือ 8 ตัวอักษรแรกพอเพื่อให้ SiteId อ่านง่ายขึ้น
  // (โอกาสชนกันของ 8 ตัวอักษรแรกจาก UUID ยังต่ำมากในทางปฏิบัติสำหรับงานนี้)
  const uuid = Utilities.getUuid().replace(/-/g, "").substring(0, 10).toUpperCase();
  return "S" + uuid;
}


// ─────────────────────────────────────────────────────────────────────────
// 3. ฟังก์ชันช่วยอ่าน/เขียนชีตแบบ row-object (กัน index หลุดเวลาคอลัมน์เปลี่ยน)
// ─────────────────────────────────────────────────────────────────────────
function _sheetToObjects(sheetName) {
  const ss = _getSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const values = sheet.getDataRange().getValues();
  const header = values[0];
  const rows = values.slice(1);
  return rows.map((row, idx) => {
    const obj = { _row: idx + 2 }; // เลขแถวจริงใน sheet (1-indexed, ข้าม header)
    header.forEach((col, i) => { obj[col] = row[i]; });
    return obj;
  });
}

function _findRowIndex(sheetName, matchFn) {
  const objs = _sheetToObjects(sheetName);
  const found = objs.find(matchFn);
  return found ? found._row : -1;
}


// ─────────────────────────────────────────────────────────────────────────
// 4. ไซต์งาน — สร้าง / รายการ / โหลดรายละเอียด
// ─────────────────────────────────────────────────────────────────────────
function createSite(siteName) {
  if (!siteName || !siteName.trim()) {
    throw new Error("กรุณาระบุชื่อไซต์งาน");
  }
  ensureSheetsExist();
  const ss = _getSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_SITES);

  const siteId = getNextSiteId();
  const now = new Date();
  sheet.appendRow([siteId, siteName.trim(), "draft", now, "", 0, 0, "", ""]);

  return { siteId: siteId, siteName: siteName.trim() };
}

/** รายการไซต์ทั้งหมด เรียงใหม่สุดก่อน สำหรับหน้าแรก */
function listSites() {
  ensureSheetsExist();
  const sites = _sheetToObjects(SHEET_SITES);
  const hotels = _sheetToObjects(SHEET_HOTELS);
  const fuel = _sheetToObjects(SHEET_FUEL);

  return sites
    .map(s => {
      const hCount = hotels.filter(h => h.SiteId === s.SiteId && h.ImgUrl).length;
      const fCount = fuel.filter(f => f.SiteId === s.SiteId && f.BillUrl).length;
      return {
        siteId: s.SiteId,
        siteName: s.SiteName,
        status: s.Status,
        createdAt: s.CreatedAt ? new Date(s.CreatedAt).toISOString() : "",
        confirmedAt: s.ConfirmedAt ? new Date(s.ConfirmedAt).toISOString() : "",
        hotelCount: hCount,
        fuelCount: fCount,
      };
    })
    .sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
}

/** โหลดข้อมูลไซต์เต็มรูปแบบ (สำหรับเข้าหน้ากรอกฟอร์ม) */
function getSiteDetail(siteId) {
  ensureSheetsExist();
  const sites = _sheetToObjects(SHEET_SITES);
  const site = sites.find(s => s.SiteId === siteId);
  if (!site) throw new Error("ไม่พบไซต์งานนี้");

  const hotels = _sheetToObjects(SHEET_HOTELS).filter(h => h.SiteId === siteId);
  const fuel = _sheetToObjects(SHEET_FUEL).filter(f => f.SiteId === siteId);

  // จัดเป็นโครงสร้าง nested ให้ frontend ใช้งานง่าย
  const hotelMap = {};
  for (let h = 1; h <= HOTEL_COUNT; h++) {
    hotelMap[h] = {};
    for (let i = 1; i <= HOTEL_ITEM_COUNT; i++) {
      hotelMap[h][i] = { imgUrl: "", desc: "" };
    }
  }
  hotels.forEach(row => {
    const h = Number(row.HotelNo), i = Number(row.ItemNo);
    if (hotelMap[h] && hotelMap[h][i]) {
      hotelMap[h][i] = { imgUrl: row.ImgUrl || "", desc: row.Desc || "" };
    }
  });

  const fuelMap = {};
  for (let n = 1; n <= FUEL_COUNT; n++) {
    fuelMap[n] = { billUrl: "", preUrl: "", postUrl: "", date: "", province: "" };
  }
  fuel.forEach(row => {
    const n = Number(row.FuelNo);
    if (fuelMap[n]) {
      fuelMap[n] = {
        billUrl: row.BillUrl || "",
        preUrl: row.PreUrl || "",
        postUrl: row.PostUrl || "",
        date: row.Date ? Utilities.formatDate(new Date(row.Date), Session.getScriptTimeZone(), "yyyy-MM-dd") : "",
        province: row.Province || "",
      };
    }
  });

  return {
    siteId: site.SiteId,
    siteName: site.SiteName,
    status: site.Status,
    startMile: Number(site.StartMile) || 0,
    endMile: Number(site.EndMile) || 0,
    carImgUrl: site.CarImgUrl || "",
    reportUrl: site.ReportUrl || "",
    hotels: hotelMap,
    fuel: fuelMap,
  };
}


// ─────────────────────────────────────────────────────────────────────────
// 5. บันทึกแต่ละหัวข้อ (โรงแรม / รถ / น้ำมัน) — แยกอิสระจากกัน
// ─────────────────────────────────────────────────────────────────────────

/** บันทึกรูป+คำอธิบายของโรงแรม 1 ช่อง */
function saveHotelItem(siteId, hotelNo, itemNo, imgUrl, desc) {
  _assertSiteEditable(siteId);
  const ss = _getSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_HOTELS);

  const rowIdx = _findRowIndex(SHEET_HOTELS, r =>
    r.SiteId === siteId && Number(r.HotelNo) === Number(hotelNo) && Number(r.ItemNo) === Number(itemNo));

  if (rowIdx > -1) {
    sheet.getRange(rowIdx, 1, 1, HOTELS_HEADER.length).setValues(
      [[siteId, hotelNo, itemNo, imgUrl, desc]]);
  } else {
    sheet.appendRow([siteId, hotelNo, itemNo, imgUrl, desc]);
  }
  return { ok: true };
}

/** บันทึกข้อมูลรถ (ไมล์เริ่ม/จบ + รูปรถ) */
function saveCarInfo(siteId, startMile, endMile, carImgUrl) {
  _assertSiteEditable(siteId);
  const rowIdx = _findRowIndex(SHEET_SITES, r => r.SiteId === siteId);
  if (rowIdx === -1) throw new Error("ไม่พบไซต์งานนี้");

  const ss = _getSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_SITES);
  // อัปเดตเฉพาะคอลัมน์ StartMile(F) EndMile(G) CarImgUrl(H) — คอลัมน์ 6,7,8
  sheet.getRange(rowIdx, 6, 1, 3).setValues([[startMile, endMile, carImgUrl || ""]]);
  return { ok: true };
}

/** บันทึกข้อมูลการเติมน้ำมัน 1 ครั้ง */
function saveFuelItem(siteId, fuelNo, billUrl, preUrl, postUrl, dateStr, province) {
  _assertSiteEditable(siteId);
  const ss = _getSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_FUEL);

  const rowIdx = _findRowIndex(SHEET_FUEL, r =>
    r.SiteId === siteId && Number(r.FuelNo) === Number(fuelNo));

  const dateVal = dateStr ? new Date(dateStr) : "";

  if (rowIdx > -1) {
    sheet.getRange(rowIdx, 1, 1, FUEL_HEADER.length).setValues(
      [[siteId, fuelNo, billUrl, preUrl, postUrl, dateVal, province]]);
  } else {
    sheet.appendRow([siteId, fuelNo, billUrl, preUrl, postUrl, dateVal, province]);
  }
  return { ok: true };
}

function _assertSiteEditable(siteId) {
  const sites = _sheetToObjects(SHEET_SITES);
  const site = sites.find(s => s.SiteId === siteId);
  if (!site) throw new Error("ไม่พบไซต์งานนี้");
  // หมายเหตุ: เดิมเคยห้ามแก้ไขหลังกดยืนยัน (confirmed) แต่ตอนนี้อนุญาตให้
  // แก้ไข/เพิ่มข้อมูลได้แม้ยืนยันไปแล้ว — กดยืนยันซ้ำได้เรื่อยๆ เพื่อสร้าง
  // รายงานเวอร์ชันใหม่ทุกครั้งที่มีการอัปเดตข้อมูล
}


// ─────────────────────────────────────────────────────────────────────────
// 6. ยืนยันไซต์ (จบทริป) — ต้องมีข้อมูลรถ+ไมล์อย่างน้อยก่อนยืนยันได้
// ─────────────────────────────────────────────────────────────────────────
function confirmSite(siteId) {
  const rowIdx = _findRowIndex(SHEET_SITES, r => r.SiteId === siteId);
  if (rowIdx === -1) throw new Error("ไม่พบไซต์งานนี้");

  const sites = _sheetToObjects(SHEET_SITES);
  const site = sites.find(s => s.SiteId === siteId);

  // เงื่อนไข: ต้องมีข้อมูลรถ+ไมล์อย่างน้อยก่อนยืนยันได้ (ตรวจทุกครั้งที่กด
  // แม้จะเคยยืนยันไปแล้วก่อนหน้า เผื่อมีคนลบข้อมูลรถออกไปทีหลัง)
  const startMile = Number(site.StartMile) || 0;
  const endMile = Number(site.EndMile) || 0;
  const hasCarImg = !!site.CarImgUrl;
  if (!(startMile > 0 || endMile > 0 || hasCarImg)) {
    throw new Error("กรุณากรอกข้อมูลรถ (เลขไมล์เริ่มต้น/จบ หรือรูปรถ) อย่างน้อย 1 อย่างก่อนยืนยันไซต์งาน");
  }

  const ss = _getSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_SITES);
  sheet.getRange(rowIdx, 3).setValue("confirmed");        // Status
  sheet.getRange(rowIdx, 5).setValue(new Date());          // ConfirmedAt

  // ── สร้างรายงาน Word + บันทึกลิงก์ลงชีต + ส่งอีเมล ──
  // ทำทุกครั้งที่กดยืนยัน (ไม่ใช่แค่ครั้งแรก) เพื่อให้ได้รายงานเวอร์ชันล่าสุด
  // เสมอ แม้จะเคยยืนยันไปแล้วและมีการแก้ไขข้อมูลเพิ่มทีหลัง
  const report = generateReport(siteId);

  // บันทึก URL ของรายงานล่าสุดไว้ในคอลัมน์ ReportUrl ของชีต Sites
  const reportUrlCol = SITES_HEADER.indexOf("ReportUrl") + 1; // 1-indexed
  sheet.getRange(rowIdx, reportUrlCol).setValue(report.viewUrl);

  let emailSent = false;
  let emailError = "";
  try {
    sendReportEmail(site.SiteName, report.file);
    emailSent = true;
  } catch (e) {
    emailError = e.message;
  }

  return {
    ok: true,
    reportUrl: report.viewUrl,
    downloadUrl: report.downloadUrl,
    emailSent: emailSent,
    emailError: emailError,
  };
}

/** ส่งอีเมลแจ้งรายงานทริป พร้อมแนบไฟล์ Word ที่สร้างไว้ */
function sendReportEmail(siteName, docxFile) {
  const subject = "Trip Report — " + siteName;
  const body =
    "ไฟล์รายงานทริปงาน \"" + siteName + "\" ถูกสร้าง/อัปเดตแล้ว\n\n" +
    "เวลา: " + new Date().toString() + "\n\n" +
    "ดูไฟล์แนบในอีเมลนี้ หรือเข้าถึงผ่าน Google Sheet ที่เก็บข้อมูล " +
    "(คอลัมน์ ReportUrl ในชีต Sites)";

  MailApp.sendEmail({
    to: EMAIL_RECIPIENT,
    subject: subject,
    body: body,
    attachments: [docxFile.getAs(MimeType.MICROSOFT_WORD)],
  });
}


// ─────────────────────────────────────────────────────────────────────────
// 7. อัปโหลดไฟล์ไป Google Drive
// ─────────────────────────────────────────────────────────────────────────
function uploadFile(base64Data, fileName) {
  try {
    const folder = DriveApp.getFolderById(FOLDER_ID);
    const contentType = base64Data.substring(5, base64Data.indexOf(';'));
    const bytes = Utilities.base64Decode(base64Data.split(',')[1]);
    const blob = Utilities.newBlob(bytes, contentType, fileName);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return "https://drive.google.com/uc?export=view&id=" + file.getId();
  } catch (e) {
    throw new Error("อัปโหลดรูปไม่สำเร็จ: " + e.toString());
  }
}


// ─────────────────────────────────────────────────────────────────────────
// 8. สร้างรายงาน Google Docs แล้วคืนลิงก์ดาวน์โหลดเป็น .docx
// ─────────────────────────────────────────────────────────────────────────
function generateReport(siteId) {
  const detail = getSiteDetail(siteId);

  const doc = DocumentApp.create("Trip_Report_" + detail.siteName + "_" + detail.siteId);
  const body = doc.getBody();
  body.setMarginTop(36).setMarginBottom(36).setMarginLeft(50).setMarginRight(50);

  // ── หัวเรื่อง ──
  const title = body.appendParagraph("Trip Report — " + detail.siteName);
  title.setHeading(DocumentApp.ParagraphHeading.TITLE);

  const distance = Math.max(0, detail.endMile - detail.startMile);
  body.appendParagraph(
    "เลขไมล์เริ่มต้น: " + detail.startMile +
    "   เลขไมล์หลังจบ: " + detail.endMile +
    "   ระยะทางรวม: " + distance + " กม."
  );

  if (detail.carImgUrl) {
    _appendImageFromUrl(body, detail.carImgUrl, "รูปรถ", 300);
  }
  body.appendHorizontalRule();

  // ── ส่วนที่ 1: โรงแรม (เรียงซ้าย→ขวา ทีละแถว พร้อมคำอธิบาย) ──
  const hHeading = body.appendParagraph("รูปภาพโรงแรม");
  hHeading.setHeading(DocumentApp.ParagraphHeading.HEADING1);

  let anyHotel = false;
  for (let h = 1; h <= HOTEL_COUNT; h++) {
    const items = [];
    for (let i = 1; i <= HOTEL_ITEM_COUNT; i++) {
      const it = detail.hotels[h][i];
      if (it.imgUrl) items.push(it);
    }
    if (items.length === 0) continue;
    anyHotel = true;

    body.appendParagraph("โรงแรมที่ " + h).setHeading(DocumentApp.ParagraphHeading.HEADING2);

    // ตาราง 2 แถว: แถวบนรูป เรียงซ้าย→ขวา / แถวล่างคำอธิบายของแต่ละรูป
    const table = body.appendTable();
    const imgRow = table.appendTableRow();
    const descRow = table.appendTableRow();
    items.forEach(it => {
      const imgCell = imgRow.appendTableCell();
      _appendImageToCell(imgCell, it.imgUrl, 150);
      const descCell = descRow.appendTableCell();
      descCell.appendParagraph(it.desc || "-");
    });
    body.appendParagraph(""); // เว้นบรรทัด
  }
  if (!anyHotel) {
    body.appendParagraph("(ยังไม่มีรูปโรงแรมที่บันทึกแล้ว)");
  }
  body.appendHorizontalRule();

  // ── ส่วนที่ 2: น้ำมัน (ใบเสร็จใหญ่ซ้าย / ไมล์ก่อน-หลังขวา) ──
  const fHeading = body.appendParagraph("บันทึกการเติมน้ำมัน");
  fHeading.setHeading(DocumentApp.ParagraphHeading.HEADING1);

  let anyFuel = false;
  for (let n = 1; n <= FUEL_COUNT; n++) {
    const f = detail.fuel[n];
    if (!f.billUrl && !f.preUrl && !f.postUrl) continue;
    anyFuel = true;

    body.appendParagraph("การเติมครั้งที่ " + n).setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph("วันที่: " + (f.date || "-") + "   จังหวัด: " + (f.province || "-"));

    const table = body.appendTable();
    const row = table.appendTableRow();

    const leftCell = row.appendTableCell();
    if (f.billUrl) {
      _appendImageToCell(leftCell, f.billUrl, 260); // ใบเสร็จ รูปใหญ่
    } else {
      leftCell.appendParagraph("(ไม่มีรูปใบเสร็จ)");
    }

    const rightCell = row.appendTableCell();
    rightCell.appendParagraph("ไมล์ก่อนเติม:");
    if (f.preUrl) _appendImageToCell(rightCell, f.preUrl, 130);
    rightCell.appendParagraph("ไมล์หลังเติม:");
    if (f.postUrl) _appendImageToCell(rightCell, f.postUrl, 130);

    body.appendParagraph("");
  }
  if (!anyFuel) {
    body.appendParagraph("(ยังไม่มีรายการเติมน้ำมันที่บันทึกแล้ว)");
  }

  doc.saveAndClose();

  // แปลงเป็น .docx แล้วคืน URL ดาวน์โหลด
  const docFile = DriveApp.getFileById(doc.getId());
  const docxBlob = docFile.getAs("application/vnd.openxmlformats-officedocument.wordprocessingml.document");
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const docxFile = folder.createFile(docxBlob).setName("Trip_Report_" + detail.siteName + ".docx");
  docxFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // ลบไฟล์ Google Docs ต้นฉบับทิ้ง (เหลือแค่ .docx)
  DriveApp.getFileById(doc.getId()).setTrashed(true);

  return {
    downloadUrl: "https://drive.google.com/uc?export=download&id=" + docxFile.getId(),
    viewUrl: "https://drive.google.com/file/d/" + docxFile.getId() + "/view",
    fileId: docxFile.getId(),
    file: docxFile,       // ใช้แนบไฟล์ส่งอีเมลโดยตรง (ไม่ต้อง fetch ซ้ำ)
    siteName: detail.siteName,
  };
}

/** ดึงรูปจาก URL (Drive) มาแปะใน body ของเอกสาร พร้อม caption */
function _appendImageFromUrl(body, url, caption, widthPx) {
  try {
    const blob = _fetchImageBlob(url);
    const img = body.appendImage(blob);
    _resizeImage(img, widthPx);
    if (caption) body.appendParagraph(caption);
  } catch (e) {
    body.appendParagraph("(ไม่สามารถโหลดรูป: " + caption + ")");
  }
}

/** ดึงรูปจาก URL มาแปะในเซลล์ตาราง */
function _appendImageToCell(cell, url, widthPx) {
  try {
    const blob = _fetchImageBlob(url);
    const img = cell.appendImage(blob);
    _resizeImage(img, widthPx);
  } catch (e) {
    cell.appendParagraph("(โหลดรูปไม่สำเร็จ)");
  }
}

function _resizeImage(img, widthPx) {
  const ratio = img.getHeight() / img.getWidth();
  img.setWidth(widthPx);
  img.setHeight(Math.round(widthPx * ratio));
}

/** แปลง Drive view-URL กลับเป็น file id แล้วดึง blob ตรงๆ (เร็ว/เสถียรกว่า fetch URL ภายนอก) */
function _fetchImageBlob(url) {
  const m = url.match(/[-\w]{25,}/); // ดึง Drive file id จาก URL
  if (m) {
    const file = DriveApp.getFileById(m[0]);
    return file.getBlob();
  }
  // เผื่อกรณีไม่ใช่ Drive URL — fetch ตรงๆ
  const response = UrlFetchApp.fetch(url);
  return response.getBlob();
}
