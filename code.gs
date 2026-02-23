/**
 * STORE LOG SYSTEM - PROFESSIONAL VERSION 2026
 * ระบบบันทึกและตรวจสอบข้อมูลการจัดสินค้า
 */
/** * ฟังก์ชันหลักสำหรับแสดงหน้าเว็บ 
 * ไม่มีการแก้ไขโครงสร้างเดิม 
 */
function doGet(e) {
  if (e.parameter.page == 'dashboard') {
    return HtmlService.createTemplateFromFile('DashboardNew').evaluate()
      .setTitle('DASHBOARD - TOP 5')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('STORE LOGGING SYSTEM')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getScriptUrl() { return ScriptApp.getService().getUrl(); }

function getFormData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const listSheet = ss.getSheetByName("List");
  const staffSheet = ss.getSheetByName("พนักงาน");
  
  const staffData = staffSheet.getRange("A2:C" + staffSheet.getLastRow()).getValues();
  const staffCodes = {};
  staffData.forEach(r => {
    if (r[0]) staffCodes[r[0]] = r[2] || r[0]; 
  });

  return {
    pickers: staffSheet.getRange("A2:A" + staffSheet.getLastRow()).getValues().flat().filter(String),
    checkers: staffSheet.getRange("B2:B" + staffSheet.getLastRow()).getValues().flat().filter(String),
    staffCodes: staffCodes, 
    products: listSheet.getRange("A2:D" + listSheet.getLastRow()).getValues().map(r => ({ code: String(r[2]), desc: String(r[3]) })).filter(r => r.desc),
    customers: [...new Set(listSheet.getRange("A2:A" + listSheet.getLastRow()).getValues().flat().filter(String))]
  };
}

/** * บันทึกข้อมูลการพิมพ์เอกสาร (Log_การพิมพ์) 
 * แก้ไข: เปลี่ยนการเก็บ Base64 เป็น Link Google Drive ตามที่คุยกัน 
 */
/** * บันทึกข้อมูลการพิมพ์เอกสาร และรูปภาพลง Google Drive
 */
function savePrintLog(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName("Log_การพิมพ์");
  const FOLDER_ID = "1WWMrr4HoaFNLqPO3_GPl-JFtfjDPdMZI"; 
  
  try {
    const folder = DriveApp.getFolderById(FOLDER_ID);
    
    const saveImageToDrive = (base64Data, fileName) => {
      // ตรวจสอบว่ามีข้อมูลรูปภาพส่งมาจริงหรือไม่
      if (!base64Data || typeof base64Data !== 'string' || !base64Data.includes(",")) return "ไม่มีรูปภาพ";
      
      try {
        const splitData = base64Data.split(",");
        const contentType = splitData[0].split(":")[1].split(";")[0];
        const bytes = Utilities.base64Decode(splitData[1]);
        const blob = Utilities.newBlob(bytes, contentType, fileName);
        const file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        return file.getUrl();
      } catch(e) { return "Error: " + e.toString(); }
    };

    const billNo = data.bill || "Unknow-Bill";
    const imgLinks = [
      saveImageToDrive(data.img1, `W1_${billNo}_${Date.now()}.jpg`),
      saveImageToDrive(data.img2, `W2_${billNo}_${Date.now()}.jpg`),
      saveImageToDrive(data.img3, `R1_${billNo}_${Date.now()}.jpg`),
      saveImageToDrive(data.img4, `R2_${billNo}_${Date.now()}.jpg`)
    ];

    if (!logSheet) {
      logSheet = ss.insertSheet("Log_การพิมพ์");
      logSheet.appendRow(["วันเวลา", "เลขที่บิล", "ร้านค้า", "ผู้จัดทำ", "เหตุผล", "ค่าใช้จ่าย", "Link ผิด1", "Link ผิด2", "Link ถูก1", "Link ถูก2"]);
    }
    
    logSheet.appendRow([
      new Date(), billNo, data.customer || "", data.creator || "", 
      data.reason || "", data.cost || 0,
      imgLinks[0], imgLinks[1], imgLinks[2], imgLinks[3]
    ]);
    
    return true;
  } catch(e) {
    return "Error: " + e.toString();
  }
}

function addData(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("ข้อมูลการจัดสินค้า");
    
    // ตรวจสอบว่ามีชีทนี้อยู่จริง
    if (!sheet) return "ไม่พบชีท 'ข้อมูลการจัดสินค้า'";

    // รวมรหัสบิล
    const fullBillNo = (data.billPrefix || "") + (data.billPeriod || "") + (data.billNum || "");
    
    // เตรียมข้อมูล 16 คอลัมน์ (A ถึง P) ตามโครงสร้างชีทของคุณ
    const rowData = [
      new Date(),                // A: วันที่รับแจ้ง
      data.company || "",        // B: บริษัท
      fullBillNo || "",          // C: เลขที่บิล
      data.errorType || "",      // D: ประเภท
      data.wrongID || "",        // E: รหัสสินค้าที่ผิด
      data.wrongName || "",      // F: รายการที่ผิด
      Number(data.qtyWrong) || 0, // G: QTY. (ผิด)
      data.rightID || "",        // H: รหัสสินค้าที่ถูกต้อง
      data.rightName || "",      // I: รายการที่ถูกต้อง
      Number(data.qtyRight) || 0, // J: QTY. (ถูก)
      data.staffPick || "",      // K: ผู้จัดสินค้า
      data.staffCheck || "",     // L: เช็คเกอร์
      data.status || "",         // M: สถานะ
      "",                        // N: พนักงานจัดส่ง 1 (ว่างไว้สำหรับ Log การพิมพ์)
      "",                        // O: พนักงานจัดส่ง 2 (ว่างไว้สำหรับ Log การพิมพ์)
      ""                         // P: ผู้ออกเอกสาร/อื่นๆ
    ];
    
    // บันทึกข้อมูล
    sheet.appendRow(rowData);
    
    return "บันทึกเรียบร้อย";
  } catch(e) {
    Logger.log("Error in addData: " + e.toString());
    return "เกิดข้อผิดพลาดในการบันทึก: " + e.message;
  }
}

/** * ดึงข้อมูลทั้งหมด 
 * แก้ไข: กรองแถวว่าง (Filter) ออก เพื่อไม่ให้ Dashboard ดึงข้อมูลขยะจากสูตรมาแสดง 
 */
function getAllData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("ข้อมูลการจัดสินค้า");
  if (!sheet || sheet.getLastRow() < 2) return [];
  
  const rawData = sheet.getRange(2, 1, sheet.getLastRow() - 1, 16).getValues();
  
  const filteredData = rawData.filter(r => {
    const billCell = String(r[2]).trim();
    return billCell !== "" && billCell !== null && billCell !== "undefined";
  });
  
  return filteredData.map(r => ({
    date: Utilities.formatDate(r[0] instanceof Date ? r[0] : new Date(), "GMT+7", "dd/MM/yyyy"),
    month: Utilities.formatDate(r[0] instanceof Date ? r[0] : new Date(), "GMT+7", "MM"),
    customer: String(r[1] || ""),
    bill: String(r[2] || ""),
    type: String(r[3] || ""),
    wName: String(r[5] || ""), 
    wQty: r[6] || 0,
    rName: String(r[8] || ""), 
    rQty: r[9] || 0,
    pick: String(r[10] || ""), 
    check: String(r[11] || ""),
    status: String(r[12] || ""),
    colN: String(r[13] || ""), 
    colO: String(r[14] || ""),
    colP: String(r[15] || "")
  })).reverse();
}

function getMonthlyWorkload() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("จำนวนบิลต่อเดือน");
  if (!sheet) return Array(12).fill(0);
  const data = sheet.getRange(2, 2, 1, 12).getValues()[0]; 
  return data.map(val => val || 0);
}

function getDataCount() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("ข้อมูลการจัดสินค้า");
  if (!sheet || sheet.getLastRow() < 2) return 0;
  
  const vals = sheet.getRange(2, 3, sheet.getLastRow() - 1, 1).getValues();
  return vals.filter(r => {
    const bill = String(r[0]).trim();
    return bill !== "" && bill !== "null" && bill !== "undefined";
  }).length;
}
