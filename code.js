// 取得存在 Google 後端的設定，不需寫死在程式碼中

const CONFIG = {
  SHEET_NAME: 'Registrations',
  CLASS_SHEET_NAME: 'Classes', 
  FOLDER_ID: '1JbSwWckxj5ttCLc_2DgJ-a8Q8qnd6u3m' // 確保這裡沒有被清空
};


/**
 * 1. 載入唯一的 SPA 母網頁
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('臺北市勞動檢查處 - 職訓課程報名網')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * 2. 獲取課程清單與即時名額
 */
function getAvailableClasses() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.CLASS_SHEET_NAME);
  const data = sheet.getDataRange().getDisplayValues(); 
  const classes = [];
  
  for (let i = 2; i < data.length; i++) {
    classes.push({
      id: data[i][0],        // A 欄：課程編號
      name: data[i][1],      // B 欄：課程名稱
      category: data[i][2],  // C 欄：課程類別 
      date: data[i][3],      // D 欄：開課日期 
      total: data[i][4],     // E 欄：總名額   
      registered: data[i][5],// F 欄：已報名人數 
      remaining: data[i][6]  // G 欄：剩餘名額  
    });
  }
  return classes;
}

/**
 * 3. 處理報名資料並寫入資料庫
 * [更新]：在 D 欄旁新增 E 欄「課程日期」，整體結構調整為 28 欄
 */
function submitRegistration(formData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    const folder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
    
    if (!sheet) throw new Error("找不到報名資料工作表(Registrations)");
    
    const timestamp = new Date();
    const uid = "REG-" + Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "yyyyMMdd-HHmmss") + "-" + Math.floor(Math.random() * 1000);
    const status = "待審核"; 
    
    // 處理身分證檔案上傳 (Base64)
    const filePrefix = `${formData.name}_${formData.idNumber}`;
    let idFrontUrl = formData.idFrontBase64 ? saveBase64FileToDrive(folder, formData.idFrontBase64, `${filePrefix}_正面`) : "";
    let idBackUrl = formData.idBackBase64 ? saveBase64FileToDrive(folder, formData.idBackBase64, `${filePrefix}_反面`) : "";
    let idOtherUrl = formData.idOtherBase64 ? saveBase64FileToDrive(folder, formData.idOtherBase64, `${filePrefix}_其他`) : ""; // 處理其他證件

    // 準備寫入報名資料表 (精準對應 A ~ AC 欄共 29 欄)
    const rowData = [
      uid,                          // A: 報名編號
      timestamp,                    // B: 報名時間
      status,                       // C: 審核狀態
      formData.course,              // D: 報名課程
      formData.courseDate,          // E: 課程日期 (新欄位)
      formData.basicQualification,  // F: 基本資格
      formData.priorityQualification || "", // G: 優先錄取資格
      formData.name,                // H: 姓名
      formData.gender,              // I: 性別
      formData.birthday,            // J: 生日
      formData.idNumber,            // K: 身分證字號
      formData.phone,               // L: 手機號碼
      formData.email,               // M: 電子郵件
      formData.education,           // N: 最高學歷
      formData.school,              // O: 畢業學校
      formData.zipCode,             // P: 戶籍郵遞區號
      formData.address,             // Q: 戶籍地址
      formData.contactAddress,      // R: 聯絡地址
      formData.emergencyName,       // S: 緊急聯絡人姓名
      formData.emergencyPhone,      // T: 緊急聯絡人電話
      formData.companyName || "",   // U: 任職公司名稱
      formData.companyTaxId || "",  // V: 公司統編
      formData.companyAddress || "",// W: 公司地址
      formData.companyPhone || "",  // X: 公司電話
      formData.companyContact || "",// Y: 公司教訓聯絡人
      idFrontUrl,                   // Z: 身分證正面連結
      idBackUrl,                    // AA: 身分證反面連結
      idOtherUrl,                   // AB: 其他證件連結
      ""                            // AC: 審核備註/補件說明
    ];
    
    // 寫入報名試算表
    sheet.appendRow(rowData);

    return { success: true, message: "報名成功！", uid: uid };
  } catch (error) {
    console.error("後端執行錯誤：", error);
    return { success: false, message: "錯誤：" + error.message };
  }
}

/**
 * 4. 檔案轉存 Google Drive 輔助函式
 */
function saveBase64FileToDrive(folder, dataUrl, fileName) {
  let mimeType = 'image/jpeg'; 
  let base64Data = dataUrl;
  if (dataUrl.indexOf('data:') === 0) {
    const parts = dataUrl.split(',');
    mimeType = parts[0].split(';')[0].split(':')[1];
    base64Data = parts[1];
  }
  let extension = '';
  if (mimeType.includes('jpeg') || mimeType.includes('jpg')) extension = '.jpg';
  else if (mimeType.includes('png')) extension = '.png';
  else if (mimeType.includes('pdf')) extension = '.pdf';

  const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, fileName + extension);
  const file = folder.createFile(blob);
  return file.getUrl();
}

/**
 * 新增：報名資料查詢功能
 */
function lookupRegistration(name, last5Id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  
  // 遍歷資料 (從第 1 列開始，因為第 0 列是標題)
  for (let i = 1; i < data.length; i++) {
    let rowName = data[i][7];   // H 欄：姓名
    let rowId = String(data[i][10]); // K 欄：身分證字號
    
    // 檢查姓名是否符合，且身分證字號結尾是否包含輸入的後五碼
    if (rowName === name && rowId.slice(-5) === last5Id) {
      return { 
        success: true, 
        data: { 
          status: data[i][2], // C 欄：狀態
          course: data[i][3], // D 欄：課程
          date: data[i][4]    // E 欄：日期
        } 
      };
    }
  }
  return { success: false, message: "查無資料，請確認姓名與身分證後五碼是否正確。" };
}