function processNPSData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const feedbackSheet = ss.getSheetByName("feedback");
    if (!feedbackSheet) {
      throw new Error('找不到 feedback sheet');
    }
    
    const npsTemplateSheet = ss.getSheetByName("NPS");
    if (!npsTemplateSheet) {
      throw new Error('找不到 NPS 模板 sheet');
    }

    // 取得表頭資料
    const headers = feedbackSheet.getRange(1, 1, 1, feedbackSheet.getLastColumn()).getValues()[0];
    const npsColumnIndex = headers.indexOf("NPS") + 1; // 找到 NPS 欄位的位置（+1 是因為 getRange 從 1 開始）
    
    if (npsColumnIndex === 0) { // 如果找不到（indexOf 回傳 -1，加 1 後為 0）
      throw new Error('找不到 NPS 欄位');
    }
    
    // 處理簡志祥的資料（第 2 列）
    const row = 2;
    const speakerName = feedbackSheet.getRange("A" + row).getValue();
    const speakerSpreadsheetUrl = feedbackSheet.getRange("B" + row).getValue();
    const speakerNPS = feedbackSheet.getRange(row, npsColumnIndex).getValue();
    
    // 開啟講者的 spreadsheet
    const speakerSS = SpreadsheetApp.openByUrl(speakerSpreadsheetUrl);
    
    // 檢查是否已存在 NPS sheet，如果有就刪除
    const existingSheet = speakerSS.getSheetByName("NPS");
    if (existingSheet) {
      speakerSS.deleteSheet(existingSheet);
    }
    
    // 複製 NPS sheet
    const newSheet = npsTemplateSheet.copyTo(speakerSS);
    newSheet.setName("NPS");
    
    // 填入 NPS 值
    newSheet.getRange("A2").setValue(speakerNPS);
    
    Logger.log(`已完成處理 ${speakerName} 的 NPS 資料`);
  } catch (error) {
    Logger.log('處理 NPS 資料時發生錯誤：' + error.toString());
    throw error;
  }
} 