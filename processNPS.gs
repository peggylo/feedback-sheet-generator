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
    
    // 找到所需欄位的索引
    const npsColumnIndex = headers.indexOf("NPS") + 1;
    const statusColumnIndex = headers.indexOf("status") + 1;
    const maxColumnIndex = headers.indexOf("收穫度Max，整天這場是我心中NO1") + 1;
    const learnedLotsColumnIndex = headers.indexOf("學到非常多新東西") + 1;
    const normalColumnIndex = headers.indexOf("普通") + 1;
    const noLearnColumnIndex = headers.indexOf("沒有學到新東西") + 1;
    const totalColumnIndex = headers.indexOf("回填問卷人數") + 1;
    
    if (npsColumnIndex === 0) {
      throw new Error('找不到 NPS 欄位');
    }
    if (statusColumnIndex === 0) {
      throw new Error('找不到 status 欄位');
    }
    
    // 處理簡志祥的資料（第 2 列）
    const row = 2;
    const speakerName = feedbackSheet.getRange("A" + row).getValue();
    const speakerSpreadsheetUrl = feedbackSheet.getRange("B" + row).getValue();
    
    // 取得各項數據
    const promotersCount = feedbackSheet.getRange(row, maxColumnIndex).getValue() + 
                          feedbackSheet.getRange(row, learnedLotsColumnIndex).getValue();
    const detractorsCount = feedbackSheet.getRange(row, normalColumnIndex).getValue() + 
                           feedbackSheet.getRange(row, noLearnColumnIndex).getValue();
    const total = feedbackSheet.getRange(row, totalColumnIndex).getValue();
    
    // 計算百分比
    const promotersPercent = Math.round((promotersCount / total) * 100 * 10) / 10; // 四捨五入到小數點一位
    const detractorsPercent = Math.round((detractorsCount / total) * 100 * 10) / 10;
    const npsScore = Math.round(promotersPercent - detractorsPercent);
    
    // 取得 feedback sheet 中的 NPS 值進行比對
    const feedbackNPS = feedbackSheet.getRange(row, npsColumnIndex).getValue();
    
    // 比對計算出的 NPS 和 feedback sheet 中的 NPS
    let status = '';
    if (npsScore === feedbackNPS) {
      status = 'done';
    } else {
      status = `計算值(${npsScore}) 與 表格值(${feedbackNPS}) 不符`;
    }
    
    // 更新 status 欄位
    feedbackSheet.getRange(row, statusColumnIndex).setValue(status);
    
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
    newSheet.getRange("A2").setValue(npsScore);
    
    // 根據 NPS 分數選擇文字模板
    const messageTemplate = npsScore >= 50 ? 
      `根據 AI DAY 參加者回饋，您的聽眾中高度推薦者占 ${promotersPercent}%；改進需求者占 ${detractorsPercent}%，NPS 為：${promotersPercent}% − ${detractorsPercent}% = ${npsScore}。多數參加者對您的分享推薦度極高，從 NPS 來看您的內容對參加者具有高度價值，真心感謝！\n\n（也說明，NPS 對主辦單位來說並不是在評價講者，比較是在理解聽眾需求、他們的疑問或期待有無被解決；衷心謝謝老師準備了第一線老師們需要的內容！🙏）` :
      `根據 AI DAY 參加者回饋，您的聽眾中高度推薦者占 ${promotersPercent}%；改進需求者占 ${detractorsPercent}%，NPS 為：${promotersPercent}% − ${detractorsPercent}% = ${npsScore}。多數參加者皆推薦您的分享，從 NPS 來看您的內容對參加者很有價值，真心感謝！\n\n（也說明，NPS 對主辦單位來說並不是在評價講者，比較是在理解聽眾需求、他們的疑問或期待有無被解決；衷心謝謝老師準備了第一線老師們需要的內容！🙏）`;
    
    // 填入回饋文字
    newSheet.getRange("A6").setValue(messageTemplate);
    
    Logger.log(`已完成處理 ${speakerName} 的 NPS 資料，狀態：${status}`);
  } catch (error) {
    Logger.log('處理 NPS 資料時發生錯誤：' + error.toString());
    throw error;
  }
} 