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

    // 取得所有資料（跳過表頭）
    const dataRange = feedbackSheet.getRange(2, 1, feedbackSheet.getLastRow() - 1, feedbackSheet.getLastColumn());
    const data = dataRange.getValues();

    // 只處理 status 為空白的資料
    data.forEach((rowData, index) => {
      const currentRow = index + 2; // 實際的列數
      const status = rowData[statusColumnIndex - 1];
      
      if (status === "") { // 只處理 status 為空白的資料
        try {
          processOneNPSData(feedbackSheet, npsTemplateSheet, rowData, currentRow, {
            nps: npsColumnIndex,
            status: statusColumnIndex,
            max: maxColumnIndex,
            learnedLots: learnedLotsColumnIndex,
            normal: normalColumnIndex,
            noLearn: noLearnColumnIndex,
            total: totalColumnIndex
          });
        } catch (error) {
          Logger.log(`處理第 ${currentRow} 列資料時發生錯誤：${error.toString()}`);
          feedbackSheet.getRange(currentRow, statusColumnIndex).setValue(`處理失敗：${error.toString()}`);
        }
      }
    });

    Logger.log('所有需處理的 NPS 資料處理完成');
  } catch (error) {
    Logger.log('處理 NPS 資料時發生錯誤：' + error.toString());
    throw error;
  }
}

function processOneNPSData(feedbackSheet, npsTemplateSheet, rowData, row, columnIndices) {
  const speakerName = rowData[0];  // A 欄
  const speakerSpreadsheetUrl = rowData[1];  // B 欄
  
  // 修改計算方式
  const promotersCount = rowData[columnIndices.max - 1] + 
                        rowData[columnIndices.learnedLots - 1];
  const detractorsCount = rowData[columnIndices.noLearn - 1]; // 只計算"沒有學到新東西"
  const total = rowData[columnIndices.total - 1];
  
  // 計算百分比
  const promotersPercent = Math.round((promotersCount / total) * 100 * 10) / 10;
  const detractorsPercent = Math.round((detractorsCount / total) * 100 * 10) / 10;
  const npsScore = Math.round(promotersPercent - detractorsPercent);
  
  // 取得 feedback sheet 中的 NPS 值進行比對
  const feedbackNPS = rowData[columnIndices.nps - 1];
  
  // 比對計算出的 NPS 和 feedback sheet 中的 NPS
  let status = '';
  if (npsScore === feedbackNPS) {
    if (npsScore >= 60) {
      status = 'done';
    } else {
      status = 'done (NPS < 60，不建立回饋表)';
    }
  } else {
    status = `計算值(${npsScore}) 與 表格值(${feedbackNPS}) 不符`;
  }
  
  // 更新 status 欄位
  feedbackSheet.getRange(row, columnIndices.status).setValue(status);
  
  // 開啟講者的 spreadsheet
  const speakerSS = SpreadsheetApp.openByUrl(speakerSpreadsheetUrl);
  
  // 檢查是否已存在 NPS sheet
  const existingSheet = speakerSS.getSheetByName("NPS");
  if (existingSheet) {
    speakerSS.deleteSheet(existingSheet);
  }
  
  // 只有 NPS >= 60 才建立新的 NPS sheet
  if (npsScore >= 60) {
    // 複製 NPS sheet
    const newSheet = npsTemplateSheet.copyTo(speakerSS);
    newSheet.setName("NPS");
    
    // 填入 NPS 值
    newSheet.getRange("A2").setValue(npsScore);
    
    // 填入回饋文字
    const messageTemplate = `根據 AI DAY 參加者回饋，您的聽眾中高度推薦與認同者占 ${promotersPercent}%，NPS 為：${promotersPercent}% − ${detractorsPercent}% = ${npsScore}，是當天的 TOP 5 講者！大多數參加者對您的分享高度認同與推薦，從 NPS 來看您的內容對參加者具有高度價值，真心感謝！\n\n（也說明，NPS 對主辦單位來說並不是在評價講者，比較是在理解聽眾需求、他們的疑問或期待有無被解決；衷心謝謝您準備了第一線老師們需要的內容！🙏）`;
    
    newSheet.getRange("A6").setValue(messageTemplate);
  }
  
  Logger.log(`已完成處理 ${speakerName} 的 NPS 資料，分數：${npsScore}，狀態：${status}`);
} 