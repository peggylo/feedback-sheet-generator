function checkDuplicateFeedback() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const feedbackSheet = ss.getSheetByName("feedback");
    if (!feedbackSheet) {
      throw new Error('找不到 feedback sheet');
    }

    // 取得所有資料
    const lastRow = feedbackSheet.getLastRow();
    const data = feedbackSheet.getRange(2, 1, lastRow - 1, 2).getValues(); // 取得所有講者的資料（A和B欄）

    // 處理每一位講者的資料
    data.forEach((row, index) => {
      const currentRow = index + 2; // 實際的列數
      const speakerName = row[0];   // A欄：講者名稱
      const speakerUrl = row[1];    // B欄：試算表網址

      try {
        // 開啟講者的試算表
        const speakerSS = SpreadsheetApp.openByUrl(speakerUrl);
        const listSheet = speakerSS.getSheetByName("list");
        if (!listSheet) {
          throw new Error('找不到 list sheet');
        }

        // 取得回饋欄位（第5欄）的所有資料
        const lastRowInList = listSheet.getLastRow();
        const feedbackRange = listSheet.getRange(2, 5, lastRowInList - 1, 1);
        const feedbackData = feedbackRange.getValues();

        // 檢查重複（排除空白）
        const duplicates = findDuplicates(feedbackData);
        
        // 根據結果更新 check 欄位
        const checkResult = duplicates.length > 0 
          ? `重複：${duplicates.join('；')}` 
          : 'done';
        
        feedbackSheet.getRange(currentRow, 18).setValue(checkResult);  // R欄（check欄位）
        
        Logger.log(`檢查完成 ${speakerName}：${checkResult}`);

      } catch (error) {
        Logger.log(`處理 ${speakerName} 的資料時發生錯誤：${error.toString()}`);
        feedbackSheet.getRange(currentRow, 18).setValue(`錯誤：${error.toString()}`);
      }
    });

    Logger.log('所有講者的回饋檢查完成');

  } catch (error) {
    Logger.log('檢查重複回饋時發生錯誤：' + error.toString());
    throw error;
  }
}

function findDuplicates(feedbackData) {
  const duplicateGroups = [];
  const contentMap = new Map();

  feedbackData.forEach((row, index) => {
    const content = row[0].toString().trim();
    if (content) {
      if (!contentMap.has(content)) {
        contentMap.set(content, []);
      }
      contentMap.get(content).push(index + 2);
    }
  });

  contentMap.forEach((rows, content) => {
    if (rows.length > 1) {
      duplicateGroups.push(`第${rows.join('列、第')}列`);
    }
  });

  return duplicateGroups;
} 