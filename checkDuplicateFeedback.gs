function checkDuplicateFeedback() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const feedbackSheet = ss.getSheetByName("feedback");
    if (!feedbackSheet) {
      throw new Error('找不到 feedback sheet');
    }

    // 先只處理第一筆資料
    const firstRow = 2;  // 從第二列開始（跳過標題）
    const speakerUrl = feedbackSheet.getRange(firstRow, 2).getValue();  // B欄位的試算表網址
    
    // 開啟講者的試算表
    const speakerSS = SpreadsheetApp.openByUrl(speakerUrl);
    const listSheet = speakerSS.getSheetByName("list");
    if (!listSheet) {
      throw new Error('找不到 list sheet');
    }

    // 取得回饋欄位（第5欄）的所有資料
    const lastRow = listSheet.getLastRow();
    const feedbackRange = listSheet.getRange(2, 5, lastRow - 1, 1);  // 跳過標題列
    const feedbackData = feedbackRange.getValues();

    // 檢查重複（排除空白）
    const duplicates = findDuplicates(feedbackData);
    
    // 根據結果更新 check 欄位
    const checkResult = duplicates.length > 0 
      ? `重複：${duplicates.join('；')}` 
      : 'done';
    
    feedbackSheet.getRange(firstRow, 18).setValue(checkResult);  // R欄（check欄位）
    
    Logger.log(`檢查完成：${checkResult}`);

  } catch (error) {
    Logger.log('檢查重複回饋時發生錯誤：' + error.toString());
    throw error;
  }
}

function findDuplicates(feedbackData) {
  const duplicateGroups = [];
  const contentMap = new Map();  // 用來記錄內容和它們出現的列數

  // 先記錄每個非空白內容出現的列數
  feedbackData.forEach((row, index) => {
    const content = row[0].toString().trim();
    if (content) {  // 排除空白內容
      if (!contentMap.has(content)) {
        contentMap.set(content, []);
      }
      contentMap.get(content).push(index + 2);  // +2 是因為要算上標題列，且 index 從 0 開始
    }
  });

  // 找出重複的群組
  contentMap.forEach((rows, content) => {
    if (rows.length > 1) {  // 如果同樣的內容出現超過一次
      duplicateGroups.push(`第${rows.join('列、第')}列`);
    }
  });

  return duplicateGroups;
} 