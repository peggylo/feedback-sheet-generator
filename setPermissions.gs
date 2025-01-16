function setViewerPermissions() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const feedbackSheet = ss.getSheetByName("feedback");
    if (!feedbackSheet) {
      throw new Error('找不到 feedback sheet');
    }
    
    // 取得表頭資料
    const headers = feedbackSheet.getRange(1, 1, 1, feedbackSheet.getLastColumn()).getValues()[0];
    
    // 找到所需欄位的索引
    const emailColumnIndex = headers.indexOf("email") + 1;
    const emailCheckColumnIndex = headers.indexOf("email check") + 1;
    const spreadsheetUrlColumnIndex = headers.indexOf("LINK") + 1;
    
    if (emailColumnIndex === 0) {
      throw new Error('找不到 email 欄位');
    }
    if (emailCheckColumnIndex === 0) {
      throw new Error('找不到 email check 欄位');
    }
    if (spreadsheetUrlColumnIndex === 0) {
      throw new Error('找不到 LINK 欄位');
    }

    // 取得所有資料（跳過表頭）
    const dataRange = feedbackSheet.getRange(2, 1, feedbackSheet.getLastRow() - 1, feedbackSheet.getLastColumn());
    const data = dataRange.getValues();

    // 處理每一列資料
    data.forEach((row, index) => {
      const currentRow = index + 2;
      const emails = row[emailColumnIndex - 1];
      const spreadsheetUrl = row[spreadsheetUrlColumnIndex - 1];
      
      if (emails && spreadsheetUrl) {
        try {
          Logger.log(`處理第 ${currentRow} 列的權限設定`);
          const result = setPermissionsForOneSheet(spreadsheetUrl, emails);
          feedbackSheet.getRange(currentRow, emailCheckColumnIndex).setValue(result);
          Logger.log(`第 ${currentRow} 列處理結果: ${result}`);
        } catch (error) {
          Logger.log(`第 ${currentRow} 列處理失敗: ${error.toString()}`);
          feedbackSheet.getRange(currentRow, emailCheckColumnIndex)
            .setValue(`失敗：${error.toString()}`);
        }
      }
    });

    Logger.log('權限設定完成');
  } catch (error) {
    Logger.log('設定權限時發生錯誤：' + error.toString());
    throw error;
  }
}

function setPermissionsForOneSheet(spreadsheetUrl, emailsString) {
  // 如果 emails 是空的，直接返回
  if (!emailsString) {
    return "沒有 email 資料";
  }

  const emails = emailsString.split(';').map(email => email.trim());
  const ss = SpreadsheetApp.openByUrl(spreadsheetUrl);
  
  let successCount = 0;
  let failCount = 0;
  let errorMessages = [];

  // 處理每個 email
  emails.forEach(email => {
    if (email) {
      try {
        ss.addViewer(email);
        successCount++;
      } catch (error) {
        failCount++;
        errorMessages.push(`${email}: ${error.message}`);
      }
    }
  });

  // 回傳結果
  if (failCount === 0) {
    return "done";
  } else if (successCount > 0) {
    return `部分完成：成功 ${successCount} 個，失敗 ${failCount} 個`;
  } else {
    return `失敗：${errorMessages.join('; ')}`;
  }
} 