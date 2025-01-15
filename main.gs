/**
 * 執行處理 demo sheet 的聽眾評價功能
 */
function executeCreateFeedbackSheets() {
  try {
    createFeedbackSheets();
    Logger.log('回饋表建立完成');
  } catch (error) {
    Logger.log('執行建立回饋表時發生錯誤：' + error.toString());
  }
}

/**
 * 執行處理 feedback sheet 的 NPS 功能
 */
function executeNPSProcess() {
  try {
    processNPSData();
    Logger.log('NPS 處理完成');
  } catch (error) {
    Logger.log('執行 NPS 處理時發生錯誤：' + error.toString());
  }
} 