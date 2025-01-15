function createFeedbackSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const speakerNames = headers.slice(5); // 假設第6欄之後是講者的欄位
  
  const urls = []; // 用於儲存新試算表名稱和網址的陣列
  const feedbackOrder = [
    "收穫度Max，整天這場是我心中NO1",
    "學到非常多新東西",
    "有學到新東西",
    "普通",
    "沒有學到新東西"
  ];

  speakerNames.forEach((speaker, index) => {
    if (!speaker) return; // 跳過空的講者欄位

    const speakerIndex = 5 + index; // 講者欄位的索引
    const filteredData = data.filter((row, i) => {
      return i === 0 || (row[3] === speaker && row.some(cell => cell));
    });

    if (filteredData.length <= 1) return;

    const speakerDisplayName = speaker.slice(-2) + "老師";
    const newSheetName = `2024 教學創新 AI DAY 問卷回饋--${speaker}`;
    const newSheet = SpreadsheetApp.create(newSheetName);
    const newSs = newSheet.getActiveSheet();

    newSs.setName("list");

    const newHeaders = [
      "教師姓名",
      "來自學校",
      "縣市",
      `【${speakerDisplayName}】分享給你的收穫程度(選項:收穫度Max,學到很多,有學到,普通,沒學到新東西)`,
      `給講者【${speakerDisplayName}】分享內容的建議或心得`
    ];
    newSs.appendRow(newHeaders);

    // 設定第一列標題格式
    const headerRange = newSs.getRange(1, 1, 1, 5);

    // 設定標題文字顏色為白色
    headerRange.setFontColor("#FFFFFF");

    // 設定背景顏色
    headerRange.getCell(1, 1).setBackground("#6C6C6C"); // A列淺灰色
    headerRange.getCell(1, 2).setBackground("#6C6C6C"); // B列淺灰色
    headerRange.getCell(1, 3).setBackground("#6C6C6C"); // C列淺灰色
    headerRange.getCell(1, 4).setBackground("#272727"); // D列深灰色
    headerRange.getCell(1, 5).setBackground("#272727"); // E列深灰色

    // 設定標題文字置中並加粗
    headerRange.setHorizontalAlignment("center").setFontWeight("bold");

    let feedbackRows = [];
    let feedbackCount = {
      "收穫度Max，整天這場是我心中NO1": 0,
      "學到非常多新東西": 0,
      "有學到新東西": 0,
      "普通": 0,
      "沒有學到新東西": 0
    };

    filteredData.slice(1).forEach(row => {
      let teacherName = row[0];
      let school = row[2];
      let city = row[1];
      const feedback = row[4];
      const suggestion = row[speakerIndex];

      if (feedback === "普通" || feedback === "沒有學到新東西") {
        teacherName = "○○○";
        school = "○○○○";
      }

      if (feedbackCount.hasOwnProperty(feedback)) {
        feedbackCount[feedback]++;
      }

      feedbackRows.push([teacherName, school, city, feedback, suggestion]);
    });

    feedbackRows.sort((a, b) => {
      return feedbackOrder.indexOf(a[3]) - feedbackOrder.indexOf(b[3]);
    });

    // 將排序後的資料填入新試算表
    let rowCount = 0;
    feedbackRows.forEach(row => {
      newSs.appendRow(row);
      rowCount++;
    });

    // 設定所有儲存格的格式
    const allRange = newSs.getRange(1, 1, rowCount + 1, 5);  // +1 是因為包含標題列
    
    // 設定垂直置中
    allRange.setVerticalAlignment("middle");
    
    // 設定前四欄水平置中
    const centerRange = newSs.getRange(1, 1, rowCount + 1, 4);
    centerRange.setHorizontalAlignment("center");

    Logger.log(`為講者 ${speaker} 建立了 ${rowCount} 列資料（不含標題列）`);

    Logger.log(`講者 ${speaker} 的收穫度統計：`);
    for (const [key, value] of Object.entries(feedbackCount)) {
      Logger.log(`${key}: ${value}`);
    }

    // 設定第五欄自動換行（除標題列外）
    const wrapRange = newSs.getRange(2, 5, rowCount, 1); // 從第2行開始的第五欄
    wrapRange.setWrap(true);

    // 凍結第一列和前三欄
    newSs.setFrozenRows(1);
    newSs.setFrozenColumns(3);

    // 刪除多餘的欄位和列
    newSs.deleteColumns(6, newSs.getMaxColumns() - 5);  // 刪除 F 欄後的所有欄位
    
    const lastRow = newSs.getLastRow();
    const totalRows = newSs.getMaxRows();
    if (lastRow < totalRows) {
      newSs.deleteRows(lastRow + 1, totalRows - lastRow);  // 刪除多餘的列
    }

    const newSheetUrl = newSheet.getUrl();

    // 將資訊填入 feedback 分頁
    const feedbackSheet = ss.getSheetByName("feedback");
    const feedbackLastRow = feedbackSheet.getLastRow();  // 改名避免變數重複宣告
    
    // 準備要填入的資料
    const feedbackData = [
      speaker,  // 講者姓名
      newSheetUrl,  // 試算表網址
      feedbackCount["收穫度Max，整天這場是我心中NO1"],
      feedbackCount["學到非常多新東西"],
      feedbackCount["有學到新東西"],
      feedbackCount["普通"],
      feedbackCount["沒有學到新東西"],
      Object.values(feedbackCount).reduce((sum, count) => sum + count, 0),
      // 計算 NPS
      calculateNPS(feedbackCount)
    ];
    
    // 在最後一列之後新增資料
    feedbackSheet.getRange(feedbackLastRow + 1, 1, 1, feedbackData.length).setValues([feedbackData]);

    urls.push([newSheetName, newSheetUrl]);
  });

  urls.forEach(urlInfo => {
    Logger.log(`試算表名稱: ${urlInfo[0]} | 試算表網址: ${urlInfo[1]}`);
  });
}

// 新增 NPS 計算函數
function calculateNPS(feedbackCount) {
  const total = Object.values(feedbackCount).reduce((sum, count) => sum + count, 0);
  if (total === 0) return 0;

  // 推廣者：給 4-5 分的人數（收穫度Max 和 學到非常多）
  const promoters = feedbackCount["收穫度Max，整天這場是我心中NO1"] + 
                   feedbackCount["學到非常多新東西"];
  
  // 被動者：給 3 分的人數（有學到新東西）
  const passives = feedbackCount["有學到新東西"];
  
  // 貶低者：給 1-2 分的人數（普通和沒學到）
  const detractors = feedbackCount["普通"] + 
                    feedbackCount["沒有學到新東西"];

  // 計算百分比並四捨五入到整數
  const promotersPercent = (promoters / total) * 100;
  const detractorsPercent = (detractors / total) * 100;
  
  // NPS = 推廣者百分比 - 貶低者百分比
  return Math.round(promotersPercent - detractorsPercent);
}