function createFeedbackSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const speakerNames = headers.slice(5); // 假設第6欄之後是講者的欄位
  
  const urls = []; // 用於儲存新試算表名稱和網址的陣列
  
  speakerNames.forEach((speaker, index) => {
    if (!speaker) return; // 跳過空的講者欄位
    
    const speakerIndex = 5 + index; // 講者欄位的索引
    const filteredData = data.filter((row, i) => {
      // 篩選出 D 欄等於講者名字的列，並排除標題列和空白行
      return i === 0 || (row[3] === speaker && row.some(cell => cell)); 
    });
    
    if (filteredData.length <= 1) return; // 如果沒有有效資料列，略過
    
    // 新試算表名稱
    const newSheetName = `2024 教學創新 AI DAY 問卷回饋--${speaker}`;
    const newSheet = SpreadsheetApp.create(newSheetName);
    const newSs = newSheet.getActiveSheet();
    
    // 新標題列
    const newHeaders = [
      "教師姓名",
      "縣市",
      "來自學校",
      `【${speaker}】分享給你的收穫程度(選項:收穫度Max,學到很多,有學到,普通,沒學到新東西)`,
      `給講者【${speaker}】分享內容的建議或心得`
    ];
    newSs.appendRow(newHeaders);
    
    // 將篩選的資料填入新試算表
    filteredData.slice(1).forEach(row => {
      const newRow = [
        row[0], // 您的姓名 → 教師姓名
        row[2], // 縣市
        row[1], // 學校 → 來自學校
        row[4], // 教學導入Round1-1
        row[speakerIndex] // 給該講者的回饋內容
      ];
      newSs.appendRow(newRow);
    });

    // 紀錄新試算表網址
    const newSheetUrl = newSheet.getUrl();
    urls.push([newSheetName, newSheetUrl]);
  });
  
  // 打印試算表網址到 Logger
  urls.forEach(urlInfo => {
    Logger.log(`試算表名稱: ${urlInfo[0]} | 試算表網址: ${urlInfo[1]}`);
  });

  // 如果需要，可以選擇將 URL 寫回到原始試算表的一個分頁
  const urlSheet = ss.getSheetByName("生成試算表網址") || ss.insertSheet("生成試算表網址");
  urlSheet.clear(); // 清空原有內容
  urlSheet.appendRow(["試算表名稱", "試算表網址"]);
  urls.forEach(urlInfo => {
    urlSheet.appendRow(urlInfo);
  });
}
