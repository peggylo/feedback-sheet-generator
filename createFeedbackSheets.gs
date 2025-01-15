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

    // 修改新試算表 sheet 名稱為 "list"
    newSs.setName("list");

    // 新標題列
    const newHeaders = [
      "教師姓名",
      "來自學校", // 第二欄
      "縣市",     // 第三欄
      `【${speaker}】分享給你的收穫程度(選項:收穫度Max,學到很多,有學到,普通,沒學到新東西)`,
      `給講者【${speaker}】分享內容的建議或心得`
    ];
    newSs.appendRow(newHeaders);

    let rowCount = 0;
    let feedbackCount = {
      "收穫度Max，整天這場是我心中NO1": 0,
      "學到非常多新東西": 0,
      "有學到新東西": 0,
      "普通": 0,
      "沒有學到新東西": 0
    };

    // 將篩選的資料填入新試算表
    filteredData.slice(1).forEach(row => {
      let teacherName = row[0];
      let school = row[1];
      const feedback = row[4]; // 教學導入Round1-1
      const suggestion = row[speakerIndex]; // 給該講者的回饋內容

      // 修改教師姓名和學校資料，若回饋為「普通」或「沒有學到新東西」
      if (feedback === "普通" || feedback === "沒有學到新東西") {
        teacherName = "○○○";
        school = "○○○○";
      }

      // 統計各種回饋的數量
      if (feedbackCount.hasOwnProperty(feedback)) {
        feedbackCount[feedback]++;
      }

      const newRow = [teacherName, school, row[2], feedback, suggestion];
      newSs.appendRow(newRow);
      rowCount++;
    });

    // 打印分別建立了幾列資料
    Logger.log(`為講者 ${speaker} 建立了 ${rowCount} 列資料（不含標題列）`);

    // 打印統計數據
    Logger.log(`講者 ${speaker} 的收穫度統計：`);
    for (const [key, value] of Object.entries(feedbackCount)) {
      Logger.log(`${key}: ${value}`);
    }

    // 凍結第一列和前三欄
    newSs.setFrozenRows(1);
    newSs.setFrozenColumns(3);

    // 設定格式：第一列文字置中
    const headerRange = newSs.getRange(1, 1, 1, 5);
    headerRange.setHorizontalAlignment("center");

    // 前四欄文字置中
    const centerRange = newSs.getRange(2, 1, newSs.getLastRow() - 1, 4);
    centerRange.setHorizontalAlignment("center");

    // 第五欄文字自動換行
    const wrapRange = newSs.getRange(2, 5, newSs.getLastRow() - 1, 1);
    wrapRange.setWrap(true);

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
