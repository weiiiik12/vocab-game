const SPREADSHEET_ID = '1etfiDXLL0_8iiSZxFI5MwTZcA59mvycTtiNJ53OGvio';

// 允許跨網域存取的 Header
function getCorsResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// 1. 處理 GitHub 來的「讀取」請求 (題庫與排行榜)
function doGet(e) {
  const action = e.parameter.action;
  
  if (action === 'getVocabData') {
    const data = getVocabData();
    return getCorsResponse({ success: true, data: data });
  }
  
  if (action === 'getLeaderboard') {
    const data = getLeaderboard();
    return getCorsResponse({ success: true, data: data });
  }
  
  return getCorsResponse({ success: false, message: '未知的指令' });
}

// 2. 處理 GitHub 來的「寫入」請求 (儲存排行榜成績)
function doPost(e) {
  try {
    const postData = JSON.parse(e.postData.contents);
    const { name, profession, difficulty, timeSeconds, errors } = postData;
    
    const success = saveScore(name, profession, difficulty, timeSeconds, errors);
    return getCorsResponse({ success: success });
  } catch (error) {
    return getCorsResponse({ success: false, error: error.toString() });
  }
}

// 3. 讀取全域題庫 (維持原樣)
function getVocabData() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheets = ss.getSheets();
    const result = {}; 
    
    sheets.forEach(sheet => {
      const sheetName = sheet.getName();
      if (sheetName === '排行榜') return;
      
      const data = sheet.getDataRange().getValues();
      if (data.length < 2) return; 
      
      const headers = data[0].map(h => String(h).trim());
      let wIdx = headers.findIndex(h => h.includes('單字') || h.toLowerCase().includes('word'));
      let tIdx = headers.findIndex(h => h.includes('翻譯') || h.includes('中文') || h.toLowerCase().includes('mean'));
      
      if (wIdx === -1) wIdx = 0; 
      if (tIdx === -1) tIdx = 1; 
      
      const wordsArray = [];
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (row.length > Math.max(wIdx, tIdx)) {
          const word = String(row[wIdx]).trim();
          const trans = String(row[tIdx]).trim();
          if (word && trans) wordsArray.push({ w: word, c: trans });
        }
      }
      if (wordsArray.length > 0) result[sheetName] = wordsArray;
    });
    return result; 
  } catch (error) {
    return {};
  }
}

// 4. 儲存通關成績 (維持原樣)
function saveScore(name, profession, difficulty, timeSeconds, errors) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('排行榜');
    if (!sheet) {
      sheet = ss.insertSheet('排行榜');
      sheet.appendRow(['時間戳記', '玩家名稱', '職業', '挑戰難度', '通關時間(秒)', '失誤次數']);
      sheet.getRange("A1:F1").setFontWeight("bold");
    }
    sheet.appendRow([new Date(), name, profession, difficulty, timeSeconds, errors]);
    return true;
  } catch (error) {
    return false;
  }
}

// 5. 讀取排行榜 (維持原樣)
function getLeaderboard() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('排行榜');
    if (!sheet) return {};

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return {};

    const result = {};
    for (let i = 1; i < data.length; i++) {
      const [timestamp, name, profession, difficulty, time, errors] = data[i];
      if (!result[difficulty]) result[difficulty] = [];
      result[difficulty].push({ 
        name: String(name), 
        profession: String(profession), 
        time: Number(time), 
        errors: Number(errors),
        date: new Date(timestamp).toLocaleDateString()
      });
    }

    for (const diff in result) {
      result[diff].sort((a, b) => {
        if (a.time !== b.time) return a.time - b.time;
        return a.errors - b.errors;
      });
      result[diff] = result[diff].slice(0, 10);
    }
    return result;
  } catch (error) {
    return {};
  }
}