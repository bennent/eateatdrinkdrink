// Mr. Wish 希望好茶 點餐系統 — Google Apps Script 後端
// 部署方式：擴充功能 → Apps Script → 貼上本檔 → 部署 → New deployment → Web app

const SHEET_NAME = '訂單紀錄';
const HEADERS = ['時間戳記', '姓名', '飲料名稱', '冰量', '甜度', '價格'];

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
  }
  return sheet;
}

// 前端送出訂單：POST { user, time, items: [{name, ice, sugar, price}, ...] }
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const { user, time, items } = payload;

    if (!user || !time || !Array.isArray(items) || items.length === 0) {
      return jsonOutput_({ status: 'error', message: '缺少必要欄位' });
    }

    const sheet = getSheet_();
    items.forEach(item => {
      sheet.appendRow([time, user, item.name, item.ice, item.sugar, item.price]);
    });

    return jsonOutput_({ status: 'ok' });
  } catch (err) {
    return jsonOutput_({ status: 'error', message: err.message });
  }
}

// 統計頁讀取所有訂單列：GET，回傳扁平陣列，每列一個飲料品項
function doGet(e) {
  try {
    const sheet = getSheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return jsonOutput_([]);
    }

    const values = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
    const rows = values.map(r => ({
      time: String(r[0]),
      user: String(r[1]),
      name: String(r[2]),
      ice: String(r[3]),
      sugar: String(r[4]),
      price: Number(r[5]) || 0,
    }));

    return jsonOutput_(rows);
  } catch (err) {
    return jsonOutput_({ status: 'error', message: err.message });
  }
}

function jsonOutput_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
