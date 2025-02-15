const ERROR_SHEET_NAME = "ErrorLog"; // エラーデータを保存するシート名

function restoreSheetsFromData() {
  const dataSs = SpreadsheetApp.openByUrl(DATA_SS_URL);
  const dataSheet = dataSs.getSheetByName(DATA_SHEET_NAME);
  const dataRange = dataSheet.getDataRange();
  const data = dataRange.getValues();
  data.shift();//見出し行を削除
  const gcSs = SpreadsheetApp.openByUrl(GANTT_CHART_SS_URL);
  // シート名と見出しの位置を取得
  const sheetsRowColIds = getSheetHeaderRowCols(gcSs);

  // シートごとのデータをまとめるオブジェクトを初期化
  const shiftDataMap = initializeShiftData(gcSs);
  let shiftErrorLog = initializeDbData();
  shiftErrorLog[0].push('Reason');

  // データシートの各行に対して処理を実行
  data.forEach(row => {
    processShiftRow(row, sheetsRowColIds, shiftDataMap, shiftErrorLog);
  });

  // 各シートにデータを書き込む
  applyShiftDataToSheets(shiftDataMap);

  // エラーログをスプレッドシートに書き込む
  saveErrorLogToSheet(shiftErrorLog);
}

// 各行ごとのシフトデータを2次元配列に変換する関数
function processShiftRow(row, sheetsRowColIds, shiftDataMap, shiftErrorLog) {
  const sheetName = row[DB_COL_ROLE.department - 1];
  const rowColIds = sheetsRowColIds[sheetName];
  const rowId = row[DB_COL_ROLE.gmail - 1].trim() + '|' + row[DB_COL_ROLE.date - 1].trim();
  const colId = row[DB_COL_ROLE.hour - 1].trim() + '|' + row[DB_COL_ROLE.minute - 1].trim();

  if (rowColIds) {
    let targetRow = rowColIds.rowIds[rowId]; // メールアドレスと日付の見出しの組み合わせ
    let targetCol = rowColIds.colIds[colId]; // 時間と分の見出しの組み合わせ

    // targetRow または targetCol が undefined の場合はエラーデータとして保存
    if (!targetRow || !targetCol) {
      const reason = !targetRow ? 'Row not found' : 'Column not found';
      shiftErrorLog.push([...row, reason]);
    } else {
      targetRow = targetRow - GC_ROW_ROLE.firstData + 1;
      targetCol = targetCol - GC_COL_ROLE.firstData + 1;
      const shiftData = shiftDataMap[sheetName];
      shiftData.values[targetRow][targetCol] = row[DB_COL_ROLE.shift - 1]; // セルの値を再配置
      shiftData.backgrounds[targetRow][targetCol] = row[DB_COL_ROLE.bgColor - 1]; // セルの背景色を再配置
    }
  } else {
    shiftErrorLog.push([...row, "sheet not found"]);
  }
}

// スプレッドシートにシフトデータを書き込む関数
function applyShiftDataToSheets(shiftDataMap) {
  const gcSs = SpreadsheetApp.openByUrl(GANTT_CHART_SS_URL);
  for (const sheetName in shiftDataMap) {
    const sheet = gcSs.getSheetByName(sheetName) || gcSs.insertSheet(sheetName);
    const shiftData = shiftDataMap[sheetName];
    if (!shiftData.values.length) continue;
    let range = sheet.getRange(GC_ROW_ROLE.firstData, GC_COL_ROLE.firstData, shiftData.values.length, shiftData.values[0].length);
    range.breakApart();
    range.setValues(shiftData.values);
    range = sheet.getRange(GC_ROW_ROLE.firstData, GC_COL_ROLE.firstData, shiftData.backgrounds.length, shiftData.backgrounds[0].length);
    range.setBackgrounds(shiftData.backgrounds);
    mergeSameValuesHorizontally(sheet, range);
  }
}

// エラーログをスプレッドシートに保存する関数
function saveErrorLogToSheet(shiftErrorLog) {
  const dataSs = SpreadsheetApp.openByUrl(DATA_SS_URL);
  let errorSheet = dataSs.getSheetByName(ERROR_SHEET_NAME) || dataSs.insertSheet(ERROR_SHEET_NAME);
  errorSheet.getRange(1, 1, shiftErrorLog.length, shiftErrorLog[0].length).setValues(shiftErrorLog);
}

// ssのメールアドレスと日付、時間と分を取得して、それぞれ何行目かを記録して連想配列にする関数
function getSheetHeaderRowCols(spreadsheet) {
  const sheets = spreadsheet.getSheets();
  const headerInfo = {};

  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    const allData = sheet.getDataRange().getValues(); // シート内のすべてのデータを一括で取得

    const gmails = allData.map(row => row[GC_COL_ROLE.gmail - 1]);
    const dates = allData.map(row => row[GC_COL_ROLE.date - 1]);
    fillEmptyValues(gmails);
    const hours = allData[GC_ROW_ROLE.hour - 1];
    const minutes = allData[GC_ROW_ROLE.minute - 1];
    fillEmptyValues(hours);

    const numRows = allData.length;
    const numCols = allData[0].length;

    const gmailAndDate = {};
    const hourAndMinute = {};

    for (let i = GC_ROW_ROLE.firstData - 1; i < numRows; i++) {
      if (!dates[i] || !gmails[i]) continue;
      const key = gmails[i].trim() + '|' + dates[i].trim();
      gmailAndDate[key] = i; // 0-based index
    }

    for (let j = GC_COL_ROLE.firstData - 1; j < numCols; j++) {
      const key = hours[j].trim() + '|' + minutes[j].trim();
      hourAndMinute[key] = j; // 0-based index
    }

    headerInfo[sheetName] = {
      rowIds: gmailAndDate,
      colIds: hourAndMinute
    };
  });

  return headerInfo;
}

// 空白セルがあれば、直近の非空白の値を引き継ぐ関数
function fillEmptyValues(array) {
  let lastNonEmptyValue = null;
  for (let i = 0; i < array.length; i++) {
    array[i] = array[i] ? array[i] : lastNonEmptyValue;
    lastNonEmptyValue = array[i];
  }
}

function initializeShiftData(gcSs) {
  const sheets = gcSs.getSheets();
  const shiftDataMap = {};

  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    const numRows = sheet.getDataRange().getNumRows() - GC_ROW_ROLE.firstData + 1;
    const numCols = sheet.getDataRange().getNumColumns() - GC_COL_ROLE.firstData + 1;

    shiftDataMap[sheetName] = {
      values: Array.from({ length: numRows }, () => Array(numCols).fill('')),  // 空の2次元配列を作成
      backgrounds: Array.from({ length: numRows }, () => Array(numCols).fill('#ffffff')) // 背景色を白に設定
    };
  })

  return shiftDataMap;
}
