// 準1~撤収日
// const FIXED_SHIFT_LIST_URL = "https://docs.google.com/spreadsheets/d/XXXXXXXXXX/edit";
// 準0
const FIXED_SHIFT_LIST_URL = "https://docs.google.com/spreadsheets/d/XXXXXXXXXX/edit";
const FIXED_SHIFT_LIST_COL_ROLE = {
  isPrepared: 9,
  url: 10
};

const FIXED_SHIFT_COL_ROLE = {
  shift: 1,
  name: 2,
  firstData: 3,
  lastData: 19
};

const FIXED_SHIFT_ROW_ROLE = {
  hourMinute: 2,
  firstData: 3
}

const FIXED_SHIFT_COLOR = "#18499d";

const PREPARED = "自動化準備完了";

function formatFixedShiftData() {
  const listSs = SpreadsheetApp.openByUrl(FIXED_SHIFT_LIST_URL);
  const listRichTextData = listSs.getDataRange().getRichTextValues();

  let dbData = [];
  let errorData = initializeDbData();
  errorData[0].push("Reason");

  for (const row of listRichTextData) {
    if (row[FIXED_SHIFT_LIST_COL_ROLE.isPrepared - 1].getText() !== PREPARED) continue;

    const ssUrl = row[FIXED_SHIFT_LIST_COL_ROLE.url - 1].getLinkUrl();

    const orgSs = SpreadsheetApp.openByUrl(ssUrl);

    const orgSheets = orgSs.getSheets();

    let data = [];
    let errorLogs = [];

    for (const sheet of orgSheets) {

      const sheetName = sheet.getName();
      if (["記入方法", "【当日人員】準備日0日目 10/31(木)"].includes(sheetName)) continue;

      const fixedShiftData = sheet.getDataRange().getValues();

      const [records, errorLog] = extractFixedShiftTo2DArray(sheetName, fixedShiftData);

      data.push(...records);
      errorLogs.push(...errorLog);
    }

    dbData.push(...data);
    errorData.push(...errorLogs);
  }

  input2DArrayToDataSheet2(dbData);
  saveErrorLogToSheet(errorData);
}

// シフトデータからDBに記入する用の2次元配列を作成
function extractFixedShiftTo2DArray(sheetName, shiftData) {
  let records = [];
  let errorLog = [];

  for (let row = FIXED_SHIFT_ROW_ROLE.firstData - 1; row < shiftData.length; row++) {
    const shiftName = shiftData[row][FIXED_SHIFT_COL_ROLE.shift - 1];
    if (!shiftName || shiftName == "合計") continue;

    for (let col = FIXED_SHIFT_COL_ROLE.firstData - 1; col < FIXED_SHIFT_COL_ROLE.lastData; col++) {
      if (!shiftData[row][col]) continue;
      const record = makeRecord2(sheetName, row, col, shiftData);
      shiftData[row][col] == 1
        ? records.push(record)
        : errorLog.push([...record, "fixed shift was not 1"]);
    }

  }

  return [records, errorLog]; // DBに入力する2次元配列を返す
}

// 1行のシフトデータを展開してレコードを作成
function makeRecord2(sheetName, row, col, values) {
  let record = [];

  // 時間の書式を整形 8:15～ ---> 8時 15-
  const hourMinute = values[FIXED_SHIFT_ROW_ROLE.hourMinute - 1][col].replace("~", "").trim();
  let [hour, minute] = hourMinute.split(":");
  hour = hour + "時";
  minute = minute + "-";

  // 日付の書式を整形 準備日11/1(金) --> 準備日 11/1(金)    ←スペースを追加した
  const date = sheetName.replace("【固定】準備日0日目 10/31(木)", "準備日0日目");

  record[DB_COL_ROLE.hour - 1] = hour;
  record[DB_COL_ROLE.minute - 1] = minute;
  record[DB_COL_ROLE.date - 1] = date;
  record[DB_COL_ROLE.shift - 1] = values[row][FIXED_SHIFT_COL_ROLE.shift - 1];
  record[DB_COL_ROLE.name - 1] = values[row][FIXED_SHIFT_COL_ROLE.name - 1];
  record[DB_COL_ROLE.bgColor - 1] = FIXED_SHIFT_COLOR;

  return record; // 1行のレコードを返す
}

// データシートに2次元配列を入力
function input2DArrayToDataSheet2(dbData) {
  const dataSs = SpreadsheetApp.openByUrl(DATA_SS_URL);
  const dataSheet = dataSs.getSheetByName(DATA_SHEET_NAME) || dataSs.insertSheet(DATA_SHEET_NAME);

  // 既存データを削除
  const lastRow = dataSheet.getLastRow();

  // データがある場合、入力
  if (dbData.length > 0) {
    dataSheet.getRange(lastRow, 1, dbData.length, dbData[0].length).setValues(dbData);
  }
}

