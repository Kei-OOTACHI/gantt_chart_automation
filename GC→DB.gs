function gcToDbAllSheets() {
  const gcSs=SpreadsheetApp.openByUrl(GANTT_CHART_SS_URL);
  const sheets = gcSs.getSheets();
  let gcData = [];
  let gcColor = [];
  let dbData = initializeDbData();  // すべてのシートのシフトデータを格納する変数

  sheets.forEach(sheet => {
    const sheetName = sheet.getName();

    const [values, backgrounds] = unmergeGc(sheet);
    gcData.push(...values);


    const records = extractGcTo2DArray(sheetName, values, backgrounds);
    dbData.push(...records);
    gcColor.push(...backgrounds);

    // let targetSheet = dataSs.getSheetByName(sheetName).clear() || dataSs.insertSheet(sheetName);
    // copyUnmergedGc(targetSheet, values, backgrounds, sheet);
    Logger.log(`シート '${sheetName}' 抽出完了`);
  });

  copyUnmergedGcGroupedByDays(gcData, gcColor);

  input2DArrayToDataSheet(dbData);
}

// DB（にこれから記入する2次元配列）を初期化する関数。見出し行とかを設定する
function initializeDbData() {
  allData = [];

  let headerRow = [];
  for (const key in DB_COL_ROLE) {
    headerRow[DB_COL_ROLE[key] - 1] = key;
  }

  allData.unshift(headerRow);
  return allData;
}

// GCのデータを抽出、結合されたセルをすべて分割、DBに転記する関数
function unmergeGc(sourceSheet) {
  const sourceRange = sourceSheet.getDataRange();
  const values = sourceRange.getValues();
  const backgrounds = sourceRange.getBackgrounds();
  const mergedRanges = sourceRange.getMergedRanges();

  // 結合範囲に対して1回だけ値を取得し、それをまとめて反映
  mergedRanges.forEach(mergedRange => {
    const startRow = mergedRange.getRow() - 1;
    const startColumn = mergedRange.getColumn() - 1;
    const rowCount = mergedRange.getNumRows();
    const columnCount = mergedRange.getNumColumns();

    // ループの外で一度だけ値と背景色を取得
    const mergedValue = values[startRow][startColumn];
    const mergedBackground = backgrounds[startRow][startColumn];

    // スライスで範囲を取得し、全セルに値と背景色を適用
    for (let i = startRow; i < startRow + rowCount; i++) {
      values[i].fill(mergedValue, startColumn, startColumn + columnCount);
      backgrounds[i].fill(mergedBackground, startColumn, startColumn + columnCount);
    }
  });

  return [values, backgrounds];
}

// 各シートのデータをDBに転記する関数
function copyUnmergedGc(targetSheet, values, backgrounds, sourceSheet) {
  // 名前と日付が空白の行を削除
  let filteredValues = [];
  let filteredBg = [];

  for (let i = 0; i < values.length; i++) {
    // 名前と日付がどちらも空白でない場合、その行を保持
    if (!(values[i][GC_COL_ROLE.name - 1] === "" && values[i][GC_COL_ROLE.date - 1] === "")) {
      filteredValues.push(values[i]); // データを残す
      filteredBg.push(backgrounds[i]); // 同じ行のデータを残す
    }
  }

  const numRows = filteredValues.length;
  const numColumns = filteredValues[0].length;
  const targetRange = targetSheet.getRange(1, 1, numRows, numColumns);

  // 値と背景色を設定
  targetRange.setValues(filteredValues);
  targetRange.setBackgrounds(filteredBg);

  // 行の高さと列の幅を設定
  const height = sourceSheet.getRowHeight(GC_ROW_ROLE.firstData);
  const width = sourceSheet.getColumnWidth(GC_COL_ROLE.firstData);
  targetSheet.setRowHeightsForced(GC_ROW_ROLE.firstData, numRows - GC_ROW_ROLE.firstData + 1, height);
  targetSheet.setColumnWidths(GC_COL_ROLE.firstData, numColumns - GC_COL_ROLE.firstData + 1, width);
}

// セルを分割した後の各シートのデータ（values）から、1シフト1行のデータを生成する関数
function extractGcTo2DArray(sheetName, values, backgrounds) {
  const numRows = values.length;
  const numColumns = values[0].length;
  const records = []; // 各シートのデータを格納する配列

  // GCのシフトデータが入っているすべてのセルに対して実行
  for (let row = GC_ROW_ROLE.firstData - 1; row < numRows; row++) {
    for (let col = GC_COL_ROLE.firstData - 1; col < numColumns; col++) {

      if (!values[row][col] && backgrounds[row][col] == "#ffffff") continue;// 何もシフトが入ってないところはスキップ
      const record = makeRecord(sheetName, row, col, values, backgrounds);
      records.push(record);
    }
  }

  return records; // 各シートのデータを返す
}

// GCで1つのセルに入っているシフトデータの背景色、シフト名、誰の・何日・何時・何分のシフトか、を1行のデータに展開する関数
function makeRecord(sheetName, row, col, values, backgrounds) {
  let record = [];

  // オプションとなるデータ転記
  for (const key in DB_COL_ROLE) {
    if (["department", "hour", "minute", "bgColor", "shift"].includes(key)) continue;
    record[DB_COL_ROLE[key] - 1] = values[row][GC_COL_ROLE[key] - 1] || "";
  }
  // 必須のデータの転記
  record[DB_COL_ROLE.department - 1] = sheetName;
  record[DB_COL_ROLE.hour - 1] = values[GC_ROW_ROLE.hour - 1][col];
  record[DB_COL_ROLE.minute - 1] = values[GC_ROW_ROLE.minute - 1][col];
  record[DB_COL_ROLE.bgColor - 1] = backgrounds[row][col];
  record[DB_COL_ROLE.shift - 1] = values[row][col];

  return record;
}
function copyUnmergedGcGroupedByDays(gcData, gcColor) {
  const dataSs=SpreadsheetApp.openByUrl(DATA_SS_URL);
  // 日付ごとにデータと背景色を同時にグループ化
  const groupedData = groupDataAndColorsByColumn(gcData, gcColor, GC_COL_ROLE.date - 1);

  // 各日付ごとにシートにデータと背景色を出力
  for (const key in groupedData) {
    let sheetName = key ? key : "日付なし";  // キーが空なら「日付なし」とする
    const { data: sheetData, colors: sheetColors } = groupedData[key]; // データと背景色を取得

    // 対応するシートを取得、なければ新規作成
    let sheet = dataSs.getSheetByName(sheetName);
    if (!sheet) {
      sheet = dataSs.insertSheet(sheetName);
    } else {
      sheet.clear();  // 既存のデータをクリア
    }

    // データの設定
    const numRows = sheetData.length;
    const numColumns = sheetData[0].length;
    const targetRange = sheet.getRange(1, 1, numRows, numColumns);
    targetRange.setValues(sheetData);  // データを書き込む

    // 背景色の設定
    const colorRange = sheet.getRange(1, 1, numRows, numColumns);
    colorRange.setBackgrounds(sheetColors);  // 背景色を書き込む

    sheet.setRowHeightsForced(1, numRows, HEADER_ROW_HEIGHT);
    sheet.setColumnWidths(1, numColumns, TIMESCALE_COLUMN_WIDTH);

    sheet.setColumnWidths(1, GC_COL_ROLE.firstData - 1, HEADER_COLUMN_WIDTH);
    sheet.setColumnWidth(GC_COL_ROLE.date, DATE_COL_WHIDTH);
    sheet.setColumnWidth(GC_COL_ROLE.gender, GENDER_COL_WHIDTH);
    insertTimeScale(sheet);
  }
}

/**
 * データ配列と背景色配列を指定列でグループ化する関数
 * @param {Array} dataArray 2次元配列（データ）
 * @param {Array} colorArray 2次元配列（背景色）
 * @param {number} columnIndex グループ化基準となる列のインデックス (0ベース)
 * @return {Object} グループ化したデータと背景色のオブジェクト
 */
function groupDataAndColorsByColumn(dataArray, colorArray, columnIndex) {
  // グループを保持するオブジェクト
  const groups = {};

  // 2次元配列をループしてグループ化
  for (let i = 0; i < dataArray.length; i++) {
    const key = dataArray[i][columnIndex];  // グループ化のキーとなる値を取得

    if (!groups[key]) {
      groups[key] = { data: [], colors: [] };  // キーがまだ存在しない場合はデータと背景色を保持する配列を作成
    }

    groups[key].data.push(dataArray[i]);  // キーに対応するデータを追加
    groups[key].colors.push(colorArray[i]);  // キーに対応する背景色を追加
  }

  return groups;  // オブジェクトを返す
}


// データシートに、抽出・統合したシフトデータを書き込む関数
function input2DArrayToDataSheet(dbData) {
  const dataSs=SpreadsheetApp.openByUrl(DATA_SS_URL);
  const dataSheet = dataSs.getSheetByName(DATA_SHEET_NAME) || dataSs.insertSheet(DATA_SHEET_NAME);

  // データシートの既存データを削除
  dataSheet.clear();

  // すべてのシートのデータを一括で追加
  if (dbData.length > 0) {
    dataSheet.getRange(1, 1, dbData.length, dbData[0].length).setValues(dbData);
  }
}

function insertTimeScale(sheet) {
  let headerRowData = [];

  for (const key in GC_COL_ROLE) {
    const index = GC_COL_ROLE[key];
    headerRowData[index - 1] = key;
  }

  let times = [];
  for (let i = HOURS.start; i <= HOURS.end; i++) {
    for (let j = 0; j < MINUTES.length; j++) {
      times.push(`${i}:${MINUTES[j]}`);
    }
  }

  headerRowData = headerRowData.slice(0, GC_COL_ROLE.firstData - 1).concat(times);

  sheet.insertRowsBefore(1, 1);
  const range = sheet.getRange(1, 1, 1, headerRowData.length);
  range.setValues([headerRowData]);//先頭行の前に1行追加
  range.setBackground("#ffffff");
  // hideOrFoldCols(sheet);//初回のみ非表示とグループ化の設定
}
