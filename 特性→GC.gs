// 24 特性調査
const CHARACTERISTICS_SS_URL = "https://docs.google.com/spreadsheets/d/XXXXXXXXXX/edit";

const CHAR_SHEET_NAME = "フォームの回答 1";

// 24 特性調査
const CHAR_COL_ROLE = {
  gmail: 3,
  name: 4,
  tel: 9,
  department: 14,
  voice: 24,
  power: 25,
  disease1: 28,
  disease2: 29
};


function setCharacteristics() {
  const gcSs = SpreadsheetApp.openByUrl(GANTT_CHART_SS_URL);
  const sheets = gcSs.getSheets();
  sheets.forEach(sheet => {
    const charSs = SpreadsheetApp.openByUrl(CHARACTERISTICS_SS_URL);
    const charSheet = charSs.getSheetByName(CHAR_SHEET_NAME);
    const charData = charSheet.getDataRange().getValues();
    const returnCols = [
      CHAR_COL_ROLE.tel,
      CHAR_COL_ROLE.voice,
      CHAR_COL_ROLE.power,
      CHAR_COL_ROLE.disease1,
      CHAR_COL_ROLE.disease2
    ];
    const inputCols = [
      GC_COL_ROLE.tel,
      GC_COL_ROLE.voice,
      GC_COL_ROLE.power,
      GC_COL_ROLE.disease1,
      GC_COL_ROLE.disease2
    ];
    searchAndWriteResults(GC_COL_ROLE.gmail, sheet, CHAR_COL_ROLE.gmail, charData, returnCols, inputCols);
  });
}

/**
 * 結合されたセルも正しく認識して値を検索し、戻り値を別の列に記入する関数
 */
function searchAndWriteResults(searchValColIndex, sheet, lookupColIndex, lookupArray, returnColIndices, inputColIndices) {

  // シートの全データ範囲を取得
  let dataRange = sheet.getDataRange();
  let data = dataRange.getValues();  // シート全体のデータを一度に取得
  let mergedRanges = dataRange.getMergedRanges();  // 結合セルの範囲を取得

  // 検索列の結合セルに対応する値を適用
  let mergedLookup = {};

  // 結合セルの範囲のうち、lookupColIndexの列の結合セルのみを処理
  mergedRanges.forEach(range => {
    let topRow = range.getRow();
    let bottomRow = range.getLastRow();
    let column = range.getColumn();

    // 結合セルがlookupColIndexの列に属する場合のみ処理する
    if (column === searchValColIndex) {
      let topCellValue = data[topRow - 1][column - 1];  // 結合セル範囲の一番上の値を取得

      // 結合された行全体に対して値をセット
      for (let i = topRow; i <= bottomRow; i++) {
        mergedLookup[i] = topCellValue;
      }
    }
  });

  // 検索列のすべての値をループ
  let results = [];

  for (let i = 1; i < data.length; i++) {  // 1行目はヘッダーと想定して2行目から始める
    let searchValue = data[i][searchValColIndex - 1];

    // 結合されている場合、結合セルの値を使う
    if (mergedLookup[i + 1]) {  // `i + 1` はスプレッドシートの行番号に合わせるため
      searchValue = mergedLookup[i + 1];
    }

    // customXLookupで検索
    let returnValues = customXLookup(searchValue, lookupArray, lookupColIndex, returnColIndices);

    // 戻り値があれば、結果を保存（なければ空文字を設定）
    if (returnValues !== null) {
      results.push(returnValues);
    } else {
      // 戻り値が見つからなかった場合、空の配列を設定
      results.push(returnColIndices.map(() => ''));
    }
  }

  const getColumn = (arr, colIndex) => arr.map(row => [row[colIndex]]);

  // 戻り値の列に配列として一括で書き込み
  for (let i = 0; i < inputColIndices.length; i++) {
    const resultsColumn = getColumn(results, i);

    // resultsColumnが空でない場合のみsetValuesを実行
    if (resultsColumn.length > 0) {
      sheet.getRange(2, inputColIndices[i], resultsColumn.length, 1).setValues(resultsColumn);
    }
  }
}

/**
 * XLOOKUPのような機能を実装した関数（複数列の戻り値対応）
 * @param {any} searchValue 検索する値
 * @param {Array} lookupArray 検索対象となる2次元配列
 * @param {number} lookupColIndex 検索する列の番号（1から始まる）
 * @param {Array<number>} returnColIndices 戻り値を取得する列の番号の配列（1から始まる）
 * @return {Array} 検索に成功した場合は戻り列の値の配列、見つからない場合はnullを返す
 */
function customXLookup(searchValue, lookupArray, lookupColIndex, returnColIndices) {
  // 検索列のインデックスを0ベースに変換
  let searchCol = lookupColIndex - 1;

  for (let i = 0; i < lookupArray.length; i++) {
    // 検索値が見つかった場合
    if (lookupArray[i][searchCol] === searchValue) {
      // 戻り値を取得する列番号の配列から、対応する値を取得
      let returnValues = returnColIndices.map(colIndex => lookupArray[i][colIndex - 1]);
      return returnValues; // 戻り値として配列を返す
    }
  }

  return null; // 見つからなかった場合はnullを返す
}