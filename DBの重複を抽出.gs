function extractDbDuplicates() {
  const dataSs=SpreadsheetApp.openByUrl(DATA_SS_URL);
  const dbSeet = dataSs.getSheetByName(DATA_SHEET_NAME);
  const orgData = dbSeet.getDataRange().getValues();
  let errorData = initializeDbData();
  errorData[0].push('Reason');

  let filteredData = [];
  const colToCheck = [DB_COL_ROLE.gmail - 1, DB_COL_ROLE.date - 1, DB_COL_ROLE.hour - 1, DB_COL_ROLE.minute - 1];
  [filteredData, errorData] = sortAndExtractDuplicates(orgData, colToCheck, errorData);

  input2DArrayToDataSheet(filteredData);
  saveErrorLogToSheet(errorData);
}

// 指定した列がすべて重複する行をエラーログに移動
function sortAndExtractDuplicates(data, columnIndexes, duplicateData) {
  // 1. ソート処理：指定した列の順にソート
  data.sort((a, b) => {
    for (let col of columnIndexes) {
      const valueA = a[col];
      const valueB = b[col];
      if (valueA < valueB) return -1;
      if (valueA > valueB) return 1;
    }
    return 0; // 全て同じ場合
  });

  // 2. 重複を探し、重複した行を別の配列に移動
  let uniqueData = [];
  let lastRow = null;
  let isDuplicate = false;

  for (let i = 0; i < data.length; i++) {
    let currentRow = data[i];

    // 前の行と比較して、指定列の値がすべて同じか確認
    if (lastRow && columnIndexes.every(col => currentRow[col] === lastRow[col])) {
      if (!isDuplicate) {
        // 前の行が最初の重複の場合、前の行をコピーして重複配列に追加
        let duplicateLastRow = [...lastRow]; // lastRowをコピー
        duplicateLastRow.push("duplicate dupeID:" + i);
        duplicateData.push(duplicateLastRow);
        isDuplicate = true;
      }
      // 現在の行もコピーして重複として追加
      let duplicateCurrentRow = [...currentRow]; // currentRowをコピー
      duplicateCurrentRow.push("duplicate dupeID:" + i);
      duplicateData.push(duplicateCurrentRow);
    } else {
      // 重複していない場合は一意データに追加
      uniqueData.push(currentRow);
      isDuplicate = false;
    }

    // 重複していない行のみを lastRow にセット
    if (!isDuplicate) {
      lastRow = currentRow;
    }
  }

  // 3. 重複を除いた配列と重複を含んだ配列を返す
  return [uniqueData, duplicateData];
}


