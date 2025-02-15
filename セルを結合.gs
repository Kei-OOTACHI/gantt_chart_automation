function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('セル結合ツール')
    .addItem('横方向に同じ値を結合', 'mergeSameValuesHorizontally')
    .addItem('縦方向に同じ値を結合', 'mergeSameValuesVertically')
    .addToUi();
}

function mergeSameValuesHorizontally(sheet = undefined, range = undefined) {
  if (!sheet) sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (!range) range = sheet.getActiveRange();
  Logger.log(sheet.getName());
  Logger.log(range.getA1Notation());
  const values = range.getValues();

  const startRow = range.getRow();
  const startColumn = range.getColumn();

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    let mergeStart = 0;

    for (let j = 1; j < row.length; j++) {
      if (row[j] !== row[mergeStart] || row[mergeStart] === "") {
        if (j - mergeStart > 1 && row[mergeStart] !== "") {
          const mergeRange = sheet.getRange(startRow + i, startColumn + mergeStart, 1, j - mergeStart);
          mergeRange.merge();
          mergeRange.setHorizontalAlignment("center");
        }
        mergeStart = j;
      }
    }

    if (row.length - mergeStart > 1 && row[mergeStart] !== "") {
      const mergeRange = sheet.getRange(startRow + i, startColumn + mergeStart, 1, row.length - mergeStart);
      mergeRange.merge();
      mergeRange.setHorizontalAlignment("center");
    }
  }
}

function mergeSameValuesVertically(sheet = undefined, range = undefined) {
  if (!sheet) sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (!range) range = sheet.getActiveRange();
  const values = range.getValues();

  const startRow = range.getRow();
  const startColumn = range.getColumn();

  for (let j = 0; j < values[0].length; j++) {
    let mergeStart = 0;

    for (let i = 1; i < values.length; i++) {
      if (values[i][j] !== values[mergeStart][j] || values[mergeStart][j] === "") {
        if (i - mergeStart > 1 && values[mergeStart][j] !== "") {
          const mergeRange = sheet.getRange(startRow + mergeStart, startColumn + j, i - mergeStart, 1);
          mergeRange.merge();
          mergeRange.setHorizontalAlignment("center");
        }
        mergeStart = i;
      }
    }

    if (values.length - mergeStart > 1 && values[mergeStart][j] !== "") {
      const mergeRange = sheet.getRange(startRow + mergeStart, startColumn + j, values.length - mergeStart, 1);
      mergeRange.merge();
      mergeRange.setHorizontalAlignment("center");
    }
  }
}