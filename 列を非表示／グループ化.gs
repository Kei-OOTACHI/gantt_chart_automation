const COLS_TO_HIDE = [3, 4, 5];
const COL_TO_FOLD = [7,8,9,10];

function hideOrFoldCols(sheet) {

  hideColumns(sheet, COLS_TO_HIDE);
  groupColumns(sheet, COL_TO_FOLD);

}

function hideColumns(sheet, columnsArray) {
  columnsArray.forEach(function (column) {
    var range = sheet.getRange(1, column);  // 1行目の列を基準に取得
    sheet.hideColumn(range);  // 指定された列を非表示にする
  });
}

function groupColumns(sheet, columnsArray) {
  columnsArray.forEach(function (column) {
    var range = sheet.getRange(1, column);  // 1行目の列を基準に取得
    range.shiftColumnGroupDepth(1).collapseGroups();
  });
}

