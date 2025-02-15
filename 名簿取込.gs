// 参考：https://chatgpt.com/share/66f6ee8b-0e58-800d-b655-ff92bee2b5db

// 24新規名簿(「名簿」は英語で "roster" らしい)
const ROSTER_SS_URL2 = 'https://docs.google.com/spreadsheets/d/XXXXXXXXXX/edit';
// // 24既存名簿
const ROSTER_SS_URL1 = "https://docs.google.com/spreadsheets/d/XXXXXXXXXX/edit";

// 24新規名簿の場合
const ROSTER_COL_ROLE2 = {
  name: 2,
  furigana: 3,
  department: 7,
  team: 8,
  gender: 9,
  gmail: 14
}

// 24既存名簿の場合
const ROSTER_COL_ROLE1 = {
  name: 2,
  furigana: 3,
  department: 5,
  team: 6,
  gender: 7,
  gmail: 12
}
// 当日シフト
// const DAYS = ["準備日 11/1(金)", "当日 11/2(土)", "当日 11/3(日)", "撤収日 11/4(月)"];
// const HOURS = { start: 8, end: 21 }

// 準0シフト
const DAYS = ["準備日0日目"];
const HOURS = { start: 17, end: 22 }
const MINUTES = ["00-", "15-", "30-", "45-"];

// 列の幅と行の高さの設定
const HEADER_COLUMN_WIDTH = 80;
const TIMESCALE_COLUMN_WIDTH = 21;
const DATE_COL_WHIDTH = 100;
const GENDER_COL_WHIDTH = 40;
const HEADER_ROW_HEIGHT = 21;
const BLANK_ROW_HEIGHT = 6;


const HEADER_COL_COLOR = "#d9d9d9";
const HOURS_ROW_COLOR = "#999999";
const MINUTES_ROW_COLOR = "#d9d9d9";
const BORDER_COLOR = "#0b5394";


function createGcFrame() {
  Logger.log("名簿1開始");
  copyMemberData(ROSTER_SS_URL1, ROSTER_COL_ROLE1);
  Logger.log("名簿2開始");
  copyMemberData(ROSTER_SS_URL2, ROSTER_COL_ROLE2);
  const gcSs = SpreadsheetApp.openByUrl(GANTT_CHART_SS_URL);
  gcSs.getSheets().forEach(sheet => {
    hideOrFoldCols(sheet);
  });
}

// 見出し列の名前などの情報を名簿から転記する関数
function copyMemberData(rosterSsUrl, rosterColRole) {
  const rosterSpreadsheet = SpreadsheetApp.openByUrl(rosterSsUrl);
  const rosterSheets = rosterSpreadsheet.getSheets();
  const gcSs = SpreadsheetApp.openByUrl(GANTT_CHART_SS_URL);
  rosterSheets.forEach(rosterSheet => {
    const sheetName = rosterSheet.getSheetName();
    //既存全体、新規全体のシートはスキップ
    if (["既存全体", "新規全体"].includes(sheetName)) return;
    Logger.log(`${sheetName} 開始`);

    const gcSheet = gcSs.getSheetByName(sheetName) || gcSs.insertSheet(sheetName);

    const headerData = extractData(rosterSheet);

    const sortedData = sortHeaderData(headerData, rosterColRole);

    const groupedData = prepareGroupedData(sortedData, rosterColRole);

    setTimescale(gcSheet);

    setHeaderDataAndFormat(groupedData, gcSheet);
  });
}

// 名簿（roster）シートからデータを抽出する関数
function extractData(sheet) {
  // すべてのデータ行を取得（1行目の見出し行を除外）
  const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
  return dataRange.getValues();
}

// データを並び替える関数
function sortHeaderData(data, rosterColRole) {
  // 所属チーム、ふりがなの順で並び替え
  data.sort((a, b) => {
    if (a[rosterColRole.team - 1] === b[rosterColRole.team - 1]) {
      return a[rosterColRole.furigana - 1].localeCompare(b[rosterColRole.furigana - 1]);
    }
    return a[rosterColRole.team - 1].localeCompare(b[rosterColRole.team - 1]);
  });
  return data;
}

// 一人当たり4行ずつ複製する関数（準備日、当日1日目、当日2日目、撤収日の4行）
function prepareGroupedData(data, rosterColRole) {
  const groupedData = [];

  // 各人のデータを処理
  data.forEach(row => {
    // 日ごとの行を追加
    DAYS.forEach(day => {
      const dailyData = createDailyData(row, rosterColRole, day);
      groupedData.push(dailyData);
    });

    // 空行を追加
    groupedData.push(createEmptyRow(groupedData[0]));
  });

  return groupedData;
}

// 日ごとの行を生成する関数
function createDailyData(row, rosterColRole, day) {
  let orderedData = [];

  for (const key in GC_COL_ROLE) {
    if (key == "day" || key == "firstData") continue;
    orderedData[GC_COL_ROLE[key] - 1] = row[rosterColRole[key] - 1] || "";
  }
  orderedData[GC_COL_ROLE.date - 1] = day;

  return orderedData;
}

// 空行を生成する関数
function createEmptyRow(dailyData) {
  return Array(dailyData.length).fill("");
}

// データをガントチャートに記入し、さらに書式も整える関数
function setHeaderDataAndFormat(data, sheet) {
  const lastRowBefInput = sheet.getLastRow();
  const startRow = Math.max(lastRowBefInput + 2, GC_ROW_ROLE.firstData); // 最後の行の次に1行空白行を入れ、その次の行からデータを開始

  // データを記入
  const range = sheet.getRange(startRow, 1, data.length, data[0].length);
  range.setValues(data);

  // 書式を整える（背景色、罫線の色垂直方向の配置、幅と高さ） 
  range.setBackground(HEADER_COL_COLOR);
  range.setBorder(true, true, true, true, true, true, BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID);
  range.setVerticalAlignment("middle");
  sheet.setColumnWidths(1, GC_COL_ROLE.firstData - 1, HEADER_COLUMN_WIDTH);
  sheet.setRowHeightsForced(startRow, data.length - 1, HEADER_ROW_HEIGHT);

  const lastCol = sheet.getLastColumn();// ここ、人の情報を見出し列に記入するより先に時間を見出し行に記入しておかないと、空白行のシフトデータエリアの枠線の色が変わらないので注意
  // 空行だけは背景色、罫線の色、高さを別指定
  for (let i = startRow - 1; i <= startRow + data.length - 1; i += DAYS.length + 1) {
    const blancRowRange = sheet.getRange(i, 1, 1, lastCol);  // 指定された行の範囲を取得
    blancRowRange.setBackground("#ffffff");
    blancRowRange.setBorder(true, true, true, true, true, true, BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID);
    sheet.setRowHeight(i, BLANK_ROW_HEIGHT);
  }

  // 日付と性別の列だけは幅別指定
  sheet.setColumnWidth(GC_COL_ROLE.date, DATE_COL_WHIDTH);
  sheet.setColumnWidth(GC_COL_ROLE.gender, GENDER_COL_WHIDTH);

  // 
  mergeSameValuesVertically(sheet, range);
}

// 見出し行に時間を記入する関数
function setTimescale(sheet) {
  let h = [];
  let min = [];
  for (let i = HOURS.start; i <= HOURS.end; i++) {
    for (const j of MINUTES) {
      h.push(i + "時");
      min.push(j);
    }
  }

  let timescale = [];
  timescale[GC_ROW_ROLE.hour - 1] = h;
  timescale[GC_ROW_ROLE.minute - 1] = min;

  const range = sheet.getRange(1, GC_COL_ROLE.firstData, timescale.length, timescale[0].length);
  const hRange = sheet.getRange(GC_ROW_ROLE.hour, GC_COL_ROLE.firstData, 1, timescale[0].length);
  const minRange = sheet.getRange(GC_ROW_ROLE.minute, GC_COL_ROLE.firstData, 1, timescale[0].length);
  range.setValues(timescale);

  hRange.setBackground(HOURS_ROW_COLOR);
  minRange.setBackground(MINUTES_ROW_COLOR);
  range.setBorder(true, true, true, true, true, true, BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID);
  sheet.setColumnWidths(GC_COL_ROLE.firstData, timescale[0].length, TIMESCALE_COLUMN_WIDTH);
  mergeSameValuesHorizontally(sheet, range);
}
