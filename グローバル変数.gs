const DATA_SHEET_NAME = "データシート";

// 当日シフト
// const GANTT_CHART_SS_URL = "https://docs.google.com/spreadsheets/d/XXXXXXXXXX/edit";
// const DATA_SS_URL = "https://docs.google.com/spreadsheets/d/XXXXXXXXXX/edit";

// 準0シフト
const GANTT_CHART_SS_URL = "https://docs.google.com/spreadsheets/d/XXXXXXXXXX/edit";
const DATA_SS_URL = "https://docs.google.com/spreadsheets/d/XXXXXXXXXX/edit";

// 当日シフト
const GC_COL_ROLE = {
  team: 1,
  name: 2,
  furigana: 3,
  gmail: 4,
  tel: 5,
  gender: 6,
  voice: 7,
  power: 8,
  disease1: 9,
  disease2: 10,
  date: 11,
  firstData: 12
}

// // 準0シフト
// const GC_COL_ROLE = {
//   team: 1,
//   name: 2,
//   furigana: 3,
//   gmail: 4,
//   tel: 5,
//   gender: 6,
//   note: 7,
//   date: 8,
//   firstData: 9
// }

const GC_ROW_ROLE = {
  hour: 1,
  minute: 2,
  firstData: 3
}

const DB_COL_ROLE = {
  // 必須
  department: 1,//GCのシートに対応
  gmail: 2,//gmail+dateでGCの行に対応
  date: 3,
  hour: 4,//hour+minuteでGCの列に対応
  minute: 5,
  bgColor: 6,//bgColorとshiftはGCのセルの情報
  shift: 7,
  // オプション
  name: 8
}



