// // 24 準1当日1,2撤収日私用調査、特性調査
// const PERSONAL_SS_URL = "https://docs.google.com/spreadsheets/d/XXXXXXXXXX/edit";

// 24 準0私用調査
const PERSONAL_SS_URL = "https://docs.google.com/spreadsheets/d/XXXXXXXXXX/edit";

const PERSONAL_SHEET_NAME = "フォームの回答 1";

// // 24 準1当日1,2撤収日私用調査、特性調査
// const PERSONAL_COL_ROLE = {
//   gmail: 3,
//   name: 4,
//   tel: 9,
//   department: 14,
//   schedule: [15, 17, 19, 21],
//   unavailableTime: [16, 18, 20, 22],
//   voice: 24,
//   power: 25,
//   disease1: 28,
//   disease2: 29
// };

// 24 準0私用調査
const PERSONAL_COL_ROLE = {
  gmail: 3,
  name: 4,

  department: 10,
  schedule: [7],
  unavailableTime: [7],

}

const SEPARATORS = { timeRange: ",", startEnd: "～", hourMinute: ":" };

const SCHEDULE_OPTIONS = {
  empty: "終日可",
  full: "終日不可",
  depends: "入れない時間帯がある",
  pending: "入れない可能性があるが詳細は未定"
}
const TIME_RANGE_REG_EX = /^([0-9]|1[0-9]|2[0-3]):[0-5][0-9].*([0-9]|1[0-9]|2[0-3]):[0-5][0-9]$/;// 準0私用調査。入れる時間が直に記入

const SHIFTDATA_OPTIONS = {
  full: { bgColor: '#000000', shift: '私用' },
  empty: { bgColor: '#FFFFFF', shift: '' },
  pending: { bgColor: '#A9A9A9', shift: '' },

  fixed: { bgColor: '#18499d', shift: '固定シフト' },
  booking: { bgColor: '0000ff', shift: 'ブッキング' },

  error: { bgColor: '#ff0000', shift: 'エラー' }
}


function formatPersonalData() {
  const sheet = SpreadsheetApp.openByUrl(PERSONAL_SS_URL).getSheetByName(PERSONAL_SHEET_NAME); // データのシート名を指定
  const data = sheet.getDataRange().getValues(); // シート全体のデータを取得
  const formattedData = initializeDbData();

  data.forEach((row, index) => {
    if (index === 0) return; // ヘッダー行はスキップ

    // availableTimesをループで生成
    const availableTimes = [];
    for (let i = 0; i < DAYS.length; i++) {
      availableTimes.push({
        date: DAYS[i],
        schedule: row[PERSONAL_COL_ROLE.schedule[i] - 1],
        unavailableTime: row[PERSONAL_COL_ROLE.unavailableTime[i] - 1]
      });
    }

    availableTimes.forEach((shiftData) => {
      for (let hour = HOURS.start; hour <= HOURS.end; hour++) {
        for (let minute of MINUTES) {
          // 時間帯の処理を分岐させる関数を呼び出し
          let min = parseInt(minute, 10);
          const { bgColor, shift } = getBgColorAndShift(shiftData, hour, min);

          // フォーマットされた行を順序通りに配置
          const formattedRow = [];
          // オプションとなるデータ転記
          for (const key in DB_COL_ROLE) {
            if (["date", "hour", "minute", "bgColor", "shift"].includes(key)) continue;
            formattedRow[DB_COL_ROLE[key] - 1] = row[PERSONAL_COL_ROLE[key] - 1] || "";
          }

          formattedRow[DB_COL_ROLE.bgColor - 1] = bgColor;
          formattedRow[DB_COL_ROLE.shift - 1] = shift;
          formattedRow[DB_COL_ROLE.date - 1] = shiftData.date;
          formattedRow[DB_COL_ROLE.hour - 1] = `${hour}時`;
          formattedRow[DB_COL_ROLE.minute - 1] = minute;

          formattedData.push(formattedRow);
        }
      }
    });
  });

  // 整形したデータを新しいシートに書き込み
  const dataSs = SpreadsheetApp.openByUrl(DATA_SS_URL);
  const outputSheet = dataSs.getSheetByName(DATA_SHEET_NAME) || dataSs.insertSheet(DATA_SHEET_NAME);
  outputSheet.getRange(1, 1, formattedData.length, formattedData[0].length).setValues(formattedData);
}



// 時間帯に応じて背景色とシフトの処理を分岐する関数
function getBgColorAndShift(shiftData, hour, minute) {
  let bgColor = "";
  let shift = "";
  let scheduledTime;

  // 入れない時間が終日不可、終日可、などの回答と同じ列に記入されていた場合に対応するための処理
  if (TIME_RANGE_REG_EX.test(shiftData.schedule)) {
    scheduledTime = SCHEDULE_OPTIONS.depends;
  } else {
    scheduledTime = shiftData.schedule;
  }

  switch (scheduledTime) {
    case SCHEDULE_OPTIONS.empty:
      bgColor = SHIFTDATA_OPTIONS.empty.bgColor;
      shift = SHIFTDATA_OPTIONS.empty.shift;
      break;

    case SCHEDULE_OPTIONS.full:
      bgColor = SHIFTDATA_OPTIONS.full.bgColor; // 黒色 (入れない時間帯)
      shift = SHIFTDATA_OPTIONS.full.shift; // シフトに「私用」を設定
      break;

    case SCHEDULE_OPTIONS.depends:
      try {
        // 15分刻みで入れない時間帯の処理
        if (isUnavailable(shiftData.unavailableTime, hour, minute)) {
          bgColor = SHIFTDATA_OPTIONS.full.bgColor; // 黒色 (入れない時間帯)
          shift = SHIFTDATA_OPTIONS.full.shift; // シフトに「私用」を設定
        } else {
          bgColor = SHIFTDATA_OPTIONS.empty.bgColor; // 入れない時間帯が無い部分は白
          shift = SHIFTDATA_OPTIONS.empty.shift; // 通常シフト
        }
      } catch (e) {
        bgColor = SHIFTDATA_OPTIONS.error.bgColor; // 赤色 (エラー)
        shift = SHIFTDATA_OPTIONS.error.shift; // エラーとして処理
      }
      break;

    case SCHEDULE_OPTIONS.pending:
      bgColor = SHIFTDATA_OPTIONS.pending.bgColor; // 灰色 (未定)
      shift = SHIFTDATA_OPTIONS.pending.shift;
      break;

    default:
      bgColor = '#ff0000'; // デフォルト
      shift = 'エラー';
      break;
  }

  return { bgColor, shift };
}

function test() {
  const res = isUnavailable()
  Logger.log(res);
}

// 入れない時間帯を15分刻みにする処理を行う関数
function isUnavailable(unavailableTime, hour, minute) {
  try {
    // カンマで複数の時間帯を分割
    const timeRanges = unavailableTime.split(SEPARATORS.timeRange);

    // 各時間帯を処理
    for (let timeRange of timeRanges) {
      const [start, end] = timeRange.split(SEPARATORS.startEnd).map(t => {
        t = t.trim();
        const timeParts = t.split(SEPARATORS.hourMinute);
        if (timeParts.length !== 2) {
          throw new Error('Invalid time format');
        }
        return timeParts.map(Number);
      });

      // 時間帯内にあるかを判定
      if ((hour > start[0] || (hour === start[0] && minute >= start[1])) &&
        (hour < end[0] || (hour === end[0] && minute < end[1]))) {
        return true; // 指定の時間帯に該当する場合
      }
    }

    return false; // いずれの時間帯にも該当しない場合

  } catch (e) {
    // 想定外の書式のエラーハンドリング
    throw new Error('Invalid time range format');
  }
}
