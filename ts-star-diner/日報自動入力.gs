/**
 * T's Star Diner — 日報の時間帯別「客数・売上・組数」を funfo から自動入力
 *
 * 「T's Star Diner 経営データ基盤」の Apps Script に、既存の取込コードとは
 * 別ファイルとして追加して使う（auth_ / post_ / props_ / dateStr_ を共用）。
 *
 * 書き込み先: マイドライブの「日報Diner本店_YYYY年M月」の日付シート（"1"〜"31"）
 *   D5:D18 … 時間帯別 客数
 *   F5:F18 … 時間帯別 売上（funfo の 売上合計 − 割引額 ＝ 純売上）
 *   O18:O20 … 組数（〜15時 / 17〜23時 / 24時〜L）
 * それ以外のセル（使用人時・仕入・光熱・コメント等）には触れない。
 *
 * 時間帯の扱い（営業日 D）:
 *   11〜12時の行 … D 03:00〜11:59（開店前の注文もここに含める）
 *   14〜15時の行 … D 14:00〜16:59（中休み 15〜17時の注文もここに含める）
 *   24〜1時 / 1〜2時 / 2〜3時の行 … 翌日 00:00〜02:59
 *
 * 使い方:
 *   1. setupNippoTrigger() を1回実行 → 毎朝7時台に前日分を日報へ自動入力
 *   2. 手動で入れたいときはメニュー「日報自動入力」から
 */

var NIPPO_FILE_PREFIX = '日報Diner本店_';

// row: 日報の行番号 / dayOffset: 0=当日, 1=翌日 / from,to: funfo に渡す時刻
var NIPPO_BANDS = [
  { row: 5,  dayOffset: 0, from: '03:00', to: '11:59' },
  { row: 6,  dayOffset: 0, from: '12:00', to: '12:59' },
  { row: 7,  dayOffset: 0, from: '13:00', to: '13:59' },
  { row: 8,  dayOffset: 0, from: '14:00', to: '16:59' },
  { row: 9,  dayOffset: 0, from: '17:00', to: '17:59' },
  { row: 10, dayOffset: 0, from: '18:00', to: '18:59' },
  { row: 11, dayOffset: 0, from: '19:00', to: '19:59' },
  { row: 12, dayOffset: 0, from: '20:00', to: '20:59' },
  { row: 13, dayOffset: 0, from: '21:00', to: '21:59' },
  { row: 14, dayOffset: 0, from: '22:00', to: '22:59' },
  { row: 15, dayOffset: 0, from: '23:00', to: '23:59' },
  { row: 16, dayOffset: 1, from: '00:00', to: '00:59' },
  { row: 17, dayOffset: 1, from: '01:00', to: '01:59' },
  { row: 18, dayOffset: 1, from: '02:00', to: '02:59' }
];

// ---------------- エントリーポイント ----------------

/** トリガー用: 前日分を日報に入力 */
function fetchNippoJob() {
  Logger.log(writeNippo_(dateStr_(-1)));
}

/** 毎朝7時台のトリガーと、メニュー表示用トリガーを設定（既存の同名トリガーは張り替え） */
function setupNippoTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    var h = t.getHandlerFunction();
    if (h === 'fetchNippoJob' || h === 'addNippoMenu') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('fetchNippoJob').timeBased().everyDays(1).atHour(7).create();
  ScriptApp.newTrigger('addNippoMenu').forSpreadsheet(SPREADSHEET_ID).onOpen().create();
  Logger.log('毎日7時台に前日分を日報へ入力するトリガーを設定しました');
}

/** 経営データ基盤を開いたときにメニューを追加（既存の onOpen とは別に動く） */
function addNippoMenu() {
  SpreadsheetApp.getUi().createMenu('日報自動入力')
    .addItem('前日分を日報に入力', 'menuNippoYesterday')
    .addItem('日付を指定して日報に入力', 'menuNippoOneDay')
    .addToUi();
}

function menuNippoYesterday() {
  SpreadsheetApp.getUi().alert(writeNippo_(dateStr_(-1)));
}

function menuNippoOneDay() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt('日報に入力する日付を YYYY-MM-DD で入力\n（その日の客数・売上・組数は上書きされます）', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  ui.alert(writeNippo_(res.getResponseText().trim()));
}

// ---------------- 本体 ----------------

/** 指定日(YYYY-MM-DD)の時間帯別実績を funfo から取得し、その月の日報に書き込む */
function writeNippo_(dateStr) {
  var date = parseDate_(dateStr);
  var fileName = NIPPO_FILE_PREFIX + date.getFullYear() + '年' + (date.getMonth() + 1) + '月';
  var files = DriveApp.getFilesByName(fileName);
  if (!files.hasNext()) return dateStr + ': 「' + fileName + '」が見つからないためスキップ';
  var sheet = SpreadsheetApp.open(files.next()).getSheetByName(String(date.getDate()));
  if (!sheet) return dateStr + ': 「' + fileName + '」にシート「' + date.getDate() + '」が無いためスキップ';

  var token = auth_();
  var nextStr = Utilities.formatDate(new Date(date.getTime() + 86400000), TZ, 'yyyy-MM-dd');
  var guests = [], sales = [], groups = {};
  var totalSales = 0, totalGuests = 0, totalGroups = 0;

  NIPPO_BANDS.forEach(function (b) {
    var r = fetchBand_(token, b.dayOffset ? nextStr : dateStr, b.from, b.to);
    guests.push([r.guests]);
    sales.push([r.sales]);
    groups[b.row] = r.groups;
    totalSales += r.sales;
    totalGuests += r.guests;
    totalGroups += r.groups;
    Utilities.sleep(300); // API負荷への配慮
  });

  // 組数: 〜15時 = 行5〜8 / 17〜23時 = 行9〜14 / 24時〜L = 行15〜18（日報の集計式と同じ区切り）
  var g = function (from, to) { var s = 0; for (var i = from; i <= to; i++) s += groups[i]; return s; };

  sheet.getRange('D5:D18').setValues(guests);
  sheet.getRange('F5:F18').setValues(sales);
  sheet.getRange('O18:O20').setValues([[g(5, 8)], [g(9, 14)], [g(15, 18)]]);

  return dateStr + ': 日報「' + fileName + '」シート' + date.getDate() + ' に入力 / 売上¥' + totalSales +
    ' / 客数' + totalGuests + ' / ' + totalGroups + '組';
}

/** 1つの時間帯の 客数・純売上・組数 を返す */
function fetchBand_(token, dateStr, from, to) {
  var c = props_();
  var json = post_(c.base + '/swo/clientApi/statistics/v2/store/sale/data', {
    orderDate: dateStr,
    orderType: 7,
    storeEmails: [c.email],
    startHourMinute: from,
    endHourMinute: to,
    userEmail: c.email
  }, token);
  var day = (json.data && json.data.amountHours && json.data.amountHours[0]) || {};
  return {
    guests: day.personNum || 0,
    sales: (day.amount || 0) - (day.discount || 0),
    groups: day.orderNum || 0
  };
}
