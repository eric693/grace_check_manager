// Holidays.gs — 國定假日（管理員在資料管理頁的「國定假日」分頁維護）
//
// 以前清單寫死在程式裡（TAIWAN_HOLIDAYS_2026），換年就要改程式；現在存在「國定假日」表，
// 每年在網頁上加新的一年就好。還沒存過的話，沿用程式內建的 2026 年清單。
//
// 每一天兩個意思：
//   ・是國定假日（請假時數、加班類型都照假日算）——清單裡的每一天都是
//   ・加倍計薪：時段計薪（ShiftPay.gs）那天時薪乘上倍率。補假日要不要加倍，各店做法不同，
//     所以每一天可以單獨設定

const SHEET_HOLIDAYS = '國定假日';
const HOLIDAY_HEADERS = ['日期', '名稱', '加倍計薪'];

// 程式內建的 2026 年清單（還沒在網頁上存過時用）
const DEFAULT_HOLIDAY_NAMES = {
  '2026-01-01': '開國紀念日', '2026-02-15': '農曆除夕前一日', '2026-02-16': '農曆除夕', '2026-02-17': '春節',
  '2026-02-18': '春節', '2026-02-19': '春節', '2026-02-20': '除夕前一日補假', '2026-02-27': '和平紀念日補假',
  '2026-02-28': '和平紀念日', '2026-04-03': '兒童節補假', '2026-04-04': '兒童節', '2026-04-05': '清明節',
  '2026-04-06': '清明節補假', '2026-05-01': '勞動節', '2026-06-19': '端午節', '2026-09-25': '中秋節',
  '2026-09-28': '教師節', '2026-10-09': '國慶日補假', '2026-10-10': '國慶日', '2026-10-25': '臺灣光復節',
  '2026-10-26': '臺灣光復節補假', '2026-12-25': '行憲紀念日'
};

let _holidayCache = null;   // 同一次執行只讀一次表（薪資計算會逐日查）

/** 所有國定假日：[{ date, name, doublePay }]，依日期排序 */
function readHolidays_() {
  if (_holidayCache) return _holidayCache;
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_HOLIDAYS);
  let list;
  if (sheet && sheet.getLastRow() >= 2) {
    const tz = Session.getScriptTimeZone();
    list = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues()
      .map(r => ({
        date: (r[0] && typeof r[0].getTime === 'function') ? Utilities.formatDate(r[0], tz, 'yyyy-MM-dd') : String(r[0] || '').trim(),
        name: String(r[1] || '').trim(),
        doublePay: String(r[2] || '').trim() !== '否'
      }))
      .filter(h => /^\d{4}-\d{2}-\d{2}$/.test(h.date));
  } else {
    list = (typeof TAIWAN_HOLIDAYS_2026 !== 'undefined' ? TAIWAN_HOLIDAYS_2026 : [])
      .map(d => ({ date: d, name: DEFAULT_HOLIDAY_NAMES[d] || '', doublePay: true }));
  }
  list.sort((a, b) => a.date.localeCompare(b.date));
  _holidayCache = list;
  return list;
}

/** 是不是國定假日（YYYY-MM-DD） */
function isNationalHoliday(dateStr) {
  return readHolidays_().some(h => h.date === dateStr);
}

/** 國定假日、而且設定要加倍計薪 */
function isDoublePayHoliday_(dateStr) {
  return readHolidays_().some(h => h.date === dateStr && h.doublePay);
}

/** API：國定假日清單（不需登入；前端判斷工作日用） */
function handleGetHolidays() {
  const list = readHolidays_();
  return {
    ok: true,
    year: new Date().getFullYear(),
    holidays: list.map(h => h.date),
    items: list
  };
}

/**
 * API（管理員）：儲存整份國定假日清單
 * 參數：holidays（JSON 陣列，每項 { date: 'yyyy-MM-dd', name, doublePay }）
 */
function handleSaveHolidays(params) {
  try {
    const admin = requireRecordsAdmin_(params.token);
    if (!admin.ok) return admin;
    let input;
    try {
      input = JSON.parse(params.holidays || '[]');
    } catch (error) {
      return { ok: false, code: 'HOLIDAY_INVALID', msg: '國定假日資料格式錯誤' };
    }
    if (!Array.isArray(input) || input.length > 500) return { ok: false, code: 'HOLIDAY_INVALID', msg: '國定假日資料格式錯誤' };

    const seen = {};
    const rows = [];
    for (const item of input) {
      const date = String((item && item.date) || '').trim();
      const m = date.match(/^(\d{4})-(\d{2})-(\d{2})$/);
      const d = m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : null;
      if (!d || d.getMonth() !== Number(m[2]) - 1 || d.getDate() !== Number(m[3])) {
        return { ok: false, code: 'HOLIDAY_DATE', msg: `日期「${date}」不正確` };
      }
      if (seen[date]) return { ok: false, code: 'HOLIDAY_DUPLICATE', msg: `日期重複：${date}` };
      seen[date] = true;
      const name = String(item.name || '').trim().slice(0, 30);
      // 日期存成文字，試算表才不會依地區把它轉成別的格式
      rows.push(["'" + date, name, item.doublePay === false ? '否' : '是']);
    }
    rows.sort((a, b) => a[0].localeCompare(b[0]));

    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const ss = SpreadsheetApp.getActive();
      let sheet = ss.getSheetByName(SHEET_HOLIDAYS);
      if (!sheet) {
        sheet = ss.insertSheet(SHEET_HOLIDAYS);
        sheet.getRange(1, 1, 1, HOLIDAY_HEADERS.length).setValues([HOLIDAY_HEADERS])
             .setFontWeight('bold').setBackground('#1f4e79').setFontColor('#ffffff');
        sheet.setFrozenRows(1);
      }
      const last = sheet.getLastRow();
      if (last > 1) sheet.getRange(2, 1, last - 1, HOLIDAY_HEADERS.length).clearContent();
      if (rows.length) sheet.getRange(2, 1, rows.length, HOLIDAY_HEADERS.length).setValues(rows);
    } finally {
      lock.releaseLock();
    }
    _holidayCache = null;

    const years = {};
    rows.forEach(r => { const y = r[0].slice(1, 5); years[y] = (years[y] || 0) + 1; });
    params.changes = Object.keys(years).sort().map(y => `${y} 年 ${years[y]} 天`).join('、') || '清空';
    params.holidays = '';
    return { ok: true, items: readHolidays_() };
  } catch (error) {
    Logger.log(' handleSaveHolidays 錯誤: ' + error.message);
    return { ok: false, msg: '儲存國定假日失敗：' + error.message };
  }
}
