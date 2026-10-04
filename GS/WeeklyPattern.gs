// WeeklyPattern.gs
//
// 固定班表：每位員工星期一到日各上什麼班（班別代碼見 ShiftTemplates.gs）。
// 存好之後可以「產生排班」，把一段日期依星期幾寫進排班表，不用每週一筆一筆排。
//
// 產生的規則：
//   ・這位員工那一天已經有排班（手動排的、或之前產生的）→ 跳過，不會重複
//   ・勾了「覆蓋」→ 這段期間由固定班表產生、而且還沒改過的排班先刪掉再重排，
//     手動排的班、手動改過的班都不會動（固定班表改了之後用這個更新未來的班）
//   ・產生的排班備註寫「固定班表」，用來分辨
//
// 要每週自動排：在 Apps Script 編輯器執行一次 installWeeklyShiftTrigger()，
// 之後每週一凌晨會自動把未來 4 週補滿（只補空的日子，不會覆蓋）。

const SHEET_WEEKLY_PATTERN = '固定班表';
const WEEKLY_PATTERN_HEADERS = ['員工ID', '員工姓名', '星期', '班別代碼', '地點'];
const WEEKLY_PATTERN_NOTE = '固定班表';
const WEEKLY_PATTERN_MAX_DAYS = 93;          // 一次最多產生約三個月
const WEEKLY_PATTERN_AUTO_DAYS = 28;         // 自動排班往後補幾天

function getWeeklyPatternSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_WEEKLY_PATTERN);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_WEEKLY_PATTERN);
    sheet.getRange(1, 1, 1, WEEKLY_PATTERN_HEADERS.length).setValues([WEEKLY_PATTERN_HEADERS])
         .setFontWeight('bold').setBackground('#1f4e79').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * 讀出固定班表
 * @returns {Array<{employeeId, employeeName, weekday, code, location}>} weekday：1 = 星期一 … 7 = 星期日
 */
function readWeeklyPattern_() {
  const data = getWeeklyPatternSheet_().getDataRange().getValues();
  const list = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const employeeId = String(row[0] || '').trim();
    const weekday = Number(row[2]);
    const code = String(row[3] || '').trim();
    if (!employeeId || !code || !(weekday >= 1 && weekday <= 7)) continue;
    list.push({
      employeeId: employeeId,
      employeeName: String(row[1] || '').trim(),
      weekday: weekday,
      code: code,
      location: String(row[4] || '').trim()
    });
  }
  return list;
}

/** "yyyy-MM-dd" → 星期（1 = 一 … 7 = 日）。用正午避開時區換日 */
function weekdayOfDateStr_(dateStr) {
  const p = dateStr.split('-').map(Number);
  const day = new Date(p[0], p[1] - 1, p[2], 12).getDay();
  return day === 0 ? 7 : day;
}

/** 從 startStr 到 endStr（含）的每一天，"yyyy-MM-dd" */
function listDateStrs_(startStr, endStr) {
  const out = [];
  const p = startStr.split('-').map(Number);
  const d = new Date(p[0], p[1] - 1, p[2], 12);
  const pad = n => String(n).padStart(2, '0');
  for (let i = 0; i <= WEEKLY_PATTERN_MAX_DAYS; i++) {
    const s = `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
    if (s > endStr) break;
    out.push(s);
    d.setDate(d.getDate() + 1);
  }
  return out;
}

/**
 * 依固定班表產生排班
 * @param {string} startStr "yyyy-MM-dd"
 * @param {string} endStr   "yyyy-MM-dd"
 * @param {{replace?: boolean, by?: string}} [options]
 * @returns {{ok: boolean, added?: number, skipped?: number, replaced?: number, msg?: string}}
 */
function generateShiftsFromPattern_(startStr, endStr, options) {
  const opts = options || {};
  const datePattern = /^\d{4}-\d{2}-\d{2}$/;
  if (!datePattern.test(startStr || '') || !datePattern.test(endStr || '')) {
    return { ok: false, code: 'WEEKLY_PATTERN_DATE', msg: '日期格式要是 yyyy-MM-dd' };
  }
  if (startStr > endStr) {
    return { ok: false, code: 'WEEKLY_PATTERN_DATE', msg: '開始日期不能晚於結束日期' };
  }
  const dates = listDateStrs_(startStr, endStr);
  if (dates[dates.length - 1] !== endStr) {
    return { ok: false, code: 'WEEKLY_PATTERN_RANGE', msg: `一次最多產生 ${WEEKLY_PATTERN_MAX_DAYS} 天` };
  }

  const pattern = readWeeklyPattern_();
  if (!pattern.length) {
    return { ok: false, code: 'WEEKLY_PATTERN_EMPTY', msg: '固定班表是空的，請先設定' };
  }

  const templates = readShiftTemplates_();
  const sheet = getShiftSheet();
  const data = sheet.getDataRange().getValues();
  const now = formatDateTime(new Date());
  const by = opts.by || Session.getActiveUser().getEmail();

  // 這段期間每位員工每天已經有的排班
  const taken = new Set();
  let replaced = 0;
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[13] === '已刪除') continue;
    const date = formatDateOnly(row[3]);
    if (date < startStr || date > endStr) continue;
    const employeeId = String(row[1]).trim();
    // 覆蓋：只刪固定班表產生、而且產生之後沒人改過的（建立時間 = 最後修改時間）
    const untouched = String(row[9]) === String(row[11]);
    if (opts.replace && String(row[8]).trim() === WEEKLY_PATTERN_NOTE && untouched) {
      sheet.getRange(i + 1, 12, 1, 3).setValues([[now, by, '已刪除']]);
      replaced++;
      continue;
    }
    taken.add(employeeId + '_' + date);
  }

  const rows = [];
  let skipped = 0;
  const unknownCodes = new Set();
  dates.forEach(date => {
    const weekday = weekdayOfDateStr_(date);
    pattern.filter(p => p.weekday === weekday).forEach(p => {
      const key = p.employeeId + '_' + date;
      if (taken.has(key)) { skipped++; return; }
      const tpl = findShiftTemplate_(p.code, templates);
      if (!tpl) { unknownCodes.add(p.code); return; }
      taken.add(key);
      rows.push([
        'SHIFT-' + Utilities.getUuid(),
        p.employeeId,
        p.employeeName,
        date,
        tpl.name,
        formatTimeOnly(tpl.startTime),
        formatTimeOnly(tpl.endTime),
        p.location,
        WEEKLY_PATTERN_NOTE,
        now, by, now, by,
        '正常',
        tpl.breakMinutes
      ]);
    });
  });

  if (rows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  }

  const result = { ok: true, added: rows.length, skipped: skipped, replaced: replaced };
  if (unknownCodes.size) result.unknownCodes = Array.from(unknownCodes);
  Logger.log(`固定班表產生排班 ${startStr}～${endStr}: ${JSON.stringify(result)}`);
  return result;
}

// ==================== API ====================

/** API：讀固定班表（管理員、排班人員） */
function handleGetWeeklyPattern(params) {
  try {
    const perm = checkSchedulingPermission(params.token);
    if (!perm.ok) return perm;
    return { ok: true, pattern: readWeeklyPattern_() };
  } catch (error) {
    Logger.log(' handleGetWeeklyPattern 錯誤: ' + error.message);
    return { ok: false, msg: error.toString() };
  }
}

/**
 * API：儲存整份固定班表（管理員、排班人員）
 * 參數：pattern（JSON 陣列，每項 { employeeId, employeeName, weekday, code, location }）
 */
function handleSaveWeeklyPattern(params) {
  try {
    const perm = checkSchedulingPermission(params.token);
    if (!perm.ok) return perm;

    let input;
    try {
      input = JSON.parse(params.pattern || '[]');
    } catch (error) {
      return { ok: false, code: 'WEEKLY_PATTERN_INVALID', msg: '固定班表資料格式錯誤' };
    }
    if (!Array.isArray(input) || input.length > 7 * 500) {
      return { ok: false, code: 'WEEKLY_PATTERN_INVALID', msg: '固定班表資料格式錯誤' };
    }

    const templates = readShiftTemplates_();
    const seen = new Set();
    const rows = [];
    for (const item of input) {
      const p = item || {};
      const employeeId = String(p.employeeId || '').trim();
      const weekday = Number(p.weekday);
      const code = String(p.code || '').trim();
      if (!code) continue;                       // 空白 = 那天休息
      if (!employeeId || !Number.isInteger(weekday) || weekday < 1 || weekday > 7) {
        return { ok: false, code: 'WEEKLY_PATTERN_INVALID', msg: '固定班表資料格式錯誤' };
      }
      const tpl = findShiftTemplate_(code, templates);
      if (!tpl) {
        return { ok: false, code: 'WEEKLY_PATTERN_UNKNOWN_CODE', msg: `找不到班別代碼「${code}」` };
      }
      const key = employeeId + '_' + weekday;
      if (seen.has(key)) continue;
      seen.add(key);
      rows.push([employeeId, String(p.employeeName || '').trim().slice(0, 50), weekday, tpl.code,
                 String(p.location || '').trim().slice(0, 50)]);
    }

    const sheet = getWeeklyPatternSheet_();
    const last = sheet.getLastRow();
    if (last > 1) sheet.getRange(2, 1, last - 1, WEEKLY_PATTERN_HEADERS.length).clearContent();
    if (rows.length) sheet.getRange(2, 1, rows.length, WEEKLY_PATTERN_HEADERS.length).setValues(rows);

    return { ok: true, pattern: readWeeklyPattern_() };
  } catch (error) {
    Logger.log(' handleSaveWeeklyPattern 錯誤: ' + error.message);
    return { ok: false, msg: error.toString() };
  }
}

/**
 * API：依固定班表產生排班（管理員、排班人員）
 * 參數：startDate、endDate（yyyy-MM-dd）、replace（'true' 時覆蓋固定班表產生、沒改過的排班）
 */
function handleGenerateShiftsFromPattern(params) {
  try {
    const perm = checkSchedulingPermission(params.token);
    if (!perm.ok) return perm;
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);  // 兩個人同時按，不會排出兩份
    try {
      return generateShiftsFromPattern_(params.startDate, params.endDate, {
        replace: String(params.replace) === 'true',
        by: perm.user ? perm.user.name : ''
      });
    } finally {
      lock.releaseLock();
    }
  } catch (error) {
    Logger.log(' handleGenerateShiftsFromPattern 錯誤: ' + error.message);
    return { ok: false, msg: error.toString() };
  }
}

// ==================== 每週自動排班 ====================

/** 觸發器執行：從今天起往後 4 週，空的日子依固定班表補上 */
function autoGenerateWeeklyShifts() {
  const tz = Session.getScriptTimeZone();
  const start = new Date();
  const end = new Date(start.getTime() + (WEEKLY_PATTERN_AUTO_DAYS - 1) * 24 * 60 * 60 * 1000);
  const result = generateShiftsFromPattern_(
    Utilities.formatDate(start, tz, 'yyyy-MM-dd'),
    Utilities.formatDate(end, tz, 'yyyy-MM-dd'),
    { replace: false, by: '自動排班' }
  );
  Logger.log('每週自動排班: ' + JSON.stringify(result));
  return result;
}

/** 在 Apps Script 編輯器執行一次：每週一凌晨 1 點自動排班（重複執行不會裝兩個） */
function installWeeklyShiftTrigger() {
  const exists = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'autoGenerateWeeklyShifts');
  if (!exists) {
    ScriptApp.newTrigger('autoGenerateWeeklyShifts')
      .timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(1).create();
  }
  const result = autoGenerateWeeklyShifts();
  Logger.log(exists ? '每週自動排班已經裝過了' : '已安裝每週自動排班（每週一 1:00）');
  return result;
}

/** 不想自動排了：在 Apps Script 編輯器執行 */
function removeWeeklyShiftTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'autoGenerateWeeklyShifts')
    .forEach(t => ScriptApp.deleteTrigger(t));
  Logger.log('已移除每週自動排班');
}
