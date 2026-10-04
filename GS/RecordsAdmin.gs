// RecordsAdmin.gs
//
// 資料管理頁（records.html）的後端：管理員不用打開試算表，直接在網頁上
//   ・查、新增、修改、刪除「打卡紀錄」
//   ・查、調整「假期餘額」
//
// 打卡紀錄沒有 ID 欄，用「列號 + 指紋」指定要改哪一筆：指紋是打卡時間、員工ID、上下班，
// 列表載入後如果有人在別處改了或刪了資料（列號位移），指紋對不上就拒絕，不會改錯筆。
//
// 這幾個 action 在 Main.gs 的 ROUTE_ACCESS 設為管理員限定，操作會寫進「管理操作記錄」。

const RECORDS_MAX_DAYS = 62;   // 打卡紀錄一次最多查兩個月

// 打卡紀錄欄位（0-based）：A 時間、B 員工ID、C 部門、D 姓名、E 上下班、F GPS、G 地點、
// H 標記（補打卡／系統虛擬卡）、I 審核（v／?／x）、J 備註
const PUNCH_COL = { time: 0, userId: 1, dept: 2, name: 3, type: 4, gps: 5, location: 6, mark: 7, audit: 8, note: 9 };

// 假期餘額表（getLeaveBalanceSheet）：A 員工ID、B 姓名、C 到職日、D～R 各假別（小時）、S 更新時間
const LEAVE_BALANCE_TYPES = [
  'ANNUAL_LEAVE', 'SICK_LEAVE', 'PERSONAL_LEAVE', 'BEREAVEMENT_LEAVE', 'MARRIAGE_LEAVE',
  'MATERNITY_LEAVE', 'PATERNITY_LEAVE', 'HOSPITALIZATION_LEAVE', 'MENSTRUAL_LEAVE', 'FAMILY_CARE_LEAVE',
  'OFFICIAL_LEAVE', 'WORK_INJURY_LEAVE', 'NATURAL_DISASTER_LEAVE', 'COMP_TIME_OFF', 'ABSENCE_WITHOUT_LEAVE'
];
const LEAVE_BALANCE_FIRST_COL = 4;                                   // D
const LEAVE_BALANCE_UPDATED_COL = LEAVE_BALANCE_FIRST_COL + LEAVE_BALANCE_TYPES.length;  // S

/** 儲存格的值是不是日期（試算表讀出來的日期、時間都是 Date） */
function isDateValue_(value) {
  return !!value && typeof value.getTime === 'function' && !isNaN(value.getTime());
}

function requireRecordsAdmin_(token) {
  const session = checkSession_(token);
  if (!session.ok || !session.user) {
    return { ok: false, code: 'ERR_SESSION_INVALID', msg: '未授權或 session 已過期' };
  }
  if (session.user.dept !== '管理員') {
    return { ok: false, code: 'PERMISSION_DENIED', msg: '需要管理員權限' };
  }
  return { ok: true, user: session.user };
}

// ==================== 打卡紀錄 ====================

function punchFingerprint_(row) {
  const t = isDateValue_(row[PUNCH_COL.time]) ? row[PUNCH_COL.time] : new Date(row[PUNCH_COL.time]);
  return [isNaN(t.getTime()) ? '' : t.getTime(), String(row[PUNCH_COL.userId]).trim(), String(row[PUNCH_COL.type]).trim()].join('|');
}

/** "yyyy-MM-dd" + "HH:mm" → Date（腳本時區）；格式不對回傳 null */
function parsePunchDateTime_(dateStr, timeStr) {
  const d = String(dateStr || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  const t = String(timeStr || '').match(/^(\d{1,2}):(\d{2})$/);
  if (!d || !t) return null;
  const h = Number(t[1]), m = Number(t[2]);
  if (h > 23 || m > 59) return null;
  const date = new Date(Number(d[1]), Number(d[2]) - 1, Number(d[3]), h, m, 0);
  return isNaN(date.getTime()) ? null : date;
}

/** 員工名單裡的姓名（有手動姓名用手動的）與部門 */
function findEmployeeForRecords_(employeeId) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_EMPLOYEES);
  const values = sheet.getDataRange().getValues();
  const id = String(employeeId || '').trim();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim() === id) {
      return { userId: id, name: String(values[i][8] || values[i][2] || '').trim(), dept: values[i][5] || '' };
    }
  }
  return null;
}

/** 這一列的狀態，給畫面顯示：一般、補打卡已核准、補打卡審核中、補打卡駁回、系統虛擬卡 */
function punchRowStatus_(row) {
  const mark = String(row[PUNCH_COL.mark] || '').trim();
  const audit = String(row[PUNCH_COL.audit] || '').trim();
  if (mark === '系統虛擬卡') return 'VIRTUAL';
  if (mark === '補打卡') {
    if (audit === 'v') return 'ADJUST_APPROVED';
    if (audit === 'x') return 'ADJUST_REJECTED';
    return 'ADJUST_PENDING';
  }
  return 'NORMAL';
}

/**
 * API：查打卡紀錄
 * 參數：startDate、endDate（yyyy-MM-dd，最多 62 天）、employeeId（可省略 = 全部）
 */
function handleAdminListPunches(params) {
  try {
    const admin = requireRecordsAdmin_(params.token);
    if (!admin.ok) return admin;

    const start = parsePunchDateTime_(params.startDate, '00:00');
    const end = parsePunchDateTime_(params.endDate, '23:59');
    if (!start || !end || start > end) {
      return { ok: false, code: 'RECORDS_DATE', msg: '請選擇正確的日期範圍' };
    }
    if ((end - start) / 86400000 > RECORDS_MAX_DAYS) {
      return { ok: false, code: 'RECORDS_RANGE', msg: `一次最多查 ${RECORDS_MAX_DAYS} 天` };
    }
    const endMs = end.getTime() + 59999;   // 含 23:59 這一整分鐘

    const tz = Session.getScriptTimeZone();
    const nameMap = getEmployeeNameMap_();
    const employeeId = String(params.employeeId || '').trim();
    const data = SpreadsheetApp.getActive().getSheetByName(SHEET_ATTENDANCE).getDataRange().getValues();
    const records = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[PUNCH_COL.time]) continue;
      const time = isDateValue_(row[PUNCH_COL.time]) ? row[PUNCH_COL.time] : new Date(row[PUNCH_COL.time]);
      if (isNaN(time.getTime()) || time < start || time.getTime() > endMs) continue;
      const uid = String(row[PUNCH_COL.userId]).trim();
      if (employeeId && uid !== employeeId) continue;

      records.push({
        row: i + 1,
        key: punchFingerprint_(row),
        date: Utilities.formatDate(time, tz, 'yyyy-MM-dd'),
        time: Utilities.formatDate(time, tz, 'HH:mm'),
        userId: uid,
        name: nameMap[uid] || String(row[PUNCH_COL.name] || ''),
        type: String(row[PUNCH_COL.type] || ''),
        location: String(row[PUNCH_COL.location] || ''),
        note: String(row[PUNCH_COL.note] || ''),
        status: punchRowStatus_(row)
      });
    }

    records.sort((a, b) => (a.date + a.time).localeCompare(b.date + b.time) || a.name.localeCompare(b.name));
    return { ok: true, records: records };
  } catch (error) {
    Logger.log(' handleAdminListPunches 錯誤: ' + error.message);
    return { ok: false, msg: error.toString() };
  }
}

function validatePunchInput_(params) {
  const time = parsePunchDateTime_(params.date, params.time);
  if (!time) return { ok: false, code: 'RECORDS_DATE', msg: '日期或時間格式不正確' };
  const type = String(params.type || '').trim();
  if (type !== '上班' && type !== '下班') return { ok: false, code: 'ERR_INVALID_PUNCH_TYPE', msg: '打卡類型要是上班或下班' };
  const note = String(params.note || '').trim();
  if (note.length > 200) return { ok: false, code: 'RECORDS_NOTE', msg: '備註最多 200 個字' };
  return { ok: true, time: time, type: type, location: String(params.location || '').trim().slice(0, 50), note: note };
}

/**
 * API：管理員新增一筆打卡（直接算數，不用再審核）
 * 參數：employeeId、date、time、type、location、note
 */
function handleAdminAddPunch(params) {
  try {
    const admin = requireRecordsAdmin_(params.token);
    if (!admin.ok) return admin;
    const input = validatePunchInput_(params);
    if (!input.ok) return input;
    const emp = findEmployeeForRecords_(params.employeeId);
    if (!emp) return { ok: false, code: 'RECORDS_NO_EMPLOYEE', msg: '找不到這位員工' };

    const note = '管理員新增（' + admin.user.name + '）' + (input.note ? '：' + input.note : '');
    const row = [input.time, emp.userId, emp.dept, emp.name, input.type, '', input.location, '', '', note];

    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_ATTENDANCE);
      sh.getRange(sh.getLastRow() + 1, 1, 1, row.length).setValues([row]);
    } finally {
      lock.releaseLock();
    }
    params.employeeName = emp.name;   // 管理操作記錄要看得到是誰的卡
    return { ok: true };
  } catch (error) {
    Logger.log(' handleAdminAddPunch 錯誤: ' + error.message);
    return { ok: false, msg: error.toString() };
  }
}

/** 找到列號且指紋相符的那一列；不符回傳錯誤 */
function lockedPunchRow_(sheet, rowNumber, key) {
  const row = Number(rowNumber);
  if (!Number.isInteger(row) || row < 2 || row > sheet.getLastRow()) {
    return { ok: false, code: 'RECORDS_STALE', msg: '這筆資料已經變動，請重新查詢' };
  }
  const values = sheet.getRange(row, 1, 1, 10).getValues()[0];
  if (punchFingerprint_(values) !== String(key || '')) {
    return { ok: false, code: 'RECORDS_STALE', msg: '這筆資料已經變動，請重新查詢' };
  }
  return { ok: true, row: row, values: values };
}

/**
 * API：修改一筆打卡（時間、上下班、地點、備註）
 * 參數：row、key（列表給的）、date、time、type、location、note
 */
function handleAdminUpdatePunch(params) {
  try {
    const admin = requireRecordsAdmin_(params.token);
    if (!admin.ok) return admin;
    const input = validatePunchInput_(params);
    if (!input.ok) return input;

    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_ATTENDANCE);
      const found = lockedPunchRow_(sh, params.row, params.key);
      if (!found.ok) return found;
      const old = found.values;
      sh.getRange(found.row, PUNCH_COL.time + 1).setValue(input.time);
      sh.getRange(found.row, PUNCH_COL.type + 1).setValue(input.type);
      sh.getRange(found.row, PUNCH_COL.location + 1).setValue(input.location);
      sh.getRange(found.row, PUNCH_COL.note + 1).setValue(input.note);

      const tz = Session.getScriptTimeZone();
      const oldTime = isDateValue_(old[PUNCH_COL.time]) ? Utilities.formatDate(old[PUNCH_COL.time], tz, 'yyyy-MM-dd HH:mm') : String(old[PUNCH_COL.time]);
      params.employeeId = String(old[PUNCH_COL.userId]);
      params.employeeName = String(old[PUNCH_COL.name] || '');
      params.before = `${oldTime} ${old[PUNCH_COL.type]} ${old[PUNCH_COL.location] || ''}`.trim();
    } finally {
      lock.releaseLock();
    }
    return { ok: true };
  } catch (error) {
    Logger.log(' handleAdminUpdatePunch 錯誤: ' + error.message);
    return { ok: false, msg: error.toString() };
  }
}

/**
 * API：刪除一筆打卡
 * 參數：row、key
 */
function handleAdminDeletePunch(params) {
  try {
    const admin = requireRecordsAdmin_(params.token);
    if (!admin.ok) return admin;

    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_ATTENDANCE);
      const found = lockedPunchRow_(sh, params.row, params.key);
      if (!found.ok) return found;
      const old = found.values;
      sh.deleteRow(found.row);

      const tz = Session.getScriptTimeZone();
      const oldTime = isDateValue_(old[PUNCH_COL.time]) ? Utilities.formatDate(old[PUNCH_COL.time], tz, 'yyyy-MM-dd HH:mm') : String(old[PUNCH_COL.time]);
      params.employeeId = String(old[PUNCH_COL.userId]);
      params.employeeName = String(old[PUNCH_COL.name] || '');
      params.deleted = `${oldTime} ${old[PUNCH_COL.type]} ${old[PUNCH_COL.location] || ''} ${old[PUNCH_COL.note] || ''}`.trim();
    } finally {
      lock.releaseLock();
    }
    return { ok: true };
  } catch (error) {
    Logger.log(' handleAdminDeletePunch 錯誤: ' + error.message);
    return { ok: false, msg: error.toString() };
  }
}

// ==================== 假期餘額 ====================

/**
 * API：所有員工的假期餘額（小時）。名單上有、但還沒有餘額資料的員工也列出來（exists: false）
 */
function handleAdminGetLeaveBalances(params) {
  try {
    const admin = requireRecordsAdmin_(params.token);
    if (!admin.ok) return admin;

    const tz = Session.getScriptTimeZone();
    const values = getLeaveBalanceSheet().getDataRange().getValues();
    const nameMap = getEmployeeNameMap_();
    const list = [];
    const seen = new Set();

    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const uid = String(row[0] || '').trim();
      if (!uid || seen.has(uid)) continue;
      seen.add(uid);
      const balances = {};
      LEAVE_BALANCE_TYPES.forEach((type, j) => { balances[type] = Number(row[LEAVE_BALANCE_FIRST_COL - 1 + j]) || 0; });
      const hire = row[2];
      const updated = row[LEAVE_BALANCE_UPDATED_COL - 1];
      list.push({
        userId: uid,
        name: nameMap[uid] || String(row[1] || ''),
        hireDate: isDateValue_(hire) ? Utilities.formatDate(hire, tz, 'yyyy-MM-dd') : String(hire || ''),
        balances: balances,
        updatedAt: isDateValue_(updated) ? Utilities.formatDate(updated, tz, 'yyyy-MM-dd HH:mm') : String(updated || ''),
        exists: true
      });
    }

    // 還沒有餘額資料的在職員工
    const empValues = SpreadsheetApp.getActive().getSheetByName(SHEET_EMPLOYEES).getDataRange().getValues();
    for (let i = 1; i < empValues.length; i++) {
      const uid = String(empValues[i][0] || '').trim();
      if (!uid || seen.has(uid) || String(empValues[i][7] || '').trim() === '停用') continue;
      seen.add(uid);
      const balances = {};
      LEAVE_BALANCE_TYPES.forEach(type => { balances[type] = 0; });
      list.push({ userId: uid, name: nameMap[uid] || String(empValues[i][2] || ''), hireDate: '', balances: balances, updatedAt: '', exists: false });
    }

    list.sort((a, b) => a.name.localeCompare(b.name));
    return { ok: true, types: LEAVE_BALANCE_TYPES, employees: list };
  } catch (error) {
    Logger.log(' handleAdminGetLeaveBalances 錯誤: ' + error.message);
    return { ok: false, msg: error.toString() };
  }
}

/**
 * API：設定一位員工的假期餘額
 * 參數：employeeId、balances（JSON：{ 假別: 小時 }，只改有給的假別）、hireDate（yyyy-MM-dd，可省略）
 */
function handleAdminSetLeaveBalance(params) {
  try {
    const admin = requireRecordsAdmin_(params.token);
    if (!admin.ok) return admin;

    const emp = findEmployeeForRecords_(params.employeeId);
    if (!emp) return { ok: false, code: 'RECORDS_NO_EMPLOYEE', msg: '找不到這位員工' };

    let input;
    try {
      input = JSON.parse(params.balances || '{}');
    } catch (e) {
      return { ok: false, code: 'RECORDS_INVALID', msg: '假期資料格式錯誤' };
    }
    const updates = {};
    for (const type of Object.keys(input || {})) {
      if (LEAVE_BALANCE_TYPES.indexOf(type) === -1) {
        return { ok: false, code: 'RECORDS_INVALID', msg: '假別不正確：' + type };
      }
      const hours = Number(input[type]);
      if (!isFinite(hours) || hours < 0 || hours > 9999) {
        return { ok: false, code: 'RECORDS_HOURS', msg: '時數要是 0～9999 的數字' };
      }
      updates[type] = Math.round(hours * 100) / 100;
    }
    let hireDate = null;
    if (params.hireDate) {
      hireDate = parsePunchDateTime_(params.hireDate, '00:00');
      if (!hireDate) return { ok: false, code: 'RECORDS_DATE', msg: '到職日格式不正確' };
    }

    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const sheet = getLeaveBalanceSheet();
      const values = sheet.getDataRange().getValues();
      let rowNumber = -1;
      for (let i = 1; i < values.length; i++) {
        if (String(values[i][0]).trim() === emp.userId) { rowNumber = i + 1; break; }
      }

      const before = [];
      if (rowNumber < 0) {
        // 還沒有餘額資料：新增一列，沒給的假別是 0
        const row = [emp.userId, emp.name, hireDate || ''];
        LEAVE_BALANCE_TYPES.forEach(type => row.push(updates[type] !== undefined ? updates[type] : 0));
        row.push(new Date());
        sheet.getRange(sheet.getLastRow() + 1, 1, 1, row.length).setValues([row]);
      } else {
        const current = values[rowNumber - 1];
        const rowValues = LEAVE_BALANCE_TYPES.map((type, j) => {
          const old = Number(current[LEAVE_BALANCE_FIRST_COL - 1 + j]) || 0;
          if (updates[type] === undefined) return old;
          if (updates[type] !== old) before.push(`${type} ${old}→${updates[type]}`);
          return updates[type];
        });
        sheet.getRange(rowNumber, LEAVE_BALANCE_FIRST_COL, 1, rowValues.length).setValues([rowValues]);
        if (hireDate) sheet.getRange(rowNumber, 3).setValue(hireDate);
        sheet.getRange(rowNumber, LEAVE_BALANCE_UPDATED_COL).setValue(new Date());
      }

      params.employeeName = emp.name;
      params.changes = before.length ? before.join('; ') : (rowNumber < 0 ? '新建餘額資料' : '沒有變更');
    } finally {
      lock.releaseLock();
    }
    return { ok: true };
  } catch (error) {
    Logger.log(' handleAdminSetLeaveBalance 錯誤: ' + error.message);
    return { ok: false, msg: error.toString() };
  }
}
