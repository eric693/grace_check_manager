// ShiftPay.gs — 時段計薪
//
// 不同時段不同時薪（例如晚班比早班高）、晚班保障時數、國定假日加倍，
// 照每一段實際打卡算到分鐘。規則存在「時段計薪設定」表，管理員在資料管理頁的「計薪規則」分頁調整，
// 以後員工異動、時薪調整、時段改了，都不用改程式。
//
// 一位員工的規則：
//   rules：[{ name: '晚班', days: [1,2,3,4,5], start: '19:00', rate: 600, minHours: 3 }, …]
//          days：1 = 星期一 … 7 = 星期日；start：表定開始時間；minHours：保障時數（不到就照這個時數算，0 = 不保障）
//   holidayMultiplier：國定假日時薪倍率（預設 2）
//   deductInsurance：要不要扣勞健保（false = 有保但不扣）
//   overtimeExtra：核准的「加班申請」要不要另外加發加班費（預設不要：時段計薪已經照實際打卡算到下班，
//                  再發加班費等於同一段時間付兩次）
//
// 算法（每一段「上班卡 → 下班卡」）：
//   1. 看上班卡時間，對應到當天適用的規則：開始時間最晚、而且上班卡不早於「開始時間前 60 分鐘」的那一條
//   2. 從「表定開始」和「實際上班卡」較晚的那個算到下班卡，精確到分鐘（早到不算，晚到照實際）
//   3. 不到保障時數就照保障時數算
//   4. 國定假日：時薪 × 倍率
//   5. 每段四捨五入到元，月合計 = 各段相加
//
// 有設定規則的員工，薪資計算（calculateHourlySalary）的基本薪資就用這裡算的金額。

const SHEET_SHIFT_PAY = '時段計薪設定';
const SHIFT_PAY_HEADERS = ['員工ID', '員工姓名', '時段規則', '國定假日倍率', '扣勞健保', '更新時間', '更新者', '加班申請另計'];
const SHIFT_PAY_EARLY_MINUTES = 60;    // 上班卡最多比表定早多久，還算是那個時段
const SHIFT_PAY_LONG_SEGMENT_MINUTES = 6 * 60;   // 一段超過這麼久，可能是忘了打休息卡

function getShiftPaySheet_() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(SHEET_SHIFT_PAY);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_SHIFT_PAY);
    sheet.getRange(1, 1, 1, SHIFT_PAY_HEADERS.length).setValues([SHIFT_PAY_HEADERS])
         .setFontWeight('bold').setBackground('#1f4e79').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  } else if (sheet.getRange(1, SHIFT_PAY_HEADERS.length).getValue() === '') {
    sheet.getRange(1, SHIFT_PAY_HEADERS.length).setValue(SHIFT_PAY_HEADERS[SHIFT_PAY_HEADERS.length - 1]);
  }
  return sheet;
}

function shiftPayMinutes_(hhmm) {
  const m = String(hhmm || '').match(/^(\d{1,2}):(\d{2})$/);
  return m ? Number(m[1]) * 60 + Number(m[2]) : null;
}

/** 檢查並整理一份規則；不合法丟出錯誤（訊息給管理員看） */
function normalizeShiftPayConfig_(input) {
  const rulesIn = Array.isArray(input && input.rules) ? input.rules : [];
  if (rulesIn.length > 20) throw new Error('時段最多 20 個');
  const rules = rulesIn.map((r, i) => {
    const label = `第 ${i + 1} 個時段`;
    const name = String(r.name || '').trim().slice(0, 20);
    if (!name) throw new Error(`${label}：請填名稱`);
    const days = (Array.isArray(r.days) ? r.days : []).map(Number).filter(d => Number.isInteger(d) && d >= 1 && d <= 7);
    if (!days.length) throw new Error(`「${name}」：請選擇適用的星期`);
    if (shiftPayMinutes_(r.start) === null) throw new Error(`「${name}」：開始時間格式要是 08:50 這種`);
    const rate = Number(r.rate);
    if (!isFinite(rate) || rate <= 0 || rate > 100000) throw new Error(`「${name}」：時薪要大於 0`);
    const minHours = Number(r.minHours || 0);
    if (!isFinite(minHours) || minHours < 0 || minHours > 24) throw new Error(`「${name}」：保障時數要在 0～24 之間`);
    return { name: name, days: Array.from(new Set(days)).sort(), start: String(r.start).padStart(5, '0'), rate: rate, minHours: minHours };
  });
  const multiplier = input && input.holidayMultiplier !== undefined && input.holidayMultiplier !== '' ? Number(input.holidayMultiplier) : 2;
  if (!isFinite(multiplier) || multiplier < 1 || multiplier > 5) throw new Error('國定假日倍率要在 1～5 之間');
  return { rules: rules, holidayMultiplier: multiplier, deductInsurance: !(input && input.deductInsurance === false),
           overtimeExtra: !!(input && input.overtimeExtra === true) };
}

/** 一位員工的時段計薪設定；沒設定（或沒有任何時段）回傳 null */
function readShiftPayConfig_(employeeId) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_SHIFT_PAY);
  if (!sheet || sheet.getLastRow() < 2) return null;
  const values = sheet.getDataRange().getValues();
  const id = String(employeeId || '').trim();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim() !== id) continue;
    try {
      const config = normalizeShiftPayConfig_({
        rules: JSON.parse(values[i][2] || '[]'),
        holidayMultiplier: values[i][3],
        deductInsurance: String(values[i][4]).trim() !== '否',
        overtimeExtra: String(values[i][7] || '').trim() === '是'
      });
      return config.rules.length ? config : null;
    } catch (error) {
      Logger.log(` 時段計薪設定第 ${i + 1} 列有誤: ${error.message}`);
      return null;
    }
  }
  return null;
}

/** 這一段該套哪一條規則 */
function pickShiftPayRule_(rules, weekday, startMinutes) {
  const candidates = rules.filter(r => r.days.indexOf(weekday) !== -1)
    .sort((a, b) => shiftPayMinutes_(a.start) - shiftPayMinutes_(b.start));
  if (!candidates.length) return null;
  let picked = null;
  candidates.forEach(r => {
    if (startMinutes >= shiftPayMinutes_(r.start) - SHIFT_PAY_EARLY_MINUTES) picked = r;
  });
  return picked || candidates[0];   // 比所有時段都早很多：當成當天第一個時段
}

/**
 * 依打卡算一位員工一個月的時段薪資
 * @param {Object} config readShiftPayConfig_ 的結果
 * @param {Array} [attendance] getEmployeeMonthlyAttendanceInternal 的結果（已經讀過就傳進來）
 * @returns {{ totalPay, totalMinutes, billedMinutes, days: Array, warnings: Array }}
 */
function computeShiftPay_(employeeId, yearMonth, config, attendance) {
  const records = attendance || getEmployeeMonthlyAttendanceInternal(employeeId, yearMonth) || [];
  const days = [];
  let totalPay = 0, totalMinutes = 0, billedMinutes = 0;
  const warnings = [];

  records.forEach(day => {
    const p = String(day.date).split('-').map(Number);
    const jsDay = new Date(p[0], p[1] - 1, p[2], 12).getDay();
    const weekday = jsDay === 0 ? 7 : jsDay;
    // 國定假日而且設定「加倍計薪」才乘倍率（補假日要不要加倍，在國定假日清單逐日設定，見 Holidays.gs）
    const holiday = typeof isDoublePayHoliday_ === 'function' ? isDoublePayHoliday_(day.date)
      : (typeof getDateType === 'function' && getDateType(day.date) === 'holiday');
    const items = [];

    (day.segments || []).forEach(seg => {
      const inMin = shiftPayMinutes_(seg.start);
      let outMin = shiftPayMinutes_(seg.end);
      if (inMin === null || outMin === null) return;
      if (outMin < inMin) outMin += 24 * 60;   // 跨午夜
      const rule = pickShiftPayRule_(config.rules, weekday, inMin);
      const flags = [];
      if (!rule) {
        flags.push('NO_RULE');
        warnings.push({ date: day.date, code: 'NO_RULE', segment: `${seg.start}–${seg.end}` });
        items.push({ rule: '', in: seg.start, out: seg.end, from: seg.start, minutes: 0, billed: 0, rate: 0, pay: 0, flags: flags });
        return;
      }
      const fromMin = Math.max(inMin, shiftPayMinutes_(rule.start));
      const minutes = Math.max(0, outMin - fromMin);
      let billed = minutes;
      if (rule.minHours > 0 && minutes > 0 && minutes < rule.minHours * 60) {
        billed = rule.minHours * 60;
        flags.push('MIN_HOURS');
      }
      if (minutes > SHIFT_PAY_LONG_SEGMENT_MINUTES) {
        flags.push('LONG');
        warnings.push({ date: day.date, code: 'LONG', segment: `${seg.start}–${seg.end}` });
      }
      const rate = rule.rate * (holiday ? config.holidayMultiplier : 1);
      if (holiday) flags.push('HOLIDAY');
      const pay = Math.round(billed / 60 * rate);
      const fromText = `${String(Math.floor(fromMin / 60) % 24).padStart(2, '0')}:${String(fromMin % 60).padStart(2, '0')}`;
      items.push({ rule: rule.name, in: seg.start, out: seg.end, from: fromText, minutes: minutes, billed: billed, rate: rate, pay: pay, flags: flags });
      totalPay += pay;
      totalMinutes += minutes;
      billedMinutes += billed;
    });

    if (day.unpaired) warnings.push({ date: day.date, code: 'UNPAIRED' });
    days.push({ date: day.date, weekday: weekday, holiday: holiday, items: items, pay: items.reduce((s, x) => s + x.pay, 0) });
  });

  return { totalPay: totalPay, totalMinutes: totalMinutes, billedMinutes: billedMinutes, days: days, warnings: warnings };
}

/** 給薪資計算用：有規則就回傳結果，沒有回傳 null */
function computeShiftPayForMonth_(employeeId, yearMonth) {
  const config = readShiftPayConfig_(employeeId);
  if (!config) return null;
  const result = computeShiftPay_(employeeId, yearMonth, config);
  result.config = config;
  return result;
}

// ==================== API ====================

/** API：一位員工的時段計薪設定（沒設定回傳空規則） */
function handleGetShiftPayConfig(params) {
  try {
    const admin = requireRecordsAdmin_(params.token);
    if (!admin.ok) return admin;
    const config = readShiftPayConfig_(params.employeeId);
    return { ok: true, config: config || { rules: [], holidayMultiplier: 2, deductInsurance: true, overtimeExtra: false } };
  } catch (error) {
    return { ok: false, msg: '讀取計薪規則失敗：' + error.message };
  }
}

/**
 * API：儲存一位員工的時段計薪設定
 * 參數：employeeId、config（JSON：{ rules, holidayMultiplier, deductInsurance }）。rules 空陣列 = 取消時段計薪
 */
function handleSaveShiftPayConfig(params) {
  try {
    const admin = requireRecordsAdmin_(params.token);
    if (!admin.ok) return admin;
    const emp = findEmployeeForRecords_(params.employeeId);
    if (!emp) return { ok: false, code: 'RECORDS_NO_EMPLOYEE', msg: '找不到這位員工' };

    let config;
    try {
      config = normalizeShiftPayConfig_(JSON.parse(params.config || '{}'));
    } catch (error) {
      return { ok: false, code: 'SHIFT_PAY_INVALID', msg: error.message };
    }

    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const sheet = getShiftPaySheet_();
      const values = sheet.getDataRange().getValues();
      let rowNumber = -1;
      for (let i = 1; i < values.length; i++) {
        if (String(values[i][0]).trim() === emp.userId) { rowNumber = i + 1; break; }
      }
      const row = [emp.userId, emp.name, JSON.stringify(config.rules), config.holidayMultiplier,
                   config.deductInsurance ? '是' : '否', new Date(), admin.user.name, config.overtimeExtra ? '是' : '否'];
      if (rowNumber < 0) sheet.getRange(sheet.getLastRow() + 1, 1, 1, row.length).setValues([row]);
      else sheet.getRange(rowNumber, 1, 1, row.length).setValues([row]);
    } finally {
      lock.releaseLock();
    }

    params.employeeName = emp.name;
    params.changes = config.rules.map(r => `${r.name} ${r.start} 時薪${r.rate}${r.minHours ? ' 保障' + r.minHours + '小時' : ''}`).join('; ') +
      `; 國定假日×${config.holidayMultiplier}; ${config.deductInsurance ? '扣' : '不扣'}勞健保; 加班申請${config.overtimeExtra ? '另計' : '不另計'}`;
    params.config = '';
    return { ok: true, config: config };
  } catch (error) {
    Logger.log(' handleSaveShiftPayConfig 錯誤: ' + error.message);
    return { ok: false, msg: '儲存計薪規則失敗：' + error.message };
  }
}
