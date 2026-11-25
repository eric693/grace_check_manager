// PayrollRules.gs
//
// 店家自己的計薪規則，在薪資算完之後（calculateMonthlySalary）再套上去：
//
//   （全勤獎金已取消：一律為 0，不再判斷遲到、忘卡、請假或排班時數）
//
//   正職（薪資類型「月薪」）
//     餐費：實際工時滿 mealMinHours 小時的天數 × mealPerDay
//     生日禮金：生日在當月，且到職滿 birthdayMinTenureMonths 個月
//     加班費的時薪改用（基本薪資 + 伙食費 + 職務加給）÷ 30 ÷ 8（在 SalaryManagement.gs）
//
//   兼職（薪資類型「時薪」）
//     餐費：排班滿 mealMinHours 小時、而且當天有出勤的天數 × mealPerDay
//     生日禮金：同正職，金額不同
//
//   兩種都有、每張薪資單各自填的：銷售獎金、預支抵扣、手動加減項目、薪資單備註。
//   這些存在「月薪資記錄」的「計薪調整」欄（JSON），重新計算時會沿用，
//   員工打開自己的薪資單（會觸發重算）也不會把管理員填的東西洗掉。
//
// 金額與門檻都在「薪資規則」設定（SystemSettings.gs 的 DEFAULT_PAYROLL_RULES）。

const PAYROLL_ADJUSTMENTS_COLUMN = '計薪調整';

// 生日說明：後端回傳 { code, params }，前端依語系翻譯（i18n 同名鍵）。
// 這裡的中文只拿來寫進試算表，給直接看試算表的人讀。
const PAYROLL_MESSAGES_ZH = {
  PAYROLL_BDAY_GIVEN: '當月壽星（到職 {hireDate}，依{source}）',
  PAYROLL_BDAY_NO_HIRE: '當月壽星，但沒有到職日期，未發生日禮金',
  PAYROLL_BDAY_TENURE: '當月壽星，到職未滿 {months} 個月（到職 {hireDate}）',
  PAYROLL_HIRE_SRC_CONFIG: '薪資設定',
  PAYROLL_HIRE_SRC_EMPLOYEES: '員工名單',
  PAYROLL_HIRE_SRC_ACCOUNT: '帳號建立日'
};

/** { code, params } → 中文 */
function payrollMessageZh_(msg) {
  if (!msg) return '';
  const params = msg.params || {};
  return String(PAYROLL_MESSAGES_ZH[msg.code] || msg.code).replace(/\{(\w+)\}/g, (m, key) => {
    const value = params[key];
    if (Array.isArray(value)) return value.join('、');
    if (typeof value === 'string' && PAYROLL_MESSAGES_ZH[value]) return PAYROLL_MESSAGES_ZH[value];
    return value === undefined ? m : String(value);
  });
}

const PAYROLL_MAX_MANUAL_ITEMS = 20;
const PAYROLL_MAX_AMOUNT = 1000000;
const PAYROLL_MAX_NOTE = 500;

function getPayrollRules_() {
  const rules = (typeof getSalaryRules_ === 'function') ? getSalaryRules_().payrollRules : null;
  return rules || DEFAULT_PAYROLL_RULES;
}

/** Date、"2026-03-05"、"2026/3/5" → Date（當地日期的 00:00）；看不懂回傳 null */
function payrollParseDate_(value) {
  if (!value) return null;
  if (value instanceof Date || Object.prototype.toString.call(value) === '[object Date]') {
    return isNaN(value.getTime()) ? null : new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }
  const m = String(value).trim().match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
  if (!m) return null;
  const date = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  return isNaN(date.getTime()) ? null : date;
}

function payrollDateKey_(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
}

/** "HH:mm" → 分鐘；看不懂回傳 null */
function payrollTimeToMinutes_(text) {
  const m = String(text || '').match(/^(\d{1,2}):(\d{2})/);
  return m ? Number(m[1]) * 60 + Number(m[2]) : null;
}

/** 一筆排班的應工作分鐘（休假、沒有時間的班 = 0） */
function payrollScheduledMinutes_(shift) {
  if (!shift) return 0;
  if (/休/.test(String(shift.shiftType || ''))) return 0;
  const start = payrollTimeToMinutes_(shift.startTime);
  const end = payrollTimeToMinutes_(shift.endTime);
  if (start === null || end === null || start === end) return 0;
  return calcShiftMinutes_(shift.startTime, shift.endTime, false, Number(shift.breakMinutes) || 0).workMinutes;
}

// ==================== 計薪調整（每張薪資單各自填的） ====================

/**
 * 整理計薪調整：去掉亂填的值，金額四捨五入到元。
 * 回傳 { ok, adjustments } 或 { ok: false, msg }
 */
function normalizePayrollAdjustments_(input) {
  const src = input && typeof input === 'object' ? input : {};
  const money = (value, label) => {
    if (value === null || value === undefined || value === '') return { ok: true, value: null };
    const n = Number(value);
    if (!isFinite(n) || n < 0 || n > PAYROLL_MAX_AMOUNT) {
      return { ok: false, msg: `${label}要介於 0～${PAYROLL_MAX_AMOUNT}` };
    }
    return { ok: true, value: Math.round(n) };
  };

  const sales = money(src.salesBonus, '銷售獎金');
  if (!sales.ok) return sales;
  const advance = money(src.advanceDeduction, '預支抵扣');
  if (!advance.ok) return advance;

  const rawItems = Array.isArray(src.manualItems) ? src.manualItems : [];
  if (rawItems.length > PAYROLL_MAX_MANUAL_ITEMS) {
    return { ok: false, msg: `手動項目最多 ${PAYROLL_MAX_MANUAL_ITEMS} 筆` };
  }
  const manualItems = [];
  for (let i = 0; i < rawItems.length; i++) {
    const item = rawItems[i] || {};
    const name = String(item.name || '').trim();
    const amount = money(item.amount, `第 ${i + 1} 筆手動項目的金額`);
    if (!amount.ok) return amount;
    if (!name && !amount.value) continue;            // 整列空白就略過
    if (!name || name.length > 30) return { ok: false, msg: `第 ${i + 1} 筆手動項目要填名稱（最多 30 字）` };
    manualItems.push({ name: name, amount: amount.value || 0, type: item.type === 'sub' ? 'sub' : 'add' });
  }

  const note = String(src.note || '').trim();
  if (note.length > PAYROLL_MAX_NOTE) return { ok: false, msg: `備註最多 ${PAYROLL_MAX_NOTE} 字` };

  return {
    ok: true,
    adjustments: {
      salesBonus: sales.value || 0,
      advanceDeduction: advance.value,       // null = 用已核准的預支申請自動帶入
      manualItems: manualItems,
      note: note
    }
  };
}

/** 讀出這張薪資單上次存的計薪調整；沒有就回傳 null */
function readSavedPayrollAdjustments_(employeeId, yearMonth) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_MONTHLY_SALARY_ENHANCED);
    if (!sheet || sheet.getLastRow() < 2) return null;
    const data = sheet.getDataRange().getValues();
    const col = data[0].map(h => String(h).trim()).indexOf(PAYROLL_ADJUSTMENTS_COLUMN);
    if (col === -1) return null;

    const salaryId = `SAL-${yearMonth}-${employeeId}`;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) !== salaryId) continue;
      const raw = String(data[i][col] || '').trim();
      if (!raw) return null;
      const parsed = JSON.parse(raw);
      const checked = normalizePayrollAdjustments_(parsed);
      if (!checked.ok) return null;
      // 生日說明的代碼一併留著：只存欄位、不重算的存檔路徑才不會把它們弄丟
      if (parsed && parsed.messages) checked.adjustments.messages = parsed.messages;
      return checked.adjustments;
    }
  } catch (error) {
    Logger.log(' 讀取計薪調整失敗（改用空白）: ' + error);
  }
  return null;
}

// ==================== 資料來源 ====================

/** 當月已核准、日期在當月的預支金額 */
function sumApprovedAdvances_(employeeId, yearMonth) {
  if (typeof SHEET_EXPENSE === 'undefined') return 0;
  const data = getSheetValues_(SHEET_EXPENSE);
  let total = 0;
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (String(row[EXPENSE_COL.EMPLOYEE_ID]).trim() !== String(employeeId).trim()) continue;
    if (String(row[EXPENSE_COL.TYPE]) !== 'advance') continue;
    if (String(row[EXPENSE_COL.STATUS]) !== EXPENSE_STATUS.APPROVED) continue;
    const date = payrollParseDate_(row[EXPENSE_COL.DATE]);
    if (!date || payrollDateKey_(date).substring(0, 7) !== yearMonth) continue;
    total += Number(row[EXPENSE_COL.AMOUNT]) || 0;
  }
  return Math.round(total);
}

/** 生日（員工基本資料 F 欄） */
function getEmployeeBirthday_(employeeId) {
  const sheetName = (typeof SHEET_EMPLOYEE_INFO !== 'undefined') ? SHEET_EMPLOYEE_INFO : '員工基本資料';
  const data = getSheetValues_(sheetName);
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(employeeId).trim()) return payrollParseDate_(data[i][5]);
  }
  return null;
}

/**
 * 到職日：薪資設定的「到職日期」→ 員工名單的到職日期 → 員工名單的建立時間（帳號建立日，保守估計）
 * @returns {{ date: Date|null, source: string }}
 */
function getEmployeeHireDate_(employeeId, config) {
  const fromConfig = payrollParseDate_(config && config['到職日期']);
  if (fromConfig) return { date: fromConfig, source: 'PAYROLL_HIRE_SRC_CONFIG' };

  const data = getSheetValues_(SHEET_EMPLOYEES);
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][EMPLOYEE_COL.USER_ID]).trim() !== String(employeeId).trim()) continue;
    const hire = payrollParseDate_(data[i][EMPLOYEE_COL.HIRE_DATE]);
    if (hire) return { date: hire, source: 'PAYROLL_HIRE_SRC_EMPLOYEES' };
    const created = payrollParseDate_(data[i][EMPLOYEE_COL.CREATED]);
    if (created) return { date: created, source: 'PAYROLL_HIRE_SRC_ACCOUNT' };
  }
  return { date: null, source: '' };
}

// ==================== 各項規則 ====================

/** 生日禮金 */
function evaluateBirthdayGift_(amount, rules, employeeId, config, yearMonth) {
  const birthday = getEmployeeBirthday_(employeeId);
  const [y, m] = yearMonth.split('-').map(Number);
  if (!birthday) return { amount: 0, message: null };
  if (birthday.getMonth() + 1 !== m) return { amount: 0, message: null };

  const hire = getEmployeeHireDate_(employeeId, config);
  if (!hire.date) return { amount: 0, message: { code: 'PAYROLL_BDAY_NO_HIRE', params: {} } };

  const monthEnd = new Date(y, m, 0);
  const eligibleFrom = new Date(hire.date.getFullYear(), hire.date.getMonth() + rules.birthdayMinTenureMonths, hire.date.getDate());
  if (eligibleFrom > monthEnd) {
    return { amount: 0, message: { code: 'PAYROLL_BDAY_TENURE',
                                   params: { months: rules.birthdayMinTenureMonths, hireDate: payrollDateKey_(hire.date) } } };
  }
  return { amount: amount, message: { code: 'PAYROLL_BDAY_GIVEN',
                                      params: { hireDate: payrollDateKey_(hire.date), source: hire.source } } };
}

// ==================== 套用 ====================

/**
 * 把計薪規則套到一筆已經算好的薪資上（直接修改 data）。
 *
 * @param {Object} data 月薪或時薪的計算結果
 * @param {Object} config 員工薪資設定
 * @param {Object} [adjustments] 這次要用的計薪調整；沒給就沿用這張薪資單上次存的
 */
function applyPayrollRules_(data, config, adjustments) {
  const rules = getPayrollRules_();
  const isFullTime = data.salaryType === '月薪';
  const isPartTime = data.salaryType === '時薪';
  if (!isFullTime && !isPartTime) return data;

  const employeeId = data.employeeId;
  const yearMonth = data.yearMonth;
  const adj = adjustments || readSavedPayrollAdjustments_(employeeId, yearMonth) ||
              { salesBonus: 0, advanceDeduction: null, manualItems: [], note: '' };

  const deductionsBefore = (Number(data.grossSalary) || 0) - (Number(data.netSalary) || 0);
  // 全勤獎金已取消：前面的計算若帶了舊設定的全勤獎金，這裡扣掉
  let gross = (Number(data.grossSalary) || 0) - (Number(data.attendanceBonus) || 0);
  let deductions = deductionsBefore;

  let mealSubsidy = 0;
  let mealDays = 0;
  let birthdayGift = 0;
  let birthdayMessage = null;

  if (rules.enabled) {
    const attendance = getEmployeeMonthlyAttendanceInternal(employeeId, yearMonth);
    const shiftMap = (typeof getEmployeeShiftMapForMonth === 'function')
      ? getEmployeeShiftMapForMonth(employeeId, yearMonth)
      : {};

    if (isFullTime) {
      // 餐費：實際工時滿門檻的天數
      mealDays = attendance.filter(r => (Number(r.workHours) || 0) >= rules.mealMinHours).length;

      const gift = evaluateBirthdayGift_(rules.fullTimeBirthdayGift, rules, employeeId, config, yearMonth);
      birthdayGift = gift.amount;
      birthdayMessage = gift.message;
    } else {
      // 餐費：排班滿門檻、而且當天有出勤
      const worked = {};
      attendance.forEach(r => { if ((Number(r.workHours) || 0) > 0) worked[r.date] = true; });
      mealDays = Object.keys(shiftMap).filter(date =>
        worked[date] && payrollScheduledMinutes_(shiftMap[date]) >= rules.mealMinHours * 60
      ).length;

      const gift = evaluateBirthdayGift_(rules.partTimeBirthdayGift, rules, employeeId, config, yearMonth);
      birthdayGift = gift.amount;
      birthdayMessage = gift.message;
    }

    mealSubsidy = mealDays * rules.mealPerDay;
  }

  const advanceAuto = sumApprovedAdvances_(employeeId, yearMonth);
  const advanceDeduction = (adj.advanceDeduction !== null && adj.advanceDeduction !== undefined)
    ? adj.advanceDeduction
    : advanceAuto;

  const manualItems = adj.manualItems || [];
  const manualAddTotal = manualItems.filter(i => i.type === 'add').reduce((s, i) => s + i.amount, 0);
  const manualSubTotal = manualItems.filter(i => i.type === 'sub').reduce((s, i) => s + i.amount, 0);
  const salesBonus = adj.salesBonus || 0;

  gross += mealSubsidy + birthdayGift + salesBonus + manualAddTotal;
  deductions += advanceDeduction + manualSubTotal;

  data.attendanceBonus = 0;           // 全勤獎金已取消
  data.attendanceQualified = false;
  data.attendanceInfo = null;
  data.attendanceNote = '';
  data.mealSubsidy = mealSubsidy;
  data.mealDays = mealDays;
  data.birthdayGift = birthdayGift;
  data.birthdayMessage = birthdayMessage;
  data.birthdayNote = payrollMessageZh_(birthdayMessage);
  data.salesBonus = salesBonus;
  data.advanceDeduction = advanceDeduction;
  data.advanceAuto = advanceAuto;
  data.manualItems = manualItems;
  data.manualAddTotal = manualAddTotal;
  data.manualSubTotal = manualSubTotal;
  data.payslipNote = adj.note || '';
  data.payrollAdjustments = {
    salesBonus: salesBonus,
    advanceDeduction: adj.advanceDeduction === undefined ? null : adj.advanceDeduction,
    manualItems: manualItems,
    note: adj.note || '',
    // 給前端翻譯用：從試算表讀回的薪資單也能依語系顯示生日說明（存檔時才會用到，讀回時會被忽略）
    messages: { attendance: null, birthday: birthdayMessage }
  };
  data.payrollRulesApplied = !!rules.enabled;
  data.grossSalary = Math.round(gross);
  data.netSalary = Math.round(gross - deductions);

  return data;
}

// ==================== API ====================

/**
 * API：管理員儲存某張薪資單的計薪調整，並用它重算、存檔。
 * 參數：employeeId、yearMonth、adjustments（JSON：salesBonus、advanceDeduction、manualItems、note）
 */
function handleSavePayrollAdjustments(params) {
  try {
    const session = checkSession_(params.token);
    if (!session.ok || !session.user || session.user.dept !== '管理員') {
      return { ok: false, code: 'PERMISSION_DENIED', msg: '需要管理員權限' };
    }

    const employeeId = String(params.employeeId || '').trim();
    const yearMonth = String(params.yearMonth || '').trim();
    if (!employeeId || !/^\d{4}-\d{2}$/.test(yearMonth)) {
      return { ok: false, code: 'PAYROLL_ADJUST_INVALID', msg: '缺少員工或年月' };
    }

    let input;
    try {
      input = JSON.parse(params.adjustments || '{}');
    } catch (error) {
      return { ok: false, code: 'PAYROLL_ADJUST_INVALID', msg: '計薪調整格式錯誤' };
    }
    const checked = normalizePayrollAdjustments_(input);
    if (!checked.ok) return { ok: false, code: 'PAYROLL_ADJUST_INVALID', msg: checked.msg };

    const calculated = calculateMonthlySalary(employeeId, yearMonth, checked.adjustments);
    if (!calculated.success) return { ok: false, msg: calculated.message };

    calculated.data.token = params.token;   // 稽核記錄要知道是誰改的
    const saved = saveMonthlySalary(calculated.data);
    delete calculated.data.token;
    if (!saved.success) return { ok: false, msg: saved.message };

    return { ok: true, data: calculated.data, salaryId: saved.salaryId };

  } catch (error) {
    Logger.log(' handleSavePayrollAdjustments 錯誤: ' + error);
    return { ok: false, msg: error.toString() };
  }
}
