// records.js — 資料管理頁（管理員）
//
// 打卡紀錄：查、新增、修改、刪除；假期餘額：查、調整。後端見 GS/RecordsAdmin.gs。
// 打卡紀錄用「列號 + 指紋（row、key）」指定要改哪一筆，資料在別處被動過時後端會拒絕，
// 這時重新查詢就好。

window.I18N_TITLE_KEY = 'RECORDS_PAGE_TITLE';

const LEAVE_TYPES = [
    'ANNUAL_LEAVE', 'COMP_TIME_OFF', 'PERSONAL_LEAVE', 'SICK_LEAVE', 'HOSPITALIZATION_LEAVE',
    'MENSTRUAL_LEAVE', 'FAMILY_CARE_LEAVE', 'BEREAVEMENT_LEAVE', 'MARRIAGE_LEAVE', 'MATERNITY_LEAVE',
    'PATERNITY_LEAVE', 'OFFICIAL_LEAVE', 'WORK_INJURY_LEAVE', 'NATURAL_DISASTER_LEAVE', 'ABSENCE_WITHOUT_LEAVE'
];
const PUNCH_STATUS = {
    NORMAL: { key: 'RECORDS_STATUS_NORMAL', cls: 'rec-badge-muted' },
    ADJUST_APPROVED: { key: 'RECORDS_STATUS_ADJUST_APPROVED', cls: 'rec-badge-muted' },
    ADJUST_PENDING: { key: 'RECORDS_STATUS_ADJUST_PENDING', cls: 'rec-badge-warn' },
    ADJUST_REJECTED: { key: 'RECORDS_STATUS_ADJUST_REJECTED', cls: 'rec-badge-bad' },
    VIRTUAL: { key: 'RECORDS_STATUS_VIRTUAL', cls: 'rec-badge-bad' }
};

let employees = [];
let punchRecords = [];
let editingPunch = null;      // 正在修改的那一筆（新增模式是 null）
let leaveEmployees = [];
let editingLeave = null;

function recMessage(text, type) {
    document.querySelectorAll('.rec-message').forEach(el => el.remove());
    const div = document.createElement('div');
    div.className = `rec-message ${type === 'error' ? 'error' : 'success'}`;
    div.setAttribute('role', type === 'error' ? 'alert' : 'status');
    div.textContent = text;
    document.body.appendChild(div);
    setTimeout(() => div.remove(), type === 'error' ? 6000 : 3000);
}

async function withButton(button, fn) {
    if (button) button.disabled = true;
    try {
        return await fn();
    } finally {
        if (button) button.disabled = false;
    }
}

function fmtHours(hours) {
    const h = Math.round((Number(hours) || 0) * 100) / 100;
    return t('RECORDS_HOURS_DAYS', { hours: h, days: Math.round(h / 8 * 100) / 100 });
}

// ==================== 初始化 ====================

let recordsStarted = false;

document.addEventListener('DOMContentLoaded', async () => {
    if (recordsStarted) return;   // 只初始化一次，否則員工選單會重複
    recordsStarted = true;
    await loadTranslations(detectLang());

    let session;
    try {
        session = await apiRequestJson('checkSession');
    } catch (error) {
        session = { ok: false };
    }
    if (!session.ok || !session.user || session.user.dept !== '管理員') {
        document.getElementById('records-denied').classList.remove('rec-hidden');
        return;
    }
    document.getElementById('records-main').classList.remove('rec-hidden');

    document.querySelectorAll('.rec-tab').forEach(tab => {
        tab.addEventListener('click', () => switchRecordsTab(tab.dataset.tab));
    });

    // 預設查這個月 1 號到今天
    const today = todayStr();
    ['punch', 'lv', 'ot'].forEach(prefix => {
        document.getElementById(prefix + '-start').value = today.slice(0, 8) + '01';
        document.getElementById(prefix + '-end').value = today;
    });
    // 請假常是預先申請的，結束日期預設到月底
    const monthEnd = new Date(Number(today.slice(0, 4)), Number(today.slice(5, 7)), 0);
    document.getElementById('lv-end').value = toLocalDateStr(monthEnd);

    document.getElementById('punch-search-btn').addEventListener('click', loadPunches);
    document.getElementById('punch-add-btn').addEventListener('click', () => openPunchForm(null));
    document.getElementById('pf-save').addEventListener('click', savePunch);
    document.getElementById('pf-cancel').addEventListener('click', closePunchForm);
    document.getElementById('lf-save').addEventListener('click', saveLeave);
    document.getElementById('lf-cancel').addEventListener('click', closeLeaveForm);
    document.getElementById('lv-search').addEventListener('click', loadLeaves);
    document.getElementById('ot-search').addEventListener('click', loadOvertime);
    document.getElementById('otf-save').addEventListener('click', saveOvertime);
    document.getElementById('otf-cancel').addEventListener('click', closeOvertimeForm);
    ['otf-start', 'otf-end'].forEach(id => document.getElementById(id).addEventListener('change', fillOvertimeHours));
    document.getElementById('br-search').addEventListener('click', browseSheet);
    document.getElementById('br-sheet').addEventListener('change', browseSheet);
    document.getElementById('br-keyword').addEventListener('keydown', e => { if (e.key === 'Enter') browseSheet(); });
    document.getElementById('br-download').addEventListener('click', downloadBrowseCsv);

    await Promise.all([loadEmployeeOptions(), loadLocationOptions()]);
    loadPunches();
});

function switchRecordsTab(name) {
    document.querySelectorAll('.rec-tab').forEach(tab => tab.classList.toggle('active', tab.dataset.tab === name));
    document.querySelectorAll('.rec-panel').forEach(panel => panel.classList.toggle('active', panel.id === 'panel-' + name));
    if (name === 'leave') loadLeaveBalances();
    if (name === 'leaves') loadLeaves();
    if (name === 'overtime') loadOvertime();
    if (name === 'browse') loadSheetList();
}

async function loadEmployeeOptions() {
    try {
        const data = await apiRequestJson('getAllUsers');
        employees = (data.ok ? data.users || [] : []).filter(u => u.userId && u.name);
    } catch (error) {
        employees = [];
    }
    const filter = document.getElementById('punch-employee');
    const form = document.getElementById('pf-employee');
    const extra = [...document.querySelectorAll('select.employee-filter')];
    employees.forEach(u => {
        filter.add(new Option(u.name, u.userId));
        form.add(new Option(u.name, u.userId));
        extra.forEach(select => select.add(new Option(u.name, u.userId)));
    });
}

async function loadLocationOptions() {
    try {
        const data = await apiRequestJson('getLocations');
        const select = document.getElementById('pf-location');
        (data.ok ? data.locations || [] : []).forEach(loc => select.add(new Option(loc.name, loc.name)));
    } catch (error) {
        console.warn('載入地點失敗:', error);
    }
}

// ==================== 打卡紀錄 ====================

async function loadPunches() {
    const start = document.getElementById('punch-start').value;
    const end = document.getElementById('punch-end').value;
    const employeeId = document.getElementById('punch-employee').value;
    const body = document.getElementById('punch-body');
    const summary = document.getElementById('punch-summary');
    if (!start || !end) {
        recMessage(t('SHIFT_WEEKLY_PICK_DATES'), 'error');
        return;
    }
    body.innerHTML = `<tr><td colspan="8" class="rec-empty">${escapeHtml(t('LOADING'))}</td></tr>`;
    summary.textContent = '';

    await withButton(document.getElementById('punch-search-btn'), async () => {
        try {
            const query = new URLSearchParams({ startDate: start, endDate: end, employeeId }).toString();
            const data = await apiRequestJson(`adminListPunches&${query}`);
            if (!data.ok) {
                body.innerHTML = '';
                recMessage(data.msg || t('RECORDS_LOAD_FAILED'), 'error');
                return;
            }
            punchRecords = data.records || [];
            renderPunches();
        } catch (error) {
            console.error('載入打卡紀錄失敗:', error);
            body.innerHTML = '';
            recMessage(t('RECORDS_LOAD_FAILED'), 'error');
        }
    });
}

function renderPunches() {
    const body = document.getElementById('punch-body');
    const summary = document.getElementById('punch-summary');
    summary.textContent = t('RECORDS_PUNCH_COUNT', { count: punchRecords.length });
    if (!punchRecords.length) {
        body.innerHTML = `<tr><td colspan="8" class="rec-empty">${escapeHtml(t('RECORDS_NO_PUNCHES'))}</td></tr>`;
        return;
    }
    body.innerHTML = '';
    let lastDate = '';
    punchRecords.forEach((rec, index) => {
        const tr = document.createElement('tr');
        if (rec.date !== lastDate) tr.className = 'date-start';
        const status = PUNCH_STATUS[rec.status] || PUNCH_STATUS.NORMAL;
        const typeCls = rec.type === '上班' ? 'rec-badge-in' : 'rec-badge-out';
        const typeLabel = rec.type === '上班' ? t('PUNCH_IN') : rec.type === '下班' ? t('PUNCH_OUT') : rec.type;
        tr.innerHTML = `
            <td>${rec.date !== lastDate ? escapeHtml(rec.date) : ''}</td>
            <td class="num">${escapeHtml(rec.time)}</td>
            <td>${escapeHtml(rec.name)}</td>
            <td><span class="rec-badge ${typeCls}">${escapeHtml(typeLabel)}</span></td>
            <td>${escapeHtml(rec.location)}</td>
            <td><span class="rec-badge ${status.cls}">${escapeHtml(t(status.key))}</span></td>
            <td>${escapeHtml(rec.note)}</td>
            <td class="ops">
                <button type="button" class="rec-btn rec-btn-secondary rec-btn-sm" data-edit="${index}">${escapeHtml(t('BTN_EDIT'))}</button>
                <button type="button" class="rec-btn rec-btn-danger rec-btn-sm" data-delete="${index}">${escapeHtml(t('BTN_DELETE'))}</button>
            </td>`;
        lastDate = rec.date;
        body.appendChild(tr);
    });
    body.querySelectorAll('[data-edit]').forEach(btn => {
        btn.addEventListener('click', () => openPunchForm(punchRecords[Number(btn.dataset.edit)]));
    });
    body.querySelectorAll('[data-delete]').forEach(btn => {
        btn.addEventListener('click', () => deletePunch(punchRecords[Number(btn.dataset.delete)], btn));
    });
}

function openPunchForm(rec) {
    editingPunch = rec;
    const form = document.getElementById('punch-form');
    const employeeSelect = document.getElementById('pf-employee');
    document.getElementById('punch-form-title').textContent = t(rec ? 'RECORDS_PUNCH_EDIT' : 'RECORDS_PUNCH_ADD');
    if (rec) {
        employeeSelect.value = rec.userId;
        document.getElementById('pf-date').value = rec.date;
        document.getElementById('pf-time').value = rec.time;
        document.getElementById('pf-type').value = rec.type;
        const loc = document.getElementById('pf-location');
        if (rec.location && ![...loc.options].some(o => o.value === rec.location)) loc.add(new Option(rec.location, rec.location));
        loc.value = rec.location;
        document.getElementById('pf-note').value = rec.note;
    } else {
        const filterEmployee = document.getElementById('punch-employee').value;
        if (filterEmployee) employeeSelect.value = filterEmployee;
        document.getElementById('pf-date').value = todayStr();
        document.getElementById('pf-time').value = '';
        document.getElementById('pf-type').value = '上班';
        document.getElementById('pf-location').value = '';
        document.getElementById('pf-note').value = '';
    }
    // 修改時不能換人：打錯人就刪掉再新增
    employeeSelect.disabled = !!rec;
    form.classList.remove('rec-hidden');
    form.scrollIntoView?.({ behavior: 'smooth', block: 'start' });
}

function closePunchForm() {
    editingPunch = null;
    document.getElementById('punch-form').classList.add('rec-hidden');
}

async function savePunch() {
    const fields = {
        employeeId: document.getElementById('pf-employee').value,
        date: document.getElementById('pf-date').value,
        time: document.getElementById('pf-time').value,
        type: document.getElementById('pf-type').value,
        location: document.getElementById('pf-location').value,
        note: document.getElementById('pf-note').value.trim()
    };
    if (!fields.employeeId || !fields.date || !fields.time) {
        recMessage(t('RECORDS_FILL_REQUIRED'), 'error');
        return;
    }
    const editing = editingPunch;
    if (editing) Object.assign(fields, { row: editing.row, key: editing.key });

    await withButton(document.getElementById('pf-save'), async () => {
        try {
            const query = new URLSearchParams(fields).toString();
            const data = await apiRequestJson(`${editing ? 'adminUpdatePunch' : 'adminAddPunch'}&${query}`);
            if (data.ok) {
                recMessage(t(editing ? 'RECORDS_PUNCH_UPDATED' : 'RECORDS_PUNCH_ADDED'), 'success');
                closePunchForm();
                loadPunches();
            } else {
                recMessage(data.msg || t('RECORDS_SAVE_FAILED'), 'error');
            }
        } catch (error) {
            console.error('儲存打卡失敗:', error);
            recMessage(t('RECORDS_SAVE_FAILED'), 'error');
        }
    });
}

async function deletePunch(rec, button) {
    const typeLabel = rec.type === '上班' ? t('PUNCH_IN') : t('PUNCH_OUT');
    if (!confirm(t('RECORDS_PUNCH_DELETE_CONFIRM', { name: rec.name, date: rec.date, time: rec.time, type: typeLabel }))) return;
    await withButton(button, async () => {
        try {
            const query = new URLSearchParams({ row: rec.row, key: rec.key }).toString();
            const data = await apiRequestJson(`adminDeletePunch&${query}`);
            if (data.ok) {
                recMessage(t('RECORDS_PUNCH_DELETED'), 'success');
                if (editingPunch && editingPunch.row === rec.row) closePunchForm();
                loadPunches();
            } else {
                recMessage(data.msg || t('RECORDS_DELETE_FAILED'), 'error');
            }
        } catch (error) {
            console.error('刪除打卡失敗:', error);
            recMessage(t('RECORDS_DELETE_FAILED'), 'error');
        }
    });
}

// ==================== 假期餘額 ====================

async function loadLeaveBalances() {
    const body = document.getElementById('leave-body');
    body.innerHTML = `<tr><td colspan="8" class="rec-empty">${escapeHtml(t('LOADING'))}</td></tr>`;
    try {
        const data = await apiRequestJson('adminGetLeaveBalances');
        if (!data.ok) {
            body.innerHTML = '';
            recMessage(data.msg || t('RECORDS_LOAD_FAILED'), 'error');
            return;
        }
        leaveEmployees = data.employees || [];
        renderLeaveBalances();
    } catch (error) {
        console.error('載入假期餘額失敗:', error);
        body.innerHTML = '';
        recMessage(t('RECORDS_LOAD_FAILED'), 'error');
    }
}

function renderLeaveBalances() {
    const body = document.getElementById('leave-body');
    if (!leaveEmployees.length) {
        body.innerHTML = `<tr><td colspan="8" class="rec-empty">${escapeHtml(t('SHIFT_NO_EMPLOYEE_DATA'))}</td></tr>`;
        return;
    }
    body.innerHTML = '';
    leaveEmployees.forEach((emp, index) => {
        const tr = document.createElement('tr');
        const b = emp.balances || {};
        tr.innerHTML = `
            <td>${escapeHtml(emp.name)}${emp.exists ? '' : ` <span class="rec-badge rec-badge-warn">${escapeHtml(t('RECORDS_LEAVE_NOT_SET'))}</span>`}</td>
            <td>${escapeHtml(emp.hireDate || '')}</td>
            <td class="num">${escapeHtml(fmtHours(b.ANNUAL_LEAVE))}</td>
            <td class="num">${escapeHtml(fmtHours(b.COMP_TIME_OFF))}</td>
            <td class="num">${escapeHtml(fmtHours(b.PERSONAL_LEAVE))}</td>
            <td class="num">${escapeHtml(fmtHours(b.SICK_LEAVE))}</td>
            <td>${escapeHtml(emp.updatedAt || '')}</td>
            <td class="ops"><button type="button" class="rec-btn rec-btn-secondary rec-btn-sm" data-adjust="${index}">${escapeHtml(t('RECORDS_ADJUST'))}</button></td>`;
        body.appendChild(tr);
    });
    body.querySelectorAll('[data-adjust]').forEach(btn => {
        btn.addEventListener('click', () => openLeaveForm(leaveEmployees[Number(btn.dataset.adjust)]));
    });
}

function openLeaveForm(emp) {
    editingLeave = emp;
    document.getElementById('leave-form-title').textContent = t('RECORDS_LEAVE_EDIT_TITLE', { name: emp.name });
    document.getElementById('lf-hire').value = emp.hireDate && /^\d{4}-\d{2}-\d{2}$/.test(emp.hireDate) ? emp.hireDate : '';
    const box = document.getElementById('leave-inputs');
    box.innerHTML = '';
    LEAVE_TYPES.forEach(type => {
        const field = document.createElement('div');
        field.className = 'rec-field';
        const id = 'lf-' + type;
        const value = (emp.balances || {})[type] || 0;
        field.innerHTML = `
            <label for="${id}">${escapeHtml(t(type))}</label>
            <input type="number" id="${id}" data-type="${type}" min="0" max="9999" step="0.5" value="${value}">
            <span class="rec-days"></span>`;
        const input = field.querySelector('input');
        const days = field.querySelector('.rec-days');
        const refresh = () => { days.textContent = fmtHours(input.value); };
        input.addEventListener('input', refresh);
        refresh();
        box.appendChild(field);
    });
    const form = document.getElementById('leave-form');
    form.classList.remove('rec-hidden');
    form.scrollIntoView?.({ behavior: 'smooth', block: 'start' });
}

function closeLeaveForm() {
    editingLeave = null;
    document.getElementById('leave-form').classList.add('rec-hidden');
}

async function saveLeave() {
    if (!editingLeave) return;
    const balances = {};
    for (const input of document.querySelectorAll('#leave-inputs input[data-type]')) {
        const value = Number(input.value);
        if (input.value === '' || !isFinite(value) || value < 0 || value > 9999) {
            recMessage(t('RECORDS_HOURS_INVALID'), 'error');
            input.focus();
            return;
        }
        balances[input.dataset.type] = value;
    }
    const params = new URLSearchParams({ employeeId: editingLeave.userId, balances: JSON.stringify(balances) });
    const hire = document.getElementById('lf-hire').value;
    if (hire) params.set('hireDate', hire);

    await withButton(document.getElementById('lf-save'), async () => {
        try {
            const data = await apiRequestJson(`adminSetLeaveBalance&${params.toString()}`);
            if (data.ok) {
                recMessage(t('RECORDS_LEAVE_SAVED', { name: editingLeave.name }), 'success');
                closeLeaveForm();
                loadLeaveBalances();
            } else {
                recMessage(data.msg || t('RECORDS_SAVE_FAILED'), 'error');
            }
        } catch (error) {
            console.error('儲存假期餘額失敗:', error);
            recMessage(t('RECORDS_SAVE_FAILED'), 'error');
        }
    });
}

// ==================== 請假紀錄 ====================

const RECORD_STATUS = {
    PENDING: { key: 'RECORDS_ST_PENDING', cls: 'rec-badge-warn' },
    APPROVED: { key: 'RECORDS_ST_APPROVED', cls: 'rec-badge-in' },
    REJECTED: { key: 'RECORDS_ST_REJECTED', cls: 'rec-badge-bad' },
    CANCELLED: { key: 'RECORDS_ST_CANCELLED', cls: 'rec-badge-muted' }
};
function statusBadge(status) {
    const s = RECORD_STATUS[String(status || '').toUpperCase()];
    return s ? `<span class="rec-badge ${s.cls}">${escapeHtml(t(s.key))}</span>` : escapeHtml(status || '');
}
function canStillChange(status) {
    const s = String(status || '').toUpperCase();
    return s === 'PENDING' || s === 'APPROVED';
}

let leaveRecords = [];

async function loadLeaves() {
    const body = document.getElementById('lv-body');
    body.innerHTML = `<tr><td colspan="8" class="rec-empty">${escapeHtml(t('LOADING'))}</td></tr>`;
    const query = new URLSearchParams({
        startDate: document.getElementById('lv-start').value,
        endDate: document.getElementById('lv-end').value,
        employeeId: document.getElementById('lv-employee').value,
        status: document.getElementById('lv-status').value
    }).toString();
    try {
        const data = await apiRequestJson(`adminListLeaves&${query}`);
        if (!data.ok) { body.innerHTML = ''; recMessage(data.msg || t('RECORDS_LOAD_FAILED'), 'error'); return; }
        leaveRecords = data.records || [];
        document.getElementById('lv-summary').textContent = t('RECORDS_PUNCH_COUNT', { count: leaveRecords.length });
        if (!leaveRecords.length) {
            body.innerHTML = `<tr><td colspan="8" class="rec-empty">${escapeHtml(t('RECORDS_NO_DATA'))}</td></tr>`;
            return;
        }
        body.innerHTML = '';
        leaveRecords.forEach((rec, index) => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td>${escapeHtml(rec.name)}</td>
                <td>${escapeHtml(t(rec.leaveType))}</td>
                <td>${escapeHtml(rec.start)}</td>
                <td>${escapeHtml(rec.end)}</td>
                <td class="num">${escapeHtml(fmtHours(rec.hours))}</td>
                <td>${statusBadge(rec.status)}</td>
                <td>${escapeHtml(rec.reason)}${rec.comment ? `<div class="rec-days">${escapeHtml(rec.comment)}</div>` : ''}</td>
                <td class="ops">${canStillChange(rec.status) ? `<button type="button" class="rec-btn rec-btn-danger rec-btn-sm" data-cancel="${index}">${escapeHtml(t('RECORDS_CANCEL_RECORD'))}</button>` : ''}</td>`;
            body.appendChild(tr);
        });
        body.querySelectorAll('[data-cancel]').forEach(btn => {
            btn.addEventListener('click', () => cancelLeave(leaveRecords[Number(btn.dataset.cancel)], btn));
        });
    } catch (error) {
        console.error('載入請假紀錄失敗:', error);
        body.innerHTML = '';
        recMessage(t('RECORDS_LOAD_FAILED'), 'error');
    }
}

async function cancelLeave(rec, button) {
    const approved = String(rec.status).toUpperCase() === 'APPROVED';
    const reason = prompt(t(approved ? 'RECORDS_LEAVE_CANCEL_APPROVED' : 'RECORDS_LEAVE_CANCEL_PENDING',
        { name: rec.name, type: t(rec.leaveType), start: rec.start, hours: rec.hours }), '');
    if (reason === null) return;
    await withButton(button, async () => {
        try {
            const query = new URLSearchParams({ row: rec.row, key: rec.key, comment: reason }).toString();
            const data = await apiRequestJson(`adminCancelLeave&${query}`);
            if (data.ok) {
                recMessage(data.refunded ? t('RECORDS_LEAVE_CANCELLED_REFUND', { hours: data.refunded }) : t('RECORDS_LEAVE_CANCELLED'), 'success');
                loadLeaves();
            } else {
                recMessage(data.msg || t('RECORDS_SAVE_FAILED'), 'error');
            }
        } catch (error) {
            console.error('取消請假失敗:', error);
            recMessage(t('RECORDS_SAVE_FAILED'), 'error');
        }
    });
}

// ==================== 加班紀錄 ====================

let overtimeRecords = [];
let editingOvertime = null;

async function loadOvertime() {
    const body = document.getElementById('ot-body');
    body.innerHTML = `<tr><td colspan="7" class="rec-empty">${escapeHtml(t('LOADING'))}</td></tr>`;
    const query = new URLSearchParams({
        startDate: document.getElementById('ot-start').value,
        endDate: document.getElementById('ot-end').value,
        employeeId: document.getElementById('ot-employee').value,
        status: document.getElementById('ot-status').value
    }).toString();
    try {
        const data = await apiRequestJson(`adminListOvertime&${query}`);
        if (!data.ok) { body.innerHTML = ''; recMessage(data.msg || t('RECORDS_LOAD_FAILED'), 'error'); return; }
        overtimeRecords = data.records || [];
        document.getElementById('ot-summary').textContent = t('RECORDS_PUNCH_COUNT', { count: overtimeRecords.length });
        if (!overtimeRecords.length) {
            body.innerHTML = `<tr><td colspan="7" class="rec-empty">${escapeHtml(t('RECORDS_NO_DATA'))}</td></tr>`;
            return;
        }
        body.innerHTML = '';
        overtimeRecords.forEach((rec, index) => {
            const tr = document.createElement('tr');
            const active = canStillChange(rec.status);
            tr.innerHTML = `
                <td>${escapeHtml(rec.date)}</td>
                <td>${escapeHtml(rec.name)}</td>
                <td class="num">${escapeHtml(rec.startTime)}～${escapeHtml(rec.endTime)}</td>
                <td class="num">${escapeHtml(t('RECORDS_HOURS_SHORT', { hours: rec.hours }))}</td>
                <td>${statusBadge(rec.status)}</td>
                <td>${escapeHtml(rec.reason)}${rec.comment ? `<div class="rec-days">${escapeHtml(rec.comment)}</div>` : ''}</td>
                <td class="ops">${active ? `
                    <button type="button" class="rec-btn rec-btn-secondary rec-btn-sm" data-edit="${index}">${escapeHtml(t('BTN_EDIT'))}</button>
                    <button type="button" class="rec-btn rec-btn-danger rec-btn-sm" data-cancel="${index}">${escapeHtml(t('RECORDS_CANCEL_RECORD'))}</button>` : ''}</td>`;
            body.appendChild(tr);
        });
        body.querySelectorAll('[data-edit]').forEach(btn => btn.addEventListener('click', () => openOvertimeForm(overtimeRecords[Number(btn.dataset.edit)])));
        body.querySelectorAll('[data-cancel]').forEach(btn => btn.addEventListener('click', () => cancelOvertime(overtimeRecords[Number(btn.dataset.cancel)], btn)));
    } catch (error) {
        console.error('載入加班紀錄失敗:', error);
        body.innerHTML = '';
        recMessage(t('RECORDS_LOAD_FAILED'), 'error');
    }
}

function openOvertimeForm(rec) {
    editingOvertime = rec;
    document.getElementById('ot-form-title').textContent = t('RECORDS_OT_EDIT_TITLE', { name: rec.name, date: rec.date });
    document.getElementById('otf-start').value = rec.startTime;
    document.getElementById('otf-end').value = rec.endTime;
    document.getElementById('otf-hours').value = rec.hours;
    const form = document.getElementById('ot-form');
    form.classList.remove('rec-hidden');
    form.scrollIntoView?.({ behavior: 'smooth', block: 'start' });
}

function closeOvertimeForm() {
    editingOvertime = null;
    document.getElementById('ot-form').classList.add('rec-hidden');
}

/** 改了開始或結束時間，時數跟著算（四捨五入到 0.5 小時，還是可以手動改） */
function fillOvertimeHours() {
    const [sh, sm] = document.getElementById('otf-start').value.split(':').map(Number);
    const [eh, em] = document.getElementById('otf-end').value.split(':').map(Number);
    if ([sh, sm, eh, em].some(n => isNaN(n))) return;
    let minutes = (eh * 60 + em) - (sh * 60 + sm);
    if (minutes <= 0) minutes += 24 * 60;
    document.getElementById('otf-hours').value = Math.round(minutes / 30) / 2;
}

async function saveOvertime() {
    if (!editingOvertime) return;
    const rec = editingOvertime;
    const fields = {
        row: rec.row, key: rec.key,
        startTime: document.getElementById('otf-start').value,
        endTime: document.getElementById('otf-end').value,
        hours: document.getElementById('otf-hours').value
    };
    const hours = Number(fields.hours);
    if (!fields.startTime || !fields.endTime || !(hours > 0 && hours <= 24)) {
        recMessage(t('RECORDS_OT_INVALID'), 'error');
        return;
    }
    await withButton(document.getElementById('otf-save'), async () => {
        try {
            const data = await apiRequestJson(`adminUpdateOvertime&${new URLSearchParams(fields).toString()}`);
            if (data.ok) {
                recMessage(t(data.salaryUpdated ? 'RECORDS_OT_UPDATED_SALARY' : 'RECORDS_OT_UPDATED'), 'success');
                closeOvertimeForm();
                loadOvertime();
            } else {
                recMessage(data.msg || t('RECORDS_SAVE_FAILED'), 'error');
            }
        } catch (error) {
            console.error('修改加班失敗:', error);
            recMessage(t('RECORDS_SAVE_FAILED'), 'error');
        }
    });
}

async function cancelOvertime(rec, button) {
    const reason = prompt(t('RECORDS_OT_CANCEL_CONFIRM', { name: rec.name, date: rec.date, hours: rec.hours }), '');
    if (reason === null) return;
    await withButton(button, async () => {
        try {
            const query = new URLSearchParams({ row: rec.row, key: rec.key, comment: reason }).toString();
            const data = await apiRequestJson(`adminCancelOvertime&${query}`);
            if (data.ok) {
                recMessage(t(data.salaryUpdated ? 'RECORDS_OT_CANCELLED_SALARY' : 'RECORDS_OT_CANCELLED'), 'success');
                if (editingOvertime && editingOvertime.row === rec.row) closeOvertimeForm();
                loadOvertime();
            } else {
                recMessage(data.msg || t('RECORDS_SAVE_FAILED'), 'error');
            }
        } catch (error) {
            console.error('取消加班失敗:', error);
            recMessage(t('RECORDS_SAVE_FAILED'), 'error');
        }
    });
}

// ==================== 其他資料表（唯讀） ====================

let browseResult = null;
let sheetListLoaded = false;

async function loadSheetList() {
    if (sheetListLoaded) return;
    const select = document.getElementById('br-sheet');
    try {
        const data = await apiRequestJson('adminListSheets');
        if (!data.ok) { recMessage(data.msg || t('RECORDS_LOAD_FAILED'), 'error'); return; }
        select.innerHTML = '';
        (data.sheets || []).forEach(sh => select.add(new Option(t('RECORDS_SHEET_OPTION', { name: sh.name, rows: sh.rows }), sh.key)));
        sheetListLoaded = true;
        if (select.options.length) browseSheet();
    } catch (error) {
        console.error('載入資料表清單失敗:', error);
        recMessage(t('RECORDS_LOAD_FAILED'), 'error');
    }
}

async function browseSheet() {
    const key = document.getElementById('br-sheet').value;
    if (!key) return;
    const head = document.getElementById('br-head');
    const body = document.getElementById('br-body');
    const summary = document.getElementById('br-summary');
    head.innerHTML = '';
    body.innerHTML = `<tr><td class="rec-empty">${escapeHtml(t('LOADING'))}</td></tr>`;
    summary.textContent = '';
    try {
        const query = new URLSearchParams({ sheet: key, keyword: document.getElementById('br-keyword').value.trim() }).toString();
        const data = await apiRequestJson(`adminBrowseSheet&${query}`);
        if (!data.ok) { body.innerHTML = ''; recMessage(data.msg || t('RECORDS_LOAD_FAILED'), 'error'); return; }
        browseResult = data;
        head.innerHTML = `<tr>${(data.headers || []).map(h => `<th>${escapeHtml(h)}</th>`).join('')}</tr>`;
        summary.textContent = data.total > data.rows.length
            ? t('RECORDS_BROWSE_LIMITED', { total: data.total, shown: data.rows.length })
            : t('RECORDS_PUNCH_COUNT', { count: data.total });
        if (!data.rows.length) {
            body.innerHTML = `<tr><td colspan="${Math.max(1, data.headers.length)}" class="rec-empty">${escapeHtml(t('RECORDS_NO_DATA'))}</td></tr>`;
            return;
        }
        body.innerHTML = data.rows.map(row => `<tr>${row.map(v => `<td>${escapeHtml(v)}</td>`).join('')}</tr>`).join('');
    } catch (error) {
        console.error('讀取資料表失敗:', error);
        body.innerHTML = '';
        recMessage(t('RECORDS_LOAD_FAILED'), 'error');
    }
}

function downloadBrowseCsv() {
    if (!browseResult || !browseResult.rows || !browseResult.rows.length) {
        recMessage(t('RECORDS_NO_DATA'), 'error');
        return;
    }
    // 開頭是 = + - @ 的儲存格加單引號，避免 Excel 當成公式執行
    const cell = v => {
        let text = String(v ?? '');
        if (/^[=+\-@]/.test(text)) text = "'" + text;
        return /[",\n\r]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
    };
    const lines = [browseResult.headers, ...browseResult.rows].map(row => row.map(cell).join(','));
    const blob = new Blob(['﻿' + lines.join('\r\n')], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `${browseResult.name}_${todayStr()}.csv`;
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(link.href), 1000);
}
