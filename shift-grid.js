// shift-grid.js — 週排班表（排班頁的預設畫面）
//
// 一人一列、星期一到日一欄，一眼看完整週。點格子跳出選單：選班別就排好（已有一個班就換掉），
// 「休息」清空那天，「自訂時間」帶到新增排班表單。每次修改後重新讀這一週，畫面跟試算表一致。
// 沒有排班權限的人只能看。
//
// 依賴 shift.js 的全域：shiftTemplates、allEmployees、employeesReady、isAdmin、isScheduler、
// LEAVE_SHIFT_TYPES、calcShiftMinutes、formatWorkHours、formatTimeOnly、showMessage、switchTab。

const GRID_DAY_KEYS = ['WEEK_MONDAY', 'WEEK_TUESDAY', 'WEEK_WEDNESDAY', 'WEEK_THURSDAY', 'WEEK_FRIDAY', 'WEEK_SATURDAY', 'WEEK_SUNDAY'];
// 班別顏色：依班別設定的順序輪流用，同一個班別每週都是同一個顏色
const GRID_COLORS = [
    { bg: '#e0ecff', fg: '#1e3a8a', dot: '#3b82f6' },
    { bg: '#dcfce7', fg: '#14532d', dot: '#22c55e' },
    { bg: '#fef3c7', fg: '#78350f', dot: '#f59e0b' },
    { bg: '#fce7f3', fg: '#831843', dot: '#ec4899' },
    { bg: '#ede9fe', fg: '#4c1d95', dot: '#8b5cf6' },
    { bg: '#cffafe', fg: '#164e63', dot: '#06b6d4' },
    { bg: '#ffedd5', fg: '#7c2d12', dot: '#f97316' },
    { bg: '#e5e7eb', fg: '#1f2937', dot: '#6b7280' }
];
const GRID_LEAVE_COLOR = { bg: '#f3f4f6', fg: '#6b7280', dot: '#9ca3af' };

let gridWeekStart = null;        // 這一週的星期一（Date，當地時間）
let gridShifts = [];
let gridHolidays = [];
let gridLoading = 0;

function gridCanEdit() {
    return typeof isAdmin !== 'undefined' && (isAdmin || isScheduler);
}

function gridMonday(date) {
    const d = new Date(date.getFullYear(), date.getMonth(), date.getDate());
    const day = d.getDay();
    d.setDate(d.getDate() - (day === 0 ? 6 : day - 1));
    return d;
}

function gridDates() {
    return Array.from({ length: 7 }, (_, i) => {
        const d = new Date(gridWeekStart);
        d.setDate(d.getDate() + i);
        return d;
    });
}

function gridColorFor(shiftType) {
    if (LEAVE_SHIFT_TYPES.includes(shiftType) || shiftType === '休假') return GRID_LEAVE_COLOR;
    const index = shiftTemplates.findIndex(tpl => tpl.name === shiftType);
    return index === -1 ? GRID_COLORS[GRID_COLORS.length - 1] : GRID_COLORS[index % (GRID_COLORS.length - 1)];
}

function gridWorkMinutes(shift) {
    const start = formatTimeOnly(shift.startTime);
    const end = formatTimeOnly(shift.endTime);
    if (!start || (start === '00:00' && end === '00:00')) return 0;
    return calcShiftMinutes(start, end, false, Number(shift.breakMinutes) || 0).workMinutes;
}

function gridEmployees() {
    // 停用、離職的不列；但這週有班的人一定列出來，排錯的班才看得到、刪得掉
    const list = (allEmployees || []).filter(e => e.userId && e.name &&
        !['停用', '離職'].includes(String(e.status || '').trim()));
    gridShifts.forEach(s => {
        if (!list.some(e => String(e.userId) === String(s.employeeId))) {
            list.push({ userId: s.employeeId, name: s.employeeName || s.employeeId });
        }
    });
    return list;
}

async function loadGridHolidays() {
    try {
        const cached = JSON.parse(localStorage.getItem('holidays_cache') || 'null');
        if (cached && Array.isArray(cached.holidays) && Date.now() - cached.savedAt < 24 * 60 * 60 * 1000) {
            gridHolidays = cached.holidays;
            return;
        }
    } catch (error) { /* 私密模式等 */ }
    try {
        const res = await apiRequestJson('getHolidays');
        if (res.ok && Array.isArray(res.holidays)) {
            gridHolidays = res.holidays;
            try { localStorage.setItem('holidays_cache', JSON.stringify({ savedAt: Date.now(), holidays: res.holidays })); } catch (error) { /* ignore */ }
        }
    } catch (error) {
        console.warn('取得國定假日失敗:', error);
    }
}

// ==================== 載入與畫面 ====================

async function initShiftGrid() {
    if (!document.getElementById('grid-table')) return;
    if (!gridWeekStart) gridWeekStart = gridMonday(new Date());
    document.getElementById('grid-prev').onclick = () => moveGridWeek(-7);
    document.getElementById('grid-next').onclick = () => moveGridWeek(7);
    document.getElementById('grid-today').onclick = () => { gridWeekStart = gridMonday(new Date()); loadShiftGrid(); };
    const apply = document.getElementById('grid-apply-pattern');
    apply.onclick = applyPatternToGridWeek;
    apply.style.display = gridCanEdit() ? '' : 'none';
    document.getElementById('grid-hint').textContent = t(gridCanEdit() ? 'GRID_HINT_EDIT' : 'GRID_HINT_VIEW');
    loadGridHolidays().then(() => renderShiftGrid());
    await loadShiftGrid();
}

function moveGridWeek(days) {
    closeGridMenu();
    gridWeekStart.setDate(gridWeekStart.getDate() + days);
    loadShiftGrid();
}

async function loadShiftGrid() {
    const table = document.getElementById('grid-table');
    if (!table || !gridWeekStart) return;
    const dates = gridDates();
    document.getElementById('grid-range').textContent = t('GRID_RANGE', {
        start: toLocalDateStr(dates[0]).replace(/-/g, '/'),
        end: toLocalDateStr(dates[6]).slice(5).replace('-', '/')
    });
    const ticket = ++gridLoading;
    table.classList.add('grid-busy');
    try {
        if (typeof employeesReady !== 'undefined') await employeesReady;
        const query = new URLSearchParams({ startDate: toLocalDateStr(dates[0]), endDate: toLocalDateStr(dates[6]) }).toString();
        const data = await apiRequestJson(`getShifts&${query}`);
        if (ticket !== gridLoading) return;   // 使用者已經換到別週
        if (!data.ok) {
            showMessage(data.msg || t('SHIFT_LOAD_FAILED'), 'error');
            return;
        }
        gridShifts = data.data || [];
        renderShiftGrid();
    } catch (error) {
        console.error('載入週排班失敗:', error);
        if (ticket === gridLoading) showMessage(t('SHIFT_LOAD_ERROR'), 'error');
    } finally {
        if (ticket === gridLoading) table.classList.remove('grid-busy');
    }
}

function renderShiftGrid() {
    const table = document.getElementById('grid-table');
    if (!table || !gridWeekStart) return;
    const dates = gridDates();
    const dateStrs = dates.map(d => toLocalDateStr(d));
    const today = toLocalDateStr(new Date());
    const people = gridEmployees();
    const editable = gridCanEdit();

    const head = dates.map((d, i) => {
        const ds = dateStrs[i];
        const off = i === 6 || gridHolidays.includes(ds);
        const cls = [off ? 'grid-off' : '', ds === today ? 'grid-today' : ''].join(' ');
        return `<th class="${cls}"><div class="grid-dow">${escapeHtml(t(GRID_DAY_KEYS[i]))}</div><div class="grid-date">${d.getMonth() + 1}/${d.getDate()}</div>${gridHolidays.includes(ds) ? `<div class="grid-holiday">${escapeHtml(t('GRID_HOLIDAY'))}</div>` : ''}</th>`;
    }).join('');

    const dayCounts = dateStrs.map(() => 0);
    const rows = people.map(emp => {
        let weekMinutes = 0;
        const cells = dateStrs.map((ds, i) => {
            const mine = gridShifts.filter(s => String(s.employeeId) === String(emp.userId) && String(s.date).slice(0, 10) === ds);
            const minutes = mine.reduce((sum, s) => sum + gridWorkMinutes(s), 0);
            weekMinutes += minutes;
            if (minutes > 0) dayCounts[i]++;
            const chips = mine.map(s => {
                const c = gridColorFor(s.shiftType);
                const start = formatTimeOnly(s.startTime), end = formatTimeOnly(s.endTime);
                const time = start && !(start === '00:00' && end === '00:00') ? `<span class="grid-chip-time">${escapeHtml(start)}–${escapeHtml(end)}</span>` : '';
                return `<span class="grid-chip" style="background:${c.bg};color:${c.fg}"><span class="grid-chip-name">${escapeHtml(s.shiftType)}</span>${time}</span>`;
            }).join('');
            const cls = ['grid-cell', editable ? 'grid-editable' : '', ds === today ? 'grid-today' : '', mine.length ? '' : 'grid-empty'].join(' ');
            const label = t('GRID_CELL_LABEL', { name: emp.name, date: `${dates[i].getMonth() + 1}/${dates[i].getDate()}` });
            return `<td class="${cls}" data-emp="${escapeHtml(emp.userId)}" data-date="${ds}"${editable ? ` tabindex="0" role="button" aria-label="${escapeHtml(label)}"` : ''}>${chips || (editable ? '<span class="grid-plus">＋</span>' : '')}</td>`;
        }).join('');
        return `<tr><th class="grid-name" scope="row">${escapeHtml(emp.name)}</th>${cells}<td class="grid-total">${weekMinutes ? escapeHtml(t('RECORDS_HOURS_SHORT', { hours: formatWorkHours(weekMinutes) })) : '—'}</td></tr>`;
    }).join('');

    table.innerHTML = `
        <thead><tr><th class="grid-name">${escapeHtml(t('SHIFT_EMPLOYEE_LABEL'))}</th>${head}<th class="grid-total">${escapeHtml(t('GRID_WEEK_HOURS'))}</th></tr></thead>
        <tbody>${rows || `<tr><td colspan="9" class="grid-none">${escapeHtml(t('SHIFT_NO_EMPLOYEE_DATA'))}</td></tr>`}</tbody>
        <tfoot><tr><th class="grid-name">${escapeHtml(t('GRID_HEADCOUNT'))}</th>${dayCounts.map(n => `<td class="grid-count">${n}</td>`).join('')}<td></td></tr></tfoot>`;

    if (editable) {
        table.querySelectorAll('td.grid-editable').forEach(td => {
            td.addEventListener('click', e => { e.stopPropagation(); openGridMenu(td); });
            td.addEventListener('keydown', e => {
                if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); openGridMenu(td); }
            });
        });
    }
    renderGridLegend();
}

function renderGridLegend() {
    const legend = document.getElementById('grid-legend');
    if (!legend) return;
    legend.innerHTML = shiftTemplates.filter(tpl => tpl.active && tpl.startTime).map(tpl => {
        const c = gridColorFor(tpl.name);
        return `<span class="grid-legend-item"><span class="grid-dot" style="background:${c.dot}"></span>${escapeHtml(tpl.name)} ${escapeHtml(tpl.startTime)}–${escapeHtml(tpl.endTime)}</span>`;
    }).join('');
}

// ==================== 格子選單 ====================

function closeGridMenu() {
    const menu = document.getElementById('grid-menu');
    if (menu) menu.remove();
    document.querySelectorAll('td.grid-active').forEach(td => td.classList.remove('grid-active'));
}

function openGridMenu(td) {
    closeGridMenu();
    const employeeId = td.dataset.emp;
    const date = td.dataset.date;
    const emp = gridEmployees().find(e => String(e.userId) === String(employeeId)) || { userId: employeeId, name: employeeId };
    const mine = gridShifts.filter(s => String(s.employeeId) === String(employeeId) && String(s.date).slice(0, 10) === date);
    const d = new Date(date + 'T12:00:00');
    const dayIndex = (d.getDay() + 6) % 7;
    const templates = shiftTemplates.filter(tpl => tpl.active && tpl.startTime);

    const menu = document.createElement('div');
    menu.id = 'grid-menu';
    menu.className = 'grid-menu';
    menu.setAttribute('role', 'dialog');
    menu.innerHTML = `
        <div class="grid-menu-title">${escapeHtml(emp.name)}・${d.getMonth() + 1}/${d.getDate()}（${escapeHtml(t(GRID_DAY_KEYS[dayIndex]))}）</div>
        ${mine.length ? `<div class="grid-menu-section">${mine.map(s => `
            <div class="grid-menu-current">
                <span><span class="grid-dot" style="background:${gridColorFor(s.shiftType).dot}"></span>${escapeHtml(s.shiftType)} ${escapeHtml(formatTimeOnly(s.startTime))}–${escapeHtml(formatTimeOnly(s.endTime))}</span>
                <button type="button" class="grid-menu-del" data-delete="${escapeHtml(s.shiftId)}">${escapeHtml(t('BTN_DELETE'))}</button>
            </div>`).join('')}</div>` : ''}
        <div class="grid-menu-label">${escapeHtml(t(mine.length === 1 ? 'GRID_CHANGE_TO' : 'GRID_PICK_SHIFT'))}</div>
        <div class="grid-menu-options">${templates.map(tpl => `
            <button type="button" class="grid-menu-opt${mine.length === 1 && mine[0].shiftType === tpl.name ? ' grid-current' : ''}" data-code="${escapeHtml(tpl.code)}">
                <span class="grid-dot" style="background:${gridColorFor(tpl.name).dot}"></span>
                <span class="grid-opt-name">${escapeHtml(tpl.name)}${mine.length === 1 && mine[0].shiftType === tpl.name ? `<span class="grid-current-tag">${escapeHtml(t('GRID_CURRENT'))}</span>` : ''}</span>
                <span class="grid-opt-time">${escapeHtml(tpl.startTime)}–${escapeHtml(tpl.endTime)}</span>
            </button>`).join('') || `<div class="grid-menu-empty">${escapeHtml(t('GRID_NO_TEMPLATES'))}</div>`}</div>
        <div class="grid-menu-footer">
            ${mine.length ? `<button type="button" class="grid-menu-off" data-off>${escapeHtml(t('GRID_SET_OFF'))}</button>` : ''}
            <button type="button" class="grid-menu-custom" data-custom>${escapeHtml(t('GRID_CUSTOM'))}</button>
            <button type="button" class="grid-menu-close" data-close>${escapeHtml(t('RECORDS_CANCEL'))}</button>
        </div>`;
    document.body.appendChild(menu);
    td.classList.add('grid-active');
    positionGridMenu(menu, td);

    menu.addEventListener('click', e => e.stopPropagation());
    menu.querySelectorAll('[data-code]').forEach(btn => btn.addEventListener('click', () => {
        const tpl = templates.find(x => x.code === btn.dataset.code);
        if (tpl) gridAssign(emp, date, tpl, mine);
    }));
    menu.querySelectorAll('[data-delete]').forEach(btn => btn.addEventListener('click', () => gridDelete([btn.dataset.delete])));
    menu.querySelector('[data-off]')?.addEventListener('click', () => gridDelete(mine.map(s => s.shiftId)));
    menu.querySelector('[data-custom]').addEventListener('click', () => {
        closeGridMenu();
        switchTab('add');
        const employeeSelect = document.getElementById('employee-select');
        if (employeeSelect) employeeSelect.value = employeeId;
        const dateInput = document.getElementById('shift-date');
        if (dateInput) dateInput.value = date;
        if (typeof refreshSameDayShifts === 'function') refreshSameDayShifts();
    });
    menu.querySelector('[data-close]').addEventListener('click', closeGridMenu);
    (menu.querySelector('.grid-menu-opt') || menu.querySelector('button'))?.focus({ preventScroll: true });
}

function positionGridMenu(menu, td) {
    const rect = td.getBoundingClientRect();
    const width = Math.min(300, window.innerWidth - 32);
    menu.style.width = width + 'px';
    let left = rect.left + window.scrollX;
    left = Math.max(16 + window.scrollX, Math.min(left, window.scrollX + window.innerWidth - width - 16));
    menu.style.left = left + 'px';
    // 下面放不下就往上開；上面也不夠就貼著畫面頂端（選單本身可以捲動）
    const height = menu.offsetHeight;
    let top = rect.bottom + 6;
    if (top + height > window.innerHeight - 16) {
        top = rect.top - height - 6;
        if (top < 16) top = Math.max(16, window.innerHeight - height - 16);
    }
    menu.style.top = (top + window.scrollY) + 'px';
}

document.addEventListener('click', closeGridMenu);
document.addEventListener('keydown', e => { if (e.key === 'Escape') closeGridMenu(); });

function setGridMenuBusy() {
    const menu = document.getElementById('grid-menu');
    if (menu) menu.querySelectorAll('button').forEach(b => { b.disabled = true; });
}

/** 選了班別：這天沒有班就新增；剛好一個班就換掉；兩個以上就再加一個 */
async function gridAssign(emp, date, tpl, mine) {
    if (mine.length === 1 && mine[0].shiftType === tpl.name &&
        formatTimeOnly(mine[0].startTime) === tpl.startTime && formatTimeOnly(mine[0].endTime) === tpl.endTime) {
        closeGridMenu();
        return;
    }
    setGridMenuBusy();
    const fields = {
        employeeId: emp.userId,
        employeeName: emp.name,
        date: date,
        shiftType: tpl.name,
        startTime: tpl.startTime,
        endTime: tpl.endTime,
        breakMinutes: tpl.breakMinutes
    };
    const update = mine.length === 1;
    if (update) {
        fields.shiftId = mine[0].shiftId;
        fields.location = mine[0].location || '';
        fields.note = mine[0].note || '';
    }
    try {
        const data = await apiRequestJson(`${update ? 'updateShift' : 'addShift'}&${new URLSearchParams(fields).toString()}`);
        closeGridMenu();
        if (data.ok) {
            showMessage(t('GRID_SAVED', { name: emp.name, shift: tpl.name }), 'success');
        } else {
            showMessage(data.msg || t(update ? 'SHIFT_UPDATE_FAILED' : 'SHIFT_ADD_FAILED'), 'error');
        }
    } catch (error) {
        console.error('週排班表儲存失敗:', error);
        closeGridMenu();
        showMessage(t('SHIFT_ADD_ERROR'), 'error');
    }
    loadShiftGrid();
}

async function gridDelete(shiftIds) {
    if (!shiftIds.length) return;
    setGridMenuBusy();
    let failed = 0;
    for (const id of shiftIds) {
        try {
            const data = await apiRequestJson(`deleteShift&shiftId=${encodeURIComponent(id)}`);
            if (!data.ok) failed++;
        } catch (error) {
            failed++;
        }
    }
    closeGridMenu();
    showMessage(failed ? t('SHIFT_DELETE_FAILED') : t('GRID_CLEARED'), failed ? 'error' : 'success');
    loadShiftGrid();
}

async function applyPatternToGridWeek() {
    const dates = gridDates();
    const start = toLocalDateStr(dates[0]);
    const end = toLocalDateStr(dates[6]);
    if (!confirm(t('SHIFT_WEEKLY_CONFIRM', { start, end }))) return;
    const button = document.getElementById('grid-apply-pattern');
    button.disabled = true;
    try {
        const data = await apiRequestJson(`generateShiftsFromPattern&${new URLSearchParams({ startDate: start, endDate: end, replace: 'false' }).toString()}`);
        if (data.ok) {
            showMessage(t('SHIFT_WEEKLY_RESULT', { added: data.added, skipped: data.skipped }), 'success');
        } else {
            showMessage(data.msg || t('SHIFT_WEEKLY_GENERATE_FAILED'), 'error');
        }
    } catch (error) {
        console.error('套用固定班表失敗:', error);
        showMessage(t('SHIFT_WEEKLY_GENERATE_FAILED'), 'error');
    } finally {
        button.disabled = false;
    }
    loadShiftGrid();
}
