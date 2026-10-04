// 打卡地點：Nominatim 地址搜尋 + 可拖曳微調的選取器地圖
// 從 script.js 拆出，只依賴 utils.js / i18n.js / libs.js 與 showNotification()。

// ==================== 地點搜尋功能 ====================

/**
 * 使用 Nominatim API 搜尋地點
 */
async function searchLocation(query) {
    if (!query || query.trim() === '') {
        return [];
    }
    
    // 先限定台灣搜尋（門牌與路段的命中率較高），沒有結果再放寬到全球
    const request = async (countryCode) => {
        const url = `https://nominatim.openstreetmap.org/search?format=json&q=${encodeURIComponent(query)}`
            + `&limit=5&accept-language=zh-TW${countryCode ? '&countrycodes=' + countryCode : ''}`;
        const response = await fetch(url);
        if (!response.ok) throw new Error('搜尋失敗');
        return await response.json();
    };
    
    try {
        let results = await request('tw');
        if (!results.length) results = await request('');
        return results;
        
    } catch (error) {
        console.error('地點搜尋錯誤:', error);
        showNotification(t('NOTIF_SEARCH_FAILED'), 'error');
        return [];
    }
}

/**
 * 顯示搜尋結果
 */
function displaySearchResults(results) {
    const resultsList = document.getElementById('search-results-list');
    const resultsContainer = document.getElementById('search-results');
    
    if (!resultsList || !resultsContainer) return;
    
    resultsList.innerHTML = '';
    
    if (results.length === 0) {
        resultsContainer.classList.add('hidden');
        showNotification(t('NOTIF_NO_PLACE_FOUND'), 'warning');
        return;
    }
    
    resultsContainer.classList.remove('hidden');
    
    results.forEach(result => {
        const li = document.createElement('li');
        li.className = 'text-sm text-gray-800 dark:text-gray-200';
        li.innerHTML = `
            <div class="font-semibold">${escapeHtml(result.display_name)}</div>
            <div class="text-xs text-gray-500 dark:text-gray-400 mt-1">
                ${parseFloat(result.lat).toFixed(6)}, ${parseFloat(result.lon).toFixed(6)}
            </div>
        `;
        
        li.addEventListener('click', () => {
            selectSearchResult(result);
        });
        
        resultsList.appendChild(li);
    });
}

/**
 * 選擇搜尋結果
 */
function selectSearchResult(result) {
    const nameInput = document.getElementById('location-name');
    const latInput = document.getElementById('location-lat');
    const lngInput = document.getElementById('location-lng');
    const addBtn = document.getElementById('add-location-btn');
    const resultsContainer = document.getElementById('search-results');
    
    if (nameInput) nameInput.value = result.display_name.split(',')[0].trim();
    if (latInput) latInput.value = parseFloat(result.lat).toFixed(6);
    if (lngInput) lngInput.value = parseFloat(result.lon).toFixed(6);
    if (addBtn) addBtn.disabled = false;
    if (resultsContainer) resultsContainer.classList.add('hidden');
    
    // 在下方小地圖標出這個點，之後可以拖曳微調
    setPickerLocation(parseFloat(result.lat), parseFloat(result.lon));
    
    showNotification(t('NOTIF_LOCATION_PICKED'), 'success');
}

// ==================== 打卡地點選取器（可拖曳微調） ====================
// 搜尋回來的座標是建物或路段中心，跟實際打卡的門口常差數十公尺，
// 所以在「新增打卡地點」表單裡放一張小地圖，標記可以拖，圓圈即時跟著半徑走。

let pickerMap = null;
let pickerMarker = null;
let pickerCircle = null;

function pickerRadius() {
    const slider = document.getElementById('location-radius');
    return slider ? parseInt(slider.value) : 200;
}

// 把座標寫回表單欄位
function writePickedCoords(lat, lng) {
    const latInput = document.getElementById('location-lat');
    const lngInput = document.getElementById('location-lng');
    const addBtn = document.getElementById('add-location-btn');
    if (latInput) latInput.value = lat.toFixed(6);
    if (lngInput) lngInput.value = lng.toFixed(6);
    if (addBtn) addBtn.disabled = false;
}

/**
 * 在選取器地圖上標出座標；地圖第一次用到時才建立。
 */
async function setPickerLocation(lat, lng) {
    writePickedCoords(lat, lng);
    
    const el = document.getElementById('location-picker-map');
    if (!el) return;
    
    try {
        await ensureLib('leaflet');
    } catch (err) {
        console.error('地圖載入失敗:', err);
        return;
    }
    
    const coords = [lat, lng];
    const radius = pickerRadius();
    
    if (!pickerMap) {
        el.innerHTML = '';
        el.classList.remove('flex', 'items-center', 'justify-center');
        pickerMap = L.map(el).setView(coords, 18);
        L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
            attribution: '&copy; OpenStreetMap'
        }).addTo(pickerMap);
        
        pickerMarker = L.marker(coords, { draggable: true }).addTo(pickerMap);
        pickerCircle = L.circle(coords, {
            color: 'blue', fillColor: '#30f', fillOpacity: 0.2, radius: radius
        }).addTo(pickerMap);
        
        // 拖曳結束就把新座標寫回欄位，圓圈也跟著移動
        pickerMarker.on('drag', (e) => pickerCircle.setLatLng(e.target.getLatLng()));
        pickerMarker.on('dragend', (e) => {
            const p = e.target.getLatLng();
            writePickedCoords(p.lat, p.lng);
            showNotification(t('NOTIF_PICKER_ADJUSTED', { lat: p.lat.toFixed(6), lng: p.lng.toFixed(6) }), 'success');
        });
        // 點地圖也能直接改點位
        pickerMap.on('click', (e) => setPickerLocation(e.latlng.lat, e.latlng.lng));
        
        setTimeout(() => pickerMap.invalidateSize(), 100);
    } else {
        pickerMap.setView(coords, Math.max(pickerMap.getZoom(), 17));
        pickerMarker.setLatLng(coords);
        pickerCircle.setLatLng(coords).setRadius(radius);
    }
}

// 分頁切回管理員時，地圖是在隱藏狀態下建立的話尺寸會歪掉
function refreshLocationPicker() {
    if (pickerMap) setTimeout(() => pickerMap.invalidateSize(), 100);
}

// ==================== 範圍調整拉桿 ====================

/**
 * 初始化範圍拉桿
 */
function initRadiusSlider() {
    const slider = document.getElementById('location-radius');
    const valueDisplay = document.getElementById('radius-value');
    
    if (!slider || !valueDisplay) return;
    
    slider.addEventListener('input', (e) => {
        const value = e.target.value;
        valueDisplay.textContent = value;
        
        //  修正：先檢查 circle 是否存在
        if (circle && currentCoords) {
            circle.setRadius(parseInt(value));
        }
        
        // 新增打卡地點的選取器地圖也要跟著改
        if (pickerCircle) {
            pickerCircle.setRadius(parseInt(value));
        }
    });
}

// ==================== 打卡地點列表（管理員：修改、刪除） ====================
// 按「編輯」把地點帶到上方「新增打卡地點」表單，沿用同一張地圖拖曳調整；
// 這時「新增地點」按鈕變成「儲存修改」，旁邊多一個「取消修改」。

let editingLocationId = null;

async function loadLocationAdminList() {
    const list = document.getElementById('location-admin-list');
    if (!list) return;
    list.innerHTML = `<li class="text-gray-500 dark:text-gray-400 text-sm">${escapeHtml(t('LOADING'))}</li>`;
    try {
        const res = await callApifetch('getLocations');
        if (!res.ok) throw new Error(res.msg || 'getLocations failed');
        const locations = res.locations || [];
        if (!locations.length) {
            list.innerHTML = `<li class="text-gray-500 dark:text-gray-400 text-sm">${escapeHtml(t('LOCATION_LIST_EMPTY'))}</li>`;
            return;
        }
        list.innerHTML = '';
        locations.forEach(loc => {
            const li = document.createElement('li');
            li.className = 'p-4 rounded-lg border border-gray-200 dark:border-gray-700 bg-gray-50 dark:bg-gray-800 flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3';
            if (loc.id === editingLocationId) li.classList.add('ring-2', 'ring-indigo-400');
            li.innerHTML = `
                <div class="min-w-0">
                    <div class="font-semibold text-gray-800 dark:text-white break-words">${escapeHtml(loc.name)}</div>
                    <div class="text-xs text-gray-500 dark:text-gray-400 mt-1">
                        ${Number(loc.lat).toFixed(6)}, ${Number(loc.lng).toFixed(6)} ｜ ${escapeHtml(t('LOCATION_RADIUS_SHORT', { radius: loc.scope }))}
                    </div>
                </div>
                <div class="flex gap-2 shrink-0">
                    <button type="button" data-edit class="px-3 py-1.5 rounded-lg text-sm font-semibold bg-indigo-100 text-indigo-700 hover:bg-indigo-200 dark:bg-indigo-900/40 dark:text-indigo-300">${escapeHtml(t('BTN_EDIT'))}</button>
                    <button type="button" data-delete class="px-3 py-1.5 rounded-lg text-sm font-semibold bg-red-600 text-white hover:bg-red-700">${escapeHtml(t('BTN_DELETE'))}</button>
                </div>`;
            li.querySelector('[data-edit]').addEventListener('click', () => startEditLocation(loc));
            li.querySelector('[data-delete]').addEventListener('click', () => deleteLocationById(loc));
            list.appendChild(li);
        });
    } catch (err) {
        console.error('載入打卡地點失敗:', err);
        list.innerHTML = `<li class="text-red-600 text-sm">${escapeHtml(t('LOCATION_LIST_FAILED'))}</li>`;
    }
}

function setLocationRadius(radius) {
    const slider = document.getElementById('location-radius');
    if (!slider) return;
    slider.value = radius;
    slider.dispatchEvent(new Event('input'));
}

function startEditLocation(loc) {
    editingLocationId = loc.id;
    document.getElementById('location-name').value = loc.name;
    setLocationRadius(loc.scope || 200);
    setPickerLocation(Number(loc.lat), Number(loc.lng));
    const addBtn = document.getElementById('add-location-btn');
    if (addBtn) {
        addBtn.textContent = t('LOCATION_SAVE_EDIT');
        addBtn.disabled = false;
    }
    document.getElementById('cancel-location-edit-btn')?.classList.remove('hidden');
    loadLocationAdminList();
    // 列表在表單下方，捲回表單讓管理員看到帶入的內容
    document.getElementById('location-name').scrollIntoView?.({ behavior: 'smooth', block: 'center' });
}

/** 清空表單、回到新增模式 */
function resetLocationForm() {
    editingLocationId = null;
    ['location-name', 'location-lat', 'location-lng', 'location-search'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.value = '';
    });
    setLocationRadius(200);
    const addBtn = document.getElementById('add-location-btn');
    if (addBtn) {
        addBtn.textContent = t('ADD_LOCATION_BTN');
        addBtn.disabled = true;
    }
    const getBtn = document.getElementById('get-location-btn');
    if (getBtn) {
        getBtn.textContent = t('GET_LOCATION_BTN');
        getBtn.disabled = false;
    }
    document.getElementById('cancel-location-edit-btn')?.classList.add('hidden');
}

/** 「新增地點／儲存修改」按鈕 */
async function submitLocationForm() {
    const name = document.getElementById('location-name').value.trim();
    const lat = document.getElementById('location-lat').value;
    const lng = document.getElementById('location-lng').value;
    const radius = document.getElementById('location-radius').value;

    if (!name || !lat || !lng) {
        showNotification(t('NOTIF_FILL_ALL_AND_LOCATION'), 'error');
        return;
    }

    const editing = editingLocationId;
    const query = new URLSearchParams({ name, lat, lng, radius });
    if (editing) query.set('id', editing);
    try {
        const res = await callApifetch(`${editing ? 'updateLocation' : 'addLocation'}&${query.toString()}`);
        if (res.ok) {
            showNotification(editing
                ? t('LOCATION_UPDATED', { count: res.renamed || 0 })
                : t('NOTIF_LOCATION_ADDED'), 'success');
            resetLocationForm();
            loadLocationAdminList();
        } else {
            showNotification((editing ? t('LOCATION_UPDATE_FAILED') : t('NOTIF_ADD_LOCATION_FAILED_MSG')) + (res.msg || ''), 'error');
        }
    } catch (err) {
        console.error(err);
        showNotification(editing ? t('LOCATION_UPDATE_FAILED') : t('NOTIF_ADD_LOCATION_FAILED_MSG'), 'error');
    }
}

async function deleteLocationById(loc) {
    if (!confirm(t('LOCATION_DELETE_CONFIRM', { name: loc.name }))) return;
    try {
        const res = await callApifetch(`deleteLocation&id=${encodeURIComponent(loc.id)}`);
        if (res.ok) {
            showNotification(t('LOCATION_DELETED', { name: loc.name }), 'success');
            if (editingLocationId === loc.id) resetLocationForm();
            loadLocationAdminList();
        } else {
            showNotification(res.msg || t('LOCATION_DELETE_FAILED'), 'error');
        }
    } catch (err) {
        console.error(err);
        showNotification(t('LOCATION_DELETE_FAILED'), 'error');
    }
}
