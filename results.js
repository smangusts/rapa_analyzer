let filteredData = [];
let currentCompName = '';
let currentCompDates = '';

document.addEventListener('DOMContentLoaded', async () => {
    await loadData();
    setupEventListeners();
});

async function loadData() {
    // Получаем все необходимые данные из хранилища
    const data = await chrome.storage.local.get([
        'competitionData', 
        'competitionUrl', 
        'competitionId', 
        'competitionName', 
        'competitionDates'
    ]);
    allData = data.competitionData || [];
    currentCompName = data.competitionName || '';
    currentCompDates = data.competitionDates || '';
    
    // Проверка на наличие новых полей (Ранг, Возраст)
    if (allData.length > 0 && !('rank' in allData[0])) {
        showOutdatedDataWarning();
    }

    filteredData = [...allData];
    
    if (allData.length > 0) {
        // Update UI Info
        if (data.competitionUrl) {
            const compInfoEl = document.getElementById('compInfo');
            if (compInfoEl) compInfoEl.textContent = `ID: ${data.competitionId || 'N/A'}`;
            document.title = `Анализ: ${data.competitionId || 'RAPA'}`;
        }
        if (data.competitionName) {
            let fullTitle = data.competitionName;
            if (data.competitionDates) {
                fullTitle += ` (${data.competitionDates})`;
            }
            document.getElementById('compName').textContent = fullTitle;
            // Скрываем отдельный элемент дат, так как они теперь в заголовке
            const datesEl = document.getElementById('compDates');
            if (datesEl) datesEl.style.display = 'none';
        }
        
        populateFilters();
        updateDashboard();
        
        // Технический лог для отладки и будущего (Future-proofing)
        console.group('🛠 Технический отчет RAPA Analyzer');
        console.log('Загружено записей:', allData.length);
        console.log('Дисциплины:', [...new Set(allData.map(d => d.discipline))]);
        console.log('Категории:', [...new Set(allData.map(d => d.category))]);
        console.groupEnd();

        showNotification('Данные успешно загружены!');
    } else {
        showNotification('Данные не найдены. Пожалуйста, запустите сбор из поп-апа.', true);
    }
}

function setupEventListeners() {
    const filterIds = ['filterName', 'filterDiscip', 'filterCategory', 'filterRank', 'filterCity', 'filterSchool'];
    filterIds.forEach(id => {
        const el = document.getElementById(id);
        el.addEventListener('change', (e) => {
            applyFilters();
        });
        if (el.tagName === 'INPUT') {
            el.addEventListener('input', applyFilters);
        }
    });

    document.getElementById('resetFilters').addEventListener('click', () => {
        filterIds.forEach(id => {
            const el = document.getElementById(id);
            el.value = '';
        });
        applyFilters();
    });

    document.getElementById('exportExcel').addEventListener('click', exportToExcel);
    
    // Управление школами (School Merger)
    document.getElementById('manageSchoolsBtn').addEventListener('click', openMergerModal);
    document.getElementById('closeMergerModal').addEventListener('click', closeMergerModal);
    document.getElementById('cancelMergerBtn').addEventListener('click', closeMergerModal);
    document.getElementById('targetSchoolSelect').addEventListener('change', handleTargetSchoolChange);
    document.getElementById('confirmMergerBtn').addEventListener('click', executeMerge);
}

function populateFilters() {
    updateAllFiltersOptions();
}

function updateAllFiltersOptions() {
    const filters = [
        { id: 'filterDiscip', field: 'discipline' },
        { id: 'filterCategory', field: 'category' },
        { id: 'filterRank', field: 'rank' },
        { id: 'filterCity', field: 'city' },
        { id: 'filterSchool', field: 'school' }
    ];

    filters.forEach(filter => {
        const select = document.getElementById(filter.id);
        const currentValue = select.value;

        // Фильтруем данные по всем ОСТАЛЬНЫМ активным фильтрам
        const otherActiveFilters = filters.filter(f => f.id !== filter.id);
        const dataForThisFilter = allData.filter(item => {
            return otherActiveFilters.every(f => {
                const val = document.getElementById(f.id).value;
                if (!val) return true;
                
                // Специальная обработка для имени (если захотим добавить его в каскад)
                if (f.id === 'filterName') return item.name.toLowerCase().includes(val.toLowerCase());
                
                return item[f.field] === val;
            });
        });

        // Получаем уникальные значения
        const uniqueValues = [...new Set(dataForThisFilter.map(item => item[filter.field]))].filter(Boolean).sort();

        // Перерисовываем список
        const firstOpt = select.options[0]; // "Все ..." или "Выберите ..."
        select.innerHTML = '';
        select.appendChild(firstOpt);

        uniqueValues.forEach(val => {
            const opt = document.createElement('option');
            opt.value = val;
            opt.textContent = val;
            if (val === currentValue) opt.selected = true;
            select.appendChild(opt);
        });

        // Если текущее значение больше не доступно, сбрасываем на пустое
        if (currentValue && !uniqueValues.includes(currentValue)) {
            select.value = '';
        }

        // Особая логика для Ранга (разблокировка)
        if (filter.id === 'filterRank') {
            const catVal = document.getElementById('filterCategory').value;
            select.disabled = !catVal;
            if (!catVal) {
                select.options[0].textContent = 'Сначала выберите категорию';
                select.value = '';
            } else {
                select.options[0].textContent = 'Все ранги';
            }
        }
    });
}

function applyFilters() {
    const qName = document.getElementById('filterName').value.toLowerCase();
    const qDiscip = document.getElementById('filterDiscip').value;
    const qCategory = document.getElementById('filterCategory').value;
    const qRank = document.getElementById('filterRank').value;
    const qCity = document.getElementById('filterCity').value;
    const qSchool = document.getElementById('filterSchool').value;

    // Обновляем доступные опции в других фильтрах
    updateAllFiltersOptions();

    filteredData = allData.filter(item => {
        const matchName = !qName || item.name.toLowerCase().includes(qName);
        const matchDiscip = !qDiscip || item.discipline === qDiscip;
        const matchCategory = !qCategory || item.category === qCategory;
        const matchRank = !qRank || item.rank === qRank;
        const matchCity = !qCity || item.city === qCity;
        const matchSchool = !qSchool || item.school === qSchool;
        
        return matchName && matchDiscip && matchCategory && matchRank && matchCity && matchSchool;
    });

    updateDashboard();
}

function showOutdatedDataWarning() {
    const banner = document.createElement('div');
    banner.className = 'warning-banner';
    banner.innerHTML = `
        <div style="display: flex; align-items: center; gap: 15px;">
            <span style="font-size: 24px;">⚠️</span>
            <div>
                <strong>Обнаружены устаревшие данные!</strong><br>
                Чтобы увидеть новые фильтры (КМС/МС), возраст и статистику по школам, пожалуйста, 
                вернитесь в расширение и нажмите кнопку <strong>"Найти участников"</strong> снова.
            </div>
        </div>
        <button onclick="this.parentElement.remove()" class="btn btn-primary" style="background: rgba(255,255,255,0.2); border: 1px solid white;">Понятно</button>
    `;
    document.querySelector('.container').prepend(banner);
}

function updateDashboard() {
    renderTable();
    renderPerformanceStatsTable();
    renderStatsTable();
    
    const uniqueParticipants = new Set(filteredData.map(item => item.name + item.dob));
    const uniqueSchools = new Set(filteredData.map(item => item.school));
    const uniqueCities = new Set(filteredData.map(item => item.city));

    document.getElementById('countTotal').textContent = filteredData.length;
    document.getElementById('countUnique').textContent = uniqueParticipants.size;
    document.getElementById('countSchools').textContent = uniqueSchools.size;
    document.getElementById('countCities').textContent = uniqueCities.size;
    document.getElementById('resultsCount').textContent = `Показано: ${filteredData.length}`;
}

function renderTable() {
    const tbody = document.querySelector('#dataTable tbody');
    tbody.innerHTML = '';

    filteredData.forEach(item => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td style="font-weight: 500;">${item.name}</td>
            <td>${item.dob}</td>
            <td style="text-align: center;">${item.age || ''}</td>
            <td><span class="badge">${item.discipline}</span></td>
            <td>${item.category}</td>
            <td>${item.rank}</td>
            <td style="text-align: center;"><span class="perf-type">${item.perfType || '—'}</span></td>
            <td>${item.school}</td>
            <td>${item.city}</td>
            <td style="font-size: 13px; color: #5f6368;">${item.trainer}</td>
        `;
        tbody.appendChild(tr);
    });
}

function renderPerformanceStatsTable() {
    const tbody = document.getElementById('performanceStatsBody');
    if (!tbody) return;
    tbody.innerHTML = '';

    // Группировка: школа + город -> типы номеров
    // Учитываем, что один номер (заявка) может иметь несколько строк (по участникам)
    // Поэтому группируем по item._raw.id (если есть) или считаем уникальные комбинации
    const statsMap = new Map();
    
    // Используем Set для ID заявок внутри каждой группы, чтобы не считать участников одного дуэта как два дуэта
    filteredData.forEach(item => {
        const key = `${item.school}|||${item.city}`;
        if (!statsMap.has(key)) {
            statsMap.set(key, { solo: new Set(), duet: new Set(), group: new Set() });
        }
        
        const appId = item._raw ? item._raw.id : (item.name + item.discipline + item.category);
        if (item.perfType === 'Соло') statsMap.get(key).solo.add(appId);
        else if (item.perfType === 'Дуэт') statsMap.get(key).duet.add(appId);
        else if (item.perfType === 'Группа') statsMap.get(key).group.add(appId);
    });

    const statsArray = Array.from(statsMap.entries()).map(([key, counts]) => {
        const [school, city] = key.split('|||');
        return {
            label: `${school} (${city})`,
            solo: counts.solo.size,
            duet: counts.duet.size,
            group: counts.group.size,
            total: counts.solo.size + counts.duet.size + counts.group.size
        };
    });

    statsArray.sort((a, b) => b.total - a.total);

    statsArray.forEach(stat => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${stat.label}</td>
            <td style="text-align: center;">${stat.solo}</td>
            <td style="text-align: center;">${stat.duet}</td>
            <td style="text-align: center;">${stat.group}</td>
            <td style="text-align: center; font-weight: 600;">${stat.total}</td>
        `;
        tbody.appendChild(tr);
    });
}

function renderStatsTable() {
    const tbody = document.querySelector('#statsTable tbody');
    if (!tbody) return;
    tbody.innerHTML = '';

    // Группировка: школа -> город -> количество уникальных участников
    const statsMap = new Map();
    
    filteredData.forEach(item => {
        const key = `${item.school}|||${item.city}`;
        if (!statsMap.has(key)) {
            statsMap.set(key, new Set());
        }
        statsMap.get(key).add(item.name + item.dob);
    });

    const statsArray = [];
    statsMap.forEach((participants, key) => {
        const [school, city] = key.split('|||');
        statsArray.push({ school, city, count: participants.size });
    });

    // Сортировка по количеству (по убыванию)
    statsArray.sort((a, b) => b.count - a.count);

    statsArray.forEach(stat => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${stat.school}</td>
            <td>${stat.city}</td>
            <td style="text-align: center; font-weight: 600;">${stat.count}</td>
        `;
        tbody.appendChild(tr);
    });
}

function exportToExcel() {
    if (filteredData.length === 0) {
        showNotification('Нет данных для экспорта!', true);
        return;
    }

    // Подготовка заголовков и данных
    const headers = ['ФИО', 'Дата рождения', 'Полных лет', 'Дисциплина', 'Категория', 'Ранг', 'Тип номера', 'Школа', 'Город', 'Тренер'];
    const data = filteredData.map(item => [
        item.name, item.dob, item.age, item.discipline, item.category, item.rank, item.perfType, item.school, item.city, item.trainer
    ]);

    // Используем оригинальные данные из хранилища для чистого экспорта
    const compNameRaw = currentCompName || 'RAPA Соревнование';
    const compDatesRaw = currentCompDates || '';

    // Создаем массив с метаданными в начале
    const aoa = [
        [compNameRaw],
        [compDatesRaw],
        [], // Пустая строка для отступа
        headers,
        ...data
    ];

    // Создаем рабочий лист
    const ws = XLSX.utils.aoa_to_sheet(aoa);

    // Добавляем автофильтр на строку заголовков (она теперь 4-я, индекс 3)
    const range = XLSX.utils.decode_range(ws['!ref']);
    ws['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { c: 0, r: 3 }, e: { c: headers.length - 1, r: 3 } }) };

    // Настройка ширины колонок
    ws['!cols'] = [
        { wch: 30 }, // ФИО
        { wch: 15 }, // Дата рождения
        { wch: 10 }, // Лет
        { wch: 20 }, // Дисциплина
        { wch: 20 }, // Категория
        { wch: 15 }, // Ранг
        { wch: 12 }, // Тип номера
        { wch: 35 }, // Школа
        { wch: 15 }, // Город
        { wch: 25 }  // Тренер
    ];

    // Жирный шрифт для заголовков (через параметры ячеек, если библиотека поддерживает в текущей версии)
    // SheetJS Community version (xlsx.full.min.js) имеет ограниченную поддержку стилей, 
    // но автофильтры и базовая структура будут работать отлично.

    // Создаем книгу
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Участники");

    // Генерируем файл и инициируем скачивание
    const baseName = compNameRaw || 'RAPA_Export';
    const cleanDates = (compDatesRaw || '').replace(/[^\d\.]/g, '_').slice(0, 20);
    const translitName = transliterate(baseName);
    const fileName = `${translitName}_${cleanDates}.xlsx`.replace(/__+/g, '_').replace(/^_|_$/g, '');
    
    XLSX.writeFile(wb, fileName);
    
    showNotification(`Excel-файл "${fileName}" создан!`);
}

function transliterate(text) {
    const cyrillicToLatin = {
        'а': 'a', 'б': 'b', 'в': 'v', 'г': 'g', 'д': 'd', 'е': 'e', 'ё': 'yo', 'ж': 'zh', 'з': 'z',
        'и': 'i', 'й': 'y', 'к': 'k', 'л': 'l', 'м': 'm', 'н': 'n', 'о': 'o', 'п': 'p', 'р': 'r',
        'с': 's', 'т': 't', 'у': 'u', 'ф': 'f', 'х': 'kh', 'ц': 'ts', 'ч': 'ch', 'ш': 'sh', 'щ': 'sch',
        'ъ': '', 'ы': 'y', 'ь': '', 'э': 'e', 'ю': 'yu', 'я': 'ya',
        'А': 'A', 'Б': 'B', 'В': 'V', 'Г': 'G', 'Д': 'D', 'Е': 'E', 'Ё:': 'Yo', 'Ж': 'Zh', 'З': 'Z',
        'И': 'I', 'Й': 'Y', 'К': 'K', 'Л': 'L', 'М': 'M', 'Н': 'N', 'О': 'O', 'П': 'P', 'Р': 'R',
        'С': 'S', 'Т': 'T', 'У': 'U', 'Ф': 'F', 'Х': 'Kh', 'Ц': 'Ts', 'Ч': 'Ch', 'Ш': 'Sh', 'Щ': 'Sch',
        'Ъ': '', 'Ы': 'Y', 'Ь': '', 'Э': 'E', 'Ю': 'Yu', 'Я': 'Ya'
    };

    let result = text.split('').map(char => cyrillicToLatin[char] || char).join('');
    // Очистка спецсимволов для имени файла
    return result.replace(/[^a-zA-Z0-0\s\-_]/g, '').replace(/\s+/g, '_');
}

function showNotification(message, isError = false) {
    const el = document.getElementById('notification');
    el.textContent = message;
    el.style.backgroundColor = isError ? '#d93025' : '#202124';
    el.classList.add('show');
    
    setTimeout(() => {
        el.classList.remove('show');
    }, 3000);
}

// --- Логика объединения школ (School Merger) ---

function openMergerModal() {
    const modal = document.getElementById('schoolMergerModal');
    const targetSelect = document.getElementById('targetSchoolSelect');
    const aliasList = document.getElementById('aliasSchoolsList');
    
    // Получаем список всех уникальных школ из текущих данных
    const uniqueSchools = [...new Set(allData.map(item => item.school))].filter(Boolean).sort();
    
    // Наполняем dropdown главной школы
    targetSelect.innerHTML = '<option value="">-- Выберите главную школу --</option>';
    uniqueSchools.forEach(school => {
        const opt = document.createElement('option');
        opt.value = school;
        opt.textContent = school;
        targetSelect.appendChild(opt);
    });
    
    // Очищаем список галочек и блокируем кнопку
    aliasList.innerHTML = '<p style="color: #5f6368; font-size: 13px; font-style: italic;">Сначала выберите главную школу сверху.</p>';
    document.getElementById('confirmMergerBtn').disabled = true;
    
    modal.classList.add('show');
}

function closeMergerModal() {
    document.getElementById('schoolMergerModal').classList.remove('show');
}

function handleTargetSchoolChange(e) {
    const targetSchool = e.target.value;
    const aliasList = document.getElementById('aliasSchoolsList');
    const confirmBtn = document.getElementById('confirmMergerBtn');
    
    if (!targetSchool) {
        aliasList.innerHTML = '<p style="color: #5f6368; font-size: 13px; font-style: italic;">Сначала выберите главную школу сверху.</p>';
        confirmBtn.disabled = true;
        return;
    }
    
    const uniqueSchools = [...new Set(allData.map(item => item.school))].filter(Boolean).sort();
    aliasList.innerHTML = '';
    
    let hasOptions = false;
    uniqueSchools.forEach(school => {
        if (school === targetSchool) return; 
        hasOptions = true;
        
        const label = document.createElement('label');
        label.className = 'checkbox-item';
        label.innerHTML = `
            <input type="checkbox" value="${school.replace(/"/g, '&quot;')}">
            <span>${school}</span>
        `;
        // Разблокируем кнопку "Объединить", если выбрана хотя бы одна галочка
        label.querySelector('input').addEventListener('change', () => {
            const checkedBoxes = aliasList.querySelectorAll('input:checked');
            confirmBtn.disabled = checkedBoxes.length === 0;
            confirmBtn.textContent = checkedBoxes.length > 0 
                ? `Объединить выбранные (${checkedBoxes.length})`
                : 'Объединить выбранные';
        });
        aliasList.appendChild(label);
    });
    
    if (!hasOptions) {
        aliasList.innerHTML = '<p style="color: #5f6368; font-size: 13px; font-style: italic;">Нет других школ для объединения.</p>';
    }
    confirmBtn.disabled = true;
    confirmBtn.textContent = 'Объединить выбранные';
}

async function executeMerge() {
    const targetSchool = document.getElementById('targetSchoolSelect').value;
    const checkedBoxes = document.getElementById('aliasSchoolsList').querySelectorAll('input:checked');
    const schoolsToMerge = Array.from(checkedBoxes).map(cb => cb.value);
    
    if (!targetSchool || schoolsToMerge.length === 0) return;
    
    // 1. Обновляем данные в оперативной памяти (allData)
    allData.forEach(item => {
        if (schoolsToMerge.includes(item.school)) {
            item.school = targetSchool;
        }
    });
    
    // 2. Достаем текущие алиасы из базы
    const storageRes = await chrome.storage.local.get(['schoolAliases', 'competitionData']);
    let aliases = storageRes.schoolAliases || {};
    
    // 3. Запоминаем новые алиасы
    schoolsToMerge.forEach(oldSchool => {
        // Ключи сохраняем в нижнем регистре для надежного матчинга при парсинге
        const lowerKey = oldSchool.trim().toLowerCase();
        aliases[lowerKey] = targetSchool;
    });
    
    // 4. Обновляем хранилище Chrome (сохраняем очищенные данные и новые алиасы)
    await chrome.storage.local.set({ 
        schoolAliases: aliases,
        competitionData: allData
    });
    
    // 5. Перезапускаем фильтры и обновляем UI
    // Убираем очищение фильтра школы, если выбранная школа слита — она уже исчезла из списков
    applyFilters();
    closeMergerModal();
    showNotification(`Выбрано школ (${schoolsToMerge.length}) успешно объединены в "${targetSchool}". Правило сохранено!`);
}
