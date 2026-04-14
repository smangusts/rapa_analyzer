let allData = [];

function normalizeSchool(name) {
  if (!name) return 'Не указана';
  let lowerName = name.toLowerCase().replace(/«/g, '').replace(/»/g, '').replace(/"/g, '').trim();
  
  if (['академия', 'академия воздушной акробатики', 'академия спорта и воздушной акробатики'].includes(lowerName)) {
    return 'Академия Воздушной Акробатики (объединенная)';
  }
  if (lowerName.includes('efiria') || lowerName.includes('эфирия')) return 'Эфирия';
  if (lowerName.includes('new trend') || lowerName.includes('new tend')) return 'New Trend';
  if (lowerName.includes('pro уровень') || lowerName.includes('pro уровнь')) return 'PRO уровень';
  if (lowerName.includes('аверс')) return 'Аверс';
  if (lowerName.includes('эклипс')) return 'Академия Эклипс';
  if (lowerName.includes('астерия') || lowerName.includes('asteria')) return 'Центр воздушной акробатики и спорта "Астерия"';
  if (lowerName.includes('астар')) return 'Астар';
  if (lowerName.includes('топ фит') || lowerName.includes('top fit')) return 'Top fit';
  if (lowerName.includes('рожковой')) return 'Школа Воздушной Акробатики Елены Рожковой';
  
  return name;
}

document.getElementById('fetchBtn').addEventListener('click', async () => {
  const errorDiv = document.getElementById('error');
  const loader = document.getElementById('loader');
  const btn = document.getElementById('fetchBtn');
  
  errorDiv.textContent = '';
  loader.style.display = 'inline-block';
  btn.disabled = true;
  document.getElementById('resultsPanel').style.display = 'none';
  allData = [];

  try {
    const tabs = await chrome.tabs.query({active: true, currentWindow: true});
    if (!tabs || !tabs.length) throw new Error("Вкладка не найдена.");
    
    const activeTab = tabs[0];
    const url = activeTab.url;
    const match = url.match(/([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})/);
    
    if (!match) {
      throw new Error("Откройте страницу соревнования на RAPA (должен быть ID в адресной строке).");
    }
    const compId = match[1];
    let page = 1;
    const take = 50;
    
    // Загружаем сохраненные алиасы школ
    let schoolAliases = {};
    try {
        const stored = await chrome.storage.local.get(['schoolAliases']);
        if (stored.schoolAliases) {
            schoolAliases = stored.schoolAliases;
        }
    } catch (e) {
        console.warn('Не удалось загрузить алиасы школ:', e);
    }

    // Получаем метаданные соревнования (название и даты)
    let compName = 'Соревнование RAPA';
    let compDates = '';
    try {
      const metaRes = await fetch(`https://referee.f-rapa.ru/api/competitions/${compId}/public`);
      if (metaRes.ok) {
        const metaJson = await metaRes.json();
        let name = metaJson.name || compName;
        const city = (metaJson.city && metaJson.city.name) ? metaJson.city.name : '';
        if (city) name += ` — ${city}`;
        compName = name;
        
        if (metaJson.date_from && metaJson.date_to) {
          const d1 = new Date(metaJson.date_from).toLocaleDateString('ru-RU');
          const d2 = new Date(metaJson.date_to).toLocaleDateString('ru-RU');
          compDates = `${d1} — ${d2}`;
        }
      }
    } catch (e) {
      console.warn('Не удалось получить метаданные соревнования:', e);
    }
    
    while (true) {
      const apiUrl = `https://referee.f-rapa.ru/api/competitions/${compId}/public/applications?page=${page}&take=${take}`;
      const res = await fetch(apiUrl);
      if (!res.ok) throw new Error("Ошибка при запросе к сайту RAPA.");
      
      const json = await res.json();
      const items = json.data || [];
      if (!items.length) break;
      
      for (const item of items) {
        const cityName = (item.city && item.city.name) ? item.city.name : 'Не указан';
        
        // Надежное извлечение и нормализация школы
        const rawSchool = item.trainer_place || '';
        const trimmedSchool = rawSchool.trim();
        const lowerSchoolKey = trimmedSchool.toLowerCase();
        
        let school = trimmedSchool;
        if (schoolAliases[lowerSchoolKey]) {
            school = schoolAliases[lowerSchoolKey]; // Применяем ручной маппинг пользователя
        } else {
            school = normalizeSchool(rawSchool); // Применяем базовую эвристику
        }
        
        const trainer = item.trainer_fullname || 'Не указан';
        
        // Более гибкое извлечение (Defensive extraction)
        const dcr = item.discipline_category_rank || {};
        const rank = dcr.name || 'Без разряда';
        
        // Пытаемся найти категорию и дисциплину разными путями
        let categoryName = 'Без категории';
        let disciplineName = 'Без дисциплины';
        
        if (dcr.category) {
            categoryName = dcr.category.name || categoryName;
            if (dcr.category.discipline) {
                disciplineName = dcr.category.discipline.name || disciplineName;
            }
        }

        // Определение типа номера
        const participants = item.participants || [];
        let perfType = 'Группа';
        if (participants.length === 1) perfType = 'Соло';
        else if (participants.length === 2) perfType = 'Дуэт';

        // Если данных критически не хватает - логируем для отладки
        if (rank === 'Без разряда' && categoryName === 'Без категории') {
            console.warn('Обнаружена неполная заявка:', item);
        }
        
        for (const p of participants) {
          const fullname = `${p.lastname || ''} ${p.firstname || ''} ${p.middlename || ''}`.trim();
          
          if (!p.lastname && !p.firstname) {
              console.warn('Участник без имени в заявке ID:', item.id);
              continue; 
          }

          // Расчет возраста
          let age = '';
          if (p.birthdate) {
            const birth = new Date(p.birthdate);
            const now = new Date();
            age = now.getFullYear() - birth.getFullYear();
            const m = now.getMonth() - birth.getMonth();
            if (m < 0 || (m === 0 && now.getDate() < birth.getDate())) {
              age--;
            }
          }

          // Разделение КМС и МС
          let refinedCategory = categoryName;
          const lowerCat = categoryName.toLowerCase();
          const lowerRank = rank.toLowerCase();
          
          if (lowerCat.includes('кмс') || lowerCat.includes('мс')) {
            if (lowerRank.includes('кмс')) refinedCategory = 'КМС';
            else if (lowerRank.includes('мс')) refinedCategory = 'МС';
          }

          allData.push({
            name: fullname,
            dob: p.birthdate || '',
            age: age,
            discipline: disciplineName,
            category: refinedCategory,
            rank: rank,
            school: school,
            city: cityName,
            trainer: trainer,
            perfType: perfType, // Тип номера (Соло/Дуэт/Группа)
            _raw: item 
          });
        }
      }
      
      const meta = json.meta || {};
      if (!meta.hasNextPage) break;
      page++;
    }
    
    // Сохраняем данные в хранилище Chrome
    await chrome.storage.local.set({ 
      competitionData: allData,
      competitionUrl: url,
      competitionId: compId,
      competitionName: compName,
      competitionDates: compDates,
      timestamp: new Date().toISOString()
    });

    // Открываем новую вкладку с красивыми результатами
    chrome.tabs.create({ url: 'results.html' });
    
    // Закрываем поп-ап
    window.close();

  } catch (err) {
    errorDiv.textContent = err.message || "Произошла неизвестная ошибка при парсинге.";
  } finally {
    loader.style.display = 'none';
    btn.disabled = false;
  }
});

// Обработка импорта из Excel
document.getElementById('importBtn').addEventListener('click', () => {
  document.getElementById('importExcel').click();
});

document.getElementById('importExcel').addEventListener('change', async (e) => {
  const file = e.target.files[0];
  if (!file) return;

  const loader = document.getElementById('loader');
  const errorDiv = document.getElementById('error');
  loader.style.display = 'inline-block';
  errorDiv.textContent = '';

  try {
    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data, { type: 'array' });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    
    // Преобразуем в массив массивов
    const aoa = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    
    if (aoa.length < 5) {
      throw new Error("Файл кажется пустым или имеет неверный формат.");
    }

    // Извлекаем метаданные: A1 - название, A2 - даты
    const compNameRaw = aoa[0] ? aoa[0][0] : 'Загружено из файла';
    const compDatesRaw = aoa[1] ? aoa[1][0] : '';
    
    // Данные начинаются с 5-й строки (индекс 4)
    const importedRows = aoa.slice(4);
    const competitionData = importedRows.filter(row => row && row[0]).map(row => ({
      name: row[0] || '',
      dob: row[1] || '',
      age: row[2] || '',
      discipline: row[3] || '',
      category: row[4] || '',
      rank: row[5] || '',
      perfType: row[6] || '',
      school: row[7] || '',
      city: row[8] || '',
      trainer: row[9] || ''
    }));

    if (!competitionData.length) {
      throw new Error("Не удалось найти данные участников в файле.");
    }

    // Сохраняем в хранилище
    await chrome.storage.local.set({ 
      competitionData: competitionData,
      competitionUrl: 'offline_file',
      competitionId: 'offline',
      competitionName: compNameRaw,
      competitionDates: compDatesRaw,
      timestamp: new Date().toISOString()
    });

    // Открываем результаты
    chrome.tabs.create({ url: 'results.html' });
    window.close();

  } catch (err) {
    errorDiv.textContent = "Ошибка импорта: " + err.message;
  } finally {
    loader.style.display = 'none';
  }
});

// --- Вспомогательные функции для Google Sheets ---

async function getAuthToken(interactive = true) {
  return new Promise((resolve, reject) => {
    chrome.identity.getAuthToken({ interactive }, (token) => {
      if (chrome.runtime.lastError) {
        reject(chrome.runtime.lastError);
      } else {
        resolve(token);
      }
    });
  });
}

async function createSpreadsheet(token, title) {
  const res = await fetch('https://sheets.googleapis.com/v4/spreadsheets', {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${token}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      properties: { title: title }
    })
  });
  if (!res.ok) throw new Error('Не удалось создать таблицу в Google');
  return await res.json();
}

async function updateSheetValues(token, spreadsheetId, range, values) {
  const res = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${range}?valueInputOption=RAW`, {
    method: 'PUT',
    headers: {
      'Authorization': `Bearer ${token}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({ values })
  });
  if (!res.ok) throw new Error('Не удалось записать данные в таблицу');
}

// --- Обработка Google Sheets ---

document.getElementById('exportGSBtn').addEventListener('click', async () => {
  const errorDiv = document.getElementById('error');
  const loader = document.getElementById('loader');
  
  // Проверяем, есть ли данные в хранилище для экспорта
  const storage = await chrome.storage.local.get(['competitionData', 'competitionName', 'competitionDates']);
  if (!storage.competitionData || !storage.competitionData.length) {
    errorDiv.textContent = 'Сначала загрузите данные онлайн или из Excel.';
    return;
  }

  loader.style.display = 'inline-block';
  errorDiv.textContent = '';

  try {
    const token = await getAuthToken();
    const title = `RAPA Анализ: ${storage.competitionName || 'Турнир'}`;
    const spreadsheet = await createSpreadsheet(token, title);
    const spreadsheetId = spreadsheet.spreadsheetId;

    const headers = ['ФИО', 'Дата рождения', 'Полных лет', 'Дисциплина', 'Категория', 'Ранг', 'Тип номера', 'Школа', 'Город', 'Тренер'];
    const rows = storage.competitionData.map(item => [
      item.name, item.dob, item.age, item.discipline, item.category, item.rank, item.perfType, item.school, item.city, item.trainer
    ]);

    const values = [
      [storage.competitionName],
      [storage.competitionDates],
      [],
      headers,
      ...rows
    ];

    await updateSheetValues(token, spreadsheetId, 'Sheet1!A1', values);
    
    // Показываем ссылку и открываем
    window.open(`https://docs.google.com/spreadsheets/d/${spreadsheetId}`, '_blank');
    loader.style.display = 'none';
    errorDiv.style.color = '#34a853';
    errorDiv.textContent = 'Успешно экспортировано в Google Таблицы!';

  } catch (err) {
    console.error(err);
    errorDiv.style.color = '#d93025';
    errorDiv.textContent = 'Ошибка Google Sheets: ' + (err.message || 'Проверьте Client ID');
    loader.style.display = 'none';
  }
});

document.getElementById('importGSBtn').addEventListener('click', async () => {
  const url = prompt('Вставьте ID Google Таблицы (из адреса документа) или полную ссылку:');
  if (!url) return;

  let spreadsheetId = url;
  const match = url.match(/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (match) spreadsheetId = match[1];

  const errorDiv = document.getElementById('error');
  const loader = document.getElementById('loader');
  loader.style.display = 'inline-block';
  errorDiv.textContent = '';

  try {
    const token = await getAuthToken();
    const res = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/Sheet1!A1:J1000`, {
      headers: { 'Authorization': `Bearer ${token}` }
    });
    
    if (!res.ok) throw new Error('Не удалось прочитать таблицу. Проверьте доступ.');
    const json = await res.json();
    const aoa = json.values;

    if (!aoa || aoa.length < 5) throw new Error('Таблица пуста или имеет неверный формат.');

    const compName = aoa[0] ? aoa[0][0] : 'Импорт из Google';
    const compDates = aoa[1] ? aoa[1][0] : '';
    
    const importedRows = aoa.slice(4);
    const competitionData = importedRows.filter(r => r && r[0]).map(row => ({
      name: row[0] || '',
      dob: row[1] || '',
      age: row[2] || '',
      discipline: row[3] || '',
      category: row[4] || '',
      rank: row[5] || '',
      perfType: row[6] || '',
      school: row[7] || '',
      city: row[8] || '',
      trainer: row[9] || ''
    }));

    await chrome.storage.local.set({ 
      competitionData,
      competitionUrl: 'google_sheets',
      competitionId: 'gs_' + spreadsheetId,
      competitionName: compName,
      competitionDates: compDates,
      timestamp: new Date().toISOString()
    });

    chrome.tabs.create({ url: 'results.html' });
    window.close();

  } catch (err) {
    errorDiv.textContent = 'Ошибка импорта: ' + err.message;
    loader.style.display = 'none';
  }
});

// Функции фильтрации переехали в results.js
