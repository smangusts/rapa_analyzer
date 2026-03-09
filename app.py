import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import collections
import io

st.set_page_config(page_title="Анализ Воздушной Гимнастики", page_icon="🎪", layout="wide")

def normalize_school(name):
    lower_name = str(name).lower().replace('«', '').replace('»', '').replace('"', '').strip()
    
    if lower_name in ['академия', 'академия воздушной акробатики', 'академия спорта и воздушной акробатики']:
        return 'Академия Воздушной Акробатики (объединенная)'
        
    if 'efiria' in lower_name or 'эфирия' in lower_name: return 'Эфирия'
    if 'new trend' in lower_name or 'new tend' in lower_name: return 'New Trend'
    if 'pro уровень' in lower_name or 'pro уровнь' in lower_name: return 'PRO уровень'
    if 'аверс' in lower_name: return 'Аверс'
    if 'эклипс' in lower_name: return 'Академия Эклипс'
    if 'астерия' in lower_name or 'asteria' in lower_name: return 'Центр воздушной акробатики и спорта "Астерия"'
    if 'астар' in lower_name: return 'Астар'
    if 'топ фит' in lower_name or 'top fit' in lower_name: return 'Top fit'
    if 'рожковой' in lower_name: return 'Школа Воздушной Акробатики Елены Рожковой'
    return name

@st.cache_data
def process_excel(file_bytes):
    """
    Парсит загруженный excel-файл в удобный Pandas Dataframe со всеми нужными полями.
    """
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(min_row=2, max_col=4, values_only=True))
    
    people = []
    current_person = None
    
    for row in rows:
        participant, disc, comp, desc = row
        
        is_new_person = False
        p_str = str(participant).strip() if participant else ''
        if p_str and p_str.lower() != 'nan':
            # Логика определения нового участника (имя + фамилия)
            if any(c.isalpha() for c in p_str) and not p_str.startswith('жен.,') and not p_str.startswith('муж.,') and 'лет ' not in p_str.lower() and not p_str.lower().startswith('дуэты'):
                is_new_person = True
    
        if is_new_person:
            if current_person:
                people.append(current_person)
            current_person = {
                'ФИО Участника': p_str,
                'Дисциплина': [],
                'Категория': [],
                'Город': 'Не указан',
                'Школа': 'Не указана',
                'Тренер': 'Не указан',
                'Возраст/Пол': ''
            }
    
        if current_person:
            # Возраст/пол может быть в ячейке participant на следующей строке
            if participant and not is_new_person:
                if 'жен.,' in str(participant) or 'муж.,' in str(participant):
                    parts = str(participant).split(',')
                    age = [pt for pt in parts if 'лет' in pt]
                    current_person['Возраст/Пол'] = age[0].strip() if age else str(participant)
            
            if disc: current_person['Дисциплина'].append(str(disc))
            if comp: current_person['Категория'].append(str(comp))
            if desc:
                desc_str = str(desc).strip()
                if desc_str.startswith('Город:'):
                    current_person['Город'] = desc_str.replace('Город:', '').strip()
                elif desc_str.startswith('Школа'):
                    current_person['Школа'] = desc_str.split(':', 1)[-1].strip()
                elif desc_str.startswith('ФИО тренера:'):
                    current_person['Тренер'] = desc_str.replace('ФИО тренера:', '').strip()
    
    if current_person:
        people.append(current_person)
        
    # Преобразуем списки в строки для удобного отображения в таблице и причесываем школы
    for p in people:
        p['Дисциплина'] = ' | '.join(list(dict.fromkeys(p['Дисциплина']))) # Убираем дубли
        p['Категория'] = ' | '.join(list(dict.fromkeys(p['Категория'])))
        p['Школа'] = normalize_school(p['Школа'])
        
    df = pd.DataFrame(people)
    return df

@st.cache_data
def process_url(url):
    import re
    import requests
    
    match = re.search(r'([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})', url)
    if not match:
        raise ValueError("Некорректная ссылка. Убедитесь, что она взята с сайта (содержит ID соревнования).")
    
    comp_id = match.group(1)
    
    people = []
    page = 1
    take = 100
    
    while True:
        api_url = f"https://referee.f-rapa.ru/api/competitions/{comp_id}/public/applications?page={page}&take={take}"
        resp = requests.get(api_url)
        resp.raise_for_status()
        
        data = resp.json()
        items = data.get('data', [])
        
        if not items:
            break
            
        for item in items:
            city_name = item.get('city', {}).get('name', 'Не указан') if item.get('city') else 'Не указан'
            school = item.get('trainer_place') or 'Не указана'
            trainer = item.get('trainer_fullname') or 'Не указан'
            
            dcr = item.get('discipline_category_rank', {})
            rank = dcr.get('name', '')
            category = dcr.get('category', {}).get('name', '')
            discipline = dcr.get('category', {}).get('discipline', {}).get('name', '')
            full_category = f"{rank} — {category}".strip(' —')
            
            for part in item.get('participants', []):
                lname = part.get('lastname') or ''
                fname = part.get('firstname') or ''
                mname = part.get('middlename') or ''
                fullname = f"{lname} {fname} {mname}".strip()
                bdate = part.get('birthdate', '')
                
                people.append({
                    'ФИО Участника': fullname,
                    'Дисциплина': discipline,
                    'Категория': full_category,
                    'Город': city_name,
                    'Школа': school,
                    'Тренер': trainer,
                    'Возраст/Пол': bdate
                })
                
        meta = data.get('meta', {})
        if not meta.get('hasNextPage'):
            break
        page += 1
        
    for p in people:
        p['Школа'] = normalize_school(p['Школа'])
        
    return pd.DataFrame(people)

st.title("🎪 Анализ соревнований по воздушной гимнастике")
st.markdown("Загрузите файл Excel формата (как `пермь.xlsx`) или вставьте ссылку на сайт RAPA, чтобы получить доступ к красивой и удобной фильтрации данных участников.")

data_source = st.radio("Выберите источник данных:", ["Ссылка на соревнование (RAPA)", "Excel файл"])

df = None

if data_source == "Excel файл":
    uploaded_file = st.file_uploader("Выберите Excel файл", type=["xlsx"])
    if uploaded_file is not None:
        with st.spinner('Обработка файла...'):
            df = process_excel(uploaded_file.getvalue())
        st.success("Файл успешно загружен и обработан!")
else:
    url_input = st.text_input("Введите ссылку (например, https://referee.f-rapa.ru/public/competition-applications/...)", "")
    if url_input:
        with st.spinner('Скачивание данных с сайта...'):
            try:
                df = process_url(url_input)
                st.success("Данные успешно загружены с сайта!")
            except Exception as e:
                st.error(f"Ошибка при загрузке: {e}")

if df is not None and not df.empty:
    
    st.sidebar.header("Панели фильтрации")
    
    # Text input for Name
    search_name = st.sidebar.text_input("🔍 Поиск по ФИО участника", "")
    
    # Multiselects
    cities = sorted(list(df['Город'].unique()))
    selected_cities = st.sidebar.multiselect("Город", cities, default=[])
    
    schools = sorted(list(df['Школа'].unique()))
    selected_schools = st.sidebar.multiselect("Школа", schools, default=[])
    
    coaches = sorted(list([c for c in df['Тренер'].unique() if c != 'Не указан']))
    selected_coaches = st.sidebar.multiselect("Тренер", coaches, default=[])
    
    # Since Discip lines might be compound strings, we do a text match, or we could extract unique tokens. Let's do text search for simplicity
    search_discipline = st.sidebar.text_input("Дисциплина (например, 'кольцо', 'полотна')", "")
    
    # Apply Filters
    filtered_df = df.copy()
    
    if search_name:
        filtered_df = filtered_df[filtered_df['ФИО Участника'].str.contains(search_name, case=False, na=False)]
        
    if selected_cities:
        filtered_df = filtered_df[filtered_df['Город'].isin(selected_cities)]
        
    if selected_schools:
        filtered_df = filtered_df[filtered_df['Школа'].isin(selected_schools)]
        
    if selected_coaches:
        filtered_df = filtered_df[filtered_df['Тренер'].isin(selected_coaches)]
        
    if search_discipline:
        filtered_df = filtered_df[filtered_df['Дисциплина'].str.contains(search_discipline, case=False, na=False)]
        
    # Metrics
    st.subheader("📊 Текущая статистика по фильтрам")
    col1, col2, col3 = st.columns(3)
    col1.metric("Всего найдено участников", len(filtered_df))
    col2.metric("Уникальных школ", filtered_df['Школа'].nunique())
    col3.metric("Уникальных городов", filtered_df['Город'].nunique())
    
    st.markdown("### Таблица участников")
    st.dataframe(filtered_df, use_container_width=True, height=600)
    
else:
    st.info("Пожалуйста, загрузите файл на панели выше, чтобы начать работу.")
