# Инструкция по настройке Google Sheets API

Для того чтобы экспорт в Google Таблицы заработал, необходимо получить ваш персональный `client_id`. Это бесплатно и занимает 5 минут.

## Шаг 1: Создание проекта в Google Cloud
1. Перейдите в [Google Cloud Console](https://console.cloud.google.com/).
2. Нажмите **"Select a project"** (или название текущего) вверху и создайте новый проект: **"New Project"** (назовите его `RAPA Analyzer`).

## Шаг 2: Включение API
1. Перейдите в раздел **"APIs & Services"** > **"Library"**.
2. Найдите и включите два API:
   - **Google Sheets API**
   - **Google Drive API**

## Шаг 3: Настройка экрана согласия (OAuth Consent Screen)
1. Перейдите в **"APIs & Services"** > **"OAuth consent screen"**.
2. Выберите тип **"External"** и нажмите "Create".
3. Заполните обязательные поля:
   - App name: `RAPA Analyzer`
   - User support email: ваш email
   - Developer contact info: ваш email
4. Нажмите "Save and Continue" во всех разделах. (Scopes можно пропустить на этом этапе).
5. На вкладке "Test users" добавьте ваш email, чтобы вы могли пользоваться приложением.

## Шаг 4: Получение Client ID
1. Перейдите в **"APIs & Services"** > **"Credentials"**.
2. Нажмите **"+ CREATE CREDENTIALS"** > **"OAuth client ID"**.
3. Application type выбирайте: **"Chrome extension"**.
4. В поле **"Item ID"** нужно вставить ID расширения.
   - Как узнать ID: Откройте `chrome://extensions`, найдите RAPA Analyzer. Его ID это длинная строка букв.
5. Нажмите "Create". Вы получите **Client ID**.

## Шаг 5: Обновление расширения
1. Откройте файл `manifest.json` в папке расширения.
2. Замените `"YOUR_CLIENT_ID_HERE.apps.googleusercontent.com"` на ваш полученный ID.
3. Сохраните файл.
4. В `chrome://extensions` нажмите кнопку "Обновить" (Reload) у расширения RAPA Analyzer.

**Готово! Теперь кнопки Google Таблиц будут работать.**
