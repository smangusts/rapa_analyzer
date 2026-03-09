#!/bin/bash

# Устанавливаем пути относительно домашней папки или текущей директории
VENV_DIR="$HOME/.gemini/antigravity/scratch/venv"
# Если скрипт запущен из папки проекта
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
APP_PATH="$SCRIPT_DIR/app.py"

# Выводим сообщение красивым цветом
echo -e "\033[1;36mЗапуск приложения для анализа соревнований...\033[0m"

# Проверяем, существует ли папка с виртуальным окружением
if [ ! -d "$VENV_DIR" ]; then
    # Пробуем локальное venv
    VENV_DIR="$SCRIPT_DIR/venv"
    if [ ! -d "$VENV_DIR" ]; then
        VENV_DIR="$SCRIPT_DIR/.venv"
    fi
fi

if [ -d "$VENV_DIR" ]; then
    source "$VENV_DIR/bin/activate"
else
    echo -e "\033[1;33mПредупреждение: Виртуальное окружение не найдено. Попытка запуска с системным python...\033[0m"
fi

# Проверяем, существует ли файл приложения
if [ ! -f "$APP_PATH" ]; then
    echo -e "\033[1;31mОшибка: Файл приложения не найден по пути $APP_PATH\033[0m"
    read -p "Нажмите Enter для выхода..."
    exit 1
fi

# Запускаем Streamlit
echo "Открываем веб-интерфейс (Streamlit)..."
streamlit run "$APP_PATH"

# Оставляем окно открытым, если приложение случайно закроется
read -p "Приложение остановлено. Нажмите Enter для выхода..."
