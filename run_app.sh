#!/bin/bash
# Активация окружения и запуск приложения Streamlit
if [ -d "venv" ]; then
    source venv/bin/activate
elif [ -d ".venv" ]; then
    source .venv/bin/activate
fi

streamlit run app.py
