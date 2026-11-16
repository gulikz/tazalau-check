#!/bin/bash
# Tazalau Check - запуск на macOS из исходного кода (двойной щелчок в Finder).
cd "$(dirname "$0")" || exit 1
if [ ! -x ".venv/bin/python" ]; then
    echo "Первый запуск: устанавливаю компоненты..."
    python3 -m venv .venv && .venv/bin/python -m pip install --upgrade pip \
        && .venv/bin/python -m pip install -r requirements.txt || { echo "Не удалось установить компоненты"; exit 1; }
fi
exec .venv/bin/python tazalau_app.py "$@"
