#!/usr/bin/env bash
# Сборка GUI-приложения для Linux (Astra 1.8 и др.) через PyInstaller.
# Результат: dist/matrix-check/matrix-check

set -euo pipefail

cd "$(dirname "$0")"

PY_CMD="python3"
USE_POETRY=0
if command -v poetry >/dev/null 2>&1; then
  USE_POETRY=1
  PY_CMD="poetry run python"
fi

echo "=== Проверка Python ==="
$PY_CMD --version

echo
echo "=== Установка зависимостей проекта ==="
if [[ "$USE_POETRY" -eq 1 ]]; then
  poetry install --extras build --no-root
else
  $PY_CMD -m pip install -U pip
  $PY_CMD -m pip install \
    "pandas>=3,<4" "openpyxl>=3.1.5,<4" "tqdm>=4.67,<5" \
    "xlrd>=2,<3" "tabulate>=0.10,<0.11" "pyinstaller>=6,<7"
fi

echo
echo "=== Проверка tkinter ==="
$PY_CMD -c "import tkinter; tkinter.Tk().destroy(); print('tkinter OK')" || {
  echo
  echo "ОШИБКА: tkinter не установлен."
  echo "На Astra Linux 1.8 выполните:"
  echo "  sudo apt install python3-tk tcl tk"
  exit 1
}

echo
echo "=== Проверка pandas/openpyxl ==="
$PY_CMD -c "import pandas, openpyxl; print('pandas/openpyxl OK')"

echo
echo "=== Установка PyInstaller ==="
$PY_CMD -m pip install -U pyinstaller

echo
echo "=== Очистка старых сборок ==="
rm -rf build dist

echo
echo "=== Сборка ==="
$PY_CMD -m PyInstaller --clean --noconfirm matrix-check-linux.spec

echo
echo "========================================"
echo " ГОТОВО!"
echo
echo " Запускайте:"
echo "   ./dist/matrix-check/matrix-check"
echo
echo " При переносе копируйте всю папку dist/matrix-check/"
echo "========================================"
