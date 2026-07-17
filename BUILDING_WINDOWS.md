# Как скомпилировать программу под Windows (PyInstaller)

## 1. Требования
- Windows 10/11 x64
- Python **3.12** (рекомендуется x64)
- Установленные зависимости проекта

## 2. Установите зависимости
Вариант A (через Poetry):
```powershell
cd z:\programs\matrix_applications
poetry install
poetry run python -m src.main
```

Вариант B (без Poetry) — если используете pip напрямую:
```powershell
cd z:\programs\matrix_applications
pip install -U pip
pip install "pandas>=3,<4" "openpyxl>=3.1.5,<4" "tqdm>=4.67,<5" "xlrd>=2,<3" "tabulate>=0.10,<0.11"
```

## 3. Установите PyInstaller
```powershell
cd z:\programs\matrix_applications
poetry run pip install pyinstaller
```

## 4. Сборка exe
Запустите скрипт:
```powershell
cd z:\programs\matrix_applications
build_windows_pyinstaller.bat
```

После успешной сборки exe будет в:
- `dist\matrix-check\matrix-check.exe`

## 5. Если не собралось с первого раза
1) Попробуйте собрать в `--onedir` (уже используется в скрипте).  
2) Если PyInstaller ругается на скрытые импорты, скажите мне текст ошибки — подскажу, какие `--hidden-import` добавить.

