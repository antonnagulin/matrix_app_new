@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

REM Сборка GUI-приложения для Windows через PyInstaller.
REM Результат: dist\matrix-check\matrix-check.exe

cd /d "%~dp0"

set "USE_POETRY=0"
where poetry >nul 2>&1
if %errorlevel%==0 (
  set "USE_POETRY=1"
  set "PY_CMD=poetry run python"
) else (
  set "PY_CMD=python"
)

echo === Проверка Python ===
%PY_CMD% --version
if errorlevel 1 (
  echo Python не найден. Установите Python 3.12 x64.
  pause
  exit /b 1
)

echo.
echo === Установка зависимостей проекта ===
if "%USE_POETRY%"=="1" (
  poetry install --extras build --no-root
  if errorlevel 1 (
    echo Не удалось установить зависимости через poetry install.
    pause
    exit /b 1
  )
) else (
  %PY_CMD% -m pip install -U pip
  %PY_CMD% -m pip install "pandas>=3,<4" "openpyxl>=3.1.5,<4" "tqdm>=4.67,<5" "xlrd>=2,<3" "tabulate>=0.10,<0.11" "pyinstaller>=6,<7"
)

echo.
echo === Проверка pandas/openpyxl ===
%PY_CMD% -c "import pandas, openpyxl; print('pandas/openpyxl OK')"
if errorlevel 1 (
  echo.
  echo ОШИБКА: pandas или openpyxl не установлены в окружении сборки.
  echo Выполните: poetry install
  pause
  exit /b 1
)

echo.
echo === Установка PyInstaller ===
%PY_CMD% -m pip install -U pyinstaller
if errorlevel 1 (
  echo Не удалось установить PyInstaller.
  pause
  exit /b 1
)

echo.
echo === Очистка старых сборок ===
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo.
echo === Сборка ===
%PY_CMD% -m PyInstaller --clean --noconfirm matrix-check.spec
if errorlevel 1 (
  echo Сборка завершилась с ошибкой.
  pause
  exit /b 1
)

echo.
echo ========================================
echo  ГОТОВО!
echo.
echo  Запускайте ТОЛЬКО:
echo    dist\matrix-check\matrix-check.exe
echo.
echo  НЕ запускайте exe из папки build\
echo  При переносе копируйте всю папку dist\matrix-check\
echo ========================================
echo.
pause
