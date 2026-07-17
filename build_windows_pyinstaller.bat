@echo off
setlocal enabledelayedexpansion

REM Сборка GUI-приложения для Windows через PyInstaller.
REM Результат: dist\matrix-check\matrix-check.exe

cd /d "%~dp0"

echo === Проверка Python ===
where poetry >nul 2>&1
if %errorlevel%==0 (
  set "PY_CMD=poetry run python"
) else (
  set "PY_CMD=python"
)

%PY_CMD% --version
if errorlevel 1 (
  echo Python не найден. Установите Python 3.12 x64.
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
