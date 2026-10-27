# Как скомпилировать программу под Windows (PyInstaller)

## 1. Требования
- Windows 10/11 **x64**
- Python **3.12 x64** (не 32-bit)
- [Microsoft Visual C++ 2015–2022 Redistributable (x64)](https://aka.ms/vs/17/release/vc_redist.x64.exe)

## 2. Установите зависимости

```powershell
cd C:\Users\User\programs\matrix_applications
poetry install
poetry run python -m src.main
```

Если без Poetry:

```powershell
pip install -U pip
pip install "pandas>=3,<4" "openpyxl>=3.1.5,<4"
pip install "pytest>=9,<10"   # для тестов (опционально)
```

## 3. Установите зависимости (обязательно)

```powershell
cd C:\Users\User\programs\matrix_applications
poetry install --extras "dev build"
poetry run python -c "import pandas, openpyxl; print('OK')"
```

Если команда выше падает — exe соберётся без `pandas` и не запустится.

## 4. Сборка exe

```powershell
cd C:\Users\User\programs\matrix_applications
build_windows_pyinstaller.bat
```

Скрипт сам выполняет `poetry install --extras build` и проверяет `pandas/openpyxl` перед сборкой.

## 5. Где exe

После успешной сборки запускайте **только**:

```
dist\matrix-check\matrix-check.exe
```

Важно:
- **НЕ** запускайте exe из папки `build\` — там временные файлы, оттуда и ошибка `Failed to load Python DLL`.
- При переносе на другой ПК копируйте **всю папку** `dist\matrix-check\`, а не один `.exe`.

## 6. Ошибка `ModuleNotFoundError: No module named 'pandas'`

Зависимости не были установлены в venv перед сборкой.

```powershell
poetry install --extras build
poetry run python -c "import pandas, openpyxl; print('OK')"
build_windows_pyinstaller.bat
```

## 7. Ошибка `Failed to load Python DLL`

1. Запускайте из `dist\matrix-check\`, не из `build\`.
2. Установите **VC++ Redistributable x64** (ссылка выше).
3. Пересоберите:
   ```powershell
   build_windows_pyinstaller.bat
   ```
4. Убедитесь, что Python **64-bit**:
   ```powershell
   python -c "import struct; print(struct.calcsize('P')*8)"
   ```
   Должно вывести `64`.

## 8. Ошибка `Tcl data directory ... not found`

PyInstaller не упаковал библиотеки tkinter. В актуальном `matrix-check.spec` это исправлено — папки `tcl8.6` и `tk8.6` кладутся в `_internal\_tcl_data` и `_internal\_tk_data`.

Пересоберите:
```powershell
build_windows_pyinstaller.bat
```

Если ошибка останется — переустановите Python с галочкой **tcl/tk and IDLE**.

## 9. Ручная сборка (если нужно)

```powershell
poetry run python -m pip install -U pyinstaller
poetry run python -m PyInstaller --clean --noconfirm matrix-check.spec
```
