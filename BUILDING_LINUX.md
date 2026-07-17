# Сборка matrix-check на Linux (Astra Linux 1.8)

Приложение с GUI на **tkinter**. Сборка через **PyInstaller** в каталог `dist/matrix-check/`.

## 1. Системные пакеты (Astra Linux 1.8)

```bash
sudo apt update
sudo apt install -y python3 python3-venv python3-pip python3-tk tcl tk \
  python3-dev build-essential xdg-utils
```

- `python3-tk` — обязателен для GUI и сборки
- `xdg-utils` — открытие папки с результатом в файловом менеджере

Python **3.11+** (в Astra 1.8 обычно `python3` → 3.11).

## 2. Poetry (рекомендуется)

```bash
curl -sSL https://install.python-poetry.org | python3 -
# или: sudo apt install python3-poetry
```

## 3. Зависимости проекта

```bash
cd /path/to/matrix_applications
poetry install --extras build --no-root
poetry run python -c "import tkinter, pandas, openpyxl; print('OK')"
```

## 4. Запуск из исходников

```bash
chmod +x run.sh
./run.sh
# или
poetry run python -m src.main
```

## 5. Сборка исполняемого файла

```bash
chmod +x build_linux_pyinstaller.sh
./build_linux_pyinstaller.sh
```

Результат:

```bash
./dist/matrix-check/matrix-check
```

При переносе на другой ПК копируйте **всю папку** `dist/matrix-check/`.

## 6. Типичные ошибки

### `ModuleNotFoundError: No module named '_tkinter'`

```bash
sudo apt install python3-tk tcl tk
```

### `Tcl data directory not found` (в собранном бинарнике)

Переустановите `python3-tk` и пересоберите:

```bash
sudo apt install --reinstall python3-tk
./build_linux_pyinstaller.sh
```

### `ModuleNotFoundError: No module named 'pandas'`

```bash
poetry install --extras build --no-root
```

### Нет прав на запуск скриптов

```bash
chmod +x run.sh build_linux_pyinstaller.sh
```

## 7. Ручная сборка

```bash
poetry run python -m pip install -U pyinstaller
rm -rf build dist
poetry run python -m PyInstaller --clean --noconfirm matrix-check-linux.spec
```

## 8. Тесты

```bash
poetry run pytest -q
```
