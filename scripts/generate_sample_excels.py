"""Генерация демо Excel-файлов без внешних зависимостей (stdlib only)."""

from __future__ import annotations

import zipfile
from pathlib import Path
from xml.sax.saxutils import escape


ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "sample_data"


def col_letter(idx: int) -> str:
    """1-based column index -> Excel letter (A, B, ..., Z, AA, ...)."""
    result = []
    while idx:
        idx, rem = divmod(idx - 1, 26)
        result.append(chr(65 + rem))
    return "".join(reversed(result))


def sheet_xml(headers: list[str], rows: list[list[str]]) -> str:
    lines = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        "<sheetData>",
    ]

    def write_row(r_idx: int, values: list[str]) -> None:
        lines.append(f'<row r="{r_idx}">')
        for c_idx, value in enumerate(values, start=1):
            ref = f"{col_letter(c_idx)}{r_idx}"
            text = escape("" if value is None else str(value))
            lines.append(
                f'<c r="{ref}" t="inlineStr"><is><t>{text}</t></is></c>'
            )
        lines.append("</row>")

    write_row(1, headers)
    for i, row in enumerate(rows, start=2):
        write_row(i, row)

    lines.append("</sheetData></worksheet>")
    return "\n".join(lines)


def write_xlsx(path: Path, headers: list[str], rows: list[list[str]]) -> None:
    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>
"""
    rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
"""
    workbook = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>
"""
    workbook_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>
"""

    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("xl/workbook.xml", workbook)
        zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        zf.writestr("xl/worksheets/sheet1.xml", sheet_xml(headers, rows))


def main() -> None:
    sprav_headers = [
        "Таб.№",
        "ФИО",
        "Штатная должность",
        "Статус занятости",
        "Увольн.",
        "Подразделение",
    ]
    # Сценарии под текущую логику:
    # - активный сотрудник
    # - уволенный (дата / статус / должность по умолч.)
    # - дубликаты ФИО (для диалога выбора)
    # - ё/точки для normalize_fio
    sprav_rows = [
        ["1001", "Иванов Иван Иванович", "Инженер", "работает", "31.12.9999", "Отдел АСУ"],
        ["1002", "Петров Пётр Петрович", "Менеджер", "работает", "31.12.9999", "Отдел кадров"],
        # дубликаты уволенного Иванова (вторая запись) — для check_duplicates
        ["1003", "Сидоров Сидор Сидорович", "Аналитик", "работает", "31.12.9999", "Аналитика"],
        ["1003", "Сидоров Сидор Сидорович", "Интеграция: штатная должность по умолч.", "уволился", "15.03.2023", "Аналитика"],
        ["1004", "Кузнецова Анна Сергеевна", "Бухгалтер", "работает", "31.12.9999", "Бухгалтерия"],
        ["1005", "Смирнов Алексей Николаевич", "Интеграция: штатная должность по умолч.", "работает", "31.12.9999", "Склад"],
        ["1006", "Васильева Мария Ивановна", "Специалист", "уволился", "01.06.2024", "Отдел продаж"],
        ["1007", "Орлов Олег Олегович", "Техник", "работает", "10.01.2022", "Сервис"],
        ["1008", "Ёлкин Егор Егорович", "Программист", "работает", "31.12.9999", "IT"],
        ["1009", "Новикова Елена Павловна", "Юрист", "работает", "31.12.9999", "Юр. отдел"],
    ]

    check_headers = ["ФИО", "Должность", "Подразделение", "Таб.№"]
    # Форматы ФИО как в реальных выгрузках (полное / инициалы / ё)
    check_rows = [
        ["Иванов И. И.", "Инженер", "Отдел АСУ", "1001"],           # активный → без подсветки
        ["Петров П. П.", "Менеджер", "Отдел кадров", "1002"],       # активный (ё→е)
        ["Сидоров С. С.", "Аналитик", "Аналитика", "1003"],         # уволен + дубликаты → оранжевый
        ["Смирнов А. Н.", "Кладовщик", "Склад", "1005"],            # уволен по должности → оранжевый
        ["Васильева М. И.", "Специалист", "Отдел продаж", "1006"],  # уволен по статусу → оранжевый
        ["Орлов О. О.", "Техник", "Сервис", "1007"],                # уволен по дате → оранжевый
        ["Елкин Е. Е.", "Программист", "IT", "1008"],               # активный (ё нормализация)
        ["Кузнецова А. С.", "Бухгалтер", "Бухгалтерия", "1004"],    # активный
        ["Фёдоров Ф. Ф.", "Курьер", "Логистика", ""],               # нет в справочнике → красный
        ["Неизвестный Н. Н.", "Стажёр", "HR", ""],                  # нет в справочнике → красный
        ["", "Пустая строка", "", ""],                              # должна игнорироваться
        ["Новикова Е. П.", "Юрист", "Юр. отдел", "1009"],           # активный
    ]

    # Справочник смены ФИО: старое имя в проверяемом файле → актуальное в основном справочнике.
    # В текущем run_check пока не подключён; формат готов к интеграции.
    renamed_headers = [
        "Старое ФИО",
        "Новое ФИО",
        "Таб.№",
        "Дата изменения",
        "Примечание",
    ]
    renamed_rows = [
        ["Фёдоров Фёдор Фёдорович", "Федоров Федор Федорович", "1010", "12.01.2025", "Замена ё→е в документах"],
        ["Иванова Анна Сергеевна", "Кузнецова Анна Сергеевна", "1004", "03.09.2024", "Смена фамилии после замужества"],
        ["Смирнов А. Н.", "Смирнов Алексей Николаевич", "1005", "20.11.2023", "Раскрытие инициалов"],
        ["Петров Петр Петрович", "Петров Пётр Петрович", "1002", "01.02.2024", "Восстановление буквы ё"],
        ["Неизвестный Николай Николаевич", "Неизвестный Н. Н.", "1011", "05.05.2025", "Кандидат, ещё не в основном справочнике"],
    ]

    files = [
        (OUT / "sprav.xlsx", sprav_headers, sprav_rows),
        (OUT / "check.xlsx", check_headers, check_rows),
        (OUT / "sprav_renamed_fio.xlsx", renamed_headers, renamed_rows),
    ]

    for path, headers, rows in files:
        write_xlsx(path, headers, rows)
        print(f"OK: {path}")


if __name__ == "__main__":
    main()
