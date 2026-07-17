"""Тесты src/write_position_eararchi.py."""

from openpyxl import Workbook

from src.write_position_eararchi import write_position_eararchi


def test_write_position_and_hierarchy():
    wb = Workbook()
    ws = wb.active
    key = (5, "ИВАНОВ ИИ", "Иванов Иван Иванович", "IT")
    record = {
        "ФИО": "Иванов Иван Иванович",
        "Штатная должность": "Инженер",
        "Орг. единица уровня 1": "Головной офис",
        "Орг. единица уровня 2": "IT",
        "Орг. единица уровня 3": "",
        "Орг. единица уровня 4": "",
    }
    not_d = {key: [record]}

    write_position_eararchi(ws, not_d, new_col_hierarchy_index=3, new_col_position_index=4)

    assert ws.cell(row=5, column=3).value == "Головной офис / IT"
    assert ws.cell(row=5, column=4).value == "Инженер"
