"""Проверка сборки словаря ФИО и обнаружения дубликатов."""

from src.filters import normalize_fio


def test_no_duplicates():
    data = [
        {"ФИО": "Иванов Иван Иванович"},
        {"ФИО": "Петров Петр Петрович"},
    ]

    fio_dict = {}
    for row in data:
        fio = normalize_fio(row["ФИО"])
        fio_dict.setdefault(fio, []).append(row)

    assert all(len(v) == 1 for v in fio_dict.values())


def test_detects_duplicates_by_normalized_fio():
    data = [
        {"ФИО": "Иванов Иван Иванович", "Таб.№": "1"},
        {"ФИО": "Иванов И. И.", "Таб.№": "2"},
        {"ФИО": "Петров П. П.", "Таб.№": "3"},
    ]

    fio_dict = {}
    for row in data:
        fio = normalize_fio(row["ФИО"])
        fio_dict.setdefault(fio, []).append(row)

    assert len(fio_dict["ИВАНОВ ИИ"]) == 2
    assert len(fio_dict["ПЕТРОВ ПП"]) == 1


def test_yo_and_dots_collapse_to_same_key():
    assert normalize_fio("Ёлкин Ё. Ё.") == normalize_fio("Елкин Е Е")
