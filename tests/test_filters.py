import pytest
import pandas as pd
from datetime import datetime
from src.filters import normalize_fio, is_active, is_dismissed

def test_normalize_fio():
    # Полное ФИО
    assert normalize_fio("Иванов Иван Иванович") == "ИВАНОВ ИИ"
    # ФИО с точками
    assert normalize_fio("Петров П. П.") == "ПЕТРОВ ПП"
    # ФИО с Ё
    assert normalize_fio("Ёлкин Е. Е.") == "ЕЛКИН ЕЕ"
    # Два слова
    assert normalize_fio("Сидоров Сергей") == "СИДОРОВ СЕРГЕЙ"
    # Пустое
    assert normalize_fio(None) == ""

def test_is_active():
    # Строка в формате дд.мм.гггг
    assert is_active("31.12.9999") is True
    assert is_active("01.01.2025") is False
    # pd.Timestamp
    assert is_active(pd.Timestamp("9999-12-31")) is True
    # datetime
    assert is_active(datetime(9999,12,31)) is True
    # None
    assert is_active(None) is False

def test_is_dismissed():
    # Сотрудник активен
    record_active = {
        "Штатная должность": "Менеджер",
        "Статус занятости": "работает",
        "Увольн.": "31.12.9999"
    }
    assert is_dismissed(record_active) is False

    # Сотрудник уволен по позиции
    record_position = {
        "Штатная должность": "Интеграция: штатная должность по умолч.",
        "Статус занятости": "работает",
        "Увольн.": "31.12.9999"
    }
    assert is_dismissed(record_position) is True

    # Сотрудник уволен по статусу
    record_status = {
        "Штатная должность": "Менеджер",
        "Статус занятости": "уволился",
        "Увольн.": "31.12.9999"
    }
    assert is_dismissed(record_status) is True

    # Сотрудник уволен по дате
    record_date = {
        "Штатная должность": "Менеджер",
        "Статус занятости": "работает",
        "Увольн.": "01.01.2025"
    }
    assert is_dismissed(record_date) is True
