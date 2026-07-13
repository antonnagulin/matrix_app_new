import pandas as pd
from datetime import datetime

from src.filters import normalize_fio, is_active, is_dismissed


class TestNormalizeFio:
    def test_full_fio(self):
        assert normalize_fio("Иванов Иван Иванович") == "ИВАНОВ ИИ"

    def test_fio_with_dots(self):
        assert normalize_fio("Петров П. П.") == "ПЕТРОВ ПП"

    def test_yo_normalized(self):
        assert normalize_fio("Ёлкин Е. Е.") == "ЕЛКИН ЕЕ"
        assert normalize_fio("ёжиков ё. ё.") == "ЕЖИКОВ ЕЕ"

    def test_two_words(self):
        assert normalize_fio("Сидоров Сергей") == "СИДОРОВ СЕРГЕЙ"

    def test_single_word(self):
        assert normalize_fio("Иванов") == "ИВАНОВ"

    def test_none_and_nan(self):
        assert normalize_fio(None) == ""
        assert normalize_fio(float("nan")) == ""

    def test_empty_and_whitespace(self):
        assert normalize_fio("") == ""
        assert normalize_fio("   ") == ""

    def test_strips_extra_spaces(self):
        assert normalize_fio("  Иванов   Иван   Иванович  ") == "ИВАНОВ ИИ"

    def test_numeric_input(self):
        assert normalize_fio(123) == "123"


class TestIsActive:
    def test_active_date_string_dot_format(self):
        assert is_active("31.12.9999") is True

    def test_inactive_date_string(self):
        assert is_active("01.01.2025") is False

    def test_iso_dash_format(self):
        assert is_active("9999-12-31") is True
        assert is_active("2020-01-15") is False

    def test_timestamp(self):
        assert is_active(pd.Timestamp("9999-12-31")) is True
        assert is_active(pd.Timestamp("2020-01-01")) is False

    def test_datetime(self):
        assert is_active(datetime(9999, 12, 31)) is True
        assert is_active(datetime(2024, 5, 1)) is False

    def test_none_and_nan(self):
        assert is_active(None) is False
        assert is_active(float("nan")) is False

    def test_invalid_string(self):
        assert is_active("не дата") is False
        assert is_active("") is False

    def test_unsupported_type(self):
        assert is_active(12345) is False


class TestIsDismissed:
    def test_active_employee(self):
        record = {
            "Штатная должность": "Менеджер",
            "Статус занятости": "работает",
            "Увольн.": "31.12.9999",
        }
        assert is_dismissed(record) is False

    def test_dismissed_by_default_position(self):
        record = {
            "Штатная должность": "Интеграция: штатная должность по умолч.",
            "Статус занятости": "работает",
            "Увольн.": "31.12.9999",
        }
        assert is_dismissed(record) is True

    def test_dismissed_by_status(self):
        record = {
            "Штатная должность": "Менеджер",
            "Статус занятости": "Уволился",
            "Увольн.": "31.12.9999",
        }
        assert is_dismissed(record) is True

    def test_dismissed_by_date(self):
        record = {
            "Штатная должность": "Менеджер",
            "Статус занятости": "работает",
            "Увольн.": "01.01.2025",
        }
        assert is_dismissed(record) is True

    def test_empty_dismiss_date_means_dismissed(self):
        record = {
            "Штатная должность": "Менеджер",
            "Статус занятости": "работает",
            "Увольн.": "",
        }
        assert is_dismissed(record) is True

    def test_missing_keys_treated_as_dismissed(self):
        assert is_dismissed({}) is True
