from src.filters import is_dismissed


def test_dismissed_by_position():
    row = {
        "Штатная должность": "Интеграция: штатная должность по умолч.",
        "Статус занятости": "",
        "Увольн.": "31.12.9999",
    }
    assert is_dismissed(row) is True


def test_dismissed_by_status():
    row = {
        "Штатная должность": "Инженер",
        "Статус занятости": "уволился",
        "Увольн.": "31.12.9999",
    }
    assert is_dismissed(row) is True


def test_dismissed_by_date():
    row = {
        "Штатная должность": "Инженер",
        "Статус занятости": "работает",
        "Увольн.": "11.12.2004",
    }
    assert is_dismissed(row) is True


def test_not_dismissed_when_active():
    row = {
        "Штатная должность": "Инженер",
        "Статус занятости": "работает",
        "Увольн.": "31.12.9999",
    }
    assert is_dismissed(row) is False
