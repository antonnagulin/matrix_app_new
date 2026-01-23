from src.filters import is_dismissed


def test_dismissed_by_position():
    row = {
        "Штатная должность": "Интеграция: штатная должность по умолч.",
        "Статус занятости": "",
        "Увольн.": ""
    }
    assert is_dismissed(row) is True


def test_dismissed_by_status():
    row = {
        "Штатная должность": "",
        "Статус занятости": "уволился",
        "Увольн.": ""
    }
    assert is_dismissed(row) is True


def test_dismissed_by_date():
    row = {
        "Штатная должность": "",
        "Статус занятости": "",
        "Увольн.": "11.12.2004"
    }
    assert is_dismissed(row) is True


