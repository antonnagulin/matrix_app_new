from src.constants import (
    FIO_COL,
    POSITION_COL,
    STATUS_COL,
    DISMISS_COL,
    ACTIVE_DATE,
    RED_FILL_COLOR,
    ORANGE_FILL_COLOR,
)


def test_column_constants():
    assert FIO_COL == "ФИО"
    assert POSITION_COL == "Штатная должность"
    assert STATUS_COL == "Статус занятости"
    assert DISMISS_COL == "Увольн."


def test_active_date_constant():
    assert ACTIVE_DATE == "31.12.9999"


def test_fill_color_constants():
    assert RED_FILL_COLOR == "FF9999"
    assert ORANGE_FILL_COLOR == "FF7500"
