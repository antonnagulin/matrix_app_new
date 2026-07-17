"""Тесты src/check_dissmised.py."""

from src.check_dissmised import check_dissmised


def _row(dismiss_date: str):
    return [{"Увольн.": dismiss_date}]


def test_filters_recent_dismissals_2025_2026():
    data = {
        (1, "A", "A", "x"): _row("2025-06-01"),
        (2, "B", "B", "x"): _row("2026-01-15"),
        (3, "C", "C", "x"): _row("2020-01-01"),
    }
    result = check_dissmised(data)
    assert len(result) == 2
    assert (1, "A", "A", "x") in result
    assert (2, "B", "B", "x") in result
    assert (3, "C", "C", "x") not in result
