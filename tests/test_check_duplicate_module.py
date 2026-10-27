"""Тесты src/check_duplicate.py."""

from src.check_duplicate import check_duplicate, second_check_duplicate
from src.constants import ACTIVE_DATE

KEY = (2, "ИВАНОВ ИИ", "Иванов И. И.", "IT")


def _active(**extra):
    base = {
        "ФИО": "Иванов Иван Иванович",
        "Штатная должность": "Инженер",
        "Статус занятости": "работает",
        "Увольн.": ACTIVE_DATE,
    }
    base.update(extra)
    return base


def _dismissed(**extra):
    return _active(
        **{
            "Штатная должность": "Интеграция: штатная должность по умолч.",
            "Статус занятости": "уволился",
            "Увольн.": "01.01.2020",
            **extra,
        }
    )


def test_check_duplicate_single_active_resolves():
    data = {KEY: [_active()]}
    new_d, not_d, dismissed, joints = check_duplicate(data)
    assert new_d == {}
    assert len(not_d) == 1
    assert dismissed == {}
    assert joints == {}


def test_check_duplicate_two_active_stays_ambiguous():
    data = {KEY: [_active(), _active(**{"Штатная должность": "Менеджер"})]}
    new_d, not_d, dismissed, joints = check_duplicate(data)
    assert len(new_d) == 1
    assert not_d == {}
    assert dismissed == {}
    assert joints == {}


def test_check_duplicate_all_dismissed():
    data = {KEY: [_dismissed(), _dismissed(**{"Увольн.": "15.03.2021"})]}
    new_d, not_d, dismissed, joints = check_duplicate(data)
    assert new_d == {}
    assert not_d == {}
    assert len(dismissed) == 1
    assert joints == {}


def test_check_duplicate_skips_internal_part_time():
    data = {
        KEY: [
            _active(),
            _active(**{"Группа сотрудников": "Внутренние совмест."}),
        ]
    }
    new_d, not_d, dismissed, joints = check_duplicate(data)
    assert new_d == {}
    assert len(not_d) == 1
    assert joints == {}


def test_check_duplicate_joints_only_not_labeled_dismissed():
    data = {
        KEY: [
            _active(**{"Группа сотрудников": "Внутренние совмест."}),
            _active(
                **{
                    "Группа сотрудников": "Внутренние совмест.",
                    "Штатная должность": "Менеджер",
                }
            ),
        ]
    }
    new_d, not_d, dismissed, joints = check_duplicate(data)
    assert new_d == {}
    assert not_d == {}
    assert dismissed == {}
    assert len(joints) == 1


def test_check_duplicate_active_plus_dismissed_resolves():
    data = {KEY: [_active(), _dismissed()]}
    new_d, not_d, dismissed, joints = check_duplicate(data)
    assert new_d == {}
    assert len(not_d) == 1
    assert dismissed == {}
    assert joints == {}


def test_second_check_duplicate_full_fio_single_match():
    key = (3, "ИВАНОВ ИИ", "Иванов Иван Иванович", "IT")
    rec = _active()
    new_d, not_d = second_check_duplicate({key: [rec]})
    assert new_d == {}
    assert len(not_d) == 1


def test_second_check_duplicate_full_fio_two_identical_records_still_ambiguous():
    key = (3, "ИВАНОВ ИИ", "Иванов Иван Иванович", "IT")
    rec = _active()
    new_d, not_d = second_check_duplicate({key: [rec, rec]})
    assert len(new_d) == 1
    assert not_d == {}


def test_second_check_duplicate_full_fio_multiple_matches():
    key = (3, "ИВАНОВ ИИ", "Иванов Иван Иванович", "IT")
    rec1 = _active()
    rec2 = _active(**{"Штатная должность": "Менеджер"})
    new_d, not_d = second_check_duplicate({key: [rec1, rec2]})
    assert len(new_d) == 1
    assert not_d == {}


def test_second_check_duplicate_initials_go_to_manual_selection():
    """Инициалы вроде «Смирнов А. Н.» не совпадают с полным ФИО в справочнике → ручной выбор."""
    key = (3, "СМИРНОВ АН", "Смирнов А. Н.", "IT")
    rec1 = _active(**{"ФИО": "Смирнов Алексей Николаевич", "Штатная должность": "Инженер"})
    rec2 = _active(**{"ФИО": "Смирнов Алексей Николаевич", "Штатная должность": "Программист"})
    remaining, resolved = second_check_duplicate({key: [rec1, rec2]})
    assert len(remaining) == 1
    assert resolved == {}
    assert len(remaining[key]) == 2


def test_second_check_duplicate_short_initials_without_dots_go_to_manual():
    key = (3, "ИВАНОВ ИИ", "Иванов И И", "IT")
    remaining, resolved = second_check_duplicate({key: [_active()]})
    assert len(remaining) == 1
    assert resolved == {}
