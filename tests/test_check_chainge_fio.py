"""Тесты src/check_chainge_fio.py."""

import pandas as pd

from src.check_chainge_fio import chainge_fio


def test_chainge_fio_no_match_silent(tmp_path, capsys):
    path = tmp_path / "renamed.xlsx"
    pd.DataFrame(
        {
            "Смена ФИО": ["Иванов Иван Иванович"],
            "Табельный номер": ["3001"],
        }
    ).to_excel(path, index=False)

    chainge_fio(["Неизвестный Н. Н."], str(path))

    assert capsys.readouterr().out == ""


def test_chainge_fio_reports_ambiguous_renames(tmp_path, capsys):
    path = tmp_path / "renamed.xlsx"
    pd.DataFrame(
        {
            "Смена ФИО": [
                "Федоров Федор Федорович",
                "Федоров Ф. Ф.",
            ],
            "Табельный номер": ["3001", "3002"],
        }
    ).to_excel(path, index=False)

    chainge_fio(["Фёдоров Ф. Ф."], str(path))

    captured = capsys.readouterr()
    assert "расхождение" in captured.out.lower() or "Найдено" in captured.out
