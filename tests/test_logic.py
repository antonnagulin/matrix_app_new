import pytest
import pandas as pd
from src.logic import run_check
from pathlib import Path

def test_run_check(tmp_path):
    # создаём справочник
    spr_df = pd.DataFrame({
        "ФИО": ["Иванов Иван Иванович", "Петров Пётр Петрович"],
        "Штатная должность": ["Менеджер", "Интеграция: штатная должность по умолч."],
        "Статус занятости": ["работает", "работает"],
        "Увольн.": ["31.12.9999", "31.12.9999"]
    })
    spr_file = tmp_path / "sprav.xlsx"
    spr_df.to_excel(spr_file, index=False)

    # создаём файл проверки
    check_df = pd.DataFrame({
        "ФИО": ["Иванов И. И.", "Петров П. П."],
        "Должность": ["Менеджер", "Менеджер"],
        "Подразделение": ["Отдел1", "Отдел2"]
    })
    check_file = tmp_path / "check.xlsx"
    check_df.to_excel(check_file, index=False)

    output_file = tmp_path / "output.xlsx"

    # запускаем проверку
    run_check(
        sprav_file=spr_file,
        check_file=check_file,
        output_file=output_file,
        check_dismissed=True,
        check_duplicates=True,
        log=None
    )

    # Проверяем, что файл создан
    assert Path(output_file).exists()
