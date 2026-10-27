"""Тесты ранней валидации путей и схемы справочника."""

from pathlib import Path

import pandas as pd
import pytest
from openpyxl import Workbook

from src.logic import run_check
from src.validation import (
    validate_output_not_input,
    validate_sheet_name,
    validate_sprav_schema,
)


def test_validate_sprav_schema_requires_fio():
    with pytest.raises(ValueError, match="обязательных"):
        validate_sprav_schema(pd.DataFrame({"Имя": ["A"]}))


def test_validate_sprav_schema_warns_on_recommended():
    warnings = validate_sprav_schema(pd.DataFrame({"ФИО": ["A"]}))
    assert "Штатная должность" in warnings
    assert "Статус занятости" in warnings
    assert "Увольн." in warnings


def test_validate_output_not_input(tmp_path):
    sprav = tmp_path / "sprav.xlsx"
    sprav.write_bytes(b"x")
    with pytest.raises(ValueError, match="совпадает"):
        validate_output_not_input(sprav, sprav_file=sprav, check_file=tmp_path / "c.xlsx")


def test_validate_sheet_name():
    validate_sheet_name(["Лист1", "Данные"], None)
    validate_sheet_name(["Лист1", "Данные"], "Данные")
    with pytest.raises(ValueError, match="не найден"):
        validate_sheet_name(["Лист1"], "НетТакого")


def test_run_check_rejects_output_equal_sprav(monkeypatch, tmp_path):
    sprav = tmp_path / "sprav.xlsx"
    check = tmp_path / "check.xlsx"
    sprav.write_bytes(b"x")
    check.write_bytes(b"x")

    monkeypatch.setattr(
        "src.logic.read_sprav_file",
        lambda *_: pd.DataFrame(
            [
                {
                    "ФИО": "Иванов Иван Иванович",
                    "Штатная должность": "Инженер",
                    "Статус занятости": "работает",
                    "Увольн.": "31.12.9999",
                }
            ]
        ),
    )
    with pytest.raises(ValueError, match="совпадает"):
        run_check(str(sprav), str(check), str(sprav))


def test_run_check_missing_fio_column_in_sprav(monkeypatch):
    wb = Workbook()
    ws = wb.active
    ws.append(["ФИО", "Отдел"])
    ws.append(["Иванов И. И.", "IT"])
    monkeypatch.setattr(
        "src.logic.read_sprav_file",
        lambda *_: pd.DataFrame({"Должность": ["x"]}),
    )
    monkeypatch.setattr("src.logic.load_check_file", lambda *a, **k: (wb, ws, None))
    monkeypatch.setattr("src.logic.save_file", lambda *_: None)

    with pytest.raises(ValueError, match="обязательных"):
        run_check("sprav.xlsx", "check.xlsx", "out.xlsx")


def test_load_check_file_unknown_sheet(tmp_path):
    from src.file_utils import load_check_file

    path = tmp_path / "check.xlsx"
    wb = Workbook()
    wb.active.title = "Основной"
    wb.active["A1"] = "ФИО"
    wb.save(path)

    with pytest.raises(ValueError, match="не найден"):
        load_check_file(path, sheet_name="Другой")
