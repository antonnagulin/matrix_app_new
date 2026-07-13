"""Тесты file_utils — чтение/запись Excel и заливки."""

import pandas as pd
from openpyxl import Workbook, load_workbook

from src.file_utils import read_sprav_file, load_check_file, save_file, get_fill
from src.constants import RED_FILL_COLOR


def test_read_sprav_fills_na_with_empty_string(tmp_path):
    path = tmp_path / "sprav.xlsx"
    pd.DataFrame({"ФИО": ["А", None], "Увольн.": [None, "01.01.2020"]}).to_excel(
        path, index=False
    )

    df = read_sprav_file(path)

    assert df.loc[1, "ФИО"] == ""
    assert df.loc[0, "Увольн."] == ""


def test_load_and_save_roundtrip(tmp_path):
    src = tmp_path / "in.xlsx"
    dst = tmp_path / "out.xlsx"

    wb = Workbook()
    wb.active["A1"] = "ФИО"
    wb.active["A2"] = "Тестов Т. Т."
    wb.save(src)

    loaded_wb, loaded_ws = load_check_file(src)
    assert loaded_ws["A2"].value == "Тестов Т. Т."

    save_file(loaded_wb, dst)
    assert load_workbook(dst).active["A2"].value == "Тестов Т. Т."


def test_get_fill_returns_solid_pattern():
    fill = get_fill(RED_FILL_COLOR)
    assert fill.fill_type == "solid"
    assert fill.start_color.rgb.endswith(RED_FILL_COLOR)
