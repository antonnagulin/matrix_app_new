from src.logic import run_check
from openpyxl.styles import PatternFill
import tempfile
import os


class DummyLog:
    def __init__(self):
        self.lines = []

    def insert(self, *_args):
        self.lines.append(_args[1])

    def see(self, *_):
        pass


def test_run_check_red_and_orange(
    sprav_df,
    check_wb,
    monkeypatch
):
    wb, ws = check_wb
    log = DummyLog()

    # мок файловых функций
    monkeypatch.setattr("src.logic.read_sprav_file", lambda _: sprav_df)
    monkeypatch.setattr("src.logic.load_check_file", lambda _: (wb, ws))
    monkeypatch.setattr("src.logic.save_file", lambda *_: None)

    run_check(
        sprav_file="sprav.xlsx",
        check_file="check.xlsx",
        output_file="out.xlsx",
        log=log
    )

    red = PatternFill(start_color="FF9999", fill_type="solid")
    orange = PatternFill(start_color="FF7500", fill_type="solid")

    assert ws["A2"].fill.start_color.rgb is not None  # Иванов → оранжевый
    assert ws["A3"].fill.start_color.rgb is not None  # Петров → красный
