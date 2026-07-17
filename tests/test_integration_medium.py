"""Интеграционный тест на sample_data/check_medium.xlsx."""

from pathlib import Path

import pytest
from openpyxl import load_workbook

from src.logic import run_check
from src.constants import RED_FILL_COLOR, ELOYY_FILL_COLOR, GREEN_FILL_COLOR

ROOT = Path(__file__).resolve().parents[1]
SPRAV = ROOT / "sample_data" / "sprav_medium.xlsx"
CHECK = ROOT / "sample_data" / "check_medium.xlsx"


@pytest.mark.skipif(not SPRAV.is_file() or not CHECK.is_file(), reason="medium sample missing")
def test_medium_sample_produces_expected_highlights(tmp_path):
    output = tmp_path / "medium_result.xlsx"
    log_lines = []

    class Log:
        def insert(self, _idx, text):
            log_lines.append(text)

        def see(self, _idx):
            pass

    run_check(
        str(SPRAV),
        str(CHECK),
        str(output),
        check_dismissed=True,
        check_duplicates=True,
        prompt_duplicate_selection=False,
        log=Log(),
    )

    assert output.exists()
    ws = load_workbook(output).active

    def _has_color(row: int, color: str) -> bool:
        return color in (ws.cell(row=row, column=1).fill.start_color.rgb or "")

    assert _has_color(17, GREEN_FILL_COLOR)
    assert _has_color(18, GREEN_FILL_COLOR)
    assert _has_color(19, ELOYY_FILL_COLOR)
    assert _has_color(20, ELOYY_FILL_COLOR)
    assert _has_color(27, RED_FILL_COLOR)

    log_text = "".join(log_lines)
    assert "Строка заголовков" in log_text
    assert "Разбор дубликатов" in log_text
