import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

from .constants import FIO_COL


def read_sprav_file(file_path: str) -> pd:
    """Чтение справочника в DataFrame"""

    return pd.read_excel(file_path, dtype=str).fillna("")


def load_check_file(file_path) -> load_workbook:
    """Открытие файла для проверки Excel"""

    wb = load_workbook(file_path)
    ws = wb.active
    return wb, ws


def save_file(wb, file_path) -> None:
    """Сохранить файл Excel"""

    wb.save(file_path)


def get_fill(color_hex: str) -> PatternFill:
    """Получить цвет для файл Excel """

    return PatternFill(start_color=color_hex, end_color=color_hex, fill_type="solid")


def clear_fill() -> PatternFill:
    """Сбросить заливку ячейки."""
    return PatternFill(fill_type=None)


def _cell_text(value) -> str:
    if value is None:
        return ""
    return str(value).strip()


def row_values(ws, row: int, max_col: int | None = None) -> list:
    """Значения ячеек строки (1-based)."""
    last_col = max_col or ws.max_column or 1
    return [ws.cell(row=row, column=c).value for c in range(1, last_col + 1)]


def find_header_row(
    ws,
    marker: str = FIO_COL,
    *,
    max_scan: int | None = None,
) -> int:
    """
    Найти номер строки (1-based) с заголовками по маркеру (обычно «ФИО»).

    Raises:
        ValueError: если маркер не найден в первых max_scan строках.
    """
    limit = max_scan if max_scan is not None else max(ws.max_row or 1, 1)
    marker_norm = marker.strip().casefold()

    for row in range(1, limit + 1):
        for col in range(1, (ws.max_column or 1) + 1):
            text = _cell_text(ws.cell(row=row, column=col).value)
            if text.casefold() == marker_norm:
                return row

    raise ValueError(
        f"Не найдена строка заголовков с колонкой «{marker}» "
        f"(проверены строки 1–{limit})."
    )


def peek_sheet_rows(file_path: str, max_rows: int = 40) -> tuple[list[list], int]:
    """
    Прочитать первые строки файла проверки для превью в UI.

    Returns:
        (rows, detected_header_row) — rows: список [номер_строки, *значения],
        detected_header_row или 0 если не найден.
    """
    wb, ws = load_check_file(file_path)
    try:
        detected = find_header_row(ws)
    except ValueError:
        detected = 0

    last = min(max_rows, ws.max_row or 1)
    max_col = min(ws.max_column or 1, 12)
    rows = []
    for r in range(1, last + 1):
        values = [_cell_text(v) for v in row_values(ws, r, max_col)]
        rows.append([r, *values])
    return rows, detected

