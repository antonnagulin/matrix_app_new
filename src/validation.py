"""Ранняя валидация путей и схемы справочника."""

from __future__ import annotations

from pathlib import Path

import pandas as pd

from .constants import DISMISS_COL, FIO_COL, POSITION_COL, STATUS_COL

# Минимально необходимые колонки справочника
SPRAV_REQUIRED_COLUMNS = (FIO_COL,)
# Желательные колонки (предупреждение, если нет)
SPRAV_RECOMMENDED_COLUMNS = (POSITION_COL, STATUS_COL, DISMISS_COL)


def _norm_path(path: str | Path) -> Path:
    return Path(path).expanduser().resolve()


def validate_output_not_input(
    output_file: str | Path,
    *,
    sprav_file: str | Path | None = None,
    check_file: str | Path | None = None,
    renamed_fio_file: str | Path | None = None,
) -> None:
    """Запретить запись результата поверх входных файлов."""
    out = _norm_path(output_file)
    inputs: list[tuple[str, str | Path]] = []
    if sprav_file:
        inputs.append(("справочника", sprav_file))
    if check_file:
        inputs.append(("проверки", check_file))
    if renamed_fio_file:
        inputs.append(("смены ФИО", renamed_fio_file))

    for label, src in inputs:
        try:
            src_path = _norm_path(src)
        except OSError:
            continue
        if src_path == out:
            raise ValueError(
                f"Файл результата совпадает с файлом {label}:\n{out}\n"
                "Укажите другое имя выходного файла."
            )
        # Также review / included рядом — stem-коллизии не проверяем жёстко;
        # достаточно не перезаписать исходники основным out.
        review = out.with_name(f"{out.stem}_review{out.suffix}")
        included = out.with_name(f"{out.stem}_included{out.suffix}")
        for side in (review, included):
            if src_path == side:
                raise ValueError(
                    f"Служебный файл результата совпадает с файлом {label}:\n{side}\n"
                    "Укажите другое имя выходного файла."
                )


def validate_sprav_schema(spr_df: pd.DataFrame) -> list[str]:
    """
    Проверить схему справочника.

    Raises:
        ValueError: нет обязательных колонок.
    Returns:
        список предупреждений (отсутствующие рекомендуемые колонки).
    """
    cols = {str(c).strip() for c in spr_df.columns}
    missing_required = [c for c in SPRAV_REQUIRED_COLUMNS if c not in cols]
    if missing_required:
        raise ValueError(
            "В справочнике нет обязательных колонок: "
            + ", ".join(f"«{c}»" for c in missing_required)
            + ".\nПроверьте файл и строку заголовков."
        )
    warnings = [c for c in SPRAV_RECOMMENDED_COLUMNS if c not in cols]
    return warnings


def validate_sheet_name(sheet_names: list[str], sheet_name: str | None) -> None:
    """Если лист задан явно — он должен существовать."""
    if not sheet_name:
        return
    if sheet_name not in sheet_names:
        available = ", ".join(sheet_names) if sheet_names else "(пусто)"
        raise ValueError(
            f"Лист «{sheet_name}» не найден в файле проверки.\n"
            f"Доступные листы: {available}"
        )
