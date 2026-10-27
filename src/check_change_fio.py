"""Проверка и авторазрешение отсутствующих ФИО по файлу смены фамилий."""

from __future__ import annotations

import pandas as pd

from .file_utils import read_sprav_file
from .filters import normalize_fio
from .fio_index import FioIndex, build_fio_index, lookup_fio, person_group_key

_NAME_COLUMNS = ("Смена ФИО", "Старое ФИО", "Новое ФИО", "ФИО")
_TAB_COLUMNS = ("Таб.№", "Табельный номер")


def _logger(write):
    def _log(msg: str = "") -> None:
        if write:
            write(msg)
        else:
            print(msg)

    return _log


def _name_columns(df: pd.DataFrame) -> list[str]:
    return [c for c in _NAME_COLUMNS if c in df.columns]


def _tab_column(df: pd.DataFrame) -> str | None:
    for col in _TAB_COLUMNS:
        if col in df.columns:
            return col
    return None


def _build_tab_index(spr_df: pd.DataFrame) -> dict[str, list]:
    """Табельный номер → список записей справочника."""
    tab_col = _tab_column(spr_df)
    if not tab_col:
        return {}
    index: dict[str, list] = {}
    for row in spr_df.to_dict("records"):
        tab = str(row.get(tab_col, "") or "").strip()
        if tab:
            index.setdefault(tab, []).append(row)
    return index


def prepare_rename_df(change_fio_file: str) -> pd.DataFrame | None:
    """Загрузить файл смены ФИО и подготовить нормализованные колонки."""
    rename_df = read_sprav_file(change_fio_file)
    name_cols = _name_columns(rename_df)
    if not name_cols:
        return None
    for col in name_cols:
        rename_df[f"_NORM_{col}"] = rename_df[col].map(normalize_fio)
    return rename_df


def build_tab_index(spr_df: pd.DataFrame) -> dict[str, list]:
    return _build_tab_index(spr_df)


def resolve_via_rename(
    fio_value: str,
    rename_df: pd.DataFrame,
    fio_index: FioIndex | dict | None = None,
    tab_index: dict[str, list] | None = None,
) -> tuple[list | None, str | None]:
    """
    Найти запись(и) справочника для отсутствующего ФИО через файл смены.

    fio_index — предпочтительно FioIndex; dict оставлен для совместимости тестов.
    """
    name_cols = _name_columns(rename_df)
    if not name_cols:
        return None, None

    item_norm = normalize_fio(fio_value)
    if not item_norm:
        return None, None

    mask = False
    for col in name_cols:
        norm_col = f"_NORM_{col}"
        if norm_col not in rename_df.columns:
            rename_df[norm_col] = rename_df[col].map(normalize_fio)
        mask = mask | (rename_df[norm_col] == item_norm)
    matched = rename_df.loc[mask]
    if matched.empty:
        return None, None

    candidate_names: list[str] = []
    for col in name_cols:
        for val in matched[col].tolist():
            if val is None or str(val).strip() == "":
                continue
            if normalize_fio(val) == item_norm:
                continue
            candidate_names.append(str(val))

    found_by_fio: dict[str, list] = {}
    for name in candidate_names:
        if isinstance(fio_index, FioIndex):
            records = lookup_fio(fio_index, name)
        elif isinstance(fio_index, dict):
            records = list(fio_index.get(normalize_fio(name), []))
        else:
            records = []
        if records:
            key = person_group_key(records[0].get("ФИО", name))
            found_by_fio[key] = records

    found_by_tab: dict[str, list] = {}
    tab_col = _tab_column(matched)
    if tab_col and tab_index:
        for tab in matched[tab_col].tolist():
            tab_key = str(tab).strip()
            if tab_key and tab_key in tab_index:
                found_by_tab[tab_key] = tab_index[tab_key]

    resolved: dict[str, list] = {}
    for records in list(found_by_fio.values()) + list(found_by_tab.values()):
        key = person_group_key(records[0].get("ФИО", ""))
        if key:
            resolved[key] = records
        else:
            resolved[f"tab:{id(records)}"] = records

    if len(resolved) == 1:
        records = next(iter(resolved.values()))
        via = "табельному номеру" if found_by_tab and not found_by_fio else "ФИО"
        return list(records), via

    return None, None


def change_fio(
    red_not_found,
    change_fio_file=None,
    write=None,
    fio_dict=None,
    fio_index=None,
    spr_df=None,
    rename_df=None,
    tab_index=None,
    *,
    chainge_fio_file=None,
):
    """
    Лог по отсутствующим ФИО относительно файла смены фамилий.

    Если rename_df уже загружен (первый проход run_check) — файл не перечитывается.
    """
    _log = _logger(write)
    resolved_map: dict[str, list] = {}
    file_path = change_fio_file or chainge_fio_file

    if rename_df is None:
        if not file_path:
            _log("Файл смены ФИО не задан — проверка пропущена.")
            return resolved_map
        rename_df = prepare_rename_df(file_path)
    if rename_df is None:
        _log(
            "В файле смены ФИО нет колонок «Смена ФИО» / «Старое ФИО» / «Новое ФИО» — "
            "проверка пропущена."
        )
        return resolved_map

    name_cols = _name_columns(rename_df)
    if tab_index is None:
        tab_index = _build_tab_index(spr_df) if spr_df is not None else {}

    if fio_index is None and spr_df is not None:
        fio_index = build_fio_index(spr_df)
    elif fio_index is None and fio_dict is not None:
        fio_index = fio_dict

    found_any = False
    for item in red_not_found:
        fio = item[1] if isinstance(item, (tuple, list)) and len(item) >= 2 else item
        fio_str = str(fio)

        records, via = resolve_via_rename(fio_str, rename_df, fio_index, tab_index)
        if not records:
            item_norm = normalize_fio(fio_str)
            mask = False
            for col in name_cols:
                mask = mask | (rename_df[f"_NORM_{col}"] == item_norm)
            matched_rows = rename_df.loc[mask]
            if matched_rows.empty:
                continue
            found_any = True
            preview_cols = [
                c
                for c in (
                    *_TAB_COLUMNS,
                    "Старое ФИО",
                    "Новое ФИО",
                    "Смена ФИО",
                    "ФИО",
                    "Дата изменения",
                    "Дата",
                    "Примечание",
                    "Комментарий",
                )
                if c in matched_rows.columns
            ]
            preview = (
                matched_rows[preview_cols]
                if preview_cols
                else matched_rows.drop(
                    columns=[c for c in matched_rows.columns if c.startswith("_NORM_")],
                    errors="ignore",
                )
            )
            _log(
                f"\nВ файле смены ФИО есть запись для «{fio_str}», "
                f"но в справочнике однозначно не найдено:"
            )
            _log(preview.head(10).to_string(index=False))
            continue

        found_any = True
        resolved_map[fio_str] = records
        sprav_fio = records[0].get("ФИО", "")
        _log(
            f"\n«{fio_str}» найден через смену ФИО ({via}) → «{sprav_fio}» "
            f"({len(records)} запис. в справочнике)"
        )

    if not found_any:
        _log("Совпадений в файле смены ФИО не найдено.")

    return resolved_map


# Совместимость со старыми именами
chainge_fio = change_fio
