"""Устаревший модуль: используйте ``src.check_change_fio``."""

from .check_change_fio import (
    build_tab_index,
    chainge_fio,
    change_fio,
    prepare_rename_df,
    resolve_via_rename,
)

__all__ = [
    "build_tab_index",
    "change_fio",
    "chainge_fio",
    "prepare_rename_df",
    "resolve_via_rename",
]
