"""Запись подразделения и должности из справочника в лист проверки."""

from .filters import join_with_slash


def write_position_hierarchy(ws, resolved, new_col_hierarchy_index, new_col_position_index):
    """
    ws                       — worksheet openpyxl
    resolved                 — словарь {(row, fio_norm, fio, pos): [dict, ...], ...}
    new_col_hierarchy_index  — колонка для иерархии
    new_col_position_index   — колонка для штатной должности
    """
    for (row, _fio_norm, _fio_value, _pos), value in resolved.items():
        hierarchy_str = join_with_slash(value[0])
        position_str = value[0].get("Штатная должность", "")
        ws.cell(row=row, column=new_col_hierarchy_index, value=hierarchy_str)
        ws.cell(row=row, column=new_col_position_index, value=position_str)


# Совместимость со старым именем
write_position_eararchi = write_position_hierarchy
