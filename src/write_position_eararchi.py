from .filters import join_with_slash


def write_position_eararchi(ws, not_d, new_col_hierarchy_index, new_col_position_index):
    """
    ws                       — worksheet openpyxl
    not_d                    — словарь вида {(row, fio_norm): [dict, ...], ...}
    new_col_hierarchy_index  — колонка для иерархии
    new_col_position_index   — колонка для штатной должности
    """
        
    for (row, _fio_norm, fio_value, pos_or_eararchy_value), value in not_d.items():

        hierarchy_str = join_with_slash(value[0])
        position_str = value[0].get("Штатная должность", "")

        ws.cell(row=row, column=new_col_hierarchy_index, value=hierarchy_str)
        ws.cell(row=row, column=new_col_position_index, value=position_str)
