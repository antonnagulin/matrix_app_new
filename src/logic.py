import sys
from tkinter.scrolledtext import ScrolledText

from src.filters import normalize_fio, is_dismissed, join_with_slash
from src.file_utils import read_sprav_file, load_check_file, get_fill, save_file
from src.constants import (
    FIO_COL, RED_FILL_COLOR, ORANGE_FILL_COLOR,
    ELOYY_FILL_COLOR, GREEN_FILL_COLOR, BROWN_FILL_COLOR
)

from src.check_chainge_fio import chainge_fio
from src.check_duplicate import check_duplicate, second_check_duplicate
from src.check_dissmised import check_dissmised

from src.paint_row_for_dict import paint_rows_by_dict
from src.write_position_eararchi import write_position_eararchi


def run_check(
    sprav_file,
    check_file,
    output_file,
    chainge_fio_file,
    log: ScrolledText=None,
):
    def write(msg=""):
        if log:
            log.insert("end", msg + "\n")
            log.see("end")
        print(msg)

    write("Чтение справочника...")
    spr_df = read_sprav_file(sprav_file)
    wb, ws = load_check_file(check_file)
    headers = [cell.value for cell in ws[1]]
    if FIO_COL not in headers:
        raise ValueError(f"В файле проверки нет колонки '{FIO_COL}'")

    fio_col_index = headers.index(FIO_COL) + 1

    # ====== Добавляем новые столбцы справа ======
    col_hierarchy_name = "Подразделение"
    col_position_name = "Штатная должность (справочник)"
    ws.insert_cols(len(headers) + 1)
    ws.cell(row=1, column=len(headers) + 1, value=col_hierarchy_name)
    ws.insert_cols(len(headers) + 2)
    ws.cell(row=1, column=len(headers) + 2, value=col_position_name)
    new_col_hierarchy_index = len(headers) + 1
    new_col_position_index = len(headers) + 2

    red_fill = get_fill(RED_FILL_COLOR)
    orange_fill = get_fill(ORANGE_FILL_COLOR)
    yellow_fill = get_fill(ELOYY_FILL_COLOR)
    green_fill = get_fill(GREEN_FILL_COLOR)
    brown_fill = get_fill(BROWN_FILL_COLOR)

    write("Проверка сотрудников...")

            # ФИО -> список записей
    fio_dict = {}
    for _, row in spr_df.iterrows():
        fio_norm = normalize_fio(row[FIO_COL])
        if fio_norm:
            fio_dict.setdefault(fio_norm, []).append(row)
    
    
    write(f"Справочник загружен ({len(fio_dict)} уникальных ФИО)")

    red_not_found = []
    orange_dismissed = []
    orange_dismissed_test = dict()
    elloyy_dublicate = dict()

    for row in range(2, ws.max_row + 1):
        fio_cell = ws.cell(row=row, column=fio_col_index)
        fio_value = fio_cell.value

        pos_or_eararchy_cell = ws.cell(row=row, column=fio_col_index+1)
        pos_or_eararchy_value = pos_or_eararchy_cell.value

        if not fio_value or str(fio_value).strip() == "":
            continue

        fio_norm = normalize_fio(fio_value)


        if fio_norm not in fio_dict:
            fio_cell.fill = red_fill
            ws.cell(row=row, column=new_col_hierarchy_index, value="")
            ws.cell(row=row, column=new_col_position_index, value="")
            red_not_found.append(str(fio_value))
            # red_not_found.append((row, fio_norm, str(fio_value), pos_or_eararchy_value))
            continue


        if len(fio_dict[fio_norm]) == 1:
            

                hierarchy_str = join_with_slash(fio_dict[fio_norm][0])
                position_str = fio_dict[fio_norm][0].get("Штатная должность", "")

                ws.cell(row=row, column=new_col_hierarchy_index, value=hierarchy_str)
                ws.cell(row=row, column=new_col_position_index, value=position_str)


                # if any(is_dismissed(rec) for rec in fio_dict[fio_norm]):
                if is_dismissed(fio_dict[fio_norm][0]):
                    # sys.exit()
                    fio_cell.fill = orange_fill
                    orange_dismissed.append(str(fio_value))
                    orange_dismissed_test[(row, fio_norm, fio_value, pos_or_eararchy_value)] = fio_dict[fio_norm]


        if len(fio_dict[fio_norm]) > 1:

            elloyy_dublicate[(row, fio_norm, fio_value, pos_or_eararchy_value)] = fio_dict[fio_norm]

            fio_cell_1 = ws.cell(row=row, column=fio_col_index)
            fio_cell_1.fill = yellow_fill

            # tmp = list()
            # for item in fio_dict[fio_norm]:
            #     tmp.append(item['Таб.№'])
            

            # dub = ws.cell(row=row, column=new_col_position_index+1, value=str(tmp))
            # dub.fill = yellow_fill

            # tmp=list()


    chainge_fio(red_not_found, chainge_fio_file)
    

    new_d, not_d, dismissed = check_duplicate(elloyy_dublicate)

    # print("perple_dublicate:", len(elloyy_dublicate))
    # print(
    #     "sum:",
    #     f"{len(new_d)} +", f"{len(not_d)} +", f"{len(dismissed)} =",
    #     len(new_d) + len(not_d) + len(dismissed)
    # )

    paint_rows_by_dict(ws, new_d, yellow_fill, fio_col_index+1)
    paint_rows_by_dict(ws, not_d, green_fill, fio_col_index+1)
    write_position_eararchi(ws, not_d, new_col_hierarchy_index, new_col_position_index)
    new_d, not_d, = second_check_duplicate(new_d)

    paint_rows_by_dict(ws, new_d, yellow_fill, fio_col_index+2)
    paint_rows_by_dict(ws, not_d, green_fill, fio_col_index+2)
    write_position_eararchi(ws, not_d, new_col_hierarchy_index, new_col_position_index)

    
    not_d = check_dissmised(orange_dismissed_test)
    paint_rows_by_dict(ws, not_d, brown_fill, fio_col_index+2)


    #     # ====== Лог: отсутствующие ======
    # write()
    # if red_not_found:
    #     write("Сотрудники, отсутствующие в справочнике (красным):")
    #     for fio in sorted(set(red_not_found)):
    #         write(f"  - {fio}")
    # else:
    #     write("Сотрудников, отсутствующих в справочнике, не найдено")

    # ====== Лог: дубликаты ======

    # write()
    # write("Поиск дубликатов в справочнике (для уволенных)...")
    # seen = set()
    # found_any = False

    # for fio in sorted(set(orange_dismissed)):
    #     fio_norm = normalize_fio(fio)
    #     records = fio_dict.get(fio_norm, [])
    #     if len(records) > 1 and fio_norm not in seen:
    #         seen.add(fio_norm)
    #         found_any = True
    #         write()
    #         write("=" * 90)
    #         write(f"Дубликаты в справочнике для: {fio}")
    #         write("=" * 90)
    #         for idx, rec in enumerate(records, start=1):
    #             write(f"\nЗапись #{idx}")
    #             for col, value in rec.items():
    #                 write(f"  {col}: {value}")

    # if not found_any:
    #         write("Дубликатов не найдено")


    save_file(wb, output_file)
    write()
    write(f"Готово! Файл сохранён: {output_file}")

