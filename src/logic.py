from .filters import normalize_fio, is_dismissed
from .file_utils import read_sprav_file, load_check_file, save_file, get_fill
from .constants import FIO_COL, RED_FILL_COLOR, ORANGE_FILL_COLOR, ELOYY_FILL_COLOR
import openpyxl
import pandas as pd
from .core import core


def remove_first_two_words(text):
    if pd.isna(text):
        return ""
    words = str(text).split()
    return " ".join(words[2:]) if len(words) > 2 else ""


def join_with_slash(row):
    razdel = remove_first_two_words(row.get("Раздел персонала", ""))
    parts = [
        'ООО "Газпром добыча Краснодар"',
        razdel,
        row.get("Орг. единица уровня 2", ""),
        row.get("Орг. единица уровня 3", ""),
        row.get("Орг. единица уровня 4", ""),
    ]
    parts = [str(p).strip() for p in parts if pd.notna(p) and str(p).strip()]
    return " / ".join(parts)


def run_check(
    sprav_file,
    check_file,
    output_file,
    check_dismissed=True,
    check_duplicates=True,
    log=None,
    chainge_fio=None,
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

    write("Проверка сотрудников...")

    core(
        fio_col_index,
        ws,
        spr_df,
        red_fill,
        new_col_hierarchy_index,
        new_col_position_index,
        join_with_slash,
        check_dismissed,
        orange_fill,
        yellow_fill,
        write,
        check_duplicates,
        wb,
        output_file,
        chainge_fio,
    )


if __name__ == "__main__":
    run_check(
        sprav_file="/home/wol/programs/scripts/matrix_applications/данные/spravochnik.xlsx",
        # sprav_file="/home/wol/programs/scripts/matrix_applications/данные/справочник работающих сотрудников.xlsx",
        # check_file="/home/wol/programs/scripts/matrix_applications/данные/Актуальные МАТРИЦЫ ДОСТУПА/Матрица доступа к  СЭД Общее делопроизводство.xlsx",
        # check_file="/home/wol/programs/scripts/matrix_applications/данные/Актуальные МАТРИЦЫ ДОСТУПА/Кадровый.xlsx",
        check_file="/home/wol/programs/scripts/matrix_applications/данные/Актуальные МАТРИЦЫ ДОСТУПА/Матрица доступа к ССД Инфотех.xlsx",
        # output_file="/home/wol/programs/scripts/matrix_applications/данные/результат2.xlsx",
        output_file="/home/wol/programs/scripts/matrix_applications/данные/результат1.xlsx",
        check_dismissed=True,
        check_duplicates=True,
        log=None,
        chainge_fio = "/home/wol/programs/scripts/matrix_applications/данные/список_поменявших_фамилию.xlsx",
    )
    # run_check(
    #     # sprav_file="/home/wol/programs/scripts/matrix_applications/данные/spravochnik.xlsx",
    #     sprav_file="/home/wol/programs/scripts/matrix_applications/данные/справочник работающих сотрудников.xlsx",
    #     check_file="/home/wol/programs/scripts/matrix_applications/данные/Актуальные МАТРИЦЫ ДОСТУПА/Матрица доступа к  СЭД Общее делопроизводство.xlsx",
    #     # check_file="/home/wol/programs/scripts/matrix_applications/данные/Актуальные МАТРИЦЫ ДОСТУПА/Матрица доступа к АС «Гран-ВУР».xlsx",

    #     output_file="/home/wol/programs/scripts/matrix_applications/данные/результат2.xlsx",
    #     # output_file="/home/wol/programs/scripts/matrix_applications/данные/результат1.xlsx",
    #     check_dismissed=True,
    #     check_duplicates=True,
    #     log=None,
    #     chainge_fio = "/home/wol/programs/scripts/matrix_applications/данные/список_поменявших_фамилию.xlsx",
    # )
