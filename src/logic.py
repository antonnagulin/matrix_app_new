from .filters import normalize_fio, is_dismissed
from .file_utils import read_sprav_file, load_check_file, save_file, get_fill
from .constants import FIO_COL, RED_FILL_COLOR, ORANGE_FILL_COLOR
import openpyxl
import pandas as pd

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
    log=None
):
    def write(msg=""):
        if log:
            log.insert("end", msg + "\n")
            log.see("end")
        print(msg)

    write("Чтение справочника...")
    spr_df = read_sprav_file(sprav_file)

    # ФИО -> список записей
    fio_dict = {}
    for _, row in spr_df.iterrows():
        fio_norm = normalize_fio(row[FIO_COL])
        if fio_norm:
            fio_dict.setdefault(fio_norm, []).append(row)

    write(f"Справочник загружен ({len(fio_dict)} уникальных ФИО)")

    wb, ws = load_check_file(check_file)
    headers = [cell.value for cell in ws[1]]

    if FIO_COL not in headers:
        raise ValueError(f"В файле проверки нет колонки '{FIO_COL}'")

    fio_col_index = headers.index(FIO_COL) + 1

    # ====== Добавляем новые столбцы справа ======
    col_hierarchy_name = "Иерархия из справочника"
    col_position_name = "Штатная должность (справочник)"
    ws.insert_cols(len(headers) + 1)
    ws.cell(row=1, column=len(headers) + 1, value=col_hierarchy_name)
    ws.insert_cols(len(headers) + 2)
    ws.cell(row=1, column=len(headers) + 2, value=col_position_name)
    new_col_hierarchy_index = len(headers) + 1
    new_col_position_index = len(headers) + 2

    red_fill = get_fill(RED_FILL_COLOR)
    orange_fill = get_fill(ORANGE_FILL_COLOR)

    write("Проверка сотрудников...")

    red_not_found = []
    orange_dismissed = []

    for row in range(2, ws.max_row + 1):
        fio_cell = ws.cell(row=row, column=fio_col_index)
        fio_value = fio_cell.value
        if not fio_value or str(fio_value).strip() == "":
            continue

        fio_norm = normalize_fio(fio_value)

        # ❌ Нет в справочнике
        if fio_norm not in fio_dict:
            fio_cell.fill = red_fill
            ws.cell(row=row, column=new_col_hierarchy_index, value="")
            ws.cell(row=row, column=new_col_position_index, value="")
            red_not_found.append(str(fio_value))
            continue

        # 🟢 Берем первую запись из справочника
        record = fio_dict[fio_norm][0]
        hierarchy_str = join_with_slash(record)
        position_str = record.get("Штатная должность", "")

        ws.cell(row=row, column=new_col_hierarchy_index, value=hierarchy_str)
        ws.cell(row=row, column=new_col_position_index, value=position_str)

        # 🟠 Проверка увольнения
        if check_dismissed:
            records = fio_dict[fio_norm]
            if any(is_dismissed(rec) for rec in records):
                fio_cell.fill = orange_fill
                orange_dismissed.append(str(fio_value))

    # ====== Лог: отсутствующие ======
    write()
    if red_not_found:
        write("Сотрудники, отсутствующие в справочнике (красным):")
        for fio in sorted(set(red_not_found)):
            write(f"  - {fio}")
    else:
        write("Сотрудников, отсутствующих в справочнике, не найдено")

    # ====== Лог: дубликаты ======
    if check_duplicates:
        write()
        write("Поиск дубликатов в справочнике (для уволенных)...")
        seen = set()
        found_any = False

        for fio in sorted(set(orange_dismissed)):
            fio_norm = normalize_fio(fio)
            records = fio_dict.get(fio_norm, [])
            if len(records) > 1 and fio_norm not in seen:
                seen.add(fio_norm)
                found_any = True
                write()
                write("=" * 90)
                write(f"Дубликаты в справочнике для: {fio}")
                write("=" * 90)
                for idx, rec in enumerate(records, start=1):
                    write(f"\nЗапись #{idx}")
                    for col, value in rec.items():
                        write(f"  {col}: {value}")

        if not found_any:
            write("Дубликатов не найдено")

    save_file(wb, output_file)
    write()
    write(f"Готово! Файл сохранён: {output_file}")
