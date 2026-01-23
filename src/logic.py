from .filters import normalize_fio, is_dismissed
from .file_utils import read_sprav_file, load_check_file, save_file, get_fill
from .constants import FIO_COL, RED_FILL_COLOR, ORANGE_FILL_COLOR

def run_check(sprav_file, check_file, output_file, check_dismissed=True, check_duplicates=True, log=None):
    def write(msg):
        if log:
            log.insert("end", msg + "\n")
            log.see("end")
        print(msg)

    write("Чтение справочника...")
    spr_df = read_sprav_file(sprav_file)

    # Создаём словарь ФИО -> список записей (для дубликатов)
    fio_dict = {}
    for _, row in spr_df.iterrows():
        fio_norm = normalize_fio(row[FIO_COL])
        if fio_norm:
            fio_dict.setdefault(fio_norm, []).append(row)
    write(f"Справочник загружен ({len(fio_dict)} уникальных ФИО)")

    wb, ws = load_check_file(check_file)
    headers = [cell.value for cell in ws[1]]
    fio_col_index = headers.index(FIO_COL) + 1

    red_fill = get_fill(RED_FILL_COLOR)
    orange_fill = get_fill(ORANGE_FILL_COLOR)

    write("Проверка сотрудников...")

    for row in range(2, ws.max_row + 1):
        fio_cell = ws.cell(row=row, column=fio_col_index)
        fio_norm = normalize_fio(fio_cell.value)
        if fio_norm not in fio_dict:
            fio_cell.fill = red_fill
            continue
        if check_dismissed:
            records = fio_dict[fio_norm]
            if any(is_dismissed(r) for r in records):
                fio_cell.fill = orange_fill

    # Дубликаты
    if check_duplicates:
        write("Поиск дубликатов в справочнике (для оранжевых)...")
        seen = set()
        for row in range(2, ws.max_row + 1):
            cell = ws.cell(row=row, column=fio_col_index)
            if cell.fill == orange_fill:
                fio_norm = normalize_fio(cell.value)
                if fio_norm not in seen and len(fio_dict.get(fio_norm, [])) > 1:
                    seen.add(fio_norm)
                    write(f"\nДубликаты: {fio_norm}")
                    for rec in fio_dict[fio_norm]:
                        write(f"  {rec.to_dict()}")

    save_file(wb, output_file)
    write(f"Готово! Файл сохранён: {output_file}")
