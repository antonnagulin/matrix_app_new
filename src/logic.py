from .filters import normalize_fio, is_dismissed
from .file_utils import read_sprav_file, load_check_file, save_file, get_fill
from .constants import FIO_COL, RED_FILL_COLOR, ORANGE_FILL_COLOR


def run_check(
    sprav_file,
    check_file,
    output_file,
    check_dismissed=True,
    check_duplicates=True,
    log=None
):
    # ====== Логгер ======
    def write(msg=""):
        if log:
            log.insert("end", msg + "\n")
            log.see("end")
        print(msg)

    # ====== Чтение справочника ======
    write("Чтение справочника...")
    spr_df = read_sprav_file(sprav_file)

    # ФИО (нормализованное) -> список записей справочника
    fio_dict = {}
    for _, row in spr_df.iterrows():
        fio_norm = normalize_fio(row[FIO_COL])
        if fio_norm:
            fio_dict.setdefault(fio_norm, []).append(row)

    write(f"Справочник загружен ({len(fio_dict)} уникальных ФИО)")

    # ====== Открытие файла проверки ======
    wb, ws = load_check_file(check_file)
    headers = [cell.value for cell in ws[1]]

    if FIO_COL not in headers:
        raise ValueError(f"В файле проверки нет колонки '{FIO_COL}'")

    fio_col_index = headers.index(FIO_COL) + 1

    red_fill = get_fill(RED_FILL_COLOR)
    orange_fill = get_fill(ORANGE_FILL_COLOR)

    write("Проверка сотрудников...")

    # ====== Списки для логов ======
    red_not_found = []       # нет в справочнике
    orange_dismissed = []    # уволенные

    # ====== Первый проход: подсветка ======
    for row in range(2, ws.max_row + 1):
        fio_cell = ws.cell(row=row, column=fio_col_index)
        fio_value = fio_cell.value
        fio_norm = normalize_fio(fio_value)

        # ❌ Нет в справочнике
        if fio_norm not in fio_dict:
            fio_cell.fill = red_fill
            red_not_found.append(str(fio_value))
            continue

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

    # ====== Лог: дубликаты в справочнике ======
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

    # ====== Сохранение ======
    save_file(wb, output_file)
    write()
    write(f"Готово! Файл сохранён: {output_file}")
