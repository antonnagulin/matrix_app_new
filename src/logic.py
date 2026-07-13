from .filters import normalize_fio, is_dismissed
from .file_utils import read_sprav_file, load_check_file, save_file, get_fill
from .constants import FIO_COL, RED_FILL_COLOR, ORANGE_FILL_COLOR
import tkinter as tk
from tkinter import messagebox, ttk
import openpyxl


def ask_include_missing(parent, fio: str) -> str:
    """
    Спросить, включить ли отсутствующего сотрудника в отчёт.

    Returns:
        'yes' — только этого сотрудника
        'no' — не включать
        'yes_all' — этого и всех следующих отсутствующих
    """
    result = ["no"]

    def on_yes() -> None:
        result[0] = "yes"
        win.destroy()

    def on_no() -> None:
        result[0] = "no"
        win.destroy()

    def on_yes_all() -> None:
        result[0] = "yes_all"
        win.destroy()

    win = tk.Toplevel(parent)
    win.title("Сотрудник не найден")
    win.resizable(False, False)
    win.grab_set()
    win.transient(parent)

    body = ttk.Frame(win, padding=16)
    body.pack(fill="both", expand=True)

    ttk.Label(
        body,
        text=f"Включить '{fio}' в отчет?",
        font=("Arial", 11),
        wraplength=420,
    ).pack(anchor="w", pady=(0, 12))

    btn_frame = ttk.Frame(body)
    btn_frame.pack(fill="x")

    ttk.Button(btn_frame, text="Да", command=on_yes, width=12).pack(side="left", padx=(0, 6))
    ttk.Button(btn_frame, text="Нет", command=on_no, width=12).pack(side="left", padx=(0, 6))
    ttk.Button(btn_frame, text="Да для всех", command=on_yes_all, width=14).pack(side="left")

    win.update_idletasks()
    px = parent.winfo_rootx() + (parent.winfo_width() - win.winfo_width()) // 2
    py = parent.winfo_rooty() + (parent.winfo_height() - win.winfo_height()) // 2
    win.geometry(f"+{max(px, 0)}+{max(py, 0)}")

    win.wait_window()
    return result[0]


SKIP_ALL_DUPLICATES = "skip_all"


def choose_actual_record(parent, fio, records):
    """
    Показывает окно с дубликатами.

    Returns:
        запись — если пользователь нажал «Выбрать»
        SKIP_ALL_DUPLICATES — сформировать отчёт без выбора остальных дубликатов
        None — отмена / закрытие окна
    """
    selected = [None]

    def on_select() -> None:
        try:
            idx = tree.selection()[0]
            selected[0] = records[int(idx)]
            win.destroy()
        except Exception:
            messagebox.showwarning("Выбор", "Выберите запись!", parent=win)

    def on_skip_all() -> None:
        selected[0] = SKIP_ALL_DUPLICATES
        win.destroy()

    win = tk.Toplevel(parent)
    win.title(f"Выбор актуальной записи для {fio}")
    win.geometry("960x480")
    win.minsize(820, 420)
    win.grab_set()
    win.transient(parent)

    container = ttk.Frame(win, padding=12)
    container.pack(fill="both", expand=True)
    container.rowconfigure(1, weight=1)
    container.columnconfigure(0, weight=1)

    ttk.Label(
        container,
        text=f"Найдены дубликаты для {fio}:",
        font=("Arial", 12, "bold"),
    ).grid(row=0, column=0, sticky="w", pady=(0, 8))

    cols = list(records[0].keys())
    tree = ttk.Treeview(container, columns=cols, show="headings", height=12)
    for col in cols:
        tree.heading(col, text=col)
        tree.column(col, width=130, anchor="w")
    for i, rec in enumerate(records):
        tree.insert("", "end", iid=str(i), values=[rec.get(c, "") for c in cols])
    tree.grid(row=1, column=0, sticky="nsew", pady=(0, 12))

    btn_frame = ttk.Frame(container)
    btn_frame.grid(row=2, column=0, sticky="ew")

    btn_font = ("Segoe UI", 10)
    tk.Button(
        btn_frame,
        text="Выбрать",
        command=on_select,
        font=btn_font,
        padx=18,
        pady=8,
        width=14,
    ).pack(side="left", padx=(0, 8))
    tk.Button(
        btn_frame,
        text="Сформировать отчёт без выбора",
        command=on_skip_all,
        font=btn_font,
        padx=18,
        pady=8,
        width=28,
    ).pack(side="left", padx=(0, 8))
    tk.Button(
        btn_frame,
        text="Отмена",
        command=win.destroy,
        font=btn_font,
        padx=18,
        pady=8,
        width=14,
    ).pack(side="right")

    win.update_idletasks()
    px = parent.winfo_rootx() + (parent.winfo_width() - win.winfo_width()) // 2
    py = parent.winfo_rooty() + (parent.winfo_height() - win.winfo_height()) // 2
    win.geometry(f"+{max(px, 0)}+{max(py, 0)}")

    win.wait_window()
    return selected[0]


def run_check(
    sprav_file,
    check_file,
    output_file,
    check_dismissed=True,
    check_duplicates=True,
    prompt_duplicate_selection=True,
    log=None,
    parent=None
):
    def write(msg=""):
        if log:
            log.insert("end", msg + "\n")
            log.see("end")
        print(msg)

    write("Чтение справочника...")
    spr_df = read_sprav_file(sprav_file)

    # ====== Словарь ФИО -> список записей ======
    fio_dict = {}
    for _, row in spr_df.iterrows():
        fio_norm = normalize_fio(row[FIO_COL])
        if fio_norm:
            fio_dict.setdefault(fio_norm, []).append(row)

    write(f"Справочник загружен ({len(fio_dict)} уникальных ФИО)")

    # ====== Загрузка файла проверки ======
    wb, ws = load_check_file(check_file)
    headers = [cell.value for cell in ws[1]]
    fio_col_index = headers.index(FIO_COL) + 1

    red_fill = get_fill(RED_FILL_COLOR)
    orange_fill = get_fill(ORANGE_FILL_COLOR)

    write("Проверка сотрудников...")

    # ====== Списки ======
    red_not_found = []       # отсутствующие
    orange_dismissed = []    # уволенные
    include_in_report = []   # выбранные красные
    repit_not_found = set()

    include_all_missing = False

    # ====== Первый проход ======
    for row in range(2, ws.max_row + 1):
        fio_cell = ws.cell(row=row, column=fio_col_index)
        fio_value = fio_cell.value
        fio_norm = normalize_fio(fio_value)

        # Игнорируем пустые строки
        if not fio_value or str(fio_value).strip() == "":
            continue

        # ❌ Нет в справочнике → спрашиваем пользователя (один раз на ФИО)
        if fio_norm not in fio_dict:
            if fio_norm not in repit_not_found:
                fio_cell.fill = red_fill
                red_not_found.append((row, fio_cell.value))

                if parent:
                    if include_all_missing:
                        include_in_report.append(row)
                    else:
                        answer = ask_include_missing(parent, fio_cell.value)
                        if answer == "yes_all":
                            include_all_missing = True
                            include_in_report.append(row)
                        elif answer == "yes":
                            include_in_report.append(row)

                repit_not_found.add(fio_norm)
            elif include_all_missing:
                include_in_report.append(row)
            continue

        # 🟠 Проверка увольнения
        if check_dismissed:
            records = fio_dict[fio_norm]
            if any(is_dismissed(rec) for rec in records):
                fio_cell.fill = orange_fill
                orange_dismissed.append(str(fio_cell.value))

    # ====== Новый XLSX для выбранных красных ======
    if include_in_report:
        new_wb = openpyxl.Workbook()
        new_ws = new_wb.active
        for col_idx, header in enumerate(headers, start=1):
            new_ws.cell(row=1, column=col_idx, value=header)
        for i, row_idx in enumerate(include_in_report, start=2):
            for col_idx, header in enumerate(headers, start=1):
                new_ws.cell(row=i, column=col_idx, value=ws.cell(row=row_idx, column=col_idx).value)
        report_file = output_file.replace(".xlsx", "_included.xlsx")
        new_wb.save(report_file)
        write(f"Сгенерирован файл с выбранными сотрудниками: {report_file}")

    # ====== Лог: красные ======
    if red_not_found:
        write("\nСотрудники, отсутствующие в справочнике (красным):")
        for _, fio in red_not_found:
            write(f"  - {fio}")

    # ====== Второй проход: дубликаты оранжевых ======
    if check_duplicates:
        write("\nПоиск дубликатов в справочнике (для уволенных)...")
        seen = set()
        found_any = False
        skip_duplicate_selection = not prompt_duplicate_selection
        for fio in sorted(set(orange_dismissed)):
            fio_norm = normalize_fio(fio)
            records = fio_dict.get(fio_norm, [])
            if len(records) > 1 and fio_norm not in seen:
                seen.add(fio_norm)
                found_any = True
                write("\n" + "="*90)
                write(f"Дубликаты в справочнике для: {fio}")
                write("="*90)

                for idx, rec in enumerate(records, start=1):
                    write(f"\nЗапись #{idx}")
                    for col, value in rec.items():
                        write(f"  {col}: {value}")

                if parent and not skip_duplicate_selection:
                    choice = choose_actual_record(parent, fio, records)
                    if choice is SKIP_ALL_DUPLICATES:
                        skip_duplicate_selection = True
                        write("Дальнейший выбор дубликатов пропущен — отчёт сформирован без ручного выбора.")
                    elif choice is not None:
                        write(f"Выбрана актуальная запись: {choice}")

        if not found_any:
            write("Дубликатов не найдено")
        elif skip_duplicate_selection and not prompt_duplicate_selection:
            write("Дубликаты записаны в журнал без ручного выбора.")

    # ====== Сохранение файла с подсветкой ======
    save_file(wb, output_file)
    write(f"\nГотово! Файл сохранён: {output_file}")
