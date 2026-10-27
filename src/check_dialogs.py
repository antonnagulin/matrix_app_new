"""Диалоги выбора для проверки сотрудников (tkinter)."""

from __future__ import annotations

import tkinter as tk
from tkinter import messagebox, ttk

SKIP_ALL_DUPLICATES = "skip_all"


def ask_include_missing(parent, fio: str) -> str:
    """
    Спросить, включить ли отсутствующего сотрудника в отчёт.

    Returns:
        'yes' — только этого сотрудника
        'no' — не включать
        'yes_all' — этого и всех следующих отсутствующих
        'no_all' — не включать этого и всех следующих отсутствующих
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

    def on_no_all() -> None:
        result[0] = "no_all"
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
    ttk.Button(btn_frame, text="Да для всех", command=on_yes_all, width=14).pack(
        side="left", padx=(0, 6)
    )
    ttk.Button(btn_frame, text="Нет для всех", command=on_no_all, width=14).pack(side="left")

    win.update_idletasks()
    px = parent.winfo_rootx() + (parent.winfo_width() - win.winfo_width()) // 2
    py = parent.winfo_rooty() + (parent.winfo_height() - win.winfo_height()) // 2
    win.geometry(f"+{max(px, 0)}+{max(py, 0)}")

    win.wait_window()
    return result[0]


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
