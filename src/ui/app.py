"""Главное окно приложения."""

from __future__ import annotations

import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk

from src.file_utils import list_sheet_names, peek_sheet_rows
from src.logic import run_check
from src.platform_utils import open_folder
from src.ui.dialog_bridge import TkDialogBridge
from src.ui.styles import COLORS, FONT_FAMILY, setup_styles


class ThreadSafeLogWriter:
    """Журнал в ScrolledText: запись из worker через root.after."""

    def __init__(self, app: "EmployeeCheckApp") -> None:
        self._app = app

    def insert(self, _index: str, text: str) -> None:
        def _append() -> None:
            widget = self._app.log
            widget.configure(state="normal")
            widget.insert("end", text)
            widget.see("end")
            widget.configure(state="disabled")
            line = text.strip()
            if line:
                last = [part for part in line.splitlines() if part.strip()]
                if last:
                    self._app.status.set(last[-1][:160])

        if threading.current_thread() is threading.main_thread():
            _append()
        else:
            self._app.root.after(0, _append)

    def see(self, _index: str) -> None:
        return

    def clear(self) -> None:
        self._app.log.configure(state="normal")
        self._app.log.delete("1.0", "end")
        self._app.log.configure(state="disabled")


class EmployeeCheckApp:
    """Проверка сотрудников по Excel-справочнику."""

    WINDOW_TITLE = "Проверка сотрудников"
    DEFAULT_WIDTH = 1040
    DEFAULT_HEIGHT = 820
    MIN_WIDTH = 900
    MIN_HEIGHT = 700

    RESULT_COLUMNS = (
        ("row", "№", 60),
        ("fio", "ФИО", 220),
        ("hierarchy", "Подразделение", 220),
        ("position", "Должность", 200),
        ("comment", "Комментарий", 280),
    )

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self._running = False
        self._last_output_dir: str | None = None
        self._dialog_bridge = TkDialogBridge(root)

        self.sprav = tk.StringVar()
        self.check = tk.StringVar()
        self.output = tk.StringVar(value="check_result.xlsx")
        self.renamed = tk.StringVar()
        self.sheet_name = tk.StringVar(value="")
        self.dismissed = tk.BooleanVar(value=True)
        self.duplicates = tk.BooleanVar(value=True)
        self.prompt_duplicates = tk.BooleanVar(value=True)
        self.include_unresolved = tk.BooleanVar(value=False)
        self.auto_header = tk.BooleanVar(value=True)
        self.header_row = tk.IntVar(value=1)
        self.header_info = tk.StringVar(value="Строка заголовков: авто")
        self.status = tk.StringVar(value="Готово к работе")

        self._configure_window()
        setup_styles(root)
        self._build_ui()
        self._log_writer = ThreadSafeLogWriter(self)
        self._bind_events()
        self._apply_window_geometry()

    def _configure_window(self) -> None:
        self.root.title(self.WINDOW_TITLE)
        self.root.minsize(self.MIN_WIDTH, self.MIN_HEIGHT)
        self.root.configure(bg=COLORS["bg"])

    def _apply_window_geometry(self) -> None:
        self.root.update_idletasks()
        width = max(self.DEFAULT_WIDTH, self.root.winfo_reqwidth())
        height = max(self.DEFAULT_HEIGHT, self.root.winfo_reqheight())
        screen_w = self.root.winfo_screenwidth()
        screen_h = self.root.winfo_screenheight()
        width = min(width, max(screen_w - 40, self.MIN_WIDTH))
        height = min(height, max(screen_h - 80, self.MIN_HEIGHT))
        x = max((screen_w - width) // 2, 0)
        y = max((screen_h - height) // 2, 0)
        self.root.geometry(f"{width}x{height}+{x}+{y}")

    def _build_ui(self) -> None:
        self._build_header()
        self._build_toolbar()

        body = ttk.Frame(self.root, padding=(16, 8, 16, 8))
        body.pack(fill="both", expand=True)

        self._build_files_section(body)
        self._build_options_section(body)
        self._build_legend(body)
        self._build_results_section(body)
        self._build_log_section(body)
        self._build_status_bar()

    def _build_header(self) -> None:
        header = tk.Frame(self.root, bg=COLORS["header_bg"])
        header.pack(fill="x")

        inner = tk.Frame(header, bg=COLORS["header_bg"])
        inner.pack(fill="both", expand=True, padx=20, pady=12)

        tk.Label(
            inner,
            text=self.WINDOW_TITLE,
            font=(FONT_FAMILY, 16, "bold"),
            bg=COLORS["header_bg"],
            fg=COLORS["header_fg"],
        ).pack(anchor="w")

        tk.Label(
            inner,
            text="Сверка ФИО из Excel-файла со справочником сотрудников",
            font=(FONT_FAMILY, 10),
            bg=COLORS["header_bg"],
            fg="#deecf9",
        ).pack(anchor="w", pady=(2, 0))

    def _build_toolbar(self) -> None:
        bar = ttk.Frame(self.root, padding=(16, 10, 16, 0))
        bar.pack(fill="x")

        self.run_btn = ttk.Button(
            bar, text="Запустить проверку", style="Accent.TButton", command=self._on_run
        )
        self.run_btn.pack(side="left")

        self.open_btn = ttk.Button(
            bar, text="Открыть папку результата", command=self._open_output_folder, state="disabled"
        )
        self.open_btn.pack(side="left", padx=(8, 0))

        ttk.Button(bar, text="Очистить таблицу", command=self._clear_results).pack(
            side="left", padx=(8, 0)
        )

        ttk.Button(bar, text="Выход", command=self._on_close).pack(side="right")

        self.progress = ttk.Progressbar(bar, mode="indeterminate", length=140)
        self.progress.pack(side="right", padx=(0, 12))

    def _build_files_section(self, parent: ttk.Frame) -> None:
        frame = ttk.LabelFrame(parent, text="Файлы", padding=12)
        frame.pack(fill="x", pady=(0, 10))
        frame.columnconfigure(1, weight=1)

        self._file_row(frame, 0, "Справочник", self.sprav, open_file=True)
        self._file_row(
            frame,
            1,
            "Файл проверки",
            self.check,
            open_file=True,
            on_change=self._on_check_file_changed,
        )

        ttk.Label(frame, text="Лист:").grid(row=2, column=0, sticky="w", padx=(0, 8), pady=6)
        self.sheet_combo = ttk.Combobox(
            frame, textvariable=self.sheet_name, state="readonly", width=40
        )
        self.sheet_combo.grid(row=2, column=1, sticky="w", pady=6)
        self.sheet_combo.bind("<<ComboboxSelected>>", lambda _e: self._on_sheet_changed())

        self._file_row(
            frame,
            3,
            "Файл смены фамилий (опц.)",
            self.renamed,
            open_file=True,
            dialog_title="Выберите файл смены фамилий",
        )
        self._file_row(frame, 4, "Файл результата", self.output, open_file=False, save_file=True)

    def _file_row(
        self,
        parent: ttk.LabelFrame,
        row: int,
        label: str,
        var: tk.StringVar,
        *,
        open_file: bool = False,
        save_file: bool = False,
        on_change=None,
        dialog_title: str | None = None,
    ) -> None:
        ttk.Label(parent, text=f"{label}:").grid(
            row=row, column=0, sticky="w", padx=(0, 8), pady=6
        )

        entry = ttk.Entry(parent, textvariable=var)
        entry.grid(row=row, column=1, sticky="ew", pady=6)

        btn_frame = ttk.Frame(parent)
        btn_frame.grid(row=row, column=2, sticky="e", padx=(8, 0), pady=6)

        if open_file:
            ttk.Button(
                btn_frame,
                text="Обзор…",
                command=lambda: self._pick_open_file(var, on_change, dialog_title),
                width=10,
            ).pack(side="left")
        elif save_file:
            ttk.Button(
                btn_frame,
                text="Сохранить как…",
                command=lambda: self._pick_save_file(var),
                width=14,
            ).pack(side="left")

        ttk.Button(
            btn_frame,
            text="✕",
            width=3,
            command=lambda: var.set(""),
        ).pack(side="left", padx=(4, 0))

    def _build_options_section(self, parent: ttk.Frame) -> None:
        frame = ttk.LabelFrame(parent, text="Настройки проверки", padding=12)
        frame.pack(fill="x", pady=(0, 10))

        ttk.Checkbutton(
            frame,
            text="Проверять уволенных сотрудников (оранжевая подсветка)",
            variable=self.dismissed,
        ).pack(anchor="w", pady=2)

        ttk.Checkbutton(
            frame,
            text="Разбирать дубликаты в справочнике (жёлтый / зелёный)",
            variable=self.duplicates,
        ).pack(anchor="w", pady=2)

        ttk.Checkbutton(
            frame,
            text="Запрашивать выбор при неоднозначных дубликатах",
            variable=self.prompt_duplicates,
        ).pack(anchor="w", pady=2)

        ttk.Checkbutton(
            frame,
            text="Включать неразрешённые дубликаты в итоговый файл",
            variable=self.include_unresolved,
        ).pack(anchor="w", pady=2)

        header_frame = ttk.Frame(frame)
        header_frame.pack(fill="x", pady=(10, 0))

        self.auto_header_cb = ttk.Checkbutton(
            header_frame,
            text="Автоопределение строки заголовков (по колонке «ФИО»)",
            variable=self.auto_header,
            command=self._on_auto_header_toggle,
        )
        self.auto_header_cb.pack(anchor="w")

        row_frame = ttk.Frame(header_frame)
        row_frame.pack(fill="x", pady=(6, 0))

        ttk.Label(row_frame, text="Строка заголовков:").pack(side="left")
        self.header_spin = ttk.Spinbox(
            row_frame,
            from_=1,
            to=500,
            width=6,
            textvariable=self.header_row,
            command=self._on_header_row_manual,
        )
        self.header_spin.pack(side="left", padx=(8, 8))
        self.header_spin.bind("<FocusOut>", lambda _e: self._on_header_row_manual())
        self.header_spin.bind("<Return>", lambda _e: self._on_header_row_manual())

        ttk.Button(
            row_frame,
            text="Выбрать в таблице…",
            command=self._pick_header_row_dialog,
        ).pack(side="left", padx=(0, 12))

        ttk.Label(row_frame, textvariable=self.header_info).pack(side="left")
        self._sync_header_controls()

    def _build_legend(self, parent: ttk.Frame) -> None:
        frame = ttk.Frame(parent)
        frame.pack(fill="x", pady=(0, 10))

        self._legend_item(
            frame, COLORS["red_marker"], "Красный — сотрудник не найден в справочнике"
        )
        self._legend_item(
            frame, COLORS["orange_marker"], "Оранжевый — сотрудник уволен или неактивен"
        )
        self._legend_item(
            frame, COLORS["yellow_marker"], "Жёлтый — дубликат в справочнике"
        )
        self._legend_item(
            frame, COLORS["green_marker"], "Зелёный — дубликат разрешён однозначно"
        )

    def _legend_item(self, parent: ttk.Frame, color: str, text: str) -> None:
        row = ttk.Frame(parent)
        row.pack(side="left", padx=(0, 24))

        swatch = tk.Label(row, width=2, bg=color, relief="solid", borderwidth=1)
        swatch.pack(side="left", padx=(0, 6))
        ttk.Label(row, text=text).pack(side="left")

    def _build_results_section(self, parent: ttk.Frame) -> None:
        frame = ttk.LabelFrame(parent, text="Результаты проверки", padding=8)
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)

        cols = [key for key, _title, _width in self.RESULT_COLUMNS]
        self.results = ttk.Treeview(
            frame,
            columns=cols,
            show="headings",
            selectmode="browse",
            height=8,
        )
        for key, title, width in self.RESULT_COLUMNS:
            self.results.heading(key, text=title)
            anchor = "center" if key == "row" else "w"
            self.results.column(
                key, width=width, minwidth=50, anchor=anchor, stretch=(key != "row")
            )

        self.results.tag_configure(
            "red", background=COLORS["red_marker"], foreground=COLORS["text"]
        )
        self.results.tag_configure(
            "orange", background=COLORS["orange_marker"], foreground=COLORS["text"]
        )
        self.results.tag_configure(
            "yellow", background=COLORS["yellow_marker"], foreground=COLORS["text"]
        )
        self.results.tag_configure(
            "green", background=COLORS["green_marker"], foreground="#ffffff"
        )

        yscroll = ttk.Scrollbar(frame, orient="vertical", command=self.results.yview)
        xscroll = ttk.Scrollbar(frame, orient="horizontal", command=self.results.xview)
        self.results.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)

        self.results.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")

    def _build_log_section(self, parent: ttk.Frame) -> None:
        frame = ttk.LabelFrame(parent, text="Журнал", padding=8)
        frame.pack(fill="both", expand=False, pady=(10, 0))

        self.log = scrolledtext.ScrolledText(
            frame,
            height=8,
            wrap="word",
            font=(FONT_FAMILY, 9),
            state="disabled",
        )
        self.log.pack(fill="both", expand=True)

    def _build_status_bar(self) -> None:
        bar = tk.Frame(self.root, bg="#e9e9e9", height=28)
        bar.pack(fill="x", side="bottom")
        bar.pack_propagate(False)

        ttk.Label(bar, textvariable=self.status, style="Status.TLabel").pack(
            side="left", fill="x", expand=True
        )

    def _bind_events(self) -> None:
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.bind("<Control-Return>", lambda _e: self._on_run())
        self.root.bind("<F5>", lambda _e: self._on_run())

    def _pick_open_file(
        self, var: tk.StringVar, on_change=None, title: str | None = None
    ) -> None:
        path = filedialog.askopenfilename(
            title=title or "Выберите Excel-файл",
            filetypes=[("Файлы Excel", "*.xlsx"), ("Все файлы", "*.*")],
        )
        if path:
            var.set(path)
            if on_change:
                on_change(path)

    def _pick_save_file(self, var: tk.StringVar) -> None:
        path = filedialog.asksaveasfilename(
            title="Сохранить результат как",
            defaultextension=".xlsx",
            filetypes=[("Файлы Excel", "*.xlsx")],
            initialfile=Path(var.get() or "check_result.xlsx").name,
        )
        if path:
            var.set(path)

    def _suggest_output_name(self, check_path: str) -> None:
        if self.output.get() and self.output.get() != "check_result.xlsx":
            return
        p = Path(check_path)
        self.output.set(str(p.with_name(f"{p.stem}_result.xlsx")))

    def _on_check_file_changed(self, check_path: str) -> None:
        self._suggest_output_name(check_path)
        self._refresh_sheets(check_path)
        self._detect_header_for_check(check_path)

    def _on_sheet_changed(self) -> None:
        path = self.check.get().strip()
        if path and Path(path).is_file():
            self._detect_header_for_check(path)

    def _refresh_sheets(self, check_path: str) -> None:
        try:
            names = list_sheet_names(check_path)
        except Exception as exc:
            self.sheet_combo["values"] = ()
            self.sheet_name.set("")
            self.status.set(f"Не удалось прочитать листы: {exc}")
            return

        self.sheet_combo["values"] = names
        if names:
            current = self.sheet_name.get()
            self.sheet_name.set(current if current in names else names[0])
        else:
            self.sheet_name.set("")

    def _selected_sheet(self) -> str | None:
        name = self.sheet_name.get().strip()
        return name or None

    def _sync_header_controls(self) -> None:
        state = "disabled" if self.auto_header.get() else "normal"
        self.header_spin.configure(state=state)

    def _on_auto_header_toggle(self) -> None:
        self._sync_header_controls()
        if self.auto_header.get():
            path = self.check.get().strip()
            if path and Path(path).is_file():
                self._detect_header_for_check(path)
            else:
                self.header_info.set("Строка заголовков: авто")
        else:
            self.header_info.set(f"Строка заголовков: вручную → {self.header_row.get()}")

    def _on_header_row_manual(self) -> None:
        if self.auto_header.get():
            return
        try:
            row = int(self.header_row.get())
        except (tk.TclError, ValueError, TypeError):
            row = 1
            self.header_row.set(1)
        if row < 1:
            row = 1
            self.header_row.set(1)
        self.header_info.set(f"Строка заголовков: вручную → {row}")

    def _detect_header_for_check(self, check_path: str) -> None:
        try:
            _rows, detected = peek_sheet_rows(
                check_path, max_rows=5, sheet_name=self._selected_sheet()
            )
        except Exception as exc:
            self.header_info.set(f"Не удалось прочитать файл: {exc}")
            return

        if detected:
            self.header_row.set(detected)
            if self.auto_header.get():
                self.header_info.set(f"Строка заголовков: авто → {detected}")
            else:
                self.header_info.set(
                    f"Строка заголовков: вручную → {self.header_row.get()}"
                )
            self.status.set(f"Заголовки найдены в строке {detected}")
        else:
            self.header_info.set("Колонка «ФИО» не найдена — выберите строку вручную")
            self.auto_header.set(False)
            self._sync_header_controls()
            self.status.set("Укажите строку заголовков вручную")

    def _pick_header_row_dialog(self) -> None:
        path = self.check.get().strip()
        if not path or not Path(path).is_file():
            messagebox.showwarning(
                self.WINDOW_TITLE,
                "Сначала выберите файл проверки.",
                parent=self.root,
            )
            return

        try:
            preview_rows, detected = peek_sheet_rows(
                path, max_rows=40, sheet_name=self._selected_sheet()
            )
        except Exception as exc:
            messagebox.showerror(
                self.WINDOW_TITLE, f"Не удалось прочитать файл:\n{exc}", parent=self.root
            )
            return

        selected = [None]

        win = tk.Toplevel(self.root)
        win.title("Выбор строки заголовков")
        win.geometry("920x480")
        win.minsize(720, 360)
        win.transient(self.root)
        win.grab_set()

        body = ttk.Frame(win, padding=12)
        body.pack(fill="both", expand=True)
        body.rowconfigure(1, weight=1)
        body.columnconfigure(0, weight=1)

        hint = (
            f"Выберите строку, в которой находятся заголовки (колонка «ФИО»). "
            f"{'Найдена автоматически: ' + str(detected) if detected else 'Автоматически не найдена.'}"
        )
        ttk.Label(body, text=hint, wraplength=860).grid(
            row=0, column=0, sticky="w", pady=(0, 8)
        )

        max_cells = max((len(r) - 1 for r in preview_rows), default=1)
        cols = ["№ строки"] + [f"C{i}" for i in range(1, max_cells + 1)]

        tree = ttk.Treeview(body, columns=cols, show="headings", height=16)
        for col in cols:
            tree.heading(col, text=col)
            tree.column(col, width=110 if col != "№ строки" else 70, anchor="w")

        for row_data in preview_rows:
            values = list(row_data) + [""] * (len(cols) - len(row_data))
            tree.insert("", "end", iid=str(row_data[0]), values=values[: len(cols)])

        if detected and str(detected) in tree.get_children():
            tree.selection_set(str(detected))
            tree.see(str(detected))
        elif tree.get_children():
            tree.selection_set(tree.get_children()[0])

        tree.grid(row=1, column=0, sticky="nsew")
        scroll = ttk.Scrollbar(body, orient="vertical", command=tree.yview)
        scroll.grid(row=1, column=1, sticky="ns")
        tree.configure(yscrollcommand=scroll.set)

        def on_ok() -> None:
            sel = tree.selection()
            if not sel:
                messagebox.showwarning("Выбор", "Выберите строку заголовков.", parent=win)
                return
            selected[0] = int(sel[0])
            win.destroy()

        btn_row = ttk.Frame(body)
        btn_row.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(12, 0))
        ttk.Button(btn_row, text="Выбрать", command=on_ok).pack(side="left")
        ttk.Button(btn_row, text="Отмена", command=win.destroy).pack(side="right")

        win.wait_window()
        if selected[0] is None:
            return

        self.auto_header.set(False)
        self.header_row.set(selected[0])
        self._sync_header_controls()
        self.header_info.set(f"Строка заголовков: вручную → {selected[0]}")
        self.status.set(f"Выбрана строка заголовков: {selected[0]}")

    def _clear_results(self, *, update_status: bool = True) -> None:
        for item in self.results.get_children():
            self.results.delete(item)
        if update_status and not self._running:
            self.status.set("Таблица результатов очищена")

    def _fill_results(self, rows: list[dict] | None) -> None:
        self._clear_results(update_status=False)
        if not rows:
            return

        for item in rows:
            values = (
                item.get("row", ""),
                item.get("fio", ""),
                item.get("hierarchy", ""),
                item.get("position", ""),
                item.get("comment", ""),
            )
            status = item.get("status") or ""
            tags = (status,) if status in {"red", "orange", "yellow", "green"} else ()
            self.results.insert("", "end", values=values, tags=tags)

    def _set_running(self, running: bool) -> None:
        self._running = running
        state = "disabled" if running else "normal"
        self.run_btn.configure(state=state)
        if running:
            self.progress.start(12)
            self.root.configure(cursor="watch")
        else:
            self.progress.stop()
            self.root.configure(cursor="")

    def _validate_inputs(self) -> bool:
        if not self.sprav.get().strip():
            messagebox.showerror(
                self.WINDOW_TITLE, "Укажите файл справочника.", parent=self.root
            )
            return False
        if not self.check.get().strip():
            messagebox.showerror(
                self.WINDOW_TITLE, "Укажите файл проверки.", parent=self.root
            )
            return False
        if not Path(self.sprav.get()).is_file():
            messagebox.showerror(
                self.WINDOW_TITLE, "Файл справочника не найден.", parent=self.root
            )
            return False
        if not Path(self.check.get()).is_file():
            messagebox.showerror(
                self.WINDOW_TITLE, "Файл проверки не найден.", parent=self.root
            )
            return False
        if not self.output.get().strip():
            messagebox.showerror(
                self.WINDOW_TITLE, "Укажите имя выходного файла.", parent=self.root
            )
            return False

        sprav = Path(self.sprav.get().strip())
        check = Path(self.check.get().strip())
        output = Path(self.output.get().strip())
        renamed = self.renamed.get().strip()
        try:
            out_resolved = output.expanduser().resolve()
            for label, path in (
                ("справочника", sprav),
                ("проверки", check),
                ("смены ФИО", Path(renamed) if renamed else None),
            ):
                if path is None:
                    continue
                if path.expanduser().resolve() == out_resolved:
                    messagebox.showerror(
                        self.WINDOW_TITLE,
                        f"Файл результата совпадает с файлом {label}.\n"
                        "Укажите другое имя выходного файла.",
                        parent=self.root,
                    )
                    return False
        except OSError:
            pass

        sheet = self._selected_sheet()
        if sheet and self.sheet_combo["values"]:
            names = list(self.sheet_combo["values"])
            if sheet not in names:
                messagebox.showerror(
                    self.WINDOW_TITLE,
                    f"Лист «{sheet}» не найден в файле проверки.",
                    parent=self.root,
                )
                return False
        return True

    def _on_run(self) -> None:
        if self._running:
            return
        if not self._validate_inputs():
            return

        self._clear_results(update_status=False)
        self._log_writer.clear()
        self._set_running(True)
        self.status.set("Выполняется проверка…")
        self.open_btn.configure(state="disabled")

        output_path = self.output.get().strip()
        header_row = None if self.auto_header.get() else int(self.header_row.get())
        renamed = self.renamed.get().strip() or None
        kwargs = dict(
            sprav_file=self.sprav.get().strip(),
            check_file=self.check.get().strip(),
            output_file=output_path,
            check_dismissed=self.dismissed.get(),
            check_duplicates=self.duplicates.get(),
            prompt_duplicate_selection=self.prompt_duplicates.get(),
            header_row=header_row,
            renamed_fio_file=renamed,
            include_unresolved_duplicates=self.include_unresolved.get(),
            sheet_name=self._selected_sheet(),
            log=self._log_writer,
            parent=self._dialog_bridge,
        )

        def worker() -> None:
            error: str | None = None
            results: list[dict] = []
            stats: dict = {}
            try:
                if renamed and not Path(renamed).is_file():
                    raise FileNotFoundError(f"Файл смены ФИО не найден: {renamed}")
                out = run_check(**kwargs)
                if isinstance(out, tuple):
                    results, stats = out
                else:
                    results = out or []
            except Exception as exc:
                error = str(exc)

            def finish() -> None:
                self._on_run_finished(output_path, error, results, stats)

            if threading.current_thread() is threading.main_thread():
                finish()
            else:
                self.root.after(0, finish)

        if getattr(self, "_run_async", True):
            self._worker = threading.Thread(target=worker, daemon=True)
            self._worker.start()
        else:
            self._worker = None
            worker()

    def _on_run_finished(
        self,
        output_path: str,
        error: str | None,
        results: list[dict] | None = None,
        stats: dict | None = None,
    ) -> None:
        self._set_running(False)

        if error:
            self.status.set("Ошибка при проверке")
            messagebox.showerror(self.WINDOW_TITLE, error, parent=self.root)
            return

        self._fill_results(results or [])
        self._last_output_dir = str(Path(output_path).resolve().parent)
        self.open_btn.configure(state="normal")
        count = len(results or [])
        stats = stats or {}
        final_rows = stats.get("final_rows", "?")
        excluded = stats.get("excluded", "?")
        self.status.set(
            f"Готово: {output_path} — в итоге {final_rows} строк, подсвечено {count}"
        )

        review = Path(output_path).with_name(Path(output_path).stem + "_review.xlsx")
        messagebox.showinfo(
            self.WINDOW_TITLE,
            f"Проверка завершена.\n\n"
            f"Итоговый файл: {output_path}\n"
            f"Строк в итоге: {final_rows}\n"
            f"Не вошло в итог: {excluded} "
            f"(не найдены={stats.get('not_found', 0)}, "
            f"уволены={stats.get('dismissed', 0)}, "
            f"неразрешённые дубли={stats.get('unresolved_dups', 0) if not self.include_unresolved.get() else 0})\n\n"
            f"Файл с подсветкой: {review}\n"
            f"В таблице: {count} подсвеченных строк.",
            parent=self.root,
        )

    def _open_output_folder(self) -> None:
        folder = self._last_output_dir
        if not folder:
            folder = str(Path(self.output.get()).resolve().parent)
        if not folder or not Path(folder).is_dir():
            messagebox.showwarning(
                self.WINDOW_TITLE, "Папка результата пока недоступна.", parent=self.root
            )
            return

        if not open_folder(folder):
            messagebox.showinfo(self.WINDOW_TITLE, folder, parent=self.root)

    def _on_close(self) -> None:
        if self._running:
            if not messagebox.askyesno(
                self.WINDOW_TITLE,
                "Проверка ещё выполняется. Закрыть приложение?",
                parent=self.root,
            ):
                return
            # Разблокировать worker, ждущий диалог
            self._dialog_bridge.close()
        self.root.destroy()
