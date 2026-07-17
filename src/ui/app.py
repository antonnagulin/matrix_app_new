"""Главное окно приложения для Windows."""

from __future__ import annotations

import os
import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk

from src.logic import run_check
from src.file_utils import peek_sheet_rows
from src.ui.styles import COLORS, FONT_FAMILY, setup_styles


class LogWriter:
  """Адаптер лога для ScrolledText с обновлением UI."""

  def __init__(self, app: "EmployeeCheckApp") -> None:
    self._app = app

  def insert(self, _index: str, text: str) -> None:
    self._app._append_log(text)
    self._app.root.update_idletasks()

  def see(self, _index: str) -> None:
    self.log.see("end")

  @property
  def log(self):
    return self._app.log


class EmployeeCheckApp:
  """Проверка сотрудников по Excel-справочнику."""

  WINDOW_TITLE = "Проверка сотрудников"
  MIN_WIDTH = 860
  MIN_HEIGHT = 640

  def __init__(self, root: tk.Tk) -> None:
    self.root = root
    self._running = False
    self._last_output_dir: str | None = None
    self._log_writer = LogWriter(self)

    self.sprav = tk.StringVar()
    self.check = tk.StringVar()
    self.output = tk.StringVar(value="check_result.xlsx")
    self.dismissed = tk.BooleanVar(value=True)
    self.duplicates = tk.BooleanVar(value=True)
    self.prompt_duplicates = tk.BooleanVar(value=True)
    self.auto_header = tk.BooleanVar(value=True)
    self.header_row = tk.IntVar(value=1)
    self.header_info = tk.StringVar(value="Строка заголовков: авто")
    self.status = tk.StringVar(value="Готово к работе")

    self._configure_window()
    setup_styles(root)
    self._build_ui()
    self._bind_events()

  def _configure_window(self) -> None:
    self.root.title(self.WINDOW_TITLE)
    self.root.minsize(self.MIN_WIDTH, self.MIN_HEIGHT)
    self.root.configure(bg=COLORS["bg"])

    if sys.platform == "win32":
      try:
        self.root.state("zoomed")
      except tk.TclError:
        self.root.geometry("1000x720")
    else:
      self.root.geometry("1000x720")

    self._center_if_needed()

  def _center_if_needed(self) -> None:
    if self.root.state() == "zoomed":
      return
    self.root.update_idletasks()
    w = self.root.winfo_width()
    h = self.root.winfo_height()
    x = (self.root.winfo_screenwidth() // 2) - (w // 2)
    y = (self.root.winfo_screenheight() // 2) - (h // 2)
    self.root.geometry(f"{w}x{h}+{x}+{y}")

  def _build_ui(self) -> None:
    self._build_header()
    self._build_toolbar()

    body = ttk.Frame(self.root, padding=(16, 8, 16, 8))
    body.pack(fill="both", expand=True)

    self._build_files_section(body)
    self._build_options_section(body)
    self._build_legend(body)
    self._build_log_section(body)
    self._build_status_bar()

  def _build_header(self) -> None:
    header = tk.Frame(self.root, bg=COLORS["header_bg"], height=72)
    header.pack(fill="x")
    header.pack_propagate(False)

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

    self.run_btn = ttk.Button(bar, text="Запустить проверку", style="Accent.TButton", command=self._on_run)
    self.run_btn.pack(side="left")

    self.open_btn = ttk.Button(bar, text="Открыть папку результата", command=self._open_output_folder, state="disabled")
    self.open_btn.pack(side="left", padx=(8, 0))

    ttk.Button(bar, text="Очистить лог", command=self._clear_log).pack(side="left", padx=(8, 0))

    ttk.Button(bar, text="Выход", command=self.root.destroy).pack(side="right")

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
    self._file_row(frame, 2, "Файл результата", self.output, open_file=False, save_file=True)

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
  ) -> None:
    ttk.Label(parent, text=f"{label}:").grid(row=row, column=0, sticky="w", padx=(0, 8), pady=6)

    entry = ttk.Entry(parent, textvariable=var)
    entry.grid(row=row, column=1, sticky="ew", pady=6)

    btn_frame = ttk.Frame(parent)
    btn_frame.grid(row=row, column=2, sticky="e", padx=(8, 0), pady=6)

    if open_file:
      ttk.Button(
        btn_frame,
        text="Обзор…",
        command=lambda: self._pick_open_file(var, on_change),
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
      text="Искать дубликаты в справочнике для уволенных",
      variable=self.duplicates,
    ).pack(anchor="w", pady=2)

    ttk.Checkbutton(
      frame,
      text="Запрашивать выбор при дубликатах (иначе — только журнал)",
      variable=self.prompt_duplicates,
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

    self._legend_item(frame, COLORS["red_marker"], "Красный — сотрудник не найден в справочнике")
    self._legend_item(frame, COLORS["orange_marker"], "Оранжевый — сотрудник уволен или неактивен")
    self._legend_item(frame, COLORS["yellow_marker"], "Жёлтый — дубликат в справочнике")
    self._legend_item(frame, COLORS["green_marker"], "Зелёный — дубликат разрешён однозначно")

  def _legend_item(self, parent: ttk.Frame, color: str, text: str) -> None:
    row = ttk.Frame(parent)
    row.pack(side="left", padx=(0, 24))

    swatch = tk.Label(row, width=2, bg=color, relief="solid", borderwidth=1)
    swatch.pack(side="left", padx=(0, 6))
    ttk.Label(row, text=text).pack(side="left")

  def _build_log_section(self, parent: ttk.Frame) -> None:
    frame = ttk.LabelFrame(parent, text="Журнал выполнения", padding=8)
    frame.pack(fill="both", expand=True)

    self.log = scrolledtext.ScrolledText(
      frame,
      font=("Consolas", 10),
      wrap="word",
      bg=COLORS["surface"],
      fg=COLORS["text"],
      relief="flat",
      borderwidth=1,
      highlightthickness=1,
      highlightbackground=COLORS["border"],
      highlightcolor=COLORS["accent"],
    )
    self.log.pack(fill="both", expand=True)
    self.log.insert("end", "Выберите файлы и нажмите «Запустить проверку».\n")
    self.log.configure(state="disabled")

  def _build_status_bar(self) -> None:
    bar = tk.Frame(self.root, bg="#e9e9e9", height=28)
    bar.pack(fill="x", side="bottom")
    bar.pack_propagate(False)

    ttk.Label(bar, textvariable=self.status, style="Status.TLabel").pack(side="left", fill="x", expand=True)

  def _bind_events(self) -> None:
    self.root.protocol("WM_DELETE_WINDOW", self._on_close)
    self.root.bind("<Control-Return>", lambda _e: self._on_run())
    self.root.bind("<F5>", lambda _e: self._on_run())

  def _pick_open_file(self, var: tk.StringVar, on_change=None) -> None:
    path = filedialog.askopenfilename(
      title="Выберите Excel-файл",
      filetypes=[("Excel", "*.xlsx"), ("Все файлы", "*.*")],
    )
    if path:
      var.set(path)
      if on_change:
        on_change(path)

  def _pick_save_file(self, var: tk.StringVar) -> None:
    path = filedialog.asksaveasfilename(
      title="Сохранить результат как",
      defaultextension=".xlsx",
      filetypes=[("Excel", "*.xlsx")],
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
    self._detect_header_for_check(check_path)

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
      _rows, detected = peek_sheet_rows(check_path, max_rows=5)
    except Exception as exc:
      self.header_info.set(f"Не удалось прочитать файл: {exc}")
      return

    if detected:
      self.header_row.set(detected)
      if self.auto_header.get():
        self.header_info.set(f"Строка заголовков: авто → {detected}")
      else:
        self.header_info.set(f"Строка заголовков: вручную → {self.header_row.get()}")
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
      preview_rows, detected = peek_sheet_rows(path, max_rows=40)
    except Exception as exc:
      messagebox.showerror(self.WINDOW_TITLE, f"Не удалось прочитать файл:\n{exc}", parent=self.root)
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
    ttk.Label(body, text=hint, wraplength=860).grid(row=0, column=0, sticky="w", pady=(0, 8))

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

  def _clear_log(self) -> None:
    self._set_log_text("")
    self._append_log("Журнал очищен.\n")
    self.status.set("Журнал очищен")

  def _set_log_text(self, text: str) -> None:
    self.log.configure(state="normal")
    self.log.delete("1.0", tk.END)
    if text:
      self.log.insert("end", text)
    self.log.configure(state="disabled")

  def _append_log(self, text: str) -> None:
    self.log.configure(state="normal")
    self.log.insert("end", text)
    self.log.see("end")
    self.log.configure(state="disabled")

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
      messagebox.showerror(self.WINDOW_TITLE, "Укажите файл справочника.", parent=self.root)
      return False
    if not self.check.get().strip():
      messagebox.showerror(self.WINDOW_TITLE, "Укажите файл проверки.", parent=self.root)
      return False
    if not Path(self.sprav.get()).is_file():
      messagebox.showerror(self.WINDOW_TITLE, "Файл справочника не найден.", parent=self.root)
      return False
    if not Path(self.check.get()).is_file():
      messagebox.showerror(self.WINDOW_TITLE, "Файл проверки не найден.", parent=self.root)
      return False
    if not self.output.get().strip():
      messagebox.showerror(self.WINDOW_TITLE, "Укажите имя выходного файла.", parent=self.root)
      return False
    return True

  def _on_run(self) -> None:
    if self._running:
      return
    if not self._validate_inputs():
      return

    self._set_log_text("")
    self._set_running(True)
    self.status.set("Выполняется проверка…")
    self.open_btn.configure(state="disabled")
    self.root.update_idletasks()

    output_path = self.output.get().strip()
    error: str | None = None

    try:
      header_row = None if self.auto_header.get() else int(self.header_row.get())
      run_check(
        sprav_file=self.sprav.get().strip(),
        check_file=self.check.get().strip(),
        output_file=output_path,
        check_dismissed=self.dismissed.get(),
        check_duplicates=self.duplicates.get(),
        prompt_duplicate_selection=self.prompt_duplicates.get(),
        header_row=header_row,
        log=self._log_writer,
        parent=self.root,
      )
    except Exception as exc:
      error = str(exc)

    self._on_run_finished(output_path, error)

  def _on_run_finished(self, output_path: str, error: str | None) -> None:
    self._set_running(False)

    if error:
      self.status.set("Ошибка при проверке")
      messagebox.showerror(self.WINDOW_TITLE, error, parent=self.root)
      return

    self._last_output_dir = str(Path(output_path).resolve().parent)
    self.open_btn.configure(state="normal")
    self.status.set(f"Готово: {output_path}")

    messagebox.showinfo(
      self.WINDOW_TITLE,
      f"Проверка завершена.\n\nРезультат сохранён:\n{output_path}",
      parent=self.root,
    )

  def _open_output_folder(self) -> None:
    folder = self._last_output_dir
    if not folder:
      folder = str(Path(self.output.get()).resolve().parent)
    if not folder or not Path(folder).is_dir():
      messagebox.showwarning(self.WINDOW_TITLE, "Папка результата пока недоступна.", parent=self.root)
      return

    if sys.platform == "win32":
      os.startfile(folder)  # noqa: S606
    else:
      messagebox.showinfo(self.WINDOW_TITLE, folder, parent=self.root)

  def _on_close(self) -> None:
    if self._running:
      if not messagebox.askyesno(
        self.WINDOW_TITLE,
        "Проверка ещё выполняется. Закрыть приложение?",
        parent=self.root,
      ):
        return
    self.root.destroy()
