"""Главное окно приложения для Windows."""

from __future__ import annotations

import os
import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk

from src.logic import run_check
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
    self._file_row(frame, 1, "Файл проверки", self.check, open_file=True, on_change=self._suggest_output_name)
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

  def _build_legend(self, parent: ttk.Frame) -> None:
    frame = ttk.Frame(parent)
    frame.pack(fill="x", pady=(0, 10))

    self._legend_item(frame, COLORS["red_marker"], "Красный — сотрудник не найден в справочнике")
    self._legend_item(frame, COLORS["orange_marker"], "Оранжевый — сотрудник уволен или неактивен")

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
      run_check(
        sprav_file=self.sprav.get().strip(),
        check_file=self.check.get().strip(),
        output_file=output_path,
        check_dismissed=self.dismissed.get(),
        check_duplicates=self.duplicates.get(),
        prompt_duplicate_selection=self.prompt_duplicates.get(),
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
