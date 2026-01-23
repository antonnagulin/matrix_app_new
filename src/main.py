import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from tkinter import ttk
from src.logic import run_check

class FullscreenApp:
    def __init__(self, root):
        self.root = root
        root.title("Проверка сотрудников")

        # ====== Полноэкранный режим ======
        try:
            root.state('zoomed')  # Windows
        except tk.TclError:
            root.attributes('-fullscreen', True)  # Linux/Mac
        self.root.bind("<Escape>", lambda e: self.root.attributes("-fullscreen", False))
        root.configure(bg="#f0f0f0")

        # ====== Переменные ======
        self.sprav = tk.StringVar()
        self.check = tk.StringVar()
        self.output = tk.StringVar(value="check_result.xlsx")
        self.dismissed = tk.BooleanVar(value=True)
        self.duplicates = tk.BooleanVar(value=True)

        self.build_ui()

    def build_ui(self):
        style = ttk.Style()
        style.configure("TLabelFrame", font=("Arial", 12, "bold"))
        style.configure("TLabel", font=("Arial", 11))
        style.configure("TButton", font=("Arial", 11, "bold"))

        # ====== Кнопки запуска и закрытия ======
        btn_frame = tk.Frame(self.root, bg="#f0f0f0")
        btn_frame.pack(fill="x", padx=20, pady=10)

        tk.Button(btn_frame, text="▶ Запустить проверку", font=("Arial", 12, "bold"),
                  bg="#4CAF50", fg="white", command=self.run).pack(side="left", padx=5)
        tk.Button(btn_frame, text="❌", font=("Arial", 12, "bold"),
                  bg="#F44336", fg="white", command=self.root.destroy).pack(side="right", padx=5)

        # ====== Блок выбора файлов ======
        file_frame = ttk.LabelFrame(self.root, text="Выбор файлов")
        file_frame.pack(fill="x", padx=20, pady=10)

        self.row_file(file_frame, "Справочник:", self.sprav, 0)
        self.row_file(file_frame, "Файл проверки:", self.check, 1)
        ttk.Label(file_frame, text="Файл результата:").grid(row=2, column=0, sticky="w", pady=5)
        ttk.Entry(file_frame, textvariable=self.output, width=50).grid(row=2, column=1, sticky="w", pady=5)

        # ====== Блок опций ======
        options_frame = ttk.LabelFrame(self.root, text="Настройки")
        options_frame.pack(fill="x", padx=20, pady=10)

        ttk.Checkbutton(options_frame, text="Проверять уволенных", variable=self.dismissed).grid(row=0, column=0, padx=10, pady=5)
        ttk.Checkbutton(options_frame, text="Искать дубликаты", variable=self.duplicates).grid(row=0, column=1, padx=10, pady=5)

        # ====== Лог выполнения (увеличенный) ======
        log_frame = ttk.LabelFrame(self.root, text="Лог выполнения")
        log_frame.pack(fill="both", expand=True, padx=20, pady=10)

        self.log = scrolledtext.ScrolledText(log_frame, font=("Consolas", 11))
        self.log.pack(fill="both", expand=True)

    def row_file(self, frame, label, var, row):
        ttk.Label(frame, text=label).grid(row=row, column=0, sticky="w", pady=5)
        ttk.Entry(frame, textvariable=var, width=50).grid(row=row, column=1, sticky="w", pady=5)
        tk.Button(frame, text="Обзор", font=("Arial", 10), bg="#2196F3", fg="white",
                  command=lambda: self.pick_file(var)).grid(row=row, column=2, padx=5)

    def pick_file(self, var):
        file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file:
            var.set(file)

    def run(self):
        if not self.sprav.get() or not self.check.get():
            messagebox.showerror("Ошибка", "Выберите файлы")
            return

        self.log.delete(1.0, tk.END)
        try:
            run_check(
                sprav_file=self.sprav.get(),
                check_file=self.check.get(),
                output_file=self.output.get(),
                check_dismissed=self.dismissed.get(),
                check_duplicates=self.duplicates.get(),
                log=self.log
            )
            messagebox.showinfo("Готово", "Проверка завершена!")
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))


if __name__ == "__main__":
    root = tk.Tk()
    app = FullscreenApp(root)
    root.mainloop()
