import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from src.logic import run_check  

from src.settings.added_person import (
    CHAINGE_FIO_FILE, FIXED_FILE1,
    OUTPUT_FOLDER, SPRAV_FILE, TARGET_FOLDER
)

# ==== Функции работы с данными ====
def normalize_fio(series):
    return series.dropna().astype(str).str.strip()

def get_columns(file_path):
    df = pd.read_excel(file_path)
    return list(df.columns)

def task_1(file_path, target_column):
    df = pd.read_excel(file_path)
    filtered = df[df[target_column].fillna(0) == 1.0]
    return filtered[['ФИО']]

def task_2(file_path):
    df = pd.read_excel(file_path)
    return df, df[['ФИО']]

def task_3(list_1, list_2):
    set_1 = set(normalize_fio(list_1['ФИО']))
    set_2 = set(normalize_fio(list_2['ФИО']))
    return set_1 - set_2

def task_4(missing_names, file_path, target_folder):
    if not missing_names:
        return None
    df = pd.read_excel(file_path)
    new_rows = pd.DataFrame({'ФИО': list(missing_names)})
    updated_df = pd.concat([df, new_rows], ignore_index=True)
    os.makedirs(target_folder, exist_ok=True)
    file_name = os.path.basename(file_path)
    new_file_path = os.path.join(target_folder, file_name)
    updated_df.to_excel(new_file_path, index=False)
    return new_file_path  # <-- возвращаем строку пути


# ==== GUI ====
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Добавление сотрудников в Excel")

        self.file2_path = tk.StringVar()
        self.selected_column = tk.StringVar()

        # Метки и поля
        tk.Label(root, text=f"Первый файл (фиксированный)").grid(row=0, column=0, sticky="w")
        tk.Label(root, text=FIXED_FILE1).grid(row=0, column=1, sticky="w")

        tk.Label(root, text="Файл 2 (для добавления)").grid(row=1, column=0, sticky="w")
        tk.Entry(root, textvariable=self.file2_path, width=50).grid(row=1, column=1)
        tk.Button(root, text="Выбрать", command=self.select_file2).grid(row=1, column=2)

        tk.Label(root, text="Столбец фильтра (значение 1)").grid(row=2, column=0, sticky="w")
        self.column_combo = ttk.Combobox(root, textvariable=self.selected_column, values=[], state="readonly")
        self.column_combo.grid(row=2, column=1, sticky="w")

        tk.Label(root, text="Папка для сохранения").grid(row=3, column=0, sticky="w")
        tk.Entry(root, width=50, state="readonly", readonlybackground="#f0f0f0",
                 textvariable=tk.StringVar(value=TARGET_FOLDER)).grid(row=3, column=1, sticky="w")

        tk.Button(root, text="Выполнить", command=self.run).grid(row=4, column=1, pady=10)

        self.log = tk.Text(root, width=80, height=20)
        self.log.grid(row=5, column=0, columnspan=3, padx=10, pady=10)

        # Получаем список столбцов из первого файла
        if os.path.exists(FIXED_FILE1):
            columns = get_columns(FIXED_FILE1)
            self.column_combo['values'] = columns
            if columns:
                self.column_combo.current(0)

    # --- Выбор второго файла ---
    def select_file2(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files","*.xlsx")])
        if path:
            self.file2_path.set(path)

    # --- Основная логика ---
    def run(self):
        self.log.delete("1.0", tk.END)
        file2 = self.file2_path.get()
        col = self.selected_column.get()
        folder = TARGET_FOLDER

        if not file2 or not col:
            messagebox.showerror("Ошибка", "Выберите файл 2 и столбец")
            return

        try:
            # --- Сотрудники со значением 1 ---
            list1 = task_1(FIXED_FILE1, col)
            self.log.insert(tk.END, f"Сотрудники со значением 1 ({col}):\n")
            if list1.empty:
                self.log.insert(tk.END, "Нет сотрудников со значением 1\n")
            else:
                for name in list1['ФИО']:
                    self.log.insert(tk.END, f"{name}\n")

            # --- Второй файл ---
            full_df2, list2 = task_2(file2)
            self.log.insert(tk.END, "\nВсе сотрудники второго файла:\n")
            for name in list2['ФИО']:
                self.log.insert(tk.END, f"{name}\n")

            # --- Отсутствующие ---
            missing = task_3(list1, list2)
            self.log.insert(tk.END, "\nОтсутствующие во втором файле:\n")
            if missing:
                for name in sorted(missing):
                    self.log.insert(tk.END, f"{name}\n")
            else:
                self.log.insert(tk.END, "Отсутствующих нет\n")

            # --- Сохранение ---
            saved_file = task_4(missing, file2, folder)
            if saved_file:
                self.log.insert(tk.END, f"\nФайл сохранён: {saved_file}\n")
            else:
                self.log.insert(tk.END, "\nНечего добавлять\n")
                self.root.update()
                self.root.destroy()
                return  # нет смысла запускать проверку

            # --- Закрываем GUI ---
            self.root.update()
            self.root.destroy()

            # --- Линейный запуск run_check ---
            file_name = os.path.basename(file2)
            output_file = os.path.join(OUTPUT_FOLDER, file_name)
            try:
                run_check(
                    sprav_file=SPRAV_FILE,
                    check_file=saved_file,
                    output_file=output_file,
                    chainge_fio_file=CHAINGE_FIO_FILE,
                    log=None  
                )
                print("\nПроверка run_check завершена.")
            except Exception as e:
                print(f"\nОшибка при запуске run_check: {e}")

        except Exception as e:
            messagebox.showerror("Ошибка", str(e))
            self.log.insert(tk.END, f"\nОшибка: {e}\n")


# ==== Запуск GUI ====
if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()