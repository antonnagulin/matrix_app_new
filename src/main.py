"""Точка входа: запуск Windows-приложения."""

import tkinter as tk

from src.ui import EmployeeCheckApp


def main() -> None:
    root = tk.Tk()
    EmployeeCheckApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
