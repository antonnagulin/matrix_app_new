"""Точка входа: запуск Windows-приложения."""

import tkinter as tk

from src.ui import EmployeeCheckApp
from src.win_dpi import enable_dpi_awareness


def main() -> None:
    enable_dpi_awareness()
    root = tk.Tk()
    EmployeeCheckApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
