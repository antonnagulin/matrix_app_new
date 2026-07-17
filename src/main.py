"""Точка входа: запуск GUI-приложения."""

import tkinter as tk

from src.platform_utils import configure_gui_root, setup_gui_platform
from src.ui import EmployeeCheckApp


def main() -> None:
    setup_gui_platform()
    root = tk.Tk()
    configure_gui_root(root)
    EmployeeCheckApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
