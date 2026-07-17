"""Кроссплатформенные утилиты GUI (Windows / Linux / Astra)."""

from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path


def setup_gui_platform() -> None:
    """Подготовка окружения до создания tk.Tk()."""
    if sys.platform != "win32":
        return

    from src.win_dpi import enable_dpi_awareness

    enable_dpi_awareness()


def configure_gui_root(root) -> None:
    """Настройка корневого окна после tk.Tk()."""
    if sys.platform != "win32":
        return

    from src.win_dpi import configure_tk_scaling

    configure_tk_scaling(root)


def open_folder(folder: str) -> bool:
    """Открыть папку в файловом менеджере. Возвращает True при успехе."""
    path = Path(folder).resolve()
    if not path.is_dir():
        return False

    folder_str = str(path)

    if sys.platform == "win32":
        os.startfile(folder_str)  # noqa: S606
        return True

    if sys.platform == "darwin":
        subprocess.run(["open", folder_str], check=False)
        return True

    for cmd in (
        ["xdg-open", folder_str],
        ["gio", "open", folder_str],
        ["kde-open", folder_str],
    ):
        try:
            result = subprocess.run(cmd, check=False, capture_output=True, timeout=10)
            if result.returncode == 0:
                return True
        except (FileNotFoundError, subprocess.TimeoutExpired, OSError):
            continue

    return False
