"""DPI awareness для чёткого отображения tkinter на Windows (в т.ч. PyInstaller exe)."""

from __future__ import annotations

import sys

if sys.platform == "win32":
    import ctypes
    from ctypes import wintypes


def enable_dpi_awareness() -> None:
    """Включить DPI awareness до создания окон tkinter.

    Для tkinter надёжнее system-DPI, чем PerMonitorV2: иначе geometry()
    может не применяться и окно остаётся высотой в одну шапку.
    """
    if sys.platform != "win32":
        return

    try:
        # PROCESS_SYSTEM_DPI_AWARE = 1
        if ctypes.windll.shcore.SetProcessDpiAwareness(1) == 0:
            return
    except (AttributeError, OSError):
        pass

    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except (AttributeError, OSError):
        pass


def configure_tk_scaling(root) -> None:
    """Синхронизировать tk scaling с DPI монитора."""
    if sys.platform != "win32":
        return

    try:
        hwnd = wintypes.HWND(root.winfo_id())
        dpi = ctypes.windll.user32.GetDpiForWindow(hwnd)
        if dpi <= 0:
            dpi = 96
        root.tk.call("tk", "scaling", dpi / 72.0)
    except Exception:
        pass
