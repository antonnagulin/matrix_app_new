"""DPI awareness для чёткого отображения tkinter на Windows (в т.ч. PyInstaller exe)."""

from __future__ import annotations

import sys


def enable_dpi_awareness() -> None:
    """Включить per-monitor DPI до создания окон tkinter."""
    if sys.platform != "win32":
        return

    import ctypes

    try:
        user32 = ctypes.windll.user32
        # PerMonitorV2 — Windows 10 1703+
        if user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4)):
            return
    except (AttributeError, OSError):
        pass

    try:
        # PROCESS_PER_MONITOR_DPI_AWARE = 2
        if ctypes.windll.shcore.SetProcessDpiAwareness(2) == 0:
            return
    except (AttributeError, OSError):
        pass

    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except (AttributeError, OSError):
        pass
