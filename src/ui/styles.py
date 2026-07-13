"""Настройка внешнего вида под Windows."""

import sys
import tkinter as tk
from tkinter import ttk


# Цвета в духе Fluent / Windows 11
COLORS = {
    "bg": "#f3f3f3",
    "surface": "#ffffff",
    "border": "#e1e1e1",
    "text": "#1a1a1a",
    "text_muted": "#5c5c5c",
    "accent": "#0078d4",
    "accent_hover": "#106ebe",
    "success": "#107c10",
    "danger": "#d13438",
    "red_marker": "#ff9999",
    "orange_marker": "#ff7500",
    "header_bg": "#0078d4",
    "header_fg": "#ffffff",
}

FONT_FAMILY = "Segoe UI" if sys.platform == "win32" else "Arial"


def setup_styles(root: tk.Tk) -> ttk.Style:
    style = ttk.Style(root)

    if sys.platform == "win32":
        for theme in ("vista", "xpnative", "clam"):
            if theme in style.theme_names():
                style.theme_use(theme)
                break

    style.configure(".", font=(FONT_FAMILY, 10))
    style.configure("TFrame", background=COLORS["bg"])
    style.configure("Surface.TFrame", background=COLORS["surface"])
    style.configure("TLabel", background=COLORS["bg"], foreground=COLORS["text"])
    style.configure("Surface.TLabel", background=COLORS["surface"], foreground=COLORS["text"])
    style.configure("Muted.TLabel", background=COLORS["surface"], foreground=COLORS["text_muted"])
    style.configure("Header.TLabel", background=COLORS["header_bg"], foreground=COLORS["header_fg"])
    style.configure("HeaderSub.TLabel", background=COLORS["header_bg"], foreground="#deecf9")
    style.configure("TLabelframe", background=COLORS["bg"], foreground=COLORS["text"])
    style.configure("TLabelframe.Label", font=(FONT_FAMILY, 10, "bold"))
    style.configure("TCheckbutton", background=COLORS["bg"])
    style.configure("TEntry", fieldbackground=COLORS["surface"])
    style.configure("Accent.TButton", font=(FONT_FAMILY, 10, "bold"))
    style.configure("Status.TLabel", background="#e9e9e9", foreground=COLORS["text_muted"], padding=(8, 4))

    return style
