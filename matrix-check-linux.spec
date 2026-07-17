# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec для matrix-check (Linux / Astra, GUI + tkinter)."""

import importlib.util
import sys
from pathlib import Path

from PyInstaller.utils.hooks import collect_all

block_cipher = None


def _tcl_tk_datas():
    """Tcl/Tk data files required by tkinter runtime hook."""
    pairs = []

    if sys.platform == "win32":
        tcl_root = Path(sys.base_prefix) / "tcl"
        mapping = {"tcl8.6": "_tcl_data", "tk8.6": "_tk_data"}
        for folder, dest in mapping.items():
            src = tcl_root / folder
            if src.is_dir():
                pairs.append((str(src), dest))
    else:
        try:
            import tkinter as tk

            root = tk.Tk()
            root.withdraw()
            tcl_lib = root.tk.eval("set tcl_library")
            tk_lib = root.tk.eval("set tk_library")
            root.destroy()
            pairs = [(tcl_lib, "_tcl_data"), (tk_lib, "_tk_data")]
        except Exception:
            for src, dest in (
                ("/usr/share/tcltk/tcl8.6", "_tcl_data"),
                ("/usr/share/tcltk/tk8.6", "_tk_data"),
                ("/usr/lib/tcl8.6", "_tcl_data"),
                ("/usr/lib/tk8.6", "_tk_data"),
            ):
                if Path(src).is_dir():
                    pairs.append((src, dest))

    if not pairs:
        raise SystemExit(
            "Не найдены библиотеки Tcl/Tk. Установите пакет python3-tk "
            "(на Astra: sudo apt install python3-tk tcl tk)."
        )
    return pairs


def _collect_package(package: str):
    if importlib.util.find_spec(package) is None:
        raise SystemExit(
            f"Модуль '{package}' не установлен. Сначала выполните: poetry install"
        )
    pkg_datas, pkg_binaries, pkg_hidden = collect_all(package)
    return pkg_datas, pkg_binaries, pkg_hidden


datas = _tcl_tk_datas()
binaries = []
hiddenimports = [
    "tkinter",
    "tkinter.ttk",
    "tkinter.filedialog",
    "tkinter.messagebox",
    "tkinter.scrolledtext",
    "pandas",
    "openpyxl",
]

for package in ("tkinter", "pandas", "openpyxl"):
    pkg_datas, pkg_binaries, pkg_hidden = _collect_package(package)
    datas += pkg_datas
    binaries += pkg_binaries
    hiddenimports += pkg_hidden

a = Analysis(
    ["src/main.py"],
    pathex=["."],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["pytest"],
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="matrix-check",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="matrix-check",
)
