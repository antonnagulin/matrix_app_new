# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec для matrix-check (Windows GUI + tkinter)."""

import sys
from pathlib import Path

from PyInstaller.utils.hooks import collect_all

block_cipher = None


def _tcl_tk_datas():
    """Tcl/Tk data files required by tkinter runtime hook (pyi_rth_tkinter)."""
    tcl_root = Path(sys.base_prefix) / "tcl"
    pairs = []
    mapping = {
        "tcl8.6": "_tcl_data",
        "tk8.6": "_tk_data",
    }
    for folder, dest in mapping.items():
        src = tcl_root / folder
        if src.is_dir():
            pairs.append((str(src), dest))
    if not pairs:
        raise SystemExit(
            f"Не найдены папки Tcl/Tk в {tcl_root}. "
            "Переустановите Python с компонентом tcl/tk."
        )
    return pairs


datas = _tcl_tk_datas()
binaries = []
hiddenimports = [
    "tkinter",
    "tkinter.ttk",
    "tkinter.filedialog",
    "tkinter.messagebox",
    "tkinter.scrolledtext",
]

for package in ("tkinter", "pandas", "openpyxl"):
    pkg_datas, pkg_binaries, pkg_hidden = collect_all(package)
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
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
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
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
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
