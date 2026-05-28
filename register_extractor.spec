# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec for TVET CDACC Register Extractor
# Build:  pyinstaller register_extractor.spec

import os
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

block_cipher = None

# ── Data files to bundle ──────────────────────────────────────────────────────
datas = []

# Project fonts folder
datas += [("fonts", "fonts")]

# ttkbootstrap themes and localization
datas += collect_data_files("ttkbootstrap")

# pdfplumber / pdfminer cmap data (needed for text extraction)
datas += collect_data_files("pdfminer")

# reportlab fonts and other assets
datas += collect_data_files("reportlab")

# ── Hidden imports ────────────────────────────────────────────────────────────
hiddenimports = []

# ttkbootstrap
hiddenimports += collect_submodules("ttkbootstrap")

# pdfplumber + pdfminer
hiddenimports += collect_submodules("pdfplumber")
hiddenimports += collect_submodules("pdfminer")

# reportlab
hiddenimports += collect_submodules("reportlab")

# openpyxl
hiddenimports += collect_submodules("openpyxl")

# Pillow (used internally by pdfplumber for image pages)
hiddenimports += collect_submodules("PIL")

# Tkinter extras sometimes missed
hiddenimports += [
    "tkinter",
    "tkinter.ttk",
    "tkinter.filedialog",
    "tkinter.messagebox",
    "tkinter.font",
    "_tkinter",
]

# ── Analysis ──────────────────────────────────────────────────────────────────
a = Analysis(
    ["gui.py"],
    pathex=["."],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        "matplotlib", "numpy", "scipy", "pandas",
        "IPython", "jupyter", "notebook",
        "PyQt5", "PyQt6", "wx", "gi",
        "unittest", "test",
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# ── Single-file EXE ───────────────────────────────────────────────────────────
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="Register Extractor",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,           # compress — requires UPX installed; set False if not available
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,      # no black console window
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # icon="icon.ico",  # uncomment and point to a .ico file to set a custom icon
)
