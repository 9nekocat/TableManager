# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec for standalone diff.exe.
Lightweight — no Flask, no templates, no web server.
"""

a = Analysis(
    ['src/diff_standalone.py'],
    pathex=['src'],
    binaries=[],
    datas=[],
    hiddenimports=[
        'openpyxl',
        'merge_engine',
        'formula_utils',
        'manager_core',
        'diff_engine',
        'excel_recalc',
        'win32com',
        'win32com.client',
        'pythoncom',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='diff',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
)
