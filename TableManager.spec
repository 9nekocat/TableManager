# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec for TableManager.
Builds a single .exe with Flask server + templates.
"""

a = Analysis(
    ['src/table_server.py'],
    pathex=['src'],
    binaries=[],
    datas=[
        ('gui/dist', 'gui/dist'),
        ('data', 'data'),
        ('src/diff_standalone.py', '.'),
        ('dist/diff.exe', '.'),
        ('art', 'art'),
        ('env', 'env'),
    ],
    hiddenimports=[
        'flask',
        'openpyxl',
        'svn_client',
        'merge_engine',
        'merge_workflow',
        'formula_utils',
        'manager_core',
        'diff_engine',
        'excel_recalc',
        'effect_recalc',
        'effect_diag',
        'toolbox',
        'ref_index',
        'table_checks',
        'processed_checks',
        'win32com',
        'win32com.client',
        'pythoncom',
        'webview',
        'webview.platforms.winforms',
        'webview.guilib',
        'clr',
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
    name='TableManager',
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
    icon='art/exe_icon.ico',
)
