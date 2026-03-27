# -*- mode: python ; coding: utf-8 -*-
import os
import importlib.util

a = Analysis(
    ['main.py'],
    pathex=[os.getcwd()],
    binaries=[],
    datas=[('icon.png', '.')],
    hiddenimports=['prodotto_dialog', 'prodotti_manager', 'email_utils', 'import_utils', 'pdf_export', 'db'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes= [],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure, a.zipped_data)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='CatalogoApp',
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
    icon=['icon.png'],
    )
