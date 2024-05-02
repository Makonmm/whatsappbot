# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_data_files
from PyInstaller import __version__ as pyinstaller_version
from PyInstaller.building.api import PYZ, EXE
from PyInstaller.building.build_main import Analysis

a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    datas=[('clientes.xlsx', '.')],  # Inclua aqui seus arquivos de dados necessários
    hiddenimports=[],
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
    [],
    exclude_binaries=True,
    name='app',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='app',
)

# Nenhuma mudança necessária na parte COLLECT

