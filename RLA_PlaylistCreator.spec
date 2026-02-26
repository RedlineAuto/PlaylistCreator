# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_data_files

datas = [('RLA-PlaylistCreator.ico', '.'), ('RA-x-RR.png', '.')]
datas += collect_data_files('customtkinter')


a = Analysis(
    ['RLA_PlaylistCreator_v2.pyw'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=['requests', 'yt_dlp', 'mutagen', 'numpy'],
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
    name='RLA_PlaylistCreator',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['RLA-PlaylistCreator.ico'],
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='RLA_PlaylistCreator',
)
