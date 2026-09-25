# -*- mode: python ; coding: utf-8 -*-
# Reduceret udgave til kolleger paa RMC2025N00400/Movia-projektet: samme
# UDTRAEK- og kravmatrix-funktioner som Extractor.exe, men uden
# inputfelt-udtraek, krydstjek og interne henvisninger.

a = Analysis(
    ['src/extractor_gui_rmc.py'],
    pathex=['src'],
    binaries=[],
    datas=[],
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
    a.binaries,
    a.datas,
    [],
    name='RMC_Procurement_Extractor',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico',
)
