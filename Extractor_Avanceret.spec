# -*- mode: python ; coding: utf-8 -*-
# Bygger den fulde/avancerede udgave af Extractor (med Excel-kommentarer,
# traadgruppering, intern_tjek m.v.) som en separat exe ved siden af den
# enklere Extractor.exe fra 17-6-2026, som nogle kolleger foretraekker.


a = Analysis(
    ['src/extractor_gui.py'],
    pathex=['src'],
    binaries=[],
    datas=[],
    hiddenimports=['krydstjek', 'semantik', 'intern_tjek'],
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
    name='Extractor_Avanceret',
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
