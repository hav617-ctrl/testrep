# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec for analyze_dialogue.exe
# Build: python -m PyInstaller analyze_dialogue.spec --clean

a = Analysis(
    ['analyze_dialogue.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        'openpyxl',
        'et_xmlfile',
        'anthropic',
        'anthropic._models',
        'anthropic.types',
        'anthropic._client',
        'httpx',
        'httpcore',
        'anyio',
        'sniffio',
        'certifi',
        'tkinter',
        'tkinter.filedialog',
        'tkinter.messagebox',
        'tkinter.scrolledtext',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='analyze_dialogue',
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
)
