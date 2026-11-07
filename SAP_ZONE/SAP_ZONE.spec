# -*- mode: python ; coding: utf-8 -*-

datas = [('zone_map.json', '.')]
binaries = []
hiddenimports = [
    'win32gui',
    'win32con',
    'win32clipboard',
    'win32process',
    'win32print',
    'pythoncom',
    'win32com',
    'win32com.client',
    'sqlite3',
    'openpyxl'
]
excludes = [
    'webview',
    'pythonnet',
    'clr_loader',
    'numpy',
    'matplotlib',
    'scipy',
    'pandas',
    'IPython',
    'PIL',
    'pytest',
    'test',
    'unittest'
]

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excludes,
    noarchive=False,
    optimize=2,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='SAP_ZONE',
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
