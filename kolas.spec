# -*- mode: python ; coding: utf-8 -*-

a = Analysis(
    ['dd.py'],
    pathex=[],
    binaries=[],
    datas=[('template.xlsx', '.')],  # template.xlsx를 exe와 같은 폴더에 포함
    hiddenimports=['selenium', 'openpyxl', 'xlrd', 'webdriver_manager', 'dotenv'],
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
    name='KOLAS자동화',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,   # 터미널 창 표시 (로그 확인용)
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
