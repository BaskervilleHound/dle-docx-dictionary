# -*- mode: python ; coding: utf-8 -*-

import os


playwright_browsers = os.path.join(os.environ['LOCALAPPDATA'], 'ms-playwright')
playwright_datas = [
    (
        os.path.join(playwright_browsers, name),
        os.path.join('ms-playwright', name),
    )
    for name in os.listdir(playwright_browsers)
    if name == '.links' or name.startswith('chromium')
]


a = Analysis(
    ['dlex.py'],
    pathex=[],
    binaries=[],
    datas=playwright_datas,
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
    name='DLEX-portable',
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
