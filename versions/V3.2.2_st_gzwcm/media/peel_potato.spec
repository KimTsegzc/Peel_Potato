# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['peel_potato.py'],
    pathex=[],
    binaries=[],
    # Include UI/help assets so they are packaged into the EXE directory
    datas=[
        ('media/icon_app.ico', 'media'),
        ('media/icon_exe.ico', 'media'),
        ('media/help.html', 'media'),
        ('data/emp.xlsx', 'data'),
        ('data/dict.xlsx', 'data'),
    ],
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
    name='peel_potato',
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
    icon='media/icon_exe.ico',
)
