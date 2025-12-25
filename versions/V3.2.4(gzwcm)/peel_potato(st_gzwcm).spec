# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['peel_potato(st_gzwcm).py'],
    pathex=[],
    binaries=[],
    # Include UI/help assets so they are packaged into the EXE directory
    datas=[
        ('media/icon_app.ico', 'media'),
        ('media/icon_exe.ico', 'media'),
        ('media/help_st_gzwcm.html', 'media'),
        ('data/emp_embed.xlsx', 'data'),
        ('data/dict_embed.xlsx', 'data'),
    ],
    hiddenimports=['PyQt6', 'PyQt6.QtCore', 'PyQt6.QtGui', 'PyQt6.QtWidgets'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib', 'scipy', 'IPython', 'jupyter', 'notebook',
        'tornado', 'zmq', 'PIL.ImageTk', 'tkinter', 'unittest',
        'test', 'tests', 'PyQt5', 'PySide2', 'PySide6'
    ],
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
    name='peel_potato(gzw_V3.2.3)',
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
