# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['peel_potato(main).py'],
    pathex=[],
    binaries=[],
    # Include UI/help assets so they are packaged into the EXE directory
    datas=[
        ('media/icon_app.ico', 'media'),
        ('media/icon_exe.ico', 'media'),
        ('media/help.html', 'media'),
    ],
    hiddenimports=[
        'PyQt6',
        'PyQt6.QtCore',
        'PyQt6.QtGui',
        'PyQt6.QtWidgets',
        'win32com',
        'win32com.client',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['matplotlib', 'scipy', 'IPython', 'jupyter', 'notebook', 'tornado', 'zmq', 'PIL.ImageTk', 'tkinter', 'unittest', 'test', 'tests', 'PyQt5', 'PySide2', 'PySide6'],
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
    strip=True,
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
