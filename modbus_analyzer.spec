# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['modbus_analyzer.py'],
    pathex=[],
    binaries=[],
    datas=[('ui', 'ui'), ('core', 'core'), ('utils', 'utils'), ('plugins', 'plugins'), ('config_and_params.xlsx', '.')],
    hiddenimports=['PyQt5.QtCore', 'PyQt5.QtGui', 'PyQt5.QtWidgets', 'pyqtgraph', 'pandas', 'numpy', 'pyserial', 'openpyxl'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['matplotlib', 'scipy', 'sklearn', 'torch', 'tensorflow', 'PIL', 'cv2'],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='modbus_analyzer',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='modbus_analyzer',
)
