# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path

project_root = Path(SPECPATH)
src_dir = project_root / "src"

datas = [
    (str(src_dir / "templates"), "templates"),
    (str(src_dir / "static"), "static"),
]

hiddenimports = [
    "statsmodels.tsa.arima.model",
    "statsmodels.tsa.holtwinters",
    "matplotlib.backends.backend_agg",
]

a = Analysis(
    [str(src_dir / "app.py")],
    pathex=[str(src_dir)],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
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
    name="PhilRice_Forecast",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
)
