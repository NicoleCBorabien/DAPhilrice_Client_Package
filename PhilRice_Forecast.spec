# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path

project = Path(SPECPATH)
a = Analysis(
    ['app.py'],
    pathex=[str(project)],
    binaries=[],
    datas=[(str(project / 'templates'), 'templates'),
           (str(project / 'static'), 'static')],
    hiddenimports=['statsmodels.tsa.arima.model',
                   'statsmodels.tsa.holtwinters',
                   'matplotlib.backends.backend_agg'],
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
    name='PhilRice_Forecast',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
)
