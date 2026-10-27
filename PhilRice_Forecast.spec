# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path
from PyInstaller.utils.hooks import collect_all

project = Path(SPECPATH)

datas = [
    (str(project / "templates"), "templates"),
    (str(project / "static"), "static"),
]

binaries = []
hiddenimports = []

for package in ["matplotlib", "statsmodels", "scipy", "PIL"]:
    package_datas, package_binaries, package_hiddenimports = collect_all(package)
    datas.extend(package_datas)
    binaries.extend(package_binaries)
    hiddenimports.extend(package_hiddenimports)

a = Analysis(
    ["app.py"],
    pathex=[str(project)],
    binaries=binaries,
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
    upx=False,
    console=False,
    disable_windowed_traceback=False,
)
