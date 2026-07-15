# PyInstaller spec — one-file console updater for Windows and macOS.
# Build from the repo root:  pyinstaller packaging/networth-update.spec
# The updater reads masters from the workbook itself, so no data files
# need to be bundled; certifi/truststore are picked up by the hooks.

import sys

a = Analysis(
    ["update_portfolio.py"],
    pathex=["../src"],
    binaries=[],
    datas=[],
    hiddenimports=["truststore"],
    hookspath=[],
    runtime_hooks=[],
    excludes=["tkinter", "test", "unittest"],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="Update Portfolio" if sys.platform == "win32" else "networth-updater",
    debug=False,
    strip=False,
    upx=False,
    console=True,
    disable_windowed_traceback=False,
)
