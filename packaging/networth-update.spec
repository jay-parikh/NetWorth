# PyInstaller spec — one-file console updater for Windows and macOS.
# Build from the repo root:  pyinstaller packaging/networth-update.spec
# Masters come from the workbook itself, but every data/*.csv the updater
# loads at runtime (rates, FMV, curated restructures, bullion proxies) must
# be bundled here; certifi/truststore are picked up by the hooks.

import sys

a = Analysis(
    ["update_portfolio.py"],
    pathex=["../src"],
    binaries=[],
    datas=[("../data/banks_in.csv", "data"),
           ("../data/fmv_2018-01-31.csv", "data"),
           ("../data/ppf_rates.csv", "data"),
           ("../data/epf_rates.csv", "data"),
           ("../data/bullion_proxies.csv", "data"),
           ("../data/restructures.csv", "data")],
    hiddenimports=["truststore", "msoffcrypto", "msoffcrypto.format.ooxml"],
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
