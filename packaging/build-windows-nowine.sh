#!/bin/bash
# Build the Windows deliverable on Linux WITHOUT Wine and WITHOUT compiling.
#
# This project is pure Python, so we assemble a real Windows-runnable app by
# downloading prebuilt Windows artifacts and arranging them:
#   * the official Windows EMBEDDABLE CPython (a genuine Windows python.exe)
#   * prebuilt Windows wheels for every dependency (pip --platform win_amd64;
#     even the one C-extension dep, charset-normalizer, ships a win_amd64 wheel)
#   * our own pure-Python wheel + the bundled data CSVs
#   * a launcher (.bat, guaranteed; plus a best-effort .exe)
#
# Run from the repo root:  packaging/build-windows-nowine.sh <version> [pyver]
set -euo pipefail
cd "$(dirname "$0")/.."
VERSION="${1:?usage: packaging/build-windows-nowine.sh <version> [pyver]}"
PYVER="${2:-3.12.7}"
PYSHORT="$(echo "$PYVER" | awk -F. '{print $1$2}')"          # 3.12.7 -> 312
CACHE="$HOME/.cache/networth-winbuild"
EMBED_ZIP="python-${PYVER}-embed-amd64.zip"
mkdir -p "$CACHE"

# host venv just for the build tooling (pip download/wheel + workbook generate)
python3 -m venv .venv-nowine 2>/dev/null || true
source .venv-nowine/bin/activate
pip -q install --upgrade pip wheel distlib >/dev/null
pip -q install -e . >/dev/null          # editable → host generate finds repo data/

STAGE="dist/NetWorth-${VERSION}-windows"
rm -rf "$STAGE" "dist/NetWorth-${VERSION}-windows.zip"
mkdir -p "$STAGE/app"

# 1. Windows embeddable CPython
if [ ! -f "$CACHE/$EMBED_ZIP" ]; then
  echo "==> downloading $EMBED_ZIP"
  curl -fsSL -o "$CACHE/$EMBED_ZIP" \
    "https://www.python.org/ftp/python/${PYVER}/${EMBED_ZIP}"
fi
mkdir -p "$STAGE/app/python"
unzip -q -o "$CACHE/$EMBED_ZIP" -d "$STAGE/app/python"

# make site-packages importable: rewrite the ._pth
cat > "$STAGE/app/python/python${PYSHORT}._pth" <<PTH
python${PYSHORT}.zip
.
Lib\\site-packages
import site
PTH

# 2. our wheel + Windows dependency wheels, unpacked into site-packages
SITE="$STAGE/app/python/Lib/site-packages"
mkdir -p "$SITE"
WHEELDIR="$(mktemp -d)"
pip -q wheel . --no-deps -w "$WHEELDIR" >/dev/null
NW_WHEEL="$(ls "$WHEELDIR"/networth-*.whl)"
pip -q install --target "$SITE" \
  --platform win_amd64 --python-version "$PYVER" \
  --implementation cp --only-binary=:all: --upgrade \
  "$NW_WHEEL" requests xlsxwriter openpyxl truststore >/dev/null

# 3. bundled data where the non-frozen DATA_DIR looks (…/Lib/data)
cp -r data "$STAGE/app/python/Lib/data"

# 4. the workbook (platform-neutral; built with the host python)
python -m networth.generate -o "$STAGE/Family_Portfolio_Tracker.xlsx" >/dev/null

# 5. launcher — a .bat: correct-by-inspection and relocatable (%~dp0), the
#    honest choice for a build that can't be execute-tested from Linux
cat > "$STAGE/Update Portfolio.bat" <<'BAT'
@echo off
REM NetWorth updater (no-Wine bundle). Double-click to run.
setlocal
cd /d "%~dp0"
"%~dp0app\python\python.exe" -m networth._packaged %*
BAT

cp packaging/README-nowine.txt "$STAGE/README.txt"

# 6. zip it (python's zipfile — no external 'zip' needed)
( cd dist && python -m zipfile -c "NetWorth-${VERSION}-windows.zip" \
    "NetWorth-${VERSION}-windows" )
echo "Built dist/NetWorth-${VERSION}-windows.zip (no Wine, no compiler)"
