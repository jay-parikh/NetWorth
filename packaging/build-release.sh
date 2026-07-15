#!/bin/bash
# Build the release zip for the CURRENT OS (macOS or Linux; use
# build-release.bat on Windows). Run from the repo root:
#   packaging/build-release.sh <version>
set -euo pipefail
cd "$(dirname "$0")/.."
VERSION="${1:?usage: packaging/build-release.sh <version>}"

python -m venv .venv-build 2>/dev/null || true
source .venv-build/bin/activate
# Old pip cannot install a pyproject-only project; upgrade first, then a
# regular (non-editable) install — editable is a dev convenience, not needed
# to build a release artifact.
pip -q install --upgrade pip setuptools wheel
pip -q install . pyinstaller

OS="$(uname -s | tr '[:upper:]' '[:lower:]')"
case "$OS" in
  darwin) OSNAME="macos" ;;
  *)      OSNAME="$OS" ;;
esac

rm -rf build dist
python -m networth.generate -o "dist/Family_Portfolio_Tracker.xlsx" 2>/dev/null || {
  mkdir -p dist && python -m networth.generate -o "dist/Family_Portfolio_Tracker.xlsx"; }
( cd packaging && pyinstaller --distpath ../dist --workpath ../build -y networth-update.spec )

STAGE="dist/NetWorth-${VERSION}-${OSNAME}"
mkdir -p "$STAGE"
cp "dist/Family_Portfolio_Tracker.xlsx" "$STAGE/"
cp packaging/README-enduser.txt "$STAGE/README.txt"
if [ "$OSNAME" = "macos" ]; then
  cp dist/networth-updater "$STAGE/"
  cp "packaging/Update Portfolio.command" "$STAGE/"
  chmod +x "$STAGE/Update Portfolio.command" "$STAGE/networth-updater"
else
  cp dist/networth-updater "$STAGE/" 2>/dev/null || cp "dist/Update Portfolio" "$STAGE/"
fi
( cd dist && zip -qr "NetWorth-${VERSION}-${OSNAME}.zip" "NetWorth-${VERSION}-${OSNAME}" )
echo "Built dist/NetWorth-${VERSION}-${OSNAME}.zip"
