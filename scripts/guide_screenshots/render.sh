#!/usr/bin/env bash
# Re-render every docs/USER-GUIDE.md screenshot from the CURRENT sample
# workbook. Run after any template change so the guide never drifts from
# the product (the docs-follow-code rule applied to pictures).
#
# Needs: LibreOffice (a headless install, or an AppImage — see below) and
# `pip install pymupdf` in the project venv. No root required:
#   curl -Lo lo.AppImage https://appimages.libreitalia.org/LibreOffice-fresh.standard-x86_64.AppImage
#   chmod +x lo.AppImage && ./lo.AppImage --appimage-extract
#   export SOFFICE=$PWD/squashfs-root/AppRun
#   export LOPYTHON=$PWD/squashfs-root/opt/libreoffice*/program/python
# (First AppImage start exits with code 81 after creating its profile —
# just start it again.)
set -euo pipefail

HERE="$(cd "$(dirname "$0")" && pwd)"
REPO="$(cd "$HERE/../.." && pwd)"
WORK="${WORK:-$(mktemp -d)}"
SOFFICE="${SOFFICE:-soffice}"
LOPYTHON="${LOPYTHON:-python3}"        # must have `uno` importable
PY="$REPO/.venv/bin/python"
PORT=2002

echo "work dir: $WORK"
mkdir -p "$WORK/pdfs" "$WORK/pdfs-masked" "$WORK/png"

# 1. build the two workbook variants at a PINNED date (deterministic shots)
"$PY" - "$WORK" << 'EOF'
import sys
from datetime import date
from networth.generate import build_workbook
from networth.sample_data import sample_portfolio
from networth import crypto

out = sys.argv[1]
today = date(2026, 7, 18)              # bump alongside a sample-data refresh
d = sample_portfolio()
d.show_capital_gains = True
build_workbook(d, f"{out}/guide-main.xlsx", today=today)
m = sample_portfolio()
m.privacy_enabled = True
m.privacy_hash = crypto.hash_password("demo-password")
build_workbook(m, f"{out}/guide-masked.xlsx", masked=True, today=today)
print("workbooks built")
EOF

# sheet password of the masked build (unprotects sheets so comments can be
# stripped before export)
SHPW=$("$PY" - "$WORK" << 'EOF'
import sys
from openpyxl import load_workbook
from networth import crypto
wb = load_workbook(f"{sys.argv[1]}/guide-masked.xlsx")
print(crypto.sheet_password(wb.defined_names["NW_Privacy"].value.strip('="')))
EOF
)

# 2. headless LibreOffice listener
"$SOFFICE" --headless --norestore --nologo \
  -env:UserInstallation=file://"$WORK"/louser \
  --accept="socket,host=localhost,port=$PORT;urp;" > /dev/null 2>&1 &
sleep 10

# 3. per-sheet PDF exports (SheetName[:RANGE]; ~ = space in a sheet name)
"$LOPYTHON" "$HERE/uno_export.py" "$WORK/guide-main.xlsx" "$WORK/pdfs" \
  "Dashboard,Projection,Settings,Amit:A1:P32,Equity:A1:P14,Equity_Sells:A1:K8,MutualFunds:A1:M8,MF_SIP:A1:K12,FixedDeposits:A1:N10,PPF:A1:L10,Bonds:A1:N10,Gold_Silver:A1:O10,NPS:A1:K10,Manual_Assets:A1:L12,Dividends:A1:S36,Capital~Gains,Tax_Rules:A1:G12,Guide"
"$LOPYTHON" "$HERE/uno_export.py" "$WORK/guide-masked.xlsx" \
  "$WORK/pdfs-masked" "Equity:A1:P14" "$SHPW"

# 4. rasterize + auto-crop, then place with the guide's names
"$PY" "$HERE/rasterize.py" "$WORK/pdfs" "$WORK/png"
"$PY" "$HERE/rasterize.py" "$WORK/pdfs-masked" "$WORK/png" masked-
IMG="$REPO/docs/images"
mkdir -p "$IMG"
for f in dashboard settings equity equity-sells capital-gains mutualfunds \
         mf-sip fixeddeposits ppf bonds gold-silver nps manual-assets \
         dividends projection tax-rules masked-equity; do
  cp "$WORK/png/$f.png" "$IMG/$f.png"
done
cp "$WORK/png/amit.png" "$IMG/person-amit.png"
cp "$WORK/png/guide.png" "$IMG/guide-tab.png"

pkill -f "soffice.bin.*port=$PORT" || true
echo "done — $(ls "$IMG" | wc -l) images in docs/images/"
