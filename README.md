# NetWorth — Family Portfolio Tracker

One Excel workbook to track a family's entire net worth — **Equity, Mutual Funds
(incl. SIPs), Fixed Deposits, PPF and Bonds** — with automatic price/NAV updates,
XIRR (annualised returns), inflation-adjusted projections, and per-person views.

- **100% local.** Your data never leaves your computer. The only network calls
  are direct downloads of public price data (AMFI NAVs, BSE/NSE bhavcopy).
  No cloud, no server, no account, no telemetry.
- **Windows & macOS.** One double-clickable updater per OS. No Python or any
  other install needed to *use* it — just Excel (or LibreOffice) to open the file.
- **Open & portable.** The workbook is generated from code against a written
  specification ([docs/SPEC.md](docs/SPEC.md)), so the whole product can be
  re-implemented on any platform or language from the spec alone.

> Status: **in development.** The current working (Windows-only, PowerShell)
> template lives in [legacy/](legacy/) and is fully usable today — see
> [legacy/README.txt](legacy/README.txt). The cross-platform rewrite is being
> built milestone by milestone; see [docs/RELEASES.md](docs/RELEASES.md).

## What you get

| Sheet | Purpose |
|---|---|
| Dashboard | Family net worth, per-person × per-asset-class matrix, portfolio XIRR, inflation check, allocation charts |
| Projection | 20-year corpus trajectory: your return vs inflation (nominal & real) |
| Per-person sheets | Each family member's allocation + pie chart |
| Equity | Holdings with live prices, day/net change, per-stock XIRR |
| MutualFunds + MF_SIP | Fund summary auto-computed from a one-row-per-purchase SIP ledger |
| FixedDeposits | Value-as-on-today and maturity value from principal/rate/dates |
| PPF | Balance tracking with reference rate |
| Bonds | Corporate/other bonds with maturity value |
| By Scrip | Family-wide exposure per stock, split by person |
| MF_Master / Stock_Master | ~14,000 AMFI schemes and ~4,500 listed stocks powering type-ahead dropdowns |
| Guide | 2-minute in-workbook manual |

Planned for v1 (see [docs/PLAN.md](docs/PLAN.md) and [docs/ROADMAP.md](docs/ROADMAP.md)):
red/green gain-loss colouring, bond maturity & coupon-aware XIRR, Indian bank
dropdown for FDs, expected value at FY-end, FMV 31-01-2018 fallback for unknown
buy prices (LTCG grandfathering), delisted/suspended stock handling, and
automatic corporate-action (split/bonus) adjustment.

## Quick start (current legacy template)

1. Download the [legacy/](legacy/) folder (or the release zip) and keep all
   files together.
2. Open `Family_Portfolio_Tracker.xlsx`, read the **Guide** sheet.
3. Replace the fictional sample rows (Amit/Priya/Rahul) with your own people
   and holdings. Yellow cells are inputs; grey columns compute themselves.
   Pick schemes/scrips from the dropdowns — ISINs fill in automatically.
4. Close the file, then double-click:
   - `UpdatePrices.bat` — refresh share prices (BSE/NSE) + stock master + XIRR
   - `UpdateNAV.bat` — refresh fund NAVs (AMFI) + XIRR
   - `UpdateFundMaster.bat` — refresh the mutual-fund scheme list

   *(Legacy updaters are Windows-only and need desktop Excel. The rewrite
   replaces all three with one cross-platform `Update Portfolio` app.)*

## Quick start (cross-platform rewrite — once released)

1. Download the zip for your OS from **GitHub Releases**.
2. Open `Family_Portfolio_Tracker.xlsx`, enter your holdings, save & close.
3. Double-click **`Update Portfolio`** (`.exe` on Windows, `.command` on Mac).
   It backs up your file, fetches the latest prices/NAVs, recomputes XIRR and
   projections, and rebuilds the workbook in place.

macOS note: the first run of an unsigned app needs *right-click → Open*.

## For developers

Everything is Python (dev-time only; end users never need it):

```
pip install -e ".[dev]"
python -m networth.generate    # build the template workbook from code
python -m networth.update      # run the updater against a workbook
pytest                         # tests incl. XIRR golden values & round-trip identity
```

Key documents:

- [docs/SPEC.md](docs/SPEC.md) — the platform-agnostic specification (sheets,
  data contracts, algorithms). **The spec is the product**; Python here is the
  reference implementation.
- [docs/PLAN.md](docs/PLAN.md) — approved architecture & feature plan.
- [docs/RELEASES.md](docs/RELEASES.md) — milestone/release plan.
- [docs/ROADMAP.md](docs/ROADMAP.md) — backlog & ideas.
- [CLAUDE.md](CLAUDE.md) — working notes/conventions for AI-assisted development.

**For maintainers:** end users edit the workbook freely — entering and updating
holdings in the input cells is the product, and the updater preserves all of it.
What must *not* happen is changing the template's **structure** (sheets, columns,
formulas, charts) by hand-editing the shipped xlsx, or saving it through
openpyxl (which silently destroys the charts). Structural changes belong in
`src/networth/generate.py`; rebuild to apply them.

## Privacy & data sources

All fetches are plain HTTPS GETs of public data, initiated from your machine:

- AMFI daily NAVs: `https://www.amfiindia.com/spages/NAVAll.txt`
- BSE bhavcopy: `https://www.bseindia.com/download/BhavCopy/Equity/...`
- NSE bhavcopy (fallback): `https://nsearchives.nseindia.com/content/cm/...`

Static reference data (bank list, PPF rate history, 31-01-2018 FMV) ships
inside the release — no fetch needed.

## License

[MIT](LICENSE)
