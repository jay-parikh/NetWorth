# NetWorth — Family Portfolio Tracker

One Excel workbook to track a family's entire net worth — **Equity, Mutual Funds
(incl. SIPs), Fixed Deposits, PPF and Bonds** — refreshed by one double-click:
live prices and NAVs, true XIRR (annualised returns), automatic split/bonus
handling, inflation-adjusted projections, and per-person views.

- **100% local.** Your data never leaves your computer. The only network calls
  are direct downloads of public market data (AMFI, BSE, NSE). No cloud, no
  server, no account, no telemetry.
- **Windows & macOS.** One double-clickable updater per OS. No Python or any
  other install needed to *use* it — just Excel (or LibreOffice) to open the file.
- **Open & portable.** The workbook is generated from code against a written
  specification ([docs/SPEC.md](docs/SPEC.md)), so the whole product can be
  re-implemented on any platform or language from the spec alone.

> Status: **v1.1 feature-complete** — v1.0 (R0–R7) plus the v1.1 milestone
> (accurate PPF interest via an optional deposit ledger, net-worth history +
> trend chart, auto-update check). 80 tests, live-verified against real
> BSE/AMFI/NSE data. See [docs/RELEASES.md](docs/RELEASES.md).

## The workbook

| Sheet | Purpose |
|---|---|
| Dashboard | Family net worth, per-person × per-asset-class matrix, portfolio & per-class XIRR, inflation check (real return), expected value at FY-end, allocation pie + per-person bar chart |
| Projection | 20-year corpus trajectory: your return vs inflation, nominal & real (live line chart) |
| One sheet per person | That member's allocation, pie chart, and their holdings across all five classes |
| Equity | Holdings with live BSE/NSE prices, day/net change, per-stock XIRR, split/bonus adjustment factor |
| MutualFunds + MF_SIP | Fund summary auto-computed from a one-row-per-purchase SIP ledger (redemption = negative amount) |
| FixedDeposits | Bank dropdown, value-as-on-today and maturity value from principal/rate/dates/compounding |
| PPF | Balance tracking with reference rate |
| Bonds | Invested/current value, **maturity value** and remaining coupons; coupon-aware XIRR |
| By Scrip | Family-wide exposure per stock, split by person |
| Corporate_Actions | Audit trail of every split/bonus/consolidation fetched for your holdings + manual rows for feed gaps |
| MF_Master / Stock_Master / Bank_Master | ~14,000 AMFI schemes, ~4,500 listed stocks and 60 Indian banks powering the dropdowns; stock status (Active/Suspended/Delisted) per held ISIN |
| Guide | The in-workbook manual, including a What's-New section |

**Colour language:** blue/yellow = you type here, grey = computed. **Green/red**
on every gain, return and XIRR figure. **Amber** = look closer: a stale price
(>7 days), a suspended/delisted stock, or an estimated cost.

**Smart behaviours** (see [docs/SPEC.md](docs/SPEC.md) for exact rules):

- **Corporate actions** — splits, bonuses and consolidations are fetched from
  **both NSE and BSE** for your holdings on every update, deduplicated, and
  applied automatically, including future ex-dates when they arrive. Your
  typed quantities and costs are never rewritten; an *Adj factor* column
  applies the multiplier at valuation time, and the Corporate_Actions sheet
  shows exactly what was applied and why. If any holding cannot be verified
  on either exchange, the run summary says so by name — nothing is ever
  skipped silently.
- **FMV 31-01-2018 fallback** — bought before Feb 2018 and don't know the
  price? Leave *Avg. cost* blank: the LTCG-grandfathering fair market value
  (bundled, 1,639 ISINs, with symbol fallback for reissued ISINs) fills in,
  amber-flagged with an explanatory comment.
- **Delisted/suspended detection** — a held stock absent from the bhavcopy for
  21+ days is marked Suspended, 180+ days Delisted, keeping its last traded
  price/date; you can always type a price manually and it persists.
- **Type-ahead dropdowns** — type the first letters, press Enter, re-open the
  dropdown: only matching schemes/stocks/banks remain (14k schemes stay
  usable). Current Microsoft 365 additionally filters the open list live as
  you type.
- **XIRR everywhere** — portfolio, per class, per fund and per stock, computed
  from dated cashflows (SIP ledger, FD start dates, bond coupons) exactly like
  Excel's `XIRR`.

## Quick start

1. Download the zip for your OS from **GitHub Releases** and extract it
   anywhere; keep the files together.
2. Open `Family_Portfolio_Tracker.xlsx` and read its **Guide** sheet
   (2 minutes). Replace the fictional sample data (Amit/Priya/Rahul) with
   your people and holdings — blue/yellow cells are inputs, dropdowns fill
   the ISINs.
3. Save, **close the file**, then double-click **`Update Portfolio`**
   (`.exe` on Windows; on macOS right-click `Update Portfolio.command` →
   Open the first time).

Each run: offers to **add a new person** (just type a name — their sheet is
created for you) → backs up your file (`backups/`, last 10 kept) → fetches
prices, NAVs, masters and corporate actions → recomputes every XIRR, the
FY-end estimate and PPF interest → records a net-worth snapshot → rebuilds the
workbook in place and prints a one-screen summary. Failed sources degrade
gracefully (old values stay, a warning tells you).

First-run OS warnings (binaries are unsigned for now): Windows SmartScreen →
*More info → Run anyway*; macOS Gatekeeper → *right-click → Open*.

**Upgrading**: your workbook is yours — to upgrade the app, just replace the
updater binary. The next run regenerates the workbook to the newest layout
with all your data preserved.

*(The original Windows-only PowerShell template that this project grew from
is preserved in [legacy/](legacy/) and still works — see
[legacy/README.txt](legacy/README.txt).)*

## Developer guide — build & run on Windows and macOS

Python is the dev-time toolchain only; end users never need it. Everything
below works identically on Windows, macOS and Linux unless a per-OS block
says otherwise.

### Prerequisites

| | Windows | macOS |
|---|---|---|
| Python **3.10+** | [python.org installer](https://www.python.org/downloads/) — defaults are fine (the `py` launcher is included; "Add to PATH" not required) | `python3` from [python.org](https://www.python.org/downloads/) or `brew install python` (the Xcode/CLT system Python 3.9 is too old) |
| Git | [git-scm.com](https://git-scm.com/) or GitHub Desktop | `xcode-select --install` or `brew install git` |
| Internet | needed once for `pip install`, and by the updater for market data | same |

No compilers, no Excel, no admin rights needed to develop. Verify:
`py -3 --version` (Windows) / `python3 --version` (macOS) prints ≥ 3.10.

### Set up (once)

**Windows** (Command Prompt or PowerShell, inside the repo folder):

```bat
py -3 -m venv .venv
.venv\Scripts\activate
pip install -e ".[dev]"
```

**macOS / Linux** (Terminal, inside the repo folder):

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -e ".[dev]"
```

`-e` is an editable install: code changes take effect without reinstalling.
Re-activate the venv (`.venv\Scripts\activate` / `source .venv/bin/activate`)
whenever you open a new shell.

### Everyday commands (venv active, any OS)

```bash
python -m networth.generate                 # build Family_Portfolio_Tracker.xlsx
                                            #   from code, with sample data
python -m networth.generate -o my.xlsx      # ...to a custom path
python -m networth.update                   # refresh the workbook in the current
                                            #   folder (must be CLOSED in Excel):
                                            #   prices, NAVs, corp actions, XIRR
python -m networth.update path\to\file.xlsx # ...a specific workbook
pytest                                      # full suite (~60 tests, ~30 s):
                                            #   XIRR golden values, feed parsers,
                                            #   corp-action scenarios, round-trip
                                            #   identity — no network needed
pytest tests/test_generate.py -q            # quick structural checks only
```

The dev loop for template changes: edit `src/networth/generate.py` → run
`python -m networth.generate` → open the xlsx and look → `pytest`. Never edit
the generated file directly (see the maintainers note below).

### Build the one-click apps (what end users download)

Run **on the OS you're building for** — PyInstaller cannot cross-compile.
From the repo root, *without* any venv active (the script makes its own):

```bat
:: Windows → dist\NetWorth-<version>-windows.zip  (+ Update Portfolio.exe)
packaging\build-release.bat 1.0.0rc3
```

```bash
# macOS  → dist/NetWorth-<version>-macos.zip
#          (+ networth-updater binary and "Update Portfolio.command" launcher)
packaging/build-release.sh 1.0.0rc3
```

Each zip contains the generated workbook, the updater app and a one-page
README — the exact layout CI publishes on releases. First-run warnings are
expected for unsigned builds: Windows SmartScreen → *More info → Run anyway*;
macOS Gatekeeper → *right-click → Open*.

### Troubleshooting

- **`python`/`py` not found** — Windows: reinstall from python.org (the
  Microsoft Store stub can shadow it; `py -3` is the reliable spelling).
  macOS: use `python3`, not `python`.
- **TLS/certificate errors during fetch** — corporate proxies/antivirus
  intercept HTTPS; the updater already validates against the OS trust store
  (`truststore`). In plain dev scripts, call
  `import truststore; truststore.inject_into_ssl()` first.
- **"workbook is open in Excel"** — the updater refuses by design; close the
  file (and check for a leftover `~$...xlsx` lock file).
- **NSE fetches are slow the first time** — the cookie warm-up request is
  normal; BSE is the primary price source anyway.
- **Charts vanished from a workbook you edited in code** — you saved it
  through openpyxl. Regenerate with `python -m networth.generate`; openpyxl
  is read-only in this codebase.

### Layout & releasing

Layout: `src/networth/generate.py` (xlsxwriter workbook builder) ·
`reader.py` (openpyxl read-only round-trip) · `update.py` (orchestrator) ·
`fetch/` (amfi, bhavcopy BSE→NSE, corporate_actions NSE+BSE) · `compute/`
(xirr, cashflows, projections) · `data/` (bundled bank list + FMV table) ·
`packaging/` (PyInstaller spec, per-OS build scripts) · `tests/`.

Releasing: push a `v*` tag — [.github/workflows/release.yml](.github/workflows/release.yml)
runs the test suite on Linux/Windows/macOS, builds both one-click zips and
attaches them to the GitHub Release (tags containing `-` are marked
pre-release automatically). To rehearse a release locally, run the
`packaging/build-release.*` script for your OS as above.

Key documents:

- [docs/SPEC.md](docs/SPEC.md) — the platform-agnostic specification (sheets,
  data contracts, algorithms). **The spec is the product**; Python here is the
  reference implementation. Behaviour changes update the spec in the same commit.
- [docs/PLAN.md](docs/PLAN.md) — approved architecture & feature plan.
- [docs/RELEASES.md](docs/RELEASES.md) — milestone plan & acceptance criteria (R0–R7 = v1).
- [docs/ROADMAP.md](docs/ROADMAP.md) — backlog & ideas (PPF ledger, capital-gains report, CAS import…).
- [CLAUDE.md](CLAUDE.md) — working notes/conventions for AI-assisted development.

**For maintainers:** end users edit the workbook freely — entering and updating
holdings in the input cells is the product, and the updater preserves all of it.
What must *not* happen is changing the template's **structure** (sheets, columns,
formulas, charts) by hand-editing the shipped xlsx, or saving it through
openpyxl (which silently destroys the charts). Structural changes belong in
`src/networth/generate.py`; rebuild to apply them.

## Privacy & data sources

All fetches are plain HTTPS GETs of public data, initiated from your machine
(TLS validated against your OS certificate store, so corporate proxies work):

- AMFI daily NAVs + scheme list: `https://www.amfiindia.com/spages/NAVAll.txt`
- BSE bhavcopy (prices, primary): `https://www.bseindia.com/download/BhavCopy/Equity/...`
- NSE bhavcopy (fallback): `https://nsearchives.nseindia.com/content/cm/...`
- NSE corporate actions (per held stock): `https://www.nseindia.com/api/corporates-corporateActions?...`
- BSE corporate actions (per held stock): `https://api.bseindia.com/BseIndiaAPI/api/DefaultData/w?...`

On each run the updater also makes one call to the GitHub releases API to tell
you if a newer version exists. It sends nothing about you (just a normal API
request) and is easy to turn off: pass `--no-update-check` or set
`NETWORTH_NO_UPDATE_CHECK=1`.

Bundled static data (no fetch needed): the Indian bank list, the 31-01-2018
FMV table, and the PPF interest-rate history — all refreshed via app releases.

Nothing about your holdings is ever uploaded, anywhere.

## License

[MIT](LICENSE)
