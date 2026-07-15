# NetWorth — Generic, Cross-Platform Family Portfolio Tracker

## Context

The existing `Portfolio_Template.zip` (in `/home/frappe/frappe-bench/NetWorth/`) is a working Excel family-portfolio tracker (Equity, MF/SIP, FD, PPF, Bonds, XIRR, inflation projection) but has three structural problems:

1. **Windows-only** — the three updaters are PowerShell + desktop-Excel COM.
2. **Not reproducible** — no generator exists; the xlsx was built by hand + raw XML surgery (openpyxl drops the 6 charts and comments on save), so every edit is fragile archaeology.
3. **Not shareable** — no repo, no spec, no docs; nothing a wider audience (or another implementer) can build from.

Goal: turn it into a public, spec-driven, local-only product at **github.com/jay-parikh/NetWorth** — Windows **and** macOS, zero cloud/server (network used only to fetch public price data: AMFI, BSE), and simple enough for a **non-developer** end user (double-click, no Python knowledge needed).

## Locked decisions (from user answers)

- End user is not a developer → ship **one-click packaged updaters** (PyInstaller executables per OS). Python is the *development* toolchain only.
- **All Python source generates the final Excel** — the workbook is a build artifact, never hand-edited XML again.
- v1 scope: core parity + red/green formatting + bond maturity + bank master + FY-end expected value, **plus** FMV 31-01-2018 / delisted handling, **plus** corporate actions (split/bonus). PPF auto-rate ledger → roadmap.
- Repo: `jay-parikh/NetWorth`. User decides when to commit/push; plan the work as **small, individually releasable commits/milestones**.
- Per memory: **no `Co-Authored-By` trailer** in any commit.

## Architecture

### Two-layer design (portability requirement)

1. **`docs/SPEC.md` — platform-agnostic specification.** Defines everything needed to reimplement in any language/platform: sheet schemas (columns, types, input vs computed), data contracts (AMFI NAVAll.txt format, BSE bhavcopy format, corporate-action feed), and algorithms in pseudocode (XIRR cashflow assembly per asset class, FEFO-style batch… n/a — corporate-action adjustment, FMV fallback rule, FY-end projection math, delisted detection). The spec is the product; Python is the *reference implementation*.
2. **Python reference implementation** — pure-Python, cross-platform, no desktop-Excel dependency.

### The "regenerate" model (replaces COM/XML surgery)

The workbook is both UI and data store. The updater does a **round-trip rebuild**:

```
read user inputs (openpyxl, read-only — safe, nothing saved)
  → fetch AMFI NAVs / BSE bhavcopy / corporate actions (local HTTP GETs, cached)
  → compute (XIRR, adjustments, projections) in Python
  → regenerate the ENTIRE workbook with xlsxwriter (charts, dropdowns,
    conditional formats, comments — all native, no XML surgery)
  → timestamped backup of the previous file first (backups/ folder, keep last N)
```

Consequences (documented in README + Guide sheet): file must be closed during update; only structured input cells round-trip (user's ad-hoc formatting outside input areas is not preserved). In exchange: deterministic builds, testable, works identically on Windows/Mac/Linux, charts never corrupt.

XIRR stays **Python-computed plain values** (as today) — live native-formula XIRR is a roadmap item.

### End-user experience

- Download a per-OS zip from GitHub Releases: `Family_Portfolio_Tracker.xlsx` + `Update Portfolio.exe` (Windows) / `Update Portfolio.command`+binary (macOS) + a 1-page README.
- One updater binary replaces today's three .bat files (prices + NAV + masters + corporate actions in one run, with a console log and pause, like today).
- macOS Gatekeeper: document right-click → Open (unsigned initially; signing → roadmap).

### Data sources (all public, fetched locally)

| Data | Source | Notes |
|---|---|---|
| MF NAVs + scheme master | AMFI NAVAll.txt | same as today |
| Stock prices + stock master | BSE bhavcopy (ISIN-keyed) | same as today; NSE fallback roadmap |
| Corporate actions | BSE corporate-actions CSV/API per held ISIN | new |
| FMV 31-01-2018 | bhavcopy of 31-Jan-2018 → bundled `data/fmv_2018-01-31.csv` (static) | new |
| Bank master | bundled `data/banks_in.csv` (RBI scheduled-bank list, static) | new |
| PPF rates | bundled `data/ppf_rates.csv` (historical quarterly table); auto-refresh → roadmap | answers "how to auto-update PPF interest": no official API exists — bundle + update via app releases |

## Repo layout

```
NetWorth/
├── README.md              # what it is, screenshots, download + 5-min quickstart, FAQ
├── CLAUDE.md              # dev guide for Claude: architecture, build/test/release
│                          #   commands, "never hand-edit the xlsx", commit conventions
├── LICENSE                # MIT
├── docs/
│   ├── SPEC.md            # full platform-agnostic spec (sheets, formats, algorithms)
│   ├── ROADMAP.md         # brainstormed backlog (below)
│   └── RELEASES.md        # milestone/release plan
├── src/networth/
│   ├── model.py           # dataclasses: Holding, Lot, CashFlow, CorporateAction…
│   ├── generate.py        # workbook builder (xlsxwriter): sheets, charts, formats
│   ├── reader.py          # round-trip input reader (openpyxl read-only)
│   ├── fetch/             # amfi.py, bse.py, corporate_actions.py (cache + retry)
│   ├── compute/           # xirr.py, projections.py, corp_action_adjust.py, fmv.py
│   └── update.py          # orchestrator = the packaged entry point
├── data/                  # bundled static CSVs (fmv, banks, ppf_rates)
├── tests/                 # pytest: xirr golden values, round-trip identity,
│                          #   corp-action scenarios, parser fixtures
├── packaging/             # PyInstaller specs + per-OS build scripts
└── legacy/                # today's Portfolio_Template contents, for reference/migration
```

## Feature specifications (v1)

- **Red/green cells**: conditional formats on all Gain/Loss, Return-%, and XIRR columns (green ≥ 0, red < 0; amber for stale/delisted prices) + Dashboard per-class and real-return verdict.
- **Bond maturity amount**: Bonds sheet gains Coupon %, Frequency, Maturity Date → computed Maturity Value column (and coupon cashflows feed bond XIRR — fixes today's "price-return only" gap).
- **Bank master**: FD sheet Bank column becomes dropdown from bundled bank list (type-ahead pattern already proven on MF/Stock masters).
- **FY-end expected value**: new Dashboard column "Expected @ 31-Mar-{FY}". Fixed income (FD/PPF/Bonds) = deterministic accrual math to FY end; market assets (Equity/MF) = current value grown at a user-input "expected return %" cell — clearly labelled as an estimate.
- **FMV 31-01-2018 fallback**: Equity buy-price column may be blank/unknown → cost falls back to bundled 31-01-2018 FMV (LTCG grandfathering value) with an amber flag + comment noting the fallback. Also stored separately so a future capital-gains report can apply the proper grandfathering rule.
- **Delisted / not-traded**: Stock_Master gains Status (Active/Suspended/Delisted) + Last-Traded-Date, derived from bhavcopy absence over N sessions + BSE status. Rows valued at last traded price, flagged amber, excluded from "today's change"; user can override price manually.
- **Corporate actions (splits, bonus, and consolidation)**: updater fetches actions for *held* ISINs (past + announced), computes adjusted qty/avg-cost from raw purchase lots, writes a visible **Corporate_Actions audit sheet** (raw → adjusted, per action). User's raw entries are never mutated — adjustments are applied at compute time, so re-runs are idempotent and future ex-dates apply automatically when they arrive. Manual-action rows supported for anything the feed misses.

## Release plan (small, traceable milestones — one releasable commit-set each)

- **R0 — scaffold**: repo init, README skeleton, CLAUDE.md, LICENSE, SPEC/ROADMAP/RELEASES skeletons, `legacy/` import. No code.
- **R1 — spec first**: complete `docs/SPEC.md` for core parity (reverse-engineer current workbook: sheet map, styles, defined names, validation formulas, XIRR cashflow model — most already captured in memory notes).
- **R2 — generator, core parity**: `generate.py` builds the full template from code (all sheets, masters, type-ahead dropdowns, 6 charts, Guide, sample data). Acceptance: feature-parity with today's xlsx.
- **R3 — cross-platform updater**: fetch + compute + regenerate + backups; replaces all three PS1 scripts. Runs via Python (dev).
- **R4 — one-click packaging**: PyInstaller builds for Windows/mac, release-zip layout, end-user README. First public-usable release.
- **R5 — quick wins**: red/green formats, bond maturity + coupon XIRR, bank master, FY-end expected value.
- **R6 — FMV + delisted**: bundled FMV data, fallback logic, status/staleness flags.
- **R7 — corporate actions**: feed fetcher, adjustment engine, audit sheet. → tag **v1.0**.

## ROADMAP.md brainstorm (beyond v1)

PPF contribution ledger with monthly-minimum-balance interest + bundled rate table auto-refresh; capital-gains tax report (STCG/LTCG with grandfathering — reuses FMV data); dividend tracking; CAS PDF import (CAMS/KFintech) and broker tradebook import (Zerodha); more asset classes (gold/SGB, NPS, EPF, real estate, cash, insurance) and liabilities/loans for true net worth; net-worth history snapshots + trend chart; asset-allocation targets with rebalance hints; goal planning; live native-XIRR formulas; NSE price fallback; multi-currency; workbook password/encryption guidance; signed/notarized binaries; Google Sheets port (proves SPEC portability).

## Verification

- `pytest`: XIRR against known Excel values; corporate-action scenarios (split, bonus, combined, future-dated); parser fixtures for AMFI/bhavcopy; **round-trip identity test** (generate → read → regenerate → semantically identical).
- Open generated xlsx in LibreOffice headless / Excel to confirm charts, dropdowns, conditional formats render (manual check on the user's Windows machine for Excel-specific behaviors, as with the current template).
- End-to-end: run packaged updater on a copy of the sample workbook; verify prices/NAV/XIRR/backup file produced.

## Notes for implementation sessions

- Never edit the generated xlsx by hand or with openpyxl-save (kills charts) — always change `generate.py` and rebuild.
- Commits: small, per-milestone, no Co-Authored-By trailer; user decides when to push to github.com/jay-parikh/NetWorth.
