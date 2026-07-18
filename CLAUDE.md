# CLAUDE.md — NetWorth development guide

Read this before touching anything. It captures the decisions, gotchas and
conventions agreed with the project owner (Jay), so any session — human or AI —
starts from the same page.

## What this project is

A **local-only, Excel-based family portfolio tracker** for Indian investors
(Equity / Mutual Funds+SIP / FD / PPF / Bonds) that a **non-developer** can use:
open the workbook, type holdings, double-click one updater app to refresh
prices, NAVs, XIRR and projections. Windows **and** macOS. No cloud, no server,
no Python knowledge required of the end user.

Repo: https://github.com/jay-parikh/NetWorth (owner decides when to push).

## The three governing decisions

1. **Spec-first, portable by design.** `docs/SPEC.md` is a platform-agnostic
   specification (sheet schemas, data contracts, algorithms in pseudocode).
   The Python code in `src/networth/` is the *reference implementation*.
   Anyone must be able to rebuild the product on any language/platform from
   the spec alone. When behaviour changes, **update the spec in the same
   commit** as the code.

2. **The workbook is a build artifact — the "regenerate" model.**
   `generate.py` builds the entire xlsx from code (xlsxwriter): all sheets,
   charts, dropdowns, formats, sample data. The updater does a round-trip:
   read user inputs (openpyxl **read-only**) → fetch public data → compute in
   Python → **regenerate the whole workbook** → timestamped backup first.
   Consequences: the file must be closed during update; only structured input
   cells survive a rebuild (documented to users).

3. **End users get one-click binaries.** PyInstaller executables per OS
   (`Update Portfolio.exe` / `.command`), shipped via GitHub Releases together
   with the generated template xlsx. Python is a dev-time dependency only.

## Critical gotchas (learned the hard way)

- **End users DO edit the workbook — that's the product.** They type holdings
  into input cells, add/delete/sort rows, and nothing may break because of it.
  The "don't edit" rules below are for *maintainers* changing the template's
  structure, never for users entering data.
- **NEVER load-and-save the xlsx with openpyxl** — it silently drops all
  6 charts and every cell comment. openpyxl is for *reading only* here.
  (One carve-out, v1.6.2: TESTS may load→edit→save to a NEW path to
  fabricate user-edited fixtures that only `read_workbook` will consume —
  the reader needs no charts/comments. Never for a shipped or user file.)
  All writing goes through xlsxwriter in `generate.py`. The legacy template
  was maintained by raw XML surgery inside the zip for exactly this reason;
  the regenerate model exists so nobody ever does that again. Likewise,
  structural changes (sheets/columns/formulas/charts) are made in
  `generate.py` and rebuilt — never by hand-editing a shipped xlsx.
- **Master sheets must stay sorted by name** (ordinal, case-insensitive).
  The type-ahead dropdowns are `OFFSET(...MATCH($C4&"*"...)...)` begins-with
  window formulas over a sorted list; an unsorted master silently breaks them.
- **Stock_Master merges are ADD-ONLY**: existing ISINs keep their existing
  display name even if the exchange renames the security — user rows key on
  the name via INDEX/MATCH, and renames would orphan them.
- **Data validations are non-blocking** (`showErrorMessage=false`): a blank
  looked-up ISIN is the "no match" signal, and users may accept a warning to
  type delisted/merged names manually. Preserve this behaviour.
- **XIRR cells are plain values** written at update time, not formulas.
  Excel's native XIRR is a roadmap item — don't half-migrate.
- NSE endpoints need browser-ish headers and a cookie warm-up
  (`www.nseindia.com` first). Since v1.2 (R8) BSE and NSE bhavcopies are
  **full peers**: same-day union merge, NSE wins price conflicts, BSE alone
  supplies scrip codes, and delisted/suspended escalation only runs on
  dual-source days. See `legacy/UpdatePrices.ps1` `Get-Bhavcopy` for a
  working recipe of the headers.

## Product principles (Jay, 2026-07-17)

- **Docs follow code, always.** README, the workbook's Guide sheet and the
  docs are updated in the same commit as the change they describe — the
  SPEC-in-same-commit rule extends to every user-facing doc.
- **Plain language for anything a user reads** (README, Guide, prompts,
  warnings, release notes): simple words, short sentences, no jargon —
  end users are not software people. Keep the meaning exact.
- **Don't intimidate at first open.** New and advanced features default to
  off/hidden (progressive disclosure); the workbook must look simple at
  first glance even as features grow.
- **Every release ships plain-English notes** on GitHub Releases, written in
  `docs/release-notes/<tag>.md` (first line `# <title>`, rest = body). The
  release workflow publishes the file when the tag is pushed;
  `scripts/publish_release_notes.py` back-fills or edits published ones.

## Conventions

- **Commits:** small, self-contained, one milestone (or sub-step) each, message
  prefixed with the milestone (`R2: ...`). **No `Co-Authored-By` trailer, ever.**
  Do not push or create the GitHub repo unless Jay explicitly asks.
- **Milestones:** R0 scaffold → R1 spec → R2 generator (core parity) →
  R3 updater → R4 packaging → R5 quick wins (red/green, bond maturity, bank
  master, FY-end value) → R6 FMV-2018 + delisted → R7 corporate actions = v1.0.
  Details and acceptance criteria: `docs/RELEASES.md`.
- **Dates in docs:** absolute (e.g. 2026-07-15), never "today/yesterday".
- Money/units: INR only for now (multi-currency is roadmap).

## Build / test / run (dev)

```bash
pip install -e ".[dev]"           # xlsxwriter, openpyxl, requests, pytest
python -m networth.generate        # → Family_Portfolio_Tracker.xlsx (repo root, gitignored)
python -m networth.update PATH     # round-trip update of an existing workbook
pytest                             # XIRR golden values, round-trip identity, parsers, corp actions
```

Verification beyond pytest: open the generated workbook in Excel (Jay's
Windows machine) or LibreOffice and check charts, type-ahead dropdowns,
conditional formats. The round-trip identity test (generate → read → regenerate
→ semantically equal) is the regression backbone.

## Layout

```
docs/SPEC.md       platform-agnostic spec — THE source of truth
docs/PLAN.md       approved plan (2026-07-15) — context for why things are this way
docs/RELEASES.md   milestone plan + acceptance criteria
docs/ROADMAP.md    post-v1 backlog / brainstorm
src/networth/      reference implementation (model, generate, reader, fetch/, compute/, update)
data/              bundled statics: banks_in.csv, ppf_rates.csv, fmv_2018-01-31.csv
legacy/            the original hand-built template + PowerShell updaters (Windows/COM).
                   Read-only reference & migration source. Do not extend it.
packaging/         PyInstaller specs / per-OS build scripts
```

## Domain notes that shape features

- **FMV 31-01-2018**: Indian LTCG "grandfathering" — for equity bought before
  2018-02-01, cost basis may use the 31-Jan-2018 fair market value (highest
  traded price that day). We bundle that day's bhavcopy as a lookup so users
  who don't know an old buy price still get sane numbers (flagged amber).
- **PPF rate** is set quarterly by the Ministry of Finance; there is **no API**.
  We bundle a historical `ppf_rates.csv`, refreshed via app releases.
- **Corporate actions** (split/bonus/consolidation) never mutate the user's raw
  purchase rows: adjustments are applied at compute time from a fetched+manual
  actions table, with a visible audit sheet — idempotent across re-runs, and
  future ex-dates take effect automatically when they arrive.
- **Delisted/suspended** scrips: detected by prolonged absence from bhavcopy +
  exchange status; valued at last traded price, flagged, excluded from
  day-change. User may override the price manually.
- **XIRR cashflow model** (per class): Equity −Invested@CostDate/+CurVal@today;
  MF from the MF_SIP ledger (redemptions = negative Amount rows);
  FD −Principal@Start/+Value@min(today, maturity); PPF estimated at Rate% from
  the as-on date (no ledger yet); Bonds −Buy@BuyDate/+CurVal@today, plus coupon
  cashflows once R5 lands. Same-date degenerate flows → return null, skip.
