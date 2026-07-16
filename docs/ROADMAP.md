# Roadmap — everything beyond v1.0, with status

Every idea stays on this list with its current status — nothing is deleted
when it ships or gets scheduled, so this page is the one place to see where
the product is going and how far it has come.

**Status markers:**
✅ **done** — shipped ·
🚧 **planned** — has acceptance criteria in [RELEASES.md](RELEASES.md)
(design detail for v1.2–v1.4 in [PLAN-v1.2.md](PLAN-v1.2.md), planned
2026-07-16) ·
⬜ **pending** — an idea, not a promise; it becomes real only when it moves
into RELEASES.md with acceptance criteria.

## Money & tax

- ✅ **PPF contribution ledger** *(shipped v1.1)* — one row per deposit,
  proper monthly-minimum-balance interest computation using the bundled
  quarterly `ppf_rates.csv` (refreshed via app releases; there is no official
  API). Replaced the flat Rate% estimate; enables true PPF XIRR.
- 🚧 **Dividend ledger** *(R9, v1.2)* — a Dividends sheet logging every
  dividend paid on held stocks per financial year, filled automatically from
  both exchanges' announcements, with a Dashboard "Dividends this FY" total
  and a dividends-by-month chart.
- ⬜ **Capital-gains tax report** — realised/unrealised STCG & LTCG per FY,
  with 31-01-2018 grandfathering applied properly (reuses the bundled FMV
  data), ₹1.25L LTCG exemption tracking, and a sell-planning helper.
- ⬜ **Dividends → true equity XIRR** — feed the Dividends sheet rows into
  the equity cashflow model, plus a per-person split of the FY dividend
  total. Kept separate from R9 deliberately: changing return semantics
  deserves its own release.
- ⬜ **Interest tracking** — FD interest payout mode (payout vs cumulative),
  bond coupon receipts ledger.

## Data in

- 🚧 **NSE as full peer source for prices** *(R8, v1.2)* — same-day BSE+NSE
  union merge; today NSE bhavcopy is fallback only (corporate actions
  already query both exchanges).
- 🚧 **Mergers / demergers / ISIN reassignments** *(R14, v1.4)* — the one
  corporate-action category with no reliable free feed for swap ratios;
  handled via a curated actions file shipped with releases plus the existing
  Manual rows.
- ⬜ **CAS import** — parse CAMS/KFintech Consolidated Account Statement PDFs
  to auto-fill the MF_SIP ledger.
- ⬜ **Broker import** — Zerodha (tradebook/holdings CSV) first; then generic
  contract-note CSV mapping.
- ⬜ **Curated-data refresh cadence** — `restructures.csv` (R14) and
  `bullion_proxies.csv` (R13) are release-refreshed curated files; decide a
  rhythm (e.g. every release, plus an out-of-band release for a major
  index-stock merger) once R14 ships.

## More of the balance sheet

- 🚧 **New asset classes** *(R12–R13, v1.3)* — Gold/SGB/Silver (live
  bullion-market rate with manual override), NPS (daily NAV), EPF (rate
  accrual), real estate, cash/savings, insurance surrender value — each
  user-selectable so only relevant sheets show (R10) and each with its own
  allocation target (R11).
- ⬜ **ESOPs/RSUs** — deferred 2026-07-16 pending multi-currency:
  foreign-listed grants need FX + foreign price feeds. Until then,
  Indian-listed vested RSUs go on the Equity sheet, and a hand-converted ₹
  value can sit in Manual_Assets (Class "Other").
- ⬜ **NPS contribution ledger** — dated contributions for exact NPS XIRR,
  replacing R13's approximate two-flow (mirrors the PPF flat-first path).
- ⬜ **EPF exact accrual** — contribution ledger + monthly-interest-on-
  running-balance rule with annual crediting, replacing R12's flat estimate.
- ⬜ **Liabilities** — loans/EMIs, so the Dashboard shows true net worth
  (assets − liabilities).

## Insight & planning

- ✅ **Net-worth history** *(shipped v1.1)* — the updater appends a dated
  snapshot row on each run; trend chart on the Dashboard (a by-class
  stacked-area companion arrives with R10).
- 🚧 **Asset-allocation targets** *(R11, v1.3)* — target % per class, drift
  red/green, plain-language rebalancing hints, actual-vs-target chart.
- ⬜ **Goal planning** — goals with target date/amount mapped to holdings;
  on/off-track verdict using the projection engine.
- ⬜ **Live native-XIRR formulas** — replace script-written XIRR values with
  Excel `XIRR()` over generated helper cashflow ranges, so returns update on
  every edit without running the updater.

## Platform & polish

- ✅ **Auto-update check** *(shipped v1.1)* — the updater compares its
  version against the latest GitHub release tag and prints a one-line hint
  (still fully local/offline-safe).
- ⬜ **Search-as-you-type inside dropdowns** — typing while the list is open
  is an Excel-owned interaction a plain .xlsx cannot provide. Current
  Microsoft 365 (≈ version 2304+) already does it natively with our
  validations (AutoComplete for data validation, contains-matching), so the
  primary answer is "update Office". If demand persists for older Excel, the
  option is a separate **macro-enabled (.xlsm) variant** with a VBA search
  popup — costs: macro security warnings everywhere, no LibreOffice, Mac
  userform quirks — kept out of the default build deliberately.
- ⬜ **Signed/notarized binaries** (macOS notarization, Windows code signing)
  to remove the right-click→Open / SmartScreen friction.
- ⬜ **Multi-currency** — foreign stocks/funds, FX rates, INR-consolidated
  view (unblocks ESOPs/RSUs above).
- ⬜ **Workbook protection guidance** — password/encryption how-to; locked
  computed columns to prevent accidental formula damage.
- ⬜ **Google Sheets port** — a second implementation of docs/SPEC.md (Apps
  Script), proving the spec's portability promise.
