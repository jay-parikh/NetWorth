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
- ✅ **Dividend ledger** *(shipped v1.2.0, R9)* — a Dividends sheet logging
  every dividend paid on held stocks per financial year, filled automatically
  from both exchanges' announcements, with a Dashboard "Dividends this FY"
  total and a dividends-by-month chart.
- ✅ **Capital-gains tax report** *(shipped v1.6.0)* — realised/unrealised
  STCG & LTCG per FY on the Capital Gains tab, fed by the new Equity_Sells
  ledger + MF_SIP FIFO; 31-01-2018 grandfathering applied properly (bundled
  FMV, corp-action-normalised), ₹1.25L LTCG exemption tracking with a
  live headroom headline, per-sale-date rates (Budget 2024 mid-FY switch),
  and a sell-planning section (gain-if-sold-today + the date each holding
  turns long-term). Indicative, for planning — the sheet says so.
- ✅ **Dividends → true equity XIRR** *(shipped v1.6.0)* — Dividends rows
  (all FYs, ex-date ≤ today) and recorded Equity_Sells round trips now feed
  the equity cashflow model, plus a per-person "Dividends FY" split on each
  person sheet. Residual idea: `_dividend_qty` could add back sold lots
  (needs ex-date-unit conversion) — deferred until **broker import** lands,
  which brings the sale data needed to do it honestly (verdict 2026-07-18).
- ✅ **Same-FY loss set-off, Sec 70(2)** *(shipped v1.6.1)* — a short-term
  loss left over after netting now reduces the same year's LTCG before the
  §112A allowance, shown in its own By-FY column (equity family only, in
  FYs whose rules are known; debt↔equity cross-bucket netting is NOT
  modelled and the sheet says so). The following related ideas were
  deliberately excluded, with reasons (verdicts 2026-07-18 — settled,
  please don't re-open as "pending"):
  - *Carry-forward across years* — can't be honest: needs loss history from
    before the user started the ledger plus proof of timely ITR filing.
  - *Debt-fund indexation* — obsolete: sales on/after 23-07-2024 are 12.5%
    **without** indexation (already in `tax_rules_in.csv`); the old 20%+
    indexation regime only covers back-dated sales whose filing windows have
    closed.
  - *Short selling* — same-day shorts already land in speculative income
    (the branch tests date equality, order-agnostic); cross-day delivery
    shorting doesn't exist in the Indian cash segment (T+1), and F&O is out
    of scope — the "sold before bought" warning is the intended typo-catcher.
- ⬜ **Interest tracking** — FD interest payout mode (payout vs cumulative),
  bond coupon receipts ledger.

## Data in

- ✅ **NSE as full peer source for prices** *(shipped v1.2.0, R8)* — same-day
  BSE+NSE union merge, NSE wins price conflicts, BSE keeps the scrip codes;
  delisted/suspended escalation only on dual-source days.
- ✅ **Mergers / demergers / ISIN reassignments** *(shipped v1.4.0, R14)* —
  curated `restructures.csv` + Manual rows: consumed ISINs price via their
  successor at the right ratio (cost & holding period carried), demergers
  append the spun-off shares with the notified cost split, all audited on
  Corporate_Actions. Keeping the curated file current is a release duty.
- ⬜ **CAS import** — parse CAMS/KFintech Consolidated Account Statement PDFs
  to auto-fill the MF_SIP ledger.
- ⬜ **Broker import** — Zerodha (tradebook/holdings CSV) first; then generic
  contract-note CSV mapping.
- ✅ **Curated-data refresh cadence** *(decided 2026-07-18)* —
  `restructures.csv` (R14) and `bullion_proxies.csv` (R13) are refreshed as
  part of **every release** (a standing item in the RELEASES.md checklist),
  plus an out-of-band release whenever a major index stock restructures.

## More of the balance sheet

- ✅ **New asset classes** *(shipped v1.3.0, R12–R13)* — Gold/SGB/Silver
  (live IBJA bullion rate with market-implied fallback and manual override),
  NPS (daily NAV from NPS Trust), EPF (rate accrual from the bundled EPFO
  table), real estate, cash/savings, insurance surrender value — each
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
- ✅ **Asset-allocation targets** *(shipped v1.3.0, R11)* — target % per
  class on Settings, drift red/green, plain-language rebalancing hints,
  actual-vs-target chart — all live formulas.
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
- ✅ **Workbook protection** *(shipped v1.5.0, exceeded the original idea)* —
  two opt-in layers sharing one password: the ••• privacy mask (display
  curtain + full sheet protection) and the Lock (real OOXML AES encryption,
  Excel's own password-to-open). Residual idea: locked computed columns in
  the *unmasked* workbook to prevent accidental formula damage.
- ⬜ **Google Sheets port** — a second implementation of docs/SPEC.md (Apps
  Script), proving the spec's portability promise.
