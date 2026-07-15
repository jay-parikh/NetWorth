# Roadmap — what's next

Ideas beyond the shipped releases (v1.0 core, v1.1 = PPF ledger + net‑worth
history + auto‑update check), roughly ordered by value vs. effort. Nothing here
is a promise — an item becomes real only when it moves into
[RELEASES.md](RELEASES.md) with acceptance criteria.

## Money & tax

- **PPF contribution ledger** — one row per deposit, proper monthly-minimum-
  balance interest computation using the bundled quarterly `ppf_rates.csv`
  (auto-refresh of that table via app releases; there is no official API).
  Replaces today's flat Rate% estimate; enables true PPF XIRR and maturity
  projection.
- **Capital-gains tax report** — realised/unrealised STCG & LTCG per FY,
  with 31-01-2018 grandfathering applied properly (reuses the bundled FMV
  data), ₹1.25L LTCG exemption tracking, and a sell-planning helper.
- **Dividend & interest tracking** — dividend ledger per scrip (feeds true
  equity XIRR), FD interest payout mode (payout vs cumulative), bond coupon
  receipts ledger.

## Data in

- **CAS import** — parse CAMS/KFintech Consolidated Account Statement PDFs to
  auto-fill the MF_SIP ledger.
- **Broker import** — Zerodha (tradebook/holdings CSV) first; then generic
  contract-note CSV mapping.
- **NSE as full peer source for prices** — today NSE bhavcopy is fallback only
  (corporate actions already query both exchanges).
- **Mergers / demergers / ISIN reassignments** — the one corporate-action
  category that cannot be auto-adjusted (no reliable free feed for swap
  ratios); today the updater flags unverifiable holdings and the
  Corporate_Actions Manual rows cover it. A curated actions file shipped with
  releases could close most of the gap.

## More of the balance sheet

- **New asset classes** — Gold/SGB, NPS, EPF, real estate, cash/savings,
  insurance (surrender value), ESOPs/RSUs.
- **Liabilities** — loans/EMIs, so the Dashboard shows true net worth
  (assets − liabilities).

## Insight & planning

- **Net-worth history** — the updater appends a dated snapshot row on each run;
  trend chart on the Dashboard.
- **Asset-allocation targets** — target % per class, drift red/green,
  rebalancing hints.
- **Goal planning** — goals with target date/amount mapped to holdings;
  on/off-track verdict using the projection engine.
- **Live native-XIRR formulas** — replace script-written XIRR values with
  Excel `XIRR()` over generated helper cashflow ranges, so returns update on
  every edit without running the updater.

## Platform & polish

- **Search-as-you-type inside dropdowns** — typing while the list is open is
  an Excel-owned interaction a plain .xlsx cannot provide. Current Microsoft
  365 (≈ version 2304+) already does it natively with our validations
  (AutoComplete for data validation, contains-matching), so the primary
  answer is "update Office". If demand persists for older Excel, the option
  is a separate **macro-enabled (.xlsm) variant** with a VBA search popup —
  costs: macro security warnings everywhere, no LibreOffice, Mac userform
  quirks — kept out of the default build deliberately.

- **Signed/notarized binaries** (macOS notarization, Windows code signing) to
  remove the right-click→Open / SmartScreen friction.
- **Auto-update check** — updater compares its version against the latest
  GitHub release tag and prints a one-line hint (still fully local/offline-safe).
- **Multi-currency** — foreign stocks/funds, FX rates, INR-consolidated view.
- **Workbook protection guidance** — password/encryption how-to; locked
  computed columns to prevent accidental formula damage.
- **Google Sheets port** — a second implementation of docs/SPEC.md (Apps
  Script), proving the spec's portability promise.
