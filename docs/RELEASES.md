# Release / milestone plan

Small, individually releasable milestones. One milestone = one coherent set of
commits, prefixed `R<n>:`. Tags follow `v0.<n>` per milestone; `v1.0` at R7.

| # | Milestone | Delivers | Acceptance criteria |
|---|---|---|---|
| R0 | Scaffold | Repo layout, README, CLAUDE.md, LICENSE, PLAN, doc skeletons, `legacy/` import | Repo browsable; legacy template usable as-is from `legacy/` |
| R1 | Spec first | Complete `docs/SPEC.md` for core parity, reverse-engineered from the legacy workbook | A competent dev could rebuild the workbook from SPEC.md alone, without opening the legacy xlsx |
| R2 | Generator (core parity) | `generate.py` + `model.py`: full template from code — all 15 sheets, masters, type-ahead dropdowns, 6 charts, Guide, sample data | Generated xlsx is feature-equivalent to `legacy/Family_Portfolio_Tracker.xlsx`; opens clean in Excel & LibreOffice; round-trip identity test passes |
| R3 | Cross-platform updater | `update.py` + `reader.py` + `fetch/` + `compute/xirr.py`: one command replaces all three PS1 scripts; timestamped backups; console log | On a workbook with sample data: prices, NAVs, masters and XIRR refresh correctly on Linux/Mac/Windows with only Python installed |
| R4 | One-click packaging | PyInstaller builds (Windows `.exe`, macOS `.command` + binary), release-zip layout, 1-page end-user README | A non-developer can download a zip, double-click, and get an updated workbook. First public-usable release |
| R5 | Quick wins | Red/green conditional formats (green ≥ 0 / red < 0 / amber stale), bond Maturity Value + coupon-aware XIRR, bank-master dropdown for FDs, "Expected @ FY-end" column | Each visible in the generated workbook; XIRR for a couponed bond matches a hand-checked value |
| R6 | FMV 2018 + delisted | `data/fmv_2018-01-31.csv` bundled; blank buy-price falls back to FMV (amber + comment); Stock_Master Status/Last-Traded-Date; stale-price flagging | Unknown-cost holding values correctly with visible flag; a delisted sample scrip shows last price, amber, excluded from day-change |
| R7 | Corporate actions → **v1.0** | BSE corporate-action fetch for held ISINs, split/bonus/consolidation adjustment engine, Corporate_Actions audit sheet, manual-action rows | Scenario tests pass (split, bonus, combined, future-dated, idempotent re-run); raw user rows never mutated |

## v1.1 — "PPF done right + your money over time"

Post-v1.0 milestone, three independent features, each its own commit + tests.

| Feature | Delivers | Acceptance criteria |
|---|---|---|
| PPF accurate interest | `data/ppf_rates.csv` (historical quarterly rates), `compute/ppf.py` (monthly-minimum-balance rule, annual crediting), a new **PPF_Ledger** sheet (Owner/Account/Date/Amount). **Optional-ledger + fallback**: accounts with ledger rows get exact balance/interest/XIRR; accounts without keep today's flat estimate. Current PPF rate auto-filled from the table for everyone. | `compute/ppf.py` matches a hand-worked example; a ledger account's Balance today = deposits + credited/accrued interest; non-ledger accounts unchanged; Dashboard/person PPF totals use the computed balance |
| Net-worth history + trend | New **History** sheet (Date + per-class + total); the updater upserts one snapshot per day; Dashboard line chart of net worth over time. History rows round-trip regeneration (they are data). | After two updater runs on different dates, History has two rows and the Dashboard chart plots them; regeneration preserves history |
| Auto-update check | The updater queries the latest GitHub release tag and prints a one-line hint when a newer version exists. Silent on network failure / no releases. | Newer remote tag → hint printed; same/older or offline → nothing; never blocks the run |

Ships as **v1.1.0** once all three land and verify.

## v1.2 — "Both exchanges + dividend income"

*For the user: prices that match their broker's app, and a Dividends tab
showing the cash their shares paid this financial year.*

Planned 2026-07-16; design detail in [PLAN-v1.2.md](PLAN-v1.2.md).
**Shipped as v1.2.0 (2026-07-16)** — R8 + R9 landed with tests
(`test_r8_dual_source.py`, `test_r9_dividends.py`, 103 total) and a live
end-to-end run (10/10 scrips BSE+NSE, 9 real dividend rows, CA-adjusted
quantities verified against real splits).
**Standing acceptance criteria for every R8–R14 row** (the design principles
in PLAN-v1.2.md): the default template stays as clean as v1.1; every new
sheet has a plain one-sentence hint; every amber cell has a comment saying
why; the Guide covers the feature in plain language; a non-developer can use
it without reading these docs.

| # | Milestone | Delivers | Acceptance criteria |
|---|---|---|---|
| R8 | NSE full peer prices | Same-day BSE+NSE bhavcopy union merge (NSE close/prev win on conflict; BSE scrip codes always retained when BSE responded); per-run source label; Suspended/Delisted escalation only on dual-source days | Fixture pair: BSE has ISIN X only, NSE has Y only, both have Z at different closes → merged result prices X, Y, Z@NSE; `codes_by_isin` keeps BSE codes for X and Z; NSE-down run completes BSE-only with no status transitions; both-down walks back then errors as before; round-trip identity passes |
| R9 | Dividends sheet | Dividend parsing from the existing dual-exchange announcement fetch (Interim/Final/Special, ₹/share variants; %-of-face skipped, counted in summary); new **Dividends** sheet (FY, Owner, Scrip, ISIN, Type, Ex-Date, Rate, Qty@ex-date est., Est. amount, Source, Details); qty estimated from lots with cost date < ex-date, CA-adjusted to ex−1; current-FY Auto rows rebuilt each run, prior-FY rows frozen, Manual rows persist and override same key; Dashboard "Dividends this FY" cell + a "Dividends by month" column chart on the Dividends sheet. Dividends do **not** feed equity XIRR yet | Parser goldens incl. `"Interim Dividend - Rs. - 5.5000"`, `"Dividend - Rs 8 Per Share"`, `"Re. 1/- per share"`, `"Dividend 250%"` (skipped); a held stock with a fixture announcement yields one row per owner with CA-adjusted pre-ex-date qty; re-run idempotent; Manual row suppresses Auto with same key; a prior-FY row survives a new-FY run; Dashboard cell sums current FY only; round-trip test extended |

## v1.3 — "Your whole balance sheet"

*For the user: switch on only the assets they own from a simple Settings tab
— gold, EPF, NPS, the house, savings, insurance — and see at a glance whether
their money is balanced the way they planned.*

**Shipped as v1.3.0 (2026-07-16)** — R10–R13 landed with tests
(`test_r10_settings.py` … `test_r13_wave2.py`, 126 total) and a live
end-to-end run with all 12 classes on: SGB priced from the merged bhavcopy,
gold/silver rated from live IBJA (₹14,167.9/g / ₹217.43/g), NPS NAV matched
from the live NPS Trust feed (262-scheme master), EPF rate auto-filled.

| # | Milestone | Delivers | Acceptance criteria |
|---|---|---|---|
| R10 | Settings + class registry + selectable classes | `ASSET_CLASSES` registry in `model.py` (dict-backed ClassXirr/HistorySnapshot; registry-driven Dashboard matrix, allocation table, person blocks, History columns, charts, snapshot, XIRR dispatch); new **Settings** sheet (per-class Yes/No + Target % + Status, drift tolerance, targets-total sanity); disabled classes: sheet hidden (never omitted), Dashboard/person/History presence omitted; `effective_enabled = enabled OR has_data` — a class with data is never hidden (Status `On (has data)` + updater warning); classic five default Yes, new classes No; header-located Dashboard reads; Dashboard visual upgrade: stacked-area "Net worth by class over time" chart, data bars on allocation values, ▲/▼ arrows on day-change | Default template shows exactly the v1.1 tab set + Settings; stacked-area chart, data bars and ▲/▼ arrows render in Excel and LibreOffice; toggling a class Yes reveals its sheet/column/row/blocks, toggling No hides them with totals unchanged; a class with data set to No stays visible with warning; round-trip identity passes for any enable combination; a v1.1 workbook (no Settings) updates cleanly with defaults; openpyxl reads hidden sheets in the round-trip test |
| R11 | Allocation targets | Dashboard allocation table gains Actual %, Target % (from Settings), Drift (absolute pp, red/green vs ±tolerance, default 5), Rebalance hint (`On target` / `Move ₹X out` / `Move ₹X in`, pre-tax, class-level); "Actual vs Target %" column chart; all live formulas — no updater run needed | With Equity 62% actual vs 55% target and tol 5: drift +7.0pp red, hint `Move ₹<0.07×total> out`; within-band → green `On target`; blank target → no drift/hint/CF; targets ≠ 100 flags amber on Settings; drift updates live on editing a holding; disabled classes absent from drift view and chart |
| R12 | New classes wave 1 | **Manual_Assets** sheet (Class dropdown: Real Estate / Cash / Insurance / Other — each its own registry class via SUMIFS filter; Value-as-on amber when > 90 days stale); **EPF** sheet (passbook balance + as-on + rate pre-filled from bundled `data/epf_rates.csv`, flat accrual to today); XIRR: two-flow for RE/Insurance/Other, none for Cash, PPF-style for EPF; History columns become label-keyed (old workbooks migrate losslessly) | An RE row with cost+date+value yields a hand-checked two-flow XIRR; Cash shows no XIRR anywhere; EPF balance accrues at the bundled rate; stale value-as-on flags amber; each enabled subclass gets its own Dashboard column/allocation row/person block/History column; a v1.1 History reads losslessly (old totals unchanged, new classes 0); round-trip passes with all wave-1 classes populated |
| R13 | New classes wave 2 | **Gold_Silver** sheet (SGB priced from the merged bhavcopy by ISIN incl. 2.5% semi-annual coupon XIRR; physical gold/silver = grams × purity × ₹/g rate); metal rate layered: IBJA daily benchmark primary → bhavcopy-implied median over bundled `data/bullion_proxies.csv` → carry-forward + amber (> 7 days) → manual Rate-override always wins; **NPS** sheet + **NPS_Master** (units × daily NAV from the public NSDL-CRA file, type-ahead scheme dropdown, add-only merge by scheme code, approximate two-flow XIRR) | An SGB row with a real ISIN prices from bhavcopy and its XIRR includes coupons (hand-checked golden); a physical-gold row values grams × purity × rate and a typed override wins; rate fallback chain exercised by fixtures (IBJA down → proxy median; both down → previous rate kept + warning); NPS units × NAV refresh from a fixture file, scheme type-ahead works over sorted NPS_Master; pie/targets/History include new classes only when enabled; round-trip extended |

## v1.4 — "Restructures"

*For the user: when companies they hold merge or split up, their holdings
adjust themselves correctly — no manual maths, and a clear audit trail.*

**Shipped as v1.4.0 (2026-07-17)** — R14 landed with scenario tests
(`test_r14_restructures.py`; Invested conserved to the rupee) and a live
check: a pre-merger HDFC Ltd lot priced via HDFC Bank at the real 42:25
ratio **times HDFC Bank's real 2025 split** (chained factor 3.36 — 25 old
shares → 84 new, matching a real demat). v1.4.0 also ships the onboarding
polish: sample rows in every class sheet (delete what you don't own and the
workbook slims itself), a console show/hide prompt for asset classes, a
colourful live console with progress lines, an always-printed version line,
and the refreshed README.

| # | Milestone | Delivers | Acceptance criteria |
|---|---|---|---|
| R14 | Mergers / demergers / ISIN reassignments | Curated `data/restructures.csv` shipped with releases (MERGER / DEMERGER / ISIN_CHANGE; ratios + cost apportionment; Σ cost_pct = 100 validated); engine: price-routing via successor ISIN + ratio folded into the existing Adj factor for mergers/ISIN changes (Invested and Cost date carry — Sec. 47); demergers: new Equity "Cost factor" column + append-once child rows with inherited cost dates, `Applied` date on the Corporate_Actions row as the idempotency token; Corporate_Actions sheet gains New ISIN / Cost % / Applied + `Curated` source; Manual row with same key overrides Curated; Stock_Master `Merged`/`Renamed` status suppresses Suspended/Delisted escalation; unlisted children amber `awaiting listing` | Merger 1:2 → priced via new ISIN, Adj factor 0.5, Invested unchanged, audit row Curated+Applied, no false Suspended, second run idempotent; ISIN change → price follows new ISIN; demerger 60/40 → parent Cost factor 0.6, child appended once with inherited cost date, **parent Invested×T + child Invested = original Invested to the rupee**, deleting the child survives the next run, unquoted child shows amber and prices on first quote; Manual override replaces curated ratios; Σ cost_pct ≠ 100 fails loudly at load; raw user rows byte-identical (appended flagged rows aside); round-trip extended |

Post-v1.4: see [ROADMAP.md](ROADMAP.md).

## v1.4.1 — "Review hardening"

*For the user: the same features, with the sharp edges filed off — cost
bases stay right through chained restructures, dividend income can't be
double-counted or mis-parsed, and one flaky feed can no longer undo work
already done.*

**Shipped as v1.4.1 (2026-07-17)** — every confirmed finding of the
2026-07-17 full-diff code review fixed, each with a regression test
(`tests/test_review_fixes.py`, 16 tests; suite at 154). Highlights:

- **Restructure engine edges** — a second demerger apportions what
  *remained* (no more 112%-of-original cost); demergers on a
  merger-successor reach lots held via the old ISIN; children are computed
  AFTER the corporate-actions refresh, apply atomically, defer (unstamped)
  when history can't be verified or the Equity sheet is full; merged
  holdings' successor symbols are queried so their dividends arrive
  (live-verified: the legacy HDFC Ltd lot now earns HDFC Bank's real
  ₹13/share on 84 chained shares).
- **Dividend parsing** — face-value wordings and word-tail "re" no longer
  parse as rates; cross-exchange dedupe keys on `(isin, ex_date, rate)` so
  differing type wordings can't double-count a payout.
- **Feed resilience** — a one-symbol corporate-actions failure keeps that
  stock's applied rows and current-FY dividends; an NSE 200-but-HTML
  bot page degrades the day to BSE-only instead of aborting.
- **Sheet safety** — Corporate_Actions capacity ×4 (200 rows) with
  Manual/Curated rows written first and loud, oldest-first overflow;
  Dividends overflow protects Manual/current-FY rows; restructure flags
  and the FMV marker coexist in the Flags cell.
- **Correctness odds & ends** — a 0.0 demerger cost factor is honoured in
  XIRR; Manual_Assets class labels match case-insensitively (unknown labels
  warn); the frozen exe bundles all six runtime CSVs.
- **Onboarding** — new classes ship Settings **No** (visible while their
  sample rows exist), so the Guide's "delete the samples and the tabs tidy
  away" promise now actually works; the still-holds-rows warning fires at
  toggle time instead of nagging every run.

## v1.4.2 — "The rest of the review"

*For the user: small honesty fixes — the sheet, the console and the docs
now say exactly what the numbers are doing.*

**Shipped as v1.4.2 (2026-07-17)** — the review's lower-severity findings,
each with a regression test (`tests/test_v142_polish.py`; suite at 163):

- "Avg cost today" applies the demerger cost factor, matching the docked
  basis a broker app shows.
- The Gold_Silver rates-as-on stamp advances only when a METAL rate
  actually arrived (an SGB-only day no longer hides a stale benchmark),
  and the stale-rate amber excludes SGB rows.
- The Dividends Qty comment now routes corrections through a Manual row
  (hand-edits to current-year Auto rows were silently undone).
- The console show/hide prompt reports EFFECTIVE visibility ("shown —
  holds rows; delete them to hide"), ending the "prompt says hidden but
  the tab is right there" confusion.
- `--pause` survives a closed stdin (the packaged entry always passes it;
  scheduled runs used to end in an EOFError traceback).
- NPS_Master rows are kept by Scheme Code — a blank PFM no longer drops a
  scheme on round-trip.
- `run()` no longer mutates caller-supplied restructure events; nameless
  restructure rows display their symbol instead of a raw ISIN; SPEC no
  longer describes the dropped GoldBeES proxy.

## v1.4.3 — "A calm first open" (shipped v1.4.3, 2026-07-17)

UX-only release: nothing new is fetched or computed differently for data the
user can see — the workbook just stops overwhelming at first glance.

| Change | Delivers | Acceptance criteria |
|---|---|---|
| Settings choice wins | A class switched off is hidden AND excluded everywhere (Dashboard, allocation, person tabs, Projection, portfolio XIRR, new History snapshots record 0) even when it holds rows; rows never deleted; one amber Dashboard notice (I1) + one updater summary line name the hidden money with its value | off-with-data class: sheet hidden, status `Hidden - has data (not counted)`, exactly one summary line, Dashboard I1 names it, rows survive round-trip; portfolio XIRR with class off == portfolio XIRR with its rows absent |
| Calm first open | Shipped template shows only the classic five (targets 40/15/20/15/10); EPF/Gold & Silver/NPS/Property/Cash/Insurance/Other ship hidden with sample rows waiting inside | fresh template hides exactly those class sheets + the reference sheets; enabling a class reveals its samples |
| Reference lists switch | The four masters + Corporate_Actions hide behind ONE Settings row (16); dropdowns/lookups keep working against hidden sheets | switch No → all five hidden; Yes → all visible; round-trips |
| Tab colours + Guide pointer | Colour-coded tab strip (navy/teal/blue/grey/gold); "New here? → Guide" note on the Dashboard | tabColor present in sheet XML per group |
| Property, plainly | "Real Estate" class renamed **Property**; old label accepted everywhere on read (Settings row, Class cells, History header, allocation) | old-label workbook round-trips into Property with settings/history/xirr intact |
| Clearer Manual Assets & Gold | Purpose-first sheet hints, per-column examples in comments, generic samples (apartment, coins, bars — no more Pune flat / bangles); Guide gains a 3-step "adding gold" box | comments/samples assert; Guide text mentions the steps |
| Simpler Settings sheet | `Show?` column wording, plain statuses (`Shown`/`Hidden`/`Hidden - has data (not counted)`), "Balance targets (optional)" section, tolerance/total moved to rows 18/19 | layout asserts + reader tolerance |

Docs in the same commit: SPEC (§2.1, §3.1–§3.3, §3.14–§3.18, §4, §6.2,
§6.8, §6.11, §7), README, packaged README.txt, Guide sheet, release note.
Suite: 170 tests.

## Release artifact layout (from R4)

```
NetWorth-<version>-windows.zip        # assembled on Linux, no Wine/compiler
└── NetWorth-<version>-windows/
    ├── Family_Portfolio_Tracker.xlsx
    ├── Update Portfolio.bat          # double-click launcher (%~dp0-relative)
    ├── README.txt                    # 1-page quick start
    └── app/                          # private embeddable CPython + code + data
NetWorth-<version>-macos.zip          # built on macOS via PyInstaller
├── Family_Portfolio_Tracker.xlsx
├── Update Portfolio.command          # tiny wrapper
├── networth-updater                  # PyInstaller binary
└── README.txt
```

Windows build (`packaging/build-windows-nowine.sh`) downloads the Windows
embeddable CPython + prebuilt `win_amd64` wheels and arranges them — no
compilation, no Wine. A single self-contained `.exe` needs a real Windows box
(`packaging/build-release.bat`) or CI; on Linux the launcher is a `.bat`
because a relocatable Windows `.exe` can't be execute-verified there.
