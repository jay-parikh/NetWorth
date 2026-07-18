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

## v1.5.0 — "Privacy: Mask + Lock" (shipped v1.5.0, 2026-07-17)

Two opt-in privacy layers sharing one password (SPEC §3.19); all four
Mask×Lock combinations are legal states. Design reviewed with an explicit
threat model and leak-vector audit.

| Change | Delivers | Acceptance criteria |
|---|---|---|
| Privacy mask (curtain) | Every number renders ••• (3-section literal format — no `-•••` sign leak; dates/text visible); all sheets protected with selection disallowed (kills status-bar SUM/copy/Go To); charts, data bars, red/green and icon sets suppressed (each chart replaced by a grey note); values intact underneath | masked build: 0 chart XMLs, 10 notes, no dataBar/iconSet/conditionalFormatting, sheetProtection+selectLockedCells everywhere, reader gets full values, round-trip identity holds |
| Password fingerprint | pbkdf2-sha256 (200k) in defined name `NW_Privacy` + at-rest state in `NW_Masked`; password itself never stored, never echoed (getpass), never in argv | hash round-trips; verify accepts/rejects; RESET clears both |
| Lock (encryption at rest) | Standard OOXML Agile encryption via msoffcrypto-tool 6.0; Excel/LO prompt natively; updater needs the password even to read; build-in-memory + encrypt + **self-verify** before atomic replace — plaintext never on disk; backups byte-copy ciphertext | encrypted file is CFB magic; wrong pw/headless exits with file byte-identical; decrypt+read identity; unconfirmed-password enable is refused (no lockout by typo) |
| Flows | First-enable set-password (loud no-recovery warning for Lock); locked update prompt (3 trial-decrypt tries); mask prompt (show / Enter=keep / RESET); both-layers "show this time? y/N"; `--lock` offline re-mask/re-encrypt; unmasked view-backups auto-purged when the mask returns | lifecycle tests cover all four states + transitions; relock provably fetches nothing |

Docs in-commit: SPEC §3.14 rows 20–22 + new §3.19 + §7; Guide privacy
section; README; packaged README.txt. Suite: 181 tests.

## v1.5.1 — "Everything explains itself" (2026-07-18)

UX pass driven by Jay's design review: the workbook must be self-explanatory
without renaming a single term. Policy is **keep + gloss** (SPEC §3 preamble):
domain terms — compute, NAV, ISIN, XIRR, corpus, PRAN, UAN, SGB, coupon,
ex-date — are correct and stay verbatim everywhere; hover comments lead with
the term and explain it in plain words.

| Change | Delivers | Acceptance criteria |
|---|---|---|
| Guide redesign | 7 tighter sections (~69 rows, was 83), frozen title banner, "Words you'll see" glossary; every instruction of the old Guide retained (52-atom sweep) | glossary + reference-lists + gold how-to + "Your choice wins" probes pass; freeze at A3 |
| Header glosses | Shared gloss strings (`_G_XIRR`, `_G_ISIN`, `_G_NAV`, `_G_CURVAL`, `_G_NETCHG`) + per-sheet comments on PRAN, UAN, SGB, FY, Ex-Date, Comp./yr, Coupon, Face Value, Drift, Real return, Corpus, # holdings — same term, same explanation, every sheet | test_v151_gloss: headers verbatim, comments lead with the term |
| Obscure text rewritten | "IBJA; market-implied fallback", "LTCG grandfathering", "monthly-minimum-balance rule", "Alt+Down", "Yellow-ish" banner — all replaced by plain words (terms like XIRR/ISIN kept) | banished-phrase sweep in test_v151_gloss |
| Privacy prompts simplified | Mask prompt is staged (one plain question; password only on "y", verified on the spot with 3 tries; RESET confirm separated); wrong password is never silent | test_mask_prompt_staged_flow + test_wrong_password_is_never_silent |

Docs in-commit: SPEC §3 keep+gloss preamble + §3.11 Guide + §7 prompt flow;
README (badge 185, R0→v1.5); release notes v1.5.1. Suite: 185 tests.

## v1.6.0 — "Honest returns + the taxman's view" (2026-07-18)

Two roadmap items Jay picked for the release: dividends (and now recorded
sales) enter the equity XIRR, and a capital-gains tax report with the
31-01-2018 grandfathering rule. Everything keep+gloss (STCG/LTCG stay
verbatim, hover comments explain), everything default-off (Settings row 17),
everything indicative — "for planning, not for filing" on the sheet itself.

| Change | Delivers | Acceptance criteria |
|---|---|---|
| Dividends + sells → equity XIRR (§6.2) | Dividend rows (all FYs, ex-date ≤ today) and complete Equity_Sells round trips append to the equity flows — the return is genuinely money-weighted; portfolio XIRR inherits; per-person "Dividends FY" SUMIFS on each person sheet row 4 | golden XIRR with a dividend > bare two-flow; future/incomplete rows excluded; all-sold ⇒ None; flow-order pins (test_r7/test_review_fixes) still green; Dashboard B17 untouched |
| Equity_Sells input sheet (§3.20) | Self-contained sale records in sell-time units (contract-note view); ISIN lookup + type-ahead; Proceeds/Gain formulas; double-entry banner ("also reduce the Quantity on the Equity tab"); worked samples incl. a blank-buy-price grandfathering demo | round-trips byte-equal through build→read→build; guarded reader (pre-v1.6 = no-op); default hidden |
| Capital-gains engine (§6.16) + tax_rules_in.csv | Grandfathering = max(cost, min(FMV, sale)) with FMV normalised across post-2018 corp actions; MF FIFO with oversell/nav-less warnings (never guess); Debt/slab split at 2023-04-01; equity + mf_equity share ONE §112A exemption; per-sale-date STCG rates across the 2024-07-23 mid-FY switch; `lt_on` sell-planning dates | 25 goldens in test_v160_capgains/test_v160_dividend_xirr incl. split-normalised FMV, mid-FY 15%/20% mix, shared-bucket headroom 0 at ₹1.4L LTCG |
| Capital Gains sheet (§3.21) + Settings switch | Computed at BUILD time (round-trip identity untouched); FY summary first, realised detail, sell-planning; headroom headline with absolute deadline; Settings row 17 default No hides the pair together; MutualFunds M "Tax type" dropdown | sheets hidden by default, visible on Yes; masked build: protected, charts stay 10, no ₹ composed into text cells; self-explanatory sweep (banners, sample rows, gloss on every new jargon header) |

Docs in-commit: SPEC §3.1/§3.5/§3.7/§3.14/§3.20/§3.21/§5.5/§6.2/§6.6/§6.16/§7;
Guide "Selling & tax" section + STCG/LTCG/Grandfathering glossary; ROADMAP two
items ✅ + curated-data cadence recorded (refresh restructures.csv +
bullion_proxies.csv every release); release notes v1.6.0.

Pre-release hardening (2026-07-18): a full multi-angle code review of the
release produced 15 confirmed findings — all fixed in-release (engine gated
+ degrades to warnings instead of crashing on a bad bundled CSV; typed-0
prices honoured; STCG tax netted; scheme-level Tax-type consensus; FIFO ₹1
dust threshold; pre-2024 debt-fund 1095-day rule row; FY labels + relock
pinned to the run date; overflow/no-Owner warnings; honest same-day-sale
wording; sample XIRR re-captured) — plus 12 regression tests (suite: 222).

Also in v1.6.0: **docs/USER-GUIDE.md** — the complete illustrated user
guide (19 screenshots rendered from the shipped sample workbook via the
reproducible scripts/guide_screenshots/ pipeline, worked examples per
feature), linked from the README and the release notes.

Also (Jay, 2026-07-18): **charts never sit on data** — the Dashboard charts
now anchor right of the person × class grid (grid-width-derived column, so
enabling more classes slides them right instead of hiding columns) and the
person-sheet pie moved right of the holding blocks (§3.3/§3.5); screenshots
re-rendered.

Also (Jay, 2026-07-18): **intraday sales surface as speculative income** —
a same-day buy/sell is no longer warned-and-skipped: it becomes a realised
row (bucket "speculative", term "Intraday") plus an "Intraday gains ₹" FY
column, slab-taxed business income shown but never mixed into STCG/LTCG or
XIRR (§6.16). Plus edge-case hardening from a systematic sweep: negative
qty/price rows warn ("check the row") in both the engine and XIRR;
Tax_Rules rejects typo numbers (rates outside 0–100, negative allowance,
non-positive holding days) and warns on duplicate (asset, date) rows;
negative NAVs count as "no NAV" in the MF FIFO. Suite: 229.

And (Jay, 2026-07-18): **Tax_Rules in the workbook (§3.22)** — the
capital-gains rate table (rates, holding periods, the ₹1.25L allowance) is
no longer release-bound: an editable Tax_Rules sheet ships prefilled with
the bundled law and is upserted over the CSV defaults by (asset,
applies-from date). A Budget change is an Excel edit; invalid rows warn and
are preserved, never guessed at or dropped.

## v1.6.1 — "Losses count, the way the law allows" (2026-07-18)

Stability patch — exactly one correctness fix, found by re-verifying the
items deferred during v1.6.0 (Jay: simple and stable, no new features).
A wrong tax number is a stability bug in a tax report.

| Change | Delivers | Acceptance criteria |
|---|---|---|
| Sec 70(2) same-FY set-off (§6.16) | A short-term loss left over after the short-term netting reduces the same year's LTCG **before** the §112A allowance — previously the excess was silently discarded and the sheet overstated tax in a loss-harvesting year. Shown in its own By-FY column "ST loss used vs LTCG ₹" (blank when zero); the LTCG column stays raw; `headroom_now` uses the post-set-off figure | net ST loss 2L + LT gain 3L ⇒ set-off 2L, allowance used 1L, headroom 25k, both taxes 0; excess loss capped at the FY's LTCG, never carried to another FY; speculative losses never feed it (Sec 73); losses < gains ⇒ nothing changes |

The other three deferrals were closed as won't-do — verdicts and reasons
live in ROADMAP.md (the one place; don't restate them elsewhere).

Pre-commit hardening (2026-07-18): a full multi-angle review of this patch
produced 15 findings, all addressed in-release — the set-off is era-gated
(no set-off shown for pre-2018 §10(38) FYs, whose LTCG was exempt), float
dust clamps to zero so "blank when zero" holds, each short-term row is
taxed at its own asset's Tax_Rules rate (an mf_equity row no longer
inherits equity's rate if a user diverges them), masked builds write the
set-off cell on every row so its presence can't leak a loss-harvest year
through the mask, the updater console prints the set-off beside the raw
figures, one shared `ltcg_eff` derivation feeds both the FY row and the
headline, the Sec 70(3) citation was corrected, and every user-facing
sentence was narrowed to claim exactly what the engine does (no
cross-bucket equity↔debt netting — now stated on the sheet). Suite: 236.

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
