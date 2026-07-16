# NetWorth — v1.2–v1.4 plan (approved 2026-07-16)

Dual-exchange market data, dividend income, the whole balance sheet, and
corporate restructures. This document does for v1.2–v1.4 what
[PLAN.md](PLAN.md) did for v1: it records the approved design and the *why*,
so implementation sessions start from decisions, not debates. Milestones and
acceptance criteria live in [RELEASES.md](RELEASES.md); the normative spec
([SPEC.md](SPEC.md)) is updated **in the same commit as the code**, per the
standing rule — this plan lists where each addition will slot in.

## Context

v1.1.0 tracks five asset classes (Equity, MF+SIP, FD, PPF+ledger, Bonds) with
BSE-primary/NSE-fallback prices and split/bonus/consolidation corporate
actions. Seven requirements were raised on 2026-07-16:

1. NSE as a full peer price source (today fallback only).
2. A dividend sheet logging all dividends on held stocks per financial year.
3. Gold/SGB/Silver as an asset class with a live daily metal rate.
4. More asset classes: NPS, EPF, real estate, cash/savings, insurance
   (surrender value), ESOPs/RSUs.
5. Asset-allocation targets — target % per class, drift red/green,
   rebalancing hints.
6. Mergers / demergers / ISIN reassignments auto-adjusted.
7. User-selectable asset classes — only applicable sheets visible, the rest
   hidden, for a neat workbook.

## Locked decisions (from user answers, 2026-07-16)

- **Sequencing: balance sheet before restructures.** v1.2 = data quality
  (R8–R9), v1.3 = balance-sheet expansion (R10–R13), v1.4 = restructures
  (R14). Manual Corporate_Actions rows already cover restructures as a
  workaround; missing asset classes have no workaround at all.
- **Metal rates come from the bullion trade's own benchmark.** Primary source
  is the IBJA daily rate (999 fine gold/silver — the rate bullion dealers and
  jewellers quote from; RBI uses IBJA 999 for SGB redemption), with a
  bhavcopy-implied fallback and a manual override column that always wins.
- **ESOPs/RSUs are deferred entirely** (2026-07-16): foreign-listed grants are
  a multi-currency problem; the class stays on [ROADMAP.md](ROADMAP.md) until
  multi-currency lands. Indian-listed vested RSUs simply go on the Equity
  sheet (Guide note).
- **Docs now, spec with code.** This plan + RELEASES.md rows are written
  up-front; SPEC.md sections are written per-milestone alongside the
  implementation, so spec and code never drift.
- **Ease of use is an acceptance criterion, not a nicety** (2026-07-16):
  every feature must keep the workbook easy, clean, simple and beautiful for
  a non-developer — see the design principles below, which bind all of
  R8–R14.

## Design principles — easy, clean, beautiful (bind every milestone)

The end user is a non-developer who opens one workbook and double-clicks one
updater. Nothing in v1.2–v1.4 may change that feel:

- **Clean by default.** A fresh workbook shows only the classic five classes
  plus Settings. New tabs, columns and charts appear only after the user
  switches a class on with a simple Yes. More features must never mean a more
  crowded first impression.
- **One glance, one answer.** The Dashboard stays the single place to look —
  dividends this year, allocation drift, every class total. Anything that
  needs study rather than a glance lives on its own sheet.
- **Plain words on the sheet.** Every new sheet gets a one-sentence hint in
  row 2 saying what to type and what fills in automatically (e.g. *"Type
  what you own. Rates and values fill in when you run the update."*).
  Headers use everyday words ("Rate today", "Value", "Move ₹ in/out"), not
  abbreviations or finance jargon. The Guide sheet gains a short,
  plain-language section per feature.
- **Colour says what words would.** The existing language is kept everywhere:
  green = gain, red = loss, amber = estimated or stale — and every amber cell
  carries a comment explaining *why* in one sentence.
- **Never scary, never silent.** Updater messages are plain sentences
  (*"Could not fetch today's gold rate — kept the last one from
  12-07-2026."*). Warnings say what to do next. Nothing fails without saying
  why in words a non-developer understands.
- **Rich, never busy.** Clean does not mean bare — the workbook should feel
  alive and rewarding to open, never boring. Each milestone ships at least
  one visual that makes its numbers *felt*, drawn from a shared vocabulary
  (all native xlsxwriter, no add-ins): charts (pie, column, line, stacked
  area), data bars on value columns, ▲/▼ icon-set arrows on change columns,
  and the green/red/amber colour language. Concretely across v1.2–v1.4: a
  "Dividends by month" chart (R9), a "Net worth by class over time" stacked
  area chart (R10), the "Actual vs Target %" chart (R11), data bars on the
  allocation table and ▲/▼ arrows on day-change (R10), and every new class
  joining the pies, trend and per-person charts automatically (R12–R13).
  The balance rule: visuals live on the Dashboard and summary areas, data
  sheets stay quiet workspaces — rich where the user looks, calm where they
  type.
- **Nothing new to learn.** Same blue input cells, same type-ahead dropdowns,
  same one-click updater. If a feature would require the user to learn a new
  skill, redesign the feature.

## Master release plan

| # | Milestone | Ships as | Effort |
|---|---|---|---|
| R8 | NSE as full peer price source | v1.2 | S |
| R9 | Dividends sheet (FY dividend ledger) | v1.2 | M |
| R10 | Settings sheet + asset-class registry + selectable classes | v1.3 | M |
| R11 | Asset-allocation targets (drift red/green, rebalance hints) | v1.3 | S |
| R12 | New classes wave 1: Manual_Assets (RE / Cash / Insurance / Other) + EPF | v1.3 | M |
| R13 | New classes wave 2: Gold_Silver (SGB + physical bullion) + NPS | v1.3 | M |
| R14 | Mergers / demergers / ISIN reassignments (curated file + engine) | v1.4 | L |

**Why this order.** R8 is foundational for everything after it: it stops
discarding the BSE scrip codes that the corporate-action/dividend lookups
need (R9), it brings NSE's SGB quotes into the merged bhavcopy (R13's gold
rate fallback), and it makes "absent from bhavcopy" mean absent from *both*
exchanges — fewer false Suspended flags that are really mergers (R14). R10's
class registry is the plumbing R11 stores its targets in and R12/R13 register
their classes with; building the new classes first would hardcode seven more
classes into the ~8 modules that already hardcode five, then refactor them
anyway. R14 goes last: it is the largest and the only one with an ongoing
data-curation duty.

## Feature specifications

### R8 — NSE as full peer price source

**What the user sees:** nothing new to do or learn — more of their stocks
simply get a price, and the prices match what their broker app shows.

Today `fetch()` in `fetch/bhavcopy.py` walks back up to 7 days and, per day,
tries BSE then NSE — **first hit wins**, and when NSE wins, `codes_by_isin`
(BSE scrip codes) is cleared, silently degrading the BSE corporate-action
lookups for that run.

**Design: same-day union merge; NSE wins on price conflicts; BSE keeps the
scrip codes.**

- Per candidate day, attempt **both** exchanges. If at least one succeeds,
  merge and stop walking back — never mix trade dates across exchanges (one
  `trade_date` keeps Closing-Price-Date and day-change semantics coherent).
- Merge rule: union of ISINs. On a dual-listed conflict, **NSE close/prev
  win** — NSE carries ~90% of cash-market turnover and is what users' broker
  apps show, so day-change matches what the family sees elsewhere.
- `codes_by_isin` is taken **exclusively from the BSE parse and retained
  whenever BSE responded** (the "NSE wins ⇒ clear codes" branch is deleted).
  Master rows: dedupe by ISIN; for *new* ISINs prefer the NSE symbol (it is
  what the NSE corporate-action API needs); the add-only merge already
  protects existing rows.
- Source labelling is **per run, not per cell**: summary reads e.g.
  `Prices: BSE+NSE 15-07-2026 (3 NSE-only)`. A family tracker does not need
  per-cell provenance.
- **Delisted/stale detection amendment (SPEC §6.5):** the Suspended (>21d) /
  Delisted (>180d) escalation runs **only on dual-source days**. On a
  single-source day (one exchange down or not yet published): quoted → Active
  as usual; unquoted → status carried forward untouched. Skipping escalation
  on degraded runs costs nothing (thresholds are in days) and prevents an
  NSE-only scrip being marked Suspended during a BSE outage.
- Failure modes: one exchange down → single-source run + summary note, no
  escalation; both down → walk back as today; all 7 days dry → existing
  RuntimeError and graceful updater degradation.

Files: `fetch/bhavcopy.py` (merge), `update.py` (escalation guard, summary),
SPEC §5.2/§5.3 (merged into one dual-source contract), §6.5, §7.

### R9 — Dividends sheet (FY dividend ledger)

**What the user sees:** a new Dividends tab that reads like a passbook — one
row per dividend their shares paid this financial year, already filled in —
and one Dashboard line, "Dividends this FY".

**Zero new HTTP.** The corporate-action fetcher already downloads every
announcement for every held symbol/scrip-code from both exchanges and
*discards* dividends. R9 adds a `parse_dividend(subject)` alongside
`parse_subject`:

- keyword `dividend`; sub-type from `interim|final|special` (default Final);
- rate from `Rs./Re./₹ N(.NN) (per share)?` variants, including BSE's
  `"Interim Dividend - Rs. - 5.5000"` and NSE's `"Dividend - Rs 8 Per Share"`;
- percent-of-face forms ("Dividend 250%") are **skipped in this milestone**
  (rare post-2010; a Manual row covers them); skipped-subject count in the
  run summary;
- dedupe on `(isin, div_type, ex_date)`, NSE wins — same rule as corporate
  actions.

**New `Dividends` sheet** (tab after Corporate_Actions; title r1, hint r2,
header r3, data r4+):

| Col | Header | Kind | Definition |
|---|---|---|---|
| A | FY | updater | e.g. `2026-27`, from ex-date (Indian FY, Apr 1–Mar 31) |
| B | Owner | updater (Auto) / input (Manual) | |
| C | Scrip | updater / input | |
| D | ISIN | updater / input | |
| E | Type | dropdown | Interim / Final / Special |
| F | Ex-Date | date | |
| G | Rate ₹/share | updater / input | |
| H | Qty @ ex-date (est.) | updater | see algorithm below; amber "(est.)" |
| I | Est. amount | computed | `=G×H`, guarded; amber |
| J | Source | updater | Auto / Manual |
| K | Details | updater | announcement free text |

The Dashboard gains one info cell: **Dividends this FY** =
`SUMIFS(Dividends!I:I, Dividends!A:A, <current FY>)` (family total;
per-person split is a roadmap nicety). The Dividends sheet gets a small
**"Dividends by month"** column chart (12 SUMPRODUCT cells over the current
FY feeding a native chart) — income the user can *see* arriving, not just a
table.

**Qty-at-ex-date estimate (planned SPEC §6.12)** — stated honestly: there is
no sell ledger, so the estimate projects *current* Equity rows backwards.

```
qty_est(owner, isin, ex) =
  Σ over Equity rows r with r.isin == isin, r.owner == owner,
    and (r.cost_date blank OR r.cost_date < ex):
      r.qty_raw × adjustment_factor(isin, r.cost_date, ex − 1 day, actions)
```

Reuses `model.adjustment_factor` evaluated as of the day before ex-date, so a
bonus/split between purchase and dividend ex-date is respected. Blank
cost-date lots count as held (consistent with the FMV-era treatment). The
limitation — rows deleted/edited after a sale make history wrong — is stated
in the sheet hint and the Guide, and H/I render amber.

**Idempotency & FY rollover:** one sheet with an FY column — no per-FY tabs.
Each run rebuilds **Auto rows of the current FY only** from feed + holdings;
every other row persists unchanged. Prior-FY Auto rows therefore freeze
automatically on the first run after Apr 1 (multi-year dividend history for
free — users need it at ITR time). Manual rows persist always and override an
Auto row with the same `(isin, div_type, ex_date)` key.

**Explicitly out of scope:** dividends do **not** feed equity XIRR in this
milestone. Changing return semantics deserves its own release; the "dividends
→ true equity XIRR" item stays on the roadmap.

Files: `fetch/corporate_actions.py`, `model.py` (`DividendRow`,
`PortfolioData.dividends`), `update.py`, `generate.py` (`_write_dividends`),
`reader.py`, Guide text; SPEC §3.14, §5.4, §6.12, §7.

### R10 — Settings sheet, asset-class registry, selectable classes

**What the user sees:** one simple Settings tab — Yes or No beside each asset
class — and the workbook shows only the tabs they actually use. This is the
"neat Excel" requirement, and it is what keeps the workbook feeling small
even as the product grows.

**The refactor everything hangs on.** The five classes are hardcoded in ~8
places: the Dashboard matrix (`generate.py:175`), the allocation table,
`_PERSON_BLOCKS` (`generate.py:281`) and the person summary rows,
`ClassXirr`/`HistorySnapshot` fields (`model.py`), `compute_all_xirr`
(`compute/cashflows.py:156`), `net_worth_snapshot` (`compute/snapshot.py:31`)
and the reader's fixed Dashboard cells. Adding seven classes by copy-paste
would roughly double that surface and make per-class visibility
unimplementable.

**Design: an `ASSET_CLASSES` registry in `model.py`** — the pattern the
persons matrix already proves, applied to classes:

```python
@dataclass(frozen=True)
class AssetClass:
    key: str              # "equity", "gold_silver", "real_estate", …
    label: str            # Dashboard/History header
    sheet: str            # data sheet, e.g. "Manual_Assets"
    value_col: str        # column summed by Dashboard/person SUMIFS
    owner_col: str = "A"
    class_filter: tuple[str, str] | None = None  # (col, label) on shared sheets
    default_enabled: bool = False
    has_xirr: bool = True
```

`ClassXirr` and `HistorySnapshot` become dict-backed (keyed by class key).
Dashboard matrix columns, allocation-table rows, person summary rows and
blocks, History columns, chart ranges, snapshot sums and cashflow dispatch
all iterate the registry filtered by *effective enablement* (below). A
shared-sheet class is just an extra SUMIFS criterion pair:
`=SUMIFS(Manual_Assets!$G:$G, Manual_Assets!$A:$A,$A6, Manual_Assets!$B:$B,"Real Estate")`.

**New `Settings` sheet** (tab after Projection, before person sheets):

| Cell(s) | Content | Kind |
|---|---|---|
| A1 | `SETTINGS` | static |
| A2 | hint: Yes/No shows or hides each asset class; a class holding data is never hidden; Target % feeds the Dashboard drift view | static |
| A3:E3 | headers `Asset class, Enabled, Target %, Status, Notes` | static |
| A4:A15 | class labels, registry order | static |
| B4:B15 | `Yes` / `No` dropdown (non-blocking, house style — not form-control checkboxes: xlsxwriter cannot make them and LibreOffice renders them poorly) | **input** |
| C4:C15 | target allocation %, blank = no target | **input** |
| D4:D15 | generator-written status: `On`, `Off`, `On (has data)` | computed value |
| A17/B17 | `Drift tolerance (± % points)` / default **5** | **input** |
| A18/B18 | `Targets total` / `=SUM(C4:C15)`, amber when non-blank and ≠ 100 | computed |

`PortfolioData` gains `class_settings` (enabled + target % per class), read
tolerantly (labels matched by name in rows 4–20; missing sheet → defaults).
**Defaults: the classic five = Yes, every new class = No** — the shipped
template looks exactly like v1.1 plus a Settings tab until the user opts in.

**Hide, not omit.** The generator always writes every class sheet (with the
user's data) and calls `worksheet.hide()` on disabled ones:

1. *Round-trip safety* — openpyxl reads hidden sheets identically, so data
   typed before disabling survives every rebuild with zero reader
   special-casing.
2. *Formula integrity* — no dangling references, no `#REF!`.
3. *Reversibility* — flip No→Yes (or just unhide in Excel) and everything is
   still there.

What *is* omitted for a disabled class (registry-driven, computed column
letters): its Dashboard matrix column, allocation-table row, person summary
row and holdings block, History column, and chart ranges. Shared sheets
(Manual_Assets) hide only when **all** their subclasses are off; dedicated
masters (NPS_Master) hide with their class; Dashboard and Guide never hide.

While the charts become registry-driven, R10 also upgrades the Dashboard's
visual richness (the "rich, never busy" principle): the net-worth trend line
gains a **stacked-area "Net worth by class over time"** companion built from
the same History columns, the allocation table's Value column gets **data
bars**, and day-change cells get **▲/▼ icon-set arrows** — all native
xlsxwriter, all sized to the surviving (enabled-classes) layout.

**Never lose data:** `effective_enabled = enabled OR has_data`. A class with
data toggled to No stays visible with Status `On (has data)` plus a note
("delete its rows to hide it"), and the updater prints a warning. A hidden
ghost value can never haunt the Dashboard total.

**Reader changes:** Settings block (tolerant, defaulting); the Dashboard
class-XIRR and FY-expected reads move from fixed cells to **header-located**
lookups — required once columns are dynamic, and it keeps v1.1 workbooks
readable (missing Settings → defaults; fixed persons/inflation cells
unchanged).

Files: `model.py`, `generate.py` (largest diff — de-hardcoding
`_PERSON_BLOCKS` and the Dashboard matrix), `reader.py`,
`compute/cashflows.py`, `compute/snapshot.py`, `compute/projections.py`;
SPEC §2 (class table becomes normative configuration), §3.13 Settings, §7.

### R11 — Asset-allocation targets

**What the user sees:** the Dashboard answers "is my money balanced the way I
planned?" at a glance — green when a class is on target, red with a plain
hint like "Move ₹1,20,000 out" when it drifts, and a small
actual-versus-target chart beside the pie.

Small once R10 exists, and deliberately **pure live formulas** — drift stays
correct the moment the user edits a holding, without running the updater.

Extend the existing Dashboard allocation table:

| Col | Header | Definition |
|---|---|---|
| A | Asset class | existing |
| B | Value | existing |
| C | XIRR | existing (updater-written) |
| D | Actual % | `= value / family total`, guarded |
| E | Target % | from Settings column C (blank = no target) |
| F | Drift | `= D − E` in percentage points; three-band red/green vs tolerance |
| G | Rebalance hint | `"On target"` when `|F| ≤ tolerance`, else `"Move ₹X out/in"` where `X = |F| × family total` |

- **Tolerance is absolute percentage points** (±5 default, Settings B17).
  Absolute reads straight off the pie ("Equity 62% vs 55% target = 7pp out");
  relative bands over-trigger on small sleeves (a 2% gold sleeve at 3% is 50%
  relative drift but irrelevant money).
- Blank target ⇒ the row shows no drift/hint and no conditional format fires.
- The hint is deliberately class-level, gross and **pre-tax** (header comment
  says so) — lot-level/tax-aware selling belongs to the capital-gains roadmap
  item.
- One clustered column chart **"Actual vs Target %"** (series D and E, rows
  with a target only) beside the allocation pie.
- Sanity: targets summing ≠ 100 flags amber on Settings (B18).

Files: `generate.py` (formulas + chart + a banded sibling of `_redgreen`),
`reader.py` (targets round-trip via Settings, already read in R10);
SPEC §3.3, §6.14.

### R12 — New classes, wave 1: Manual_Assets + EPF (no new fetchers)

**What the user sees:** one friendly sheet for the things they value by hand
— the house, savings accounts, insurance — where they type a value and a
date and the Dashboard picks it up; and an EPF sheet that works exactly like
the PPF sheet they already know (copy the balance from the passbook, done).

**Sheet allocation principle:** a class earns a dedicated sheet only when it
needs class-specific columns, fetch plumbing, or per-row math. Real estate,
cash/savings, insurance surrender value and "other" are all "manual current
value + optional cost/date" — four separate tabs would be pure clutter, so
they share one sheet while remaining **separate registry classes** (own
Dashboard column, toggle, target and History column each) via the
`class_filter` SUMIFS mechanism.

**`Manual_Assets` sheet:**

| Col | Header | Kind | Definition |
|---|---|---|---|
| A | Owner | input | |
| B | Class | input | dropdown: Real Estate / Cash / Insurance / Other |
| C | Description | input | "2BHK Baner", "HDFC savings", "LIC Jeevan…" |
| D | Institution / Ref | input | bank, insurer, registrar |
| E | Invested / Cost | input (optional) | RE: purchase cost; Insurance: premiums paid; Cash: blank |
| F | Cost / Start date | input (optional) | XIRR anchor |
| G | Current value ₹ | input | market estimate / balance / surrender value |
| H | Value as-on | input (date) | amber when > 90 days old — manual values rot silently |
| I | Net chg. | computed | `G − E`, guarded, red/green |
| J | Notes | input | |
| K | Key | helper | |

**`EPF` sheet** — deliberately congruent with PPF (flat estimate first, exact
ledger later, the path PPF itself took):

| Col | Header | Kind | Definition |
|---|---|---|---|
| A | Owner | input | |
| B | Establishment / UAN | input | |
| C | Member ID | input | |
| D | Current Balance | input | from the EPFO passbook |
| E | Balance as-on | input (date) | passbook date |
| F | Rate % (ref) | input | pre-filled with the current EPFO rate from bundled `data/epf_rates.csv` (`fy_start, rate_pct`; annual, no API — release-refreshed, the `ppf_rates.csv` precedent verbatim) |
| G | Notes | input | |
| H | Balance today | computed | `= D × (1 + F/100) ^ YEARFRAC(E, TODAY())`, guarded — flat accrual; Dashboard sums H |
| J | Key | helper | |

EPF's real rule (monthly interest on running balance on EPF-wage
contributions, credited annually) needs a contribution ledger — that ledger +
exact accrual is a roadmap follow-up mirroring SPEC §6.10.

**XIRR models (SPEC §6.2 table additions):** RE / Insurance / Other =
`−Invested @ Cost date, +Current value @ today` (skip when E, F or G blank);
**Cash is excluded from XIRR entirely** (`has_xirr=False`, blank cell in the
allocation table); EPF follows the PPF flat path. FY-end projection: EPF
accrues at its own rate to FY-end; RE / Cash / Insurance / Other are held
flat — estimating property appreciation would be false precision (header
comment says so).

**History schema migration (SPEC §6.11 revision):** the History sheet's
columns become **label-keyed**. Write: `Date` + the label of every class that
is effective-enabled or has nonzero history + `Total`. Read: map columns by
header name (unknown ignored, absent → 0). A v1.1 History therefore reads
losslessly; old totals recompute identically; new classes show 0 in old rows
— which is true (they were not tracked then).

Files: `model.py` (rows + registry entries + `epf_rates` loader),
`generate.py`, `reader.py`, `compute/cashflows.py`, `compute/snapshot.py`,
`compute/projections.py`, `data/epf_rates.csv`; SPEC §3.17, §3.18, §5.5,
§6.2, §6.11.

### R13 — New classes, wave 2: Gold_Silver + NPS (new data contracts)

**What the user sees:** a Gold & Silver sheet where they type grams (and
purity for jewellery) and today's value appears at the bullion-market rate —
with a column to type their jeweller's rate instead if they prefer; and an
NPS sheet that works like the mutual-fund sheet they already know (pick the
scheme from a dropdown, type units, NAV fills in).

**`Gold_Silver` sheet:**

| Col | Header | Kind | Definition |
|---|---|---|---|
| A | Owner | input | |
| B | Type | input | dropdown: SGB / Gold / Silver |
| C | Description / Series | input | "SGB 2023-24 Ser II", "Bangles 22K" |
| D | ISIN | input | SGB only — drives bhavcopy pricing |
| E | Qty | input | SGB: units (1 unit = 1 g); metal: grams |
| F | Purity | input | blank = 1 (SGB always 1); 22K jewellery = 0.916 |
| G | Buy Price ₹/unit | input | per gram/unit |
| H | Buy Date | input | XIRR anchor |
| I | Rate today (auto) | updater | SGB: bhavcopy close by ISIN; metal: ₹/g fine-metal rate (below) |
| J | Rate override | input | manual ₹/unit — **always wins over I** |
| K | Cur. val | computed | `= E × F(=1 if blank) × (J if set else I)`, guarded |
| L | Invested | computed | `= E × G`, guarded |
| M | Net chg. | computed | `K − L`, red/green |
| N | Maturity | input | SGB (8y); blank for metal |
| O | Key | helper | |

- **SGBs need almost no new machinery**: they trade on the NSE/BSE cash
  market, so the R8-merged bhavcopy prices them by ISIN (SGB liquidity is on
  NSE — another reason R8 comes first). SGB XIRR reuses the bond coupon
  engine: 2.5% p.a. semi-annual on the row's Buy Price — a documented
  approximation (the statutory coupon is on issue price; an Issue-Price
  column is not worth the width).
- **Physical metal rate (planned SPEC §5.7) — layered, per the locked
  decision:**
  1. **Primary: IBJA daily benchmark rate** — 999 fine gold and silver, the
     rate the bullion trade itself quotes from (RBI uses IBJA 999 for SGB
     redemption). Normalised to ₹/gram, tolerant parse, browser-ish headers.
     Stated plainly in the spec: there is no committed public API, making
     this the flakiest contract in the product — hence the layers below.
  2. **Fallback: bhavcopy-implied median** over a bundled
     `data/bullion_proxies.csv` (`metal, match ∈ {symbol_prefix, isin}, key,
     grams_per_unit, note`): gold from quoted SGB tranches (₹/g by
     construction) + a Gold-ETF ISIN; silver from Silver-ETF ISIN(s) with
     grams-per-unit measured at release time (expense-ratio drift is why the
     file is release-refreshed — the `ppf_rates.csv` precedent). Median
     across proxies damps tranche-specific illiquidity. Typically 2–4% below
     the IBJA retail rate; the Guide says so.
  3. **Degrade:** both sources down → keep the previous rate + as-on date,
     amber when the as-on is > 7 days old, warn in the summary.
  4. **Manual override (column J) always wins.**
- Purity converts jewellery to fine-metal terms (`grams × purity × rate`).

**`NPS` sheet:**

| Col | Header | Kind | Definition |
|---|---|---|---|
| A | Owner | input | |
| B | PRAN | input | free text |
| C | Scheme | input | type-ahead dropdown over `NPS_SchemeList` |
| D | Scheme Code | computed | INDEX/MATCH on NPS_Master, manual override allowed (existing `_manual` pattern) |
| E | Units | input | from the CRA statement |
| F | Current NAV | updater | daily NAV by scheme code |
| G | Cur. val | computed | `= E × F`, guarded |
| H | Total contributed | input (optional) | enables approximate XIRR |
| I | First contribution | input (optional, date) | |
| J | XIRR (approx) | updater | two-flow when H+I present; header comment: "approximate — dated-contribution ledger is roadmap" |
| K | Key | helper | |

Plus **`NPS_Master`** (Scheme Code, Scheme Name, PFM, NAV, NAV Date) — sorted
by name (the dropdown sort rule), refreshed-stamp cell, **add-only merge
keyed by scheme code** (SPEC §6.4 pattern).

**NPS NAV contract (planned SPEC §5.6):** the NSDL-CRA public daily NAV file
(npscra.nsdl.co.in — plain CSV: date, PFM code/name, scheme code/name, NAV;
no key), NPS Trust site as fallback; tolerant column-by-header parse; keyed
by scheme code; exact URL pinned during implementation (both hosts have
shuffled paths before — the day-cache + degrade contract covers outages).

Files: `model.py`, `generate.py`, `reader.py`, `fetch/nps.py`,
`fetch/bullion.py`, `compute/cashflows.py`, `data/bullion_proxies.csv`;
SPEC §3.15, §3.16, §5.5, §5.6, §5.7, §6.2, §6.13.

### R14 — Mergers / demergers / ISIN reassignments

**What the user sees:** when a company they hold merges, demerges or changes
its ISIN, their holdings quietly become the new company's shares at the right
ratio and cost — no manual maths, no scary error — and the Corporate_Actions
tab shows in plain rows exactly what was adjusted and when.

No reliable free feed publishes swap ratios, so the design pairs a **curated
data file shipped with releases** (the `ppf_rates.csv` / FMV precedent) with
the existing Manual-row escape hatch.

**`data/restructures.csv`:**

```
ex_date, type ∈ {MERGER, DEMERGER, ISIN_CHANGE}, old_isin, old_name,
new_isin, new_name, new_symbol, ratio_from, ratio_to, cost_pct, details
```

- MERGER (old absorbed): `ratio_from:ratio_to` = new shares per old;
  `cost_pct = 100` (full cost carries — Sec. 47 tax-neutral).
- DEMERGER: one row per resulting security, grouped by `(old_isin, ex_date)`
  — a parent-retention row (`new_isin = old_isin`, 1:1, parent's retained %)
  plus one row per child (child shares per parent share, the company-notified
  income-tax cost apportionment). **The loader validates Σ cost_pct = 100 per
  event and fails loudly otherwise.**
- ISIN_CHANGE: 1:1, cost_pct 100.
- Scope: NSE-500-ish events likely to touch retail portfolios; curation is an
  ongoing release duty (stated in the spec). Anything missed → Manual row;
  **a Manual row with the same `(old_isin, type, ex_date)` key overrides the
  curated row.**

The Corporate_Actions sheet stays the single audit trail, gaining **New
ISIN**, **Cost %** and **Applied** columns and a `Curated` source value.
Curated rows are rewritten each run like Auto rows, *except* `Applied`
persists.

**Engine (planned SPEC §6.15) — user rows are never edited:**

- **MERGER / ISIN_CHANGE — no new rows.** For each equity row whose ISIN
  resolves to an applied event's `old_isin` (chain-resolved old→new→newer by
  ex-date, cycle-capped): the updater *price-routes* the row — the successor
  ISIN's close/prev/date land in the row's price cells; the merger ratio
  folds into the existing Adj-factor column via `adjustment_factor`; Invested
  (raw qty × cost) is untouched, **which is exactly right** — merger cost
  basis and holding period (Cost date) carry in full. The Flags column gets
  `MERGED→<new name>`; Stock_Master status becomes `Merged`/`Renamed`, which
  **suppresses** the §6.5 Suspended/Delisted escalation for consumed ISINs
  (their bhavcopy absence is expected, not distress). The row keeps the old
  display name (Stock_Master is add-only, so the lookup keeps resolving).
- **DEMERGER — the one case needing two mechanisms:**
  1. A new Equity column **T "Cost factor"** (updater-written, blank = 1):
     the Invested formula becomes `qty × cost × T`, and the parent row gets
     `T = cost_pct/100`. (The user's Avg-cost cell cannot be rewritten —
     that would mutate an input.)
  2. **Append-once child rows**, one per owner-lot × child security:
     `qty = parent qty (CA-adjusted) × ratio`, per-share apportioned cost,
     and **Cost date inherited from the parent lot** — Indian CGT: demerged
     shares inherit the holding period, which is what makes capital gains
     come out right. Rows are flagged `DEMERGER:<old_isin>@<ex_date>`.
     The event's **Applied date (persisted on the Corporate_Actions row) is
     the single idempotency token**: re-runs skip applied events, so a user
     deleting a child row (say, after selling it) is respected. The child's
     name/ISIN/symbol enter Stock_Master immediately (add-only compatible —
     it is a new ISIN), so the appended row's lookup resolves pre-listing.
- **Child not yet listed** (typically 4–12 weeks): blank price, amber
  `awaiting listing`, excluded from day-change, no escalation; it prices
  automatically on first bhavcopy appearance — R8's union maximises how soon.
- **Known accepted wart (documented):** By-Scrip and person-sheet equity
  blocks group merged rows under the old name until the user re-keys them;
  values are unaffected.

Invariants the tests pin down: for a 60/40 demerger,
`parent Invested × T + child Invested = original Invested` **to the rupee**;
raw user rows byte-identical before/after (appended flagged rows aside);
second run idempotent.

Files: `model.py` (CorporateAction extension, restructures loader, resolve
chain), `update.py` (apply step ordered *before* price matching),
`generate.py` (Equity col T, CA-sheet columns, Invested formula),
`reader.py`, `data/restructures.csv`; SPEC §3.6, §5.8, §6.15, §6.5.

## Planned SPEC.md placement (written per-milestone, with the code)

| Addition | Section | Milestone |
|---|---|---|
| Dual-source bhavcopy contract (replaces §5.2 + §5.3) | §5.2–5.3 merged | R8 |
| §6.5 single-source escalation guard | §6.5 revision | R8 |
| Dividends sheet | §3.13 | R9 |
| Settings sheet | §3.14 | R10 |
| Gold_Silver sheet | §3.15 | R13 |
| NPS + NPS_Master sheets | §3.16 | R13 |
| EPF sheet | §3.17 | R12 |
| Manual_Assets sheet | §3.18 | R12 |
| NPS daily NAV contract | §5.6 | R13 |
| Bullion reference rate contract (IBJA + proxy fallback) | §5.7 | R13 |
| Restructures curated-file contract | §5.8 | R14 |
| `epf_rates.csv`, `bullion_proxies.csv`, `restructures.csv` | §5.5 additions | R12/R13/R14 |
| Dividend qty-at-ex-date estimate | §6.12 | R9 |
| Bullion rate derivation | §6.13 | R13 |
| Allocation drift & rebalance hint | §6.14 | R11 |
| Restructure engine | §6.15 | R14 |
| Class registry as normative configuration | §2 | R10 |
| History label-keyed columns | §6.11 revision | R12 |
| Updater flow: Settings read, hidden sheets, dividend/restructure steps | §7 | R9–R14 |

(§3.13+ / §5.6+ / §6.12+ are free today — SPEC currently ends at §3.12, §5.5,
§6.11.)

## Risks

- **IBJA rate source instability** — no committed API; mitigated by the
  three-layer design (proxy median fallback → carry-forward + amber → manual
  override wins). The product never *blocks* on it.
- **Exchange free-text drift** (dividend subjects, CA subjects) — tolerant
  regex + skipped-subject counts in the summary; Manual rows as backstop.
- **`restructures.csv` curation is an ongoing release duty** — scoped to
  retail-relevant events; Σ cost_pct validation fails loudly; Manual override
  per event.
- **NPS file location churn** — dual host, tolerant parse, URL pinned at
  implementation, day-cache + degrade.
- **R10 is the largest `generate.py` diff** (de-hardcoding `_PERSON_BLOCKS`
  and the Dashboard matrix) — golden-file the block layout; LibreOffice +
  Excel manual check for hidden sheets and shifted chart ranges (standing
  CLAUDE.md verification rule).
- **Dividend amounts are estimates** (no sell ledger) — amber formatting +
  explicit hint text; the estimate improves for free if a sell ledger ever
  lands.

## Verification

Per milestone, following the house pattern:

1. **Unit tests with hand-checked values** — dividend subject parser goldens;
   bullion median derivation; EPF accrual; drift/hint formula values; merger
   and 60/40-demerger arithmetic (Invested conservation to the rupee).
2. **Structural tests on the generated workbook** (openpyxl + zipfile) — new
   sheets/columns/dropdowns/conditional formats present; hidden-sheet state
   per Settings; chart ranges match the surviving layout.
3. **Injected-data integration tests** — `update.run(path, price_data=…,
   amfi_data=…, ca_data=…)` with fabricated fixtures (the existing DI seam):
   dual-exchange merge fixtures (R8), dividend announcements (R9), NPS/bullion
   fixtures (R13), restructure scenarios incl. idempotent re-runs (R14). One
   `tests/test_r<n>_*.py` per milestone, mirroring `test_r7_corporate_actions.py`.
4. **Round-trip identity** extended for every new data sheet and for hidden
   sheets in all enable/disable combinations — the regression backbone.
5. **Backward compatibility**: a v1.1 workbook (no Settings, 5-class History)
   updates cleanly; History totals unchanged; missing Settings → defaults.
6. Manual Excel (Jay's Windows machine) / LibreOffice check per CLAUDE.md.
7. **UX check per milestone** (the design principles above are acceptance
   criteria): every new sheet has a plain one-sentence hint in row 2; every
   amber cell carries a comment saying why; the Guide covers the feature in
   plain language; updater messages read as friendly sentences; the default
   template stays visually as clean as v1.1.
