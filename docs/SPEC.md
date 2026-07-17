# NetWorth — Product Specification

> **In one sentence:** this document describes *exactly* what the workbook and
> the updater must do — every tab, every number, every data source — so that
> anyone could rebuild NetWorth from scratch, in any language, without seeing
> the code.

**Version:** covers v1.0 – v1.4 · **Status:** normative.
Reverse‑engineered from the original template and extended with the approved
features in [PLAN.md](PLAN.md).

**This document is the product.** The Python in `src/networth/` is its
*reference implementation*; a port in any other language is conformant if it
satisfies this spec. Where the workbook and this spec disagree on a cosmetic
detail, **this spec wins**. When behaviour changes, the spec changes in the
same commit.

### How to read it

| You want to… | Go to |
|---|---|
| Understand the design rules | §1 Scope & principles |
| Know what every tab/column is | §3 Workbook specification |
| Fetch or parse a data source | §5 Data contracts |
| Reproduce a calculation | §6 Algorithms (pseudocode) |
| Know what the updater does, in order | §7 Updater behaviour |
| Ship it | §8 Packaging · §9 Portability checklist |

Section numbers are stable references — code comments cite them (e.g. "SPEC §6.10").

---

## 1. Scope & principles

1. **Local-only.** All computation happens on the user's machine. Network
   access is limited to HTTPS GETs of public market data (§5). No server, no
   account, no telemetry, no upload of any user data, ever.
2. **Excel workbook as UI and datastore.** The user's only interface is one
   `.xlsx` file. They enter holdings into *input cells* and may freely add,
   delete and sort data rows. Nothing in the flow may break because of normal
   data entry.
3. **The workbook is a build artifact.** A generator program produces the
   entire workbook from code; an updater refreshes it via the round-trip
   model (§7). Maintainers never hand-edit a shipped workbook's structure.
4. **Non-developer end users.** Refreshing prices must be a double-click
   action on both Windows and macOS, with no runtime dependencies beyond the
   shipped executable.
5. **Deterministic & testable.** Given the same inputs (user rows + fetched
   data + date), generation is reproducible.

### Terminology

- **Input cell/column** — user-entered; must round-trip unchanged through an
  update. Visually distinct (§3.2).
- **Computed column** — in-sheet Excel formula; recalculates live in Excel.
- **Updater-written value** — plain value computed by the updater at run time
  (prices, NAVs, XIRR); goes stale until the next run.
- **Owner / Person** — a family member name; the partitioning key of every
  holding row.
- **Master sheet** — machine-managed lookup list (MF_Master, Stock_Master,
  and in v1 Bank_Master); never edited by hand.

---

## 2. Configuration

The generator takes a configuration (defaults in parentheses):

| Key | Meaning |
|---|---|
| `persons` | ordered list of family member names (sample: Amit, Priya, Rahul); max 10 shown on the Dashboard matrix |
| `locale.date_format` | display format for dates (`dd-mm-yyyy`) |
| `inflation_default` | Dashboard inflation input default (7 %) |
| `expected_return_default` | v1: Dashboard expected-return input default for the FY-end estimate (10 %) |
| `row_budgets` | max data rows per sheet: Equity 137, MutualFunds 60, MF_SIP 500, FixedDeposits/PPF/Bonds ≥ 30, By Scrip ≥ 60 |
| `sample_data` | whether to include the fictional sample rows (on for the released template) |

Person names appear in three places that must stay consistent: the Dashboard
matrix input cells, one per-person sheet each, and the person columns of
By Scrip. The generator derives all three from `persons`.

### 2.1 Asset-class registry (v1.3, R10 — normative)

Every per-class surface is derived from ONE ordered registry; adding an asset
class means adding a registry row plus its sheet writer/reader/computes —
never editing the Dashboard/person/History logic:

| Field | Meaning |
|---|---|
| `key` | stable identifier; the per-class attribute on XIRR/History records |
| `label` | header text everywhere (Dashboard, person sheets, History, Settings) |
| `value_col` / `owner_col` | the SUMIFS ranges the Dashboard/person totals read |
| `sheets` | sheet group hidden together when the class is off |
| `person_rows` | data rows of the class's person-sheet block |
| `default_enabled` | new-workbook default (classic five = Yes; later classes = No) |
| `has_xirr` | blank the allocation-table XIRR cell when false (e.g. Cash) |

Registry order since v1.4.3 (12 classes — Settings rows 4–15 exactly):
Equity (sheets: Equity, By Scrip, Dividends; 40 person rows) → Mutual Funds
(MutualFunds, MF_SIP; 20) → Fixed Deposits (FixedDeposits; 15) → PPF (PPF,
PPF_Ledger; 10) → EPF (EPF; 10; default off) → Bonds (Bonds; 15) → Gold &
Silver (Gold_Silver; 10; default off) → NPS (NPS; 10; default off) →
Property / Cash / Insurance / Other (all on the shared Manual_Assets sheet
via `class_filter`; no person block; default off; Cash has
`has_xirr = false`). **"Property" was labelled "Real Estate" before
v1.4.3** — readers accept the old label wherever labels are matched
(Settings rows, Manual_Assets Class cells, the History header, the
allocation table) so old workbooks read seamlessly.

**Reference sheets (v1.4.3):** the four masters (MF_Master, Stock_Master,
Bank_Master, NPS_Master) plus the Corporate_Actions audit tab form a fixed
REFERENCE set whose visibility is driven solely by the Settings "Reference
lists" switch (§3.14), hidden by default. Every dropdown and INDEX/MATCH
formula resolves against hidden sheets, so nothing breaks — the tabs are
simply out of a first-time user's face.

**Enablement (normative — CHANGED in v1.4.3):** the user's Settings choice
wins: `shown = enabled`, rows or no rows. A class switched off is hidden
and **excluded from every displayed number** — Dashboard matrix and
allocation, person sheets, Projection / FY-expected (§6.8), the portfolio
XIRR (§6.2), and new History snapshots record 0 for it (§6.11). Data is
never deleted: its sheets are **hidden, never omitted** — openpyxl reads
hidden sheets, formulas keep resolving, and flipping Yes brings everything
(and its numbers) back. Awareness is mandatory whenever an off class holds
rows: the Dashboard carries a one-line amber notice (merged I1:P1 —
`Hidden, not counted: <labels> — switch on in Settings to include.`) and
the updater prints one matching summary line naming each such class with
its measured value. Surfaces driven by the enabled set: Dashboard matrix
columns (Total and Expected-@-FY columns shift left), allocation-table rows
and pie range, person summary rows and holding blocks (stacked from row 14
in registry order, each `person_rows` deep, one blank row between blocks),
and chart series. The History sheet's COLUMNS still include any class with
nonzero recorded history — data preservation, §6.11 — only its chart series
is dropped. (The pre-v1.4.3 rule was `enabled OR has_data`; v1.4.3 made the
user's choice authoritative.)

---

## 3. Workbook specification

### 3.1 Sheet map (tab order)

| # | Sheet | Kind | Purpose |
|---|---|---|---|
| 1 | Dashboard | mixed | family net worth, per-person × class matrix, XIRR, inflation, charts |
| 2 | Projection | computed | 20-year corpus trajectory table + line chart |
| v1.3 | Settings | input (§3.14) | per-class Yes/No + Target % + drift tolerance |
| 3… | one per person (e.g. Amit) | computed | that person's allocation + pie chart |
| … | Equity | data entry | stock holdings |
| … | MutualFunds | computed summary | one row per (owner, scheme), derived from MF_SIP |
| … | MF_SIP | data entry | one row per MF purchase/redemption |
| … | MF_Master | reference (hidden) | AMFI scheme list (~14k rows); Reference-lists switch §3.14 |
| … | Stock_Master | reference (hidden) | listed-stock list (~4.5k rows); Reference-lists switch |
| v1 | Bank_Master | reference (hidden) | bundled Indian bank list (Bank Name, Type; sorted; §3.11) |
| … | FixedDeposits | data entry | FDs |
| … | PPF | data entry | PPF accounts |
| v1.1 | PPF_Ledger | data entry | one row per PPF deposit (optional; §6.10) |
| v1.3 | EPF | data entry (§3.17) | EPF accounts — passbook balance + rate accrual (default off) |
| … | Bonds | data entry | corporate/other bonds |
| v1.3 | Gold_Silver | data entry (§3.15) | SGBs + physical gold/silver at the daily rate (default off) |
| v1.3 | NPS | data entry (§3.16) | NPS accounts — units × daily NAV (default off) |
| v1.3 | NPS_Master | reference (hidden) | NPS scheme list (§3.16); Reference-lists switch |
| v1.3 | Manual_Assets | data entry (§3.18) | hand-valued assets: Property / Cash / Insurance / Other (default off) |
| … | By Scrip | computed | family-wide exposure per stock |
| v1 | Corporate_Actions | reference (hidden, §6.7) | fetched + manual corporate actions and their effect; Reference-lists switch |
| v1.2 | Dividends | mixed (§3.13) | FY dividend ledger — auto + manual rows, by-month chart |
| v1.1 | History | updater data | one net-worth snapshot per day (§6.11) |
| … | Guide | static text | 2-minute manual |

Defined names (workbook scope):

```
MF_SchemeList  = MF_Master!$B$4:INDEX(MF_Master!$B:$B, COUNTA(MF_Master!$B:$B)+2)
Stock_NameList = Stock_Master!$B$4:INDEX(Stock_Master!$B:$B, COUNTA(Stock_Master!$B:$B)+2)
Bank_NameList  = (v1) same pattern over Bank_Master
```

### 3.2 Visual language

- **Title row 1** per sheet: bold sheet title, e.g. `EQUITY HOLDINGS`.
- **Hint row 2** (where present): one-line grey instruction text.
- **Header row 3**: bold on grey fill. Data starts at **row 4**.
  (Dashboard and person sheets have their own layouts, §3.3/§3.5.)
- **Input cells**: blue font; the “fill me” cells of the Dashboard are pale
  yellow. **Computed cells**: default font on light grey. This contrast is a
  spec requirement; exact shades are implementation-chosen.
- **Dates** display as `dd-mm-yyyy`; **money** as thousands-separated with
  0–2 decimals; **percentages** with 1–2 decimals.
- **Red/green (v1):** conditional formats on every Net chg., Day chg.,
  Return % and XIRR column/cell: value > 0 → green font (optionally pale
  green fill); value < 0 → red; blank → untouched. **Amber** fill marks
  degraded data: stale price (§6.5), delisted scrip, or FMV-fallback cost
  (§6.6). Colours must also work when a row is inserted/sorted (apply to the
  whole column range, not per-cell).
- **Cell comments** carry field help on headers (e.g. “Redemption = negative
  Amount”). Comments are part of the template; implementations must preserve
  the ability to regenerate them (openpyxl cannot — see CLAUDE.md).
- **Tab colours (v1.4.3):** the tab strip is colour-coded so it explains
  itself at a glance — navy `#1F4E79` for the overview tabs (Dashboard,
  Projection, Settings), teal `#31859C` for person tabs, blue `#4472C4` for
  every data-entry tab, grey `#A6A6A6` for the automatic tabs (By Scrip,
  Dividends, History) and the reference sheets, gold `#BF8F00` for the
  Guide.

### 3.3 Dashboard

Layout. Since v1.3 (R10) the class columns are the **effective-enabled set
in registry order** (§2.1); `<T>` below is the Total column (first after the
classes, `G` with the classic five) and `<X>` the Expected-@-FY column after
it. `<L>` is the last allocation-table row (`19 + #enabled`).

| Cell(s) | Content | Kind |
|---|---|---|
| A1 | `FAMILY PORTFOLIO — NET WORTH TRACKER` | static |
| I1:P1 (merged) | v1.4.3 hidden-money notice — present only when a switched-off class holds rows: `Hidden, not counted: <labels> — switch on in Settings to include.` (amber) | generator-written |
| I2:P2 (merged) | `New here? The Guide tab (last tab) walks you through everything.` | static |
| A2 / B2 | `As on` / `=TODAY()` | computed |
| A3 / B3 | `Family net worth` / `=<T>16` | computed |
| E3 | inflation % p.a. input (default 7) | **input** |
| B4 | Portfolio XIRR across all classes | **updater-written** |
| E4 | real return `=IF(B4="","",(1+B4)/(1+E3/100)-1)` | computed |
| F4 | verdict `=IF(B4="","",IF(B4>E3/100,"Beats inflation ✓","Below inflation ✗"))` | computed |
| row 5 | headers `Person, <enabled class labels…>, Total, Expected @ 31-Mar-<FY>` | static |
| A6:A15 | up to 10 person names | **input** (pre-filled from `persons`) |
| B6:…15 | per class: `=IF($A6="","",SUMIFS(<class value col>, <class owner col>, $A6))` | computed |
| `<T>`6:15 | `=IF($A6="","",SUM(B6:<last class col>6))` | computed |
| row 16 | `TOTAL` + column sums | computed |
| A17 / B17 | `Dividends FY <label>` cell (§3.13; present when Equity is enabled) | computed |
| A18 | `Allocation by asset class` | static |
| A19:G19 | headers `Asset class, Value, XIRR, Actual %, Target %, Drift, Rebalance hint` | static |
| A20:B`<L>` | one row per enabled class; Value `=<class col>16` (data-bar CF) | computed |
| C20:C`<L>` | per-class XIRR (blank when `has_xirr` is false) | **updater-written** |
| D20:D`<L>` | Actual % `=IF(<T>16=0,"",B20/<T>16)` | computed (v1.3, R11) |
| E20:E`<L>` | Target % `=IF(Settings!C<row>="","",Settings!C<row>/100)` — `<row>` is the class's **registry** Settings row, stable regardless of what is enabled | computed |
| F20:F`<L>` | Drift `=IF(E="","",D-E)` — green within ±tolerance (Settings B17), red outside, untouched when no target | computed |
| G20:G`<L>` | `=IF(E="","",IF(ABS(F)<=tol/100,"On target","Move ₹"&TEXT(ABS(F)*<T>16,"#,##0")&IF(F>0," out"," in")))` — indicative, pre-tax, class-level | computed |
| v1: D2/E2 | `Expected return % p.a.` label + input (default 10) for the FY-end estimate | **input** |
| `<X>`5, 6:15 | `Expected @ 31-Mar-<FY>` header + per-person values (§6.8) | **updater-written** |
| `<X>`16 | total `=IF(SUM(...)=0,"",SUM(...))` | computed |

All of D–G are live formulas — drift updates the moment a holding is edited,
no updater run needed (the glanceable property, §6.13).

Charts on Dashboard: **pie** "Allocation by asset class", **column** "Actual
vs Target %" (series D and E over the class labels; v1.3/R11), **bar** "Net
worth by person", **line** "Net worth over time" and **stacked area** "Net
worth by class over time" (both over History, §6.11).

### 3.4 Projection

Row 4 to row 24 (n = 0…20), columns:

| Col | Formula (row for year n) | Meaning |
|---|---|---|
| A | `=YEAR(TODAY())+n` | calendar year |
| B | `=Dashboard!$B$3*(1+Dashboard!$B$4)^n` | corpus at portfolio XIRR |
| C | `=Dashboard!$B$3*(1+Dashboard!$E$3/100)^n` | corpus growing at inflation (break-even line) |
| D | `=B/(1+Dashboard!$E$3/100)^n` | real (inflation-deflated) value of B |

Chart: **line**, “Corpus trajectory — portfolio return vs inflation
(20 years)”, series B and C (and optionally D) over A. Everything on this
sheet is live formulas — no updater involvement.

### 3.5 Person sheets (one per configured person)

| Cell(s) | Content |
|---|---|
| A1 | `<Name> — PORTFOLIO` |
| A2/B2 | `Owner` / the person's name (single source for this sheet's formulas) |
| A3/B3 | `Net worth` / `=B11` |
| A5:C5 | headers `Asset class, Value, # holdings` |
| A6:A10 | Equity, Mutual Funds, Fixed Deposits, PPF, Bonds |
| B6:B10 | `=SUMIFS(<class value col>, <class owner col>, $B$2)` (same column map as Dashboard) |
| C6:C10 | `=COUNTIF(<class owner col>, $B$2)` |
| A11:C11 | `Total` + sums |

Chart: **pie** “<Name> — allocation” over A6:B10.

### 3.6 Equity

Header row 3, data rows 4…140. Columns:

| Col | Header | Kind | Definition |
|---|---|---|---|
| A | Owner | input | person name |
| B | ISIN | computed | `=IF($C4="","",IFERROR(INDEX(Stock_Master!$C:$C,MATCH($C4,Stock_Master!$B:$B,0)),""))` — blank means “no master match”, user may overtype an ISIN manually (validation is non-blocking) |
| C | Scrip | input | type-ahead dropdown over `Stock_NameList` (§3.12) |
| D | Quantity | input | as-purchased (raw) quantity |
| E | Avg. cost | input | as-purchased average cost/share; **v1: may be left blank → FMV fallback §6.6** |
| F | Closing Price | updater | last close by ISIN |
| G | Prev. close | updater | previous close |
| H | Closing Price Date | updater | bhavcopy date used (a real date cell — the stale-price amber conditional format keys on `TODAY()-$H4>7`) |
| I | Cur. val | computed | `=IF($D4="","",$D4*$F4)` (v1: uses adjusted qty, §6.7) |
| J | Invested | computed | `=IF(OR($D4="",$E4=""),"",$D4*$E4*IF($T4="",1,$T4))` — × the v1.4 demerger Cost factor, blank = 1 |
| K | Net chg. | computed | `=I−J` guarded |
| L | Day chg. | computed | `=IF(OR($G4="",$D4=""),"",$D4*($F4-$G4))` |
| M | Cost date | input | drives per-row return annualisation & XIRR cashflows |
| N | XIRR (per row) | computed | `=IF(OR($M4="",N($J4)=0,$I4="",TODAY()<=$M4),"",($I4/$J4)^(365/(TODAY()-$M4))-1)` — simple two-flow annualisation |
| v1: O | Qty today | computed | `=IF($D4="","",$D4*IF($S4="",1,$S4))` — post-split/bonus share count, the **demat view**; feeds By Scrip and the person sheets |
| v1: P | Avg cost today | computed | `=IF(OR($D4="",$E4=""),"",$E4*IF($T4="",1,$T4)/IF($S4="",1,$S4))` — cost per share in today's share terms; × the v1.4 demerger Cost factor so a post-demerger row matches the docked basis a broker app shows |
| Q | Key | computed helper | `=IF($A4="","",$A4&"#"&COUNTIF($A$4:$A4,$A4))` stable per-owner sequence id |
| v1: R | Flags | updater helper | `FMV` (§6.6 fallback), `MERGED→<name>` / `ISIN→<isin>` (row priced via a successor, §6.15), `DEMERGER:<old_isin>@<ex_date>` (an appended child row) — flags round-trip regeneration. When a row carries both a restructure flag and `FMV` they are joined with `" | "`; neither may evict the other (the reader splits on the separator) |
| v1: S | Adj factor | **updater-written** | split/bonus **and merger-ratio** multiplier since Cost date (§6.7/§6.15, chain-aware); blank = 1. `Cur. val` and `Day chg.` use `Quantity*IF($S4="",1,$S4)*price` |
| v1.4: T | Cost factor | **updater-written** | demerger cost retention (§6.15): the parent keeps `cost_pct/100` of its cost basis, the rest moves to the appended child row; blank = 1. The user's Avg. cost cell is never rewritten |

Below the data block, one **updater-written** cell holds the equity-class
XIRR (legacy: N142). v1 additions: status/staleness amber flags (§6.5),
FMV-fallback marking on E (§6.6), adjustment columns (§6.7).

### 3.7 MutualFunds (summary) and MF_SIP (ledger)

**MF_SIP** — one row per purchase, SIP instalment or redemption
(redemption = negative Amount). Header row 3, data rows 4…503:

| Col | Header | Kind | Definition |
|---|---|---|---|
| A | Owner | input | |
| B | Fund House | computed | `INDEX(MF_Master!$A:$A, MATCH($C, MF_Master!$B:$B, 0))` guarded |
| C | Scheme Name | input | type-ahead dropdown over `MF_SchemeList` |
| D | ISIN | computed | `INDEX(MF_Master!$C:$C, MATCH($C, …))` guarded |
| E | Date | input | |
| F | Amount | input | negative = redemption |
| G | NAV on date | input | |
| H | Units | computed | `=IF(OR($F4="",$G4=""),"",$F4/$G4)` |

J1/J2: label `Portfolio MF XIRR` + **updater-written** value.

**MutualFunds** — one row per (owner, scheme); the user enters only A and C:

| Col | Header | Kind | Definition |
|---|---|---|---|
| A | Owner | input | |
| B / D | Fund House / ISIN | computed | master lookups as above |
| C | Scheme Name | input | dropdown |
| E | Units | computed | `=SUMIFS(MF_SIP!$H:$H, MF_SIP!$A:$A,$A4, MF_SIP!$D:$D,$D4)` |
| F | Avg cost NAV | computed | `=H/E` guarded |
| G | Current NAV | updater | AMFI by ISIN |
| H | Invested | computed | `=SUMIFS(MF_SIP!$F:$F, …)` |
| I | Cur. val | computed | `=E*G` guarded |
| J | Net chg. | computed | `=I−H` |
| K | Return % | computed | `=J/H` guarded |
| L | XIRR | **updater-written** | true XIRR from that (owner, ISIN)'s MF_SIP cashflows + current value |
| N | Key | helper | as in Equity |

### 3.8 FixedDeposits

| Col | Header | Kind | Definition |
|---|---|---|---|
| A | Owner | input | |
| B | Bank / Institution | input | **v1:** type-ahead dropdown over `Bank_NameList`, free text still allowed |
| C | FD No. | input | |
| D | Principal | input | |
| E | Rate % p.a. | input | |
| F / G | Start / Maturity Date | input | |
| H | Comp./yr | input | compounding periods per year (4 = quarterly) |
| I | Value as on today | computed | `=D*(1+(E/100)/H)^(H*YEARFRAC(F, MIN(TODAY(),G)))` |
| J | Maturity Value | computed | same with `YEARFRAC(F,G)` |
| L | Key | helper | |

### 3.9 PPF

| Col | Header | Kind |
|---|---|---|
| A | Owner | input |
| B | Institution | input |
| C | Account No. | input |
| D | Current Balance | input |
| E | Balance as-on | input (date) |
| F | Rate % (ref) | input (default 7.1; v1: generator pre-fills current rate from `data/ppf_rates.csv`) |
| G | Notes | input |
| I | Key | helper |

No ledger in v1 — value grows at Rate% from the as-on date for XIRR and
FY-end purposes (contribution ledger is roadmap).

### 3.10 Bonds

| Col | Header | Kind | Definition |
|---|---|---|---|
| A | Owner | input | |
| B | Issuer / Bond | input | |
| C | ISIN | input | |
| D | Qty | input | |
| E | Face Value | input | |
| F | Buy Price | input | per unit |
| G | Current Price | input (updater fills when the ISIN trades on the exchange) | |
| H | Coupon % p.a. | input | |
| I | Maturity Date | input | |
| J | Invested | computed | `=D*F` guarded |
| K | Cur. val | computed | `=D*G` guarded |
| L | Net chg. | computed | `=K−J` |
| M | Buy Date | input | required for XIRR; rows without it are skipped |
| N | Key | helper | |
| v1: O | Maturity Value | computed | `=IF(OR($D4="",$E4=""),"",$D4*$E4)` — redemption at face. Cumulative/zero bonds: set H = 0 and Face Value = redemption amount |
| v1: P | Coupons till maturity | computed | `=IF(OR($D4="",$E4="",$H4="",$I4="",$I4<=TODAY()),"",$D4*$E4*($H4/100)*YEARFRAC(TODAY(),$I4))` — simple, non-reinvested |

v1: bond XIRR (per row and class) includes coupon cashflows: `+D*E*(H/100)/f`
on each coupon date from Buy Date to today (f = coupon frequency, default
annual) — see §6.3.

### 3.11 By Scrip, masters, Guide

**By Scrip** — data rows from 4; A ISIN (input or updater-synced from Equity),
B Scrip lookup, C `=SUMIF(Equity!$B:$B,$A4,Equity!$D:$D)` total qty, one
column per configured person `=SUMIFS(Equity!$D:$D, Equity!$B:$B,$A4,
Equity!$A:$A,"<Person>")`, last column Cur. val
`=SUMIF(Equity!$B:$B,$A4,Equity!$I:$I)`.

**MF_Master** — A1 title, A2 hint, D2 `Refreshed:` + E2 date (updater),
row 3 headers `Fund Name, Scheme Name, ISIN`, data from row 4,
**sorted by Scheme Name** (ordinal, case-insensitive). Source: AMFI (§5.1).

**Stock_Master** — headers `Symbol, Stock Name, ISIN`, same layout, sorted by
Stock Name. Merge policy is **add-only** (§6.4). v1 adds `Status` and
`Last Traded` columns (§6.5).

**Bank_Master (v1)** — headers `Bank Name, Type`, sorted by name, seeded from
`data/banks_in.csv` (RBI scheduled banks + major SFBs/co-ops). Static; only
release updates refresh it.

**Guide** — plain-text manual covering: inputs vs computed colours, how to add
people/holdings, dropdown usage, what the updater does, backups, and the v1
flags (amber = stale/delisted/FMV-estimated).

### 3.12 Type-ahead dropdowns (normative mechanism)

List validation with:

```
=OFFSET(<Master>!$B$3,
        IFERROR(MATCH($C4&"*", <NameList>, 0), 1), 0,
        MAX(1, COUNTIF(<NameList>, $C4&"*")), 1)
```

- Begins-with filtering: typing a prefix then opening the dropdown shows only
  matching entries. Requires the master **sorted** by that column.
- **Interaction is two-step by Excel's design**: the window formula evaluates
  against the cell's *committed* value, so the user must type the prefix,
  press Enter, then re-open the dropdown (arrow click or Alt+Down on the
  selected cell). Excel offers no live suggestions while typing in a
  validation cell (only recent Microsoft 365 builds add native autocomplete).
  The input tip and Guide must state this two-step flow explicitly.
- `showErrorMessage = false` (non-blocking): users may keep free text (e.g. a
  delisted scheme); the lookup columns then stay blank, which downstream
  formulas treat as "fill ISIN manually".
- Input tip on the cell explains the behaviour.
- Applied ranges (legacy): Equity C4:C140, MutualFunds C4:C63, MF_SIP C4:C503;
  v1 adds FixedDeposits B4:B<n> over `Bank_NameList`.

### 3.13 Dividends (v1.2, R9)

FY dividend ledger: one row per dividend event × owner. Title r1, hint r2
(plain-language: rows fill in automatically; amounts are estimates, amber),
header r3, data r4..203.

| Col | Header | Kind | Definition |
|---|---|---|---|
| A | FY | updater / input | Indian financial year of the ex-date, e.g. `2026-27`; the updater backfills a blank FY on Manual rows from F |
| B | Owner | updater / input | one row per owner holding the stock at ex-date |
| C | Scrip | updater / input | Stock_Master display name |
| D | ISIN | updater / input | |
| E | Type | dropdown | Interim / Final / Special (non-blocking validation) |
| F | Ex-Date | date | |
| G | Rate ₹/share | updater / input | parsed from the announcement (§5.4) |
| H | Qty @ ex-date (est.) | updater | §6.12 estimate; **amber**; user-correctable on Manual rows |
| I | Est. amount | computed | `=IF(OR(G="",H=""),"",G*H)` — **amber** (estimate; the exact credit is on the bank statement) |
| J | Source | updater | Auto / Manual |
| K | Details | updater / input | announcement free text |

**Lifecycle (normative).** On every update run: Auto rows whose ex-date falls
in the **current FY** are rebuilt from the feed + current holdings; all other
rows persist unchanged (prior-FY Auto rows therefore freeze on the first run
after Apr 1 — multi-year record for free). Manual rows always persist and
suppress an Auto row with the same `(isin, ex_date)` key — the Type is
deliberately NOT part of the key, because the exchanges word the same event
differently (§5.4 dedupe rule). Current-FY Auto rows of an ISIN the feed
could NOT verify this run are **kept, not rebuilt** — a one-symbol outage
must never delete income already on the sheet. Re-runs are idempotent. If
the feed is unreachable entirely, the sheet is left as-is. **Capacity:** the
sheet holds 200 data rows; if an assembly exceeds it, the OLDEST prior-FY
Auto rows give way (Manual and current-FY rows never do) and the run warns
with the dropped count.

**By-month chart.** Columns M/N rows 4..15 hold the current FY's months
(Apr..Mar) and `SUMPRODUCT(rate × qty × month × FY)` sums; a column chart
("Dividends by month — FY <label>") renders them. The current-FY label is
stamped at build time (the updater regenerates the workbook, keeping it
fresh). The Dashboard shows one cell: `Dividends FY <label>` =
`SUMIFS(Dividends!I:I, Dividends!A:A, "<label>")`.

Dividends do **not** feed equity XIRR (roadmap; changing return semantics
deserves its own release).

### 3.14 Settings (v1.3, R10; simplified in v1.4.3)

The one place the user tunes the workbook. Title r1, hint r2 ("Show? — Yes
shows a tab, No hides it. Nothing is ever deleted…"), header r3
(`Asset class, Show?, Target %, Status, Notes`), one row per registry class
from r4 (rows 4–15 reserved), then:

| Cell | Content | Kind |
|---|---|---|
| B4:B15 | `Yes` / `No` dropdown (non-blocking) — show or hide the class | **input** |
| C4:C15 | target allocation %, blank = no target (R11 drift view; header comment says "optional") | **input** |
| D4:D15 | `Shown` / `Hidden` / `Hidden - has data (not counted)` | generator-written |
| E4:E15 | note (for the has-data case: rows are saved but not counted; switch to Yes to include) | generator-written |
| A16/B16 | `Reference lists` — Yes/No for the REFERENCE sheets (§2.1); default **No** | **input** |
| A17 | `Balance targets (optional)` section label | static |
| A18/B18 | `Drift tolerance (± % points)`, default **5** | **input** |
| A19/B19 | `Targets total` = `SUM(C4:C15)`, **amber** when non-zero and ≠ 100 | computed |

Reader rules: match class rows by label anywhere in rows 4–20 (tolerant;
"Real Estate" accepted for Property); a missing `Reference lists` row
(pre-v1.4.3 workbook) ⇒ No; missing sheet (pre-v1.3) ⇒ registry defaults;
the user's No always round-trips unchanged. Real form-control checkboxes
are deliberately not used (xlsxwriter cannot write them; LibreOffice
renders them poorly) — the Yes/No validation dropdown is the normative
control.

### 3.15 Gold_Silver (v1.3, R13; default off)

SGBs price from the merged bhavcopy by ISIN (they trade on the cash market
at ~₹/gram); physical metal values grams × purity × the daily reference rate
(§5.7). Title r1, hint r2 (+ H2/I2 `Rates as on` stamp, amber trigger),
header r3, data r4..53, TOTAL r55.

| Col | Header | Kind | Definition |
|---|---|---|---|
| A | Owner | input | |
| B | Type | input | dropdown: SGB / Gold / Silver |
| C | Description / Series | input | "SGB 2023-24 Ser II", "Gold coins, 2 x 10 g (24K)", "Silver bar, 1 kg" (header comment carries these examples) |
| D | ISIN | input | SGB only — drives bhavcopy pricing |
| E | Qty (g / units) | input | SGB: units (1 unit = 1 g); metal: grams |
| F | Purity | input | blank = 1 (SGB always 1); 22K = 0.916, 18K = 0.75 |
| G | Buy Price ₹/unit | input | per gram/unit; XIRR outflow |
| H | Buy Date | input | XIRR anchor |
| I | Rate today (auto) | **updater** | SGB: exchange close; metal: §5.7 ₹/g rate. **Amber** on METAL rows when the I2 rates-as-on stamp is > 7 days old (SGB rows are exempt — their closes carry their own dates and the share-price staleness rules apply). I2 itself is stamped only when a metal rate actually arrived — an SGB-only pricing day must not refresh it and hide a stale benchmark |
| J | Rate override | input | user's ₹/unit (e.g. the jeweller's board rate) — **always wins over I** |
| K | Cur. val | computed | `=E × (F or 1) × (J else I)`, guarded; the class value column |
| L | Invested | computed | `=E × G`, guarded |
| M | Net chg. | computed | `K − L`, red/green |
| N | Maturity | input | SGB (8 years); blank for metal |
| O | Key | helper | |

SGB XIRR includes the statutory 2.5 % p.a. semi-annual coupon computed on
the row's **Buy Price** — a documented approximation (the statutory base is
issue price; an extra column is not worth the width).

### 3.16 NPS + NPS_Master (v1.3, R13; default off)

Units × daily NAV, exactly the mutual-fund mental model. NPS sheet: title
r1, hint r2, header r3, data r4..43, TOTAL r45.

| Col | Header | Kind | Definition |
|---|---|---|---|
| A | Owner | input | |
| B | PRAN | input | free text |
| C | Scheme | input | type-ahead dropdown over `NPS_SchemeList` (§3.12) |
| D | Scheme Code | computed | INDEX/MATCH on NPS_Master; manual override allowed (plain text beats the formula) |
| E | Units | input | from the CRA statement |
| F | Current NAV | **updater** | daily NAV by scheme code (§5.6) |
| G | Cur. val | computed | `=E × F`, guarded; the class value column |
| H | Total contributed | input (optional) | enables the approximate XIRR |
| I | First contribution | input (optional, date) | |
| J | XIRR | **updater** | approximate two-flow (−H @ I, +G today); header comment states the approximation — a dated-contribution ledger is roadmap |
| K | Key | helper | |

**NPS_Master**: `Scheme Code, Scheme Name, PFM` + refreshed stamp (E2),
sorted by scheme name (the dropdown sort rule, §3.12), **add-only merge
keyed by scheme code** (§6.4 pattern). The reader keeps a row when its
**Scheme Code** is non-empty — PFM is descriptive, and a blank PFM must
never drop a scheme from the master (the MF/Stock masters key on their
ISIN column instead). A REFERENCE sheet since v1.4.3 — visibility follows
the Settings "Reference lists" switch (§2.1), not the NPS class.

### 3.17 EPF (v1.3, R12; default off)

Deliberately congruent with PPF's flat path: passbook balance in, accrual
out. Title r1, hint r2, header r3, data r4..43, TOTAL r45.

| Col | Header | Kind | Definition |
|---|---|---|---|
| A | Owner | input | |
| B | Establishment / UAN | input | |
| C | Member ID | input | |
| D | Current Balance | input | EPFO passbook closing balance (employee + employer) |
| E | Balance as-on | input (date) | the passbook date |
| F | Rate % | input | blank ⇒ updater fills the latest `epf_rates.csv` rate |
| G | Notes | input | |
| H | Balance today | computed | `=IF(D="","",IF(OR(E="",F=""),D,D*(1+F/100)^YEARFRAC(E,TODAY())))` — flat accrual; the class value column |
| J | Key | helper | person-block lookup |

Exact monthly-run accrual + a contribution ledger is a roadmap follow-up
(mirroring PPF's flat-first history); the H header comment states the
estimate nature.

### 3.18 Manual_Assets (v1.3, R12; shared sheet, four registry classes, default off)

One sheet for every hand-valued asset; the `Class` column routes each row to
its own registry class (Property / Cash / Insurance / Other — each with
its own Dashboard column, allocation row, target and History column via the
`class_filter` SUMIFS criterion, §2.1). The sheet hides only when ALL four
are off. Title r1, hint r2 ("Things you value yourself… only two numbers
matter"), header r3, data r4..63, TOTAL r65.

| Col | Header | Kind | Definition |
|---|---|---|---|
| A | Owner | input | |
| B | Class | input | dropdown (non-blocking): Property / Cash / Insurance / Other. Matching is **case-insensitive** — Excel's SUMIFS already is, and the reader canonicalises a typed variant ("property", and the pre-v1.4.3 "Real Estate") to the dropdown label so both sides agree. A value matching NO label counts in no class (only the sheet TOTAL sees it); the updater warns naming the row |
| C | Description | input | "Apartment (self-occupied)", "Savings account balance", "Life policy - surrender value today" (the header comment carries these examples) |
| D | Institution / Ref | input | bank, insurer, registrar |
| E | Invested / Cost | input (optional) | RE: purchase cost; Insurance: premiums paid; enables Net chg. + XIRR |
| F | Cost date | input (optional) | XIRR anchor |
| G | Current value ₹ | input | THE number: market estimate / balance / surrender value; the class value column |
| H | Value as-on | input (date) | **amber when > 90 days old** — hand-typed values rot silently |
| I | Net chg. | computed | `=IF(OR(E="",G=""),"",G−E)`, red/green |
| J | Notes | input | |
| K | Key | helper | |

No per-person holding block (the sheet itself is the overview); person
sheets show one summary row per subclass. XIRR per §6.2 (Cash excluded).

---

## 4. Sample data

The released template ships with fictional holdings for three people (Amit,
Priya, Rahul) using **real ISINs** so the first updater run works end-to-end.
MF samples must be real AMFI (Scheme Name, Fund House, ISIN) triples; equity
samples real BSE scrips. **EVERY asset class carries sample rows** (incl. a
real SGB ISIN, generic gold coins / 22K jewellery / a silver bar, an NPS
scheme from the seeded master, an EPF passbook line, and deliberately
generic Property ("Apartment (self-occupied)") / Cash / Insurance / Other
rows). Targets sit on the five default-on classes (40/15/20/15/10 — sum
100) so the drift view demonstrates itself. **v1.4.3 calm first open:** the
classic five ship Settings **Yes** and are all a new user sees; every newer
class ships **No and therefore hidden**, its sample rows waiting inside as
a worked example the moment it is switched on. The stored ClassXirr carries
figures only for the shown classes (the allocation table lists only those);
hidden classes get theirs computed when enabled. Onboarding = replace the
sample rows with your own and switch on what you own — nothing needs
deleting to keep the workbook tidy.

---

## 5. Data contracts

All fetches: plain HTTPS GET, ≤ 2 retries, per-day local cache (`cache/`),
graceful degradation (a failed source leaves old values in place and reports
it — never blanks user-visible data).

### 5.1 AMFI daily NAVs + scheme master

`https://www.amfiindia.com/spages/NAVAll.txt` — `;`-separated text:

```
Scheme Code;ISIN Div Payout/ ISIN Growth;ISIN Div Reinvestment;Scheme Name;Net Asset Value;Date
```

Parsing rules: lines without `;` are section headers — the most recent one is
the current **Fund House**; skip blanks/headers; a scheme yields up to two
(ISIN → NAV) entries (both ISIN columns); NAV `N.A.` → skip. Date format
`dd-MMM-yyyy`.

### 5.2 BSE bhavcopy (dual source with NSE, since v1.2 / R8)

`https://www.bseindia.com/download/BhavCopy/Equity/BhavCopy_BSE_CM_0_0_0_<yyyymmdd>_F_0000.CSV`

Common-format CSV. Columns are located **by header name**, tolerantly:
ISIN ∈ {ISIN, ISIN_CODE, …}; Close ∈ {ClsPric, Close, LAST, …} (never a
column containing “prev”); Prev ∈ {PrvsClsgPric, PrevClose, …}; also read
TckrSymb/FinInstrmNm for the master, FinInstrmId (BSE scrip code) for the
corporate-actions lookup, and **HghPric** when building FMV data.

**Merge rule (normative).** Both exchanges are fetched for the **same trade
date** and merged — dates are never mixed across exchanges. Not published on
holidays: try today, then walk back up to 7 calendar days, stopping on the
first day where **at least one** exchange answers; record the date actually
used (→ Closing Price Date column). The merged result is:

```
prices        = union of ISINs; on a dual-listed conflict NSE close/prev win
                (deeper cash-market liquidity — matches broker apps)
codes_by_isin = from the BSE parse ONLY, retained whenever BSE responded
                (NSE's FinInstrmId is not a BSE scrip code)
master rows   = deduped by ISIN; NSE symbol preferred for NEW ISINs (it is
                what the NSE corporate-actions API needs); the add-only
                merge (§6.4) protects existing rows regardless
source label  = per RUN, not per cell: "BSE+NSE <date>" (or the single
                exchange that answered), plus an NSE-only count in the
                console summary
```

If only one exchange published for the chosen day, the run proceeds
single-source: quoted rows update normally, but the §6.5 status escalation
is skipped (absence from one exchange is not evidence of anything). Both
failing for all 7 days is the only hard failure (updater then degrades
gracefully, keeping old prices).

### 5.3 NSE bhavcopy (peer source)

`https://nsearchives.nseindia.com/content/cm/BhavCopy_NSE_CM_0_0_0_<yyyymmdd>_F_0000.csv.zip`
— zip containing one CSV, same common format, parsed and merged per §5.2.
NSE requires a browser-like `User-Agent` and a cookie warm-up GET on
`https://www.nseindia.com/` first. A 200 response whose body is NOT a valid
zip (NSE serves bot-challenge HTML pages with status 200) is treated exactly
like NSE-unavailable: the day proceeds single-source on BSE — it must never
abort the fetch or discard already-parsed BSE data.

### 5.4 Corporate actions (v1, R7; dual-source since v1.0.0-rc)

Per held stock, fetch historical + announced actions from **both exchanges**
and deduplicate:

```
NSE: https://www.nseindia.com/api/corporates-corporateActions?index=equities&symbol=<SYM>
     (cookie warm-up required, like §5.3; free-text field: subject)
BSE: https://api.bseindia.com/BseIndiaAPI/api/DefaultData/w?Fdate=&Purpose=&TDate=
     &ddlcategorys=E&ddlindustrys=&scripcode=<CODE>&segment=0&strSearch=S
     (Referer: https://www.bseindia.com/ required; free-text field: Purpose;
      Ex_date format "28 Oct 2024"; scrip codes come from the daily BSE
      bhavcopy's FinInstrmId column — no extra mapping source)
```

The free text is classified identically for both: `Bonus A:B` / "Bonus issue
A:B" → BONUS; "split"/"sub-division"/"Stock Split" with "From <face> To
<face>" → SPLIT; "consolidation" → CONSOLIDATION. **Dividends (v1.2, R9):**
a subject containing "dividend" with a rupee amount ("Rs 8 Per Share",
"Rs. - 5.5000", "Re. 1/-", "₹2.50") yields a dividend record — type from
`interim|final|special` (default Final), rate in ₹/share. Two guards keep
garbage out of the rate: the currency token must not sit inside a word (the
"re" ending "…Per Sha**re** 2024" is not `Re.`), and any "face value of
Rs.N" phrase is masked BEFORE the rate search, so "Dividend - 300% on face
value of Rs.2/- each" parses no rate at all. Percent-of-face wordings
("Dividend 250%") are skipped and **counted only when the ex-date falls in
the current FY** — the feeds carry decades of history, and warning about a
2004 record is noise — and the updater reports the count so the user can
add a Manual row. Dividend records dedupe on `(isin, ex_date, rate)`, NSE
wins — deliberately NOT on the type: the exchanges word the same event
differently ("Dividend" → Final on NSE vs "Interim Dividend" on BSE) and a
type-keyed dedupe would double-count it; two genuinely distinct same-day
payouts differ in rate and both survive. Everything else (rights, AGMs,
buybacks) is ignored. The normative contract is the record, not the URLs:

```
{ symbol, isin, ex_date, type ∈ {SPLIT, BONUS, CONSOLIDATION},
  ratio_from, ratio_to, source ∈ {Auto, Manual}, details }
```

SPLIT/CONSOLIDATION ratio = old face : new face (e.g. 10:2 split → factor 5).
BONUS ratio A:B = A new shares per B held (factor 1 + A/B).

**Dedupe rule:** records from the two exchanges merge on
`(isin, type, ex_date)` — ex-dates are exchange-synchronised — NSE record
wins. Manual rows on the Corporate_Actions sheet take precedence over an
Auto row with the same key.

**Coverage rule (never skip silently, never revert):** the fetch reports
which ISINs were successfully answered by at least one exchange. Any held
ISIN answered by neither MUST surface as a user-visible warning naming the
scrip ("corporate actions could NOT be verified for: …"), so an unverifiable
holding is a known condition, not a silent gap. One security failing must
not abort the sweep, and — critically — must not REVERT that security: the
Auto action rows (and current-FY Auto dividend rows, §3.13) of an unverified
ISIN survive the rebuild untouched, so an already-applied split's quantities
never snap back because one endpoint blocked one symbol on one day. Only an
all-security/all-source failure degrades wholesale (keep every existing
row). Since v1.4 (§6.15), merged holdings' SUCCESSOR symbols are queried
too — the successor's dividends and later actions concern the old-ISIN
rows. Mergers/demergers/ISIN reassignments remain out of scope for
auto-adjustment from these feeds (no reliable free feed publishes ratios) —
the Curated + Manual paths cover them.

### 5.5 Bundled static data (in `data/`, refreshed only by releases)

| File | Shape | Source |
|---|---|---|
| `fmv_2018-01-31.csv` | `isin, symbol, fmv` | NSE bhavcopy of 2018-01-31 (EQ/BE/BZ series, 1,639 ISINs); FMV = that day's **high** price (IT Act grandfathering definition). The symbol column enables lookup when a later corporate action reissued the ISIN (e.g. HDFC Bank post-split) |
| `banks_in.csv` | `bank_name, type` | RBI scheduled commercial banks list + major SFB/payment/co-op banks |
| `ppf_rates.csv` *(roadmap — ships with the PPF contribution ledger)* | `from_date, to_date, rate_pct` | MoF quarterly notifications, historical to present (no official API exists) |
| `epf_rates.csv` *(v1.3, R12)* | `fy_start, rate_pct` | EPFO annual declared rates, historical to present (no official API — refreshed via releases); the updater fills a blank EPF Rate % with the latest row |
| `bullion_proxies.csv` *(v1.3, R13)* | `metal, match ∈ {symbol_prefix, isin}, key, grams_per_unit, note` | exchange-traded ₹/gram proxies for the §5.7 fallback (SGB prefix for gold, SilverBeES for silver; GoldBeES was dropped 2026-07-17 — its live grams-per-unit had drifted ~17 % from the nominal 0.01 g, exactly the expense-ratio decay this column warns about); release-refreshed |

### 5.6 NPS daily NAVs + scheme master (v1.3, R13)

```
PRIMARY : https://npstrust.org.in/nav-report-excel
          (despite the name: TAB-separated text, verified 2026-07-16 —
           ID, DATE OF NAV, PFM NAME, SCHEME ID, SCHEME NAME, NAV VALUE;
           one row per scheme, latest published day)
FALLBACK: https://npscra.nsdl.co.in/download/NAVReport.csv (same record)
```

Columns located by header name, tolerantly; delimiter sniffed (tab vs
comma); keyed by **SCHEME ID** (e.g. `SM001003`); rows without a positive
NAV are skipped. Feeds the NPS_Master add-only merge and per-row NAV refresh
(code resolved from the row's Scheme Code column — lookup or override). No
API key; failure keeps old NAVs with a summary warning.

### 5.7 Bullion reference rate (v1.3, R13) — layered by design

The flakiest data in the product, so it must never block a run:

```
1. PRIMARY : IBJA daily benchmark — https://www.ibjarates.com/
             stable span ids: lblGold999_PM (₹/10 g), lblSilver999_PM
             (₹/kg); _AM variants earlier in the day. Normalise to ₹/gram.
             The rate the bullion trade quotes from; RBI uses IBJA 999 for
             SGB redemption. No committed API — parse defensively, return
             nothing on any doubt.
2. FALLBACK: market-implied ₹/g = median(close / grams_per_unit) over the
             quoted proxies of data/bullion_proxies.csv (SGB tranches ≈
             ₹/g fine gold; SilverBeES ≈ 1 g) from the bhavcopy already
             fetched — zero extra HTTP. Typically 2–4 % below the IBJA
             retail benchmark (Guide says so). Only proxies whose implied
             rate verifiably tracks the metal stay in the file (§5.5 —
             GoldBeES was dropped for unit drift).
3. DEGRADE : both fail ⇒ keep each row's previous Rate today AND the old
             rates-as-on stamp (amber past 7 days) + a summary warning.
             The stamp advances only when a metal rate actually arrived —
             SGB pricing alone never refreshes it (§3.15).
4. The sheet's Rate-override column always wins over the auto rate.
```

### 5.8 Curated restructures — `data/restructures.csv` (v1.4, R14)

No reliable free feed publishes merger/demerger swap ratios, so the product
ships a **curated, release-refreshed** file (the ppf_rates/fmv precedent —
keeping it current is an ongoing release duty; anything missed is covered by
a Manual row on Corporate_Actions, which overrides a Curated row with the
same `(old_isin, type, ex_date)` key):

```
ex_date, type ∈ {MERGER, DEMERGER, ISIN_CHANGE}, old_isin, old_name,
old_symbol, new_isin, new_name, new_symbol, ratio_from, ratio_to,
cost_pct, details
```

- **MERGER** (old security absorbed): `ratio_from:ratio_to` = A new shares
  per B old; `cost_pct = 100` — cost basis and holding period carry in full
  (Sec. 47 tax-neutral).
- **DEMERGER**: one row per resulting security, grouped by
  `(old_isin, ex_date)` — a parent-retention row (`new_isin = old_isin`,
  1:1) plus one row per spun-off child (child shares per parent share, the
  company-notified income-tax cost apportionment). **The loader validates
  Σ cost_pct = 100 per event and fails loudly** — a silently wrong split
  would corrupt capital-gains numbers.
- **ISIN_CHANGE**: 1:1, `cost_pct = 100`.

Scope: index-grade events likely to touch retail portfolios (shipped v1.4.0:
the HDFC Ltd → HDFC Bank merger). Rows load as `source = Curated`, rewritten
from the file each run **except the Applied date, which persists** (§6.15).

---

## 6. Algorithms (normative pseudocode)

### 6.1 XIRR solver

Inputs: cashflows `[(date, amount)]`, sign convention outflow < 0, inflow > 0.

```
guard: < 2 flows, all same date, all same sign, |sum of days| == 0  → null
f(r) = Σ amount_i / (1 + r)^(days_i / 365)      days_i from first flow date
solve f(r) = 0 by Newton from r=0.1, fall back to bisection on [-0.9999, 10]
tolerance 1e-7, max 100 iterations; no root → null
```

null results render as blank cells, never 0 or an error.

### 6.2 Cashflow assembly

Per asset class (skip rows with missing required inputs):

| Class | Outflows | Inflows |
|---|---|---|
| Equity | −Invested @ Cost date (× the §6.15 cost factor when set — **0 is a real value**, meaning the parent retained no cost; only blank means 1) | +Cur. val @ today |
| Mutual Funds | one per MF_SIP row: −Amount @ Date (redemptions are +) | +Cur. val @ today per (owner, ISIN) |
| Fixed Deposits | −Principal @ Start | +Value-as-on @ min(today, maturity) |
| PPF | −(Balance discounted at Rate% back to as-on date)… in practice: −Balance @ as-on | +Balance·(1+Rate%)^(days/365) @ today |
| Bonds | −Qty·BuyPrice @ Buy Date (skip if no Buy Date) | +Qty·CurrentPrice @ today; **v1:** plus each coupon `+Qty·Face·(Coupon%/f)` on its historical coupon date |
| EPF *(v1.3)* | −Balance @ as-on (PPF flat path verbatim) | +Balance·(1+Rate%)^(days/365) @ today |
| Gold & Silver *(v1.3)* | −Qty·BuyPrice @ Buy Date | +Cur. val @ today; SGB rows add each historical semi-annual coupon `+Qty·BuyPrice·1.25%` (§3.15 approximation note) |
| NPS *(v1.3)* | −Total contributed @ First contribution (row skipped when either optional input is blank) | +Units·NAV @ today |
| Property / Insurance / Other *(v1.3)* | −Invested @ Cost date (row skipped when Invested, Cost date or Value is blank) | +Current value @ today |
| Cash *(v1.3)* | **excluded from XIRR entirely** (`has_xirr` false, §2.1) — a balance has no meaningful money-weighted return | |

Class XIRR = solver over that class's union. Portfolio XIRR = solver over
the union of the **enabled** classes only (v1.4.3, §2.1) — a switched-off
class contributes nothing to the family figure; its per-class value may
still be computed (harmless, and ready when re-enabled) but is written
nowhere while hidden. Written as plain values to: Dashboard B4, the
allocation table's XIRR column, Equity class cell, MutualFunds L column +
MF_SIP J2.

### 6.3 Coupon schedule (v1)

From Maturity Date step backwards by `12/f` months (f default 1 = annual) to
Buy Date; coupons with date ≤ today enter the XIRR cashflows; future coupons
+ redemption feed only the Maturity Value / FY-end figures.

### 6.4 Master merge (add-only)

```
new = fetched list;  existing = current master rows (key: ISIN)
for isin in new:  if isin not in existing → append (symbol, name, isin)
never rename or delete an existing row (user rows key on the NAME)
resort whole table by name (ordinal, case-insensitive)   # dropdown requirement
write refresh date to E2
```

MF_Master is regenerated wholesale from AMFI each refresh (same sort rule)
but must also preserve any ISIN currently referenced by a user row even if
AMFI drops it (append with its last-known names).

### 6.5 Delisted / stale detection (v1)

```
for each held ISIN, at update time:
  quoted in the merged bhavcopy   → Status=Active, LastTraded=bhavcopy date
  absent — SINGLE-SOURCE run      → carry the previous status forward
                                    untouched (absence from one exchange is
                                    not evidence; prevents a false Suspended
                                    during a one-exchange outage)
  absent from BOTH exchanges:
    ≤ 21 calendar days            → keep last price/status; the live amber
                                    "stale" conditional format fires anyway
                                    once Closing Price Date is > 7 days old
    > 21 calendar days            → Status=Suspended (amber via status CF)
    > 180 calendar days           → Status=Delisted (amber via status CF)
Suspended/Delisted rows keep their last price and Closing Price Date; the
updater never overwrites an unquoted row's price, so a manual price typed
into F simply persists. Skipping escalation on single-source days loses
nothing — the thresholds are in days, not runs. v1.4: ISINs consumed by a
restructure carry status Merged/Renamed instead and are EXEMPT from this
escalation (§6.15) — their absence is expected, not distress.
```

Status + Last Traded live in Stock_Master columns D/E (written only for held
ISINs). Equity surfaces both flags with conditional formats: stale via
`TODAY()-$H4>7` on the price cells, suspended/delisted via an INDEX/MATCH
status lookup on the Scrip cell — both live formulas, no stored flags.

### 6.6 FMV 31-01-2018 fallback (v1)

For an Equity row with Quantity and Cost date but **blank Avg. cost**:

```
if Cost date < 2018-02-01 and Avg. cost is blank:
    fmv = FMV by ISIN, else FMV by exchange symbol   # ISIN may have been
                                                     # reissued post-split
    if fmv: write it into E (amber format + explanatory comment),
            set the Q-column flag to "FMV"
else: row stays without Invested/XIRR (as today)
```

The Q flag makes the fallback round-trip regeneration (the cell keeps its
amber + comment and is never mistaken for a user-typed cost), and lets a
future capital-gains report apply the true grandfathering rule (higher of
cost vs min(FMV, sale price)).

### 6.7 Corporate-action adjustment (v1)

User rows always hold **raw, as-purchased** Quantity / Avg. cost. Never
mutate them. At update time:

```
for each Equity row (isin, qty_raw, cost_raw, cost_date):
  factor = Π over actions a on isin where a.ex_date > cost_date and a.ex_date ≤ today:
      SPLIT:         old_face/new_face          (10:2 → 5)
      BONUS A:B:     1 + A/B
      CONSOLIDATION: old/new (< 1)
  qty_adj  = qty_raw · factor
  cost_adj = cost_raw / factor
```

The updater writes `factor` into the Equity **S (Adj factor)** column (blank
when 1); the sheet's Cur. val / Day chg. formulas multiply Quantity by it,
while Invested (qty_raw·cost_raw) is unchanged by construction. XIRR and the
FY-end estimate use the adjusted current value. Idempotent: recomputed from
raw + action list every run; a future-dated action has factor 1 until its
ex-date arrives.

**Demat view — zero user action:** columns **O (Qty today)** and **P (Avg
cost today)** re-express the holding in post-action terms (`D×factor`,
`E×cost_factor÷factor` — the §6.15 demerger retention applies to the basis
too) so the sheet matches the user's demat/broker app after every
split/bonus/demerger, purely from the Corporate_Actions sheet content. By
Scrip quantities and the person-sheet Equity blocks read O/P (not raw D/E).

The **Corporate_Actions sheet** is the audit trail: columns
`Symbol, ISIN, Type (dropdown), Ex-Date, Ratio From, Ratio To, Factor
(=IF(type="BONUS",1+E/F,E/F), computed), Source, Details` (+ the §6.15
restructure columns), data rows 4..203. Auto rows are rewritten from the
feed each run; Manual rows are user inputs and persist (they also override
an Auto row with the same isin/type/ex-date). **Row order & capacity:**
Manual and Curated rows are written FIRST — they carry user data and the
§6.15 Applied stamps and must never fall past the last row. If the assembly
still exceeds capacity, the OLDEST Auto rows (by ex-date) are dropped and
the run warns with the count — never a silent truncation.

### 6.8 Expected value at FY-end (v1)

FY end = next 31 March ≥ today. Per holding:

```
FD    : same compound formula with YEARFRAC(Start, min(FYend, Maturity))
PPF   : Balance·(1+Rate%)^(YEARFRAC(as-on, FYend))
EPF   : Balance·(1+Rate%)^(YEARFRAC(as-on, FYend))            (v1.3)
Bonds : Qty·CurrentPrice + coupons falling in (today, FYend]   (redemption if Maturity ≤ FYend: Qty·Face instead)
Equity/MF: CurVal·(1+ExpectedReturn%)^(YEARFRAC(today, FYend)) — estimate,
           driven by the Dashboard "Expected return %" input
Gold & Silver / NPS: market-linked — same ExpectedReturn% growth   (v1.3)
Manual (Property/Cash/Insurance/Other): held FLAT at Current value (v1.3 —
           estimating property or surrender-value appreciation would be
           false precision; the header comment says so)
```

Classes switched off in Settings contribute nothing (v1.4.3, §2.1).
Aggregated per person + TOTAL into the Dashboard `Expected @ 31-Mar-<FY>`
column; the estimate nature is stated in the header comment.

### 6.9 Red/green rules (v1)

Applied by the generator as column-range conditional formats (see §3.2).
Precedence: amber (data quality) overrides red/green on the affected cells.

### 6.10 PPF interest — optional ledger + fallback (v1.1)

Official rule: interest accrues each month on the **minimum balance between
the close of the 5th and the last day of the month**, and is **credited on 31
March**. A deposit on/before the 5th earns that month; a later one does not.
Rates are bundled (`data/ppf_rates.csv`, `from_date,rate_pct` ascending step
table; quarterly since Apr-2016, annual before) and refreshed via releases —
no API exists. No withdrawals modelled (accumulation), so the monthly minimum
is the balance as of the 5th.

```
ppf_value(deposits, rates, as_of) -> (balance, total_interest):
  walk months from the first deposit to as_of
  each completed month: interest = (credited + deposits_on_or_before_5th)
                                   * rate(mid-month)/1200 ; accrue
  add all the month's deposits to the credited balance
  at each 31 March: credit the FY's accrued interest (annual compounding)
  balance = credited + interest accrued since the last 31 March
```

**Optional ledger:** the **PPF_Ledger** sheet holds one row per deposit
(Owner, Account No., Date, Amount). An account with matching ledger rows gets
exact `balance_today`, `interest_earned` and a real dated-cashflow XIRR,
written by the updater into PPF columns H/I/J. An account with no ledger keeps
today's behaviour: Balance today (H) is the live formula `=IF($D="","",$D)`
(the typed Current Balance). The updater also auto-fills a blank Rate% (F)
with the current bundled rate. Dashboard and person-sheet PPF totals sum
**Balance today (H)**, so both paths flow through identically. Class PPF XIRR
and the FY-end estimate use ledger accrual where a ledger exists, else the
flat estimate.

### 6.11 Net-worth history (v1.1; label-keyed since v1.3/R10)

The updater records one dated snapshot per run into the **History** sheet.
**Columns are label-keyed, not positional**: the header row is `Date` + the
label of every class that is enabled OR carries nonzero history (recorded
numbers are never dropped by a toggle) + `Total` (`=SUM` across that row's
class columns). The reader maps columns back by header label ("Real Estate"
accepted for Property) — unknown labels are ignored, absent classes read as
0 — so a pre-v1.3 workbook (fixed five columns) reads losslessly and old
totals recompute identically. Per-class values are computed in Python to
mirror the Dashboard (equity qty×factor×close, MF units×NAV, FD compound
value, PPF Balance-today, bonds qty×price; FD uses actual/365, a hair off
Excel's 30/360 — immaterial for a trend). **v1.4.3:** before the snapshot
is stored, every switched-off class is zeroed — "not counted" holds for the
trend too, and the run's one-line warning states the measured value that
was left out. Old rows keep whatever they recorded (they were true then).
**One row per calendar day**: a re-run on the same day overwrites that
day's row; rows are capped to the most recent `HISTORY_LAST_ROW-3`. The
Dashboard carries a line chart over History Date × Total plus a
stacked-area chart whose series cover only the currently-shown classes
("Net worth by class over time"). History rows are **data** — the reader
loads them and the generator writes them back, so they survive
regeneration.

### 6.12 Dividend quantity at ex-date (v1.2, R9)

```
qty_est(owner, isin, ex) =
  Σ over Equity rows r where
        (r.isin == isin OR resolve(r.isin) as of ex−1 == isin)   # §6.15
        AND r.owner == owner
        AND (r.cost_date is blank OR r.cost_date < ex):
    r.qty × chained_adjustment_factor(r.isin, r.cost_date, ex − 1 day, actions)
```

The CA factor is evaluated **as of the day before the ex-date**, so a
split/bonus between purchase and dividend ex-date adjusts the count, while
later actions do not. The chain-aware form (§6.15) makes a lot still keyed
to a merged-away ISIN earn the SUCCESSOR's dividends at the merger-adjusted
share count. Blank cost-date lots count as held (consistent with the
FMV-era treatment, §6.6). **Known, documented limitation:** there is no sell
ledger, so the estimate projects the *current* rows backwards — rows deleted
or reduced after a sale make history wrong. Hence the amber "(est.)"
formatting on Qty/Amount, the sheet-hint sentence, and the user's ability to
correct the Qty on a Manual row. An event yields one row per owner with
qty_est > 0; each row's Est. amount = rate × qty (a live formula).

### 6.13 Allocation drift & rebalance hint (v1.3, R11)

```
for each effective-enabled class c with a non-blank Target %:
  actual_c  = value_c / family_total            (live formula)
  drift_c   = actual_c − target_c               (percentage points, absolute)
  verdict_c = |drift_c| ≤ tolerance   → green, "On target"
              otherwise               → red, "Move ₹|drift_c × total| out|in"
classes with a blank target show nothing (no drift, no hint, no CF)
sanity: Settings B18 = Σ targets, amber when non-zero and ≠ 100
```

Tolerance (Settings B17, default 5) is **absolute percentage points** —
relative bands over-trigger on small sleeves (a 2% gold sleeve at 3% is 50%
relative drift but irrelevant money). The hint is deliberately class-level,
gross and pre-tax (the header comment says so); lot-level tax-aware selling
belongs to the capital-gains roadmap item. Everything here is Excel formulas
— correct the moment the user edits a holding, without running the updater.

### 6.14 Bullion rate application (v1.3, R13)

```
per Gold_Silver row, at update time:
  Type = SGB           → Rate today = merged-bhavcopy close for the ISIN
  Type = Gold | Silver → Rate today = §5.7 layered ₹/g rate for the metal
any rate written this run → I2 "Rates as on" stamp = run date
no rate obtainable       → row keeps its previous Rate today; the stamp
                           keeps its old date; amber CF fires past 7 days
valuation (sheet formula, live): Cur. val = Qty × (Purity or 1)
                                 × (Rate override if set, else Rate today)
```

The Rate-override precedence is a sheet formula, not updater logic — a user
typing their jeweller's rate sees the value change instantly.

### 6.15 Restructure engine — mergers / demergers / ISIN changes (v1.4, R14)

Events come from §5.8 (Curated) and Manual rows; a Manual row overrides a
Curated one with the same `(old_isin, type, ex_date, new_isin)` key —
`new_isin` is part of the key because a demerger's retention row and child
rows share the first three fields and must track their Applied dates
independently. Curated events are surfaced when their old ISIN is held
**directly or via a restructure chain** (a demerger announced on a
merger-successor concerns the old-ISIN lots too). Routing/pricing runs
**before pricing**; demerger child creation runs **after the §5.4
corporate-actions refresh** (see below). The engine NEVER edits a user cell.

**MERGER / ISIN_CHANGE — no new rows.**

```
resolve(isin): follow old→new hops (ex_date ≤ today), cycle-capped at 10
pricing     : the row's close/prev/date come from resolve(isin)'s quote
qty factor  : the merger ratio (A new per B old = A/B) folds into the
              existing Adj factor S via the chain-aware factor — later
              splits/bonuses on the SUCCESSOR keep applying to the row
cost basis  : untouched — Invested and Cost date carry in full (Sec. 47);
              Flags column shows "MERGED→<name>" / "ISIN→<isin>"
status      : the consumed ISIN's Stock_Master status becomes Merged /
              Renamed, which EXEMPTS it from the §6.5 Suspended/Delisted
              escalation (its bhavcopy absence is expected, not distress)
```

**DEMERGER — the one case needing two mechanisms.** For each event (parent
row `new_isin = old_isin` with the retention `cost_pct`; child rows with
their own `new_isin`, ratio and `cost_pct`):

```
1. Cost factor (column T, recomputed each run like S):
   per equity row with cost_date < ex ≤ today:
     T = Π retention cost_pct/100 over such events (chain-aware)
   Invested = Quantity × Avg. cost × T          (the user's cells stay put)
2. Append-once child rows (per owner-lot × child), on the FIRST run where
   ex ≤ today and the event's Applied is blank. Runs AFTER the §5.4
   corporate-actions refresh — a fresh workbook's sheet holds no history
   yet, and a child qty frozen from an incomplete table would be wrong
   forever. Matching is chain-aware: a lot demerges when its own ISIN OR
   resolve(its ISIN) as of ex−1 equals the event's old ISIN.
     qty       = lot qty × chained_adjustment_factor(lot_isin, cost_date,
                 ex−1) × A/B            (merger ratios fold into the count)
     avg cost  = lot raw invested × Π(retention cost_pct/100 of every
                 EARLIER demerger on the chain, i.e. cost_adjustment_factor
                 as of ex−1) × child cost_pct/100 ÷ qty     (per share) —
                 a second demerger apportions what REMAINED, not the
                 original cost, or totals inflate past 100%
     cost date = the PARENT lot's cost date        (Indian CGT: demerged
                 shares inherit the holding period)
     flag      = "DEMERGER:<old_isin>@<ex_date>"
   The event's Applied date (persisted on its Corporate_Actions row) is the
   single idempotency token: re-runs skip applied events, so a user deleting
   a child row (e.g. after selling it) is respected.
safety gates — an event that cannot apply SAFELY is skipped, warned about,
and NOT stamped, so it retries on a later run:
   · verification: on a live run, the parent ISIN must be among the ISINs
     the §5.4 fetch successfully checked (injected/test data is trusted);
   · capacity: all of an event's child rows must fit within the Equity
     sheet's data rows — children land atomically per event (a partial
     append that retried would duplicate its early rows).
invariant: parent Invested×T + Σ child Invested = original Invested
           (to the rupee; conservation is a test), and this holds through
           CHAINED events (merger→demerger, demerger→demerger)
```

The successor's `(symbol, name, isin)` joins Stock_Master immediately
(add-only compatible — a new ISIN), so appended rows resolve before listing;
an unlisted child simply has a blank price (excluded from day-change, no
escalation) and prices automatically on first bhavcopy appearance. Known
accepted wart: By-Scrip / person-sheet blocks group a merged row under its
old name until the user re-keys it — values are unaffected.

---

## 7. Updater behaviour

One entry point (`Update Portfolio`), replacing the legacy three scripts:

```
1. locate workbook (same folder as the executable; default filename, else
   the single *.xlsx present; else prompt)
2. INTERACTIVE (console runs only, i.e. stdin.isatty(); skipped when headless
   or --no-prompt): show the current people and offer to add new person
   sheet(s). Names entered here (or via repeatable --add-person NAME) are
   appended to the person list — regeneration then creates each new person's
   sheet, Dashboard row and By-Scrip column automatically (§2). Deduped
   case-insensitively; capped at the Dashboard's 10 people. v1.4: then offer
   to SHOW/HIDE asset classes — a numbered list of the registry classes with
   their current state; chosen numbers flip the Settings Yes/No (§3.14), the
   easy alternative to editing the sheet. v1.4.3: the listed state IS the
   Settings choice — "shown", "hidden", or "hidden — holds rows (not
   counted)" — and the per-toggle confirmation spells out what will happen
   (off with rows: "will be hidden — its rows are saved but won't be
   counted until you show it again"). Toggling OFF a class that holds rows
   also warns in the summary; every run additionally carries the single
   hidden-money awareness line (§2.1). Prompting must never hang or break a
   run — any error is swallowed.
3. refuse politely if the file is locked/open (detect via exclusive-open probe)
4. backup:  backups/<name>.backup-YYYYMMDD-HHMMSS.xlsx   (keep newest 10)
5. READ  — all input columns of all data sheets (hidden ones too) + persons
           + Settings (§3.14; missing sheet ⇒ defaults) + Dashboard cells
           (openpyxl read-only; header row located by matching known header
           names within rows 1–5, and dynamic Dashboard/History columns by
           header label, so user row edits/sorts never break it)
6. FETCH — AMFI, bhavcopy (BSE+NSE same-day union merge, §5.2), corporate
           actions + dividends (NSE+BSE, §5.4; merged holdings' successor
           symbols included); per-source failure ⇒ keep previous values,
           note in summary
7. COMPUTE — restructure routing before pricing and demerger children after
           the CA refresh (§6.15), masters merge (§6.4), prices/NAVs by
           ISIN, status flags (§6.5), FMV fallbacks (§6.6), corp-action
           factors (§6.7), dividend rows (§6.12: rebuild current-FY Auto,
           freeze the rest, keep unverified), PPF ledger accrual (§6.10),
           XIRR (§6.1–6.3), FY-end estimates (§6.8), net-worth snapshot
           (§6.11)
8. REGENERATE — build the complete workbook (xlsxwriter): structure from this
           spec + user inputs + computed/updater values; sheets of
           switched-off classes hidden, reference sheets per the
           Reference-lists switch, tab colours applied (§2.1/§3.2);
           atomic replace (write temp file, then swap)
9. REPORT — console summary (rows matched/unmatched per sheet, sources used,
           XIRR figures, PPF/history/added-people, backup path). v1.4: the
           console is a product surface — banner with version, live
           "Fetching …" lines during network stages, per-line icons, ANSI
           colour when the terminal supports it (NO_COLOR respected; plain
           text when redirected; emoji stripped when the console encoding
           can't carry them), a highlighted net-worth footer, and an
           always-printed version line ("update available" / "on the latest
           release" / "couldn't check"). Pause before closing when launched
           by double-click — but a pause with no stdin (scheduled/headless
           run of the packaged entry, which always passes --pause) proceeds
           silently instead of crashing; exit code 0/1
```

Round-trip invariant (the regression backbone): `generate → read → regenerate`
with no new data must be semantically identical (same cells, formulas,
validations, charts, formats).

---

## 8. Packaging & platform

- Reference implementation: Python ≥ 3.10; deps `xlsxwriter`, `openpyxl`,
  `requests` (+ `pytest` dev). openpyxl is **read-only** in this codebase.
- End-user artifacts per OS, built with PyInstaller (console app):
  Windows `Update Portfolio.exe`; macOS `networth-updater` binary + a
  `Update Portfolio.command` wrapper (`cd "$(dirname "$0")" && ./networth-updater; read -p "Done."`).
  First-run notes: Windows SmartScreen “Run anyway”; macOS right-click→Open.
  The PyInstaller spec MUST bundle every `data/*.csv` the updater reads at
  runtime (§5.5 — rates, FMV, bullion proxies, curated restructures): the
  loaders degrade silently on a missing file, so an unbundled CSV disables
  its feature only in the frozen build, invisible to dev-environment tests
  (a test asserts the spec's datas list for exactly this reason).
- Release zip layout: see [RELEASES.md](RELEASES.md).
- The workbook itself must open correctly in desktop Excel (Windows & Mac)
  and LibreOffice ≥ 7.x (charts, dropdowns, conditional formats).

## 9. Portability compliance checklist

An alternative implementation conforms if it:

1. produces a workbook matching §3 (sheets, columns, formulas, defined names,
   validations, charts, visual language);
2. implements the §5 data contracts with the same tolerant parsing;
3. reproduces §6 algorithms bit-for-bit on the shared test vectors
   (`tests/` golden files: XIRR values, corp-action scenarios, FMV cases);
4. honours the §7 updater invariants (backup-first, read-only inputs,
   atomic replace, round-trip identity, graceful source failure);
5. never transmits user data anywhere.
