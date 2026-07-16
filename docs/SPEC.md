# NetWorth — Product Specification

> **In one sentence:** this document describes *exactly* what the workbook and
> the updater must do — every tab, every number, every data source — so that
> anyone could rebuild NetWorth from scratch, in any language, without seeing
> the code.

**Version:** covers v1.0 + v1.1 · **Status:** normative.
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

---

## 3. Workbook specification

### 3.1 Sheet map (tab order)

| # | Sheet | Kind | Purpose |
|---|---|---|---|
| 1 | Dashboard | mixed | family net worth, per-person × class matrix, XIRR, inflation, charts |
| 2 | Projection | computed | 20-year corpus trajectory table + line chart |
| 3… | one per person (e.g. Amit) | computed | that person's allocation + pie chart |
| … | Equity | data entry | stock holdings |
| … | MutualFunds | computed summary | one row per (owner, scheme), derived from MF_SIP |
| … | MF_SIP | data entry | one row per MF purchase/redemption |
| … | MF_Master | master | AMFI scheme list (~14k rows) |
| … | Stock_Master | master | listed-stock list (~4.5k rows) |
| v1 | Bank_Master | master | bundled Indian bank list (Bank Name, Type; sorted; §3.11) |
| … | FixedDeposits | data entry | FDs |
| … | PPF | data entry | PPF accounts |
| … | Bonds | data entry | corporate/other bonds |
| … | By Scrip | computed | family-wide exposure per stock |
| … | Guide | static text | 2-minute manual |
| v1 | Corporate_Actions | audit (§6.7) | fetched + manual corporate actions and their effect |

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

### 3.3 Dashboard

Layout (columns A–G):

| Cell(s) | Content | Kind |
|---|---|---|
| A1 | `FAMILY PORTFOLIO — NET WORTH TRACKER` | static |
| A2 / B2 | `As on` / `=TODAY()` | computed |
| A3 / B3 | `Family net worth` / `=G16` | computed |
| E3 | inflation % p.a. input (default 7) | **input** |
| B4 | Portfolio XIRR across all classes | **updater-written** |
| E4 | real return `=IF(B4="","",(1+B4)/(1+E3/100)-1)` | computed |
| F4 | verdict `=IF(B4="","",IF(B4>E3/100,"Beats inflation ✓","Below inflation ✗"))` | computed |
| A5:G5 | headers `Person, Equity, Mutual Funds, Fixed Deposits, PPF, Bonds, Total` | static |
| A6:A15 | up to 10 person names | **input** (pre-filled from `persons`) |
| B6:F15 | per class: `=IF($A6="","",SUMIFS(<class value col>, <class owner col>, $A6))` — Equity!I, MutualFunds!I, FixedDeposits!I, PPF!D, Bonds!K | computed |
| G6:G15 | `=IF($A6="","",SUM(B6:F6))` | computed |
| A16, B16:G16 | `TOTAL` + column sums | computed |
| A18 | `Allocation by asset class` | static |
| A19:B19 | headers `Asset class, Value` | static |
| A20:B24 | Equity `=B16`, Mutual Funds `=C16`, Fixed Deposits `=D16`, PPF `=E16`, Bonds `=F16` | computed |
| C20:C24 | per-class XIRR | **updater-written** |
| v1: D2/E2 | `Expected return % p.a.` label + input (default 10) for the FY-end estimate | **input** |
| v1: H5, H6:H15 | `Expected @ 31-Mar-<FY>` header + per-person values (§6.8) | **updater-written** |
| v1: H16 | total `=IF(SUM(H6:H15)=0,"",SUM(H6:H15))` | computed |

Charts on Dashboard: **pie** “Allocation by asset class” (A20:B24) and
**bar** “Net worth by person” (A6:A15 vs G6:G15).

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
| J | Invested | computed | `=IF(OR($D4="",$E4=""),"",$D4*$E4)` |
| K | Net chg. | computed | `=I−J` guarded |
| L | Day chg. | computed | `=IF(OR($G4="",$D4=""),"",$D4*($F4-$G4))` |
| M | Cost date | input | drives per-row return annualisation & XIRR cashflows |
| N | XIRR (per row) | computed | `=IF(OR($M4="",N($J4)=0,$I4="",TODAY()<=$M4),"",($I4/$J4)^(365/(TODAY()-$M4))-1)` — simple two-flow annualisation |
| v1: O | Qty today | computed | `=IF($D4="","",$D4*IF($S4="",1,$S4))` — post-split/bonus share count, the **demat view**; feeds By Scrip and the person sheets |
| v1: P | Avg cost today | computed | `=IF(OR($D4="",$E4=""),"",$E4/IF($S4="",1,$S4))` — cost per share in today's share terms |
| Q | Key | computed helper | `=IF($A4="","",$A4&"#"&COUNTIF($A$4:$A4,$A4))` stable per-owner sequence id |
| v1: R | Flags | updater helper | `FMV` marks an avg-cost filled by the §6.6 fallback so the flag round-trips regeneration |
| v1: S | Adj factor | **updater-written** | split/bonus multiplier since Cost date (§6.7); blank = 1. `Cur. val` and `Day chg.` use `Quantity*IF($S4="",1,$S4)*price`; `Invested` stays raw |

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

---

## 4. Sample data

The released template ships with fictional holdings for three people (Amit,
Priya, Rahul) using **real ISINs** so the first updater run works end-to-end.
MF samples must be real AMFI (Scheme Name, Fund House, ISIN) triples; equity
samples real BSE scrips. Sample rows are ordinary input rows — deleting them
is the onboarding step.

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
`https://www.nseindia.com/` first.

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
<face>" → SPLIT; "consolidation" → CONSOLIDATION; everything else (dividends,
rights, AGMs, buybacks) is ignored. The normative contract is the record, not
the URLs:

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

**Coverage rule (never skip silently):** the fetch reports which ISINs were
successfully answered by at least one exchange. Any held ISIN answered by
neither MUST surface as a user-visible warning naming the scrip ("corporate
actions could NOT be verified for: …"), so an unverifiable holding is a known
condition, not a silent gap. One security failing must not abort the sweep;
only an all-security/all-source failure degrades (keep existing rows).
Mergers/demergers/ISIN reassignments remain out of scope for auto-adjustment
(no reliable free feed) — the Manual-row path covers them.

### 5.5 Bundled static data (in `data/`, refreshed only by releases)

| File | Shape | Source |
|---|---|---|
| `fmv_2018-01-31.csv` | `isin, symbol, fmv` | NSE bhavcopy of 2018-01-31 (EQ/BE/BZ series, 1,639 ISINs); FMV = that day's **high** price (IT Act grandfathering definition). The symbol column enables lookup when a later corporate action reissued the ISIN (e.g. HDFC Bank post-split) |
| `banks_in.csv` | `bank_name, type` | RBI scheduled commercial banks list + major SFB/payment/co-op banks |
| `ppf_rates.csv` *(roadmap — ships with the PPF contribution ledger)* | `from_date, to_date, rate_pct` | MoF quarterly notifications, historical to present (no official API exists) |

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
| Equity | −Invested @ Cost date | +Cur. val @ today |
| Mutual Funds | one per MF_SIP row: −Amount @ Date (redemptions are +) | +Cur. val @ today per (owner, ISIN) |
| Fixed Deposits | −Principal @ Start | +Value-as-on @ min(today, maturity) |
| PPF | −(Balance discounted at Rate% back to as-on date)… in practice: −Balance @ as-on | +Balance·(1+Rate%)^(days/365) @ today |
| Bonds | −Qty·BuyPrice @ Buy Date (skip if no Buy Date) | +Qty·CurrentPrice @ today; **v1:** plus each coupon `+Qty·Face·(Coupon%/f)` on its historical coupon date |

Class XIRR = solver over that class's union; Portfolio XIRR = solver over the
union of **all** classes. Written as plain values to: Dashboard B4, C20:C24,
Equity class cell, MutualFunds L column + MF_SIP J2.

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
nothing — the thresholds are in days, not runs.
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
`E÷factor`) so the sheet matches the user's demat/broker app after every
split/bonus, purely from the Corporate_Actions sheet content. By Scrip
quantities and the person-sheet Equity blocks read O/P (not raw D/E).

The **Corporate_Actions sheet** is the audit trail: columns
`Symbol, ISIN, Type (dropdown), Ex-Date, Ratio From, Ratio To, Factor
(=IF(type="BONUS",1+E/F,E/F), computed), Source, Details`. Auto rows are
rewritten from the feed each run; Manual rows are user inputs and persist
(they also override an Auto row with the same isin/type/ex-date).

### 6.8 Expected value at FY-end (v1)

FY end = next 31 March ≥ today. Per holding:

```
FD    : same compound formula with YEARFRAC(Start, min(FYend, Maturity))
PPF   : Balance·(1+Rate%)^(YEARFRAC(as-on, FYend))
Bonds : Qty·CurrentPrice + coupons falling in (today, FYend]   (redemption if Maturity ≤ FYend: Qty·Face instead)
Equity/MF: CurVal·(1+ExpectedReturn%)^(YEARFRAC(today, FYend)) — estimate,
           driven by the Dashboard "Expected return %" input
```

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

### 6.11 Net-worth history (v1.1)

The updater records one dated snapshot per run into the **History** sheet
(Date, Equity, Mutual Funds, Fixed Deposits, PPF, Bonds; Total is
`=SUM(B:F)`). Per-class values are computed in Python to mirror the Dashboard
(equity qty×factor×close, MF units×NAV, FD compound value, PPF Balance-today,
bonds qty×price; FD uses actual/365, a hair off Excel's 30/360 — immaterial
for a trend). **One row per calendar day**: a re-run on the same day overwrites
that day's row; rows are capped to the most recent `HISTORY_LAST_ROW-3`. The
Dashboard carries a line chart over History Date × Total. History rows are
**data** — the reader loads them and the generator writes them back, so they
survive regeneration.

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
   case-insensitively; capped at the Dashboard's 10 people. Prompting must
   never hang or break a run — any error is swallowed.
3. refuse politely if the file is locked/open (detect via exclusive-open probe)
4. backup:  backups/<name>.backup-YYYYMMDD-HHMMSS.xlsx   (keep newest 10)
5. READ  — all input columns of all data sheets + persons + settings cells
           (openpyxl read-only; header row located by matching known header
           names within rows 1–5, so user row edits/sorts never break it)
6. FETCH — AMFI, bhavcopy (BSE+NSE same-day union merge, §5.2), corporate
           actions (NSE+BSE); per-source failure ⇒ keep previous values,
           note in summary
7. COMPUTE — masters merge (§6.4), prices/NAVs by ISIN, status flags (§6.5),
           FMV fallbacks (§6.6), corp-action factors (§6.7), PPF ledger
           accrual (§6.10), XIRR (§6.1–6.3), FY-end estimates (§6.8),
           net-worth snapshot (§6.11)
8. REGENERATE — build the complete workbook (xlsxwriter): structure from this
           spec + user inputs + computed/updater values; atomic replace
           (write temp file, then swap)
9. REPORT — console summary (rows matched/unmatched per sheet, sources used,
           XIRR figures, PPF/history/added-people, backup path); pause before
           closing when launched by double-click; a non-blocking GitHub
           release check may print an update hint; exit code 0/1
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
