<div align="center">

# 💰 NetWorth

### Your family's entire wealth, in one Excel file — refreshed with a double‑click.

Shares · Mutual Funds · FDs · PPF · EPF · NPS · Gold & Silver · Bonds ·
Property · Cash · Insurance — **one workbook**, live prices, honest returns,
automatic handling of the messy stuff (splits, bonuses, mergers, dividends).
All on **your own computer**. Nothing ever uploaded.

![Windows](https://img.shields.io/badge/Windows-0078D6?logo=windows&logoColor=white)
![macOS](https://img.shields.io/badge/macOS-000000?logo=apple&logoColor=white)
![Linux](https://img.shields.io/badge/Linux-FCC624?logo=linux&logoColor=black)
&nbsp;
![License](https://img.shields.io/badge/license-MIT-blue)
![Python](https://img.shields.io/badge/python-3.10%2B-blue)
![Local](https://img.shields.io/badge/your%20data-100%25%20local-brightgreen)
![Tests](https://img.shields.io/badge/tests-229%20passing-brightgreen)
![Release](https://img.shields.io/badge/release-v1.6.0-blue)

</div>

---

## The 10‑second pitch

You probably track your money in five apps, two brokers' statements, an EPFO
passbook and a guess. NetWorth is a single Excel file that already knows how
to do this. You type **what you own**; a double‑click fills in **what it's
worth** — today's prices, fund NAVs, gold rate, NPS NAV — and answers the
questions that actually matter:

> *What are we worth? Is it growing faster than inflation? Is the mix what we
> planned? What did our shares pay us this year? And if I sell — what will
> the taxman say?*

No account. No app. No cloud. A spreadsheet you fully own, forever.

---

## Contents

- [Start in five minutes](#-start-in-five-minutes)
- [What's in the workbook](#-whats-in-the-workbook)
- [The smart bits](#-the-smart-bits-what-it-quietly-does-for-you)
- [Your data stays yours](#-your-data-stays-yours)
- [For developers](#-for-developers)
- [Project documents](#-project-documents)
- [License](#-license)

---

## 🚀 Start in five minutes

You need Excel (or the free LibreOffice) and zero technical skills.

1. **Download** the zip for your computer from the **Releases** page and
   unzip it. Keep the files together.
2. **Open** `Family_Portfolio_Tracker.xlsx` and just look around. It opens
   **calm**: the five everyday tabs (shares, mutual funds, FDs, PPF, bonds)
   with sample rows showing exactly how each works. Gold & silver, EPF,
   NPS, property, cash, insurance are all there too — one switch away on
   the Settings tab, each with a worked example waiting inside.
3. **Make it yours, organically.** Replace the sample family (Amit / Priya /
   Rahul) with your names. Overtype sample rows with your own holdings —
   dropdowns find your funds and stocks, IDs fill themselves in. **Switch
   on what you own, switch off what you don't** — your choice always wins,
   and nothing is ever deleted.
4. **Save, close, double‑click `Update Portfolio`.** It backs your file up,
   fetches everything, recomputes every number, and prints a friendly
   one‑screen summary. It'll even ask if you want to add a family member or
   show/hide an asset class — just type a number.

> 🪜 **You don't have to learn it. You get to notice it.** Start with two
> stocks if you like. One day a stock splits and *Qty today* just… updates.
> A dividend lands and the Dividends tab logged it before your bank SMS
> arrived. HDFC merges into HDFC Bank and your old shares quietly become the
> new ones — at the right ratio, with your costs intact. That's the product.

Want the full tour? The **[illustrated user guide](docs/USER-GUIDE.md)**
walks through every feature with screenshots and worked examples.

*First run only:* Windows may show SmartScreen (**More info → Run anyway**),
macOS a Gatekeeper prompt (**right‑click → Open**) — the apps aren't
code‑signed yet. Nothing is installed; nothing leaves your machine.

**Upgrading later:** your workbook is yours forever. Replace the updater app;
the next run migrates your file to the newest layout with all data intact.

---

## 📊 What's in the workbook

You only ever type into **blue/yellow** cells; **grey** ones calculate
themselves. **Green** = gain, **red** = loss, **amber** = *take a look*
(a stale price, an estimate, a value you haven't refreshed in 90 days).

| Tab | What it shows |
|---|---|
| **Dashboard** | Family net worth, person × asset‑class grid, XIRR returns, inflation check, FY‑end estimate, dividends this year — plus five charts: allocation pie, **actual‑vs‑target**, per‑person, net‑worth trend, and a stacked **net worth by class over time** |
| **Projection** | Your corpus over 20 years: your return vs inflation |
| **Settings** | One **Show?** Yes/No per asset class (show only what you own — a hidden class keeps its rows but isn't counted, and one amber line on the Dashboard reminds you), a switch for the reference tabs, and your optional target allocation % per class |
| **One tab per person** | Each family member's holdings and allocation |
| **Equity** | Live prices, day/total change, per‑stock return, post‑split quantities, ▲/▼ day arrows |
| **Mutual Funds** *(+ SIP ledger)* | Fund summary auto‑built from one‑row‑per‑purchase |
| **Fixed Deposits / PPF / EPF** | Value today and at maturity; PPF interest by the official rules; EPF from your passbook balance at the declared rate |
| **Gold & Silver** | SGBs priced like shares; jewellery/coins at the **daily bullion (IBJA) rate** × grams × purity — or your jeweller's rate if you type it |
| **NPS** | Units × daily NAV from NPS Trust, scheme picked from a dropdown |
| **Bonds** | Value, maturity amount, coupon‑aware returns |
| **Manual Assets** | Property, cash, insurance surrender value, anything else — you type today's value, it joins the family total |
| **Dividends** | Every dividend your shares declared this financial year, logged automatically, with a by‑month chart — and counted in your return |
| **Equity Sells** *(v1.6, optional)* | One row per share sale, straight from your contract note — feeds the tax view and your true return (a sale with its old buy price left blank counts in the tax view only) |
| **Capital Gains** *(v1.6, optional)* | STCG & LTCG per year, the ₹1.25L tax‑free allowance you've used, an indicative tax figure, and the date each holding turns long‑term — with the pre‑2018 **grandfathering** rule applied for you. Intraday (same‑day) trades show separately as speculative income, never mixed in |
| **Tax Rules** *(v1.6, optional)* | The capital‑gains rates, holding periods and allowance the report uses — editable in your workbook, so a Budget change needs no new app version |
| **History** | A dated net‑worth snapshot per update — feeds the trend charts |
| **Guide** | The 2‑minute manual, right inside the file |

The tab strip is colour‑coded — navy for overview, teal for family members,
blue where you type, grey for the automatic tabs, gold for the Guide. The
reference tabs (stock/fund/bank/pension name lists and the corporate‑actions
audit trail) stay tucked away until you flip **Reference lists** to Yes on
Settings.

---

## 🧠 The smart bits (what it quietly does for you)

- **Both exchanges, merged.** Prices come from **BSE and NSE together** —
  the union of both bhavcopies, so NSE‑only and BSE‑only listings all price,
  and the numbers match your broker's app.
- **Splits & bonuses, automatically** — fetched from both exchanges,
  de‑duplicated, applied from each row's purchase date. Your typed numbers
  never change; *Qty today* matches your demat.
- **Mergers & demergers, correctly.** Old shares price as the new company at
  the right ratio (cost and purchase date carry in full — that's the tax
  rule). A demerger *appends* the spun‑off shares as their own row with the
  company‑notified cost split and the **inherited holding period**, and the
  audit tab shows exactly what was done.
- **Dividends, logged.** Every dividend declared on your holdings this FY
  appears on its own tab with an estimated amount — income you can *see*.
- **Am I balanced?** Give any class a target % and the Dashboard answers in
  colour: green "On target", or red with a plain hint like *"Move ₹1,20,000
  out"* — live, the moment you edit a holding.
- **PPF done properly** (monthly‑minimum‑balance rule, historical rates),
  **EPF from one passbook line**, **forgotten pre‑2018 costs** filled with
  the tax grandfathering value, **delisted stocks** flagged instead of
  silently going stale.
- **Honest returns everywhere** — XIRR from real dated cashflows: portfolio,
  per class, per fund, per stock. Since v1.6, dividends and recorded sales
  count in it too — the return you see is the return you got.
- **The taxman's view** *(v1.6, optional)* — record your sales and the
  Capital Gains tab shows STCG/LTCG per year, your ₹1.25L allowance
  headroom, and when each holding turns long‑term. Indicative — for
  planning, not for filing (it says so itself). And the rates aren't
  locked in code: they live on an editable **Tax_Rules** tab in your own
  workbook, so a Budget change is a 30‑second edit, not a new download.

---

## 🔒 Your data stays yours

Every fetch is a plain download of **public** data, started from your machine:

| What | Source |
|---|---|
| Share & SGB prices | BSE + NSE bhavcopies (merged) |
| Fund NAVs | AMFI |
| Corporate actions & dividends | NSE + BSE announcements |
| NPS NAVs | NPS Trust |
| Gold & silver rate | IBJA (with a market‑implied fallback) |

Reference data (banks, PPF & EPF rate history, the 2018 fair‑value table,
curated merger/demerger ratios) is **bundled** — refreshed by app releases,
no fetch needed. A once‑per‑run version check against GitHub Releases tells
you when a newer version exists (and says "you're on the latest" otherwise);
turn it off with `--no-update-check`.

**Nothing about your holdings is ever uploaded, anywhere.**

And since v1.5, two optional privacy layers guard the file itself — both
off until you want them, sharing one password (Settings tab):

- **Privacy mask** — every number shows as `•••` until you type your
  password in the update window. A curtain against people glancing at your
  screen; honest about being a curtain (type RESET if you forget).
- **Lock file** — real encryption (Excel's own "password to open",
  AES‑256). Without the password the file is unreadable anywhere — a lost
  laptop or a synced folder reveals nothing. Real also means: **no
  recovery if you forget the password.** Write it down.

---

## 🛠️ For developers

The product is a **specification** ([docs/SPEC.md](docs/SPEC.md)); the Python
here is its reference implementation. The workbook is a **build artifact** —
code generates it; the updater reads your inputs, fetches, recomputes, and
regenerates it. End users never need Python.

```
┌── generate.py ──►  Family_Portfolio_Tracker.xlsx   (xlsxwriter: 29 sheets, 10 charts, formulas)
│                            │
│                     you edit inputs
│                            ▼
└── update.py  ◄── reader.py (openpyxl, read‑only)
        │  fetch/  (amfi · bhavcopy BSE+NSE · corporate_actions+dividends · nps · bullion)
        │  compute/ (xirr · cashflows · ppf · projections · snapshot)
        └─►  regenerate the workbook (atomic, with a backup)
```

Everything per‑asset‑class flows from **one registry** (`ASSET_CLASSES` in
`model.py`): Dashboard columns, person blocks, History columns, Settings
rows, sheet visibility. Adding a class = one registry row + its sheet
writer/reader/computes.

### Set up once

```bash
python3 -m venv .venv && source .venv/bin/activate   # Windows: py -3 -m venv .venv && .venv\Scripts\activate
pip install -e ".[dev]"
```

Needs **Python 3.10+**. No compilers, no Excel, no admin rights.

### Everyday commands

```bash
python -m networth.generate            # build the workbook from code (sample data)
python -m networth.update <file.xlsx>  # refresh a workbook (must be closed in Excel)
pytest                                 # 229 tests — golden values, parsers, scenario
                                       #   suites per milestone, round-trip identity
```

**Dev loop for template changes:** edit `generate.py` → `python -m
networth.generate` → open the xlsx and look → `pytest`. Never hand‑edit the
generated file or save it through openpyxl (that drops the charts) — change
the code and rebuild.

### Build the double‑click apps

| Target | Command | Where to run |
|---|---|---|
| **macOS** | `packaging/build-release.sh <version>` | a Mac |
| **Windows** | `packaging/build-windows-nowine.sh <version>` | **Linux/Mac/Windows** — no Wine, no compiler |

NetWorth is pure Python, so the Windows app is assembled by *downloading*
prebuilt pieces — the official embeddable CPython plus `win_amd64` wheels —
and arranging them next to our code and data. Nothing is compiled or
emulated. Releasing: push a `v*` tag →
[`.github/workflows/release.yml`](.github/workflows/release.yml) tests on
three OSes, builds the apps and attaches them to the GitHub Release.

---

## 📚 Project documents

| Document | What it's for |
|---|---|
| [docs/USER-GUIDE.md](docs/USER-GUIDE.md) | **For users** — every feature walked through with screenshots and worked examples |
| [docs/SPEC.md](docs/SPEC.md) | **The product** — every sheet, data contract and algorithm, platform‑agnostic |
| [docs/RELEASES.md](docs/RELEASES.md) | Milestone plan & acceptance criteria (R0 → v1.6) |
| [docs/ROADMAP.md](docs/ROADMAP.md) | Everything else, with ✅ / 🚧 / ⬜ status |
| [docs/PLAN.md](docs/PLAN.md) · [docs/PLAN-v1.2.md](docs/PLAN-v1.2.md) | The approved architecture & design decisions |
| [CLAUDE.md](CLAUDE.md) | Working notes & conventions for contributors |

> **Maintainer's golden rule:** end users edit the workbook freely — that's
> the product, and the updater preserves it. Structural changes (sheets,
> columns, formulas, charts) live in `src/networth/generate.py`; rebuild to
> apply them.

---

## 📄 License

[MIT](LICENSE) — © 2026 Jay Parikh. Use it, fork it, make it yours.

*(The original Windows‑only PowerShell template this grew from is preserved
under [legacy/](legacy/) and still works.)*
