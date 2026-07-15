<div align="center">

# 💰 NetWorth

### One Excel workbook for your family's entire net worth — refreshed with a double‑click.

Track **Equity · Mutual Funds · Fixed Deposits · PPF · Bonds** in a single file.
Live prices, real returns (XIRR), automatic tax & corporate‑action handling —
all on **your own computer**.

![License](https://img.shields.io/badge/license-MIT-blue)
![Platform](https://img.shields.io/badge/works%20on-Windows%20·%20macOS%20·%20Linux-lightgrey)
![Python](https://img.shields.io/badge/python-3.10%2B-blue)
![Local](https://img.shields.io/badge/your%20data-100%25%20local-brightgreen)
![Tests](https://img.shields.io/badge/tests-88%20passing-brightgreen)

</div>

---

## Contents

- [Why NetWorth](#-why-networth)
- [What's in the workbook](#-whats-in-the-workbook)
- [Quick start (for everyone)](#-quick-start-for-everyone)
- [The smart bits](#-the-smart-bits-what-it-does-for-you)
- [Your data stays yours](#-your-data-stays-yours)
- [For developers](#-for-developers)
- [Project documents](#-project-documents)
- [License](#-license)

---

## ✨ Why NetWorth

|  |  |
|---|---|
| 🔒 **100% local** | Your holdings never leave your computer. The only internet use is downloading *public* prices (AMFI, BSE, NSE). No cloud, no account, no sign‑up, no tracking. |
| 🪟🍎 **Windows & macOS** | One thing to double‑click. No Python, no setup — just Excel (or the free LibreOffice) to open the file. |
| 🧩 **Open & rebuildable** | The workbook is generated from code against a written [specification](docs/SPEC.md), so it can be re‑created — or re‑implemented in any language — from the spec alone. |

---

## 📊 What's in the workbook

One file, a tab for everything. You only ever type into the **blue/yellow** cells;
the **grey** ones calculate themselves.

| Tab | What it shows |
|---|---|
| **Dashboard** | Family net worth, a person × asset‑class grid, returns (XIRR), an inflation check, an FY‑end estimate, and charts — allocation, per‑person, and net worth over time |
| **Projection** | Your money over the next 20 years: your return vs inflation |
| **One tab per person** | Each family member's holdings and allocation |
| **Equity** | Shares with live prices, day/total change, per‑stock return, and post‑split/bonus quantities |
| **Mutual Funds** *(+ ledger)* | Fund summary, auto‑built from a one‑row‑per‑purchase SIP ledger |
| **Fixed Deposits** | Value today and at maturity, from principal / rate / dates |
| **PPF** *(+ ledger)* | Balance & interest — exact, by the official rules, if you log deposits |
| **Bonds** | Value, maturity amount, and coupon‑aware returns |
| **Corporate Actions** | A transparent record of every split/bonus applied to your stocks |
| **History** | A dated net‑worth snapshot each time you update — feeds the trend chart |
| **Guide** | A 2‑minute manual, right inside the file |

> **Colours tell the story:** blue/yellow = *you type here*, grey = *calculated*.
> **Green** = gain, **red** = loss, **amber** = *look closer* (a stale price, a
> delisted stock, or an estimated cost).

---

## 🚀 Quick start (for everyone)

You need **5 minutes** and Excel (or LibreOffice). No technical skills.

1. **Download** the zip for your computer from the **Releases** page and unzip it
   somewhere. Keep the files together.
2. **Open** `Family_Portfolio_Tracker.xlsx` and skim the **Guide** tab. Replace the
   sample family (Amit / Priya / Rahul) with your own — pick funds and stocks from
   the dropdowns and the ID (ISIN) fills itself in.
3. **Save and close** the file, then **double‑click `Update Portfolio`**
   (`.bat` / `.exe` on Windows; on macOS, right‑click `Update Portfolio.command`
   → **Open** the first time).

That's it. Each run:

> 💾 backs up your file  →  🌐 fetches the latest prices, NAVs & corporate actions
> →  🧮 recomputes every return, PPF interest and the FY‑end estimate  →  📸 saves a
> net‑worth snapshot  →  ✅ rewrites the workbook and prints a one‑screen summary.

It even **offers to add a new family member** — just type a name and their tab is
built for you. If the internet or a data source is down, your old numbers stay put
and it tells you.

*First run only:* Windows may show a SmartScreen warning (**More info → Run anyway**)
and macOS a Gatekeeper prompt (**right‑click → Open**) — the apps aren't
code‑signed yet. Nothing is installed; nothing is uploaded.

**Upgrading later** is painless: your workbook is yours forever — just replace the
updater app, and the next run brings the file up to the newest layout with all your
data intact.

---

## 🧠 The smart bits (what it does for you)

You don't have to know any of this works — it just does. But here's what's quietly
handled for you:

- **Splits & bonuses, automatically.** Corporate actions are fetched from **both
  NSE and BSE**, de‑duplicated, and applied to your holdings — past *and* future
  ex‑dates. Your typed quantities are never overwritten; a *Qty today* column shows
  the post‑bonus count, matching your demat. If a stock can't be verified on either
  exchange, the summary **says so by name** — nothing is skipped silently.
- **PPF, done properly.** List your deposits and the balance, interest and return
  are computed by the *official* rule (interest on the monthly‑minimum balance,
  historical rates). Don't want to? Just type your current balance instead.
- **Old shares with a forgotten cost.** Bought before Feb 2018 and don't know the
  price? Leave it blank — the tax "grandfathering" value (31‑Jan‑2018 fair value)
  fills in, clearly flagged.
- **Delisted / suspended stocks** are detected and marked, keeping their last known
  price instead of quietly going stale.
- **Real returns everywhere** — portfolio, per asset class, per fund, per stock —
  computed from actual dated cashflows, exactly like Excel's `XIRR`.

---

## 🔒 Your data stays yours

Every fetch is a plain download of **public** data, started from your machine:

| What | Source |
|---|---|
| Fund NAVs | AMFI (`amfiindia.com`) |
| Share prices | BSE (primary), NSE (fallback) |
| Corporate actions | NSE + BSE |

Reference data (Indian banks, PPF rate history, the 2018 fair‑value table) is
**bundled** — no fetch needed. There's also a once‑per‑run check of the GitHub
Releases page to tell you if a new version exists (it sends nothing about you; turn
it off with `--no-update-check`).

**Nothing about your holdings is ever uploaded, anywhere.**

---

## 🛠️ For developers

The product is a **specification** ([docs/SPEC.md](docs/SPEC.md)); the Python here
is its reference implementation. The workbook is a **build artifact** — code
generates it; the updater reads your inputs, fetches data, recomputes, and
regenerates it. End users never need Python.

```
┌── generate.py ──►  Family_Portfolio_Tracker.xlsx   (xlsxwriter: sheets, charts, formulas)
│                            │
│                     you edit inputs
│                            ▼
└── update.py  ◄── reader.py (openpyxl, read‑only)
        │  fetch/  (amfi · bhavcopy BSE→NSE · corporate_actions)
        │  compute/ (xirr · cashflows · ppf · projections · snapshot)
        └─►  regenerate the workbook (atomic, with a backup)
```

### Set up once

<table>
<tr><th>Windows</th><th>macOS / Linux</th></tr>
<tr><td>

```bat
py -3 -m venv .venv
.venv\Scripts\activate
pip install -e ".[dev]"
```

</td><td>

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -e ".[dev]"
```

</td></tr>
</table>

Needs **Python 3.10+** and git. No compilers, no Excel, no admin rights. `-e` is an
editable install (code changes take effect immediately).

### Everyday commands

```bash
python -m networth.generate            # build the workbook from code (sample data)
python -m networth.update <file.xlsx>  # refresh a workbook (must be closed in Excel)
pytest                                 # 88 tests — XIRR golden values, parsers,
                                       #   corp‑action & PPF scenarios, round‑trip identity
```

**Dev loop for template changes:** edit `generate.py` → `python -m networth.generate`
→ open the xlsx and look → `pytest`. Never hand‑edit the generated file or save it
through openpyxl (that drops the charts) — change the code and rebuild.

### Build the double‑click apps

| Target | Command | Where to run |
|---|---|---|
| **macOS** | `packaging/build-release.sh 1.1.0` | a Mac |
| **Windows** | `packaging/build-windows-nowine.sh 1.1.0` | **Linux/Mac/Windows** — no Wine, no compiler |

Because NetWorth is **pure Python**, the Windows app is assembled by *downloading*
prebuilt Windows pieces — the official embeddable CPython plus `win_amd64` dependency
wheels (`pip download --platform win_amd64`) — and arranging them next to our code
and data. Nothing is compiled or emulated. The launcher is `Update Portfolio.bat`
(relocatable, correct‑by‑inspection); a single self‑contained `.exe` comes from
`packaging\build-release.bat` on real Windows, or from CI.

### Pre‑release checklist (build & verify on Ubuntu)

With your dev venv active, before tagging a release:

```bash
# 0 · tests green
pytest -q

# 1 · build the Windows bundle (version must match the tag you'll cut)
packaging/build-windows-nowine.sh 1.1.0

# 2 · the zip carries genuine Windows binaries + our code + data
python -m zipfile -l dist/NetWorth-1.1.0-windows.zip \
  | grep -E "win_amd64|python\.exe|networth/update\.py|Lib/data/ppf_rates|Update Portfolio\.bat"

# 3 · the bundled code imports and the entry point is wired
PYTHONPATH="dist/NetWorth-1.1.0-windows/app/python/Lib/site-packages" \
  python -c "import networth, networth._packaged as p; print('bundled', networth.__version__, '| entry:', callable(p.run))"

# 4 · the shipped workbook opens with all its tabs
python -c "from openpyxl import load_workbook; wb=load_workbook('dist/NetWorth-1.1.0-windows/Family_Portfolio_Tracker.xlsx'); print(len(wb.sheetnames), 'sheets')"
```

Expect step 2 to list all five markers, step 3 to print `entry: True`, and step 4 to
report **19 sheets**. The only thing you *can't* verify on Linux is running the
Windows binary itself — do that final smoke test on a real Windows PC (or in CI)
before you publish.

### Releasing

Push a `v*` tag → [`.github/workflows/release.yml`](.github/workflows/release.yml)
runs the tests on Linux/Windows/macOS, builds the per‑OS apps, and attaches them to
the GitHub Release (a tag with a `-`, e.g. `v1.2.0-rc.1`, is marked pre‑release).

---

## 📚 Project documents

| Document | What it's for |
|---|---|
| [docs/SPEC.md](docs/SPEC.md) | **The product** — every sheet, data contract and algorithm, platform‑agnostic |
| [docs/RELEASES.md](docs/RELEASES.md) | Milestone plan & acceptance criteria |
| [docs/ROADMAP.md](docs/ROADMAP.md) | What's next |
| [docs/PLAN.md](docs/PLAN.md) | The approved architecture & decisions |
| [CLAUDE.md](CLAUDE.md) | Working notes & conventions for contributors |

> **Maintainer's golden rule:** end users edit the workbook freely — that's the
> product, and the updater preserves it. What must *not* happen is changing the
> template's **structure** (sheets, columns, formulas, charts) by hand. Structural
> changes live in `src/networth/generate.py`; rebuild to apply them.

---

## 📄 License

[MIT](LICENSE) — © 2026 Jay Parikh. Use it, fork it, make it yours.

*(The original Windows‑only PowerShell template this grew from is preserved under
[legacy/](legacy/) and still works.)*
