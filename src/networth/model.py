"""Data model and layout constants shared by the generator and the reader.

Every dataclass field is either a user input or an updater-written value
(SPEC §1 terminology). Computed columns are Excel formulas emitted by the
generator and never stored here — that is what makes the round trip
(generate → read → regenerate) lossless.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from pathlib import Path

# ---------------------------------------------------------------- layout ----
# Row numbers are 1-based Excel rows. Data sheets: title r1, hint r2 (where
# present), header r3, data from r4 (SPEC §3.2).

HEADER_ROW = 3
FIRST_DATA_ROW = 4

EQUITY_LAST_ROW = 1503         # data rows 4..1503 (v1.7: was 253 — lots are
                               # rows, and an imported tradebook of a long
                               # family history needs the room; SPEC §3.6)
EQUITY_TOTAL_ROW = EQUITY_LAST_ROW + 2     # TOTAL sits 2 under the data
MF_LAST_ROW = 113              # v1.6.2: was 63
MF_TOTAL_ROW = MF_LAST_ROW + 2
SIP_LAST_ROW = 3003            # v1.7: was 1003 — an imported family CAS
                               # (2 investors × 8 funds × 10 years) needs
                               # verbatim room; exact history beats condensing
FD_LAST_ROW = 53
FD_TOTAL_ROW = 55
PPF_LAST_ROW = 43
PPF_TOTAL_ROW = 45
PPF_LEDGER_LAST_ROW = 503
BOND_LAST_ROW = 53
BOND_TOTAL_ROW = 55
BYSCRIP_LAST_ROW = 153          # v1.7.1: was 29 — auto-synced from Equity, so
                                # it must hold every distinct held ISIN
CA_LAST_ROW = 203               # Corporate_Actions data rows 4..203
DIV_LAST_ROW = 203              # Dividends data rows 4..203 (SPEC §3.13)
EQSELL_LAST_ROW = 203           # Equity_Sells data rows 4..203 (SPEC §3.20)
TAXRULES_LAST_ROW = 33          # Tax_Rules data rows 4..33 (SPEC §3.22)
MA_LAST_ROW = 63                # Manual_Assets data rows 4..63 (SPEC §3.18)
EPF_LAST_ROW = 43               # EPF data rows 4..43 (SPEC §3.17)
GS_LAST_ROW = 53                # Gold_Silver data rows 4..53 (SPEC §3.15)
NPS_LAST_ROW = 43               # NPS data rows 4..43 (SPEC §3.16)
HISTORY_LAST_ROW = 400          # net-worth snapshots, one per day (rows 4..400)
IMPORTMAP_LAST_ROW = 103        # Import_Map rows 4..103, both tables (§3.23)

DASH_PERSON_FIRST = 6          # Dashboard person matrix rows 6..15

# v1.6.2: the workbook's fixed tab names. A person may not shadow one (their
# tab would collide), and xlsxwriter enforces Excel's tab rules — both used
# to crash the whole update AFTER the network fetch; now names are adjusted.
RESERVED_SHEET_NAMES = frozenset({
    "Dashboard", "Projection", "Settings", "Equity", "Equity_Sells",
    "MutualFunds", "MF_SIP", "MF_Master", "Stock_Master", "Bank_Master",
    "FixedDeposits", "PPF", "PPF_Ledger", "EPF", "Bonds", "Gold_Silver",
    "NPS", "NPS_Master", "Manual_Assets", "By Scrip", "Corporate_Actions",
    "Dividends", "Capital Gains", "Tax_Rules", "History", "Guide",
    "Import_Map",
})
_SHEET_BAD_CHARS = set("[]:*?/\\")


def parse_yes_no(txt, default: bool) -> bool:
    """THE Yes/No reading (v1.6.2): yes/y ⇒ True, no/n ⇒ False
    (case-insensitive), anything else ⇒ default. The reader's wrapper adds
    a warning on garbage; the read-only peeks share this same truth so the
    interactive prompts can never disagree with the build (a 'Y' that the
    build masks but the peek reports as off would skip the password
    prompt)."""
    t = (str(txt) if txt is not None else "").strip().casefold()
    if t in ("yes", "y", "true", "on", "1"):
        return True
    if t in ("no", "n", "false", "off", "0"):
        return False
    return default


def person_sheet_name(name: str, taken: set[str]) -> str:
    """An Excel-legal TAB name for a person: ≤31 chars, none of []:*?/\\,
    no leading/trailing apostrophe, unique (case-insensitive) against
    `taken` and the fixed sheets. Only the tab is adjusted — every cell,
    total and Owner match keeps the person's full typed name."""
    # strip apostrophes AFTER truncating — a 31-char cut can itself end on
    # a ' (Excel forbids leading/trailing apostrophes in tab names)
    s = "".join("-" if c in _SHEET_BAD_CHARS else c
                for c in name.strip())[:31].strip("'") or "Person"
    low = ({t.casefold() for t in taken}
           | {t.casefold() for t in RESERVED_SHEET_NAMES})
    base, n = s, 2
    while s.casefold() in low:
        suffix = f"-{n}"
        s = base[:31 - len(suffix)] + suffix
        n += 1
    return s


def person_tab_map(persons: list[str]) -> dict[str, str]:
    """name → Excel tab for every person, in Dashboard order — THE single
    mapping (v1.6.2): the generator builds sheets from it, the reader warns
    from it, and the add-person prompt predicts from it, so no surface can
    ever disagree about which tab a person gets."""
    taken: set[str] = set()
    out: dict[str, str] = {}
    for p in persons:
        tab = person_sheet_name(p, taken)
        taken.add(tab)
        out[p] = tab
    return out


# v1.6.2: every user-input sheet's row budget, ONE table (the reader warns
# when a sheet holds more typed rows than the regenerated file can keep —
# silent truncation was data loss). (PortfolioData attr, last row, sheet.)
# Corporate_Actions / Dividends / History are updater-managed and have
# their own overflow handling.
CAPACITIES: list[tuple[str, int, str]] = [
    ("equity", EQUITY_LAST_ROW, "Equity"),
    ("mutual_funds", MF_LAST_ROW, "MutualFunds"),
    ("sip", SIP_LAST_ROW, "MF_SIP"),
    ("fixed_deposits", FD_LAST_ROW, "FixedDeposits"),
    ("ppf", PPF_LAST_ROW, "PPF"),
    ("ppf_ledger", PPF_LEDGER_LAST_ROW, "PPF_Ledger"),
    ("bonds", BOND_LAST_ROW, "Bonds"),
    ("bullion", GS_LAST_ROW, "Gold_Silver"),
    ("nps", NPS_LAST_ROW, "NPS"),
    ("epf", EPF_LAST_ROW, "EPF"),
    ("manual_assets", MA_LAST_ROW, "Manual_Assets"),
    ("equity_sells", EQSELL_LAST_ROW, "Equity_Sells"),
    ("tax_rules", TAXRULES_LAST_ROW, "Tax_Rules"),
    ("import_map", IMPORTMAP_LAST_ROW, "Import_Map"),
    ("imported_files", IMPORTMAP_LAST_ROW, "Import_Map"),
]
DASH_PERSON_LAST = 15
DASH_TOTAL_ROW = 16
PROJECTION_YEARS = 20          # Projection rows 4..24 (n = 0..20)

# Person sheets: summary from r5, then per-class holding blocks (SPEC §3.5).
# Since R10 the blocks stack dynamically over the ENABLED classes; each class
# contributes a fixed number of data rows (AssetClass.person_rows below).
PERSON_BLOCKS_START = 14

# Settings sheet (SPEC §3.14): class rows 4..15, the Reference-lists row,
# the Capital-gains-report row (v1.6), the optional balance-targets cells,
# then the privacy block (§3.19). Readers are label-driven, so the v1.6
# one-row shift is invisible to older workbooks.
SETTINGS_FIRST_ROW = 4
SETTINGS_LAST_ROW = 15
SETTINGS_REF_ROW = 16          # "Reference lists" Yes/No (masters + actions tab)
SETTINGS_CG_ROW = 17           # "Capital gains report" Yes/No (v1.6, §3.21)
SETTINGS_TOL_ROW = 19
SETTINGS_SUM_ROW = 20
SETTINGS_PRIVACY_ROW = 22      # "Privacy mask" Yes/No (••• curtain)
SETTINGS_LOCK_ROW = 23         # "Lock file (encryption)" Yes/No

TEMPLATE_FILENAME = "Family_Portfolio_Tracker.xlsx"

# In a PyInstaller one-file binary the bundled data lives under _MEIPASS.
import sys as _sys
if getattr(_sys, "frozen", False):
    DATA_DIR = Path(getattr(_sys, "_MEIPASS", ".")) / "data"
else:
    DATA_DIR = Path(__file__).resolve().parents[2] / "data"


# -------------------------------------------------------- asset classes ----

@dataclass(frozen=True)
class AssetClass:
    """One asset class — the registry entry that drives every per-class
    surface (Dashboard matrix column, allocation row, person block, History
    column, Settings row, sheet visibility). Adding a class = adding a row
    here plus its sheet writer/reader/computes (SPEC §2)."""
    key: str                      # attribute name on ClassXirr/HistorySnapshot
    label: str                    # header text everywhere
    value_col: str                # SUMIFS value range, e.g. "Equity!$I:$I"
    owner_col: str                # SUMIFS owner range, e.g. "Equity!$A:$A"
    sheets: tuple[str, ...]       # sheets hidden together when the class is off
    person_rows: int              # data rows in the person-sheet block (0 = none)
    default_enabled: bool = True
    has_xirr: bool = True
    # (range, label) extra SUMIFS criterion for classes sharing a sheet,
    # e.g. ("Manual_Assets!$B:$B", "Real Estate")
    class_filter: tuple[str, str] | None = None


def _manual(key: str, label: str) -> AssetClass:
    return AssetClass(key, label, "Manual_Assets!$G:$G", "Manual_Assets!$A:$A",
                      ("Manual_Assets",), person_rows=0, default_enabled=False,
                      has_xirr=(key != "cash"),
                      class_filter=("Manual_Assets!$B:$B", label))


ASSET_CLASSES: list[AssetClass] = [
    AssetClass("equity", "Equity", "Equity!$I:$I", "Equity!$A:$A",
               ("Equity", "By Scrip", "Dividends"), person_rows=40),
    AssetClass("mutual_funds", "Mutual Funds", "MutualFunds!$I:$I",
               "MutualFunds!$A:$A", ("MutualFunds", "MF_SIP"),
               person_rows=20),
    AssetClass("fixed_deposits", "Fixed Deposits", "FixedDeposits!$I:$I",
               "FixedDeposits!$A:$A", ("FixedDeposits",),
               person_rows=15),
    AssetClass("ppf", "PPF", "PPF!$H:$H", "PPF!$A:$A",
               ("PPF", "PPF_Ledger"), person_rows=10),
    AssetClass("epf", "EPF", "EPF!$H:$H", "EPF!$A:$A",
               ("EPF",), person_rows=10, default_enabled=False),
    AssetClass("bonds", "Bonds", "Bonds!$K:$K", "Bonds!$A:$A",
               ("Bonds",), person_rows=15),
    AssetClass("gold_silver", "Gold & Silver", "Gold_Silver!$K:$K",
               "Gold_Silver!$A:$A", ("Gold_Silver",), person_rows=10,
               default_enabled=False),
    AssetClass("nps", "NPS", "NPS!$G:$G", "NPS!$A:$A",
               ("NPS",), person_rows=10, default_enabled=False),
    _manual("real_estate", "Property"),
    _manual("cash", "Cash"),
    _manual("insurance", "Insurance"),
    _manual("other_assets", "Other"),
]

# Reference sheets (SPEC §3.14): the four name lists + the actions audit tab.
# Visible only when the Settings "Reference lists" switch is Yes — their
# formulas (type-ahead dropdowns, INDEX/MATCH lookups) work fine hidden.
REFERENCE_SHEETS = ("MF_Master", "Stock_Master", "Bank_Master", "NPS_Master",
                    "Corporate_Actions", "Import_Map")

# Manual_Assets Class-column values, in dropdown order. "Property" was
# labelled "Real Estate" before v1.4.3 — the reader canonicalises old rows.
MANUAL_CLASS_LABELS = ["Property", "Cash", "Insurance", "Other"]


@dataclass
class ClassSetting:
    """One Settings-sheet row (SPEC §3.14): the user's Yes/No plus the
    R11 allocation target."""
    enabled: bool = True
    target_pct: float | None = None


def default_class_settings() -> dict[str, "ClassSetting"]:
    return {c.key: ClassSetting(enabled=c.default_enabled)
            for c in ASSET_CLASSES}


def _manual_rows(data: "PortfolioData", label: str) -> list:
    return [r for r in data.manual_assets if r.asset_class == label]


def class_has_data(data: "PortfolioData", key: str) -> bool:
    """Whether the class holds any user rows — drives the Settings status
    text and the "hidden but not counted" notices, never visibility."""
    return bool({
        "equity": lambda: data.equity,
        "mutual_funds": lambda: data.mutual_funds or data.sip,
        "fixed_deposits": lambda: data.fixed_deposits,
        "ppf": lambda: data.ppf or data.ppf_ledger,
        "epf": lambda: data.epf,
        "bonds": lambda: data.bonds,
        "gold_silver": lambda: data.bullion,
        "nps": lambda: data.nps,
        "real_estate": lambda: _manual_rows(data, "Property"),
        "cash": lambda: _manual_rows(data, "Cash"),
        "insurance": lambda: _manual_rows(data, "Insurance"),
        "other_assets": lambda: _manual_rows(data, "Other"),
    }.get(key, lambda: True)())


def effective_enabled(data: "PortfolioData", cls: AssetClass) -> bool:
    """The user's Settings choice wins (SPEC §3.14, changed in v1.4.3):
    a class switched off is hidden and not counted even when it holds rows.
    Rows are never deleted — switching back on restores everything."""
    setting = data.class_settings.get(cls.key)
    return setting.enabled if setting else cls.default_enabled


def enabled_classes(data: "PortfolioData") -> list[AssetClass]:
    return [c for c in ASSET_CLASSES if effective_enabled(data, c)]


def off_with_data_classes(data: "PortfolioData") -> list[AssetClass]:
    """Classes switched off that still hold rows — the ones the Dashboard
    notice line and the updater's one-line warning must mention."""
    return [c for c in ASSET_CLASSES
            if not effective_enabled(data, c) and class_has_data(data, c.key)]


# ------------------------------------------------------------------ rows ----

@dataclass
class EquityRow:
    owner: str = ""
    scrip: str = ""                    # dropdown pick from Stock_Master
    qty: float | None = None
    avg_cost: float | None = None
    close: float | None = None         # updater-written
    prev_close: float | None = None    # updater-written
    close_date: date | None = None     # updater-written; drives the stale-price flag
    cost_date: date | None = None
    isin_override: str = ""            # user typed an ISIN over the lookup
    fmv_used: bool = False             # avg_cost filled from the 31-01-2018 FMV (§6.6)
    ca_factor: float | None = None     # updater-written split/bonus multiplier (§6.7)
    # updater-written (§6.15): demerger cost apportionment (Invested × this);
    # blank = 1. The user's Avg. cost cell is never rewritten.
    cost_factor: float | None = None
    # updater-written informational flag ("MERGED→<name>", "DEMERGER:<isin>@<date>")
    flag: str = ""
    # import-written (§6.18): the date this row's Quantity is true AS OF.
    # A broker holdings file reports the POST-split/bonus share count, so
    # corporate actions up to this date must never re-apply to it — the
    # adjustment window runs from here, not from the (older) Cost date.
    # Blank for typed rows: their Quantity is as-bought (as of Cost date).
    qty_asof: date | None = None


def qty_anchor(row) -> date | None:
    """The date an Equity row's Quantity is expressed in units of — the
    starting point of every corporate-action adjustment window (§6.18)."""
    return row.qty_asof or row.cost_date


@dataclass
class EquitySellRow:
    """One realised sale, self-contained (v1.6, SPEC §3.20/§6.16).

    The Equity sheet stays a NET current-holdings snapshot, so a sale is its
    own record here — the user copies the contract note and also reduces the
    Equity quantity. All prices/quantities are in SELL-TIME share units (what
    the broker's P&L shows); the engine never CA-adjusts these inputs. A blank
    buy_price on a pre-2018-02-01 purchase means "apply the 31-01-2018
    grandfathering value" (§6.6/§6.16) — computed in the report, never written
    back into the input cell.
    """
    owner: str = ""
    scrip: str = ""                    # dropdown pick from Stock_Master
    isin_override: str = ""            # user typed an ISIN over the lookup
    qty: float | None = None           # shares sold
    buy_date: date | None = None
    buy_price: float | None = None     # ₹/share; blank + pre-Feb-2018 → FMV path
    sell_date: date | None = None
    sell_price: float | None = None    # ₹/share
    notes: str = ""


@dataclass
class MFRow:
    owner: str = ""
    scheme: str = ""                   # dropdown pick from MF_Master
    current_nav: float | None = None   # updater-written
    xirr: float | None = None          # updater-written
    fund_house_override: str = ""
    isin_override: str = ""
    # v1.6 (§6.16): Equity / Debt for the Capital Gains tab; blank = Equity
    tax_type: str = ""


@dataclass
class SIPRow:
    owner: str = ""
    scheme: str = ""
    txn_date: date | None = None
    amount: float | None = None        # negative = redemption
    nav: float | None = None
    units_override: float | None = None
    fund_house_override: str = ""
    isin_override: str = ""


@dataclass
class FDRow:
    owner: str = ""
    bank: str = ""
    fd_no: str = ""
    principal: float | None = None
    rate: float | None = None
    start: date | None = None
    maturity: date | None = None
    comp_per_year: int | None = None


@dataclass
class PPFRow:
    owner: str = ""
    institution: str = ""
    account_no: str = ""
    balance: float | None = None       # fallback current balance (no-ledger path)
    as_on: date | None = None
    rate: float | None = None          # updater auto-fills the current rate if blank
    notes: str = ""
    # updater-written when a PPF_Ledger exists for this (owner, account_no):
    balance_today: float | None = None
    interest_earned: float | None = None
    xirr: float | None = None


@dataclass
class PPFLedgerRow:
    """One PPF deposit (SPEC §6.10). Optional — accounts with no ledger rows
    use the flat estimate from PPFRow.balance."""
    owner: str = ""
    account_no: str = ""
    txn_date: date | None = None
    amount: float | None = None


@dataclass
class BullionRow:
    """One gold/silver holding (SPEC §3.15). SGBs price from the bhavcopy by
    ISIN like bonds; physical metal values grams × purity × the daily ₹/g
    reference rate (§5.7) — a typed Rate override always wins."""
    owner: str = ""
    metal_type: str = ""               # SGB | Gold | Silver
    description: str = ""              # "SGB 2023-24 Ser II", "Gold coins, 2 x 10 g"
    isin: str = ""                     # SGB only
    qty: float | None = None           # SGB: units (1 unit = 1 g); metal: grams
    purity: float | None = None        # blank = 1 (SGB always 1); 22K = 0.916
    buy_price: float | None = None     # ₹ per gram/unit
    buy_date: date | None = None
    rate_auto: float | None = None     # updater-written ₹/unit
    rate_override: float | None = None # user-typed ₹/unit — wins over auto
    maturity: date | None = None       # SGB (8 years); blank for metal


@dataclass
class NPSRow:
    """One NPS account × scheme (SPEC §3.16): units × daily NAV. The XIRR is
    an approximate two-flow until a contribution ledger lands (roadmap)."""
    owner: str = ""
    pran: str = ""
    scheme: str = ""                   # dropdown pick from NPS_Master
    units: float | None = None         # from the CRA statement
    current_nav: float | None = None   # updater-written
    total_contributed: float | None = None   # optional → approx XIRR
    first_contribution: date | None = None   # optional → approx XIRR
    xirr: float | None = None          # updater-written (approximate)
    scheme_code_override: str = ""     # user typed a code over the lookup


@dataclass
class ManualAssetRow:
    """One hand-valued asset (SPEC §3.18): real estate, cash/savings,
    insurance surrender value, or anything else. The user types the current
    value; the as-on date flags staleness (amber past 90 days)."""
    owner: str = ""
    asset_class: str = ""              # Property | Cash | Insurance | Other
    description: str = ""
    institution: str = ""
    invested: float | None = None      # optional; enables Net chg. + XIRR
    cost_date: date | None = None
    value: float | None = None         # THE number: current worth in ₹
    as_on: date | None = None
    notes: str = ""


@dataclass
class EPFRow:
    """One EPF account (SPEC §3.17) — deliberately congruent with PPFRow's
    flat path: passbook balance + as-on + rate → accrued Balance today.
    Exact monthly accrual + a contribution ledger is a roadmap follow-up."""
    owner: str = ""
    establishment: str = ""            # employer / UAN
    member_id: str = ""
    balance: float | None = None       # from the EPFO passbook
    as_on: date | None = None
    rate: float | None = None          # updater auto-fills current EPFO rate
    notes: str = ""


@dataclass
class BondRow:
    owner: str = ""
    issuer: str = ""
    isin: str = ""
    qty: float | None = None
    face: float | None = None
    buy_price: float | None = None
    cur_price: float | None = None     # updater fills when the ISIN trades
    coupon: float | None = None
    maturity: date | None = None
    buy_date: date | None = None


@dataclass
class ScripRef:
    """By Scrip row — ISIN and display name are both user inputs."""
    isin: str = ""
    name: str = ""


RESTRUCTURE_TYPES = ("MERGER", "DEMERGER", "ISIN_CHANGE")


@dataclass
class CorporateAction:
    """One corporate action (SPEC §5.4/§6.7/§6.15). Auto rows are rewritten
    from the feed and Curated rows from data/restructures.csv on every update
    (Curated keeps its Applied date); Manual rows are the user's, persist,
    and override either with the same (isin, type, ex_date) key."""
    symbol: str = ""
    isin: str = ""                     # restructures: the OLD isin
    type: str = ""                     # SPLIT | BONUS | CONSOLIDATION | §6.15 types
    ex_date: date | None = None
    ratio_from: float | None = None    # SPLIT/CONSOL: old face; BONUS/MERGER/DEMERGER: A of A:B
    ratio_to: float | None = None      # SPLIT/CONSOL: new face; BONUS/MERGER/DEMERGER: B of A:B
    source: str = "Manual"             # Auto | Curated | Manual
    details: str = ""
    # §6.15 restructure fields
    new_isin: str = ""                 # successor security (DEMERGER: per row)
    new_name: str = ""
    new_symbol: str = ""
    cost_pct: float | None = None      # cost-basis apportionment (Σ=100 per event)
    applied: date | None = None        # demerger append-once idempotency token

    def factor(self) -> float:
        """Quantity multiplier (SPEC §6.7/§6.15)."""
        if self.type in ("DEMERGER", "ISIN_CHANGE"):
            return 1.0                 # parent qty unchanged / pure rename
        if not (self.ratio_from and self.ratio_to):
            return 1.0
        if self.type == "BONUS":
            return 1.0 + self.ratio_from / self.ratio_to
        # SPLIT >1, CONSOLIDATION <1, MERGER = A new shares per B old
        return self.ratio_from / self.ratio_to


def adjustment_factor(isin: str, cost_date: date | None, today: date,
                      actions: list[CorporateAction]) -> float:
    """Product of factors for actions with cost_date < ex_date ≤ today."""
    f = 1.0
    for a in actions:
        if a.isin != isin or not a.ex_date or a.ex_date > today:
            continue
        if cost_date and a.ex_date <= cost_date:
            continue
        f *= a.factor()
    return f


def _next_hop(isin: str, since: date | None, today: date,
              actions: list[CorporateAction]) -> "CorporateAction | None":
    hops = [a for a in actions
            if a.isin == isin and a.type in ("MERGER", "ISIN_CHANGE")
            and a.new_isin and a.new_isin != isin
            and a.ex_date and a.ex_date <= today
            and (not since or a.ex_date > since)]
    return min(hops, key=lambda a: a.ex_date) if hops else None


def resolve_isin(isin: str, actions: list[CorporateAction],
                 today: date) -> str:
    """Follow the MERGER/ISIN_CHANGE chain old→new→newer (SPEC §6.15),
    cycle-capped. Prices/status for a consumed ISIN route to the survivor."""
    cur: str = isin
    since: date | None = None
    for _ in range(10):
        hop = _next_hop(cur, since, today, actions)
        if hop is None:
            return cur
        cur, since = hop.new_isin, hop.ex_date
    return cur


def chained_adjustment_factor(isin: str, cost_date: date | None, today: date,
                              actions: list[CorporateAction]) -> float:
    """adjustment_factor that follows restructure chains: the merger ratio
    folds in (via factor()), and later splits/bonuses on the SUCCESSOR ISIN
    keep applying to the old-ISIN row."""
    f = 1.0
    cur, since = isin, cost_date
    for _ in range(10):
        f *= adjustment_factor(cur, since, today, actions)
        hop = _next_hop(cur, since, today, actions)
        if hop is None:
            return f
        cur, since = hop.new_isin, hop.ex_date
    return f


def cost_adjustment_factor(isin: str, cost_date: date | None, today: date,
                           actions: list[CorporateAction]) -> float:
    """Demerger cost-basis retention (SPEC §6.15): the product of the
    parent-retention cost_pct of every DEMERGER whose ex-date falls in
    (cost_date, today], following restructure chains like the qty factor.
    Merger/ISIN-change cost carries in full (Sec. 47) — factor 1."""
    f = 1.0
    cur, since = isin, cost_date
    for _ in range(10):
        for a in actions:
            if (a.type != "DEMERGER" or a.isin != cur or not a.ex_date
                    or a.ex_date > today or a.cost_pct is None):
                continue
            if a.new_isin and a.new_isin != cur:
                continue               # a child row of the event, not retention
            if since and a.ex_date <= since:
                continue
            f *= a.cost_pct / 100
        hop = _next_hop(cur, since, today, actions)
        if hop is None:
            return f
        cur, since = hop.new_isin, hop.ex_date
    return f


def load_restructures(data_dir: Path = DATA_DIR) -> list[CorporateAction]:
    """Curated mergers/demergers/ISIN changes (SPEC §5.8) — release-refreshed
    like ppf_rates/fmv; no free feed publishes swap ratios. Validates that
    every DEMERGER event's cost_pct sums to 100 and FAILS LOUDLY otherwise:
    silently wrong cost basis would corrupt capital-gains numbers."""
    import csv
    from collections import defaultdict
    from datetime import datetime as _dt

    out: list[CorporateAction] = []
    with open(data_dir / "restructures.csv", newline="", encoding="utf-8") as f:
        for r in csv.DictReader(f):
            if not (r.get("old_isin") or "").strip():
                continue
            out.append(CorporateAction(
                symbol=(r.get("old_symbol") or "").strip(),
                isin=r["old_isin"].strip(),
                type=r["type"].strip().upper(),
                ex_date=_dt.strptime(r["ex_date"].strip(), "%Y-%m-%d").date(),
                ratio_from=float(r["ratio_from"]) if r.get("ratio_from") else None,
                ratio_to=float(r["ratio_to"]) if r.get("ratio_to") else None,
                source="Curated",
                details=(r.get("details") or "").strip(),
                new_isin=(r.get("new_isin") or "").strip(),
                new_name=(r.get("new_name") or "").strip(),
                new_symbol=(r.get("new_symbol") or "").strip(),
                cost_pct=float(r["cost_pct"]) if r.get("cost_pct") else None,
            ))
    sums: dict[tuple, float] = defaultdict(float)
    for a in out:
        if a.type == "DEMERGER":
            sums[(a.isin, a.ex_date)] += a.cost_pct or 0.0
    for (isin, ex), total in sums.items():
        if abs(total - 100.0) > 1e-6:
            raise ValueError(
                f"restructures.csv: DEMERGER {isin} @ {ex} cost_pct sums to "
                f"{total}, not 100 — refusing to corrupt cost bases")
    return out


def fy_label(d: date) -> str:
    """Indian financial year of a date, e.g. 2026-06-15 → '2026-27'."""
    start = d.year if d.month >= 4 else d.year - 1
    return f"{start}-{str(start + 1)[2:]}"


@dataclass
class DividendRow:
    """One dividend event × owner (SPEC §3.13/§6.12). Auto rows whose ex-date
    falls in the CURRENT financial year are rebuilt from the feed on every
    update; prior-FY rows freeze automatically; Manual rows persist and
    override an Auto row with the same (isin, div_type, ex_date) key."""
    fy: str = ""                       # e.g. "2026-27", from the ex-date
    owner: str = ""
    scrip: str = ""
    isin: str = ""
    div_type: str = ""                 # Interim | Final | Special
    ex_date: date | None = None
    rate: float | None = None          # ₹ per share
    qty: float | None = None           # estimated holding at ex-date (§6.12)
    source: str = "Manual"             # Auto | Manual
    details: str = ""


@dataclass
class HistorySnapshot:
    """One dated net-worth snapshot (SPEC §6.11). Updater-written data that
    round-trips regeneration; one row per calendar day. One field per
    registry class (label-keyed on the sheet)."""
    snap_date: date | None = None
    equity: float = 0.0
    mutual_funds: float = 0.0
    fixed_deposits: float = 0.0
    ppf: float = 0.0
    epf: float = 0.0
    bonds: float = 0.0
    gold_silver: float = 0.0
    nps: float = 0.0
    real_estate: float = 0.0
    cash: float = 0.0
    insurance: float = 0.0
    other_assets: float = 0.0

    @property
    def total(self) -> float:
        return sum(getattr(self, c.key, 0.0) for c in ASSET_CLASSES)


@dataclass
class ClassXirr:
    """Updater-written plain values (SPEC §6.2). One field per registry
    class (cash stays None — has_xirr is false there)."""
    portfolio: float | None = None
    equity: float | None = None
    mutual_funds: float | None = None
    fixed_deposits: float | None = None
    ppf: float | None = None
    epf: float | None = None
    bonds: float | None = None
    gold_silver: float | None = None
    nps: float | None = None
    real_estate: float | None = None
    cash: float | None = None
    insurance: float | None = None
    other_assets: float | None = None


@dataclass
class Masters:
    mf_rows: list[tuple[str, str, str]] = field(default_factory=list)      # fund, scheme, isin
    stock_rows: list[tuple[str, str, str]] = field(default_factory=list)   # symbol, name, isin
    nps_rows: list[tuple[str, str, str]] = field(default_factory=list)     # code, scheme, pfm
    mf_refreshed: str = ""
    stock_refreshed: str = ""
    nps_refreshed: str = ""
    # held-ISIN trading status (SPEC §6.5): isin -> (Active|Suspended|Delisted, last traded)
    stock_status: dict[str, tuple[str, date | None]] = field(default_factory=dict)


@dataclass
class ImportMapRow:
    """One folio/account → person mapping (v1.7 import, SPEC §3.23).

    Written once by the import prompt, then persisted so the same statement
    never asks again. Owner is a normal dropdown cell — the user can fix a
    wrong answer in Excel and re-run.
    """
    source: str = ""                   # "fund statement (CAS)" / broker name
    account: str = ""                  # folio number / client id, as printed
    name_hint: str = ""                # investor name on the statement
    owner: str = ""                    # a Dashboard person


@dataclass
class ImportedFileRow:
    """One already-imported (or declined) file (SPEC §3.23) — the
    never-nag memory. Delete a row in Excel to be asked about that file
    again."""
    file: str = ""                     # file name only, never the path
    fingerprint: str = ""              # sha256[:12] of the file bytes
    imported_on: date | None = None
    decision: str = ""                 # "imported" | "skipped"


@dataclass
class PortfolioData:
    """Everything the generator needs beyond the spec itself."""
    persons: list[str] = field(default_factory=list)
    equity: list[EquityRow] = field(default_factory=list)
    mutual_funds: list[MFRow] = field(default_factory=list)
    sip: list[SIPRow] = field(default_factory=list)
    fixed_deposits: list[FDRow] = field(default_factory=list)
    ppf: list[PPFRow] = field(default_factory=list)
    ppf_ledger: list[PPFLedgerRow] = field(default_factory=list)
    epf: list["EPFRow"] = field(default_factory=list)
    bonds: list[BondRow] = field(default_factory=list)
    bullion: list["BullionRow"] = field(default_factory=list)
    nps: list["NPSRow"] = field(default_factory=list)
    manual_assets: list["ManualAssetRow"] = field(default_factory=list)
    bullion_rate_asof: date | None = None      # updater-written (§5.7)
    by_scrip: list[ScripRef] = field(default_factory=list)
    corporate_actions: list["CorporateAction"] = field(default_factory=list)
    dividends: list["DividendRow"] = field(default_factory=list)
    equity_sells: list["EquitySellRow"] = field(default_factory=list)  # v1.6 §3.20
    tax_rules: list["TaxRule"] = field(default_factory=list)           # v1.6 §3.22
    import_map: list[ImportMapRow] = field(default_factory=list)       # v1.7 §3.23
    imported_files: list[ImportedFileRow] = field(default_factory=list)
    history: list["HistorySnapshot"] = field(default_factory=list)
    inflation_pct: float = 7
    expected_return_pct: float = 10        # drives the FY-end estimate (SPEC §6.8)
    # per-class Yes/No + allocation target, from the Settings sheet (SPEC §3.14)
    class_settings: dict[str, "ClassSetting"] = field(
        default_factory=default_class_settings)
    # Settings "Reference lists" switch: shows/hides REFERENCE_SHEETS
    show_references: bool = False
    # Settings "Capital gains report" switch (v1.6 §3.21): shows/hides the
    # Equity_Sells input sheet + the Capital Gains report sheet together
    show_capital_gains: bool = False
    # privacy (SPEC §3.19): the user's two switches + the stored password
    # fingerprint (never the password). masked_at_rest mirrors the NW_Masked
    # defined name — what state the file on disk was in when read.
    privacy_enabled: bool = False          # Mask: figures show as •••
    lock_enabled: bool = False             # Lock: whole-file encryption
    privacy_hash: str = ""                 # pbkdf2-sha256$iter$salt$hash
    masked_at_rest: bool = False
    drift_tolerance_pct: float = 5         # Settings drift band (R11)
    fy_expected: dict[str, float] = field(default_factory=dict)  # updater-written, per person
    xirr: ClassXirr = field(default_factory=ClassXirr)
    masters: Masters = field(default_factory=Masters)
    # reader-detected problems (v1.6.2): text in a number cell, too many
    # rows for a sheet, an implausible date… — carried on the data object
    # (no signature churn) and surfaced into the updater's warning list.
    # NOT part of round-trip identity: a regenerated workbook starts clean.
    warnings: list[str] = field(default_factory=list)


def load_fmv(data_dir: Path = DATA_DIR) -> tuple[dict[str, float], dict[str, float]]:
    """31-01-2018 FMV (day's high, IT-Act grandfathering) keyed by ISIN and by
    symbol. The symbol map matters: some ISINs changed after later corporate
    actions (e.g. HDFC Bank post-split), so the 2018 ISIN may not match today's."""
    import csv

    by_isin: dict[str, float] = {}
    by_symbol: dict[str, float] = {}
    with open(data_dir / "fmv_2018-01-31.csv", newline="", encoding="utf-8") as f:
        rdr = csv.DictReader(f)
        for r in rdr:
            fmv = float(r["fmv"])
            by_isin[r["isin"]] = fmv
            by_symbol[r["symbol"]] = fmv
    return by_isin, by_symbol


@dataclass
class TaxRule:
    """One capital-gains regime row (SPEC §6.16), bundled in tax_rules_in.csv.

    Keyed by asset + effective_from DATE (not FY): Budget 2024 changed the
    equity rates mid-year on 2024-07-23, so the rule in force is resolved per
    SALE date. stcg_pct None = taxed at the user's slab (shown as words, no
    amount computed).
    """
    asset: str = ""                    # equity | mf_equity | mf_debt
    effective_from: date | None = None
    lt_days: int = 365                 # held longer than this = long-term
    stcg_pct: float | None = None      # None = at your slab
    ltcg_pct: float | None = None
    ltcg_exempt: float = 0.0           # §112A yearly exemption (₹)
    notes: str = ""


def load_tax_rules(data_dir: Path = DATA_DIR) -> list[TaxRule]:
    """Bundled Indian capital-gains rules (SPEC §5.5/§6.16), refreshed via app
    releases when a Budget changes them. Malformed rows raise loudly — a wrong
    tax table is worse than no tax table (load_restructures precedent)."""
    import csv

    rules: list[TaxRule] = []
    with open(data_dir / "tax_rules_in.csv", newline="", encoding="utf-8") as f:
        for r in csv.DictReader(f):
            try:
                rules.append(TaxRule(
                    asset=r["asset"].strip(),
                    effective_from=date.fromisoformat(r["effective_from"].strip()),
                    lt_days=int(r["lt_days"]),
                    stcg_pct=float(r["stcg_pct"]) if r["stcg_pct"].strip() else None,
                    ltcg_pct=float(r["ltcg_pct"]) if r["ltcg_pct"].strip() else None,
                    ltcg_exempt=float(r["ltcg_exempt_inr"] or 0),
                    notes=(r.get("notes") or "").strip(),
                ))
            except (KeyError, ValueError) as e:
                raise ValueError(f"tax_rules_in.csv: bad row {r!r}: {e}") from e
    rules.sort(key=lambda t: (t.asset, t.effective_from))
    return rules


def tax_rule_for(rules: list[TaxRule], asset: str, on: date) -> TaxRule | None:
    """The rule in force for `asset` on a given (sale) date — the newest
    effective_from that is <= the date. None if no rule covers it (e.g. a
    pre-FY-2018-19 sale: §10(38) era, tax shown as '—')."""
    best = None
    for t in rules:
        if t.asset == asset and t.effective_from and t.effective_from <= on:
            best = t
    return best


TAXRULE_ASSETS = ("equity", "mf_equity", "mf_debt")


def effective_tax_rules(user_rows: list[TaxRule]
                        ) -> tuple[list[TaxRule], list[TaxRule], list[str]]:
    """The rules the engine actually uses (SPEC §3.22): the bundled CSV is
    only the DEFAULT — the workbook's Tax_Rules rows are upserted over it by
    (asset, applies-from date), so a Budget change is an Excel edit, not an
    app release. Returns (valid rules sorted, invalid workbook rows kept for
    display, warnings). An invalid row (unknown asset / missing date) is
    never computed with and never silently dropped: it stays on the sheet
    with a warning until the user fixes it."""
    warnings: list[str] = []
    try:
        rules = load_tax_rules()
    except OSError:
        warnings.append("tax_rules_in.csv missing - using the workbook's "
                        "Tax_Rules rows alone")
        rules = []
    except ValueError:
        warnings.append("tax_rules_in.csv has a bad row - using the "
                        "workbook's Tax_Rules rows alone")
        rules = []
    by_key = {(t.asset, t.effective_from): t for t in rules}
    invalid: list[TaxRule] = []
    user_keys: set = set()
    for t in user_rows:
        asset = (t.asset or "").strip().casefold()
        if not (asset in TAXRULE_ASSETS and t.effective_from):
            invalid.append(t)
            warnings.append(
                f"Tax_Rules: the row '{t.asset or '(no asset)'} / "
                f"{t.effective_from or 'no date'}' needs an Asset of "
                "equity, mf_equity or mf_debt AND an Applies-from date - "
                "ignored for now, fix it on the Tax_Rules tab")
            continue
        # a rate outside 0-100, a negative allowance or a non-positive
        # holding period can only be a typo — never compute with it
        if (t.lt_days <= 0 or t.ltcg_exempt < 0
                or not 0 <= (t.stcg_pct or 0) <= 100
                or not 0 <= (t.ltcg_pct or 0) <= 100):
            invalid.append(t)
            warnings.append(
                f"Tax_Rules: the {asset} row from {t.effective_from} has a "
                "number that can't be right (rates are 0-100, days and "
                "allowance can't be negative) - ignored for now, fix it on "
                "the Tax_Rules tab")
            continue
        key = (asset, t.effective_from)
        if key in user_keys:
            warnings.append(
                f"Tax_Rules: two rows for {asset} / {t.effective_from} - "
                "the lower row wins; remove one")
        user_keys.add(key)
        by_key[key] = TaxRule(
            asset=asset, effective_from=t.effective_from,
            lt_days=t.lt_days, stcg_pct=t.stcg_pct, ltcg_pct=t.ltcg_pct,
            ltcg_exempt=t.ltcg_exempt, notes=t.notes)
    valid = sorted(by_key.values(), key=lambda t: (t.asset, t.effective_from))
    return valid, invalid, warnings


def load_epf_rates(data_dir: Path = DATA_DIR) -> list[tuple[int, float]]:
    """Bundled EPFO annual rates (SPEC §5.5): [(fy_start_year, rate_pct)],
    ascending. No official API — refreshed via app releases like ppf_rates."""
    import csv

    with open(data_dir / "epf_rates.csv", newline="", encoding="utf-8") as f:
        rdr = csv.DictReader(f)
        rows = [(int(r["fy_start"]), float(r["rate_pct"])) for r in rdr]
    rows.sort()
    return rows


def current_epf_rate(rates: list[tuple[int, float]] | None = None) -> float:
    rates = rates if rates is not None else load_epf_rates()
    return rates[-1][1]


def load_banks(data_dir: Path = DATA_DIR) -> list[tuple[str, str]]:
    """Bundled Indian bank list for the FD dropdown (SPEC §5.5)."""
    import csv

    with open(data_dir / "banks_in.csv", newline="", encoding="utf-8") as f:
        rdr = csv.reader(f)
        next(rdr)
        rows = [(r[0], r[1] if len(r) > 1 else "") for r in rdr if r and r[0]]
    rows.sort(key=lambda r: r[0].casefold())
    return rows


def load_masters(data_dir: Path = DATA_DIR,
                 mf_refreshed: str = "", stock_refreshed: str = "") -> Masters:
    """Load the committed seed masters (refreshed for real by the updater)."""
    import csv

    def rows(name: str) -> list[tuple[str, str, str]]:
        with open(data_dir / name, newline="", encoding="utf-8") as f:
            rdr = csv.reader(f)
            next(rdr)
            return [(r[0], r[1], r[2]) for r in rdr if len(r) >= 3 and r[2]]

    return Masters(
        mf_rows=rows("seed_mf_master.csv"),
        stock_rows=rows("seed_stock_master.csv"),
        nps_rows=rows("seed_nps_master.csv"),
        mf_refreshed=mf_refreshed,
        stock_refreshed=stock_refreshed,
        nps_refreshed=mf_refreshed,
    )
