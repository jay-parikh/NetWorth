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

EQUITY_LAST_ROW = 140          # data rows 4..140, dropdown range end
EQUITY_TOTAL_ROW = 142
MF_LAST_ROW = 63
MF_TOTAL_ROW = 65
SIP_LAST_ROW = 503
FD_LAST_ROW = 53
FD_TOTAL_ROW = 55
PPF_LAST_ROW = 43
PPF_TOTAL_ROW = 45
PPF_LEDGER_LAST_ROW = 503
BOND_LAST_ROW = 53
BOND_TOTAL_ROW = 55
BYSCRIP_LAST_ROW = 29
CA_LAST_ROW = 53                # Corporate_Actions data rows 4..53
DIV_LAST_ROW = 203              # Dividends data rows 4..203 (SPEC §3.13)
HISTORY_LAST_ROW = 400          # net-worth snapshots, one per day (rows 4..400)

DASH_PERSON_FIRST = 6          # Dashboard person matrix rows 6..15
DASH_PERSON_LAST = 15
DASH_TOTAL_ROW = 16
PROJECTION_YEARS = 20          # Projection rows 4..24 (n = 0..20)

# Person sheets: summary from r5, then per-class holding blocks (SPEC §3.5).
# Since R10 the blocks stack dynamically over the ENABLED classes; each class
# contributes a fixed number of data rows (AssetClass.person_rows below).
PERSON_BLOCKS_START = 14

# Settings sheet (SPEC §3.14): class rows 4..15, then tolerance/total cells.
SETTINGS_FIRST_ROW = 4
SETTINGS_LAST_ROW = 15
SETTINGS_TOL_ROW = 17
SETTINGS_SUM_ROW = 18

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
    person_rows: int              # data rows in the person-sheet block
    default_enabled: bool = True
    has_xirr: bool = True


ASSET_CLASSES: list[AssetClass] = [
    AssetClass("equity", "Equity", "Equity!$I:$I", "Equity!$A:$A",
               ("Equity", "By Scrip", "Corporate_Actions", "Dividends",
                "Stock_Master"), person_rows=40),
    AssetClass("mutual_funds", "Mutual Funds", "MutualFunds!$I:$I",
               "MutualFunds!$A:$A", ("MutualFunds", "MF_SIP", "MF_Master"),
               person_rows=20),
    AssetClass("fixed_deposits", "Fixed Deposits", "FixedDeposits!$I:$I",
               "FixedDeposits!$A:$A", ("FixedDeposits", "Bank_Master"),
               person_rows=15),
    AssetClass("ppf", "PPF", "PPF!$H:$H", "PPF!$A:$A",
               ("PPF", "PPF_Ledger"), person_rows=10),
    AssetClass("bonds", "Bonds", "Bonds!$K:$K", "Bonds!$A:$A",
               ("Bonds",), person_rows=15),
]


@dataclass
class ClassSetting:
    """One Settings-sheet row (SPEC §3.14): the user's Yes/No plus the
    R11 allocation target."""
    enabled: bool = True
    target_pct: float | None = None


def default_class_settings() -> dict[str, "ClassSetting"]:
    return {c.key: ClassSetting(enabled=c.default_enabled)
            for c in ASSET_CLASSES}


def class_has_data(data: "PortfolioData", key: str) -> bool:
    """A class holding user rows is never hidden (SPEC §3.14)."""
    return bool({
        "equity": lambda: data.equity,
        "mutual_funds": lambda: data.mutual_funds or data.sip,
        "fixed_deposits": lambda: data.fixed_deposits,
        "ppf": lambda: data.ppf or data.ppf_ledger,
        "bonds": lambda: data.bonds,
    }.get(key, lambda: True)())


def effective_enabled(data: "PortfolioData", cls: AssetClass) -> bool:
    setting = data.class_settings.get(cls.key)
    enabled = setting.enabled if setting else cls.default_enabled
    return enabled or class_has_data(data, cls.key)


def enabled_classes(data: "PortfolioData") -> list[AssetClass]:
    return [c for c in ASSET_CLASSES if effective_enabled(data, c)]


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


@dataclass
class MFRow:
    owner: str = ""
    scheme: str = ""                   # dropdown pick from MF_Master
    current_nav: float | None = None   # updater-written
    xirr: float | None = None          # updater-written
    fund_house_override: str = ""
    isin_override: str = ""


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


@dataclass
class CorporateAction:
    """One split/bonus/consolidation (SPEC §5.4/§6.7). Auto rows are rewritten
    from the feed on every update; Manual rows are the user's and persist."""
    symbol: str = ""
    isin: str = ""
    type: str = ""                     # SPLIT | BONUS | CONSOLIDATION
    ex_date: date | None = None
    ratio_from: float | None = None    # SPLIT/CONSOLIDATION: old face; BONUS: A
    ratio_to: float | None = None      # SPLIT/CONSOLIDATION: new face; BONUS: B
    source: str = "Manual"             # Auto | Manual
    details: str = ""

    def factor(self) -> float:
        """Quantity multiplier (SPEC §6.7)."""
        if not (self.ratio_from and self.ratio_to):
            return 1.0
        if self.type == "BONUS":
            return 1.0 + self.ratio_from / self.ratio_to
        return self.ratio_from / self.ratio_to   # SPLIT >1, CONSOLIDATION <1


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
    round-trips regeneration; one row per calendar day."""
    snap_date: date | None = None
    equity: float = 0.0
    mutual_funds: float = 0.0
    fixed_deposits: float = 0.0
    ppf: float = 0.0
    bonds: float = 0.0

    @property
    def total(self) -> float:
        return (self.equity + self.mutual_funds + self.fixed_deposits
                + self.ppf + self.bonds)


@dataclass
class ClassXirr:
    """Updater-written plain values (SPEC §6.2)."""
    portfolio: float | None = None
    equity: float | None = None
    mutual_funds: float | None = None
    fixed_deposits: float | None = None
    ppf: float | None = None
    bonds: float | None = None


@dataclass
class Masters:
    mf_rows: list[tuple[str, str, str]] = field(default_factory=list)      # fund, scheme, isin
    stock_rows: list[tuple[str, str, str]] = field(default_factory=list)   # symbol, name, isin
    mf_refreshed: str = ""
    stock_refreshed: str = ""
    # held-ISIN trading status (SPEC §6.5): isin -> (Active|Suspended|Delisted, last traded)
    stock_status: dict[str, tuple[str, date | None]] = field(default_factory=dict)


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
    bonds: list[BondRow] = field(default_factory=list)
    by_scrip: list[ScripRef] = field(default_factory=list)
    corporate_actions: list["CorporateAction"] = field(default_factory=list)
    dividends: list["DividendRow"] = field(default_factory=list)
    history: list["HistorySnapshot"] = field(default_factory=list)
    inflation_pct: float = 7
    expected_return_pct: float = 10        # drives the FY-end estimate (SPEC §6.8)
    # per-class Yes/No + allocation target, from the Settings sheet (SPEC §3.14)
    class_settings: dict[str, "ClassSetting"] = field(
        default_factory=default_class_settings)
    drift_tolerance_pct: float = 5         # Settings B17 (R11 drift band)
    fy_expected: dict[str, float] = field(default_factory=dict)  # updater-written, per person
    xirr: ClassXirr = field(default_factory=ClassXirr)
    masters: Masters = field(default_factory=Masters)


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
        mf_refreshed=mf_refreshed,
        stock_refreshed=stock_refreshed,
    )
