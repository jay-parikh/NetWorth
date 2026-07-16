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
MA_LAST_ROW = 63                # Manual_Assets data rows 4..63 (SPEC §3.18)
EPF_LAST_ROW = 43               # EPF data rows 4..43 (SPEC §3.17)
GS_LAST_ROW = 53                # Gold_Silver data rows 4..53 (SPEC §3.15)
NPS_LAST_ROW = 43               # NPS data rows 4..43 (SPEC §3.16)
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
    AssetClass("epf", "EPF", "EPF!$H:$H", "EPF!$A:$A",
               ("EPF",), person_rows=10, default_enabled=False),
    AssetClass("bonds", "Bonds", "Bonds!$K:$K", "Bonds!$A:$A",
               ("Bonds",), person_rows=15),
    AssetClass("gold_silver", "Gold & Silver", "Gold_Silver!$K:$K",
               "Gold_Silver!$A:$A", ("Gold_Silver",), person_rows=10,
               default_enabled=False),
    AssetClass("nps", "NPS", "NPS!$G:$G", "NPS!$A:$A",
               ("NPS", "NPS_Master"), person_rows=10, default_enabled=False),
    _manual("real_estate", "Real Estate"),
    _manual("cash", "Cash"),
    _manual("insurance", "Insurance"),
    _manual("other_assets", "Other"),
]

# Manual_Assets Class-column values, in dropdown order
MANUAL_CLASS_LABELS = ["Real Estate", "Cash", "Insurance", "Other"]


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
    """A class holding user rows is never hidden (SPEC §3.14)."""
    return bool({
        "equity": lambda: data.equity,
        "mutual_funds": lambda: data.mutual_funds or data.sip,
        "fixed_deposits": lambda: data.fixed_deposits,
        "ppf": lambda: data.ppf or data.ppf_ledger,
        "epf": lambda: data.epf,
        "bonds": lambda: data.bonds,
        "gold_silver": lambda: data.bullion,
        "nps": lambda: data.nps,
        "real_estate": lambda: _manual_rows(data, "Real Estate"),
        "cash": lambda: _manual_rows(data, "Cash"),
        "insurance": lambda: _manual_rows(data, "Insurance"),
        "other_assets": lambda: _manual_rows(data, "Other"),
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
    # updater-written (§6.15): demerger cost apportionment (Invested × this);
    # blank = 1. The user's Avg. cost cell is never rewritten.
    cost_factor: float | None = None
    # updater-written informational flag ("MERGED→<name>", "DEMERGER:<isin>@<date>")
    flag: str = ""


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
class BullionRow:
    """One gold/silver holding (SPEC §3.15). SGBs price from the bhavcopy by
    ISIN like bonds; physical metal values grams × purity × the daily ₹/g
    reference rate (§5.7) — a typed Rate override always wins."""
    owner: str = ""
    metal_type: str = ""               # SGB | Gold | Silver
    description: str = ""              # "SGB 2023-24 Ser II", "Bangles 22K"
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
    asset_class: str = ""              # Real Estate | Cash | Insurance | Other
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
        mf_refreshed=mf_refreshed,
        stock_refreshed=stock_refreshed,
    )
