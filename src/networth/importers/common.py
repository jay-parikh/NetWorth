"""Normalized import models + hardened Indian-format parsing (SPEC §6.17).

Every parser (CAS PDF, broker CSV) reduces its file to these dataclasses;
the merge engine never sees a source format. The number/date helpers are
deliberately strict: anything outside the recognised shapes returns None
so the caller can quarantine the line — a wrong value must never survive
by leniency (the never-garbage contract).
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date, datetime

# sanity bounds (§6.17): outside these, a "number" is a misparse
EARLIEST_DATE = date(1990, 1, 1)
MAX_AMOUNT = 1_000_000_000.0           # ₹100 crore on one line = misparse
UNIT_TOLERANCE = 0.001                 # balance reconciliation, in units


@dataclass
class ImportedSipTxn:
    """One mutual-fund transaction from a statement."""
    folio: str = ""
    isin: str = ""
    scheme_name: str = ""
    fund_house: str = ""
    txn_date: date | None = None
    amount: float | None = None        # negative = redemption / switch-out
    nav: float | None = None
    units: float | None = None         # signed like amount
    txn_type: str = ""                 # PURCHASE|REDEMPTION|SWITCH_IN|SWITCH_OUT|DIV_REINVEST


@dataclass
class ImportedTrade:
    """One (already collapsed) broker equity trade."""
    account: str = ""
    isin: str = ""
    symbol: str = ""
    trade_date: date | None = None
    qty: float | None = None           # always positive; side carries direction
    price: float | None = None
    side: str = ""                     # BUY | SELL


@dataclass
class ImportedHolding:
    """One current-holdings line (broker holdings export)."""
    account: str = ""
    isin: str = ""
    name: str = ""
    qty: float | None = None
    avg_cost: float | None = None


@dataclass
class ImportBatch:
    """Everything one parsed file contributes, plus its own audit data."""
    source: str = ""                   # "cas" | broker key, for messages/flags
    path: str = ""
    fingerprint: str = ""              # sha256[:12] of the file — never-nag
    # (account/folio id, investor-name hint) pairs → the owner-mapping prompt
    accounts: list[tuple[str, str]] = field(default_factory=list)
    sip_txns: list[ImportedSipTxn] = field(default_factory=list)
    trades: list[ImportedTrade] = field(default_factory=list)
    holdings: list[ImportedHolding] = field(default_factory=list)
    # statement-declared closing unit balance per (folio, isin) — the
    # reconciliation anchor; a fund without one can still import, but a
    # fund WITH one must sum to it or it is refused whole
    closing_units: dict[tuple[str, str], float] = field(default_factory=dict)
    # (folio, isin) funds the parser REFUSED (mid-history, unreadable
    # balance): the statement can't speak for these, so the merge must not
    # replace or extend the same ISIN from the folios that did parse
    partial: set = field(default_factory=set)
    warnings: list[str] = field(default_factory=list)


# ---- hardened parsing helpers ----------------------------------------------

_NUM_RE = re.compile(r"^\(?-?[0-9][0-9,]*(?:\.[0-9]+)?\)?$")


def parse_inr(text) -> float | None:
    """A money/units number as Indian statements print it, else None.

    Accepts lakh grouping (1,23,456.78), plain floats, a leading minus and
    the accountant's parenthesised negative. Refuses everything else —
    including empty cells, dates and column-shifted text — so misparses
    quarantine instead of becoming figures.
    """
    if text is None:
        return None
    s = str(text).strip().replace("₹", "").replace("Rs.", "").strip()
    if not s or not _NUM_RE.match(s):
        return None
    neg = s.startswith("(") and s.endswith(")")
    s = s.strip("()").replace(",", "")
    try:
        v = float(s)
    except ValueError:
        return None
    v = -v if neg else v
    if abs(v) >= MAX_AMOUNT:
        return None
    return v


_DATE_FORMATS = ("%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d")
# month names resolved by OUR table, never strptime's %b — %b is locale-
# dependent and "Jan" fails outright on a non-English system locale
_MONTHS = {m: i + 1 for i, m in enumerate(
    ("jan", "feb", "mar", "apr", "may", "jun",
     "jul", "aug", "sep", "oct", "nov", "dec"))}
_NAMED_DATE_RE = re.compile(
    r"^(\d{1,2})[-/ ]([A-Za-z]{3,9})[-/ ](\d{2,4})$")


def parse_date_any(text, today: date | None = None) -> date | None:
    """A transaction date in any shape Indian statements use, else None.

    Bounds: 1990-01-01 .. today (a future or ancient 'date' is a misparse,
    not a transaction).
    """
    if text is None:
        return None
    if isinstance(text, datetime):
        d = text.date()
    elif isinstance(text, date):
        d = text
    else:
        s = str(text).strip()
        d = None
        m = _NAMED_DATE_RE.match(s)
        if m:
            month = _MONTHS.get(m.group(2)[:3].casefold())
            year = int(m.group(3))
            if year < 100:
                year += 2000 if year < 70 else 1900
            if month:
                try:
                    d = date(year, month, int(m.group(1)))
                except ValueError:
                    return None
        if d is None:
            for fmt in _DATE_FORMATS:
                try:
                    d = datetime.strptime(s, fmt).date()
                    break
                except ValueError:
                    continue
        if d is None:
            return None
    if d < EARLIEST_DATE or (today is not None and d > today):
        return None
    return d


def triangle_ok(amount: float | None, units: float | None,
                nav: float | None) -> bool:
    """The amount = units × NAV self-check every statement row must pass.

    A shifted column or mangled comma almost never satisfies the identity,
    so this single gate catches most parse corruption. Rows missing one of
    the three legs can't be proven — the caller treats them as unproven,
    not as passed.
    """
    if amount is None or units is None or nav is None or nav <= 0:
        return False
    return abs(amount - units * nav) <= max(1.0, abs(amount) * 0.001)
