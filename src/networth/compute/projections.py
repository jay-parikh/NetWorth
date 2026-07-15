"""Expected value at FY end (SPEC §6.8) — per person and total.

Fixed income is deterministic accrual; market assets grow at the user's
"Expected return %" Dashboard input. An estimate by construction — the
Dashboard column header says so.
"""

from __future__ import annotations

from datetime import date

from ..model import PortfolioData
from .cashflows import coupon_dates


def fy_end(today: date) -> date:
    """Next 31 March ≥ today."""
    end = date(today.year, 3, 31)
    return end if today <= end else date(today.year + 1, 3, 31)


def _yf(a: date, b: date) -> float:
    return max(0.0, (b - a).days / 365.0)


def fy_expected_by_person(data: PortfolioData, today: date | None = None
                          ) -> dict[str, float]:
    today = today or date.today()
    end = fy_end(today)
    growth = (1 + data.expected_return_pct / 100) ** _yf(today, end)
    out: dict[str, float] = {p: 0.0 for p in data.persons}

    def add(owner: str, value: float) -> None:
        if owner in out:
            out[owner] += value

    for r in data.equity:
        if r.qty and r.close:
            add(r.owner, r.qty * (r.ca_factor or 1.0) * r.close * growth)

    nav_units: dict[tuple[str, str], float] = {}
    for s in data.sip:
        u = s.units_override if s.units_override is not None else (
            s.amount / s.nav if s.amount and s.nav else None)
        if u is not None:
            key = (s.owner, s.scheme)
            nav_units[key] = nav_units.get(key, 0.0) + u
    for m in data.mutual_funds:
        units = nav_units.get((m.owner, m.scheme))
        if units and m.current_nav:
            add(m.owner, units * m.current_nav * growth)

    for r in data.fixed_deposits:
        if not (r.principal and r.rate and r.start and r.maturity and r.comp_per_year):
            continue
        asof = min(end, r.maturity)
        n = r.comp_per_year
        add(r.owner, r.principal * (1 + (r.rate / 100) / n) ** (n * _yf(r.start, asof)))

    for r in data.ppf:
        if r.balance and r.rate and r.as_on:
            add(r.owner, r.balance * (1 + r.rate / 100) ** _yf(r.as_on, end))

    for r in data.bonds:
        if not (r.qty and (r.cur_price or r.face)):
            continue
        if r.maturity and r.face and r.maturity <= end:
            value = r.qty * r.face                       # redeemed within the FY
        else:
            value = r.qty * (r.cur_price or r.face)
        if r.face and r.coupon and r.maturity:
            value += r.qty * r.face * (r.coupon / 100) * len(
                coupon_dates(r.maturity, today, end))
        add(r.owner, value)

    return {p: round(v, 2) for p, v in out.items() if v}
