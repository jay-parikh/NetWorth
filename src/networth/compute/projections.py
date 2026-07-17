"""Expected value at FY end (SPEC §6.8) — per person and total.

Fixed income is deterministic accrual; market assets grow at the user's
"Expected return %" Dashboard input. An estimate by construction — the
Dashboard column header says so.
"""

from __future__ import annotations

from datetime import date

from ..model import PortfolioData, enabled_classes
from .cashflows import coupon_dates, flat_accrual


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
    # a switched-off class is not counted anywhere (SPEC §3.14, v1.4.3)
    on = {c.key for c in enabled_classes(data)}

    def add(owner: str, value: float) -> None:
        if owner in out:
            out[owner] += value

    if "equity" in on:
        for r in data.equity:
            if r.qty and r.close:
                add(r.owner, r.qty * (r.ca_factor or 1.0) * r.close * growth)

    if "mutual_funds" in on:
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

    if "fixed_deposits" in on:
        for r in data.fixed_deposits:
            if not (r.principal and r.rate and r.start and r.maturity
                    and r.comp_per_year):
                continue
            asof = min(end, r.maturity)
            n = r.comp_per_year
            add(r.owner,
                r.principal * (1 + (r.rate / 100) / n) ** (n * _yf(r.start, asof)))

    if "ppf" in on:
        from .cashflows import ppf_ledger_by_account
        from .ppf import load_ppf_rates, ppf_value
        led = ppf_ledger_by_account(data)
        rates = load_ppf_rates() if led else []
        for r in data.ppf:
            deposits = led.get((r.owner, r.account_no))
            if deposits:                    # exact ledger accrual to FY end
                bal, _ = ppf_value(deposits, rates, end)
                add(r.owner, bal)
            elif r.balance and r.rate and r.as_on:
                add(r.owner, flat_accrual(r.balance, r.rate, r.as_on, end))

    if "bonds" in on:
        for r in data.bonds:
            if not (r.qty and (r.cur_price or r.face)):
                continue
            if r.maturity and r.face and r.maturity <= end:
                value = r.qty * r.face               # redeemed within the FY
            else:
                value = r.qty * (r.cur_price or r.face)
            if r.face and r.coupon and r.maturity:
                value += r.qty * r.face * (r.coupon / 100) * len(
                    coupon_dates(r.maturity, today, end))
            add(r.owner, value)

    # market-linked new classes grow at the Expected-return input, like Equity
    if "gold_silver" in on:
        from .cashflows import bullion_value
        for r in data.bullion:
            v = bullion_value(r)
            if v:
                add(r.owner, v * growth)
    if "nps" in on:
        for r in data.nps:
            if r.units and r.current_nav:
                add(r.owner, r.units * r.current_nav * growth)

    if "epf" in on:
        for r in data.epf:                # accrues at its own rate to FY end
            if not r.balance:
                continue
            if r.as_on and r.rate:
                add(r.owner, flat_accrual(r.balance, r.rate, r.as_on, end))
            else:
                add(r.owner, r.balance)

    # hand-valued assets are held FLAT to FY end — estimating property or
    # surrender-value appreciation would be false precision (SPEC §6.8)
    key_by_manual = {"Property": "real_estate", "Cash": "cash",
                     "Insurance": "insurance", "Other": "other_assets"}
    for r in data.manual_assets:
        if r.value and key_by_manual.get(r.asset_class, "") in on:
            add(r.owner, r.value)

    return {p: round(v, 2) for p, v in out.items() if v}
