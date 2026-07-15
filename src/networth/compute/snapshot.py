"""Net-worth snapshot (SPEC §6.11) — per-class current value at a point in time.

Mirrors the Dashboard's per-class totals in Python so the updater can record a
dated History row. FD interest uses actual/365 (consistent with the XIRR
engine), a hair different from Excel's 30/360 YEARFRAC — immaterial for a trend.
"""

from __future__ import annotations

from datetime import date

from ..model import HistorySnapshot, PortfolioData


def _yf(a: date, b: date) -> float:
    return max(0.0, (b - a).days / 365.0)


def _mf_units(data: PortfolioData) -> dict[tuple[str, str], float]:
    units: dict[tuple[str, str], float] = {}
    for s in data.sip:
        if not (s.owner and s.scheme):
            continue
        u = s.units_override if s.units_override is not None else (
            s.amount / s.nav if s.amount and s.nav else None)
        if u is not None:
            units[(s.owner, s.scheme)] = units.get((s.owner, s.scheme), 0.0) + u
    return units


def net_worth_snapshot(data: PortfolioData, today: date) -> HistorySnapshot:
    equity = sum(r.qty * (r.ca_factor or 1.0) * r.close
                 for r in data.equity if r.qty and r.close)

    units = _mf_units(data)
    mutual_funds = sum(units.get((m.owner, m.scheme), 0.0) * m.current_nav
                       for m in data.mutual_funds if m.current_nav)

    fixed_deposits = 0.0
    for r in data.fixed_deposits:
        if r.principal and r.rate and r.start and r.maturity and r.comp_per_year:
            asof = min(today, r.maturity)
            n = r.comp_per_year
            fixed_deposits += r.principal * (1 + (r.rate / 100) / n) ** (
                n * _yf(r.start, asof))

    ppf = sum((r.balance_today if r.balance_today is not None else (r.balance or 0.0))
              for r in data.ppf)

    bonds = sum(r.qty * r.cur_price for r in data.bonds if r.qty and r.cur_price)

    return HistorySnapshot(snap_date=today, equity=round(equity, 2),
                           mutual_funds=round(mutual_funds, 2),
                           fixed_deposits=round(fixed_deposits, 2),
                           ppf=round(ppf, 2), bonds=round(bonds, 2))


def upsert_snapshot(history: list[HistorySnapshot], snap: HistorySnapshot,
                    keep: int) -> list[HistorySnapshot]:
    """One row per day: replace any existing row for snap's date, keep sorted,
    and cap to the most recent `keep` rows."""
    out = [h for h in history if h.snap_date and h.snap_date != snap.snap_date]
    out.append(snap)
    out.sort(key=lambda h: h.snap_date)
    return out[-keep:]
