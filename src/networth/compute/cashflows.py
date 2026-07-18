"""Cashflow assembly per asset class (SPEC §6.2) and XIRR aggregation.

Rows missing a required input are silently skipped — same behaviour as the
legacy PowerShell Update-PortfolioXirr. All results may be None (blank cell).
"""

from __future__ import annotations

from datetime import date

from ..model import ClassXirr, PortfolioData, enabled_classes
from .xirr import xirr

Flow = tuple[date, float]


def _yearfrac(a: date, b: date) -> float:
    # clamped like its siblings (projections._yf, snapshot._yf): a reversed
    # pair must never DISCOUNT a balance (v1.6.2)
    return max(0.0, (b - a).days / 365.0)


def flat_accrual(balance: float, rate_pct: float, as_on: date,
                 to: date) -> float:
    """The PPF/EPF flat path (SPEC §6.2/§6.10): a passbook balance grown at
    its annual rate from its as-on date. THE single definition — cashflows,
    projections and the net-worth snapshot all use it, so the three surfaces
    can never drift apart."""
    return balance * (1 + rate_pct / 100) ** _yearfrac(as_on, to)


def equity_flows(data: PortfolioData, today: date) -> list[Flow]:
    flows: list[Flow] = []
    for r in data.equity:
        if not (r.qty and r.avg_cost and r.close and r.cost_date):
            continue
        if r.cost_date >= today:
            continue
        # cost_factor = demerger cost retention (§6.15); the spun-off share
        # of the cost lives on the appended child row. 0.0 is legitimate
        # (parent retains nothing) — only blank means "no demerger".
        cf = r.cost_factor if r.cost_factor is not None else 1.0
        flows.append((r.cost_date, -r.qty * r.avg_cost * cf))
        flows.append((today, r.qty * (r.ca_factor or 1.0) * r.close))
    # v1.6: dividends and recorded sales are real cash — they belong in the
    # return. Appended AFTER the per-row pairs (tests pin flows[0]/[1] as the
    # first row's outflow/inflow). Dividend rows are estimates (amber, §6.12);
    # a future ex-date the feed already announced must not count yet.
    for d in data.dividends:
        if d.ex_date and d.rate and d.qty and d.ex_date <= today:
            flows.append((d.ex_date, d.rate * d.qty))
    for s in data.equity_sells:
        # a typed 0 price is real data (write-off sale, bonus-share cost) —
        # only None means "not entered", so the guards test is-not-None
        if not (s.qty and s.buy_date and s.sell_date) or s.sell_price is None:
            continue
        if (s.qty < 0 or s.sell_price < 0
                or (s.buy_price is not None and s.buy_price < 0)):
            continue
        # same-day (intraday) round trips stay out DELIBERATELY: they are
        # speculative business income (shown in the tax view's Intraday
        # bucket, §6.16), not investment return — and a zero-duration pair
        # is a degenerate flow for a money-weighted rate anyway
        if s.sell_date > today or s.sell_date <= s.buy_date:
            continue
        # blank buy price = pre-Feb-2018 grandfathering path (§6.16): the
        # buy leg is unknown here, so only complete round trips enter XIRR
        if s.buy_price is not None:
            flows.append((s.buy_date, -s.qty * s.buy_price))
            flows.append((s.sell_date, s.qty * s.sell_price))
    return flows


def _sip_units(row) -> float | None:
    if row.units_override is not None:
        return row.units_override
    if row.amount and row.nav and row.nav > 0:  # a negative NAV is noise
        return row.amount / row.nav             # (same guard as _mf_units)
    return None


def mf_flows_by_fund(data: PortfolioData, today: date,
                     nav_by_key: dict[tuple[str, str], float]
                     ) -> dict[tuple[str, str], list[Flow]]:
    """Flows per (owner, scheme). nav_by_key gives the current NAV per fund."""
    funds: dict[tuple[str, str], list[Flow]] = {}
    units: dict[tuple[str, str], float] = {}
    for r in data.sip:
        if not (r.owner and r.scheme and r.txn_date and r.amount):
            continue
        if r.txn_date > today:           # v1.6.2: a pre-typed future row
            continue                     # takes effect when its date arrives
        u = _sip_units(r)
        if not u:
            # v1.6.2: a row whose units can't be known (no NAV, or a typed
            # 0) is left out ENTIRELY — counting only its cash side swung
            # XIRR wildly (buys negative, redemptions positive). run()
            # warns about every such row so it is never silent.
            continue
        key = (r.owner, r.scheme)
        funds.setdefault(key, []).append((r.txn_date, -r.amount))
        units[key] = units.get(key, 0.0) + u    # signed: redemptions reduce
    for key, flows in funds.items():
        nav = nav_by_key.get(key)
        if nav and units.get(key):
            flows.append((today, units[key] * nav))
    return funds


def fd_flows(data: PortfolioData, today: date) -> list[Flow]:
    flows: list[Flow] = []
    for r in data.fixed_deposits:
        if not (r.principal and r.rate and r.start and r.maturity and r.comp_per_year):
            continue
        if r.start >= today:
            continue
        asof = min(today, r.maturity)
        n = r.comp_per_year
        value = r.principal * (1 + (r.rate / 100) / n) ** (n * _yearfrac(r.start, asof))
        flows.append((r.start, -r.principal))
        flows.append((asof, value))
    return flows


def ppf_ledger_by_account(data: PortfolioData, today: date | None = None
                          ) -> dict[tuple[str, str], list[Flow]]:
    """Group PPF deposits by (owner, account_no). With `today` (v1.6.2,
    pass it — every real caller does), future-dated deposits are excluded
    HERE so all three surfaces (class XIRR, the sheet's per-row figures,
    projections) see the identical ledger; a pre-typed deposit takes
    effect when its date arrives. An all-future ledger reads as empty and
    the account falls back to its typed-balance path."""
    from collections import defaultdict
    led: dict[tuple[str, str], list[Flow]] = defaultdict(list)
    for lr in data.ppf_ledger:
        if lr.owner and lr.account_no and lr.txn_date and lr.amount:
            if today is not None and lr.txn_date > today:
                continue
            led[(lr.owner, lr.account_no)].append((lr.txn_date, lr.amount))
    return led


def ppf_flows(data: PortfolioData, today: date) -> list[Flow]:
    """PPF class cashflows: exact ledger flows where a ledger exists, else the
    flat balance-grows-at-rate estimate (SPEC §6.2/§6.10)."""
    from .ppf import load_ppf_rates, ppf_value

    led = ppf_ledger_by_account(data, today)
    rates = load_ppf_rates() if led else []
    flows: list[Flow] = []
    for r in data.ppf:
        deposits = led.get((r.owner, r.account_no))
        if deposits:
            bal, _ = ppf_value(deposits, rates, today)
            flows.extend((d, -a) for d, a in deposits)
            if bal:
                flows.append((today, bal))
            continue
        if not (r.balance and r.rate and r.as_on) or r.as_on >= today:
            continue
        flows.append((r.as_on, -r.balance))
        flows.append((today, flat_accrual(r.balance, r.rate, r.as_on, today)))
    return flows


def epf_flows(data: PortfolioData, today: date) -> list[Flow]:
    """EPF class cashflows (SPEC §6.2, v1.3): the PPF flat path verbatim —
    passbook balance at as-on, accrued at the row's rate to today."""
    flows: list[Flow] = []
    for r in data.epf:
        if not (r.balance and r.rate and r.as_on) or r.as_on >= today:
            continue
        flows.append((r.as_on, -r.balance))
        flows.append((today, flat_accrual(r.balance, r.rate, r.as_on, today)))
    return flows


def manual_asset_flows(data: PortfolioData, today: date,
                       label: str) -> list[Flow]:
    """Two-flow per hand-valued row (SPEC §6.2, v1.3): −Invested at cost
    date, +Current value today. Rows without all three inputs are skipped;
    Cash never reaches here (has_xirr = false)."""
    flows: list[Flow] = []
    for r in data.manual_assets:
        if r.asset_class != label:
            continue
        if not (r.invested and r.cost_date and r.value):
            continue
        if r.cost_date >= today:
            continue
        flows.append((r.cost_date, -r.invested))
        flows.append((today, r.value))
    return flows


def bullion_value(r) -> float | None:
    """Cur. val of one Gold_Silver row — mirrors the sheet's K column:
    qty × purity × (override wins over auto)."""
    rate = r.rate_override if r.rate_override is not None else r.rate_auto
    if not (r.qty and rate):
        return None
    return r.qty * (r.purity or 1.0) * rate


def bullion_flows(data: PortfolioData, today: date) -> list[Flow]:
    """Gold/Silver class cashflows (SPEC §6.2, v1.3): −qty·BuyPrice at Buy
    Date, +Cur. val today; SGB rows additionally earn the statutory 2.5% p.a.
    semi-annual coupon on the row's Buy Price (documented approximation —
    the statutory base is issue price)."""
    flows: list[Flow] = []
    for r in data.bullion:
        value = bullion_value(r)
        if not (r.qty and r.buy_price and r.buy_date and value):
            continue
        if r.buy_date >= today:
            continue
        flows.append((r.buy_date, -r.qty * r.buy_price))
        if r.metal_type == "SGB" and r.maturity:
            amount = r.qty * r.buy_price * 0.0125       # 2.5%/2 per half-year
            flows.extend((c, amount) for c in
                         coupon_dates(r.maturity, r.buy_date, today, freq=2))
        flows.append((today, value))
    return flows


def nps_flows(data: PortfolioData, today: date) -> list[Flow]:
    """NPS class cashflows (SPEC §6.2, v1.3) — approximate two-flow per row:
    −Total contributed at First contribution, +units·NAV today. Rows missing
    either optional input are skipped (their value still counts elsewhere)."""
    flows: list[Flow] = []
    for r in data.nps:
        if not (r.total_contributed and r.first_contribution
                and r.units and r.current_nav):
            continue
        if r.first_contribution >= today:
            continue
        flows.append((r.first_contribution, -r.total_contributed))
        flows.append((today, r.units * r.current_nav))
    return flows


def coupon_dates(maturity: date, after: date, before: date,
                 freq: int = 1) -> list[date]:
    """Coupon dates in (after, before], stepping back 12/freq months from
    maturity (SPEC §6.3)."""
    step_months = 12 // freq
    dates = []
    y, m, d = maturity.year, maturity.month, maturity.day
    while True:
        try:
            c = date(y, m, d)
        except ValueError:              # e.g. 29 Feb stepping into a non-leap year
            c = date(y, m, 28)
        if c <= after:
            break
        if c <= before:
            dates.append(c)
        m -= step_months
        while m < 1:
            m += 12
            y -= 1
    dates.sort()
    return dates


def bond_coupon_flows(r, today: date) -> list[Flow]:
    """Historical coupons (buy_date, today] — feed XIRR (SPEC §6.3)."""
    if not (r.qty and r.face and r.coupon and r.maturity and r.buy_date):
        return []
    amount = r.qty * r.face * (r.coupon / 100)
    return [(c, amount) for c in coupon_dates(r.maturity, r.buy_date, today)]


def bond_flows(data: PortfolioData, today: date) -> list[Flow]:
    flows: list[Flow] = []
    for r in data.bonds:
        if not (r.qty and r.buy_price and r.cur_price and r.buy_date):
            continue
        if r.buy_date >= today:
            continue
        flows.append((r.buy_date, -r.qty * r.buy_price))
        flows.extend(bond_coupon_flows(r, today))
        flows.append((today, r.qty * r.cur_price))
    return flows


def compute_all_xirr(data: PortfolioData, today: date | None = None) -> ClassXirr:
    """Class + portfolio XIRR, plus per-fund XIRR written into data.mutual_funds."""
    today = today or date.today()

    nav_by_key = {
        (m.owner, m.scheme): m.current_nav
        for m in data.mutual_funds if m.current_nav
    }
    per_fund = mf_flows_by_fund(data, today, nav_by_key)
    for m in data.mutual_funds:
        flows = per_fund.get((m.owner, m.scheme))
        m.xirr = xirr(flows) if flows else None

    eq = equity_flows(data, today)
    mf = [f for flows in per_fund.values() for f in flows]
    fd = fd_flows(data, today)
    ppf = ppf_flows(data, today)
    epf = epf_flows(data, today)
    bonds = bond_flows(data, today)
    gs = bullion_flows(data, today)
    nps = nps_flows(data, today)
    re_ = manual_asset_flows(data, today, "Property")
    ins = manual_asset_flows(data, today, "Insurance")
    oth = manual_asset_flows(data, today, "Other")
    # Cash is excluded from XIRR entirely (has_xirr = false, SPEC §2.1)

    for r in data.nps:                      # per-row approximate XIRR
        if (r.total_contributed and r.first_contribution
                and r.units and r.current_nav and r.first_contribution < today):
            r.xirr = xirr([(r.first_contribution, -r.total_contributed),
                           (today, r.units * r.current_nav)])
        else:
            r.xirr = None

    # the family/portfolio XIRR combines only the classes that are counted —
    # a switched-off class is excluded everywhere (SPEC §3.14, v1.4.3); its
    # per-class figure is still computed (harmless, and ready when re-enabled)
    flows_by_key = {"equity": eq, "mutual_funds": mf, "fixed_deposits": fd,
                    "ppf": ppf, "epf": epf, "bonds": bonds, "gold_silver": gs,
                    "nps": nps, "real_estate": re_, "insurance": ins,
                    "other_assets": oth}
    on = {c.key for c in enabled_classes(data)}
    return ClassXirr(
        portfolio=xirr([f for k, fl in flows_by_key.items() if k in on
                        for f in fl]),
        equity=xirr(eq),
        mutual_funds=xirr(mf),
        fixed_deposits=xirr(fd),
        ppf=xirr(ppf),
        epf=xirr(epf),
        bonds=xirr(bonds),
        gold_silver=xirr(gs),
        nps=xirr(nps),
        real_estate=xirr(re_),
        insurance=xirr(ins),
        other_assets=xirr(oth),
    )
