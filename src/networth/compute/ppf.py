"""PPF interest engine (SPEC §6.10) — the official monthly-minimum-balance rule.

Public Provident Fund interest is calculated on the **lowest balance between
the close of the 5th day and the last day of each month**, and **credited once
a year on 31 March**. A deposit made on or before the 5th earns interest that
month; a later deposit does not. Rates are set (quarterly since Apr-2016,
annually before) by the Ministry of Finance — bundled in data/ppf_rates.csv,
refreshed via app releases (there is no API).

No withdrawals are modelled (PPF is accumulation-only here), so within a month
the minimum balance is simply the balance as of the 5th.
"""

from __future__ import annotations

import csv
from calendar import monthrange
from datetime import date
from pathlib import Path

from ..model import DATA_DIR


def load_ppf_rates(data_dir: Path = DATA_DIR) -> list[tuple[date, float]]:
    """Bundled rate history as an ascending [(from_date, rate_pct)] step table."""
    rows: list[tuple[date, float]] = []
    with open(data_dir / "ppf_rates.csv", newline="", encoding="utf-8") as f:
        rdr = csv.DictReader(f)
        for r in rdr:
            y, m, d = (int(x) for x in r["from_date"].split("-"))
            rows.append((date(y, m, d), float(r["rate_pct"])))
    rows.sort()
    return rows


def rate_on(rates: list[tuple[date, float]], d: date) -> float:
    """Annual rate % in effect on date d (step function)."""
    pct = rates[0][1]
    for from_date, r in rates:
        if from_date <= d:
            pct = r
        else:
            break
    return pct


def current_rate(rates: list[tuple[date, float]] | None = None) -> float:
    rates = rates or load_ppf_rates()
    return rates[-1][1]


def _next_month(y: int, m: int) -> tuple[int, int]:
    return (y + 1, 1) if m == 12 else (y, m + 1)


def ppf_value(deposits: list[tuple[date, float]],
              rates: list[tuple[date, float]],
              as_of: date) -> tuple[float, float]:
    """Return (balance_as_of, total_interest) for a PPF account.

    balance_as_of includes interest credited at past 31-March year-ends plus
    interest accrued (not yet credited) since the last 31 March. Deposits are
    (date, amount); amounts are contributions (no withdrawals).
    """
    deposits = sorted((d, a) for d, a in deposits if a)
    if not deposits or deposits[0][0] > as_of:
        return 0.0, 0.0

    credited = 0.0        # balance of deposits + interest credited at FY ends
    accrued_fy = 0.0      # interest accrued this FY, not yet credited
    total_interest = 0.0
    di, n = 0, len(deposits)
    y, m = deposits[0][0].year, deposits[0][0].month

    while (y, m) <= (as_of.year, as_of.month):
        on_before_5 = after_5 = 0.0
        while di < n and (deposits[di][0].year, deposits[di][0].month) == (y, m):
            d0, amt = deposits[di]
            if d0.day <= 5:
                on_before_5 += amt
            else:
                after_5 += amt
            di += 1

        month_end = date(y, m, monthrange(y, m)[1])
        if month_end <= as_of:
            # completed month: interest on the 5th-to-end minimum balance
            min_balance = credited + on_before_5
            interest = min_balance * rate_on(rates, date(y, m, 15)) / 1200.0
            accrued_fy += interest
            total_interest += interest
            credited += on_before_5 + after_5
            if m == 3:                        # 31 March → credit the year
                credited += accrued_fy
                accrued_fy = 0.0
        else:
            # in-progress month: deposits are in the account, no interest yet
            credited += on_before_5 + after_5

        y, m = _next_month(y, m)

    return credited + accrued_fy, total_interest


def ppf_cashflows(deposits: list[tuple[date, float]], balance: float,
                  as_of: date) -> list[tuple[date, float]]:
    """XIRR cashflows for a ledgered PPF account: deposits out, balance in."""
    flows = [(d, -a) for d, a in deposits if a]
    if flows and balance:
        flows.append((as_of, balance))
    return flows
