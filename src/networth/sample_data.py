"""Fictional sample portfolio shipped with the released template.

Three people (Amit, Priya, Rahul) with real ISINs/scheme names so the first
updater run works end-to-end (SPEC §4). Prices/NAVs/XIRR figures are the
updater-written values captured on 2026-07-13/14 — the user's first update
overwrites them. Deleting the sample rows is the onboarding step.
"""

from __future__ import annotations

from datetime import date

from .model import (
    BondRow, ClassXirr, EquityRow, FDRow, MFRow, PPFLedgerRow, PPFRow,
    PortfolioData, SIPRow, ScripRef, load_masters,
)

PERSONS = ["Amit", "Priya", "Rahul"]

_CLOSE_DATE = date(2026, 7, 13)
_FMV_DATE = date(2018, 1, 31)   # grandfathering date, see Guide

_EQUITY = [
    # owner, scrip (Stock_Master name), qty, avg cost, close, prev close
    ("Amit", "RELIANCE INDUSTRIES LTD.", 50, 964.9, 1520, 1512.3),
    ("Amit", "TATA CONSULTANCY SERVICES LTD.", 12, 2916.05, 3450, 3462.1),
    ("Amit", "INFOSYS LTD.", 40, 1173.85, 1610, 1598.7),
    ("Amit", "HDFC BANK LTD.", 60, 1076, 1690, 1701.5),
    ("Priya", "ITC LTD.", 200, 282.6, 435, 432.9),
    ("Priya", "ICICI BANK LTD.", 80, 365, 1240, 1231),
    ("Priya", "WIPRO LTD.", 150, 317, 245, 247.2),
    ("Rahul", "LARSEN & TOUBRO LTD.", 25, 1470, 3620, 3608.4),
    ("Rahul", "STATE BANK OF INDIA", 120, 327, 815, 812.5),
    ("Rahul", "HINDUSTAN UNILEVER LTD.", 30, 1370, 2410, 2422.8),
]

_PPFCF = "Parag Parikh Flexi Cap Fund - Direct Plan - Growth"
_SBI_LC = "SBI Large Cap FUND-DIRECT PLAN -GROWTH"

_BY_SCRIP_ISINS = [
    "INE040A01034", "INE030A01027", "INE090A01021", "INE009A01021",
    "INE154A01025", "INE018A01030", "INE002A01018", "INE062A01020",
    "INE467B01029", "INE075A01022",
]


def sample_portfolio() -> PortfolioData:
    masters = load_masters(mf_refreshed="14-07-2026", stock_refreshed="14-07-2026")
    name_by_isin = {isin: name for _sym, name, isin in masters.stock_rows}

    return PortfolioData(
        persons=list(PERSONS),
        equity=[
            EquityRow(owner=o, scrip=s, qty=q, avg_cost=c, close=cl,
                      prev_close=pv, close_date=_CLOSE_DATE, cost_date=_FMV_DATE)
            for o, s, q, c, cl, pv in _EQUITY
        ],
        mutual_funds=[
            MFRow(owner="Amit", scheme=_PPFCF, current_nav=91.2, xirr=0.36695),
            MFRow(owner="Priya", scheme=_SBI_LC, current_nav=96.4, xirr=0.073869),
        ],
        sip=[
            SIPRow("Amit", _PPFCF, date(2026, 1, 15), 10000, 78.42),
            SIPRow("Amit", _PPFCF, date(2026, 2, 16), 10000, 80.11),
            SIPRow("Amit", _PPFCF, date(2026, 3, 16), 10000, 82.05),
            SIPRow("Amit", _PPFCF, date(2026, 4, 15), 10000, 84.6),
            SIPRow("Amit", _PPFCF, date(2026, 5, 15), 10000, 86.9),
            SIPRow("Amit", _PPFCF, date(2026, 6, 15), 10000, 88.75),
            SIPRow("Priya", _SBI_LC, date(2024, 8, 20), 150000, 84.2),
        ],
        fixed_deposits=[
            FDRow("Amit", "HDFC Bank", "FD-2024-0117", 500000, 7.1,
                  date(2024, 4, 1), date(2027, 4, 1), 4),
            FDRow("Rahul", "State Bank of India", "FD-2025-0342", 300000, 6.8,
                  date(2025, 1, 15), date(2028, 1, 15), 4),
        ],
        ppf=[
            # Priya: no ledger → flat estimate from the current balance
            PPFRow("Priya", "Post Office", "PPF-104522", 860000,
                   date(2026, 3, 31), 7.1, "Matures 2031"),
            # Amit: ledgered (see ppf_ledger) → Balance today computed from deposits;
            # Current Balance left blank so the ledger is the single source of truth
            PPFRow("Amit", "SBI", "PPF-778101", None, None, None,
                   "Balance computed from PPF_Ledger"),
        ],
        ppf_ledger=[
            PPFLedgerRow("Amit", "PPF-778101", date(2021, 4, 5), 150000),
            PPFLedgerRow("Amit", "PPF-778101", date(2022, 4, 4), 150000),
            PPFLedgerRow("Amit", "PPF-778101", date(2023, 4, 5), 150000),
            PPFLedgerRow("Amit", "PPF-778101", date(2024, 4, 5), 150000),
            PPFLedgerRow("Amit", "PPF-778101", date(2025, 4, 4), 150000),
        ],
        bonds=[
            BondRow("Rahul", "8.5% NHAI Bond 2029", "INE906B07CB9", 50, 1000,
                    1000, 1015, 8.5, date(2029, 1, 15), date(2024, 1, 15)),
        ],
        by_scrip=[ScripRef(isin=i, name=name_by_isin.get(i, "")) for i in _BY_SCRIP_ISINS],
        inflation_pct=7,
        xirr=ClassXirr(
            portfolio=0.0676209694,
            equity=0.0664365522,
            mutual_funds=0.0876706058,
            fixed_deposits=0.0721401469,
            ppf=0.071,
            bonds=0.0059830784,
        ),
        masters=masters,
    )
