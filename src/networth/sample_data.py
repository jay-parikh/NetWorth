"""Fictional sample portfolio shipped with the released template.

Three people (Amit, Priya, Rahul) with real ISINs/scheme names so the first
updater run works end-to-end (SPEC §4). Prices/NAVs/XIRR figures are the
updater-written values captured on 2026-07-13/14 — the user's first update
overwrites them. Onboarding = replace the sample rows with your own; classes
you don't own stay switched off in Settings (their samples wait inside the
hidden tabs as worked examples).
"""

from __future__ import annotations

from datetime import date

from .model import (
    ASSET_CLASSES, BondRow, BullionRow, ClassSetting, ClassXirr, DividendRow,
    EPFRow, EquityRow, EquitySellRow, FDRow, ManualAssetRow, MFRow, NPSRow,
    PPFLedgerRow, PPFRow, PortfolioData, SIPRow, ScripRef, load_masters,
    load_tax_rules,
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
            MFRow(owner="Amit", scheme=_PPFCF, current_nav=91.2, xirr=0.36695,
                  tax_type="Equity"),
            MFRow(owner="Priya", scheme=_SBI_LC, current_nav=96.4,
                  xirr=0.073869, tax_type="Equity"),
        ],
        # worked sale examples (v1.6 §3.20): one normal round trip, one
        # pre-2018 purchase with the buy price left blank on purpose — the
        # Capital Gains tab shows the 31-Jan-2018 grandfathering rule on it.
        # (The Equity rows above are what's owned NOW, after these sales.)
        equity_sells=[
            EquitySellRow("Amit", "INFOSYS LTD.", "", 10,
                          date(2024, 1, 10), 1400.0,
                          date(2026, 5, 20), 1650.0,
                          "Sold to rebalance"),
            EquitySellRow("Priya", "STATE BANK OF INDIA", "", 50,
                          date(2016, 6, 1), None,
                          date(2026, 4, 15), 820.0,
                          "Old purchase - buy price unknown"),
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
        # every class ships a couple of rows that SHOW how it works. Classes
        # whose Settings toggle is No start hidden (v1.4.3: the toggle wins),
        # so these samples are the worked example a user finds waiting the
        # moment they switch a class on.
        epf=[
            EPFRow("Amit", "AcmeCorp / UAN 100200300400", "MH/BAN/12345/678",
                   1500000, date(2026, 3, 31), 8.25,
                   "From the EPFO passbook"),
        ],
        bullion=[
            BullionRow("Amit", "SGB", "SGB Jun-2028 (2020-21 Ser III)",
                       "IN0020200104", 20, None, 4889, date(2020, 6, 30),
                       rate_auto=14093.7, maturity=date(2028, 6, 30)),
            # generic coin/bar/jewellery examples: weigh in grams, pick the
            # metal, purity blank for 24K/999 (22K jewellery = 0.916)
            BullionRow("Priya", "Gold", "Gold coins, 2 x 10 g (24K)", "",
                       20, None, 6450, date(2023, 10, 12), rate_auto=14167.9),
            BullionRow("Priya", "Gold", "Jewellery, 40 g (22K)", "", 40,
                       0.916, 4800, date(2021, 11, 4), rate_auto=14167.9),
            BullionRow("Rahul", "Silver", "Silver bar, 1 kg", "", 1000,
                       None, 72, date(2022, 3, 10), rate_auto=217.43),
        ],
        nps=[
            NPSRow("Amit", "110012345678",
                   "SBI PENSION FUND SCHEME E - TIER I",
                   units=1200, current_nav=56.7834, total_contributed=40000,
                   first_contribution=date(2019, 6, 1), xirr=0.0776),
        ],
        manual_assets=[
            # deliberately generic examples — what you paid (optional) and
            # what it's worth today are the only numbers that matter
            ManualAssetRow("Amit", "Property", "Apartment (self-occupied)",
                           "Purchase deed", 4500000,
                           date(2016, 7, 15), 9000000, date(2026, 7, 1)),
            ManualAssetRow("Priya", "Cash", "Savings account balance",
                           "HDFC Bank", None, None, 250000, date(2026, 7, 10)),
            ManualAssetRow("Rahul", "Insurance",
                           "Life policy - surrender value today", "LIC",
                           380000, date(2015, 5, 1), 520000,
                           date(2026, 6, 1)),
            # "Other" is the hand-converted ₹ home for anything without a
            # feed — e.g. foreign RSUs until multi-currency lands (ROADMAP)
            ManualAssetRow("Priya", "Other", "Employee shares (US), in ₹",
                           "Employer plan", 240000, date(2024, 1, 10), 300000,
                           date(2026, 7, 1)),
        ],
        bullion_rate_asof=date(2026, 7, 16),
        # first open shows only the classic five (registry defaults); the
        # other classes ship hidden WITH their sample rows inside, so
        # switching one on in Settings reveals a worked example instantly.
        # Targets on the visible classes light up the drift view (sum 100).
        class_settings={
            c.key: ClassSetting(enabled=c.default_enabled, target_pct={
                "equity": 40, "mutual_funds": 15, "fixed_deposits": 20,
                "ppf": 15, "bonds": 10,
            }.get(c.key))
            for c in ASSET_CLASSES
        },
        by_scrip=[ScripRef(isin=i, name=name_by_isin.get(i, "")) for i in _BY_SCRIP_ISINS],
        dividends=[
            # a frozen prior-FY record + a current-FY row the updater refreshes
            DividendRow(fy="2025-26", owner="Priya", scrip="ITC LTD.",
                        isin="INE154A01025", div_type="Final",
                        ex_date=date(2025, 6, 4), rate=7.85, qty=200,
                        source="Auto", details="Final Dividend - Rs 7.85 Per Share"),
            DividendRow(fy="2026-27", owner="Amit",
                        scrip="TATA CONSULTANCY SERVICES LTD.",
                        isin="INE467B01029", div_type="Interim",
                        ex_date=date(2026, 7, 2), rate=10, qty=12,
                        source="Auto", details="Interim Dividend - Rs 10 Per Share"),
        ],
        inflation_pct=7,
        # class XIRRs live in the Dashboard allocation table, which lists
        # only the SHOWN classes — hidden classes carry no stored figure
        # (the updater computes theirs the moment they're switched on)
        # portfolio/equity re-captured 2026-07-18 under the v1.6 flow
        # semantics (dividends + recorded sells count in the return)
        xirr=ClassXirr(
            portfolio=0.0691672176,
            equity=0.0666771890,
            mutual_funds=0.0876706058,
            fixed_deposits=0.0721401469,
            ppf=0.071,
            bonds=0.0059830784,
        ),
        masters=masters,
        # the Tax_Rules sheet starts as the bundled law verbatim (§3.22) —
        # from then on the workbook's rows are the source of truth
        tax_rules=load_tax_rules(),
    )
