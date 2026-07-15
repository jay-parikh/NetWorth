"""PPF interest engine — hand-checked against the official rule (SPEC §6.10)."""

from datetime import date

import pytest

from networth.compute.ppf import (
    current_rate, load_ppf_rates, ppf_value, rate_on,
)

FLAT = [(date(2000, 1, 1), 7.1)]        # a single-rate table for clean checks


def test_full_year_single_deposit_earns_the_annual_rate():
    # 100000 deposited on/before the 5th, held a full FY at 7.1% → exactly 7100
    bal, interest = ppf_value([(date(2020, 4, 1), 100000)], FLAT, date(2021, 3, 31))
    assert interest == pytest.approx(7100.0, abs=0.01)
    assert bal == pytest.approx(107100.0, abs=0.01)


def test_deposit_after_the_5th_loses_that_month():
    # deposit on the 6th → April earns nothing, only 11 months count
    bal, interest = ppf_value([(date(2020, 4, 6), 100000)], FLAT, date(2021, 3, 31))
    assert interest == pytest.approx(100000 * 0.071 / 12 * 11, abs=0.01)
    assert bal == pytest.approx(100000 + interest, abs=0.01)


def test_annual_compounding_across_two_years():
    # year 1 credits 7100 at 31-Mar-2021; year 2 earns 7.1% on 107100
    bal, interest = ppf_value([(date(2020, 4, 1), 100000)], FLAT, date(2022, 3, 31))
    y2 = 107100 * 0.071
    assert bal == pytest.approx(107100 + y2, abs=0.02)
    assert interest == pytest.approx(7100 + y2, abs=0.02)


def test_accrued_but_not_yet_credited_shows_in_balance():
    # 3 completed months (Apr,May,Jun) accrued, not yet credited (before Mar)
    bal, interest = ppf_value([(date(2020, 4, 1), 120000)], FLAT, date(2020, 6, 30))
    monthly = 120000 * 0.071 / 12
    assert interest == pytest.approx(3 * monthly, abs=0.01)
    assert bal == pytest.approx(120000 + 3 * monthly, abs=0.01)


def test_no_deposits_or_future_deposits():
    assert ppf_value([], FLAT, date(2024, 1, 1)) == (0.0, 0.0)
    assert ppf_value([(date(2025, 1, 1), 5000)], FLAT, date(2024, 1, 1)) == (0.0, 0.0)


def test_rate_step_table():
    rates = load_ppf_rates()
    assert rate_on(rates, date(2016, 5, 1)) == 8.1
    assert rate_on(rates, date(2016, 10, 1)) == 8.0
    assert rate_on(rates, date(2018, 2, 1)) == 7.6
    assert rate_on(rates, date(2020, 4, 1)) == 7.1
    assert rate_on(rates, date(2026, 7, 1)) == 7.1        # held since Apr-2020
    assert rate_on(rates, date(2010, 1, 1)) == 8.0        # between 2003 and 2011
    assert current_rate(rates) == 7.1


def test_real_rate_history_blended():
    # a 2019 deposit spans the 7.9% (Jul-2019) and 7.1% (Apr-2020) regimes;
    # just assert it grows and interest is positive & sane
    bal, interest = ppf_value([(date(2019, 4, 1), 150000)],
                              load_ppf_rates(), date(2021, 3, 31))
    assert 150000 < bal < 150000 * 1.20
    assert interest > 0


# ---- workbook integration (optional-ledger + fallback) ----

def _wb_tools():
    from networth.generate import build_workbook
    from networth.reader import read_workbook
    from networth.sample_data import sample_portfolio
    from networth.fetch.bhavcopy import PriceData
    from networth.fetch.amfi import AmfiData
    from networth.update import run
    return build_workbook, read_workbook, sample_portfolio, PriceData, AmfiData, run


def test_ledger_account_computed_nonledger_fallback(tmp_path):
    build_workbook, read_workbook, sample_portfolio, PriceData, AmfiData, run = _wb_tools()
    path = tmp_path / "wb.xlsx"
    build_workbook(sample_portfolio(), str(path))
    s = run(path, price_data=PriceData(trade_date=date(2026, 7, 15), source="T"),
            amfi_data=AmfiData(), ca_data=[], today=date(2026, 7, 15))
    assert s["ppf_ledgered"] == 1
    back = read_workbook(str(path))
    amit = next(r for r in back.ppf if r.owner == "Amit")
    priya = next(r for r in back.ppf if r.owner == "Priya")
    # ledgered: 5x150000 deposits → principal 750k plus interest
    assert amit.balance_today > 750000 and amit.interest_earned > 100000
    assert 0.06 < amit.xirr < 0.08
    assert amit.rate == 7.1                       # blank rate auto-filled
    # non-ledger: computed balance not written (sheet formula = Current Balance)
    assert priya.balance_today is None


def test_ppf_balance_today_formula_and_dashboard(tmp_path):
    from openpyxl import load_workbook
    build_workbook, *_ , sample_portfolio, _P, _A, _run = _wb_tools()
    path = tmp_path / "wb.xlsx"
    build_workbook(sample_portfolio(), str(path))
    wb = load_workbook(path)
    ppf = wb["PPF"]
    assert ppf["H3"].value == "Balance today"
    # before any update, every account's Balance today = its Current Balance
    assert ppf["H4"].value == '=IF($D4="","",$D4)'
    # Dashboard + person sheets total the computed Balance-today column (H)
    assert "SUMIFS(PPF!$H:$H" in wb["Dashboard"]["E6"].value
    assert "SUMIFS(PPF!$H:$H" in wb["Amit"]["B9"].value
    # ledger sheet exists with the right headers
    assert wb["PPF_Ledger"]["A3"].value == "Owner"
    assert wb["PPF_Ledger"]["D3"].value == "Amount"


def test_ppf_ledger_roundtrips(tmp_path):
    build_workbook, read_workbook, sample_portfolio, *_ = _wb_tools()
    path = tmp_path / "wb.xlsx"
    data = sample_portfolio()
    build_workbook(data, str(path))
    back = read_workbook(str(path))
    assert len(back.ppf_ledger) == len(data.ppf_ledger) == 5
    assert back.ppf_ledger[0].account_no == "PPF-778101"
    assert back.ppf_ledger[0].amount == 150000
