"""R13: Gold_Silver + NPS (SPEC §3.15/§3.16, §5.6/§5.7, §6.14)."""

from dataclasses import asdict
from datetime import date

import pytest
from openpyxl import load_workbook

from networth.compute.cashflows import bullion_flows, bullion_value
from networth.fetch.amfi import AmfiData
from networth.fetch.bhavcopy import PriceData
from networth.fetch.bullion import derive_from_bhavcopy, parse_ibja
from networth.fetch.nps import NpsData, parse as nps_parse
from networth.generate import build_workbook
from networth.model import BullionRow, ClassSetting, NPSRow
from networth.reader import read_workbook
from networth.sample_data import sample_portfolio
from networth.update import run

TODAY = date(2026, 7, 15)

IBJA_HTML = """
<tr><td>Gold 999</td><td><span id="lblGold999_AM">141713</span></td>
<td><span id="lblGold999_PM">141679</span></td></tr>
<tr><td>Silver 999</td><td><span id="lblSilver999_AM">218485</span></td>
<td><span id="lblSilver999_PM">217434</span></td></tr>
"""

NPS_TSV = ("ID\tDATE OF NAV\tPFM NAME\tSCHEME ID\tSCHEME NAME\tNAV VALUE\n"
           "1\t2026-07-15\tSBI Pension Funds Pvt. Ltd.\tSM001003\t"
           "SBI PENSION FUND SCHEME E - TIER I\t56.7834\n"
           "2\t2026-07-15\tSBI Pension Funds Pvt. Ltd.\tSM001004\t"
           "SBI PENSION FUND SCHEME C - TIER I\t45.1000\n"
           "3\t2026-07-15\tBroken Row\tSM999999\tBROKEN\t0\n")


def test_parse_ibja_normalises_to_per_gram():
    rates = parse_ibja(IBJA_HTML)
    assert rates == {"gold": 14167.9, "silver": 217.43}   # PM wins, /10 & /1000
    # AM fallback when PM spans are missing
    am_only = IBJA_HTML.replace("lblGold999_PM", "x").replace("lblSilver999_PM", "y")
    assert parse_ibja(am_only) == {"gold": 14171.3, "silver": 218.49}
    assert parse_ibja("<html>nothing here</html>") == {}


def test_market_implied_fallback_median():
    prices = PriceData()
    # three SGB tranches straddling the spot + a silver ETF unit ≈ 1 g
    for i, (sym, isin, close) in enumerate([
            ("SGBAUG28", "IN0020180AA1", 13900.0),
            ("SGBDEC29", "IN0020190BB2", 14050.0),
            ("SGBMAY31", "IN0020210CC3", 14200.0),
            ("SILVERBEES", "INF204KB27I4", 216.0)]):
        prices.prices[isin] = {"close": close, "prev": close}
        prices.master_rows.append((sym, sym, isin))
    rates = derive_from_bhavcopy(prices)
    assert rates["gold"] == 14050.0            # median of the three tranches
    assert rates["silver"] == 216.0


def test_nps_parse_tab_separated_by_header():
    out = nps_parse(NPS_TSV)
    assert out.nav_by_code["SM001003"] == 56.7834
    assert out.nav_date == date(2026, 7, 15)
    assert ("SM001003", "SBI PENSION FUND SCHEME E - TIER I",
            "SBI Pension Funds Pvt. Ltd.") in out.master_rows
    assert "SM999999" not in out.nav_by_code   # zero NAV skipped


def test_sgb_coupon_flows_hand_checked():
    data = sample_portfolio()
    data.bullion = [BullionRow(owner="Amit", metal_type="SGB",
                               description="SGB 2023-24", isin="IN0020230AA0",
                               qty=10, buy_price=6000,
                               buy_date=date(2023, 8, 1),
                               rate_auto=14000, maturity=date(2031, 8, 1))]
    flows = sorted(bullion_flows(data, TODAY))
    # coupons: 2.5%/2 × 10 × 6000 = 750 each Feb-1/Aug-1 in (buy, today]
    assert flows[0] == (date(2023, 8, 1), -60000)
    coupons = [f for f in flows if f[1] == 750]
    assert [c[0] for c in coupons] == [date(2024, 2, 1), date(2024, 8, 1),
                                       date(2025, 2, 1), date(2025, 8, 1),
                                       date(2026, 2, 1)]
    assert flows[-1] == (TODAY, 140000)


def test_bullion_value_override_wins():
    r = BullionRow(metal_type="Gold", qty=100, purity=0.916,
                   rate_auto=14000, rate_override=14500)
    assert bullion_value(r) == pytest.approx(100 * 0.916 * 14500)
    r.rate_override = None
    assert bullion_value(r) == pytest.approx(100 * 0.916 * 14000)


def _wave2_data():
    data = sample_portfolio()
    data.class_settings["gold_silver"] = ClassSetting(enabled=True)
    data.class_settings["nps"] = ClassSetting(enabled=True)
    data.bullion = [
        BullionRow(owner="Amit", metal_type="SGB", description="SGB 2023-24",
                   isin="IN0020230AA0", qty=10, buy_price=6000,
                   buy_date=date(2023, 8, 1), maturity=date(2031, 8, 1)),
        BullionRow(owner="Priya", metal_type="Gold",
                   description="Bangles 22K", qty=50, purity=0.916,
                   buy_price=5200, buy_date=date(2021, 11, 4)),
    ]
    data.nps = [NPSRow(owner="Amit", pran="110012345678",
                       scheme="SBI PENSION FUND SCHEME E - TIER I",
                       units=1000, total_contributed=30000,
                       first_contribution=date(2020, 4, 1))]
    return data


def test_updater_prices_sgb_rates_metal_and_navs_nps(tmp_path):
    data = _wave2_data()
    path = tmp_path / "wb.xlsx"
    build_workbook(data, str(path))

    prices = PriceData(trade_date=TODAY, source="BSE+NSE",
                       sources=["BSE", "NSE"])
    prices.prices["IN0020230AA0"] = {"close": 14100.0, "prev": 14050.0}
    summary = run(path, price_data=prices, amfi_data=AmfiData(), ca_data=[],
                  div_data=[], bullion_rates={"gold": 14167.9},
                  nps_data=nps_parse(NPS_TSV), today=TODAY)
    assert "1 SGB(s) priced" in summary["bullion"]
    assert summary["nps_matched"] == 1

    back = read_workbook(str(path))
    sgb = next(b for b in back.bullion if b.metal_type == "SGB")
    gold = next(b for b in back.bullion if b.metal_type == "Gold")
    assert sgb.rate_auto == 14100.0            # exchange close by ISIN
    assert gold.rate_auto == 14167.9           # injected IBJA rate
    assert back.bullion_rate_asof == TODAY
    npsrow = back.nps[0]
    assert npsrow.current_nav == 56.7834       # via NPS_Master code lookup
    assert npsrow.xirr is not None and npsrow.xirr > 0
    # master merged and sorted for the type-ahead dropdown
    names = [n for _c, n, _p in back.masters.nps_rows]
    assert names == sorted(names, key=str.casefold)

    # rate failure leaves rates + stamp untouched, warns
    summary2 = run(path, price_data=prices, amfi_data=AmfiData(), ca_data=[],
                   div_data=[], bullion_rates={}, nps_data=NpsData(),
                   today=date(2026, 7, 20))
    assert any("gold/silver rate unavailable" in w
               for w in summary2["warnings"])
    back2 = read_workbook(str(path))
    assert next(b for b in back2.bullion
                if b.metal_type == "Gold").rate_auto == 14167.9


def test_wave2_sheets_and_round_trip(tmp_path):
    data = _wave2_data()
    path = tmp_path / "wb.xlsx"
    build_workbook(data, str(path))
    wb = load_workbook(path)
    assert wb["Gold_Silver"].sheet_state == "visible"
    assert wb["NPS_Master"].sheet_state == "visible"   # hides with NPS only
    gs = wb["Gold_Silver"]
    assert 'IF($J4="",$I4,$J4)' in gs["K4"].value      # override wins, live
    nps = wb["NPS"]
    assert "MATCH($C4,NPS_Master!$B:$B,0)" in nps["D4"].value
    assert wb.defined_names["NPS_SchemeList"] is not None

    back = read_workbook(str(path))
    assert asdict(back.bullion[0]) == asdict(data.bullion[0])
    assert asdict(back.nps[0]) == asdict(data.nps[0])
    path2 = tmp_path / "wb2.xlsx"
    build_workbook(back, str(path2))
    assert asdict(read_workbook(str(path2))) == asdict(back)
