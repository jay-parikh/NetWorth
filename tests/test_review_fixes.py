"""Regression tests for the 2026-07-17 code-review fixes (v1.2–v1.4 review).

One test (or small group) per confirmed finding, in the review's severity
order. Each would have failed before its fix.
"""

from datetime import date, timedelta
from pathlib import Path

import pytest

from networth.fetch.amfi import AmfiData
from networth.fetch.bhavcopy import PriceData, fetch
from networth.fetch.corporate_actions import dedupe_dividends, parse_dividend
from networth.generate import build_workbook
from networth.model import (CorporateAction, DividendRow, EquityRow,
                            ManualAssetRow)
from networth.reader import read_workbook
from networth.sample_data import sample_portfolio
from networth.update import run
import networth.update as U

TODAY = date(2026, 7, 15)
RELIANCE = "INE002A01018"
NEWCO = "INE0NEWCO014"
CHILD = "INE0CHILD012"
CHILD2 = "INE0CHILD204"


def _prices(**extra):
    p = PriceData(trade_date=TODAY, source="T", sources=["BSE", "NSE"])
    p.prices.update(extra)
    return p


def _demerger(parent_pct, child_pct, ex, parent_isin=RELIANCE,
              child_isin=CHILD, symbol="RELIANCE"):
    return [
        CorporateAction(symbol=symbol, isin=parent_isin, type="DEMERGER",
                        ex_date=ex, ratio_from=1, ratio_to=1,
                        source="Curated", new_isin=parent_isin,
                        new_name="PARENT", cost_pct=parent_pct,
                        details="parent retention"),
        CorporateAction(symbol=symbol, isin=parent_isin, type="DEMERGER",
                        ex_date=ex, ratio_from=1, ratio_to=1,
                        source="Curated", new_isin=child_isin,
                        new_name=f"CHILDCO {child_isin[-2:]}",
                        new_symbol="CHILDCO", cost_pct=child_pct,
                        details="spun-off child"),
    ]


def _merger(ratio_from=1, ratio_to=1, ex=date(2026, 5, 1)):
    return CorporateAction(symbol="RELIANCE", isin=RELIANCE, type="MERGER",
                           ex_date=ex, ratio_from=ratio_from,
                           ratio_to=ratio_to, source="Curated",
                           new_isin=NEWCO, new_name="NEWCO LTD.",
                           new_symbol="NEWCO", cost_pct=100)


# 1 — Corporate_Actions capacity: Manual rows must never overflow away ------

def test_manual_ca_rows_survive_a_flooded_feed(tmp_path):
    data = sample_portfolio()
    manual = CorporateAction(symbol="RELIANCE", isin=RELIANCE, type="SPLIT",
                             ex_date=date(2026, 1, 5), ratio_from=5,
                             ratio_to=5, source="Manual", details="mine")
    data.corporate_actions.append(manual)
    path = tmp_path / "wb.xlsx"
    build_workbook(data, str(path))
    flood = [CorporateAction(symbol="RELIANCE", isin=RELIANCE, type="SPLIT",
                             ex_date=date(2010, 1, 1) + timedelta(days=i),
                             ratio_from=5, ratio_to=5, source="Auto")
             for i in range(250)]
    summary = run(path, price_data=_prices(), amfi_data=AmfiData(),
                  ca_data=flood, div_data=[], today=TODAY)
    assert any("Corporate_Actions sheet is full" in w
               for w in summary["warnings"])
    back = read_workbook(str(path))
    assert len(back.corporate_actions) == 200            # sheet capacity
    assert back.corporate_actions[0].source == "Manual"  # user rows lead
    # the drop is deterministic: the OLDEST Auto rows gave way
    auto_dates = [a.ex_date for a in back.corporate_actions
                  if a.source == "Auto"]
    assert min(auto_dates) == date(2010, 1, 1) + timedelta(days=51)


# 2 — a second demerger on the same parent scales by earlier retention ------

def test_sequential_demergers_conserve_invested(tmp_path):
    data = sample_portfolio()
    path = tmp_path / "wb.xlsx"
    build_workbook(data, str(path))
    events = (_demerger(60, 40, date(2026, 5, 1), child_isin=CHILD)
              + _demerger(70, 30, date(2026, 6, 1), child_isin=CHILD2))
    summary = run(path, price_data=_prices(), amfi_data=AmfiData(),
                  ca_data=[], div_data=[], restructures=events, today=TODAY)
    assert summary["restructure_children"] == 2
    back = read_workbook(str(path))
    parent = next(r for r in back.equity
                  if r.scrip == "RELIANCE INDUSTRIES LTD.")
    child_a = next(r for r in back.equity if r.isin_override == CHILD)
    child_b = next(r for r in back.equity if r.isin_override == CHILD2)
    original = 50 * 964.9
    assert parent.cost_factor == pytest.approx(0.6 * 0.7)
    assert child_a.qty * child_a.avg_cost == pytest.approx(0.4 * original)
    # child B's basis comes from what REMAINED after demerger A (60%), not
    # from the full original cost — the 112%-of-original bug
    assert child_b.qty * child_b.avg_cost == pytest.approx(
        0.6 * 0.3 * original)
    total = (parent.qty * parent.avg_cost * parent.cost_factor
             + child_a.qty * child_a.avg_cost + child_b.qty * child_b.avg_cost)
    assert total == pytest.approx(original, abs=0.01)


# 3 — children are computed AFTER the CA refresh, and defer when unverified -

def test_demerger_defers_until_history_verified(tmp_path, monkeypatch):
    data = sample_portfolio()
    path = tmp_path / "wb.xlsx"
    build_workbook(data, str(path))
    events = _demerger(60, 40, date(2026, 5, 1))

    def _fetch_reliance_down(symbols, bse_codes, **kw):
        return [], set(symbols.values()) - {RELIANCE}, [], 0

    monkeypatch.setattr(U.ca_mod, "fetch", _fetch_reliance_down)
    summary = run(path, price_data=_prices(), amfi_data=AmfiData(),
                  restructures=events, today=TODAY)
    assert any("deferred" in w for w in summary["warnings"])
    assert summary["restructure_children"] == 0
    back = read_workbook(str(path))
    dem = next(a for a in back.corporate_actions if a.type == "DEMERGER"
               and a.new_isin == CHILD)
    assert dem.applied is None                       # NOT stamped → retries
    # next run the feed answers (injected = trusted): the child appears, and
    # its qty includes the bonus the feed knows about (1:1 → ×2)
    bonus = CorporateAction(symbol="RELIANCE", isin=RELIANCE, type="BONUS",
                            ex_date=date(2026, 4, 10), ratio_from=1,
                            ratio_to=1, source="Auto", details="1:1 bonus")
    summary2 = run(path, price_data=_prices(), amfi_data=AmfiData(),
                   ca_data=[bonus], div_data=[], restructures=events,
                   today=TODAY)
    assert summary2["restructure_children"] == 1
    back2 = read_workbook(str(path))
    child = next(r for r in back2.equity if r.isin_override == CHILD)
    assert child.qty == pytest.approx(100)           # 50 × 2 (bonus) × 1:1


# 4 — a one-symbol feed failure keeps that stock's Auto rows ----------------

def test_partial_ca_feed_failure_preserves_applied_rows(tmp_path, monkeypatch):
    data = sample_portfolio()
    split = CorporateAction(symbol="RELIANCE", isin=RELIANCE, type="SPLIT",
                            ex_date=date(2026, 5, 2), ratio_from=10,
                            ratio_to=5, source="Auto", details="10→5 split")
    data.corporate_actions.append(split)
    data.dividends.append(DividendRow(
        fy="2026-27", owner="Amit", scrip="RELIANCE INDUSTRIES LTD.",
        isin=RELIANCE, div_type="Interim", ex_date=date(2026, 5, 20),
        rate=9, qty=100, source="Auto", details="Interim Dividend Rs 9"))
    path = tmp_path / "wb.xlsx"
    build_workbook(data, str(path))

    def _fetch_reliance_down(symbols, bse_codes, **kw):
        return [], set(symbols.values()) - {RELIANCE}, [], 0

    monkeypatch.setattr(U.ca_mod, "fetch", _fetch_reliance_down)
    summary = run(path, price_data=_prices(), amfi_data=AmfiData(),
                  today=TODAY)
    assert any("existing rows are kept" in w for w in summary["warnings"])
    back = read_workbook(str(path))
    kept = [a for a in back.corporate_actions
            if a.isin == RELIANCE and a.type == "SPLIT"]
    assert len(kept) == 1                            # not reverted
    rel = next(r for r in back.equity if r.scrip == "RELIANCE INDUSTRIES LTD.")
    assert rel.ca_factor == pytest.approx(2.0)       # still applied
    divs = [d for d in back.dividends if d.isin == RELIANCE]
    assert len(divs) == 1 and divs[0].fy == "2026-27"


# 5 — demerger children never silently fall off the Equity sheet ------------

def test_demerger_blocked_by_full_equity_sheet_retries(tmp_path):
    from networth.model import EQUITY_LAST_ROW, FIRST_DATA_ROW
    cap = EQUITY_LAST_ROW - FIRST_DATA_ROW + 1       # sheet capacity
    data = sample_portfolio()
    while len(data.equity) < cap:
        data.equity.append(EquityRow(owner="Amit", scrip="INFOSYS LTD.",
                                     qty=1, avg_cost=1,
                                     cost_date=date(2020, 1, 1)))
    path = tmp_path / "wb.xlsx"
    build_workbook(data, str(path))
    events = _demerger(60, 40, date(2026, 5, 1))
    summary = run(path, price_data=_prices(), amfi_data=AmfiData(),
                  ca_data=[], div_data=[], restructures=events, today=TODAY)
    assert any("Equity sheet is full" in w for w in summary["warnings"])
    assert summary["restructure_children"] == 0
    back = read_workbook(str(path))
    assert not [r for r in back.equity if r.isin_override == CHILD]
    dem = next(a for a in back.corporate_actions
               if a.type == "DEMERGER" and a.new_isin == CHILD)
    assert dem.applied is None                       # retried next run
    # the user frees rows → the same event now applies
    back.equity = [r for r in back.equity
                   if not (r.scrip == "INFOSYS LTD." and r.qty == 1)]
    build_workbook(back, str(path))
    summary2 = run(path, price_data=_prices(), amfi_data=AmfiData(),
                   ca_data=[], div_data=[], restructures=events, today=TODAY)
    assert summary2["restructure_children"] == 1


# 6 — a demerger on a merger-successor reaches lots held via the old ISIN ---

def test_demerger_on_merger_successor_creates_children(tmp_path):
    data = sample_portfolio()
    path = tmp_path / "wb.xlsx"
    build_workbook(data, str(path))
    events = ([_merger(42, 25, ex=date(2023, 7, 13))]
              + _demerger(80, 20, date(2026, 5, 1), parent_isin=NEWCO,
                          symbol="NEWCO"))
    summary = run(path, price_data=_prices(NEWCO={"close": 300, "prev": 299}),
                  amfi_data=AmfiData(), ca_data=[], div_data=[],
                  restructures=events, today=TODAY)
    assert summary["restructure_children"] == 1
    back = read_workbook(str(path))
    rel = next(r for r in back.equity if r.scrip == "RELIANCE INDUSTRIES LTD.")
    child = next(r for r in back.equity if r.isin_override == CHILD)
    original = 50 * 964.9
    assert child.qty == pytest.approx(50 * 42 / 25)  # merger ratio folds in
    assert rel.cost_factor == pytest.approx(0.8)     # docked via the chain
    assert (rel.qty * rel.avg_cost * rel.cost_factor
            + child.qty * child.avg_cost) == pytest.approx(original, abs=0.01)


# 7 — merged holdings earn the successor's dividends ------------------------

def test_merged_holding_gets_successor_dividends(tmp_path):
    data = sample_portfolio()
    path = tmp_path / "wb.xlsx"
    build_workbook(data, str(path))
    div = DividendRow(scrip="NEWCO", isin=NEWCO, div_type="Final",
                      ex_date=date(2026, 6, 15), rate=10, source="Auto",
                      details="Final Dividend - Rs 10 Per Share")
    run(path, price_data=_prices(NEWCO={"close": 300, "prev": 299}),
        amfi_data=AmfiData(), ca_data=[], div_data=[div],
        restructures=[_merger(1, 2)], today=TODAY)   # 1 new per 2 old
    back = read_workbook(str(path))
    got = [d for d in back.dividends if d.isin == NEWCO]
    assert len(got) == 1 and got[0].owner == "Amit"
    assert got[0].qty == pytest.approx(25)           # 50 old × 1/2


def test_ca_fetch_queries_the_successor_symbol(tmp_path, monkeypatch):
    data = sample_portfolio()
    path = tmp_path / "wb.xlsx"
    build_workbook(data, str(path))
    seen: dict = {}

    def _spy(symbols, bse_codes, **kw):
        seen.update(symbols)
        return [], set(symbols.values()), [], 0

    monkeypatch.setattr(U.ca_mod, "fetch", _spy)
    run(path, price_data=_prices(NEWCO={"close": 300, "prev": 299}),
        amfi_data=AmfiData(), restructures=[_merger(1, 2)], today=TODAY)
    assert seen.get("NEWCO") == NEWCO                # successor queried


# 8 — dividend rate parser: face values and 'Sha-re' are not rates ----------

def test_parse_dividend_rejects_face_value_and_word_tails():
    assert parse_dividend(
        "Dividend - 300% on face value of Rs.2/- each") is None
    assert parse_dividend("Interim Dividend 250% Per Share 2024") is None
    # the real formats still parse
    assert parse_dividend("Final Dividend - Rs 7.85 Per Share") == ("Final", 7.85)
    assert parse_dividend("Interim Dividend - Rs. - 8.0000") == ("Interim", 8.0)
    assert parse_dividend("Dividend Re. 1/-") == ("Final", 1.0)
    assert parse_dividend("Special Dividend of ₹2.50") == ("Special", 2.5)


# 9 — the same payout worded differently per exchange counts once -----------

def test_dedupe_dividends_ignores_type_wording():
    nse = [DividendRow(scrip="X", isin="INE1", div_type="Final",
                       ex_date=date(2026, 6, 1), rate=8)]
    bse = [DividendRow(scrip="X", isin="INE1", div_type="Interim",
                       ex_date=date(2026, 6, 1), rate=8)]
    out = dedupe_dividends(nse, bse)
    assert len(out) == 1 and out[0].div_type == "Final"   # NSE wins
    # two genuinely distinct same-day payouts (different rates) both survive
    bse2 = [DividendRow(scrip="X", isin="INE1", div_type="Special",
                        ex_date=date(2026, 6, 1), rate=3)]
    assert len(dedupe_dividends(nse, bse2)) == 2


# 10 — the frozen exe bundles every runtime CSV -----------------------------

def test_pyinstaller_spec_bundles_all_runtime_csvs():
    spec = (Path(__file__).resolve().parents[1]
            / "packaging" / "networth-update.spec").read_text()
    for name in ("banks_in.csv", "fmv_2018-01-31.csv", "ppf_rates.csv",
                 "epf_rates.csv", "bullion_proxies.csv", "restructures.csv",
                 "tax_rules_in.csv"):
        assert name in spec, f"{name} missing from the PyInstaller datas"


# 11 — an NSE 200-but-HTML response degrades to BSE-only --------------------

def test_nse_html_challenge_degrades_to_bse_only():
    from test_r8_dual_source import BSE_CSV, _FakeSession, _Resp
    ymd = TODAY.strftime("%Y%m%d")
    sess = _FakeSession({
        f"BhavCopy_BSE_CM_0_0_0_{ymd}": _Resp(text=BSE_CSV),
        f"BhavCopy_NSE_CM_0_0_0_{ymd}": _Resp(
            content=b"<html>bot check</html>"),
        "www.nseindia.com": _Resp(text="warmup"),
    })
    out = fetch(session=sess, today=TODAY)
    assert out.sources == ["BSE"] and out.prices


# 12 — a 0.0 cost factor is a real value, not "blank" -----------------------

def test_zero_cost_factor_zeroes_the_xirr_outflow():
    from networth.compute.cashflows import equity_flows
    from networth.model import PortfolioData
    data = PortfolioData(equity=[
        EquityRow(owner="A", scrip="X", qty=10, avg_cost=100, close=50,
                  cost_date=date(2020, 1, 1), cost_factor=0.0)])
    flows = equity_flows(data, TODAY)
    assert flows[0] == (date(2020, 1, 1), 0.0)       # not -1000


# 13 — Manual_Assets class labels: Excel and Python must agree --------------

def test_manual_class_case_normalises_and_unknown_warns(tmp_path):
    data = sample_portfolio()
    data.manual_assets.append(ManualAssetRow(
        owner="Amit", asset_class="real estate", description="typed lower",
        value=100000, as_on=date(2026, 7, 1)))
    data.manual_assets.append(ManualAssetRow(
        owner="Amit", asset_class="RealEstate", description="typo'd",
        value=50000, as_on=date(2026, 7, 1)))
    path = tmp_path / "wb.xlsx"
    build_workbook(data, str(path))
    back = read_workbook(str(path))
    classes = [r.asset_class for r in back.manual_assets]
    assert "Property" in classes and "real estate" not in classes
    assert "RealEstate" in classes                   # unknown: kept, warned
    summary = run(path, price_data=_prices(), amfi_data=AmfiData(),
                  ca_data=[], div_data=[], today=TODAY)
    assert any("unrecognised Class 'RealEstate'" in w
               for w in summary["warnings"])


# 14 — a restructure flag never evicts the FMV marker -----------------------

def test_flag_and_fmv_marker_coexist_across_round_trip(tmp_path):
    data = sample_portfolio()
    data.equity[0].fmv_used = True
    data.equity[0].flag = "MERGED→NEWCO LTD."
    path = tmp_path / "wb.xlsx"
    build_workbook(data, str(path))
    back = read_workbook(str(path))
    assert back.equity[0].fmv_used is True
    assert back.equity[0].flag == "MERGED→NEWCO LTD."


# 15 — the shipped sample honours the delete-to-hide onboarding -------------

def test_sample_ships_new_classes_toggled_no():
    data = sample_portfolio()
    for key in ("equity", "mutual_funds", "fixed_deposits", "ppf", "bonds"):
        assert data.class_settings[key].enabled is True
    for key in ("epf", "gold_silver", "nps", "real_estate", "cash",
                "insurance", "other_assets"):
        assert data.class_settings[key].enabled is False, key
