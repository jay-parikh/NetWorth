"""R14: mergers / demergers / ISIN reassignments (SPEC §5.8/§6.15)."""

from datetime import date

import pytest

from networth.fetch.amfi import AmfiData
from networth.fetch.bhavcopy import PriceData
from networth.generate import build_workbook
from networth.model import (CorporateAction, chained_adjustment_factor,
                            cost_adjustment_factor, load_restructures,
                            resolve_isin)
from networth.reader import read_workbook
from networth.sample_data import sample_portfolio
from networth.update import run

TODAY = date(2026, 7, 15)
RELIANCE = "INE002A01018"
NEWCO = "INE0NEWCO014"
CHILD = "INE0CHILD012"


def _merger(ratio_from=1, ratio_to=2, ex=date(2026, 5, 1), type_="MERGER"):
    return CorporateAction(symbol="RELIANCE", isin=RELIANCE, type=type_,
                           ex_date=ex, ratio_from=ratio_from,
                           ratio_to=ratio_to, source="Curated",
                           new_isin=NEWCO, new_name="NEWCO LTD.",
                           new_symbol="NEWCO", cost_pct=100,
                           details="test merger")


def _demerger(parent_pct=60.0, child_pct=40.0, ex=date(2026, 5, 1)):
    return [
        CorporateAction(symbol="RELIANCE", isin=RELIANCE, type="DEMERGER",
                        ex_date=ex, ratio_from=1, ratio_to=1,
                        source="Curated", new_isin=RELIANCE,
                        new_name="RELIANCE INDUSTRIES LTD.",
                        cost_pct=parent_pct, details="parent retention"),
        CorporateAction(symbol="RELIANCE", isin=RELIANCE, type="DEMERGER",
                        ex_date=ex, ratio_from=1, ratio_to=1,
                        source="Curated", new_isin=CHILD,
                        new_name="CHILDCO LTD.", new_symbol="CHILDCO",
                        cost_pct=child_pct, details="spun-off child"),
    ]


# ------------------------------------------------------------ model layer --

def test_shipped_file_loads_and_bad_sum_fails_loudly(tmp_path):
    events = load_restructures()
    hdfc = next(e for e in events if e.isin == "INE001A01036")
    assert hdfc.type == "MERGER" and hdfc.new_isin == "INE040A01034"
    assert hdfc.factor() == pytest.approx(42 / 25)

    bad = tmp_path / "restructures.csv"
    bad.write_text(
        "ex_date,type,old_isin,old_name,old_symbol,new_isin,new_name,"
        "new_symbol,ratio_from,ratio_to,cost_pct,details\n"
        "2026-05-01,DEMERGER,INE0X,X,X,INE0X,X,X,1,1,60,parent\n"
        "2026-05-01,DEMERGER,INE0X,X,X,INE0Y,Y,Y,1,1,30,child\n")
    with pytest.raises(ValueError, match="sums to 90"):
        load_restructures(tmp_path)


def test_resolve_and_chained_factor():
    merger = _merger(ratio_from=42, ratio_to=25, ex=date(2023, 7, 13))
    split_on_successor = CorporateAction(
        symbol="NEWCO", isin=NEWCO, type="SPLIT", ex_date=date(2025, 1, 10),
        ratio_from=10, ratio_to=5, source="Auto")
    actions = [merger, split_on_successor]
    assert resolve_isin(RELIANCE, actions, TODAY) == NEWCO
    assert resolve_isin("INE_OTHER", actions, TODAY) == "INE_OTHER"
    # merger ratio folds in AND the successor's later split keeps applying
    f = chained_adjustment_factor(RELIANCE, date(2020, 1, 1), TODAY, actions)
    assert f == pytest.approx((42 / 25) * 2)
    # a lot bought AFTER the merger (already in new terms) gets only the split
    f2 = chained_adjustment_factor(NEWCO, date(2024, 1, 1), TODAY, actions)
    assert f2 == pytest.approx(2)


def test_cost_adjustment_factor_retention_only():
    events = _demerger(60, 40)
    f = cost_adjustment_factor(RELIANCE, date(2018, 1, 31), TODAY, events)
    assert f == pytest.approx(0.6)     # retention row only, child row ignored
    # lots bought after the ex-date keep full cost
    assert cost_adjustment_factor(RELIANCE, date(2026, 6, 1), TODAY, events) == 1.0


# --------------------------------------------------------------- scenarios --

def test_merger_prices_via_successor_invested_untouched(tmp_path):
    data = sample_portfolio()
    path = tmp_path / "wb.xlsx"
    build_workbook(data, str(path))

    prices = PriceData(trade_date=TODAY, source="BSE+NSE", sources=["BSE", "NSE"])
    prices.prices[NEWCO] = {"close": 3000.0, "prev": 2980.0}
    isin_by_name = {n: i for _s, n, i in data.masters.stock_rows}
    for row in data.equity:                      # everyone else stays quoted
        if row.scrip != "RELIANCE INDUSTRIES LTD.":
            prices.prices[isin_by_name[row.scrip]] = {"close": 100.0, "prev": 99.0}
    summary = run(path, price_data=prices, amfi_data=AmfiData(), ca_data=[],
                  div_data=[], restructures=[_merger(1, 2)], today=TODAY)
    assert summary["restructure_children"] == 0

    back = read_workbook(str(path))
    rel = next(r for r in back.equity if r.scrip == "RELIANCE INDUSTRIES LTD.")
    assert rel.close == 3000.0                     # priced via the successor
    assert rel.ca_factor == pytest.approx(0.5)     # 1 new per 2 old
    assert rel.cost_factor is None                 # cost carries in full
    assert rel.qty == 50 and rel.avg_cost == 964.9 # user cells byte-identical
    assert rel.flag == "MERGED→NEWCO LTD."
    assert back.masters.stock_status[RELIANCE][0] == "Merged"
    assert NEWCO in {i for _s, _n, i in back.masters.stock_rows}
    ca = next(a for a in back.corporate_actions if a.type == "MERGER")
    assert ca.source == "Curated" and ca.applied == TODAY

    # second run: idempotent — same rows, no escalation of the consumed ISIN
    summary2 = run(path, price_data=prices, amfi_data=AmfiData(), ca_data=[],
                   div_data=[], restructures=[_merger(1, 2)],
                   today=date(2026, 9, 30))       # 77 days later, still Merged
    assert summary2["suspended"] == 0
    back2 = read_workbook(str(path))
    assert back2.masters.stock_status[RELIANCE][0] == "Merged"
    assert len(back2.equity) == len(back.equity)


def test_isin_change_routes_price_without_factor(tmp_path):
    data = sample_portfolio()
    path = tmp_path / "wb.xlsx"
    build_workbook(data, str(path))
    prices = PriceData(trade_date=TODAY, source="T", sources=["BSE", "NSE"])
    prices.prices[NEWCO] = {"close": 1600.0, "prev": 1590.0}
    run(path, price_data=prices, amfi_data=AmfiData(), ca_data=[],
        div_data=[], restructures=[_merger(type_="ISIN_CHANGE")], today=TODAY)
    back = read_workbook(str(path))
    rel = next(r for r in back.equity if r.scrip == "RELIANCE INDUSTRIES LTD.")
    assert rel.close == 1600.0
    assert rel.ca_factor is None                   # 1:1, S stays blank
    assert rel.flag == f"ISIN→{NEWCO}"
    assert back.masters.stock_status[RELIANCE][0] == "Renamed"


def test_demerger_60_40_conserves_invested_to_the_rupee(tmp_path):
    data = sample_portfolio()
    path = tmp_path / "wb.xlsx"
    build_workbook(data, str(path))

    prices = PriceData(trade_date=TODAY, source="T", sources=["BSE", "NSE"])
    summary = run(path, price_data=prices, amfi_data=AmfiData(), ca_data=[],
                  div_data=[], restructures=_demerger(60, 40), today=TODAY)
    assert summary["restructure_children"] == 1

    back = read_workbook(str(path))
    parent = next(r for r in back.equity
                  if r.scrip == "RELIANCE INDUSTRIES LTD.")
    child = next(r for r in back.equity if r.scrip == "CHILDCO LTD.")
    original = 50 * 964.9
    assert parent.cost_factor == pytest.approx(0.6)
    assert parent.qty == 50 and parent.avg_cost == 964.9   # untouched
    assert child.qty == 50 and child.isin_override == CHILD
    assert child.cost_date == parent.cost_date            # holding period inherited
    assert child.flag.startswith("DEMERGER:" + RELIANCE)
    parent_invested = parent.qty * parent.avg_cost * parent.cost_factor
    child_invested = child.qty * child.avg_cost
    assert parent_invested + child_invested == pytest.approx(original, abs=0.01)
    # unquoted child: no price yet, prices on first appearance
    assert child.close is None

    # re-run: applied → no duplicate children
    run(path, price_data=prices, amfi_data=AmfiData(), ca_data=[],
        div_data=[], restructures=_demerger(60, 40), today=TODAY)
    back2 = read_workbook(str(path))
    assert len([r for r in back2.equity if r.scrip == "CHILDCO LTD."]) == 1

    # the user deletes the child (sold it) — the next run respects that
    back2.equity = [r for r in back2.equity if r.scrip != "CHILDCO LTD."]
    build_workbook(back2, str(path))
    run(path, price_data=prices, amfi_data=AmfiData(), ca_data=[],
        div_data=[], restructures=_demerger(60, 40), today=TODAY)
    back3 = read_workbook(str(path))
    assert not [r for r in back3.equity if r.scrip == "CHILDCO LTD."]
    # child prices automatically once quoted
    prices.prices[CHILD] = {"close": 250.0, "prev": 245.0}
    run(path, price_data=prices, amfi_data=AmfiData(), ca_data=[],
        div_data=[], restructures=[], today=TODAY)   # works without the event too


def test_manual_row_overrides_curated(tmp_path):
    data = sample_portfolio()
    manual = _merger(2, 1)                 # user says 2 new per 1 old
    manual.source = "Manual"
    data.corporate_actions.append(manual)
    path = tmp_path / "wb.xlsx"
    build_workbook(data, str(path))
    prices = PriceData(trade_date=TODAY, source="T", sources=["BSE", "NSE"])
    prices.prices[NEWCO] = {"close": 900.0, "prev": 890.0}
    run(path, price_data=prices, amfi_data=AmfiData(), ca_data=[],
        div_data=[], restructures=[_merger(1, 2)], today=TODAY)
    back = read_workbook(str(path))
    rel = next(r for r in back.equity if r.scrip == "RELIANCE INDUSTRIES LTD.")
    assert rel.ca_factor == pytest.approx(2.0)      # the Manual ratio won
    kinds = [(a.source, a.ratio_from) for a in back.corporate_actions
             if a.type == "MERGER"]
    assert ("Manual", 2) in kinds and ("Curated", 1) not in kinds


def test_restructure_columns_round_trip(tmp_path):
    data = sample_portfolio()
    ev = _merger(42, 25, ex=date(2023, 7, 13))
    ev.applied = date(2026, 7, 1)
    data.corporate_actions.append(ev)
    data.equity[0].flag = "MERGED→NEWCO LTD."
    data.equity[0].cost_factor = 0.6
    path = tmp_path / "wb.xlsx"
    build_workbook(data, str(path))
    back = read_workbook(str(path))
    got = next(a for a in back.corporate_actions if a.type == "MERGER")
    assert (got.new_isin, got.cost_pct, got.applied) == (
        NEWCO, 100, date(2026, 7, 1))
    assert back.equity[0].flag == "MERGED→NEWCO LTD."
    assert back.equity[0].cost_factor == 0.6
    from dataclasses import asdict
    path2 = tmp_path / "wb2.xlsx"
    build_workbook(back, str(path2))
    assert asdict(read_workbook(str(path2))) == asdict(back)
