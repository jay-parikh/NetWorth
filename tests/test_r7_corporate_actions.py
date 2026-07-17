"""R7: corporate-action parsing, factor math, updater integration, audit sheet."""

from datetime import date

import pytest
from openpyxl import load_workbook

from networth.fetch.amfi import AmfiData
from networth.fetch.bhavcopy import PriceData
from networth.fetch.corporate_actions import (
    dedupe, parse_bse_records, parse_records, parse_subject,
)
from networth.generate import build_workbook
from networth.model import CorporateAction, adjustment_factor
from networth.reader import read_workbook
from networth.sample_data import sample_portfolio
from networth.update import run

TODAY = date(2026, 7, 15)


# ---- subject parsing ----

@pytest.mark.parametrize("subject,expected", [
    ("Face Value Split (Sub-Division) - From Rs 10/- Per Share To Rs 2/- Per Share",
     ("SPLIT", 10.0, 2.0)),
    ("Face Value Split From Re 1/- To Re 0.50", ("SPLIT", 1.0, 0.5)),
    ("Bonus 1:2", ("BONUS", 1.0, 2.0)),
    ("Bonus Issue 3:1", ("BONUS", 3.0, 1.0)),
    ("Consolidation of Shares - From Re 1/- Per Share To Rs 10/- Per Share",
     ("CONSOLIDATION", 1.0, 10.0)),
    ("Dividend - Rs 8 Per Share", None),
    ("Annual General Meeting", None),
    ("Rights 1:5 @ Premium Rs 90", None),
])
def test_parse_subject(subject, expected):
    assert parse_subject(subject) == expected


def test_parse_records():
    records = [
        {"subject": "Bonus 1:1", "exDate": "20-Sep-2024"},
        {"subject": "Dividend - Rs 5 Per Share", "exDate": "01-Feb-2025"},
        {"subject": "Face Value Split From Rs 2 To Re 1", "exDate": "2025-06-10"},
    ]
    out = parse_records(records, "INE0TEST0001", "TESTCO")
    assert len(out) == 2
    assert out[0].type == "BONUS" and out[0].ex_date == date(2024, 9, 20)
    assert out[1].type == "SPLIT" and out[1].factor() == 2.0
    assert all(a.source == "Auto" and a.isin == "INE0TEST0001" for a in out)


def test_parse_bse_records():
    # real Purpose strings observed live on api.bseindia.com (2026-07-15)
    records = [
        {"Purpose": "Bonus issue 1:1", "Ex_date": "28 Oct 2024"},
        {"Purpose": "Bonus issue 2:3", "Ex_date": "15 Jun 2010"},
        {"Purpose": "Stock  Split From Rs.10/- to Rs.2/-", "Ex_date": "14 Jul 2011"},
        {"Purpose": "Dividend - Rs. - 4.2500", "Ex_date": "26 Apr 2001"},
        {"Purpose": "Final Dividend", "Ex_date": "05 Jun 2026"},
    ]
    out = parse_bse_records(records, "INE0TEST0001", "TESTCO")
    assert [(a.type, a.ratio_from, a.ratio_to, a.ex_date) for a in out] == [
        ("BONUS", 1.0, 1.0, date(2024, 10, 28)),
        ("BONUS", 2.0, 3.0, date(2010, 6, 15)),
        ("SPLIT", 10.0, 2.0, date(2011, 7, 14)),
    ]
    assert all(a.source == "Auto" for a in out)


def test_cross_source_dedupe():
    nse = [CorporateAction(symbol="NESTLEIND", isin="INE239A01024", type="SPLIT",
                           ex_date=date(2024, 1, 5), ratio_from=10, ratio_to=1,
                           source="Auto", details="Face Value Split From Rs 10 To Re 1")]
    bse = [CorporateAction(symbol="NESTLEIND", isin="INE239A01024", type="SPLIT",
                           ex_date=date(2024, 1, 5), ratio_from=10, ratio_to=1,
                           source="Auto", details="Stock  Split From Rs.10/- to Rs.1/-"),
           CorporateAction(symbol="NESTLEIND", isin="INE239A01024", type="BONUS",
                           ex_date=date(2025, 8, 8), ratio_from=1, ratio_to=1,
                           source="Auto", details="Bonus issue 1:1")]
    merged = dedupe(nse, bse)
    assert len(merged) == 2                      # split deduped, bonus kept
    split = next(a for a in merged if a.type == "SPLIT")
    assert "Face Value" in split.details         # NSE record won


# ---- factor math ----

def _act(kind, ex, frm, to, isin="INE0TEST0001"):
    return CorporateAction(symbol="T", isin=isin, type=kind, ex_date=ex,
                           ratio_from=frm, ratio_to=to)


def test_adjustment_factor_composition():
    actions = [
        _act("SPLIT", date(2020, 6, 1), 10, 2),        # ×5
        _act("BONUS", date(2022, 3, 1), 1, 1),         # ×2
        _act("CONSOLIDATION", date(2023, 5, 1), 2, 10),  # ×0.2
        _act("SPLIT", date(2030, 1, 1), 10, 5),        # future — ignored
    ]
    # bought before everything → 5 * 2 * 0.2 = 2.0
    assert adjustment_factor("INE0TEST0001", date(2019, 1, 1), TODAY, actions) \
        == pytest.approx(2.0)
    # bought between split and bonus → 2 * 0.2 = 0.4
    assert adjustment_factor("INE0TEST0001", date(2021, 1, 1), TODAY, actions) \
        == pytest.approx(0.4)
    # bought after all past actions → 1
    assert adjustment_factor("INE0TEST0001", date(2024, 1, 1), TODAY, actions) == 1.0
    # different ISIN untouched
    assert adjustment_factor("INE0OTHER001", date(2019, 1, 1), TODAY, actions) == 1.0


# ---- end-to-end through the updater ----

def test_update_applies_actions_and_keeps_manual(tmp_path):
    data = sample_portfolio()
    isin_by_name = {n: i for _s, n, i in data.masters.stock_rows}
    rel = isin_by_name["RELIANCE INDUSTRIES LTD."]
    itc = isin_by_name["ITC LTD."]
    # a pre-existing manual action (user-entered) must survive the update
    data.corporate_actions.append(CorporateAction(
        symbol="ITC", isin=itc, type="BONUS", ex_date=date(2024, 1, 5),
        ratio_from=1, ratio_to=2, source="Manual", details="typed by user"))
    path = tmp_path / "wb.xlsx"
    build_workbook(data, str(path))

    fetched = [CorporateAction(symbol="RELIANCE", isin=rel, type="SPLIT",
                               ex_date=date(2025, 1, 10), ratio_from=10,
                               ratio_to=5, source="Auto",
                               details="Face Value Split From Rs 10 To Rs 5")]
    summary = run(path, price_data=PriceData(trade_date=TODAY, source="TEST"),
                  amfi_data=AmfiData(), ca_data=fetched, today=TODAY)
    assert summary["ca_rows"] == 2
    assert summary["ca_adjusted_rows"] == 2      # RELIANCE row + ITC row

    back = read_workbook(str(path))
    sources = {a.source for a in back.corporate_actions}
    assert sources == {"Auto", "Manual"}
    rel_row = next(r for r in back.equity if r.scrip == "RELIANCE INDUSTRIES LTD.")
    itc_row = next(r for r in back.equity if r.scrip == "ITC LTD.")
    assert rel_row.ca_factor == pytest.approx(2.0)     # 10/5 split
    assert itc_row.ca_factor == pytest.approx(1.5)     # bonus 1:2
    # raw inputs untouched (SPEC §6.7: never mutate user rows)
    assert rel_row.qty == 50 and rel_row.avg_cost == 964.9

    # idempotent: run again with the same feed → same factors, no duplicates
    summary2 = run(path, price_data=PriceData(trade_date=TODAY, source="TEST"),
                   amfi_data=AmfiData(), ca_data=fetched, today=TODAY)
    assert summary2["ca_rows"] == 2
    back2 = read_workbook(str(path))
    rel2 = next(r for r in back2.equity if r.scrip == "RELIANCE INDUSTRIES LTD.")
    assert rel2.ca_factor == pytest.approx(2.0)


def test_factor_flows_into_valuation_formulas(tmp_path):
    data = sample_portfolio()
    data.equity[0].ca_factor = 2.0
    path = tmp_path / "wb.xlsx"
    build_workbook(data, str(path))
    wb = load_workbook(path)
    e = wb["Equity"]
    assert e["S4"].value == 2.0
    assert e["I4"].value == '=IF($D4="","",$D4*IF($S4="",1,$S4)*$F4)'
    # demat-view columns derive from the same factor, zero user action
    assert e["O4"].value == '=IF($D4="","",$D4*IF($S4="",1,$S4))'
    assert e["P4"].value == '=IF(OR($D4="",$E4=""),"",$E4*IF($T4="",1,$T4)/IF($S4="",1,$S4))'
    # family exposure counts post-action shares
    bs = wb["By Scrip"]
    assert "Equity!$O:$O" in bs["C4"].value
    ca = wb["Corporate_Actions"]
    assert ca["A3"].value == "Symbol"
    assert ('IF($C4="BONUS",1+$E4/$F4,'
            'IF(OR($C4="DEMERGER",$C4="ISIN_CHANGE"),1,$E4/$F4))'
            in ca["G4"].value)


def test_coverage_warning_for_unverified_holdings(tmp_path, monkeypatch):
    import networth.update as U
    data = sample_portfolio()
    path = tmp_path / "wb.xlsx"
    build_workbook(data, str(path))
    isin_by_name = {n: i for _s, n, i in data.masters.stock_rows}
    all_held = {isin_by_name[r.scrip] for r in data.equity}
    wipro = isin_by_name["WIPRO LTD."]

    # both exchanges answered for everyone except WIPRO
    monkeypatch.setattr(U.ca_mod, "fetch",
                        lambda *a, **kw: ([], all_held - {wipro}, [], 0))
    summary = run(path, price_data=PriceData(trade_date=TODAY, source="TEST"),
                  amfi_data=AmfiData(), today=TODAY)
    assert summary["ca_unverified"] == ["WIPRO LTD."]
    assert any("could NOT be verified" in w and "WIPRO LTD." in w
               for w in summary["warnings"])

    # full coverage → no warning
    monkeypatch.setattr(U.ca_mod, "fetch", lambda *a, **kw: ([], all_held, [], 0))
    summary2 = run(path, price_data=PriceData(trade_date=TODAY, source="TEST"),
                   amfi_data=AmfiData(), today=TODAY)
    assert summary2["ca_unverified"] == []
    assert not any("could NOT be verified" in w for w in summary2["warnings"])


def test_bhavcopy_captures_bse_codes():
    from networth.fetch.bhavcopy import parse
    csv_text = ("TradDt,FinInstrmId,ISIN,TckrSymb,FinInstrmNm,ClsPric,PrvsClsgPric\n"
                "2026-07-15,500325,INE002A01018,RELIANCE,RELIANCE INDUSTRIES LTD.,1291,1296.85\n")
    out = parse(csv_text)
    assert out.codes_by_isin == {"INE002A01018": "500325"}


def test_xirr_uses_adjusted_value(tmp_path):
    from networth.compute.cashflows import equity_flows
    data = sample_portfolio()
    row = data.equity[0]
    base = equity_flows(data, TODAY)
    row.ca_factor = 2.0
    doubled = equity_flows(data, TODAY)
    # terminal inflow for that row doubles, outflow unchanged
    assert doubled[1][1] == pytest.approx(2 * base[1][1])
    assert doubled[0][1] == base[0][1]
