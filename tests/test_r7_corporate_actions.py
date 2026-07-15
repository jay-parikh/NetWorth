"""R7: corporate-action parsing, factor math, updater integration, audit sheet."""

from datetime import date

import pytest
from openpyxl import load_workbook

from networth.fetch.amfi import AmfiData
from networth.fetch.bhavcopy import PriceData
from networth.fetch.corporate_actions import parse_records, parse_subject
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
    assert e["R4"].value == 2.0
    assert e["I4"].value == '=IF($D4="","",$D4*IF($R4="",1,$R4)*$F4)'
    ca = wb["Corporate_Actions"]
    assert ca["A3"].value == "Symbol"
    assert 'IF($C4="BONUS",1+$E4/$F4,$E4/$F4)' in ca["G4"].value


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
