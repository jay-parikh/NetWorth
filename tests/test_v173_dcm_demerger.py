"""v1.7.3 curated data — DCM Shriram Industries three-way demerger.

Reported by a user (2026-08-09) holding DCMSRIND from before the December
2025 demerger: Corporate_Actions showed only the 2021 split. Demergers
are never auto-fetched — they ship in data/restructures.csv, a standing
release duty (ROADMAP R14).

Facts (company announcement 03-02-2026, NCLT sanction 21-11-2025,
scheme effective 17-12-2025, record date 26-12-2025): 1 share of each
resultant company per share held, cost apportioned 42.66% DCMSRIND /
25.22% DSFCL / 32.12% DCMSIL.
"""

from collections import defaultdict
from datetime import date

import pytest

from networth import model as M
from networth.compute.restructures import apply_demergers
from networth.model import load_restructures

TODAY = date(2026, 8, 9)
DCM = "INE843D01027"
DSFCL = "INE0OFM01015"
DCMSIL = "INE0OU201013"


def _dcm_events():
    return [e for e in load_restructures() if e.isin == DCM]


def test_shipped_file_carries_the_dcm_demerger():
    evs = _dcm_events()
    assert len(evs) == 3
    assert {e.new_isin for e in evs} == {DCM, DSFCL, DCMSIL}
    by_isin = {e.new_isin: e for e in evs}
    assert by_isin[DCM].cost_pct == pytest.approx(42.66)      # retention
    assert by_isin[DSFCL].cost_pct == pytest.approx(25.22)
    assert by_isin[DCMSIL].cost_pct == pytest.approx(32.12)
    for e in evs:
        assert e.type == "DEMERGER"
        assert e.ex_date == date(2025, 12, 26)
        assert (e.ratio_from, e.ratio_to) == (1, 1)
        assert e.factor() == 1.0        # a demerger never changes parent qty


def test_every_shipped_demerger_apportions_exactly_one_hundred():
    # generic guard: any future curated event with a typo'd split fails here
    groups = defaultdict(float)
    for e in load_restructures():
        if e.type == "DEMERGER":
            groups[(e.isin, e.ex_date)] += e.cost_pct or 0.0
    assert groups, "no demergers shipped?"
    for key, total in groups.items():
        assert total == pytest.approx(100.0), f"{key} sums to {total}"


def test_pre_demerger_lot_gains_two_children_with_the_cost_split():
    d = M.PortfolioData(persons=["Amit"])
    d.masters.stock_rows = [
        ("DCMSRIND", "DCM SHRIRAM INDUSTRIES LTD.", DCM),
        ("DSFCL", "DCM SHRIRAM FINE CHEM LTD", DSFCL),
        ("DCMSIL", "DCM SHRIRAM INTERNATIONAL", DCMSIL)]
    # 100 shares bought in 2019 — before BOTH the 2021 split and the demerger
    d.equity = [M.EquityRow(owner="Amit", scrip="DCM SHRIRAM INDUSTRIES LTD.",
                            qty=100, avg_cost=90.0,
                            cost_date=date(2019, 6, 10))]
    split = M.CorporateAction(symbol="DCMSRIND", isin=DCM, type="SPLIT",
                              ex_date=date(2021, 9, 15), ratio_from=10,
                              ratio_to=2, source="Auto")
    evs = _dcm_events()
    d.corporate_actions = [split] + evs

    added, warns = apply_demergers(d, evs, ca_checked={DCM}, ca_trusted=False,
                                   price_data=None, today=TODAY)
    assert added == 2 and not warns
    assert len(d.equity) == 3

    parent = d.equity[0]
    assert parent.qty == 100                      # raw row never rewritten
    f = M.chained_adjustment_factor(DCM, M.qty_anchor(parent), TODAY,
                                    d.corporate_actions)
    assert f == pytest.approx(5.0)                # the 2021 split only
    assert M.cost_adjustment_factor(DCM, M.qty_anchor(parent), TODAY,
                                    d.corporate_actions) == pytest.approx(0.4266)

    kids = {r.isin_override: r for r in d.equity[1:]}
    assert set(kids) == {DSFCL, DCMSIL}
    for isin, r in kids.items():
        assert r.qty == pytest.approx(500.0)      # 1:1 on the post-split count
        assert r.cost_date == date(2019, 6, 10)   # holding period inherited
        assert r.flag.startswith("DEMERGER:")

    # every rupee accounted for: 100 x 90 = 9,000 across the three rows
    total = sum((r.qty or 0) * (r.avg_cost or 0)
                * M.cost_adjustment_factor(r.isin_override or DCM,
                                           M.qty_anchor(r), TODAY,
                                           d.corporate_actions)
                for r in d.equity)
    assert total == pytest.approx(9000.0, abs=0.05)


def test_lot_bought_after_the_record_date_gets_no_children():
    d = M.PortfolioData(persons=["Amit"])
    d.masters.stock_rows = [("DCMSRIND", "DCM SHRIRAM INDUSTRIES LTD.", DCM)]
    d.equity = [M.EquityRow(owner="Amit", scrip="DCM SHRIRAM INDUSTRIES LTD.",
                            qty=50, avg_cost=40.0,
                            cost_date=date(2026, 1, 15))]
    evs = _dcm_events()
    d.corporate_actions = list(evs)
    added, _w = apply_demergers(d, evs, ca_checked={DCM}, ca_trusted=False,
                                price_data=None, today=TODAY)
    assert added == 0 and len(d.equity) == 1
