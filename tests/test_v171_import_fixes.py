"""v1.7.1 import fixes — regression tests (SPEC §6.18, §3.11).

Four defects hit with real broker files on 2026-07-19:
1. imported holdings double-applied corporate actions (Qty today ×5/×15);
2. fund units / NCDs inside a holdings file leaked onto the Equity sheet;
3. demat-held funds (broker platforms) had no import route at all;
4. By Scrip never auto-filled, so imported scrips were missing there.
"""

from datetime import date, timedelta

import pytest

from networth import model as M
from networth.importers.common import (ImportBatch, ImportedHolding,
                                       ImportedTrade)
from networth.importers.merge import (_isin_class, merge_equity_batches,
                                      merge_fund_holdings)
from networth.update import _dividend_qty, sync_by_scrip

TODAY = date(2026, 7, 19)
RELIANCE = "INE002A01018"
INFY = "INE009A01021"
FUND = "INF204K01UN9"           # a fund unit (INF prefix)
ETF = "INF204KB14I2"            # an ETF: INF prefix but listed = equity
NCD = "INE342T07718"            # series 07 = debt, not equity


def _data(persons=("Amit",)):
    d = M.PortfolioData(persons=list(persons))
    d.masters.stock_rows = [("RELIANCE", "RELIANCE INDUSTRIES", RELIANCE),
                            ("INFY", "INFOSYS", INFY),
                            ("SILVERBEES", "NIPPON SILVER ETF", ETF)]
    d.masters.mf_rows = [("Nippon", "Nippon India Liquid Fund", FUND)]
    return d


def _holding(qty=45.0, avg=137.30, isin=RELIANCE, name="RELIANCE",
             account="ZR001"):
    return ImportedHolding(account=account, isin=isin, name=name,
                           qty=qty, avg_cost=avg)


def _split(ex, isin=RELIANCE, frm=10, to=2):
    # face 10 → 2 = 5 new shares per old one
    return M.CorporateAction(symbol="X", isin=isin, type="SPLIT",
                             ex_date=ex, ratio_from=frm, ratio_to=to,
                             source="Manual")


OWNER_MAP = {"ZR001": "Amit"}


# ---- 1: the corporate-action anchor -----------------------------------------

def test_holdings_row_is_anchored_and_history_never_reapplies():
    d = _data()
    d.corporate_actions = [_split(date(2015, 3, 1))]     # long before import
    b = ImportBatch(source="broker holdings", holdings=[_holding()])
    rep = merge_equity_batches(d, [b], OWNER_MAP, TODAY)
    assert rep.eq_added == 1
    row = d.equity[0]
    assert row.qty_asof == TODAY
    # the window the updater stamps Adj factor from starts at the anchor:
    # the 2015 split is behind it, so the factor is exactly 1
    f = M.chained_adjustment_factor(RELIANCE, M.qty_anchor(row), TODAY,
                                    d.corporate_actions)
    assert f == pytest.approx(1.0)


def test_action_after_the_import_still_adjusts_the_anchored_row():
    d = _data()
    b = ImportBatch(source="broker holdings", holdings=[_holding()])
    merge_equity_batches(d, [b], OWNER_MAP, TODAY)
    row = d.equity[0]
    later = TODAY + timedelta(days=30)
    actions = [_split(TODAY + timedelta(days=10))]
    f = M.chained_adjustment_factor(RELIANCE, M.qty_anchor(row), later,
                                    actions)
    assert f == pytest.approx(5.0)


def test_typed_rows_keep_the_cost_date_window():
    r = M.EquityRow(owner="Amit", scrip="RELIANCE INDUSTRIES", qty=10,
                    avg_cost=900.0, cost_date=date(2014, 1, 1))
    f = M.chained_adjustment_factor(RELIANCE, M.qty_anchor(r), TODAY,
                                    [_split(date(2015, 3, 1))])
    assert f == pytest.approx(5.0)


def test_demerger_child_not_spawned_when_broker_already_lists_it():
    from networth.compute.restructures import apply_demergers
    child_isin = "INE0CHILD012"
    ev = M.CorporateAction(symbol="RELIANCE", isin=RELIANCE, type="DEMERGER",
                           ex_date=date(2026, 4, 30), ratio_from=1,
                           ratio_to=1, source="Curated", new_isin=child_isin,
                           new_name="CHILDCO LTD.", cost_pct=40)
    d = _data()
    d.corporate_actions = [ev]
    d.equity = [M.EquityRow(owner="Amit", scrip="RELIANCE INDUSTRIES",
                            qty=45, avg_cost=137.3, qty_asof=TODAY,
                            flag="IMPORTED:broker holdings")]
    added, _w = apply_demergers(d, [ev], ca_checked=set(), ca_trusted=True,
                                price_data=None, today=TODAY)
    assert added == 0 and len(d.equity) == 1   # the broker file has CHILDCO

    # ...but a demerger AFTER the import date must still spawn the child
    ev2 = M.CorporateAction(symbol="RELIANCE", isin=RELIANCE,
                            type="DEMERGER", ex_date=TODAY + timedelta(days=5),
                            ratio_from=1, ratio_to=1, source="Curated",
                            new_isin=child_isin, new_name="CHILDCO LTD.",
                            cost_pct=40)
    d.corporate_actions = [ev2]
    added, _w = apply_demergers(d, [ev2], ca_checked=set(), ca_trusted=True,
                                price_data=None,
                                today=TODAY + timedelta(days=10))
    assert added == 1


def test_dividend_estimate_divides_anchored_rows_back():
    # 45 post-split shares anchored today; the dividend's ex-date precedes
    # the split, when the holder had 45/5 = 9 shares
    rows = [M.EquityRow(owner="Amit", scrip="RELIANCE INDUSTRIES", qty=45,
                        avg_cost=137.3, qty_asof=TODAY)]
    actions = [_split(date(2026, 6, 1))]
    n = _dividend_qty("Amit", RELIANCE, date(2026, 5, 1), rows,
                      {"RELIANCE INDUSTRIES": RELIANCE}, actions)
    assert n == pytest.approx(9.0)


def test_holdings_cross_check_compares_in_todays_terms():
    # a typed pre-split lot of 9 shares IS the broker's 45 — no warning
    d = _data()
    d.corporate_actions = [_split(date(2015, 3, 1))]
    d.equity = [M.EquityRow(owner="Amit", scrip="RELIANCE INDUSTRIES",
                            qty=9, avg_cost=680.0,
                            cost_date=date(2014, 1, 1))]
    b = ImportBatch(source="broker holdings", holdings=[_holding(qty=45.0)])
    rep = merge_equity_batches(d, [b], OWNER_MAP, TODAY)
    assert rep.eq_added == 0
    assert not any("check for missing history" in w or "sheet now holds" in w
                   for w in rep.warnings)


# ---- 2: the ISIN classification gate ----------------------------------------

def test_isin_classes():
    stock = {RELIANCE: "RELIANCE INDUSTRIES", ETF: "NIPPON SILVER ETF",
             NCD: "NFL-10.30%-31-3-28-PVT"}
    assert _isin_class(RELIANCE, stock) == "equity"
    assert _isin_class(ETF, stock) == "equity"        # listed ETF = equity
    assert _isin_class(FUND, stock) == "fund"
    # the bhavcopy (and so the master) lists traded NCDs — the series
    # digits must outrank master membership, or bonds import as shares
    assert _isin_class(NCD, stock) == "debt"
    assert _isin_class("US0378331005", stock) == "other"


def test_ncd_in_holdings_file_is_refused_to_the_bonds_sheet():
    d = _data()
    b = ImportBatch(source="broker holdings", holdings=[
        _holding(qty=10, avg=1000.0, isin=NCD, name="NFL-10.30%-31-3-28")])
    rep = merge_equity_batches(d, [b], OWNER_MAP, TODAY)
    assert rep.eq_added == 0 and not d.equity
    assert any("bond/debenture" in w and "Bonds sheet" in w
               for w in rep.warnings)


def test_fund_rows_never_become_equity_rows():
    d = _data()
    b = ImportBatch(source="broker holdings", holdings=[
        _holding(qty=100, avg=25.0, isin=FUND, name="Nippon Liquid")])
    rep = merge_equity_batches(d, [b], OWNER_MAP, TODAY)
    assert rep.eq_added == 0 and not d.equity
    # and not with noise either: merge_fund_holdings owns the reporting
    assert not rep.warnings


def test_listed_etf_imports_as_equity():
    d = _data()
    b = ImportBatch(source="broker holdings", holdings=[
        _holding(qty=2350, avg=218.22, isin=ETF, name="SILVERBEES")])
    rep = merge_equity_batches(d, [b], OWNER_MAP, TODAY)
    assert rep.eq_added == 1
    assert d.equity[0].scrip == "NIPPON SILVER ETF"


def test_fund_trades_are_refused_with_the_cas_pointer():
    d = _data()
    b = ImportBatch(source="broker tradebook", trades=[
        ImportedTrade(account="ZR001", isin=FUND, symbol="LIQUIDBEES",
                      trade_date=date(2026, 5, 1), qty=10, price=1000.0,
                      side="BUY")])
    rep = merge_equity_batches(d, [b], OWNER_MAP, TODAY)
    assert rep.eq_added == 0 and not d.equity
    line = rep.stocks[0]
    assert not line.ok and "fund units" in line.reason


# ---- 3: fund holdings → one opening MF_SIP line -----------------------------

def test_fund_holding_opens_one_dated_line_and_is_idempotent():
    d = _data()
    b = ImportBatch(source="broker holdings", holdings=[
        _holding(qty=3499.805, avg=10.0, isin=FUND, name="Nippon Liq")])
    rep = merge_fund_holdings(d, [b], OWNER_MAP, TODAY)
    assert rep.sip_added == 1 and len(d.sip) == 1
    row = d.sip[0]
    assert row.scheme == "Nippon India Liquid Fund"    # master name wins
    assert row.txn_date == TODAY
    assert row.units_override == pytest.approx(3499.805)
    assert row.amount == pytest.approx(34998.05)
    assert row.nav == pytest.approx(10.0)
    assert ("Amit", "Nippon India Liquid Fund") in {
        (r.owner, r.scheme) for r in d.mutual_funds}
    # second run: cross-check only, adds nothing, raises no mismatch
    rep2 = merge_fund_holdings(d, [b], OWNER_MAP, TODAY)
    assert rep2.sip_added == 0 and len(d.sip) == 1
    assert not any("check for missing history" in w for w in rep2.warnings)


def test_fund_holding_mismatch_warns_and_never_doubles():
    d = _data()
    d.sip = [M.SIPRow(owner="Amit", scheme="Nippon India Liquid Fund",
                      txn_date=date(2024, 1, 1), amount=10000.0, nav=10.0,
                      units_override=1000.0)]
    b = ImportBatch(source="broker holdings", holdings=[
        _holding(qty=3499.805, avg=10.0, isin=FUND, name="Nippon Liq")])
    rep = merge_fund_holdings(d, [b], OWNER_MAP, TODAY)
    assert rep.sip_added == 0 and len(d.sip) == 1
    assert any("3499.8" in w and "1000.000" in w for w in rep.warnings)


def test_fund_holding_without_avg_cost_is_refused():
    d = _data()
    b = ImportBatch(source="broker holdings", holdings=[
        ImportedHolding(account="ZR001", isin=FUND, name="Nippon Liq",
                        qty=100.0, avg_cost=None)])
    rep = merge_fund_holdings(d, [b], OWNER_MAP, TODAY)
    assert rep.sip_added == 0 and not d.sip
    assert not rep.funds[0].ok
    assert "average cost" in rep.funds[0].reason


def test_fund_holding_unmapped_account_reasks():
    d = _data()
    b = ImportBatch(source="broker holdings", holdings=[
        _holding(qty=100, avg=25.0, isin=FUND, name="Nippon Liq",
                 account="UNKNOWN")])
    rep = merge_fund_holdings(d, [b], {}, TODAY)
    assert rep.sip_added == 0 and not d.sip
    assert "not matched to a person" in rep.funds[0].reason


def test_fund_holding_unknown_isin_keeps_name_and_override():
    d = _data()
    other = "INF999X01999"
    b = ImportBatch(source="broker holdings", holdings=[
        _holding(qty=50, avg=100.0, isin=other, name="Brand New Fund Gr")])
    rep = merge_fund_holdings(d, [b], OWNER_MAP, TODAY)
    assert rep.sip_added == 1
    assert d.sip[0].scheme == "Brand New Fund Gr"
    assert d.sip[0].isin_override == other


def test_fund_holdings_capacity_defers_untouched(monkeypatch):
    monkeypatch.setattr(M, "SIP_LAST_ROW", M.FIRST_DATA_ROW - 1)
    d = _data()
    b = ImportBatch(source="broker holdings", holdings=[
        _holding(qty=100, avg=25.0, isin=FUND, name="Nippon Liq")])
    rep = merge_fund_holdings(d, [b], OWNER_MAP, TODAY)
    assert rep.deferred and not d.sip and not d.mutual_funds
    assert rep.sip_added == 0


# ---- 4: By Scrip auto-sync --------------------------------------------------

def test_by_scrip_fills_missing_held_isins_add_only():
    d = _data()
    d.equity = [
        M.EquityRow(owner="Amit", scrip="RELIANCE INDUSTRIES", qty=10),
        M.EquityRow(owner="Amit", scrip="INFOSYS", qty=5),
        M.EquityRow(owner="Amit", scrip="", isin_override="INE0ABC01234",
                    qty=3),
        M.EquityRow(owner="Amit", scrip="INFOSYS", qty=None),  # empty row
    ]
    d.by_scrip = [M.ScripRef(isin=RELIANCE, name="My own label")]
    warn = sync_by_scrip(d, {"RELIANCE INDUSTRIES": RELIANCE,
                             "INFOSYS": INFY})
    assert warn is None
    assert [r.isin for r in d.by_scrip] == [RELIANCE, "INE0ABC01234", INFY]
    assert d.by_scrip[0].name == "My own label"        # user row untouched
    assert d.by_scrip[2].name == "INFOSYS"             # master display name


def test_run_end_to_end_anchor_and_by_scrip(tmp_path):
    """Real-file scenario through run(): a stock with a historical
    split, imported from a holdings file — Adj factor must stay blank
    (broker qty is already post-split), the anchor must round-trip, and
    By Scrip must gain every held ISIN."""
    from networth.fetch.amfi import AmfiData
    from networth.generate import build_workbook
    from networth.reader import read_workbook
    from networth.sample_data import sample_portfolio
    from networth.update import run

    d = sample_portfolio()
    d.corporate_actions.append(_split(date(2020, 6, 1)))   # Manual, persists
    path = tmp_path / "wb.xlsx"
    build_workbook(d, str(path))
    b = ImportBatch(source="broker holdings", holdings=[
        _holding(qty=45.0, avg=137.30, account="Z1")])
    b.fingerprint = "beef12345678"
    run(path, price_data=None, amfi_data=AmfiData(), ca_data=[],
        div_data=[], today=TODAY, import_batches=[b],
        import_owner_map={"Z1": "Priya"},
        import_decisions=[("holdings.csv", "beef12345678", "imported")])
    back = read_workbook(str(path))
    row = next(r for r in back.equity
               if r.owner == "Priya" and "IMPORTED" in r.flag)
    assert row.qty == pytest.approx(45.0)
    assert row.qty_asof == TODAY               # the anchor round-trips
    assert row.ca_factor is None               # 2020 split NOT re-applied
    # the typed sample rows keep the split: cost-dated before the ex-date
    typed = next(r for r in back.equity
                 if r.owner == "Amit" and r.scrip.startswith("RELIANCE"))
    assert typed.ca_factor == pytest.approx(5.0)
    # By Scrip auto-sync: every held ISIN now has a row
    isin_by_name = {n: i for _s, n, i in back.masters.stock_rows}
    held = {r.isin_override or isin_by_name.get(r.scrip, "")
            for r in back.equity if r.qty}
    held.discard("")
    assert held <= {r.isin for r in back.by_scrip}


def test_by_scrip_overflow_warns_instead_of_truncating_silently(monkeypatch):
    monkeypatch.setattr(M, "BYSCRIP_LAST_ROW", M.FIRST_DATA_ROW)  # 1 row
    d = _data()
    d.equity = [
        M.EquityRow(owner="Amit", scrip="RELIANCE INDUSTRIES", qty=10),
        M.EquityRow(owner="Amit", scrip="INFOSYS", qty=5),
    ]
    warn = sync_by_scrip(d, {"RELIANCE INDUSTRIES": RELIANCE,
                             "INFOSYS": INFY})
    assert warn and "By Scrip sheet is full" in warn
    assert len(d.by_scrip) == 1
