"""v1.7.4 — naming hand-entered restructures, NSE resilience, unlisted shares.

All three came from user reports (2026-08-09):

1. a hand-entered demerger had no way to NAME the new company, so the row
   (and Stock_Master) were labelled with the raw ISIN — and because the
   master is add-only, that placeholder was permanent;
2. an NSE-only holding (Nippon India Silver ETF, SILVERBEES) never got a
   fresh price: one refused NSE request dropped the whole day's NSE data;
3. unlisted/pre-IPO shares had no honest place to carry a price.
"""

from datetime import date

import pytest

from networth import model as M
from networth.compute.restructures import apply_demergers
from networth.fetch import bhavcopy
from networth.model import load_restructures
from networth.update import _merge_stock_master

TODAY = date(2026, 8, 9)
CHILD = "INE0CHILD012"
PARENT = "INE002A01018"
TATA = "INE155A01022"
TMCV = "INE1TAE01010"


# ---- 1: naming a hand-entered restructure ----------------------------------

def _manual_demerger(new_name="", new_symbol=""):
    return [
        M.CorporateAction(symbol="RELIANCE", isin=PARENT, type="DEMERGER",
                          ex_date=date(2026, 5, 1), ratio_from=1, ratio_to=1,
                          source="Manual", new_isin=PARENT, cost_pct=70),
        M.CorporateAction(symbol="RELIANCE", isin=PARENT, type="DEMERGER",
                          ex_date=date(2026, 5, 1), ratio_from=1, ratio_to=1,
                          source="Manual", new_isin=CHILD, cost_pct=30,
                          new_name=new_name, new_symbol=new_symbol),
    ]


def _data_with_lot():
    d = M.PortfolioData(persons=["Jay"])
    d.masters.stock_rows = [("RELIANCE", "RELIANCE INDUSTRIES", PARENT)]
    d.equity = [M.EquityRow(owner="Jay", scrip="RELIANCE INDUSTRIES", qty=10,
                            avg_cost=1000.0, cost_date=date(2020, 1, 1))]
    return d


def test_child_row_uses_the_typed_new_name():
    d = _data_with_lot()
    evs = _manual_demerger(new_name="CHILDCO LTD.", new_symbol="CHILDCO")
    d.corporate_actions = list(evs)
    added, _w = apply_demergers(d, evs, ca_checked={PARENT}, ca_trusted=False,
                                price_data=None, today=TODAY)
    assert added == 1
    assert d.equity[1].scrip == "CHILDCO LTD."


def test_symbol_is_used_when_only_it_is_given():
    d = _data_with_lot()
    evs = _manual_demerger(new_symbol="CHILDCO")
    d.corporate_actions = list(evs)
    apply_demergers(d, evs, ca_checked={PARENT}, ca_trusted=False,
                    price_data=None, today=TODAY)
    assert d.equity[1].scrip == "CHILDCO"


def test_isin_placeholder_is_upgraded_when_the_feed_knows_the_name():
    # a restructure seeded the master before the security listed
    existing = [("RELIANCE", "RELIANCE INDUSTRIES", PARENT),
                (CHILD, CHILD, CHILD)]
    fetched = [("CHILDCO", "CHILDCO LTD.", CHILD)]
    rows, added, renamed = _merge_stock_master(existing, fetched)
    assert added == 0                      # the ISIN was already known
    assert renamed == {CHILD: "CHILDCO LTD."}
    assert (dict((i, n) for _s, n, i in rows))[CHILD] == "CHILDCO LTD."
    assert sorted(rows, key=lambda r: r[1].casefold()) == rows   # still sorted


def test_a_real_name_is_never_overwritten_by_the_feed():
    # add-only still protects names a user row may point at (SPEC §6.4)
    existing = [("TATAMOTORS", "TATA MOTORS LTD.", TATA)]
    fetched = [("TMPV", "TATA MOTORS PASSENGER VEHICLES LTD", TATA)]
    rows, added, renamed = _merge_stock_master(existing, fetched)
    assert added == 0 and renamed == {}
    assert rows == existing


# ---- 2: NSE resilience -----------------------------------------------------

CSV = ("TckrSymb,ISIN,ClsPric,PrvsClsgPric,FinInstrmNm\n"
       "SILVERBEES,INF204KC1402,102.5,101.0,NIPPON INDIA SILVER ETF\n")


class _Resp:
    def __init__(self, status=200, content=b"", text=""):
        self.status_code, self.content, self.text = status, content, text


def _zipped(csv_text: str) -> bytes:
    import io
    import zipfile
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        z.writestr("bhav.csv", csv_text)
    return buf.getvalue()


class _Session:
    """Refuses the archive until it has been 'cookied' by a warm-up page,
    which is exactly how NSE's edge behaves."""

    def __init__(self, *, refusals=0, only_legacy=False):
        self.refusals, self.only_legacy = refusals, only_legacy
        self.calls: list[str] = []

    def get(self, url, headers=None, timeout=None):
        self.calls.append(url)
        if "nseindia.com/all-reports" in url or url.endswith("nseindia.com/"):
            return _Resp(200, b"<html>ok</html>", "<html>ok</html>")
        legacy = "content/historical/EQUITIES" in url
        if self.only_legacy and not legacy:
            return _Resp(404)
        if self.refusals > 0:
            self.refusals -= 1
            html = "<html>Access Denied</html>"
            return _Resp(200, html.encode(), html)      # bot-challenge page
        return _Resp(200, _zipped(CSV))


def test_nse_retries_after_a_refusal_instead_of_losing_the_day():
    sess = _Session(refusals=2)
    text = bhavcopy._get_nse(sess, date(2026, 8, 7), 10)
    assert text and "INF204KC1402" in text
    assert any("all-reports" in u for u in sess.calls)   # cookie warm-up ran


def test_nse_falls_back_to_the_legacy_archive():
    sess = _Session(only_legacy=True)
    text = bhavcopy._get_nse(sess, date(2026, 8, 7), 10)
    assert text and "SILVERBEES" in text
    assert any("content/historical/EQUITIES" in u for u in sess.calls)


def test_nse_gives_up_quietly_when_every_attempt_is_refused():
    sess = _Session(refusals=99)
    assert bhavcopy._get_nse(sess, date(2026, 8, 7), 10) is None


def test_an_nse_only_etf_prices_from_the_nse_bhavcopy():
    # the ETF's INF-prefixed ISIN must not disqualify it as an equity price
    p = bhavcopy.parse(CSV)
    assert p.prices["INF204KC1402"]["close"] == pytest.approx(102.5)
    assert ("SILVERBEES", "NIPPON INDIA SILVER ETF",
            "INF204KC1402") in p.master_rows


# ---- 3: unlisted shares ----------------------------------------------------

def test_unlisted_price_values_the_holding_and_yields_to_the_market():
    unlisted = M.EquityRow(owner="Jay", scrip="SOME PRIVATE LTD", qty=100,
                           avg_cost=50.0, cost_date=date(2024, 1, 1),
                           manual_price=80.0)
    assert M.effective_price(unlisted) == pytest.approx(80.0)
    # the day it lists, the exchange close takes over with no edit at all
    unlisted.close = 95.0
    assert M.effective_price(unlisted) == pytest.approx(95.0)


def test_unlisted_holding_counts_in_net_worth_and_returns():
    from networth.compute.cashflows import equity_flows
    from networth.compute.snapshot import net_worth_snapshot
    d = M.PortfolioData(persons=["Jay"])
    d.masters.stock_rows = []
    d.equity = [M.EquityRow(owner="Jay", scrip="SOME PRIVATE LTD", qty=100,
                            avg_cost=50.0, cost_date=date(2024, 1, 1),
                            manual_price=80.0)]
    snap = net_worth_snapshot(d, TODAY)
    assert snap.equity == pytest.approx(8000.0)
    flows = equity_flows(d, TODAY)
    assert flows[0][1] == pytest.approx(-5000.0)     # invested
    assert flows[1][1] == pytest.approx(8000.0)      # worth today


def test_unlisted_row_is_never_flagged_delisted():
    # no close date ever arrived, so there is nothing to escalate from
    r = M.EquityRow(owner="Jay", scrip="SOME PRIVATE LTD", qty=10,
                    avg_cost=10.0, manual_price=25.0)
    assert r.close_date is None and M.effective_price(r) == 25.0


def test_unlisted_price_round_trips(tmp_path):
    from networth.generate import build_workbook
    from networth.reader import read_workbook
    from networth.sample_data import sample_portfolio
    d = sample_portfolio()
    d.equity.append(M.EquityRow(owner="Amit", scrip="SOME PRIVATE LTD",
                                qty=100, avg_cost=50.0,
                                cost_date=date(2024, 1, 1),
                                manual_price=80.0))
    p = tmp_path / "wb.xlsx"
    build_workbook(d, str(p))
    back = read_workbook(str(p))
    row = next(r for r in back.equity if r.scrip == "SOME PRIVATE LTD")
    assert row.manual_price == pytest.approx(80.0)
    assert row.close is None


# ---- curated data: the Tata Motors demerger --------------------------------

def test_shipped_file_carries_the_tata_motors_demerger():
    evs = [e for e in load_restructures() if e.isin == TATA]
    assert len(evs) == 2
    by_isin = {e.new_isin: e for e in evs}
    assert by_isin[TATA].cost_pct == pytest.approx(68.85)      # retained PV
    assert by_isin[TMCV].cost_pct == pytest.approx(31.15)      # CV company
    for e in evs:
        assert e.ex_date == date(2025, 10, 14)
        assert e.factor() == 1.0 and (e.ratio_from, e.ratio_to) == (1, 1)
        assert e.new_name                       # curated rows always name it
