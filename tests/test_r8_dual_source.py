"""R8: BSE + NSE as full peer price sources (SPEC §5.2/5.3, §6.5 guard)."""

import io
import zipfile
from datetime import date, timedelta

import pytest

from networth.fetch.amfi import AmfiData
from networth.fetch.bhavcopy import PriceData, fetch, merge, parse
from networth.generate import build_workbook
from networth.reader import read_workbook
from networth.sample_data import sample_portfolio
from networth.update import run

TODAY = date(2026, 7, 15)

# ISIN X: BSE-only (with scrip code). Y: NSE-only. Z: dual-listed, closes differ.
BSE_CSV = """TradDt,Sgmt,FinInstrmId,ISIN,TckrSymb,FinInstrmNm,ClsPric,PrvsClsgPric
2026-07-15,CM,500001,INE000X01011,XONBSE,X ONLY ON BSE LTD.,100.0,99.0
2026-07-15,CM,500003,INE000Z01013,ZBOTH,Z DUAL LISTED LTD.,200.0,198.0
"""
NSE_CSV = """TradDt,Sgmt,FinInstrmId,ISIN,TckrSymb,FinInstrmNm,ClsPric,PrvsClsgPric
2026-07-15,CM,99901,INE000Y01012,YONNSE,Y ONLY ON NSE LTD.,150.0,149.0
2026-07-15,CM,99903,INE000Z01013,ZBOTHN,Z DUAL LISTED LTD.,201.5,199.5
"""


def test_merge_union_nse_wins_bse_keeps_codes():
    out = merge(parse(BSE_CSV), parse(NSE_CSV))
    # union of ISINs
    assert set(out.prices) == {"INE000X01011", "INE000Y01012", "INE000Z01013"}
    # dual-listed conflict: NSE close/prev win
    assert out.prices["INE000Z01013"] == {"close": 201.5, "prev": 199.5}
    # single-listed rows keep their own exchange's quote
    assert out.prices["INE000X01011"]["close"] == 100.0
    assert out.prices["INE000Y01012"]["close"] == 150.0
    # scrip codes come from BSE only, retained for every BSE-quoted ISIN
    assert out.codes_by_isin == {"INE000X01011": "500001",
                                 "INE000Z01013": "500003"}
    # master rows: deduped by ISIN, NSE symbol preferred for the dual listing
    by_isin = {isin: sym for sym, _n, isin in out.master_rows}
    assert by_isin["INE000Z01013"] == "ZBOTHN"
    assert by_isin["INE000X01011"] == "XONBSE"
    assert out.source == "BSE+NSE" and out.sources == ["BSE", "NSE"]
    assert out.nse_only == {"INE000Y01012"}


def test_merge_single_source_runs():
    bse_only = merge(parse(BSE_CSV), None)
    assert bse_only.source == "BSE" and bse_only.sources == ["BSE"]
    assert bse_only.codes_by_isin  # kept
    nse_only = merge(None, parse(NSE_CSV))
    assert nse_only.source == "NSE" and nse_only.sources == ["NSE"]
    assert nse_only.codes_by_isin == {}  # NSE ids are not BSE scrip codes


class _Resp:
    def __init__(self, status=200, text="", content=b""):
        self.status_code = status
        self.text = text
        self.content = content


def _nse_zip(csv_text: str) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        z.writestr("bhav.csv", csv_text)
    return buf.getvalue()


class _FakeSession:
    """Routes by URL substring; unrouted URLs get an empty 404."""

    def __init__(self, routes: dict):
        self.routes = routes

    def get(self, url, **_kw):
        for frag, resp in self.routes.items():
            if frag in url:
                return resp
        return _Resp(status=404)


def test_fetch_merges_both_exchanges_same_day():
    ymd = TODAY.strftime("%Y%m%d")
    sess = _FakeSession({
        f"BhavCopy_BSE_CM_0_0_0_{ymd}": _Resp(text=BSE_CSV),
        f"BhavCopy_NSE_CM_0_0_0_{ymd}": _Resp(content=_nse_zip(NSE_CSV)),
        "www.nseindia.com": _Resp(text="warmup"),
    })
    out = fetch(session=sess, today=TODAY)
    assert out.trade_date == TODAY
    assert out.source == "BSE+NSE"
    assert set(out.prices) == {"INE000X01011", "INE000Y01012", "INE000Z01013"}


def test_fetch_proceeds_single_source_when_nse_down():
    ymd = TODAY.strftime("%Y%m%d")
    sess = _FakeSession({
        f"BhavCopy_BSE_CM_0_0_0_{ymd}": _Resp(text=BSE_CSV),
        "www.nseindia.com": _Resp(text="warmup"),
        # no NSE bhavcopy route → 404
    })
    out = fetch(session=sess, today=TODAY)
    assert out.source == "BSE" and out.sources == ["BSE"]
    assert "INE000Y01012" not in out.prices


def test_fetch_walks_back_and_never_mixes_dates():
    # today: only NSE published; yesterday: both. The walk must STOP at today
    # with a single-source NSE result, not pair today's NSE with yesterday's BSE.
    ymd0 = TODAY.strftime("%Y%m%d")
    ymd1 = (TODAY - timedelta(days=1)).strftime("%Y%m%d")
    sess = _FakeSession({
        f"BhavCopy_NSE_CM_0_0_0_{ymd0}": _Resp(content=_nse_zip(NSE_CSV)),
        f"BhavCopy_BSE_CM_0_0_0_{ymd1}": _Resp(text=BSE_CSV),
        f"BhavCopy_NSE_CM_0_0_0_{ymd1}": _Resp(content=_nse_zip(NSE_CSV)),
        "www.nseindia.com": _Resp(text="warmup"),
    })
    out = fetch(session=sess, today=TODAY)
    assert out.trade_date == TODAY and out.source == "NSE"

    # with nothing published today, the walk-back lands on yesterday, merged
    sess2 = _FakeSession({
        f"BhavCopy_BSE_CM_0_0_0_{ymd1}": _Resp(text=BSE_CSV),
        f"BhavCopy_NSE_CM_0_0_0_{ymd1}": _Resp(content=_nse_zip(NSE_CSV)),
        "www.nseindia.com": _Resp(text="warmup"),
    })
    out2 = fetch(session=sess2, today=TODAY)
    assert out2.trade_date == TODAY - timedelta(days=1)
    assert out2.source == "BSE+NSE"


def test_fetch_raises_when_both_down_for_a_week():
    sess = _FakeSession({"www.nseindia.com": _Resp(text="warmup")})
    with pytest.raises(RuntimeError):
        fetch(session=sess, today=TODAY)


def test_single_source_run_never_escalates_status(tmp_path):
    """A one-exchange outage must not mark quiet scrips Suspended (SPEC §6.5)."""
    data = sample_portfolio()
    path = tmp_path / "wb.xlsx"
    build_workbook(data, str(path))
    isin_by_name = {n: i for _s, n, i in data.masters.stock_rows}
    wipro = isin_by_name["WIPRO LTD."]

    # round 1 (dual source): everyone quoted → Active
    prices = PriceData(trade_date=TODAY, source="BSE+NSE", sources=["BSE", "NSE"])
    for row in data.equity:
        prices.prices[isin_by_name[row.scrip]] = {"close": 100.0, "prev": 99.0}
    run(path, price_data=prices, amfi_data=AmfiData(), ca_data=[], today=TODAY)

    # round 2, 30 days on, WIPRO absent — but single-source: NO escalation
    later = TODAY + timedelta(days=30)
    prices2 = PriceData(trade_date=later, source="BSE", sources=["BSE"])
    for row in data.equity:
        if row.scrip != "WIPRO LTD.":
            prices2.prices[isin_by_name[row.scrip]] = {"close": 101.0, "prev": 100.0}
    summary = run(path, price_data=prices2, amfi_data=AmfiData(), ca_data=[], today=later)
    assert summary["suspended"] == 0
    back = read_workbook(str(path))
    st, last = back.masters.stock_status[wipro]
    assert st == "Active" and last == TODAY   # carried forward untouched

    # round 3, same absence but dual-source: escalation resumes
    prices3 = PriceData(trade_date=later, source="BSE+NSE", sources=["BSE", "NSE"])
    for row in data.equity:
        if row.scrip != "WIPRO LTD.":
            prices3.prices[isin_by_name[row.scrip]] = {"close": 102.0, "prev": 101.0}
    summary3 = run(path, price_data=prices3, amfi_data=AmfiData(), ca_data=[], today=later)
    assert summary3["suspended"] == 1
    back3 = read_workbook(str(path))
    assert back3.masters.stock_status[wipro][0] == "Suspended"


def test_nse_only_count_reaches_summary(tmp_path):
    data = sample_portfolio()
    path = tmp_path / "wb.xlsx"
    build_workbook(data, str(path))
    isin_by_name = {n: i for _s, n, i in data.masters.stock_rows}
    prices = PriceData(trade_date=TODAY, source="BSE+NSE", sources=["BSE", "NSE"])
    isins = [isin_by_name[row.scrip] for row in data.equity]
    for i in isins:
        prices.prices[i] = {"close": 100.0, "prev": 99.0}
    prices.nse_only = {isins[0]}
    summary = run(path, price_data=prices, amfi_data=AmfiData(), ca_data=[], today=TODAY)
    assert summary["nse_only_matched"] == 1
    assert summary["price_source"].startswith("BSE+NSE")
