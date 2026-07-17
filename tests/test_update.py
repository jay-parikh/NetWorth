"""End-to-end updater run with injected data sources (SPEC §7)."""

from datetime import date

import pytest
from openpyxl import load_workbook

from networth.fetch.amfi import AmfiData
from networth.fetch.bhavcopy import PriceData
from networth.generate import build_workbook
from networth.reader import read_workbook
from networth.sample_data import sample_portfolio
from networth.update import run

TODAY = date(2026, 7, 15)


@pytest.fixture()
def workbook(tmp_path, monkeypatch):
    path = tmp_path / "Family_Portfolio_Tracker.xlsx"
    build_workbook(sample_portfolio(), str(path))
    return path


def _fake_sources():
    data = sample_portfolio()
    isin_by_name = {n: i for _s, n, i in data.masters.stock_rows}
    prices = PriceData(trade_date=TODAY, source="TEST")
    for row in data.equity:
        isin = isin_by_name[row.scrip]
        prices.prices[isin] = {"close": round(row.close * 1.02, 2),
                               "prev": row.close}
    prices.master_rows = [("NEWCO", "NEWLY LISTED CO", "INE0NEW00001")]
    prices.prices["INE906B07CB9"] = {"close": 1020.0, "prev": 1015.0}  # the bond

    isin_by_scheme = {s: i for _f, s, i in data.masters.mf_rows}
    amfi = AmfiData(nav_date="15-Jul-2026")
    for m in data.mutual_funds:
        amfi.nav_by_isin[isin_by_scheme[m.scheme]] = round(m.current_nav * 1.01, 2)
    amfi.master_rows = [(f, s, i) for f, s, i in
                        (("PPFAS Mutual Fund", m.scheme, isin_by_scheme[m.scheme])
                         for m in data.mutual_funds)]
    return prices, amfi


def test_update_run(workbook):
    prices, amfi = _fake_sources()
    summary = run(workbook, price_data=prices, amfi_data=amfi, ca_data=[],
                  today=TODAY)

    assert summary["equity_matched"] == 10
    assert summary["stocks_added"] == 1
    assert summary["bonds_matched"] == 1
    assert summary["mf_matched"] == 2
    # the only line is the standing hidden-money awareness note (v1.4.3) —
    # the shipped sample keeps its off-classes' rows inside hidden tabs
    assert [w for w in summary["warnings"]
            if not w.startswith("hidden and not counted")] == []
    assert (workbook.parent / "backups").exists()

    back = read_workbook(str(workbook))
    # prices moved +2%, prev = old close, date stamped
    assert back.equity[0].close == pytest.approx(1520 * 1.02, abs=0.01)
    assert back.equity[0].prev_close == 1520
    assert back.equity[0].close_date == TODAY
    # bond priced from the feed
    assert back.bonds[0].cur_price == 1020.0
    # stock master merged add-only: new listing present, old names intact
    isins = {i for _s, _n, i in back.masters.stock_rows}
    assert "INE0NEW00001" in isins and "INE002A01018" in isins
    # MF master preserved referenced schemes even though the fake AMFI list
    # dropped one of them
    schemes = {s for _f, s, i in back.masters.mf_rows}
    assert "SBI Large Cap FUND-DIRECT PLAN -GROWTH" in schemes
    # XIRR recomputed and plausible
    assert back.xirr.portfolio is not None
    assert -0.5 < back.xirr.portfolio < 1.5
    assert back.xirr.fixed_deposits == pytest.approx(0.0721, abs=0.01)
    # workbook still structurally intact (charts survive the regenerate)
    import zipfile
    with zipfile.ZipFile(workbook) as z:
        assert sum(1 for n in z.namelist() if n.startswith("xl/charts/chart")) == 10


def test_update_degrades_gracefully(workbook):
    class Boom:
        def __getattr__(self, _):
            raise RuntimeError("network down")

    # both sources fail: values stay, warnings recorded, file still regenerated
    import networth.update as U
    orig = (U.bhav_mod.fetch, U.amfi_mod.fetch, U.ca_mod.fetch)
    U.bhav_mod.fetch = lambda **kw: (_ for _ in ()).throw(RuntimeError("bhav down"))
    U.amfi_mod.fetch = lambda **kw: (_ for _ in ()).throw(RuntimeError("amfi down"))
    U.ca_mod.fetch = lambda *a, **kw: (_ for _ in ()).throw(RuntimeError("ca down"))
    try:
        summary = run(workbook, today=TODAY)
    finally:
        U.bhav_mod.fetch, U.amfi_mod.fetch, U.ca_mod.fetch = orig

    failures = [w for w in summary["warnings"]
                if not w.startswith("hidden and not counted")]
    assert len(failures) == 3
    back = read_workbook(str(workbook))
    assert back.equity[0].close == 1520          # unchanged
    assert back.xirr.portfolio is not None        # XIRR still recomputed


def test_update_refuses_open_file(workbook, tmp_path):
    lock = workbook.with_name("~$" + workbook.name)
    lock.write_bytes(b"")
    with pytest.raises(SystemExit):
        run(workbook, price_data=PriceData(), amfi_data=AmfiData(), today=TODAY)
