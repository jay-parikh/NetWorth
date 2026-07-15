"""Round-trip identity: generate → read → regenerate must be lossless (SPEC §7)."""

from dataclasses import asdict
from datetime import date

from networth.generate import build_workbook
from networth.model import EquityRow, SIPRow
from networth.reader import read_workbook
from networth.sample_data import sample_portfolio


def _comparable(data) -> dict:
    return {
        "persons": data.persons,
        "equity": [asdict(r) for r in data.equity],
        "mutual_funds": [asdict(r) for r in data.mutual_funds],
        "sip": [asdict(r) for r in data.sip],
        "fixed_deposits": [asdict(r) for r in data.fixed_deposits],
        "ppf": [asdict(r) for r in data.ppf],
        "bonds": [asdict(r) for r in data.bonds],
        "by_scrip": [asdict(r) for r in data.by_scrip],
        "inflation_pct": data.inflation_pct,
        "xirr": asdict(data.xirr),
        "mf_master_len": len(data.masters.mf_rows),
        "stock_master_len": len(data.masters.stock_rows),
        "mf_first": data.masters.mf_rows[0],
        "stock_last": data.masters.stock_rows[-1],
        "refreshed": (data.masters.mf_refreshed, data.masters.stock_refreshed),
    }


def test_roundtrip_identity(tmp_path):
    original = sample_portfolio()
    p1 = tmp_path / "one.xlsx"
    p2 = tmp_path / "two.xlsx"

    build_workbook(original, str(p1))
    read1 = read_workbook(str(p1))
    assert _comparable(read1) == _comparable(original)

    build_workbook(read1, str(p2))
    read2 = read_workbook(str(p2))
    assert _comparable(read2) == _comparable(read1)


def test_roundtrip_preserves_overrides(tmp_path):
    data = sample_portfolio()
    # user typed an ISIN + free-text scrip for a delisted stock, and manual units
    data.equity.append(EquityRow(owner="Amit", scrip="SOME DELISTED CO",
                                 isin_override="INE000TEST01", qty=10, avg_cost=5,
                                 cost_date=date(2017, 3, 1)))
    data.sip.append(SIPRow(owner="Rahul", scheme="Parag Parikh Flexi Cap Fund - "
                           "Direct Plan - Growth", txn_date=date(2025, 12, 1),
                           amount=5000, nav=None, units_override=57.25))
    path = tmp_path / "wb.xlsx"
    build_workbook(data, str(path))
    back = read_workbook(str(path))

    eq = back.equity[-1]
    assert eq.isin_override == "INE000TEST01"
    assert eq.scrip == "SOME DELISTED CO"
    sip = back.sip[-1]
    assert sip.units_override == 57.25
    assert sip.nav is None
