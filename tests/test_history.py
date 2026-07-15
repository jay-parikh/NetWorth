"""Net-worth history + trend (SPEC §6.11): snapshot math, one-per-day, round-trip."""

from datetime import date

import pytest
from openpyxl import load_workbook

from networth.compute.snapshot import net_worth_snapshot, upsert_snapshot
from networth.fetch.amfi import AmfiData
from networth.fetch.bhavcopy import PriceData
from networth.generate import build_workbook
from networth.model import HistorySnapshot
from networth.reader import read_workbook
from networth.sample_data import sample_portfolio
from networth.update import run


def test_snapshot_matches_hand_computed_classes():
    data = sample_portfolio()
    snap = net_worth_snapshot(data, date(2026, 7, 15))
    # equity = Σ qty × close (no CA factors set pre-update)
    eq = sum(r.qty * r.close for r in data.equity if r.qty and r.close)
    assert snap.equity == pytest.approx(round(eq, 2), abs=0.01)
    assert snap.total == pytest.approx(
        snap.equity + snap.mutual_funds + snap.fixed_deposits + snap.ppf + snap.bonds)
    assert snap.bonds == pytest.approx(50 * 1015)     # sample NHAI bond


def test_upsert_is_one_row_per_day():
    hist: list[HistorySnapshot] = []
    hist = upsert_snapshot(hist, HistorySnapshot(date(2026, 7, 15), equity=100), keep=400)
    hist = upsert_snapshot(hist, HistorySnapshot(date(2026, 7, 16), equity=110), keep=400)
    # same-day re-run overwrites, does not append
    hist = upsert_snapshot(hist, HistorySnapshot(date(2026, 7, 16), equity=115), keep=400)
    assert [h.snap_date for h in hist] == [date(2026, 7, 15), date(2026, 7, 16)]
    assert hist[-1].equity == 115


def test_upsert_caps_to_keep():
    hist: list[HistorySnapshot] = []
    for i in range(1, 10):
        hist = upsert_snapshot(hist, HistorySnapshot(date(2026, 1, i)), keep=3)
    assert len(hist) == 3
    assert hist[0].snap_date == date(2026, 1, 7)      # kept the most recent 3


def test_two_runs_two_rows_and_roundtrip(tmp_path):
    path = tmp_path / "wb.xlsx"
    build_workbook(sample_portfolio(), str(path))

    def do_run(day):
        return run(path, price_data=PriceData(trade_date=day, source="T"),
                   amfi_data=AmfiData(), ca_data=[], today=day)

    s1 = do_run(date(2026, 7, 15))
    s2 = do_run(date(2026, 7, 16))
    assert s1["history_points"] == 1 and s2["history_points"] == 2
    assert s2["net_worth"] > 0

    back = read_workbook(str(path))
    assert [h.snap_date for h in back.history] == [date(2026, 7, 15), date(2026, 7, 16)]
    # regeneration preserves the history rows (they are data)
    build_workbook(back, str(path))
    again = read_workbook(str(path))
    assert len(again.history) == 2
    assert again.history[1].total == pytest.approx(back.history[1].total, abs=0.01)

    # Dashboard trend chart reads the History sheet
    wb = load_workbook(path)
    assert wb["History"]["A3"].value == "Date"
    assert wb["History"]["G4"].value == "=SUM(B4:F4)"


def test_same_day_rerun_does_not_duplicate(tmp_path):
    path = tmp_path / "wb.xlsx"
    build_workbook(sample_portfolio(), str(path))
    for _ in range(3):
        run(path, price_data=PriceData(trade_date=date(2026, 7, 15), source="T"),
            amfi_data=AmfiData(), ca_data=[], today=date(2026, 7, 15))
    back = read_workbook(str(path))
    assert len(back.history) == 1
