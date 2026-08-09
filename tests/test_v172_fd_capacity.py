"""v1.7.2 — FixedDeposits row budget raised 50 → 100 (Jay, 2026-08-09).

His real workbook holds 54 FDs (a laddered family portfolio), and the
updater refused to run rather than truncate. The budget is now 100, and
a sheet filled to the EXACT cap must round-trip losslessly.
"""

from datetime import date

from networth.generate import build_workbook
from networth.model import FD_LAST_ROW, FIRST_DATA_ROW, FDRow
from networth.reader import read_workbook
from networth.sample_data import sample_portfolio
from networth.update import _refuse_overfull


def _fds(n):
    return [FDRow(owner="Amit", bank="HDFC Bank", fd_no=f"FD{i:04d}",
                  principal=100000.0, rate=7.1, start=date(2025, 4, 1),
                  maturity=date(2027, 4, 1), comp_per_year=4)
            for i in range(n)]


def test_fd_budget_is_one_hundred_rows():
    assert FD_LAST_ROW - FIRST_DATA_ROW + 1 == 100


def test_fd_sheet_round_trips_at_the_cap(tmp_path):
    cap = FD_LAST_ROW - FIRST_DATA_ROW + 1
    d = sample_portfolio()
    d.fixed_deposits = _fds(cap)
    p = tmp_path / "fds.xlsx"
    build_workbook(d, str(p))
    back = read_workbook(str(p))
    assert len(back.fixed_deposits) == cap
    assert back.fixed_deposits[-1].fd_no == f"FD{cap - 1:04d}"
    assert not any("FixedDeposits holds" in w for w in back.warnings)
    _refuse_overfull(back)              # at the cap is fine, not overfull


def test_the_workbook_jay_could_not_update_now_fits(tmp_path):
    # 54 rows — the exact size that hit "can only save 50" on v1.7.1
    d = sample_portfolio()
    d.fixed_deposits = _fds(54)
    p = tmp_path / "fifty4.xlsx"
    build_workbook(d, str(p))
    back = read_workbook(str(p))
    assert len(back.fixed_deposits) == 54
    _refuse_overfull(back)
