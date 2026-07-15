"""XIRR golden values (SPEC §6.1) — checked against Excel's XIRR."""

from datetime import date

import pytest

from networth.compute.xirr import xirr


def test_excel_documentation_example():
    # Microsoft's documented XIRR example → 0.373362535
    flows = [
        (date(2008, 1, 1), -10000),
        (date(2008, 3, 1), 2750),
        (date(2008, 10, 30), 4250),
        (date(2009, 2, 15), 3250),
        (date(2009, 4, 1), 2750),
    ]
    assert xirr(flows) == pytest.approx(0.373362535, abs=1e-6)


def test_two_flow_closed_form():
    flows = [(date(2020, 1, 1), -1000), (date(2021, 1, 1), 1100)]
    days = (date(2021, 1, 1) - date(2020, 1, 1)).days
    expected = (1100 / 1000) ** (365 / days) - 1
    assert xirr(flows) == pytest.approx(expected, abs=1e-7)


def test_negative_return():
    flows = [(date(2022, 6, 1), -5000), (date(2024, 6, 1), 3000)]
    r = xirr(flows)
    assert r is not None and -0.9 < r < 0


def test_degenerate_same_date():
    assert xirr([(date(2024, 1, 1), -100), (date(2024, 1, 1), 110)]) is None


def test_degenerate_one_sign():
    assert xirr([(date(2024, 1, 1), 100), (date(2025, 1, 1), 110)]) is None
    assert xirr([(date(2024, 1, 1), -100)]) is None
    assert xirr([]) is None


def test_solution_actually_zeroes_npv():
    flows = [(date(2023, m, 5), -10000) for m in range(1, 13)]
    flows.append((date(2026, 7, 1), 150000))
    r = xirr(flows)
    t0 = flows[0][0]
    npv = sum(a / (1 + r) ** ((d - t0).days / 365) for d, a in flows)
    assert abs(npv) < 1e-4
