"""XIRR solver (SPEC §6.1) — Newton with bisection fallback.

Matches Excel's XIRR convention: f(r) = Σ amount_i / (1+r)^(days_i/365),
days measured from the first cashflow date. Returns None (never raises, never
0) on degenerate input so callers render a blank cell.
"""

from __future__ import annotations

from datetime import date

LOW, HIGH = -0.9999, 10.0
TOL = 1e-7
MAX_ITER = 100


def xirr(flows: list[tuple[date, float]]) -> float | None:
    flows = [(d, a) for d, a in flows if a]
    if len(flows) < 2:
        return None
    flows.sort(key=lambda f: f[0])
    t0 = flows[0][0]
    days = [(d - t0).days for d, _ in flows]
    if days[-1] == 0:
        return None                      # all flows on the same date
    amounts = [a for _, a in flows]
    if all(a > 0 for a in amounts) or all(a < 0 for a in amounts):
        return None                      # no sign change → no IRR

    def f(r: float) -> float:
        return sum(a / (1.0 + r) ** (d / 365.0) for a, d in zip(amounts, days))

    def fprime(r: float) -> float:
        return sum(-(d / 365.0) * a / (1.0 + r) ** (d / 365.0 + 1)
                   for a, d in zip(amounts, days))

    # v1.6.2: absurd-but-typeable dates make the power term explode — a
    # year-9999 placeholder overflows float range (~7,400-year span) and an
    # ~80-year span underflows f(LOW)'s denominator to 0.0. Both are
    # divergence, not errors: keep the docstring's never-raises promise.
    try:
        # Newton from a mild positive guess
        r = 0.1
        for _ in range(MAX_ITER):
            fr = f(r)
            if abs(fr) < TOL:
                return r
            fp = fprime(r)
            if fp == 0:
                break
            nxt = r - fr / fp
            if nxt <= LOW or nxt >= HIGH or nxt != nxt:
                break
            if abs(nxt - r) < TOL:
                return nxt
            r = nxt

        # bisection fallback on [LOW, HIGH]
        lo, hi = LOW, HIGH
        flo, fhi = f(lo), f(hi)
        if flo * fhi > 0:
            return None
        for _ in range(200):
            mid = (lo + hi) / 2
            fm = f(mid)
            if abs(fm) < TOL or (hi - lo) / 2 < TOL:
                return mid
            if flo * fm < 0:
                hi, fhi = mid, fm
            else:
                lo, flo = mid, fm
        return None
    except (OverflowError, ZeroDivisionError):
        return None
