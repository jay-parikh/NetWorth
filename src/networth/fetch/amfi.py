"""AMFI NAVAll.txt fetcher/parser (SPEC §5.1).

Format: `;`-separated lines
    Scheme Code;ISIN Div Payout/ ISIN Growth;ISIN Div Reinvestment;Scheme Name;Net Asset Value;Date
Lines without `;` are fund-house section headers; blanks and the column
header are skipped. A scheme yields up to two ISIN entries.
"""

from __future__ import annotations

from dataclasses import dataclass, field

URL = "https://www.amfiindia.com/spages/NAVAll.txt"


@dataclass
class AmfiData:
    nav_by_isin: dict[str, float] = field(default_factory=dict)
    # (fund_house, scheme_name, isin) — one row per ISIN, dropdown master shape
    master_rows: list[tuple[str, str, str]] = field(default_factory=list)
    nav_date: str = ""


def parse(text: str) -> AmfiData:
    out = AmfiData()
    fund_house = ""
    seen: set[str] = set()
    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue
        if ";" not in line:
            fund_house = line
            continue
        parts = [p.strip() for p in line.split(";")]
        if len(parts) < 6 or parts[0] == "Scheme Code":
            continue
        _code, isin1, isin2, scheme, nav_s, date_s = parts[:6]
        try:
            nav = float(nav_s)
        except ValueError:
            continue
        for isin in (isin1, isin2):
            if isin and isin not in ("-", "N.A.") and isin not in seen:
                seen.add(isin)
                out.nav_by_isin[isin] = nav
                out.master_rows.append((fund_house, scheme, isin))
                out.nav_date = date_s
    return out


# v1.6.2: the live feed carries ~14k schemes; a 200-OK maintenance page
# parses to almost nothing, and a wholesale master replace from it would
# gut the fund list. A floor, not a target — injected test data never goes
# through fetch(), so tiny fixtures stay exempt on purpose.
AMFI_MIN_SCHEMES = 100


def fetch(session=None, timeout: int = 60) -> AmfiData:
    import requests
    sess = session or requests.Session()
    resp = sess.get(URL, timeout=timeout)
    resp.raise_for_status()
    data = parse(resp.text)
    if len(data.nav_by_isin) < AMFI_MIN_SCHEMES:
        raise ValueError("AMFI returned almost nothing - keeping the "
                         "existing fund list and NAVs")
    return data
