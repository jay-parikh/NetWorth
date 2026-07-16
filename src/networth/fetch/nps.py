"""NPS daily NAVs + scheme master (SPEC §5.6).

PRIMARY: https://npstrust.org.in/nav-report-excel — despite the name it is
tab-separated text (verified 2026-07-16): columns ID, DATE OF NAV, PFM NAME,
SCHEME ID, SCHEME NAME, NAV VALUE; one row per scheme, latest published day.
FALLBACK: the NSDL-CRA download (npscra.nsdl.co.in) with the same logical
record. Columns are located by header name, tolerantly; keyed by SCHEME ID
(e.g. SM001003). No API key, plain GET, graceful degradation.
"""

from __future__ import annotations

import csv
import io
from dataclasses import dataclass, field
from datetime import date, datetime

URLS = ["https://npstrust.org.in/nav-report-excel",
        "https://npscra.nsdl.co.in/download/NAVReport.csv"]
_HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
            "Accept": "*/*"}

_CODE_KEYS = ("scheme id", "scheme code", "scheme_id")
_NAME_KEYS = ("scheme name", "scheme_name")
_PFM_KEYS = ("pfm name", "pfm", "pfm_name")
_NAV_KEYS = ("nav value", "nav", "net asset value")
_DATE_KEYS = ("date of nav", "date", "nav date")


@dataclass
class NpsData:
    nav_by_code: dict[str, float] = field(default_factory=dict)
    # (scheme code, scheme name, pfm) for the NPS_Master add-only merge
    master_rows: list[tuple[str, str, str]] = field(default_factory=list)
    nav_date: date | None = None


def _find(headers: list[str], wanted: tuple[str, ...]) -> str | None:
    lowered = {h.lower().strip(): h for h in headers}
    for w in wanted:
        if w in lowered:
            return lowered[w]
    return None


def _as_date(s: str) -> date | None:
    for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%d-%b-%Y"):
        try:
            return datetime.strptime(s.strip(), fmt).date()
        except (ValueError, AttributeError):
            continue
    return None


def parse(text: str) -> NpsData:
    """Tab- or comma-separated, header-located; rows without a positive NAV
    or a scheme code are skipped."""
    delim = "\t" if "\t" in text.splitlines()[0] else ","
    rdr = csv.DictReader(io.StringIO(text), delimiter=delim)
    headers = rdr.fieldnames or []
    cc = _find(headers, _CODE_KEYS)
    nc = _find(headers, _NAME_KEYS)
    pc = _find(headers, _PFM_KEYS)
    vc = _find(headers, _NAV_KEYS)
    dc = _find(headers, _DATE_KEYS)
    if not cc or not vc:
        raise ValueError(f"NPS NAV columns not found. Header: {headers}")
    out = NpsData()
    for row in rdr:
        code = (row.get(cc) or "").strip()
        if not code:
            continue
        try:
            nav = float(row.get(vc) or 0)
        except ValueError:
            continue
        if nav <= 0:
            continue
        if code not in out.nav_by_code:
            out.nav_by_code[code] = nav
            name = (row.get(nc) or "").strip() if nc else code
            pfm = (row.get(pc) or "").strip() if pc else ""
            out.master_rows.append((code, name or code, pfm))
        if out.nav_date is None and dc:
            out.nav_date = _as_date(row.get(dc) or "")
    return out


def fetch(session=None, timeout: int = 60) -> NpsData:
    import requests
    sess = session or requests.Session()
    last_err: Exception | None = None
    for url in URLS:
        try:
            resp = sess.get(url, headers=_HEADERS, timeout=timeout)
            resp.raise_for_status()
            out = parse(resp.text)
            if out.nav_by_code:
                return out
        except Exception as e:  # noqa: BLE001 — try the next host
            last_err = e
    raise RuntimeError(f"NPS NAV feed unreachable (both hosts): {last_err}")
