"""Equity bhavcopy fetcher/parser — BSE primary, NSE fallback (SPEC §5.2/5.3).

Columns are located by header name, tolerantly (the exchanges' common format
today: ISIN / ClsPric / PrvsClsgPric / TckrSymb / FinInstrmNm). Bhavcopies
are not published on holidays: walk back up to MAX_BACK days from today.
"""

from __future__ import annotations

import csv
import io
import zipfile
from dataclasses import dataclass, field
from datetime import date, timedelta

BSE_URL = ("https://www.bseindia.com/download/BhavCopy/Equity/"
           "BhavCopy_BSE_CM_0_0_0_{ymd}_F_0000.CSV")
NSE_URL = ("https://nsearchives.nseindia.com/content/cm/"
           "BhavCopy_NSE_CM_0_0_0_{ymd}_F_0000.csv.zip")
NSE_WARMUP = "https://www.nseindia.com/"
MAX_BACK = 7

_HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
            "Accept": "*/*", "Accept-Language": "en-US,en;q=0.9"}

ISIN_KEYS = ("isin", "isin_code", "isin no", "isin no.", "isin code")
CLOSE_KEYS = ("clspric", "close", "close_price", "last", "lasttradedprice")
PREV_KEYS = ("prvsclsgpric", "prevclose", "prev close", "previous close")
SYMBOL_KEYS = ("tckrsymb", "symbol", "sc_name", "scrip name")
NAME_KEYS = ("fininstrmnm", "security name", "sc_name")
CODE_KEYS = ("fininstrmid", "sc_code", "scrip code", "scrip_code")


@dataclass
class PriceData:
    prices: dict[str, dict] = field(default_factory=dict)  # isin -> {close, prev}
    # (symbol, name, isin) rows for the Stock_Master add-only merge
    master_rows: list[tuple[str, str, str]] = field(default_factory=list)
    # isin -> exchange instrument code; meaningful only for BSE (scrip code,
    # e.g. 500325) — feeds the BSE corporate-actions lookup
    codes_by_isin: dict[str, str] = field(default_factory=dict)
    trade_date: date | None = None
    source: str = ""


def _find_col(headers: list[str], wanted: tuple[str, ...], *, close: bool = False) -> str | None:
    lowered = {h.lower().strip(): h for h in headers}
    for w in wanted:
        if w in lowered:
            h = lowered[w]
            if close and ("prvs" in h.lower() or "prev" in h.lower()):
                continue
            return h
    if close:  # tolerate exotic close-column names, but never a "prev" one
        for low, h in lowered.items():
            if "cls" in low and "prvs" not in low and "prev" not in low:
                return h
    return None


def parse(csv_text: str) -> PriceData:
    out = PriceData()
    rdr = csv.DictReader(io.StringIO(csv_text))
    headers = rdr.fieldnames or []
    ic = _find_col(headers, ISIN_KEYS)
    cc = _find_col(headers, CLOSE_KEYS, close=True)
    pc = _find_col(headers, PREV_KEYS)
    sc = _find_col(headers, SYMBOL_KEYS)
    nc = _find_col(headers, NAME_KEYS)
    idc = _find_col(headers, CODE_KEYS)
    if not ic or not cc:
        raise ValueError(f"ISIN/Close columns not found. Header: {headers}")
    for row in rdr:
        isin = (row.get(ic) or "").strip()
        if not isin:
            continue
        try:
            close = float(row.get(cc) or 0)
        except ValueError:
            continue
        if close <= 0:
            continue
        prev = None
        if pc:
            try:
                p = float(row.get(pc) or 0)
                prev = p if p > 0 else None
            except ValueError:
                pass
        if isin not in out.prices:
            out.prices[isin] = {"close": close, "prev": prev}
            sym = (row.get(sc) or "").strip() if sc else ""
            name = (row.get(nc) or "").strip() if nc else sym
            out.master_rows.append((sym, name or sym, isin))
            if idc:
                code = (row.get(idc) or "").strip()
                if code:
                    out.codes_by_isin[isin] = code
    return out


def _get_bse(sess, d: date, timeout: int) -> str | None:
    resp = sess.get(BSE_URL.format(ymd=d.strftime("%Y%m%d")),
                    headers=_HEADERS, timeout=timeout)
    if resp.status_code == 200 and resp.text.lstrip()[:6].lower() != "<html>":
        return resp.text
    return None


def _get_nse(sess, d: date, timeout: int) -> str | None:
    sess.get(NSE_WARMUP, headers=_HEADERS, timeout=timeout)   # cookie warm-up
    resp = sess.get(NSE_URL.format(ymd=d.strftime("%Y%m%d")),
                    headers=_HEADERS, timeout=timeout)
    if resp.status_code == 200:
        with zipfile.ZipFile(io.BytesIO(resp.content)) as z:
            inner = [n for n in z.namelist() if n.lower().endswith(".csv")]
            if inner:
                return z.read(inner[0]).decode("utf-8", errors="replace")
    return None


def fetch(session=None, today: date | None = None, timeout: int = 60) -> PriceData:
    import requests
    sess = session or requests.Session()
    today = today or date.today()
    for back in range(MAX_BACK + 1):
        d = today - timedelta(days=back)
        for getter, source in ((_get_bse, "BSE"), (_get_nse, "NSE")):
            try:
                text = getter(sess, d, timeout)
            except requests.RequestException:
                text = None
            if text:
                try:
                    out = parse(text)
                except ValueError:
                    continue
                if out.prices:
                    out.trade_date = d
                    out.source = source
                    if source != "BSE":
                        # NSE's FinInstrmId is not a BSE scrip code
                        out.codes_by_isin = {}
                    return out
    raise RuntimeError(f"No bhavcopy available in the last {MAX_BACK} days (BSE and NSE)")
