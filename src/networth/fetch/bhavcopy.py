"""Equity bhavcopy fetcher/parser — BSE + NSE as full peers (SPEC §5.2/5.3).

Both exchanges are fetched for the SAME trade date and merged: union of
ISINs, NSE close/prev win on dual-listed conflicts (deeper liquidity — what
broker apps show), while BSE alone contributes `codes_by_isin` (scrip codes
feeding the BSE corporate-actions API). If only one exchange published for a
day the run proceeds single-source; dates are never mixed across exchanges.

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
    # exchanges that actually contributed to this day's data ("BSE"/"NSE");
    # the §6.5 status escalation only runs when BOTH did
    sources: list[str] = field(default_factory=list)
    # ISINs quoted on NSE but not BSE that day (console summary only)
    nse_only: set[str] = field(default_factory=set)


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
        try:
            with zipfile.ZipFile(io.BytesIO(resp.content)) as z:
                inner = [n for n in z.namelist() if n.lower().endswith(".csv")]
                if inner:
                    return z.read(inner[0]).decode("utf-8", errors="replace")
        except zipfile.BadZipFile:
            # a 200 that isn't a zip (bot-challenge HTML page): treat as
            # NSE-unavailable so the day degrades to BSE-only per SPEC §5.2
            return None
    return None


def merge(bse: PriceData | None, nse: PriceData | None) -> PriceData:
    """Union of both exchanges' bhavcopies for one trade date (SPEC §5.2).

    NSE close/prev win on dual-listed conflicts; scrip codes come only from
    BSE (NSE's FinInstrmId is not a BSE scrip code); for ISINs new to the
    master, the NSE symbol is preferred (it is what the NSE corporate-actions
    API needs — the add-only merge protects existing rows regardless).
    """
    out = PriceData()
    if bse:
        out.prices.update(bse.prices)
        out.codes_by_isin = dict(bse.codes_by_isin)
        out.sources.append("BSE")
    if nse:
        if bse:
            out.nse_only = set(nse.prices) - set(bse.prices)
        out.prices.update(nse.prices)
        out.sources.append("NSE")
    seen: set[str] = set()
    for src in (nse, bse):
        if src is None:
            continue
        for sym, name, isin in src.master_rows:
            if isin not in seen:
                seen.add(isin)
                out.master_rows.append((sym, name, isin))
    out.source = "+".join(out.sources)
    return out


def fetch(session=None, today: date | None = None, timeout: int = 60) -> PriceData:
    import requests
    sess = session or requests.Session()
    today = today or date.today()
    for back in range(MAX_BACK + 1):
        d = today - timedelta(days=back)
        parsed: dict[str, PriceData] = {}
        for getter, source in ((_get_bse, "BSE"), (_get_nse, "NSE")):
            try:
                text = getter(sess, d, timeout)
            except requests.RequestException:
                text = None
            if not text:
                continue
            try:
                p = parse(text)
            except ValueError:
                continue
            if p.prices:
                parsed[source] = p
        if parsed:
            out = merge(parsed.get("BSE"), parsed.get("NSE"))
            out.trade_date = d
            return out
    raise RuntimeError(f"No bhavcopy available in the last {MAX_BACK} days (BSE and NSE)")
