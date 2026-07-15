"""Corporate-actions fetcher (SPEC §5.4) — NSE per-symbol API.

GET https://www.nseindia.com/api/corporates-corporateActions?index=equities&symbol=X
(after the usual cookie warm-up) returns JSON records whose free-text
`subject` describes the action. Only splits, bonuses and consolidations are
kept; dividends, rights, AGMs etc. are ignored. The parsing contract is the
CorporateAction record — anything the feed misses can be entered as a Manual
row on the Corporate_Actions sheet.
"""

from __future__ import annotations

import re
from datetime import date, datetime

from ..model import CorporateAction

API = ("https://www.nseindia.com/api/corporates-corporateActions"
       "?index=equities&symbol={symbol}")
WARMUP = "https://www.nseindia.com/"
_HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
            "Accept": "application/json", "Accept-Language": "en-US,en;q=0.9"}

_NUM = r"(?:rs\.?|re\.?|rs|re)?\s*\.?\s*(\d+(?:\.\d+)?)\s*(?:/-)?"
_FROMTO = re.compile(r"fro?m\s*" + _NUM + r".*?to\s*" + _NUM, re.I)
_BONUS = re.compile(r"bonus[^\d]{0,20}(\d+)\s*:\s*(\d+)", re.I)
_SPLIT_HINT = re.compile(r"splt|split|sub-?divi", re.I)
_CONSOL_HINT = re.compile(r"consolidat", re.I)


def parse_subject(subject: str) -> tuple[str, float, float] | None:
    """→ (type, ratio_from, ratio_to) or None for non-adjusting actions."""
    m = _BONUS.search(subject)
    if m:
        return "BONUS", float(m.group(1)), float(m.group(2))
    if _SPLIT_HINT.search(subject) or _CONSOL_HINT.search(subject):
        m = _FROMTO.search(subject)
        if m:
            frm, to = float(m.group(1)), float(m.group(2))
            kind = "CONSOLIDATION" if (_CONSOL_HINT.search(subject) or to > frm) else "SPLIT"
            return kind, frm, to
    return None


def _parse_date(s: str) -> date | None:
    for fmt in ("%d-%b-%Y", "%d-%m-%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(s.strip(), fmt).date()
        except (ValueError, AttributeError):
            continue
    return None


def parse_records(records: list[dict], isin: str, symbol: str) -> list[CorporateAction]:
    out = []
    for rec in records:
        subject = rec.get("subject") or rec.get("purpose") or ""
        parsed = parse_subject(subject)
        if not parsed:
            continue
        ex = _parse_date(rec.get("exDate") or rec.get("exdate") or "")
        if not ex:
            continue
        kind, frm, to = parsed
        out.append(CorporateAction(symbol=symbol, isin=isin, type=kind,
                                   ex_date=ex, ratio_from=frm, ratio_to=to,
                                   source="Auto", details=subject.strip()))
    return out


def fetch(symbols: dict[str, str], session=None, timeout: int = 30
          ) -> list[CorporateAction]:
    """symbols: exchange symbol → ISIN, for the held stocks only."""
    import requests
    sess = session or requests.Session()
    sess.get(WARMUP, headers=_HEADERS, timeout=timeout)
    actions: list[CorporateAction] = []
    failures = 0
    for symbol, isin in sorted(symbols.items()):
        try:
            resp = sess.get(API.format(symbol=symbol), headers=_HEADERS,
                            timeout=timeout)
            resp.raise_for_status()
            records = resp.json()
            if isinstance(records, dict):
                records = records.get("data", [])
            actions.extend(parse_records(records, isin, symbol))
        except Exception:  # noqa: BLE001 — one bad symbol must not kill the run
            failures += 1
    if failures == len(symbols) and symbols:
        raise RuntimeError("corporate-actions feed unreachable for every symbol")
    return actions
