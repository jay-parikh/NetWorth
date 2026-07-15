"""Corporate-actions fetcher (SPEC §5.4) — NSE and BSE, deduplicated.

NSE: GET https://www.nseindia.com/api/corporates-corporateActions?index=equities&symbol=X
(after the usual cookie warm-up); the free-text `subject` describes the action.
BSE: GET api.bseindia.com .../DefaultData/w?...&scripcode=<code> (Referer
required); the free-text `Purpose` describes the action; scrip codes come from
the daily BSE bhavcopy.

Only splits, bonuses and consolidations are kept; dividends, rights, AGMs etc.
are ignored. Records from both exchanges are deduplicated on
(isin, type, ex_date) — the ex-date is exchange-synchronised. The fetcher also
reports WHICH ISINs were successfully checked, so the updater can warn about
any holding it could not verify instead of skipping it silently. Anything the
feeds miss (mergers, demergers) is entered as a Manual row on the
Corporate_Actions sheet.
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

BSE_API = ("https://api.bseindia.com/BseIndiaAPI/api/DefaultData/w"
           "?Fdate=&Purpose=&TDate=&ddlcategorys=E&ddlindustrys="
           "&scripcode={code}&segment=0&strSearch=S")
_BSE_HEADERS = {"User-Agent": _HEADERS["User-Agent"],
                "Accept": "application/json",
                "Referer": "https://www.bseindia.com/"}

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
    for fmt in ("%d-%b-%Y", "%d-%m-%Y", "%Y-%m-%d", "%d %b %Y"):
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


def parse_bse_records(records: list[dict], isin: str, symbol: str
                      ) -> list[CorporateAction]:
    out = []
    for rec in records:
        purpose = rec.get("Purpose") or rec.get("purpose") or ""
        parsed = parse_subject(purpose)
        if not parsed:
            continue
        ex = _parse_date(rec.get("Ex_date") or rec.get("ex_date") or "")
        if not ex:
            continue
        kind, frm, to = parsed
        out.append(CorporateAction(symbol=symbol, isin=isin, type=kind,
                                   ex_date=ex, ratio_from=frm, ratio_to=to,
                                   source="Auto", details=purpose.strip()))
    return out


def dedupe(*action_lists: list[CorporateAction]) -> list[CorporateAction]:
    """Merge sources on (isin, type, ex_date); earlier lists win."""
    seen: set[tuple] = set()
    out: list[CorporateAction] = []
    for actions in action_lists:
        for a in actions:
            key = (a.isin, a.type, a.ex_date)
            if key not in seen:
                seen.add(key)
                out.append(a)
    return out


def fetch(nse_symbols: dict[str, str], bse_codes: dict[str, str] | None = None,
          session=None, timeout: int = 30
          ) -> tuple[list[CorporateAction], set[str]]:
    """Query NSE (symbol → ISIN) and BSE (scrip code → ISIN) for the held
    stocks; return (deduplicated actions, ISINs successfully checked on at
    least one exchange). Raises only if every query on both exchanges failed."""
    import requests
    sess = session or requests.Session()
    bse_codes = bse_codes or {}
    checked: set[str] = set()
    nse_actions: list[CorporateAction] = []
    bse_actions: list[CorporateAction] = []
    attempts = failures = 0

    if nse_symbols:
        try:
            sess.get(WARMUP, headers=_HEADERS, timeout=timeout)
        except requests.RequestException:
            pass
        for symbol, isin in sorted(nse_symbols.items()):
            attempts += 1
            try:
                resp = sess.get(API.format(symbol=symbol), headers=_HEADERS,
                                timeout=timeout)
                resp.raise_for_status()
                records = resp.json()
                if isinstance(records, dict):
                    records = records.get("data", [])
                nse_actions.extend(parse_records(records, isin, symbol))
                checked.add(isin)
            except Exception:  # noqa: BLE001 — one bad symbol must not kill the run
                failures += 1

    isin_to_symbol = {isin: sym for sym, isin in nse_symbols.items()}
    for code, isin in sorted(bse_codes.items()):
        attempts += 1
        try:
            resp = sess.get(BSE_API.format(code=code), headers=_BSE_HEADERS,
                            timeout=timeout)
            resp.raise_for_status()
            records = resp.json()
            if isinstance(records, dict):
                records = records.get("Table", [])
            bse_actions.extend(
                parse_bse_records(records, isin, isin_to_symbol.get(isin, code)))
            checked.add(isin)
        except Exception:  # noqa: BLE001
            failures += 1

    if attempts and failures == attempts:
        raise RuntimeError(
            "corporate-actions feeds unreachable (NSE and BSE, every security)")
    return dedupe(nse_actions, bse_actions), checked
