"""Corporate-actions fetcher (SPEC §5.4) — NSE and BSE, deduplicated.

NSE: GET https://www.nseindia.com/api/corporates-corporateActions?index=equities&symbol=X
(after the usual cookie warm-up); the free-text `subject` describes the action.
BSE: GET api.bseindia.com .../DefaultData/w?...&scripcode=<code> (Referer
required); the free-text `Purpose` describes the action; scrip codes come from
the daily BSE bhavcopy.

Splits, bonuses and consolidations feed the adjustment engine; dividend
announcements feed the Dividends sheet (SPEC §3.13) — rights, AGMs, buybacks
etc. are ignored. Records from both exchanges are deduplicated on
(isin, type, ex_date) — the ex-date is exchange-synchronised — NSE wins. The
fetcher also reports WHICH ISINs were successfully checked, so the updater can
warn about any holding it could not verify instead of skipping it silently.
Anything the feeds miss (mergers, demergers) is entered as a Manual row on the
Corporate_Actions sheet.
"""

from __future__ import annotations

import re
from datetime import date, datetime

from ..model import CorporateAction, DividendRow

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


_DIV_HINT = re.compile(r"dividend", re.I)
_DIV_TYPE = re.compile(r"\b(interim|final|special)\b", re.I)
# "Rs 8 Per Share", "Rs. - 5.5000", "Re. 1/-", "₹2.50" — a rupee amount, never
# a bare number (that would swallow "Dividend 250%" percent-of-face forms).
# The lookbehind keeps "re" from matching inside a word ("...Per Share 2024").
_DIV_RATE = re.compile(
    r"(?<![a-z])(?:rs|re|₹)\s*\.?\s*[-–]?\s*(\d+(?:\.\d+)?)\s*(?:/-)?", re.I)
# "...300% on face value of Rs.2/- each": the rupee amount is the FACE VALUE,
# not the rate — masked out before the rate search so these hit the
# skip-and-warn path instead of parsing garbage
_FACE_VALUE = re.compile(
    r"face\s*value\s*(?:of\s*)?(?:rs|re|₹)\s*\.?\s*\d+(?:\.\d+)?\s*(?:/-)?"
    r"(?:\s*each)?", re.I)


def parse_dividend(subject: str) -> tuple[str, float] | None:
    """→ (Interim|Final|Special, ₹/share) or None.

    Percent-of-face forms ("Dividend 250%") return None — rare since face
    values shrank post-2010; a Manual row on the Dividends sheet covers them
    (the updater counts and reports every skipped dividend subject)."""
    if not _DIV_HINT.search(subject):
        return None
    m = _DIV_RATE.search(_FACE_VALUE.sub(" ", subject))
    if not m:
        return None
    t = _DIV_TYPE.search(subject)
    return (t.group(1).capitalize() if t else "Final"), float(m.group(1))


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


def parse_dividend_records(records: list[dict], isin: str, symbol: str,
                           skip_since: date | None = None
                           ) -> tuple[list[DividendRow], int]:
    """Dividend announcements from either exchange's record shape.

    Returns (feed rows — no owner/qty yet, the updater expands those —,
    count of subjects that could not be parsed). The count is gated to
    ex-dates ≥ `skip_since` (the current FY start): the feeds carry decades
    of history and warning about a 2004 percent-of-face dividend is noise."""
    out: list[DividendRow] = []
    skipped = 0
    for rec in records:
        subject = (rec.get("subject") or rec.get("purpose")
                   or rec.get("Purpose") or "")
        if not _DIV_HINT.search(subject):
            continue
        parsed = parse_dividend(subject)
        ex = _parse_date(rec.get("exDate") or rec.get("exdate")
                         or rec.get("Ex_date") or rec.get("ex_date") or "")
        if not parsed:
            if ex and (skip_since is None or ex >= skip_since):
                skipped += 1
            continue
        if not ex:
            continue
        div_type, rate = parsed
        out.append(DividendRow(scrip=symbol, isin=isin, div_type=div_type,
                               ex_date=ex, rate=rate, source="Auto",
                               details=subject.strip()))
    return out, skipped


def dedupe_dividends(*div_lists: list[DividendRow]) -> list[DividendRow]:
    """Merge sources on (isin, ex_date, rate); earlier lists win.

    Deliberately NOT keyed on div_type: the exchanges word the same event
    differently ("Dividend" → Final on NSE, "Interim Dividend" on BSE) and
    keying on the type would double-count it. Two genuinely distinct payouts
    on one ex-date differ in rate, which stays in the key."""
    seen: set[tuple] = set()
    out: list[DividendRow] = []
    for divs in div_lists:
        for d in divs:
            key = (d.isin, d.ex_date, d.rate)
            if key not in seen:
                seen.add(key)
                out.append(d)
    return out


def fetch(nse_symbols: dict[str, str], bse_codes: dict[str, str] | None = None,
          session=None, timeout: int = 30, div_skip_since: date | None = None
          ) -> tuple[list[CorporateAction], set[str], list[DividendRow], int]:
    """Query NSE (symbol → ISIN) and BSE (scrip code → ISIN) for the held
    stocks; return (deduplicated actions, ISINs successfully checked on at
    least one exchange, deduplicated dividend announcements, count of
    dividend subjects that could not be parsed). Raises only if every query
    on both exchanges failed."""
    import requests
    sess = session or requests.Session()
    bse_codes = bse_codes or {}
    checked: set[str] = set()
    nse_actions: list[CorporateAction] = []
    bse_actions: list[CorporateAction] = []
    nse_divs: list[DividendRow] = []
    bse_divs: list[DividendRow] = []
    attempts = failures = div_skipped = 0

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
                    # v1.6.2: a dict WITHOUT the expected key is a changed/
                    # maintenance response, not "no actions" — treating it
                    # as empty would mark the ISIN checked and let the
                    # updater drop its kept Auto rows. Fail it instead so
                    # the rows are preserved.
                    if "data" not in records:
                        raise ValueError("unexpected NSE response shape")
                    records = records["data"]
                if not isinstance(records, list):
                    raise ValueError("unexpected NSE response shape")
                nse_actions.extend(parse_records(records, isin, symbol))
                divs, skipped = parse_dividend_records(records, isin, symbol,
                                                       div_skip_since)
                nse_divs.extend(divs)
                div_skipped += skipped
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
                if "Table" not in records:      # same guard as the NSE path
                    raise ValueError("unexpected BSE response shape")
                records = records["Table"]
            if not isinstance(records, list):
                raise ValueError("unexpected BSE response shape")
            symbol = isin_to_symbol.get(isin, code)
            bse_actions.extend(parse_bse_records(records, isin, symbol))
            divs, skipped = parse_dividend_records(records, isin, symbol,
                                                   div_skip_since)
            bse_divs.extend(divs)
            div_skipped += skipped
            checked.add(isin)
        except Exception:  # noqa: BLE001
            failures += 1

    if attempts and failures == attempts:
        raise RuntimeError(
            "corporate-actions feeds unreachable (NSE and BSE, every security)")
    return (dedupe(nse_actions, bse_actions), checked,
            dedupe_dividends(nse_divs, bse_divs), div_skipped)
