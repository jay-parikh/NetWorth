"""Curated corporate-action data, fetched at run time (SPEC §5.9).

Splits and bonuses come from the exchanges. Mergers and demergers do not:
no free feed publishes a swap ratio, and the cost apportionment exists
only inside a company's own filing. So those events are curated by hand —
and until v1.7.7 the only way to deliver one was a whole new app version,
which meant a user who hit a demerger waited for a release.

This module makes the curated file behave like every other feed: the
updater downloads the project's current `data/restructures.csv` on each
run, so an event added today reaches every user on their next update with
nothing for them to do. The file that ships inside the app stays the
offline baseline.

Trust rules (the same shape as the AMFI/bhavcopy fetchers):
- ONE parser, `model.parse_restructures`, judges both copies — the
  downloaded file can never be accepted on weaker terms than the shipped
  one, and its cost apportionments must still sum to 100;
- a file carrying FEWER events than the app already ships is distrusted
  wholesale, so a truncated download can never remove a merger the user's
  holdings depend on. The floor lives in here, not in the caller: a gate
  that must be passed as an argument is a gate someone will forget;
- anything else wrong — offline, proxy, 404, malformed — leaves the
  bundled copy in place. A corporate-action refresh must never be able to
  break an update run.
"""

from __future__ import annotations

CURATED_URL = ("https://raw.githubusercontent.com/jay-parikh/NetWorth/"
               "master/data/restructures.csv")
_HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
            "Accept": "text/csv, text/plain, */*"}


def fetch_restructures(session=None, timeout=(3, 5)):
    """The project's current curated events, or None. Transport only — the
    caller's `refresh_restructures` decides whether to trust the result.
    Never raises: every failure path is a quiet fallback."""
    from ..model import parse_restructures
    try:
        import requests
        sess = session or requests.Session()
        resp = sess.get(CURATED_URL, timeout=timeout, headers=_HEADERS)
        if resp.status_code != 200:
            return None
        # the file is utf-8 (company names carry en/em dashes); a response
        # without a charset would decode as ISO-8859-1 and mojibake its way
        # into the workbook — the bundled path reads explicit utf-8
        resp.encoding = "utf-8"
        # anything that isn't this file (an HTML error page, a truncated
        # body) yields no rows at all, so the parser IS the format check
        return parse_restructures(resp.text or "")
    except Exception:                        # noqa: BLE001 — see docstring
        return None


def refresh_restructures(bundled: list, session=None, timeout=(3, 5)
                         ) -> tuple[list, int]:
    """The bundled curated events topped up from the project (§5.9).

    Returns (events, how many rows were added or corrected) — the single
    entry point, so the distrust-empty floor and the Σ = 100 re-check can
    never be bypassed by a caller that forgets an argument.
    """
    fetched = fetch_restructures(session=session, timeout=timeout)
    if not fetched or len(fetched) < len(bundled):
        # fewer events than the app already ships: distrust it wholesale
        # rather than lose a merger the user's holdings depend on
        return bundled, 0
    return merge_restructures(bundled, fetched)


def merge_restructures(bundled: list, fetched: list | None) -> tuple[list, int]:
    """Bundled events, with any event the fetch also carries REPLACED by the
    fetched version.

    Replacement is per EVENT, not per row (§5.9). A demerger is only
    coherent as a whole: if an upstream correction moves a child to a
    different ISIN, a row-by-row union would keep the stale child alongside
    the new one — two valid files merging into a list that apportions more
    than 100% of the cost, appending a holding that does not exist. Whole
    events swap, and the merged list is re-validated; if the result is not
    coherent the bundled list stands, because a wrong cost basis is worse
    than an out-of-date one.
    """
    from ..model import (restructure_event_key, restructure_key,
                         validate_restructures)
    if not fetched:
        return bundled, 0
    incoming = {restructure_event_key(a) for a in fetched}
    merged = [a for a in bundled
              if restructure_event_key(a) not in incoming] + list(fetched)
    try:
        validate_restructures(merged)
    except ValueError:
        return bundled, 0
    # dataclass equality covers EVERY curated field — a hand-listed subset
    # would under-report a corrected ticker or description
    was = {restructure_key(a): a for a in bundled}
    changed = sum(1 for a in fetched if was.get(restructure_key(a)) != a)
    # rows a correction DROPPED count too: the user's sheet may already
    # carry one, so the summary must not read "nothing new"
    kept = {restructure_key(a) for a in fetched}
    changed += sum(1 for a in bundled
                   if restructure_event_key(a) in incoming
                   and restructure_key(a) not in kept)
    return merged, changed
