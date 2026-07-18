"""CAMS/KFintech detailed CAS text parser (SPEC §6.17).

Line-oriented state machine over the extracted PDF text — the amfi.py
regex style, owned drift and all. It does NOT need to be perfect to be
safe: every fund it emits must still reconcile against the statement's
own closing unit balance in merge.py, so a missed or mangled line turns
into a loud per-fund refusal, never a wrong number on a sheet.
"""

from __future__ import annotations

import re
from datetime import date

from .common import (ImportBatch, ImportedSipTxn, parse_date_any, parse_inr)

_FOLIO_RE = re.compile(r"Folio\s*No\s*[.:]*\s*([0-9][0-9/ ]*[0-9]|[0-9]+)")
_ISIN_RE = re.compile(r"ISIN\s*[.:]*\s*(INF[A-Z0-9]{9})")
_OPEN_RE = re.compile(r"Opening\s+Unit\s+Balance[^0-9\-]*([\d,.\-]+)")
_CLOSE_RE = re.compile(r"Closing\s+Unit\s+Balance[^0-9\-]*([\d,.\-]+)")
_TXN_DATE_RE = re.compile(r"^\s*(\d{2}-[A-Za-z]{3}-\d{4})\s+(.*)$")
_NUM_TOKEN_RE = re.compile(r"\(?-?[0-9][0-9,]*(?:\.[0-9]+)?\)?")
_NAME_RE = re.compile(r"^[A-Z][A-Z .]{2,60}$")

# fee/charge lines carry a date + money but no units — informational in a
# CAS (the charge is inside the purchase amount); skipped by design.
# Short tokens match as WHOLE WORDS so e.g. a scheme name containing the
# letters can never hide a real transaction.
_CHARGE_RE = re.compile(
    r"stamp\s*duty|security\s+transaction|\*{2,}|\b(?:stt|tds)\b",
    re.IGNORECASE)

_OUTFLOW = ("redemption", "switch out", "switch-out", "switchout",
            "lateral shift out")
_TYPE_WORDS = (
    ("SWITCH_IN", ("switch in", "switch-in", "switchin",
                   "lateral shift in")),
    ("SWITCH_OUT", _OUTFLOW[1:]),
    ("REDEMPTION", ("redemption",)),
    ("DIV_REINVEST", ("reinvest",)),
    ("SEGREGATION", ("segregat",)),
    ("BONUS", ("bonus",)),
    ("PURCHASE", ("purchase", "systematic", "sip", "subscription",
                  "new fund offer", "nfo")),
)


def looks_like_cas(text: str) -> bool:
    head = text[:4000].casefold()
    return ("consolidated account statement" in head
            and ("cams" in head or "kfintech" in head or "kfin" in head
                 or "karvy" in head))


def _txn_type(desc: str) -> str:
    d = desc.casefold()
    for ttype, words in _TYPE_WORDS:
        if any(w in d for w in words):
            return ttype
    return "PURCHASE" if "dividend" not in d else "DIV_REINVEST"


def parse(text: str, today: date, path: str = "") -> ImportBatch:
    """Detailed CAS text → ImportBatch. Raises ValueError with a
    plain-words message for a summary-variant statement."""
    batch = ImportBatch(source="fund statement (CAS)", path=path)
    folio = isin = scheme = ""
    fund_house = ""
    opening: dict[tuple[str, str], float] = {}
    bad_balance: set[tuple[str, str]] = set()
    names: list[str] = []          # candidate investor-name hints, in order
    folio_hint: dict[str, str] = {}
    prev_text = ""                 # last plain line — KFin puts the scheme
    lines = text.splitlines()      # name on its OWN line above the ISIN
    for raw in lines:
        line = raw.strip()
        if not line:
            continue

        m = _ISIN_RE.search(line)
        if m:
            isin = m.group(1)
            # the scheme name is the text before the ISIN/advisor tail on
            # this line, or the previous plain line (multi-line headers)
            head = line[:m.start()].strip(" -:")
            head = re.sub(r"\(Advisor.*$", "", head).strip(" -:")
            if len(head) <= 3 and len(prev_text) > 3:
                head = re.sub(r"\(Advisor.*$", "", prev_text).strip(" -:")
            scheme = head if len(head) > 3 else ""
            if "-" in scheme:
                fund_house = scheme.split("-")[0].strip()
            if not scheme:
                batch.warnings.append(
                    f"a fund near folio {folio or '?'} has no readable name")

        m = _FOLIO_RE.search(line)
        if m and "folio" in line.casefold():
            folio = m.group(1).replace(" ", "")
            if names:
                folio_hint.setdefault(folio, names[-1])

        m = _OPEN_RE.search(line)
        if m and folio and isin:
            v = parse_inr(m.group(1))
            if v is None:
                # a balance line we SAW but couldn't read: the fund can't
                # be trusted at all (defaulting to 0 would disguise a
                # mid-history statement as since-inception)
                bad_balance.add((folio, isin))
            else:
                opening[(folio, isin)] = v

        m = _CLOSE_RE.search(line)
        if m and folio and isin:
            v = parse_inr(m.group(1))
            if v is None:
                bad_balance.add((folio, isin))
            else:
                batch.closing_units[(folio, isin)] = v

        m = _TXN_DATE_RE.match(line)
        if m and folio and isin:
            desc_and_nums = m.group(2)
            tokens = _NUM_TOKEN_RE.findall(desc_and_nums)
            desc = _NUM_TOKEN_RE.sub("", desc_and_nums).strip(" -:")
            if _CHARGE_RE.search(desc):
                continue                       # fee line: no units, by design
            if len(tokens) < 3:
                batch.warnings.append(
                    f"{scheme or isin}: couldn't read the line "
                    f"'{line[:70]}' - if it is a transaction, that fund "
                    "will be refused by the balance check")
                continue
            # layout: ... amount units nav [unit balance]
            nums = [parse_inr(t) for t in tokens[-4:]]
            if len(nums) == 4:
                amount, units, nav, _bal = nums
            else:
                amount, units, nav = nums[-3:]
            ttype = _txn_type(desc)
            # sign by TYPE, not by print style — some layouts print
            # redemption amounts positive with negative units
            if ttype in ("REDEMPTION", "SWITCH_OUT"):
                amount = -abs(amount) if amount is not None else None
                units = -abs(units) if units is not None else None
            elif ttype in ("PURCHASE", "SWITCH_IN", "DIV_REINVEST"):
                amount = abs(amount) if amount is not None else None
                units = abs(units) if units is not None else None
            elif ttype in ("BONUS", "SEGREGATION"):
                # keep whatever amount was PARSED — the merge gate refuses
                # a bonus line carrying money (evidence of a column shift);
                # forcing 0 here would destroy that evidence
                amount = amount or 0.0
            batch.sip_txns.append(ImportedSipTxn(
                folio=folio, isin=isin, scheme_name=scheme,
                fund_house=fund_house,
                txn_date=parse_date_any(m.group(1), today),
                amount=amount, nav=nav, units=units, txn_type=ttype))
            continue

        if _NAME_RE.match(line) and not any(
                w in line.casefold() for w in ("folio", "isin", "balance",
                                               "registrar", "nav", "kyc",
                                               "pan", "total", "statement")):
            names.append(line.title())

        # remember a plain line as the possible scheme name of the NEXT
        # ISIN line (KFin multi-line headers); marker lines never qualify
        low = line.casefold()
        if not any(w in low for w in ("folio", "isin", "balance",
                                      "registrar", "nav on", "kyc", "pan",
                                      "page ", "statement", "email",
                                      "opening", "closing", "valuation")):
            prev_text = line

    if batch.closing_units and not batch.sip_txns:
        raise ValueError(
            "this looks like the SUMMARY statement - it has balances but "
            "no transactions. Request the DETAILED one (pick 'Detailed' "
            "and 'Since inception' on the website) and try again.")
    if not batch.sip_txns and not batch.closing_units:
        raise ValueError(
            "couldn't find any fund transactions in this file - is it a "
            "CAMS/KFintech account statement?")

    def _drop_fund(key: tuple[str, str], why: str) -> None:
        # refused funds are also marked `partial` so the merge never
        # replaces or extends the same ISIN from a sibling folio that DID
        # parse — half a fund is worse than none
        batch.warnings.append(why)
        batch.sip_txns = [t for t in batch.sip_txns
                          if (t.folio, t.isin) != key]
        batch.closing_units.pop(key, None)
        batch.partial.add(key)

    # a balance line that was seen but unreadable poisons its whole fund —
    # without trustworthy balances neither the mid-history check nor the
    # closing reconciliation can prove anything about it
    for key in sorted(bad_balance):
        _drop_fund(key, f"folio {key[0]}: a unit-balance line couldn't be "
                        "read reliably - nothing was imported for this fund")

    # a fund whose statement starts mid-history can't be trusted as a
    # complete ledger — refuse it and say why
    for key, open_units in sorted(opening.items()):
        if abs(open_units) > 0.001:
            _drop_fund(key, f"folio {key[0]}: the statement starts "
                            f"mid-history ({open_units:g} units already "
                            "held) - request a 'Since inception' statement "
                            "to import this fund")

    # every CAS prints a closing balance per fund; one we never found means
    # the layout defeated us — refuse rather than import unreconciled
    for key in sorted({(t.folio, t.isin) for t in batch.sip_txns}):
        if key not in batch.closing_units:
            _drop_fund(key, f"folio {key[0]}: the fund's closing balance "
                            "line couldn't be found, so its history can't "
                            "be checked - nothing was imported for it")

    batch.accounts = sorted({(t.folio, folio_hint.get(t.folio, ""))
                             for t in batch.sip_txns})
    return batch
