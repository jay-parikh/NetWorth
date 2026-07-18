"""The one merge engine: reconcile ImportBatches into PortfolioData (§6.17).

Never-garbage contract, enforced here regardless of which parser produced
the batch:

- every transaction must prove itself (amount = units × NAV triangle);
  an unprovable row poisons its WHOLE fund — funds import atomically or
  not at all, so a statement that can't be read reliably changes nothing;
- when the statement declares a closing unit balance, the parsed history
  must sum to it (±0.001 unit) or the fund is refused — a silently
  dropped line is caught by arithmetic, not by luck;
- a negative running-units balance mid-history (sell before buy) refuses
  the fund — it proves ordering or parse corruption;
- the engine mutates `data` only after EVERY gate has passed; a deferred
  import (capacity) leaves the workbook untouched.

Statement-wins (Jay, 2026-07-18): for funds the statement covers, typed
rows are replaced by the exact history — but only when the caller passed
replace=True, which update.py sets ONLY after an interactive confirmation.
Headless runs always append-only (multiset dedupe), never destructive.
"""

from __future__ import annotations

from collections import Counter
from dataclasses import dataclass, field
from datetime import date

from .. import model as M
from .common import ImportBatch, ImportedSipTxn, UNIT_TOLERANCE, triangle_ok

# txn types that legitimately carry units without money changing hands;
# exempt from the triangle but still inside the balance reconciliation.
# Their MF_SIP rows land with amount 0 — valuation counts the units via
# units_override, the return figure (XIRR) ignores the zero cash leg.
_UNITS_ONLY_TYPES = {"BONUS", "SEGREGATION"}


@dataclass
class FundLine:
    """One fund's fate in the preview/summary — every fund ends in a tick
    or a plain-words reason, never silence."""
    owner: str = ""                    # "" = folio not mapped to a person
    folio: str = ""
    isin: str = ""
    scheme: str = ""
    added: int = 0
    skipped: int = 0                   # already on the sheet (dedupe)
    replaced: int = 0                  # typed rows removed (statement-wins)
    invested: float = 0.0              # Σ positive amounts, for the preview
    units: float = 0.0                 # Σ signed units
    reconciled: bool = False           # matched the statement's own balance
    ok: bool = True
    reason: str = ""                   # why the fund was refused / skipped


@dataclass
class MergeReport:
    funds: list[FundLine] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    deferred: list[str] = field(default_factory=list)
    sip_added: int = 0
    sip_skipped: int = 0
    sip_replaced: int = 0
    stocks: list[FundLine] = field(default_factory=list)   # equity, per ISIN
    eq_added: int = 0                  # lots landed on the Equity sheet
    eq_skipped: int = 0                # already present (re-import)
    eq_reduced: int = 0                # typed lots consumed by netted sells
    sells_added: int = 0               # rows landed on Equity_Sells


def _sip_row_key(row, isin_by_scheme) -> tuple:
    isin = row.isin_override or isin_by_scheme.get(row.scheme, "")
    return (row.owner, isin, row.txn_date, round(row.amount or 0.0, 2))


def _validate_fund(txns: list[ImportedSipTxn], anchor: float | None,
                   today: date) -> str:
    """All gates for one (folio, isin) fund. Returns '' or the plain-words
    refusal reason. One bad row refuses the whole fund (atomicity)."""
    running = 0.0
    total = 0.0
    for t in txns:
        if t.txn_date is None or t.txn_date > today:
            return "a transaction has no readable date"
        if t.units is None:
            return (f"the {t.txn_date:%d-%m-%Y} line has no units - "
                    "the statement could not be read reliably")
        if t.txn_type in _UNITS_ONLY_TYPES:
            if t.amount not in (None, 0, 0.0):
                return (f"the {t.txn_date:%d-%m-%Y} bonus line carries an "
                        "amount - the statement could not be read reliably")
            if anchor is None:
                # units that no money proves and no balance can check are
                # unprovable — the contract says refuse, never trust
                return (f"the {t.txn_date:%d-%m-%Y} bonus/segregation line "
                        "can't be checked (the statement's closing balance "
                        "wasn't readable) - nothing was imported for this "
                        "fund")
        elif not triangle_ok(t.amount, t.units, t.nav):
            return (f"the {t.txn_date:%d-%m-%Y} line doesn't add up "
                    "(amount ≠ units × NAV) - the statement could not be "
                    "read reliably")
        # sign must agree with the transaction TYPE — a gate here, not in
        # any one parser, so every future statement format is covered
        if (t.txn_type in ("REDEMPTION", "SWITCH_OUT")
                and ((t.amount or 0) > 0 or t.units > 0)):
            return (f"the {t.txn_date:%d-%m-%Y} sale line reads as money "
                    "coming in - the statement could not be read reliably")
        if (t.txn_type in ("PURCHASE", "SWITCH_IN", "DIV_REINVEST")
                and ((t.amount or 0) < 0 or t.units < 0)):
            return (f"the {t.txn_date:%d-%m-%Y} purchase line reads as "
                    "money going out - the statement could not be read "
                    "reliably")
        running += t.units
        if running < -UNIT_TOLERANCE:
            return (f"more units sold than bought by {t.txn_date:%d-%m-%Y} "
                    "- the history is out of order or incomplete")
        total += t.units
    if anchor is not None and abs(total - anchor) > UNIT_TOLERANCE:
        return (f"the transactions sum to {total:.3f} units but the "
                f"statement says {anchor:.3f} - some lines could not be "
                "read, so nothing was imported for this fund")
    return ""


def merge_sip_batches(data, batches: list[ImportBatch],
                      owner_map: dict[str, str], today: date,
                      *, replace: bool = False,
                      allow_condense: bool = False) -> MergeReport:
    """Merge every batch's MF transactions into data.sip / data.mutual_funds.

    Atomic: validates and stages everything first, mutates `data` only at
    the very end; a capacity problem defers the WHOLE import untouched —
    unless the user agreed up front (allow_condense, interactive-only), in
    which case the OLDEST whole financial years are rolled into one
    opening line per fund (condense_txns: totals conserved, so every
    reconciliation still proves itself) using the SMALLEST cutoff that
    fits. Headless runs never condense.
    """
    rep = MergeReport()
    isin_by_scheme = {s: i for _f, s, i in data.masters.mf_rows}
    by_isin = {i: (f, s) for f, s, i in data.masters.mf_rows if i}
    persons = set(data.persons)

    # ---- group txns per fund = (folio, isin), keep statement order ----
    funds: dict[tuple[str, str], list[ImportedSipTxn]] = {}
    anchors: dict[tuple[str, str], float] = {}
    # ISINs where the parser refused ANOTHER folio (mid-history, unreadable
    # balance): the statement can't speak for the whole fund, so replacing
    # or extending the fund from the folios that DID parse could silently
    # lose the refused folio's history. Refused ISIN-wide — conservative
    # on purpose (the refused folio's owner is usually unknowable here).
    partial_isins: set[str] = set()
    for b in batches:
        rep.warnings.extend(b.warnings)
        partial_isins.update(isin for _folio, isin in b.partial)
        for t in b.sip_txns:
            funds.setdefault((t.folio, t.isin), []).append(t)
        anchors.update(b.closing_units)

    # ---- gates per fund, then stage accepted rows ----
    staged: dict[tuple[str, str], list] = {}     # (owner, isin) -> SIPRows
    lines: dict[tuple[str, str], FundLine] = {}
    fund_anchor: dict[tuple[str, str], float] = {}   # Σ folio anchors
    unanchored: set[tuple[str, str]] = set()
    for (folio, isin), txns in sorted(funds.items()):
        txns.sort(key=lambda t: (t.txn_date or today))
        scheme_stmt = txns[0].scheme_name
        line = FundLine(folio=folio, isin=isin, scheme=scheme_stmt)
        rep.funds.append(line)
        owner = owner_map.get(folio, "")
        if owner not in persons:
            line.ok, line.reason = False, (
                "folio not matched to a person - left out. Run again to "
                "be asked, or fix the Owner on the Import_Map sheet")
            continue
        line.owner = owner
        if isin in partial_isins:
            line.ok, line.reason = False, (
                "another folio of this fund couldn't be read since "
                "inception - importing only part of a fund could lose or "
                "double history, so it was left alone")
            continue
        reason = _validate_fund(txns, anchors.get((folio, isin)), today)
        if reason:
            line.ok, line.reason = False, reason
            continue
        line.reconciled = (folio, isin) in anchors
        # master name wins so the sheet's INDEX/MATCH lookups resolve;
        # unknown ISINs keep the statement name + an isin_override (which
        # also marks the ISIN as referenced, surviving master refresh)
        fund_house, scheme = by_isin.get(isin, ("", ""))
        rows = []
        for t in txns:
            rows.append(M.SIPRow(
                owner=owner,
                scheme=scheme or t.scheme_name,
                txn_date=t.txn_date,
                amount=round(t.amount or 0.0, 2),
                nav=t.nav,
                units_override=round(t.units, 4) if t.units is not None
                else None,
                fund_house_override="" if scheme else t.fund_house,
                isin_override="" if scheme else isin))
            line.invested += max(t.amount or 0.0, 0.0)
            line.units += t.units or 0.0
        key = (owner, isin)
        staged.setdefault(key, []).extend(rows)
        # the post-merge invariant needs the SUM of anchors when several
        # folios feed one (owner, fund) — a single folio's anchor would
        # false-alarm on the combined sheet total
        if line.reconciled:
            fund_anchor.setdefault(key, 0.0)
            fund_anchor[key] += anchors[(folio, isin)]
        else:
            unanchored.add(key)
        if key not in lines:
            lines[key] = line
        else:                                    # same fund, second folio
            lines[key].units += line.units

    if not staged:
        return rep

    # ---- apply: statement-wins or multiset append, staged first --------
    new_sip = list(data.sip)
    for (owner, isin), rows in staged.items():
        line = lines[(owner, isin)]
        existing_idx = [i for i, r in enumerate(new_sip)
                        if _sip_row_key(r, isin_by_scheme)[:2] == (owner, isin)]
        stmt_keys = sorted(_sip_row_key(r, isin_by_scheme) for r in rows)
        have_keys = sorted(_sip_row_key(new_sip[i], isin_by_scheme)
                           for i in existing_idx)
        if stmt_keys == have_keys:               # nothing new — idempotent
            line.skipped = len(rows)
            rep.sip_skipped += len(rows)
            continue
        if replace and existing_idx:
            # the delete below is folio-blind (the sheet stores no folio),
            # so it may only run when the statement plausibly covers
            # EVERYTHING the sheet holds for this fund. More money typed
            # than the whole statement shows = a folio the statement
            # doesn't know about - keep the typed rows and say so.
            typed_net = sum(new_sip[i].amount or 0.0 for i in existing_idx)
            stmt_net = sum(r.amount or 0.0 for r in rows)
            if stmt_net > 0 and typed_net > stmt_net * 1.05 + 1000:
                line.ok, line.reason = False, (
                    f"your sheet records ₹{typed_net:,.0f} in this fund "
                    f"but the whole statement shows ₹{stmt_net:,.0f} - "
                    "the statement may be missing a folio (another "
                    "account), so your typed rows were kept")
                continue
            for i in reversed(existing_idx):
                del new_sip[i]
            line.replaced = len(existing_idx)
            rep.sip_replaced += len(existing_idx)
            new_sip.extend(rows)
            line.added = len(rows)
            rep.sip_added += len(rows)
        else:
            # append-only multiset upsert: each existing (owner, isin,
            # date, amount) occurrence absorbs one statement row
            have = Counter(have_keys)
            for r in rows:
                k = _sip_row_key(r, isin_by_scheme)
                if have[k] > 0:
                    have[k] -= 1
                    line.skipped += 1
                    rep.sip_skipped += 1
                else:
                    new_sip.append(r)
                    line.added += 1
                    rep.sip_added += 1

    # ---- post-merge invariant: replaced funds must equal the anchors ----
    for (owner, isin), line in lines.items():
        if not (line.ok and line.reconciled and replace):
            continue
        # only when EVERY contributing folio declared a balance can the
        # sheet total be held to it
        if (owner, isin) in unanchored:
            continue
        anchor = fund_anchor.get((owner, isin))
        if anchor is None:
            continue
        total = sum(
            (r.units_override if r.units_override is not None else
             (r.amount / r.nav if r.amount and r.nav and r.nav > 0 else 0.0))
            for r in new_sip
            if _sip_row_key(r, isin_by_scheme)[:2] == (owner, isin))
        if abs(total - anchor) > UNIT_TOLERANCE:
            rep.warnings.append(
                f"{line.scheme}: the sheet now holds {total:.3f} units but "
                f"the statement says {anchor:.3f} - other rows of this fund "
                "exist on the sheet; check for doubled history")

    # ---- MutualFunds summary rows for new (owner, scheme) pairs --------
    new_mf = list(data.mutual_funds)
    have_mf = {(r.owner, r.scheme) for r in new_mf}
    for (owner, isin), rows in staged.items():
        if not lines[(owner, isin)].added:
            continue
        scheme = rows[0].scheme
        if (owner, scheme) not in have_mf:
            new_mf.append(M.MFRow(owner=owner, scheme=scheme,
                                  isin_override=rows[0].isin_override))
            have_mf.add((owner, scheme))

    # ---- capacity: defer the WHOLE import, workbook untouched ----------
    sip_cap = M.SIP_LAST_ROW - M.FIRST_DATA_ROW + 1
    mf_cap = M.MF_LAST_ROW - M.FIRST_DATA_ROW + 1
    if len(new_sip) > sip_cap:
        if allow_condense:
            attempt = _condensed_retry(data, batches, owner_map, today,
                                       replace, sip_cap)
            if attempt is not None:
                return attempt
        rep.deferred.append(
            f"the import would need {len(new_sip)} MF_SIP rows but the "
            f"sheet can only save {sip_cap} - nothing was imported. "
            "Move very old rows to another file, or import fewer files "
            "at once, and run again.")
        rep.sip_added = rep.sip_replaced = rep.sip_skipped = 0
        for line in rep.funds:
            if line.ok:
                line.added = line.skipped = line.replaced = 0
                line.ok, line.reason = False, "deferred - sheet full"
        return rep
    if len(new_mf) > mf_cap:
        rep.deferred.append(
            f"the import would need {len(new_mf)} MutualFunds rows but the "
            f"sheet can only save {mf_cap} - nothing was imported.")
        rep.sip_added = rep.sip_replaced = rep.sip_skipped = 0
        for line in rep.funds:
            if line.ok:
                line.added = line.skipped = line.replaced = 0
                line.ok, line.reason = False, "deferred - sheet full"
        return rep

    data.sip = new_sip
    data.mutual_funds = new_mf
    return rep


# ---- equity: broker trades → lots, netted FIFO (SPEC §6.18) ----------------

def _lot_key(owner, isin, cost_date, qty) -> tuple:
    return (owner, isin, cost_date, round(qty or 0.0, 3))


def _sell_key(owner, isin, sell_date, qty) -> tuple:
    return (owner, isin, sell_date, round(qty or 0.0, 3))


def _eq_isin(row, isin_by_name) -> str:
    return row.isin_override or isin_by_name.get(row.scrip, "")


# exchange series tokens brokers append to symbols ("AJMERA EQ") — the
# stock master keys on the bare symbol
_SERIES_SUFFIXES = ("EQ", "BE", "BZ", "BL", "SM", "ST", "XT", "GB", "IV",
                    "N1", "N2", "N3")


def _symbol_isin(symbol, isin_by_symbol) -> str:
    s = (symbol or "").upper().strip()
    if s in isin_by_symbol:
        return isin_by_symbol[s]
    parts = s.rsplit(None, 1)
    if len(parts) == 2 and parts[1] in _SERIES_SUFFIXES:
        return isin_by_symbol.get(parts[0], "")
    return ""


def _opening_ok(openings, owner: str, isin: str, symbols) -> bool:
    """Did the user confirm pre-2018 for this stock? Keys may be the ISIN
    or (when the file had no ISIN column) the raw symbol."""
    if not openings:
        return False
    if (owner, isin) in openings:
        return True
    return any((owner, (sym or "").upper().strip()) in openings
               for sym in symbols)


FMV_DATE = date(2018, 1, 31)   # the grandfathering cut-off (§6.6)


def merge_equity_batches(data, batches: list[ImportBatch],
                         owner_map: dict[str, str], today: date,
                         *, cg_on: bool = False,
                         replace: bool = False,
                         pre2018_openings: set | None = None) -> MergeReport:
    """Merge broker trades/holdings into data.equity (+ data.equity_sells).

    Chronological FIFO per (owner, ISIN): buys append lots, sells consume
    them — imported lots first, then (only with replace=True, i.e. after a
    human confirmed) the sheet's typed lots. Refusal gates per ISIN keep a
    half-read file or a split-crossing history from ever landing. Sells
    become Equity_Sells rows only while the Capital-gains switch is on.
    Re-import is stateless-idempotent: an ISIN whose computed OUTCOME
    (surviving lots + sells) is already on the sheet is skipped whole —
    that check is what makes re-running a file a no-op (netting must never
    replay). The per-lot multiset absorb further down has a different job:
    it lets lots the user ALREADY typed by hand stand in for identical
    imported ones. SIP uses a per-row multiset instead because statement
    rows ARE the outcome there — no derivation between file and sheet.
    """
    rep = MergeReport()
    isin_by_name = {name: isin for _s, name, isin in data.masters.stock_rows}
    name_by_isin = {isin: name for _s, name, isin in data.masters.stock_rows
                    if isin}
    isin_by_symbol = {sym.upper(): isin
                      for sym, _n, isin in data.masters.stock_rows if isin}
    persons = set(data.persons)

    # ---- pool + resolve identity ----
    groups: dict[tuple[str, str], list] = {}
    bad_idents: list[str] = []
    src_by_group: dict[tuple[str, str], str] = {}
    for b in batches:
        rep.warnings.extend(b.warnings)
        # an unreadable line whose stock can't even be identified can't
        # poison its ISIN's group (there is no group to poison) — the only
        # safe scope left is the whole file's trades. Fail safe, loudly.
        if any(t.side == "BAD"
               and not (t.isin or _symbol_isin(t.symbol, isin_by_symbol))
               for t in b.trades):
            rep.warnings.append(
                f"{b.source}: a line couldn't be read well enough to even "
                "tell which stock it belongs to - no shares were imported "
                "from this file. Fix or re-export the file and try again.")
            continue
        for t in b.trades:
            isin = t.isin or _symbol_isin(t.symbol, isin_by_symbol)
            if not isin:
                if t.symbol and t.symbol not in bad_idents:
                    bad_idents.append(t.symbol)
                continue
            owner = owner_map.get(t.account, "")
            groups.setdefault((owner, isin), []).append(t)
            src_by_group.setdefault((owner, isin), b.source)
    for sym in bad_idents:
        rep.warnings.append(
            f"'{sym}' isn't in the stock list and the file has no ISIN "
            "for it - its trades were left out. Add the ISIN column to "
            "the export, or add the holding by hand.")

    new_equity = list(data.equity)
    new_sells = list(data.equity_sells)
    # sell-driven cuts to TYPED rows are only PLANNED here — the row
    # objects are shared with data.equity, so mutating them before the
    # capacity gate would leak into a deferred (supposedly untouched)
    # workbook. Applied at the very end, after every gate has passed.
    pending_cuts: dict[int, tuple] = {}    # id(row) -> (row, qty consumed)

    def _eff_qty(r) -> float:
        cut = pending_cuts.get(id(r))
        return (r.qty or 0.0) - (cut[1] if cut else 0.0)

    for (owner, isin), trades in sorted(groups.items()):
        disp = name_by_isin.get(isin, isin)
        line = FundLine(owner=owner, isin=isin, scheme=disp)
        rep.stocks.append(line)
        if owner not in persons:
            line.ok, line.reason = False, (
                "account not matched to a person - left out. Run again to "
                "be asked, or fix the Owner on the Import_Map sheet")
            continue
        if any(t.side == "BAD" for t in trades):
            line.ok, line.reason = False, (
                "a line of this stock couldn't be read reliably - nothing "
                "was imported for it")
            continue
        trades.sort(key=lambda t: (t.trade_date, t.side))
        window = (trades[0].trade_date, trades[-1].trade_date)
        ca_hit = next(
            (a for a in data.corporate_actions
             if a.isin == isin and a.ex_date
             and window[0] <= a.ex_date <= window[1]), None)
        if ca_hit is not None:
            line.ok, line.reason = False, (
                f"a {ca_hit.type.lower()} on {ca_hit.ex_date:%d-%m-%Y} falls "
                "inside this history - quantities before and after it don't "
                "compare, so nothing was imported for this stock. Add the "
                "current holding by hand instead.")
            continue

        # ---- FIFO replay ----
        lots: list[list] = []              # [date, qty, price] imported lots
        sheet_rows = [r for r in new_equity
                      if r.owner == owner and _eq_isin(r, isin_by_name) == isin]
        sheet_pool = sorted(
            (r for r in sheet_rows if r.qty),
            key=lambda r: (r.cost_date or date(1900, 1, 1)))
        # typed lots may carry a split/bonus factor context (§6.7): if any
        # corporate action postdates the OLDEST typed lot, broker (post-CA)
        # sale quantities don't compare with the typed (raw) quantities —
        # never net across that boundary
        sheet_ca_blocked = bool(sheet_pool) and any(
            a.isin == isin and a.ex_date
            and a.ex_date > min((r.cost_date or date(1900, 1, 1))
                                for r in sheet_pool)
            for a in data.corporate_actions)
        reductions: dict[int, float] = {}  # id(row) -> qty consumed
        pairs: list[tuple] = []            # (buy_date, buy_price, qty, sell)
        shortfall = None
        opened = 0.0                       # units supplied by the pre-2018
        for t in trades:                   # opening lot (Jay, 2026-07-18)
            if t.side == "BUY":
                lots.append([t.trade_date, t.qty, t.price])
                continue
            need = t.qty
            for lot in lots:
                if need <= 0:
                    break
                take = min(lot[1], need)
                if take > 0:
                    pairs.append((lot[0], lot[2], take, t))
                    lot[1] -= take
                    need -= take
            if need > 1e-9 and replace and not sheet_ca_blocked:
                for r in sheet_pool:
                    if need <= 0:
                        break
                    avail = (r.qty or 0.0) - reductions.get(id(r), 0.0)
                    take = min(avail, need)
                    if take > 0:
                        pairs.append((r.cost_date, r.avg_cost, take, t))
                        reductions[id(r)] = reductions.get(id(r), 0.0) + take
                        need -= take
            if need > 1e-9 and _opening_ok(
                    pre2018_openings, owner, isin,
                    (t.symbol for t in trades)):
                # the user confirmed these shares predate Feb 2018: the
                # missing buy becomes an opening lot at the grandfathering
                # date with a blank cost — the official 31-01-2018 value
                # fills it in (§6.6), exactly like a typed old holding
                pairs.append((FMV_DATE, None, need, t))
                opened += need
                need = 0.0
            if need > 1e-9:
                shortfall = t
                break
        if shortfall is not None:
            line.ok, line.reason = False, (
                f"the file sells more than it buys by "
                f"{shortfall.trade_date:%d-%m-%Y} - the history is "
                "incomplete. Import the earlier tradebook file(s) in the "
                "same run"
                + (", or answer yes when asked about pre-2018 shares"
                   if pre2018_openings is not None else "")
                + (" (a corporate action also affects your typed rows of "
                   "this stock, so they can't cover the sale - adjust "
                   "them by hand)" if sheet_ca_blocked and replace else "")
                + ", or add the holding by hand.")
            continue
        if reductions and not cg_on:
            # consuming a typed lot leaves no trace unless the sale lands
            # on Equity_Sells — and without that record a re-import would
            # shrink the same lot again. Correct-or-refuse, never silently
            # wrong twice.
            line.ok, line.reason = False, (
                "this file sells shares that are typed on your Equity "
                "sheet. Turn ON 'Capital gains report' in Settings so the "
                "sale can be recorded, then import this file again - or "
                "adjust the typed rows by hand.")
            continue
        if opened:
            rep.warnings.append(
                f"{disp}: {opened:g} share(s) sold without a matching buy "
                "were taken as held from before Feb 2018"
                + (" - the official 31-01-2018 value stands in as their "
                   "cost (amber)" if cg_on else
                   "; the sale is netted off. Turn on 'Capital gains "
                   "report' in Settings first if you want it recorded"))

        survivors = [(d, q, p) for d, q, p in lots if q > 1e-9]
        cand_lots = [M.EquityRow(
            owner=owner, scrip=name_by_isin.get(isin, "") or
            (trades[0].symbol or isin),
            isin_override="" if isin in name_by_isin else isin,
            qty=round(q, 3), avg_cost=round(p, 4), cost_date=d,
            flag=f"IMPORTED:{src_by_group[(owner, isin)]}")
            for d, q, p in survivors]
        cand_sells = []
        if cg_on:
            for buy_date, buy_price, qty, t in pairs:
                cand_sells.append(M.EquitySellRow(
                    owner=owner, scrip=name_by_isin.get(isin, "") or
                    (trades[0].symbol or isin),
                    isin_override="" if isin in name_by_isin else isin,
                    qty=round(qty, 3), buy_date=buy_date,
                    buy_price=round(buy_price, 4)
                    if buy_price is not None else None,
                    sell_date=t.trade_date, sell_price=t.price,
                    notes="imported"))

        # ---- stateless idempotency: outcome already on the sheet? ----
        have_lots = Counter(_lot_key(r.owner, _eq_isin(r, isin_by_name),
                                     r.cost_date, r.qty) for r in sheet_rows)
        want_lots = Counter(_lot_key(r.owner, isin, r.cost_date, r.qty)
                            for r in cand_lots)
        have_sells = Counter(
            _sell_key(s.owner, _eq_isin(s, isin_by_name), s.sell_date, s.qty)
            for s in new_sells)
        want_sells = Counter(_sell_key(s.owner, isin, s.sell_date, s.qty)
                             for s in cand_sells)
        if (not (want_lots - have_lots) and not (want_sells - have_sells)):
            line.skipped = len(cand_lots)
            rep.eq_skipped += len(cand_lots)
            continue

        # ---- commit this ISIN (typed-row cuts stay PLANNED, see above) ----
        for r in sheet_pool:
            took = reductions.get(id(r), 0.0)
            if took <= 0:
                continue
            pending_cuts[id(r)] = (r, took)
            rep.eq_reduced += 1
        # lots already present (e.g. typed by hand) absorb candidates —
        # judged at their post-cut quantity, exactly what commit will leave
        have_lots = Counter(_lot_key(r.owner, _eq_isin(r, isin_by_name),
                                     r.cost_date, _eff_qty(r))
                            for r in new_equity
                            if r.owner == owner
                            and _eq_isin(r, isin_by_name) == isin
                            and _eff_qty(r) > 0)
        for r in cand_lots:
            k = _lot_key(r.owner, isin, r.cost_date, r.qty)
            if have_lots[k] > 0:
                have_lots[k] -= 1
                line.skipped += 1
                rep.eq_skipped += 1
            else:
                new_equity.append(r)
                line.added += 1
                rep.eq_added += 1
                line.invested += (r.qty or 0) * (r.avg_cost or 0)
                line.units += r.qty or 0
        for s in cand_sells:
            k = _sell_key(s.owner, isin, s.sell_date, s.qty)
            if have_sells[k] > 0:
                have_sells[k] -= 1
            else:
                new_sells.append(s)
                rep.sells_added += 1

    # ---- holdings: cross-check + no-history fallback ----
    undated = 0
    for b in batches:
        for h in b.holdings:
            isin = h.isin or _symbol_isin(h.name, isin_by_symbol)
            owner = owner_map.get(h.account, "")
            if not isin:
                # never silent: broker holdings files mix in fund units,
                # NCDs and preference shares that aren't listed equity
                rep.warnings.append(
                    f"'{h.name}' isn't in the stock list and the file has "
                    "no ISIN for it - left out. Fund units come via your "
                    "fund statement (CAS); bonds/NCDs go on the Bonds "
                    "sheet by hand.")
                continue
            if owner not in persons:
                rep.warnings.append(
                    f"'{h.name}': account not matched to a person - "
                    "left out")
                continue
            rows = [r for r in new_equity
                    if r.owner == owner and _eq_isin(r, isin_by_name) == isin]
            traded = (owner, isin) in groups
            if rows or traded:
                sheet_qty = sum(_eff_qty(r) for r in rows)
                if abs(sheet_qty - (h.qty or 0.0)) > 0.001:
                    rep.warnings.append(
                        f"{h.name or isin}: the broker file shows "
                        f"{h.qty:g} but the sheet now holds {sheet_qty:g} "
                        "- check for missing history or a corporate action")
                continue
            disp = name_by_isin.get(isin, "") or h.name or isin
            costless = h.avg_cost is None
            pre2018 = costless and _opening_ok(pre2018_openings, owner,
                                               isin, [h.name])
            new_equity.append(M.EquityRow(
                owner=owner, scrip=disp,
                isin_override="" if isin in name_by_isin else isin,
                qty=h.qty, avg_cost=h.avg_cost,
                cost_date=FMV_DATE if pre2018 else None,
                flag=f"IMPORTED:{b.source}"))
            rep.eq_added += 1
            rep.stocks.append(FundLine(
                owner=owner, isin=isin, scheme=disp, added=1,
                reason="", units=h.qty or 0.0,
                invested=(h.qty or 0.0) * (h.avg_cost or 0.0)))
            if pre2018:
                # the run's FMV pass fills the official 31-01-2018 value
                rep.warnings.append(
                    f"{disp}: the broker has no cost for this old holding "
                    "- taken as held from before Feb 2018, the official "
                    "2018 value stands in (amber)")
            elif costless:
                rep.warnings.append(
                    f"{disp}: the broker file has no buy price - the cost "
                    "is left blank (amber). Old paper shares? Put "
                    "31-01-2018 as the Buy date and the official value "
                    "fills in")
            else:
                undated += 1

    if undated:
        rep.warnings.append(
            f"{undated} holding(s) added with the broker's average cost "
            "but no buy date - values are right today; fill Buy dates "
            "when you want the exact tax view and return figure")

    # ---- capacity: defer the WHOLE equity import, workbook untouched ----
    # (cuts are still only planned, so returning here really does leave
    # data.equity byte-for-byte as it was)
    eq_final = (len(new_equity)
                - sum(1 for r, took in pending_cuts.values()
                      if (r.qty or 0.0) - took <= 0))
    eq_cap = M.EQUITY_LAST_ROW - M.FIRST_DATA_ROW + 1
    sell_cap = M.EQSELL_LAST_ROW - M.FIRST_DATA_ROW + 1
    if eq_final > eq_cap or len(new_sells) > sell_cap:
        sheet = "Equity" if eq_final > eq_cap else "Equity_Sells"
        n = eq_final if sheet == "Equity" else len(new_sells)
        cap = eq_cap if sheet == "Equity" else sell_cap
        rep.deferred.append(
            f"the import would need {n} {sheet} rows but the sheet can "
            f"only save {cap} - no shares were imported. Move old rows to "
            "another file and run again.")
        rep.eq_added = rep.eq_skipped = rep.eq_reduced = rep.sells_added = 0
        for line in rep.stocks:
            if line.ok:
                line.added = line.skipped = 0
                line.ok, line.reason = False, "deferred - sheet full"
        return rep

    # ---- every gate passed: NOW the shared rows may change ----
    for r, took in pending_cuts.values():
        r.qty = round((r.qty or 0.0) - took, 3)
        if r.qty <= 0:
            new_equity.remove(r)
    data.equity = new_equity
    data.equity_sells = new_sells
    return rep


# ---- condensing (the over-cap fallback the user consents to up front) ------

def _condensed_retry(data, batches, owner_map, today, replace, sip_cap):
    """Over-cap + user consent: retry the merge with the oldest whole
    financial years rolled up, using the SMALLEST cutoff that fits (least
    detail lost). Each trial is a full re-run of every gate — a condensed
    fund still has to prove itself. Returns the successful report, or None
    (caller falls back to the plain deferral). Never touches the caller's
    batches; `data` is only mutated by a SUCCESSFUL trial (the recursive
    call's own end-of-function assignment)."""
    from dataclasses import replace as dc_replace
    years = sorted({t.txn_date.year for b in batches
                    for t in b.sip_txns if t.txn_date})
    if not years:
        return None
    # candidate cutoffs at FY starts, least condensing first; the current
    # financial year is never rolled up
    this_fy = today.year if today.month >= 4 else today.year - 1
    for y in range(years[0] + 1, this_fy + 1):
        cutoff = date(y, 4, 1)
        trial = [dc_replace(b, sip_txns=condense_txns(b.sip_txns, cutoff))
                 for b in batches]
        rep = merge_sip_batches(data, trial, owner_map, today,
                                replace=replace, allow_condense=False)
        if not rep.deferred:
            rep.warnings.append(
                f"the full history wouldn't fit on the MF_SIP sheet, so "
                f"transactions before {cutoff:%b %Y} were rolled into one "
                "opening line per fund (as you agreed) - totals still "
                "match the statement's closing balances; the return "
                "figure treats those old years approximately")
            return rep
    return None


def condense_txns(txns: list[ImportedSipTxn], cutoff: date
                  ) -> list[ImportedSipTxn]:
    """Per fund, roll transactions dated before `cutoff` into ONE opening
    purchase: Σamount (net), Σunits, NAV = amount/units (so the triangle
    holds by construction), date = |amount|-weighted mean (first-order
    preserves the XIRR duration). Conserves both totals, so the balance
    reconciliation still proves the fund. A fund whose old rows net to a
    non-positive amount or units is left verbatim — condensing it would
    misstate history. The caller owns telling the user that condensed
    years make the return figure approximate and blur FIFO cost basis.
    """
    by_fund: dict[tuple[str, str], list[ImportedSipTxn]] = {}
    for t in txns:
        by_fund.setdefault((t.folio, t.isin), []).append(t)
    out: list[ImportedSipTxn] = []
    for (folio, isin), rows in sorted(by_fund.items()):
        old = [t for t in rows if t.txn_date and t.txn_date < cutoff]
        keep = [t for t in rows if not (t.txn_date and t.txn_date < cutoff)]
        units = sum(t.units or 0.0 for t in old)
        amount = sum(t.amount or 0.0 for t in old)
        if len(old) < 2 or units <= 0 or amount <= 0:
            out.extend(rows)
            continue
        weights = sum(abs(t.amount or 0.0) for t in old) or 1.0
        mean_ord = round(sum((t.amount and abs(t.amount) or 0.0)
                             * t.txn_date.toordinal()
                             for t in old) / weights)
        first = old[0]
        out.append(ImportedSipTxn(
            folio=folio, isin=isin, scheme_name=first.scheme_name,
            fund_house=first.fund_house,
            txn_date=date.fromordinal(mean_ord),
            amount=round(amount, 2), nav=round(amount / units, 4),
            units=round(units, 4), txn_type="OPENING"))
        out.extend(keep)
    return out
