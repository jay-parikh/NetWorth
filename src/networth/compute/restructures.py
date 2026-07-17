"""Restructure engine (SPEC §6.15) — mergers / demergers / ISIN changes.

Extracted from update.run() so each phase is a unit-testable function that
maps onto the spec: curated intake (events fold into the workbook's actions
table), successor labelling, and the append-once demerger child creation
with its safety gates. Price ROUTING for consumed ISINs stays in the
updater's pricing loop (it is one line of resolve_isin); everything else
about restructures lives here.
"""

from __future__ import annotations

from datetime import date, timedelta

from .. import model as M
from ..model import (RESTRUCTURE_TYPES, CorporateAction, EquityRow,
                     PortfolioData, chained_adjustment_factor,
                     cost_adjustment_factor, resolve_isin)


def integrate_restructures(data: PortfolioData,
                           restructures: list[CorporateAction],
                           today: date) -> list[CorporateAction]:
    """Fold the curated file into the workbook's actions and return the
    ACTIVE restructure events (ex-date arrived).

    Rules (SPEC §6.15): only events touching a held security — directly or
    via a restructure chain — surface on the audit sheet; a Manual row with
    the same (isin, type, ex_date, new_isin) key overrides a Curated one;
    a Curated row's Applied date survives the rewrite; successor securities
    join Stock_Master immediately (add-only safe: new ISINs) so name
    lookups resolve before listing.
    """
    isin_by_name = {name: isin for _s, name, isin in data.masters.stock_rows}
    held = {row.isin_override or isin_by_name.get(row.scrip, "")
            for row in data.equity}
    held.discard("")
    # chain-aware: a demerger/merger on a SUCCESSOR of a held ISIN concerns
    # this workbook too (the held lots resolve into it)
    held |= {resolve_isin(i, restructures, today) for i in held}
    # new_isin is part of the key: a demerger's retention row and child rows
    # share (isin, type, ex_date) and must track Applied independently
    manual_keys = {(a.isin, a.type, a.ex_date, a.new_isin)
                   for a in data.corporate_actions if a.source == "Manual"}
    prev_applied = {(a.isin, a.type, a.ex_date, a.new_isin): a.applied
                    for a in data.corporate_actions if a.applied}
    curated = [c for c in restructures
               if c.isin in held
               and (c.isin, c.type, c.ex_date, c.new_isin) not in manual_keys]
    for c in curated:                        # Applied survives the rewrite
        c.applied = c.applied or prev_applied.get(
            (c.isin, c.type, c.ex_date, c.new_isin))
    data.corporate_actions = ([a for a in data.corporate_actions
                               if a.source != "Curated"] + curated)
    events = [a for a in data.corporate_actions
              if a.type in RESTRUCTURE_TYPES and a.ex_date
              and a.ex_date <= today]

    if events:
        known = {i for _s, _n, i in data.masters.stock_rows}
        fresh = [(a.new_symbol or a.new_isin,
                  # display-name fallback: a symbol reads far better than a
                  # raw ISIN when a Manual row omits the name
                  a.new_name or a.new_symbol or a.new_isin,
                  a.new_isin)
                 for a in events
                 if a.new_isin and a.new_isin != a.isin
                 and a.new_isin not in known]
        if fresh:
            data.masters.stock_rows.extend(fresh)
            data.masters.stock_rows.sort(key=lambda r: r[1].casefold())
    return events


def consumed_label(isin: str, events: list[CorporateAction]) -> str | None:
    """"Merged" / "Renamed" when the ISIN was consumed by a restructure,
    else None. Drives Stock_Master status, the Flags column and the
    §6.5-escalation exemption."""
    hops = [a for a in events
            if a.isin == isin and a.type in ("MERGER", "ISIN_CHANGE")
            and a.new_isin and a.new_isin != isin]
    if not hops:
        return None
    return "Merged" if any(h.type == "MERGER" for h in hops) else "Renamed"


def _event_name(ev: CorporateAction) -> str:
    return ev.new_name or ev.new_symbol or ev.new_isin


def apply_demergers(data: PortfolioData, events: list[CorporateAction], *,
                    ca_checked: set[str], ca_trusted: bool,
                    price_data, today: date) -> tuple[int, list[str]]:
    """Append-once demerger child rows (SPEC §6.15) → (children added,
    warnings).

    Must run AFTER the corporate-actions refresh so quantities come from
    the full split/bonus history. The persisted Applied date is the
    idempotency token — re-runs skip applied events, so a user deleting a
    child row is respected. Safety gates: an event whose parent the feed
    could not verify this run (unless `ca_trusted`, i.e. injected data), or
    whose children would overflow the Equity sheet, is NOT stamped —
    children land atomically per event and the event retries next update.
    """
    children_added = 0
    warnings: list[str] = []
    unstamped: set[int] = set()
    equity_cap = M.EQUITY_LAST_ROW - M.FIRST_DATA_ROW + 1
    if not events:
        return 0, warnings
    isin_by_name = {name: isin for _s, name, isin in data.masters.stock_rows}
    for ev in events:
        if (ev.type != "DEMERGER" or ev.applied
                or not ev.new_isin or ev.new_isin == ev.isin):
            continue
        if not ca_trusted and ev.isin not in ca_checked:
            # applying now would freeze child quantities computed from an
            # unverified (possibly empty) actions table — defer
            unstamped.add(id(ev))
            warnings.append(
                f"demerger of {_event_name(ev)} deferred: "
                f"{ev.symbol or ev.isin}'s split/bonus history could not "
                f"be verified this run — it will apply on the next update")
            continue
        eve = ev.ex_date - timedelta(days=1)
        # children land ATOMICALLY per event: a half-applied event that is
        # retried next run would otherwise duplicate its early rows
        new_children: list[EquityRow] = []
        fits = True
        for lot in data.equity:
            lot_isin = lot.isin_override or isin_by_name.get(lot.scrip, "")
            if not lot_isin or not lot.qty:
                continue
            # chain-aware: a lot still keyed to a merged-away ISIN that
            # resolved into ev.isin by the ex-date demerges too
            if (lot_isin != ev.isin and resolve_isin(
                    lot_isin, data.corporate_actions, eve) != ev.isin):
                continue
            if lot.cost_date and lot.cost_date >= ev.ex_date:
                continue
            qty_adj = lot.qty * chained_adjustment_factor(
                lot_isin, lot.cost_date, eve, data.corporate_actions)
            ratio = (ev.ratio_from / ev.ratio_to
                     if ev.ratio_from and ev.ratio_to else 1.0)
            child_qty = round(qty_adj * ratio, 4)
            if child_qty <= 0:
                continue
            if len(data.equity) + len(new_children) >= equity_cap:
                unstamped.add(id(ev))
                warnings.append(
                    f"Equity sheet is full ({equity_cap} rows) — could not "
                    f"append the demerged {_event_name(ev)} row(s); free up "
                    f"rows and run the updater again")
                fits = False
                break
            child_cost = None
            if lot.avg_cost and ev.cost_pct is not None:
                # the cost available to apportion is the original cost × the
                # retention of every EARLIER demerger on this chain
                prior_cf = cost_adjustment_factor(
                    lot_isin, lot.cost_date, eve, data.corporate_actions)
                child_cost = round(lot.qty * lot.avg_cost * prior_cf
                                   * ev.cost_pct / 100 / child_qty, 4)
            child = EquityRow(
                owner=lot.owner, scrip=_event_name(ev),
                isin_override=ev.new_isin, qty=child_qty,
                avg_cost=child_cost,
                # Indian CGT: demerged shares inherit the holding period
                cost_date=lot.cost_date,
                flag=f"DEMERGER:{ev.isin}@{ev.ex_date.isoformat()}")
            # price the child in the same run (the bhavcopy is fetched)
            quote = price_data.prices.get(ev.new_isin) if price_data else None
            if quote:
                child.close = quote["close"]
                child.prev_close = quote["prev"]
                child.close_date = price_data.trade_date or today
            new_children.append(child)
        if fits:
            data.equity.extend(new_children)
            children_added += len(new_children)
    for ev in events:                            # stamp the audit trail
        if not ev.applied and id(ev) not in unstamped:
            ev.applied = today
    return children_added, warnings
