"""Updater — one command replaces the legacy three .bat scripts (SPEC §7).

Round trip: read inputs → fetch AMFI + bhavcopy → compute XIRR → regenerate
the workbook (backup first, atomic replace). Each source failure degrades
gracefully: previous values stay in place and the summary says so.
"""

from __future__ import annotations

import argparse
import glob
import os
import shutil
import sys
from datetime import date
from pathlib import Path

from . import __version__
from . import model as M
from .compute.cashflows import compute_all_xirr, ppf_ledger_by_account
from .compute.projections import fy_expected_by_person
from .compute.ppf import current_rate, load_ppf_rates, ppf_cashflows, ppf_value
from .compute.snapshot import net_worth_snapshot, upsert_snapshot
from .compute.xirr import xirr
from .fetch import amfi as amfi_mod
from .fetch import bhavcopy as bhav_mod
from .fetch import corporate_actions as ca_mod
from .model import (ASSET_CLASSES, MANUAL_CLASS_LABELS, RESTRUCTURE_TYPES,
                    DividendRow, EquityRow, chained_adjustment_factor,
                    class_has_data, cost_adjustment_factor, fy_label,
                    load_fmv, load_restructures, resolve_isin)
from .generate import build_workbook
from .reader import read_workbook

KEEP_BACKUPS = 10


# ------------------------------------------------------------- console UI --
# The updater is many users' ONLY window into what happened — it should feel
# alive, not like a log file. Colours degrade gracefully: plain text when the
# terminal can't (old cmd.exe, redirected output), emoji stripped when the
# console encoding can't carry them.

def _supports_color() -> bool:
    if os.environ.get("NO_COLOR"):
        return False
    if not hasattr(sys.stdout, "isatty") or not sys.stdout.isatty():
        return False
    if os.name == "nt":
        os.system("")          # enables ANSI escape processing on Windows 10+
    return True


_COLOR = _supports_color()


def _c(code: str, text: str) -> str:
    return f"\033[{code}m{text}\033[0m" if _COLOR else text


def _say(text: str = "", **kw) -> None:
    """print() that never crashes on a console that can't render emoji."""
    try:
        print(text, **kw)
    except UnicodeEncodeError:
        print(text.encode("ascii", "ignore").decode(), **kw)


def _step(text: str) -> None:
    _say(_c("36", f"  ⏳ {text}"))


def _banner(title: str) -> None:
    line = "─" * 62
    _say(_c("36", line))
    _say(_c("1;36", f"  💰 {title}"))
    _say(_c("36", line))


def _fail(msg: str) -> "SystemExit":
    print(f"ERROR: {msg}", file=sys.stderr)
    return SystemExit(1)


def locate_workbook(arg: str | None) -> Path:
    if arg:
        p = Path(arg)
        if not p.exists():
            raise _fail(f"workbook not found: {p}")
        return p
    default = Path(M.TEMPLATE_FILENAME)
    if default.exists():
        return default
    candidates = [p for p in glob.glob("*.xlsx") if not p.startswith("~$")]
    if len(candidates) == 1:
        return Path(candidates[0])
    raise _fail("no workbook given and no single .xlsx found here — "
                "run next to your tracker file or pass its path")


def ensure_closed(path: Path) -> None:
    try:
        with open(path, "r+b"):
            pass
    except PermissionError:
        raise _fail(f"{path.name} is open in Excel — close it and run again")
    lock = path.with_name("~$" + path.name)
    if lock.exists():
        raise _fail(f"{path.name} looks open in Excel (found {lock.name}) — "
                    "close it and run again")


def make_backup(path: Path) -> Path:
    from datetime import datetime
    bdir = path.parent / "backups"
    bdir.mkdir(exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    dest = bdir / f"{path.stem}.backup-{stamp}{path.suffix}"
    shutil.copy2(path, dest)
    backups = sorted(bdir.glob(f"{path.stem}.backup-*{path.suffix}"))
    for old in backups[:-KEEP_BACKUPS]:
        old.unlink()
    return dest


def _dividend_qty(owner: str, isin: str, ex: date, equity_rows,
                  isin_by_name: dict[str, str], actions) -> float:
    """Estimated shares held at the ex-date (SPEC §6.12): lots bought before
    the ex-date, at the CA-adjusted count as of the day before. A lot still
    keyed to a merged-away ISIN earns the SUCCESSOR's dividends — the chain
    resolves as of the ex-date and the merger ratio folds into the count.
    There is no sell ledger — the estimate projects the CURRENT rows
    backwards, which the sheet hint and Guide state plainly (hence the amber
    '(est.)' columns)."""
    from datetime import timedelta
    as_of = ex - timedelta(days=1)
    total = 0.0
    for row in equity_rows:
        row_isin = row.isin_override or isin_by_name.get(row.scrip, "")
        if not row_isin or row.owner != owner or not row.qty:
            continue
        if row_isin != isin and resolve_isin(row_isin, actions, as_of) != isin:
            continue
        if row.cost_date and row.cost_date >= ex:
            continue
        total += row.qty * chained_adjustment_factor(row_isin, row.cost_date,
                                                     as_of, actions)
    return total


def _merge_stock_master(existing: list[tuple[str, str, str]],
                        fetched: list[tuple[str, str, str]]) -> tuple[list, int]:
    """Add-only (SPEC §6.4): known ISINs keep their names so user rows survive."""
    known = {isin for _s, _n, isin in existing}
    merged = list(existing)
    added = 0
    for sym, name, isin in fetched:
        if isin not in known:
            merged.append((sym, name, isin))
            known.add(isin)
            added += 1
    merged.sort(key=lambda r: r[1].casefold())
    return merged, added


def _replace_mf_master(existing: list[tuple[str, str, str]],
                       fetched: list[tuple[str, str, str]],
                       referenced: set[str]) -> list[tuple[str, str, str]]:
    """Wholesale AMFI refresh, preserving ISINs a user row still references."""
    new_isins = {isin for _f, _s, isin in fetched}
    merged = list(fetched)
    merged.extend(row for row in existing
                  if row[2] in referenced and row[2] not in new_isins)
    merged.sort(key=lambda r: r[1].casefold())
    return merged


def run(path: Path, *, price_data=None, amfi_data=None, ca_data=None,
        div_data=None, bullion_rates=None, nps_data=None, restructures=None,
        add_persons: list[str] | None = None,
        toggle_classes: list[str] | None = None,
        today: date | None = None, do_backup: bool = True) -> dict:
    today = today or date.today()
    summary: dict = {"warnings": []}

    ensure_closed(path)
    data = read_workbook(str(path))

    # add new people (declaratively: they land in data.persons, so regeneration
    # creates their sheet, Dashboard row, By-Scrip column — everything)
    if add_persons:
        have = {p.casefold() for p in data.persons}
        added = []
        for name in add_persons:
            name = name.strip()
            if name and name.casefold() not in have and len(data.persons) < 10:
                data.persons.append(name)
                have.add(name.casefold())
                added.append(name)
        summary["persons_added"] = added

    # show/hide asset classes chosen at the console (v1.4) — writes the same
    # Settings the user could edit by hand, then regeneration applies it
    if toggle_classes:
        from .model import ClassSetting
        key_by_label = {c.label.casefold(): c.key for c in ASSET_CLASSES}
        label_by_key = {c.key: c.label for c in ASSET_CLASSES}
        toggled = []
        for name in toggle_classes:
            key = key_by_label.get(name.strip().casefold())
            if not key:
                continue
            s = data.class_settings.setdefault(key, ClassSetting())
            s.enabled = not s.enabled
            toggled.append(f"{label_by_key[key]} → "
                           f"{'shown' if s.enabled else 'hidden'}")
            # switched off while still holding rows: stays visible (never
            # lose data — SPEC §3.14). Warn at the moment of the toggle only;
            # No-with-data is otherwise the normal delete-to-hide state the
            # shipped sample starts in, and nagging every run would drown it.
            if not s.enabled and class_has_data(data, key):
                summary["warnings"].append(
                    f"{label_by_key[key]} still holds rows, so it stays "
                    f"visible — delete or move its rows to hide it")
        summary["classes_toggled"] = toggled

    # a Manual_Assets row with a Class the dropdown doesn't know lands in NO
    # class column: it shows in the sheet TOTAL but in no person/family total
    # (the dropdown is non-blocking by design — so say it, never guess)
    for r in data.manual_assets:
        if r.asset_class and r.asset_class not in MANUAL_CLASS_LABELS:
            summary["warnings"].append(
                f"Manual_Assets row '{r.description or r.owner}' has "
                f"unrecognised Class '{r.asset_class}' — it is counted in no "
                f"class total; pick Real Estate / Cash / Insurance / Other")

    if do_backup:
        summary["backup"] = str(make_backup(path))

    # ---- restructures (SPEC §6.15): curated file + Manual rows; the one
    # corporate-action category with no free feed. Runs BEFORE pricing so
    # consumed ISINs route to their successor and demerger children get
    # priced in the same run. A malformed curated file fails loudly.
    if restructures is None:
        try:
            restructures = load_restructures()
        except OSError:
            restructures = []
    _isin_by_name = {name: isin for _s, name, isin in data.masters.stock_rows}
    _held = {row.isin_override or _isin_by_name.get(row.scrip, "")
             for row in data.equity}
    _held.discard("")
    # chain-aware: a demerger/merger on a SUCCESSOR of a held ISIN concerns
    # this workbook too (the held lots resolve into it)
    _held |= {resolve_isin(i, restructures, today) for i in _held}
    # new_isin is part of the key: a demerger's retention row and child rows
    # share (isin, type, ex_date) and must track Applied independently
    manual_keys = {(a.isin, a.type, a.ex_date, a.new_isin)
                   for a in data.corporate_actions if a.source == "Manual"}
    prev_applied = {(a.isin, a.type, a.ex_date, a.new_isin): a.applied
                    for a in data.corporate_actions if a.applied}
    # only events touching a HELD security surface on the audit sheet — the
    # curated file covers everyone, the workbook shows what matters here
    curated = [c for c in restructures
               if c.isin in _held
               and (c.isin, c.type, c.ex_date, c.new_isin) not in manual_keys]
    for c in curated:                        # Applied survives the rewrite
        c.applied = c.applied or prev_applied.get(
            (c.isin, c.type, c.ex_date, c.new_isin))
    data.corporate_actions = ([a for a in data.corporate_actions
                               if a.source != "Curated"] + curated)
    restructure_events = [a for a in data.corporate_actions
                          if a.type in RESTRUCTURE_TYPES and a.ex_date
                          and a.ex_date <= today]

    def _consumed_label(isin: str) -> str | None:
        hops = [a for a in restructure_events
                if a.isin == isin and a.type in ("MERGER", "ISIN_CHANGE")
                and a.new_isin and a.new_isin != isin]
        if not hops:
            return None
        return "Merged" if any(h.type == "MERGER" for h in hops) else "Renamed"

    if restructure_events:
        # successor securities join the master immediately (add-only safe:
        # they are new ISINs) so name lookups resolve before listing
        known = {i for _s, _n, i in data.masters.stock_rows}
        fresh = [(a.new_symbol or a.new_isin, a.new_name or a.new_isin,
                  a.new_isin)
                 for a in restructure_events
                 if a.new_isin and a.new_isin != a.isin
                 and a.new_isin not in known]
        if fresh:
            data.masters.stock_rows.extend(fresh)
            data.masters.stock_rows.sort(key=lambda r: r[1].casefold())
    # demerger child rows are appended AFTER the corporate-actions refresh
    # below: child quantities need the FULL split/bonus history, and a fresh
    # workbook's sheet has none of it yet (§6.15)

    # ---- fetch (graceful per source; the _step lines keep the console alive
    # during the slow network minute) ----
    if price_data is None:
        _step("Fetching share & bond prices (BSE + NSE)…")
        try:
            price_data = bhav_mod.fetch(today=today)
        except Exception as e:  # noqa: BLE001 — any fetch failure degrades
            summary["warnings"].append(f"price fetch failed, keeping old prices: {e}")
    if amfi_data is None:
        _step("Fetching mutual-fund NAVs (AMFI)…")
        try:
            amfi_data = amfi_mod.fetch()
        except Exception as e:  # noqa: BLE001
            summary["warnings"].append(f"AMFI fetch failed, keeping old NAVs: {e}")

    stamp = today.strftime("%d-%m-%Y")

    # ---- equity prices + stock master + trading status (SPEC §6.5) ----
    if price_data:
        isin_by_name = {name: isin for _s, name, isin in data.masters.stock_rows}
        trade_date = price_data.trade_date or today
        # absence only means anything when BOTH exchanges answered (SPEC §6.5);
        # on a single-source day, unquoted rows carry their status forward
        dual_source = len(getattr(price_data, "sources", ())) >= 2
        nse_only = getattr(price_data, "nse_only", set())
        matched = 0
        nse_only_hits = 0
        suspended = 0
        name_by_isin_px = {isin: name for _s, name, isin in data.masters.stock_rows}
        for row in data.equity:
            isin = row.isin_override or isin_by_name.get(row.scrip, "")
            if not isin:
                continue
            # a consumed ISIN routes to its survivor (SPEC §6.15): the row
            # prices as the new security while the user's cells stay put
            routed = (resolve_isin(isin, restructure_events, today)
                      if restructure_events else isin)
            quote = price_data.prices.get(routed)
            if routed != isin:
                label = _consumed_label(isin) or "Merged"
                _st, last = data.masters.stock_status.get(isin, ("", None))
                data.masters.stock_status[isin] = (label, last or row.close_date)
                if not row.flag:
                    row.flag = (f"MERGED→{name_by_isin_px.get(routed, routed)}"
                                if label == "Merged" else
                                f"ISIN→{routed}")
            if quote:
                row.close = quote["close"]
                if quote["prev"]:
                    row.prev_close = quote["prev"]
                row.close_date = trade_date
                if routed == isin:
                    data.masters.stock_status[isin] = ("Active", trade_date)
                matched += 1
                nse_only_hits += routed in nse_only
            elif dual_source and routed == isin:
                # absent from both exchanges: escalate by how long it has been
                # (consumed ISINs are exempt — their absence is expected)
                prev_status, last = data.masters.stock_status.get(isin, ("", None))
                last = last or row.close_date
                if prev_status == "Delisted":
                    continue
                if last and (today - last).days > 180:
                    data.masters.stock_status[isin] = ("Delisted", last)
                elif last and (today - last).days > 21:
                    data.masters.stock_status[isin] = ("Suspended", last)
                    suspended += 1
                elif last:
                    data.masters.stock_status[isin] = (prev_status or "Active", last)
        summary["suspended"] = suspended
        summary["nse_only_matched"] = nse_only_hits
        data.masters.stock_rows, added = _merge_stock_master(
            data.masters.stock_rows, price_data.master_rows)
        data.masters.stock_refreshed = stamp
        summary["equity_matched"] = matched
        summary["equity_total"] = sum(1 for r in data.equity if r.scrip or r.isin_override)
        summary["stocks_added"] = added
        summary["price_source"] = f"{price_data.source} {trade_date.strftime('%d-%m-%Y')}"

        # bonds trade on the exchanges too — refresh when the ISIN is quoted
        bonds_matched = 0
        for b in data.bonds:
            quote = price_data.prices.get(b.isin)
            if quote:
                b.cur_price = quote["close"]
                bonds_matched += 1
        summary["bonds_matched"] = bonds_matched

    # ---- fund NAVs + MF master ----
    if amfi_data:
        isin_by_scheme = {scheme: isin for _f, scheme, isin in data.masters.mf_rows}
        matched = 0
        referenced: set[str] = set()
        for rows in (data.mutual_funds, data.sip):
            for row in rows:
                isin = row.isin_override or isin_by_scheme.get(row.scheme, "")
                if isin:
                    referenced.add(isin)
        for row in data.mutual_funds:
            isin = row.isin_override or isin_by_scheme.get(row.scheme, "")
            nav = amfi_data.nav_by_isin.get(isin)
            if nav:
                row.current_nav = nav
                matched += 1
        data.masters.mf_rows = _replace_mf_master(
            data.masters.mf_rows, amfi_data.master_rows, referenced)
        data.masters.mf_refreshed = stamp
        summary["mf_matched"] = matched
        summary["mf_total"] = len(data.mutual_funds)

    # ---- corporate actions (SPEC §6.7): fetch NSE+BSE, keep manual, factor rows,
    # and warn about any holding NEITHER exchange could verify — never skip silently
    isin_by_name = {name: isin for _s, name, isin in data.masters.stock_rows}
    symbol_by_isin = {isin: sym for sym, _n, isin in data.masters.stock_rows}
    name_by_isin = {isin: name for _s, name, isin in data.masters.stock_rows}
    held_isins = {row.isin_override or isin_by_name.get(row.scrip, "")
                  for row in data.equity}
    held_isins.discard("")
    # merged holdings live on as their successor (SPEC §6.15): query the
    # successor's feeds too, so ITS dividends and later splits arrive
    routed_isins = ({resolve_isin(i, data.corporate_actions, today)
                     for i in held_isins} if restructure_events else set())
    query_isins = held_isins | routed_isins
    div_skipped = 0
    ca_injected = ca_data is not None      # injected data is trusted as-is
    ca_checked: set[str] = set()           # ISINs verified by the live fetch
    preserve_isins: set[str] = set()       # feed failed → keep their old rows
    if ca_data is None:
        symbols = {symbol_by_isin[i]: i for i in query_isins if symbol_by_isin.get(i)}
        bse_codes = {code: isin
                     for isin, code in getattr(price_data, "codes_by_isin", {}).items()
                     if isin in query_isins} if price_data else {}
        try:
            checked: set[str] = set()
            if symbols or bse_codes:
                _step("Checking corporate actions & dividends (NSE + BSE)…")
                fy_start = date(today.year if today.month >= 4 else today.year - 1,
                                4, 1)
                ca_data, checked, fetched_divs, div_skipped = ca_mod.fetch(
                    symbols, bse_codes, div_skip_since=fy_start)
                if div_data is None:
                    div_data = fetched_divs
            else:
                ca_data = []
                if div_data is None:
                    div_data = []
            ca_checked = checked
            # a restructured (merged/renamed) ISIN is expected to be absent
            # from the CA feeds — the curated file IS its verification
            consumed = {i for i in held_isins if _consumed_label(i)}
            unchecked = query_isins - checked - consumed
            preserve_isins = unchecked
            summary["ca_unverified"] = sorted(
                name_by_isin.get(i, i) for i in unchecked)
            if unchecked:
                summary["warnings"].append(
                    "corporate actions could NOT be verified for: "
                    + ", ".join(summary["ca_unverified"])
                    + " — their existing rows are kept, but NEW splits/bonuses "
                      "may be missing; add any you know of as Manual rows on "
                      "the Corporate_Actions sheet")
        except Exception as e:  # noqa: BLE001
            summary["warnings"].append(
                f"corporate-actions fetch failed, keeping existing rows: {e}")
    else:
        summary["ca_unverified"] = []
    if ca_data is not None:
        manual = [a for a in data.corporate_actions if a.source != "Auto"]
        manual_keys = {(a.isin, a.type, a.ex_date) for a in manual}
        # a one-symbol feed failure must not revert that stock's already-
        # applied rows: Auto rows of unverified ISINs survive the rebuild
        kept_auto = [a for a in data.corporate_actions
                     if a.source == "Auto" and a.isin in preserve_isins]
        auto = kept_auto + [a for a in ca_data
                            if (a.isin, a.type, a.ex_date) not in manual_keys]
        auto.sort(key=lambda a: (a.symbol, a.ex_date or today))
        # Manual/Curated rows lead: they carry user data and the demerger
        # Applied stamps, so they must never fall past the sheet's last row
        ca_cap = M.CA_LAST_ROW - M.FIRST_DATA_ROW + 1
        overflow = len(manual) + len(auto) - ca_cap
        if overflow > 0:
            oldest = sorted(auto, key=lambda a: a.ex_date or today)[:overflow]
            drop_ids = {id(a) for a in oldest}
            auto = [a for a in auto if id(a) not in drop_ids]
            summary["warnings"].append(
                f"Corporate_Actions sheet is full — dropped the {overflow} "
                f"oldest Auto row(s); adjusted quantities of very old lots "
                f"may be affected")
        data.corporate_actions = manual + auto
        summary["ca_rows"] = len(data.corporate_actions)

    # ---- demerger children (SPEC §6.15): append-once child rows per lot,
    # AFTER the corporate-actions refresh so quantities come from the full
    # split/bonus history. The persisted Applied date is the idempotency
    # token — re-runs skip applied events, so a user deleting a child row is
    # respected. An event that cannot be applied safely this run (feed not
    # verified, sheet full) is NOT stamped and retries on the next update.
    children_added = 0
    unstamped: set[int] = set()
    equity_cap = M.EQUITY_LAST_ROW - M.FIRST_DATA_ROW + 1
    if restructure_events:
        from datetime import timedelta
        for ev in restructure_events:
            if (ev.type != "DEMERGER" or ev.applied
                    or not ev.new_isin or ev.new_isin == ev.isin):
                continue
            if not ca_injected and ev.isin not in ca_checked:
                # applying now would freeze child quantities computed from an
                # unverified (possibly empty) actions table — defer
                unstamped.add(id(ev))
                summary["warnings"].append(
                    f"demerger of {ev.new_name or ev.new_isin} deferred: "
                    f"{ev.symbol or ev.isin}'s split/bonus history could not "
                    f"be verified this run — it will apply on the next update")
                continue
            eve = ev.ex_date - timedelta(days=1)
            # children land ATOMICALLY per event: a half-applied event that
            # is retried next run would otherwise duplicate its early rows
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
                    summary["warnings"].append(
                        f"Equity sheet is full ({equity_cap} rows) — could "
                        f"not append the demerged "
                        f"{ev.new_name or ev.new_isin} row(s); free up rows "
                        f"and run the updater again")
                    fits = False
                    break
                child_cost = None
                if lot.avg_cost and ev.cost_pct is not None:
                    # the cost available to apportion is the original cost ×
                    # the retention of every EARLIER demerger on this chain
                    prior_cf = cost_adjustment_factor(
                        lot_isin, lot.cost_date, eve, data.corporate_actions)
                    child_cost = round(lot.qty * lot.avg_cost * prior_cf
                                       * ev.cost_pct / 100 / child_qty, 4)
                child = EquityRow(
                    owner=lot.owner, scrip=ev.new_name or ev.new_isin,
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
        for ev in restructure_events:            # stamp the audit trail
            if not ev.applied and id(ev) not in unstamped:
                ev.applied = today
    summary["restructure_children"] = children_added

    adjusted = 0
    for row in data.equity:
        isin = row.isin_override or isin_by_name.get(row.scrip, "")
        # chained: merger ratios fold in, and later splits on the successor
        # keep applying to the old-ISIN row (SPEC §6.15)
        f = chained_adjustment_factor(isin, row.cost_date, today,
                                      data.corporate_actions)
        row.ca_factor = f if abs(f - 1.0) > 1e-9 else None
        adjusted += row.ca_factor is not None
        cf = cost_adjustment_factor(isin, row.cost_date, today,
                                    data.corporate_actions)
        row.cost_factor = cf if abs(cf - 1.0) > 1e-9 else None
    summary["ca_adjusted_rows"] = adjusted

    # ---- dividends (SPEC §6.12): rebuild current-FY Auto rows from the feed,
    # freeze prior FYs, let Manual rows override the same (isin, type, ex-date)
    fy_now = fy_label(today)
    for d in data.dividends:                     # backfill FY typed-less rows
        if not d.fy and d.ex_date:
            d.fy = fy_label(d.ex_date)
    if div_data is not None:
        manual = [d for d in data.dividends if d.source != "Auto"]
        # keyed WITHOUT div_type (matching dedupe_dividends): the exchanges
        # word the same event differently, and a Manual correction must
        # override the Auto row however the feed typed it
        manual_keys = {(d.isin, d.ex_date) for d in manual}
        frozen_auto = [d for d in data.dividends
                       if d.source == "Auto" and d.fy != fy_now]
        # a feed failure for one symbol must not delete its current-FY rows
        kept = [d for d in data.dividends
                if d.source == "Auto" and d.fy == fy_now
                and d.isin in preserve_isins]
        kept_keys = {(d.isin, d.ex_date) for d in kept}
        fresh: list[DividendRow] = []
        for ev in div_data:
            if (not ev.ex_date or fy_label(ev.ex_date) != fy_now
                    or (ev.isin, ev.ex_date) in manual_keys
                    or (ev.isin, ev.ex_date) in kept_keys):
                continue
            for owner in sorted({r.owner for r in data.equity if r.owner}):
                qty = _dividend_qty(owner, ev.isin, ev.ex_date,
                                    data.equity, isin_by_name,
                                    data.corporate_actions)
                if qty <= 0:
                    continue
                fresh.append(DividendRow(
                    fy=fy_now, owner=owner,
                    scrip=name_by_isin.get(ev.isin, ev.scrip or ev.isin),
                    isin=ev.isin, div_type=ev.div_type, ex_date=ev.ex_date,
                    rate=ev.rate, qty=round(qty, 3), source="Auto",
                    details=ev.details))
        fresh.sort(key=lambda d: (d.ex_date or today, d.scrip, d.owner))
        rows = frozen_auto + kept + fresh + manual
        # never let Manual/current-FY rows fall past the sheet's last row:
        # on overflow the OLDEST frozen prior-FY Auto rows give way, loudly
        div_cap = M.DIV_LAST_ROW - M.FIRST_DATA_ROW + 1
        if len(rows) > div_cap:
            overflow = len(rows) - div_cap
            oldest = sorted(frozen_auto,
                            key=lambda d: d.ex_date or today)[:overflow]
            drop_ids = {id(d) for d in oldest}
            rows = [d for d in rows if id(d) not in drop_ids]
            summary["warnings"].append(
                f"Dividends sheet is full — dropped the {overflow} oldest "
                f"prior-year Auto row(s); copy old years elsewhere if you "
                f"want to keep the full record")
        data.dividends = rows
        summary["dividend_rows"] = len(fresh)
    if div_skipped:
        summary["warnings"].append(
            f"{div_skipped} dividend announcement(s) could not be parsed "
            "(e.g. percent-of-face-value wording) — add them as Manual rows "
            "on the Dividends sheet if they are yours")
    summary["dividends_fy_total"] = round(sum(
        (d.rate or 0) * (d.qty or 0)
        for d in data.dividends if d.fy == fy_now), 2)

    # ---- FMV 31-01-2018 fallback for unknown old costs (SPEC §6.6) ----
    fmv_filled = 0
    try:
        fmv_by_isin, fmv_by_symbol = load_fmv()
    except OSError:
        fmv_by_isin, fmv_by_symbol = {}, {}
    if fmv_by_isin:
        isin_by_name = {name: isin for _s, name, isin in data.masters.stock_rows}
        symbol_by_isin = {isin: sym for sym, _n, isin in data.masters.stock_rows}
        cutoff = date(2018, 2, 1)
        for row in data.equity:
            if row.avg_cost is not None or not row.qty:
                continue
            if not row.cost_date or row.cost_date >= cutoff:
                continue
            isin = row.isin_override or isin_by_name.get(row.scrip, "")
            # the 2018 ISIN may differ from today's (post-split reissues) —
            # fall back to the exchange symbol
            fmv = fmv_by_isin.get(isin) or fmv_by_symbol.get(symbol_by_isin.get(isin, ""))
            if fmv:
                row.avg_cost = fmv
                row.fmv_used = True
                fmv_filled += 1
    summary["fmv_filled"] = fmv_filled

    # ---- gold & silver (SPEC §5.7): SGBs from the bhavcopy; physical metal
    # from IBJA, falling back to the bhavcopy-implied median, else kept as-is
    if data.bullion:
        from .fetch import bullion as bullion_mod
        rates = bullion_rates
        rate_src = "injected"
        if rates is None:
            _step("Fetching today's gold & silver rate (IBJA)…")
            rates = bullion_mod.fetch_ibja()
            rate_src = "IBJA"
            if not rates:
                try:
                    rates = bullion_mod.derive_from_bhavcopy(price_data)
                    rate_src = "market-implied (SGB/ETF closes)"
                except OSError:
                    rates = {}
        sgb_priced = metal_rated = metal_rows = 0
        for r in data.bullion:
            if r.metal_type == "SGB":
                if r.isin and price_data:
                    quote = price_data.prices.get(r.isin)
                    if quote:
                        r.rate_auto = quote["close"]
                        sgb_priced += 1
            elif r.metal_type in ("Gold", "Silver"):
                metal_rows += 1
                rate = rates.get(r.metal_type.lower())
                if rate:
                    r.rate_auto = rate
                    metal_rated += 1
        if metal_rated or sgb_priced:
            data.bullion_rate_asof = today
        if metal_rows and not metal_rated:
            summary["warnings"].append(
                "gold/silver rate unavailable (IBJA and market-implied both "
                "failed) — kept the previous rates; the Rate override column "
                "always wins")
        summary["bullion"] = (f"{sgb_priced} SGB(s) priced, "
                              f"{metal_rated}/{metal_rows} metal row(s) rated"
                              + (f" ({rate_src})" if metal_rated else ""))

    # ---- NPS (SPEC §5.6): daily NAVs + scheme master, keyed by scheme code
    if data.nps and nps_data is None:
        from .fetch import nps as nps_mod
        _step("Fetching NPS NAVs (NPS Trust)…")
        try:
            nps_data = nps_mod.fetch()
        except Exception as e:  # noqa: BLE001
            summary["warnings"].append(
                f"NPS NAV fetch failed, keeping old NAVs: {e}")
    if nps_data:
        known = {code for code, _n, _p in data.masters.nps_rows}
        merged = list(data.masters.nps_rows)
        merged.extend(row for row in nps_data.master_rows
                      if row[0] not in known)
        merged.sort(key=lambda r: r[1].casefold())      # dropdown sort rule
        data.masters.nps_rows = merged
        data.masters.nps_refreshed = stamp
        code_by_scheme = {name: code
                          for code, name, _p in data.masters.nps_rows}
        matched = 0
        for r in data.nps:
            code = r.scheme_code_override or code_by_scheme.get(r.scheme, "")
            nav = nps_data.nav_by_code.get(code)
            if nav:
                r.current_nav = nav
                matched += 1
        summary["nps_matched"] = matched
        summary["nps_total"] = len(data.nps)

    # ---- EPF: auto-fill blank rates from the bundled EPFO table (SPEC §3.17)
    if data.epf:
        try:
            from .model import current_epf_rate
            epf_rate = current_epf_rate()
            for r in data.epf:
                if r.rate is None:
                    r.rate = epf_rate
        except OSError:
            pass

    # ---- PPF: exact balance/interest/XIRR for ledgered accounts (SPEC §6.10) ----
    rates = load_ppf_rates()
    led = ppf_ledger_by_account(data)
    ledgered = 0
    for row in data.ppf:
        if row.rate is None:
            row.rate = current_rate(rates)          # auto-fill blank rate
        deposits = led.get((row.owner, row.account_no))
        if deposits:
            bal, interest = ppf_value(deposits, rates, today)
            row.balance_today = round(bal, 2)
            row.interest_earned = round(interest, 2)
            row.xirr = xirr(ppf_cashflows(deposits, bal, today))
            ledgered += 1
        else:
            row.balance_today = None                 # generate → Balance today = D
            row.interest_earned = None
            row.xirr = None
    summary["ppf_ledgered"] = ledgered

    # ---- XIRR + FY-end estimate (always recomputed from current values) ----
    data.xirr = compute_all_xirr(data, today)
    summary["portfolio_xirr"] = data.xirr.portfolio
    data.fy_expected = fy_expected_by_person(data, today)
    summary["fy_expected_total"] = round(sum(data.fy_expected.values()), 2)

    # ---- net-worth history: one snapshot per day (SPEC §6.11) ----
    snap = net_worth_snapshot(data, today)
    data.history = upsert_snapshot(data.history, snap,
                                   keep=M.HISTORY_LAST_ROW - M.FIRST_DATA_ROW + 1)
    summary["net_worth"] = round(snap.total, 2)
    summary["history_points"] = len(data.history)

    # ---- regenerate, atomically ----
    tmp = path.with_name(path.stem + ".new" + path.suffix)
    build_workbook(data, str(tmp))
    os.replace(tmp, path)
    summary["workbook"] = str(path)
    return summary


def _row(icon: str, label: str, text: str) -> None:
    _say(f"  {icon} {_c('36', f'{label:<11}')} {text}")


def _print_summary(s: dict) -> None:
    _say(_c("36", "  " + "─" * 60))
    if "price_source" in s:
        nse_only = f", {s['nse_only_matched']} NSE-only" if s.get("nse_only_matched") else ""
        _row("📈", "Prices", f"{s['equity_matched']}/{s['equity_total']} stocks matched "
             f"({s['price_source']}{nse_only}) · {s['bonds_matched']} bond(s) priced")
    if "mf_matched" in s:
        _row("📊", "Fund NAVs", f"{s['mf_matched']}/{s['mf_total']} funds matched (AMFI)")
    if s.get("bullion"):
        _row("🪙", "Gold/Silver", s["bullion"])
    if "nps_matched" in s:
        _row("🏛️", "NPS NAVs", f"{s['nps_matched']}/{s['nps_total']} scheme(s) matched (NPS Trust)")
    if s.get("ca_rows") is not None:
        coverage = ("all held stocks verified" if not s.get("ca_unverified")
                    else _c("33", f"{len(s['ca_unverified'])} stock(s) UNVERIFIED"))
        _row("🔀", "Corp acts", f"{s['ca_rows']} action(s) on file · "
             f"{s['ca_adjusted_rows']} holding(s) adjusted · {coverage}")
    if s.get("restructure_children"):
        _row("🧬", "Restructure", f"{s['restructure_children']} new holding row(s) "
             f"appended from a demerger (cost & dates inherited)")
    if s.get("dividend_rows") is not None:
        _row("💸", "Dividends", f"₹{s.get('dividends_fy_total', 0):,.0f} declared this "
             f"financial year ({s['dividend_rows']} row(s) refreshed)")
    if s.get("fmv_filled"):
        _row("🕰️", "FMV filled", f"{s['fmv_filled']} old row(s) got the "
             f"31-01-2018 grandfathering cost")
    if s.get("suspended"):
        _row("⚠️", "Suspended", _c("33", f"{s['suspended']} held scrip(s) not "
                                         f"traded for 21+ days (amber)"))
    if s.get("ppf_ledgered"):
        _row("🏦", "PPF", f"{s['ppf_ledgered']} account(s) computed from the "
             f"deposit ledger (exact interest)")
    x = s.get("portfolio_xirr")
    if x is not None:
        colour = "32" if x >= 0 else "31"
        _row("🎯", "Return", f"portfolio XIRR {_c(colour, f'{x:.2%}')} a year")
    else:
        _row("🎯", "Return", "not enough dated cashflows yet")
    if s.get("fy_expected_total"):
        _row("🔮", "FY-end est", f"₹{s['fy_expected_total']:,.0f} (family total)")
    if s.get("persons_added"):
        _row("👥", "Added", f"sheet(s) for {', '.join(s['persons_added'])}")
    if s.get("classes_toggled"):
        _row("🗂️", "Classes", "; ".join(s["classes_toggled"]))
    if s.get("backup"):
        _row("💾", "Backup", s["backup"])
    for w in s["warnings"]:
        _say(_c("33", f"  ⚠️  {w}"))
    _say(_c("36", "  " + "─" * 60))
    if s.get("net_worth") is not None:
        _say("  " + _c("1;32", f"🏆 Family net worth: ₹{s['net_worth']:,.0f}")
             + f"   ({s.get('history_points', 0)} day(s) of history — "
               f"watch the Dashboard trend grow)")
    _say(f"  ✅ Updated {s['workbook']}")


RELEASES_API = "https://api.github.com/repos/jay-parikh/NetWorth/releases/latest"
RELEASES_PAGE = "https://github.com/jay-parikh/NetWorth/releases"


def _parse_version(v: str) -> tuple | None:
    """Parse 'v1.2.0' / '1.1.0rc3' → a sortable tuple; a final release sorts
    above any pre-release of the same base (…,1,0) > (…,0,N)."""
    import re
    m = re.match(r"v?(\d+)\.(\d+)\.(\d+)(?:[.\-]?(?:rc|a|b|alpha|beta|pre)\.?(\d+))?",
                 v.strip())
    if not m:
        return None
    major, minor, patch, pre = m.group(1), m.group(2), m.group(3), m.group(4)
    if pre is None:
        return (int(major), int(minor), int(patch), 1, 0)
    return (int(major), int(minor), int(patch), 0, int(pre))


def _latest_release(session=None, timeout: int = 8):
    """→ (parsed_version, tag, url) of the latest GitHub release, or None
    when unreachable / no releases. Never raises."""
    try:
        import requests
        sess = session or requests.Session()
        resp = sess.get(RELEASES_API, timeout=timeout,
                        headers={"Accept": "application/vnd.github+json"})
        if resp.status_code != 200:
            return None
        data = resp.json()
        tag = data.get("tag_name", "") or ""
        parsed = _parse_version(tag)
        if parsed is None:
            return None
        return parsed, tag, data.get("html_url") or RELEASES_PAGE
    except Exception:  # noqa: BLE001 — offline / rate-limited / private repo
        return None


def check_for_update(current: str, session=None, timeout: int = 8) -> str | None:
    """Return a one-line hint if a newer GitHub release exists, else None.
    Never raises — a nicety that must never disturb the update."""
    cur = _parse_version(current)
    if cur is None:
        return None
    latest = _latest_release(session, timeout)
    if latest and latest[0] > cur:
        return f"Update available: {latest[1]} (you have v{current}) — {latest[2]}"
    return None


def version_line(current: str, session=None, timeout: int = 8) -> str:
    """Always says SOMETHING about versions (v1.4): silence used to look
    broken when the user was simply up to date."""
    latest = _latest_release(session, timeout)
    cur = _parse_version(current)
    if latest and cur and latest[0] > cur:
        return f"Update available: {latest[1]} (you have v{current}) — {latest[2]}"
    if latest:
        return f"Version    : v{current} — you're on the latest release"
    return (f"Version    : v{current} (couldn't reach GitHub to check for a "
            f"newer one)")


def peek_persons(path: Path) -> list[str]:
    """Read just the current people from the Dashboard, cheaply (read-only)."""
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True)
    try:
        dash = wb["Dashboard"]
        return [str(dash.cell(r, 1).value).strip()
                for r in range(6, 16)
                if dash.cell(r, 1).value not in (None, "")]
    finally:
        wb.close()


def peek_class_states(path: Path) -> list[tuple[str, bool]]:
    """(label, shown?) per asset class from the Settings sheet, cheaply
    (read-only). Missing sheet (pre-v1.3 workbook) → registry defaults."""
    from openpyxl import load_workbook
    states = {c.label: c.default_enabled for c in ASSET_CLASSES}
    wb = load_workbook(path, read_only=True)
    try:
        if "Settings" in wb.sheetnames:
            st = wb["Settings"]
            for row in st.iter_rows(min_row=4, max_row=20, max_col=2):
                label = str(row[0].value or "").strip()
                if label in states and row[1].value is not None:
                    states[label] = str(row[1].value).strip().casefold() != "no"
    finally:
        wb.close()
    return [(c.label, states[c.label]) for c in ASSET_CLASSES]


def prompt_class_toggle(path: Path) -> list[str]:
    """Interactively show/hide asset classes — easier than editing the
    Settings sheet by hand. Returns the labels to flip."""
    states = peek_class_states(path)
    print("\nYour asset classes — the workbook shows only what's on:")
    for i, (label, shown) in enumerate(states, 1):
        print(f"  {i:2}. {label:14} [{'shown' if shown else 'hidden'}]")
    print("Show or hide something? Type its number(s), e.g. 7 or 7 9 — "
          "or just press Enter to skip.")
    try:
        raw = input("  Toggle (blank to continue): ").strip()
    except EOFError:
        return []
    toggles: list[str] = []
    for tok in raw.replace(",", " ").split():
        if tok.isdigit() and 1 <= int(tok) <= len(states):
            label, shown = states[int(tok) - 1]
            toggles.append(label)
            if shown:
                print(f"  ✓ {label} will be hidden (if it still holds rows, "
                      f"it stays visible until you delete them)")
            else:
                print(f"  ✓ {label} will be shown")
    return toggles


def prompt_new_persons(existing: list[str]) -> list[str]:
    """Interactively collect new person names (double-click / Terminal only).
    Adding a person here is optional — you can also just type a name in a yellow
    Dashboard cell."""
    print("\nPeople currently tracked: " + (", ".join(existing) or "(none)"))
    slots = 10 - len(existing)
    if slots <= 0:
        print("(the Dashboard already lists the maximum of 10 people)")
        return []
    print("Add a new person's sheet? Type a name and press Enter, "
          "or just press Enter to skip.")
    added: list[str] = []
    seen = {p.casefold() for p in existing}
    while len(added) < slots:
        try:
            name = input("  New person (blank to continue): ").strip()
        except EOFError:
            break
        if not name:
            break
        if name.casefold() in seen:
            print(f"  '{name}' is already tracked.")
            continue
        seen.add(name.casefold())
        added.append(name)
        print(f"  ✓ will add a sheet for {name} "
              f"(then record holdings with '{name}' in the Owner columns)")
    return added


def _use_os_trust_store() -> None:
    """Validate TLS against the OS certificate store (like browsers and the
    legacy PowerShell did) so corporate/AV proxies don't break the fetch."""
    try:
        import truststore
        truststore.inject_into_ssl()
    except Exception:  # noqa: BLE001 — certifi fallback is fine
        pass


def main(argv: list[str] | None = None) -> int:
    _use_os_trust_store()
    parser = argparse.ArgumentParser(
        description="Refresh prices, NAVs and XIRR in a tracker workbook.")
    parser.add_argument("workbook", nargs="?", help="path to the .xlsx (default: "
                        f"{M.TEMPLATE_FILENAME} next to the current directory)")
    parser.add_argument("--pause", action="store_true",
                        help="wait for Enter before exiting (double-click launchers)")
    parser.add_argument("--no-update-check", action="store_true",
                        help="don't check GitHub for a newer version")
    parser.add_argument("--no-prompt", action="store_true",
                        help="never ask interactive questions (e.g. add a person)")
    parser.add_argument("--add-person", action="append", metavar="NAME", default=[],
                        help="add a person's sheet without prompting (repeatable)")
    args = parser.parse_args(argv)

    code = 0
    _banner(f"NetWorth v{__version__} — updating your family portfolio")
    try:
        path = locate_workbook(args.workbook)
        new_persons = list(args.add_person)
        # interactive only when attached to a real console — never hangs a
        # headless/scheduled run
        interactive = sys.stdin.isatty() and not args.no_prompt
        toggles: list[str] = []
        if interactive:
            try:
                new_persons += prompt_new_persons(peek_persons(path))
                toggles = prompt_class_toggle(path)
            except Exception:  # noqa: BLE001 — prompting must never break the run
                pass
        summary = run(path, add_persons=new_persons, toggle_classes=toggles)
        _print_summary(summary)
    except SystemExit as e:
        code = e.code if isinstance(e.code, int) else 1
    except Exception as e:  # noqa: BLE001 — last-resort user-facing message
        print(f"ERROR: {e}", file=sys.stderr)
        code = 1

    if not (args.no_update_check or os.environ.get("NETWORTH_NO_UPDATE_CHECK")):
        line = version_line(__version__)
        _say("  " + (_c("1;33", line) if "Update available" in line
                     else _c("2", line)))

    if args.pause:
        input("\nPress Enter to close...")
    return code


if __name__ == "__main__":
    raise SystemExit(main())
