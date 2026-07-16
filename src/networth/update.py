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
from .model import (ASSET_CLASSES, DividendRow, adjustment_factor,
                    class_has_data, fy_label, load_fmv)
from .generate import build_workbook
from .reader import read_workbook

KEEP_BACKUPS = 10


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
    the ex-date, at the CA-adjusted count as of the day before. There is no
    sell ledger — the estimate projects the CURRENT rows backwards, which the
    sheet hint and Guide state plainly (hence the amber '(est.)' columns)."""
    from datetime import timedelta
    as_of = ex - timedelta(days=1)
    total = 0.0
    for row in equity_rows:
        row_isin = row.isin_override or isin_by_name.get(row.scrip, "")
        if row_isin != isin or row.owner != owner or not row.qty:
            continue
        if row.cost_date and row.cost_date >= ex:
            continue
        total += row.qty * adjustment_factor(isin, row.cost_date, as_of, actions)
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
        div_data=None, add_persons: list[str] | None = None,
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

    # a class switched off on Settings but still holding rows stays visible
    # (never lose data, never let a hidden value haunt the totals — SPEC §3.14)
    for cls in ASSET_CLASSES:
        s = data.class_settings.get(cls.key)
        if s and not s.enabled and class_has_data(data, cls.key):
            summary["warnings"].append(
                f"{cls.label} is set to No on Settings but still holds rows — "
                f"kept visible; delete or move its rows to hide it")

    if do_backup:
        summary["backup"] = str(make_backup(path))

    # ---- fetch (graceful per source) ----
    if price_data is None:
        try:
            price_data = bhav_mod.fetch(today=today)
        except Exception as e:  # noqa: BLE001 — any fetch failure degrades
            summary["warnings"].append(f"price fetch failed, keeping old prices: {e}")
    if amfi_data is None:
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
        for row in data.equity:
            isin = row.isin_override or isin_by_name.get(row.scrip, "")
            if not isin:
                continue
            quote = price_data.prices.get(isin)
            if quote:
                row.close = quote["close"]
                if quote["prev"]:
                    row.prev_close = quote["prev"]
                row.close_date = trade_date
                data.masters.stock_status[isin] = ("Active", trade_date)
                matched += 1
                nse_only_hits += isin in nse_only
            elif dual_source:
                # absent from both exchanges: escalate by how long it has been
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
    div_skipped = 0
    if ca_data is None:
        symbols = {symbol_by_isin[i]: i for i in held_isins if symbol_by_isin.get(i)}
        bse_codes = {code: isin
                     for isin, code in getattr(price_data, "codes_by_isin", {}).items()
                     if isin in held_isins} if price_data else {}
        try:
            checked: set[str] = set()
            if symbols or bse_codes:
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
            unchecked = held_isins - checked
            summary["ca_unverified"] = sorted(
                name_by_isin.get(i, i) for i in unchecked)
            if unchecked:
                summary["warnings"].append(
                    "corporate actions could NOT be verified for: "
                    + ", ".join(summary["ca_unverified"])
                    + " — quantities for these may miss splits/bonuses; add any "
                      "you know of as Manual rows on the Corporate_Actions sheet")
        except Exception as e:  # noqa: BLE001
            summary["warnings"].append(
                f"corporate-actions fetch failed, keeping existing rows: {e}")
    else:
        summary["ca_unverified"] = []
    if ca_data is not None:
        manual = [a for a in data.corporate_actions if a.source != "Auto"]
        manual_keys = {(a.isin, a.type, a.ex_date) for a in manual}
        auto = [a for a in ca_data
                if (a.isin, a.type, a.ex_date) not in manual_keys]
        auto.sort(key=lambda a: (a.symbol, a.ex_date or today))
        data.corporate_actions = auto + manual
        summary["ca_rows"] = len(data.corporate_actions)

    adjusted = 0
    for row in data.equity:
        isin = row.isin_override or isin_by_name.get(row.scrip, "")
        f = adjustment_factor(isin, row.cost_date, today, data.corporate_actions)
        row.ca_factor = f if abs(f - 1.0) > 1e-9 else None
        adjusted += row.ca_factor is not None
    summary["ca_adjusted_rows"] = adjusted

    # ---- dividends (SPEC §6.12): rebuild current-FY Auto rows from the feed,
    # freeze prior FYs, let Manual rows override the same (isin, type, ex-date)
    fy_now = fy_label(today)
    for d in data.dividends:                     # backfill FY typed-less rows
        if not d.fy and d.ex_date:
            d.fy = fy_label(d.ex_date)
    if div_data is not None:
        manual = [d for d in data.dividends if d.source != "Auto"]
        manual_keys = {(d.isin, d.div_type, d.ex_date) for d in manual}
        frozen_auto = [d for d in data.dividends
                       if d.source == "Auto" and d.fy != fy_now]
        fresh: list[DividendRow] = []
        for ev in div_data:
            if (not ev.ex_date or fy_label(ev.ex_date) != fy_now
                    or (ev.isin, ev.div_type, ev.ex_date) in manual_keys):
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
        data.dividends = frozen_auto + fresh + manual
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


def _print_summary(s: dict) -> None:
    if "price_source" in s:
        nse_only = f", {s['nse_only_matched']} NSE-only" if s.get("nse_only_matched") else ""
        print(f"Prices     : {s['equity_matched']}/{s['equity_total']} scrips matched "
              f"({s['price_source']}{nse_only}), {s['stocks_added']} new listings, "
              f"{s['bonds_matched']} bond(s) priced")
    if "mf_matched" in s:
        print(f"Fund NAVs  : {s['mf_matched']}/{s['mf_total']} funds matched (AMFI)")
    if s.get("ca_rows") is not None:
        coverage = ("all held stocks verified" if not s.get("ca_unverified")
                    else f"{len(s['ca_unverified'])} stock(s) UNVERIFIED")
        print(f"Corp acts  : {s['ca_rows']} action(s) on file (NSE+BSE), "
              f"{s['ca_adjusted_rows']} holding(s) adjusted, {coverage}")
    if s.get("dividend_rows") is not None:
        print(f"Dividends  : {s['dividend_rows']} row(s) refreshed for this FY, "
              f"{s.get('dividends_fy_total', 0):,.0f} declared this financial year")
    if s.get("fmv_filled"):
        print(f"FMV filled : {s['fmv_filled']} row(s) got the 31-01-2018 grandfathering cost")
    if s.get("suspended"):
        print(f"Suspended  : {s['suspended']} held scrip(s) not traded for 21+ days (amber)")
    if s.get("ppf_ledgered"):
        print(f"PPF        : {s['ppf_ledgered']} account(s) computed from the deposit "
              f"ledger (exact interest)")
    x = s.get("portfolio_xirr")
    print(f"XIRR       : portfolio {x:.2%}" if x is not None else
          "XIRR       : not enough dated cashflows yet")
    if s.get("fy_expected_total"):
        print(f"FY-end est : {s['fy_expected_total']:,.0f} (family total)")
    if s.get("persons_added"):
        print(f"Added      : sheet(s) for {', '.join(s['persons_added'])}")
    if s.get("net_worth") is not None:
        print(f"Net worth  : {s['net_worth']:,.0f}  "
              f"({s.get('history_points', 0)} day(s) of history)")
    if s.get("backup"):
        print(f"Backup     : {s['backup']}")
    for w in s["warnings"]:
        print(f"WARNING    : {w}")
    print(f"Updated    : {s['workbook']}")


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


def check_for_update(current: str, session=None, timeout: int = 8) -> str | None:
    """Return a one-line hint if a newer GitHub release exists, else None.
    Never raises — a nicety that must never disturb the update."""
    cur = _parse_version(current)
    if cur is None:
        return None
    try:
        import requests
        sess = session or requests.Session()
        resp = sess.get(RELEASES_API, timeout=timeout,
                        headers={"Accept": "application/vnd.github+json"})
        if resp.status_code != 200:
            return None
        data = resp.json()
        tag = data.get("tag_name", "") or ""
        url = data.get("html_url") or RELEASES_PAGE
    except Exception:  # noqa: BLE001 — offline / rate-limited / private repo
        return None
    latest = _parse_version(tag)
    if latest and latest > cur:
        return f"Update available: {tag} (you have v{current}) — {url}"
    return None


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
    try:
        path = locate_workbook(args.workbook)
        new_persons = list(args.add_person)
        # interactive only when attached to a real console — never hangs a
        # headless/scheduled run
        interactive = sys.stdin.isatty() and not args.no_prompt
        if interactive:
            try:
                new_persons += prompt_new_persons(peek_persons(path))
            except Exception:  # noqa: BLE001 — prompting must never break the run
                pass
        summary = run(path, add_persons=new_persons)
        _print_summary(summary)
    except SystemExit as e:
        code = e.code if isinstance(e.code, int) else 1
    except Exception as e:  # noqa: BLE001 — last-resort user-facing message
        print(f"ERROR: {e}", file=sys.stderr)
        code = 1

    if not (args.no_update_check or os.environ.get("NETWORTH_NO_UPDATE_CHECK")):
        hint = check_for_update(__version__)
        if hint:
            print(hint)

    if args.pause:
        input("\nPress Enter to close...")
    return code


if __name__ == "__main__":
    raise SystemExit(main())
