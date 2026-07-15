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

from . import model as M
from .compute.cashflows import compute_all_xirr
from .compute.projections import fy_expected_by_person
from .fetch import amfi as amfi_mod
from .fetch import bhavcopy as bhav_mod
from .fetch import corporate_actions as ca_mod
from .model import adjustment_factor, load_fmv
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
        today: date | None = None, do_backup: bool = True) -> dict:
    today = today or date.today()
    summary: dict = {"warnings": []}

    ensure_closed(path)
    data = read_workbook(str(path))
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
        matched = 0
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
            else:
                # not traded today: escalate by how long it has been absent
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

    # ---- corporate actions (SPEC §6.7): fetch auto, keep manual, factor rows ----
    isin_by_name = {name: isin for _s, name, isin in data.masters.stock_rows}
    symbol_by_isin = {isin: sym for sym, _n, isin in data.masters.stock_rows}
    held_isins = {row.isin_override or isin_by_name.get(row.scrip, "")
                  for row in data.equity}
    held_isins.discard("")
    if ca_data is None:
        symbols = {symbol_by_isin[i]: i for i in held_isins if symbol_by_isin.get(i)}
        try:
            ca_data = ca_mod.fetch(symbols) if symbols else []
        except Exception as e:  # noqa: BLE001
            summary["warnings"].append(
                f"corporate-actions fetch failed, keeping existing rows: {e}")
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

    # ---- XIRR + FY-end estimate (always recomputed from current values) ----
    data.xirr = compute_all_xirr(data, today)
    summary["portfolio_xirr"] = data.xirr.portfolio
    data.fy_expected = fy_expected_by_person(data, today)
    summary["fy_expected_total"] = round(sum(data.fy_expected.values()), 2)

    # ---- regenerate, atomically ----
    tmp = path.with_name(path.stem + ".new" + path.suffix)
    build_workbook(data, str(tmp))
    os.replace(tmp, path)
    summary["workbook"] = str(path)
    return summary


def _print_summary(s: dict) -> None:
    if "price_source" in s:
        print(f"Prices     : {s['equity_matched']}/{s['equity_total']} scrips matched "
              f"({s['price_source']}), {s['stocks_added']} new listings, "
              f"{s['bonds_matched']} bond(s) priced")
    if "mf_matched" in s:
        print(f"Fund NAVs  : {s['mf_matched']}/{s['mf_total']} funds matched (AMFI)")
    if s.get("ca_rows") is not None:
        print(f"Corp acts  : {s['ca_rows']} action(s) on file, "
              f"{s['ca_adjusted_rows']} holding(s) adjusted (splits/bonuses)")
    if s.get("fmv_filled"):
        print(f"FMV filled : {s['fmv_filled']} row(s) got the 31-01-2018 grandfathering cost")
    if s.get("suspended"):
        print(f"Suspended  : {s['suspended']} held scrip(s) not traded for 21+ days (amber)")
    x = s.get("portfolio_xirr")
    print(f"XIRR       : portfolio {x:.2%}" if x is not None else
          "XIRR       : not enough dated cashflows yet")
    if s.get("fy_expected_total"):
        print(f"FY-end est : {s['fy_expected_total']:,.0f} (family total)")
    if s.get("backup"):
        print(f"Backup     : {s['backup']}")
    for w in s["warnings"]:
        print(f"WARNING    : {w}")
    print(f"Updated    : {s['workbook']}")


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
    args = parser.parse_args(argv)

    code = 0
    try:
        summary = run(locate_workbook(args.workbook))
        _print_summary(summary)
    except SystemExit as e:
        code = e.code if isinstance(e.code, int) else 1
    except Exception as e:  # noqa: BLE001 — last-resort user-facing message
        print(f"ERROR: {e}", file=sys.stderr)
        code = 1
    if args.pause:
        input("\nPress Enter to close...")
    return code


if __name__ == "__main__":
    raise SystemExit(main())
