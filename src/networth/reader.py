"""Round-trip reader — reconstructs PortfolioData from an existing workbook.

openpyxl is used strictly READ-ONLY (saving through it destroys charts and
comments). Formula cells come back as "=..." strings, which is exactly how we
tell a generator-written lookup apart from a user's manual override (SPEC §7:
only structured inputs and updater-written values round-trip).

Robustness rules: the header row is located by its known header names within
rows 1–5, and data rows are scanned to the sheet's end, so users inserting,
deleting or sorting rows never breaks the read.
"""

from __future__ import annotations

from datetime import date, datetime

from openpyxl import load_workbook

from .model import (
    BondRow, ClassXirr, EquityRow, FDRow, Masters, MFRow, PPFRow,
    PortfolioData, ScripRef, SIPRow,
)


def _as_date(v) -> date | None:
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    return None


def _as_float(v) -> float | None:
    return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else None


def _as_str(v) -> str:
    return v.strip() if isinstance(v, str) else ("" if v is None else str(v))


def _is_formula(v) -> bool:
    return isinstance(v, str) and v.startswith("=")


def _manual(v) -> str:
    """A user override in a normally-formula column: plain text, not a formula."""
    return "" if _is_formula(v) or v is None else _as_str(v)


def _header_row(ws, first_header: str) -> int:
    for r in range(1, 6):
        if _as_str(ws.cell(r, 1).value) == first_header:
            return r
    raise ValueError(f"{ws.title}: header row with '{first_header}' not found in rows 1-5")


def _data_rows(ws, header: int):
    for r in range(header + 1, ws.max_row + 1):
        # TOTAL rows carry no Owner/ISIN in column A and 'TOTAL' in column C
        if _as_str(ws.cell(r, 3).value) == "TOTAL":
            continue
        yield r


def read_workbook(path: str) -> PortfolioData:
    wb = load_workbook(path, read_only=False, data_only=False)
    data = PortfolioData()

    dash = wb["Dashboard"]
    data.persons = [
        _as_str(dash.cell(r, 1).value)
        for r in range(6, 16) if _as_str(dash.cell(r, 1).value)
    ]
    infl = _as_float(dash["E3"].value)
    data.inflation_pct = infl if infl is not None else 7
    exp = _as_float(dash["E2"].value)
    data.expected_return_pct = exp if exp is not None else 10
    for r in range(6, 16):
        person = _as_str(dash.cell(r, 1).value)
        fy = _as_float(dash.cell(r, 8).value)
        if person and fy is not None:
            data.fy_expected[person] = fy
    data.xirr = ClassXirr(
        portfolio=_as_float(dash["B4"].value),
        equity=_as_float(dash["C20"].value),
        mutual_funds=_as_float(dash["C21"].value),
        fixed_deposits=_as_float(dash["C22"].value),
        ppf=_as_float(dash["C23"].value),
        bonds=_as_float(dash["C24"].value),
    )

    ws = wb["Equity"]
    h = _header_row(ws, "Owner")
    for r in _data_rows(ws, h):
        owner = _as_str(ws.cell(r, 1).value)
        scrip = _manual(ws.cell(r, 3).value)
        if not owner and not scrip:
            continue
        data.equity.append(EquityRow(
            owner=owner, scrip=scrip,
            qty=_as_float(ws.cell(r, 4).value),
            avg_cost=_as_float(ws.cell(r, 5).value),
            close=_as_float(ws.cell(r, 6).value),
            prev_close=_as_float(ws.cell(r, 7).value),
            close_date=_as_str(ws.cell(r, 8).value),
            cost_date=_as_date(ws.cell(r, 13).value),
            isin_override=_manual(ws.cell(r, 2).value),
        ))

    ws = wb["MutualFunds"]
    h = _header_row(ws, "Owner")
    for r in _data_rows(ws, h):
        owner = _as_str(ws.cell(r, 1).value)
        scheme = _manual(ws.cell(r, 3).value)
        if not owner and not scheme:
            continue
        data.mutual_funds.append(MFRow(
            owner=owner, scheme=scheme,
            current_nav=_as_float(ws.cell(r, 7).value),
            xirr=_as_float(ws.cell(r, 12).value),
            fund_house_override=_manual(ws.cell(r, 2).value),
            isin_override=_manual(ws.cell(r, 4).value),
        ))

    ws = wb["MF_SIP"]
    h = _header_row(ws, "Owner")
    for r in _data_rows(ws, h):
        owner = _as_str(ws.cell(r, 1).value)
        scheme = _manual(ws.cell(r, 3).value)
        if not owner and not scheme:
            continue
        units = ws.cell(r, 8).value
        data.sip.append(SIPRow(
            owner=owner, scheme=scheme,
            txn_date=_as_date(ws.cell(r, 5).value),
            amount=_as_float(ws.cell(r, 6).value),
            nav=_as_float(ws.cell(r, 7).value),
            units_override=None if _is_formula(units) else _as_float(units),
            fund_house_override=_manual(ws.cell(r, 2).value),
            isin_override=_manual(ws.cell(r, 4).value),
        ))

    ws = wb["FixedDeposits"]
    h = _header_row(ws, "Owner")
    for r in _data_rows(ws, h):
        owner = _as_str(ws.cell(r, 1).value)
        if not owner:
            continue
        data.fixed_deposits.append(FDRow(
            owner=owner,
            bank=_as_str(ws.cell(r, 2).value),
            fd_no=_as_str(ws.cell(r, 3).value),
            principal=_as_float(ws.cell(r, 4).value),
            rate=_as_float(ws.cell(r, 5).value),
            start=_as_date(ws.cell(r, 6).value),
            maturity=_as_date(ws.cell(r, 7).value),
            comp_per_year=(int(v) if (v := _as_float(ws.cell(r, 8).value)) else None),
        ))

    ws = wb["PPF"]
    h = _header_row(ws, "Owner")
    for r in _data_rows(ws, h):
        owner = _as_str(ws.cell(r, 1).value)
        if not owner:
            continue
        data.ppf.append(PPFRow(
            owner=owner,
            institution=_as_str(ws.cell(r, 2).value),
            account_no=_as_str(ws.cell(r, 3).value),
            balance=_as_float(ws.cell(r, 4).value),
            as_on=_as_date(ws.cell(r, 5).value),
            rate=_as_float(ws.cell(r, 6).value),
            notes=_as_str(ws.cell(r, 7).value),
        ))

    ws = wb["Bonds"]
    h = _header_row(ws, "Owner")
    for r in _data_rows(ws, h):
        owner = _as_str(ws.cell(r, 1).value)
        if not owner:
            continue
        data.bonds.append(BondRow(
            owner=owner,
            issuer=_as_str(ws.cell(r, 2).value),
            isin=_as_str(ws.cell(r, 3).value),
            qty=_as_float(ws.cell(r, 4).value),
            face=_as_float(ws.cell(r, 5).value),
            buy_price=_as_float(ws.cell(r, 6).value),
            cur_price=_as_float(ws.cell(r, 7).value),
            coupon=_as_float(ws.cell(r, 8).value),
            maturity=_as_date(ws.cell(r, 9).value),
            buy_date=_as_date(ws.cell(r, 13).value),
        ))

    ws = wb["By Scrip"]
    h = _header_row(ws, "ISIN")
    for r in _data_rows(ws, h):
        isin = _as_str(ws.cell(r, 1).value)
        if not isin:
            continue
        data.by_scrip.append(ScripRef(isin=isin, name=_as_str(ws.cell(r, 2).value)))

    def master_rows(sheet: str) -> list[tuple[str, str, str]]:
        m = wb[sheet]
        out = []
        for r in range(4, m.max_row + 1):
            isin = _as_str(m.cell(r, 3).value)
            if isin:
                out.append((_as_str(m.cell(r, 1).value), _as_str(m.cell(r, 2).value), isin))
        return out

    data.masters = Masters(
        mf_rows=master_rows("MF_Master"),
        stock_rows=master_rows("Stock_Master"),
        mf_refreshed=_as_str(wb["MF_Master"]["E2"].value),
        stock_refreshed=_as_str(wb["Stock_Master"]["E2"].value),
    )
    wb.close()
    return data
