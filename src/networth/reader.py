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
    ASSET_CLASSES, MANUAL_CLASS_LABELS, BondRow, BullionRow, ClassSetting,
    ClassXirr, CorporateAction, DividendRow, EPFRow, EquityRow, FDRow,
    HistorySnapshot, ManualAssetRow, Masters, MFRow, NPSRow, PPFLedgerRow,
    PPFRow, PortfolioData, ScripRef, SIPRow,
)

# the Class dropdown is non-blocking, so users can type any casing — Excel's
# SUMIFS matches case-insensitively and the Python side must agree, so typed
# variants are canonicalised on read ("real estate" → "Real Estate")
_CANON_CLASS = {label.casefold(): label for label in MANUAL_CLASS_LABELS}


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
    # the FY-expected column moves with the number of enabled classes —
    # locate it by its "Expected @ ..." header in row 5
    fy_col = None
    for c in range(2, 30):
        if _as_str(dash.cell(5, c).value).startswith("Expected @"):
            fy_col = c
            break
    if fy_col:
        for r in range(6, 16):
            person = _as_str(dash.cell(r, 1).value)
            fy = _as_float(dash.cell(r, fy_col).value)
            if person and fy is not None:
                data.fy_expected[person] = fy

    # class XIRRs live in the allocation table — locate its header row and
    # map rows back by class label (layout is dynamic since R10)
    data.xirr = ClassXirr(portfolio=_as_float(dash["B4"].value))
    key_by_label = {c.label: c.key for c in ASSET_CLASSES}
    for r in range(15, 40):
        if _as_str(dash.cell(r, 1).value) == "Asset class":
            for rr in range(r + 1, r + 1 + len(ASSET_CLASSES) + 5):
                key = key_by_label.get(_as_str(dash.cell(rr, 1).value))
                if key is None:
                    break
                setattr(data.xirr, key, _as_float(dash.cell(rr, 3).value))
            break

    if "Settings" in wb.sheetnames:
        st = wb["Settings"]
        label_rows = {_as_str(st.cell(r, 1).value): r for r in range(4, 21)}
        for cls in ASSET_CLASSES:
            r = label_rows.get(cls.label)
            if r is None:
                continue
            enabled_txt = _as_str(st.cell(r, 2).value).casefold()
            data.class_settings[cls.key] = ClassSetting(
                enabled=(enabled_txt != "no") if enabled_txt
                else cls.default_enabled,
                target_pct=_as_float(st.cell(r, 3).value),
            )
        tol_row = next((r for r in range(16, 25)
                        if _as_str(st.cell(r, 1).value).startswith("Drift tolerance")),
                       None)
        if tol_row:
            tol = _as_float(st.cell(tol_row, 2).value)
            if tol is not None:
                data.drift_tolerance_pct = tol
    # no Settings sheet (pre-v1.3 workbook) → registry defaults already set

    ws = wb["Equity"]
    h = _header_row(ws, "Owner")
    for r in _data_rows(ws, h):
        owner = _as_str(ws.cell(r, 1).value)
        scrip = _manual(ws.cell(r, 3).value)
        if not owner and not scrip:
            continue
        # R = Flags: "FMV", a restructure flag, or both joined with " | "
        flag_parts = [p.strip()
                      for p in _as_str(ws.cell(r, 18).value).split("|")]
        data.equity.append(EquityRow(
            owner=owner, scrip=scrip,
            qty=_as_float(ws.cell(r, 4).value),
            avg_cost=_as_float(ws.cell(r, 5).value),
            close=_as_float(ws.cell(r, 6).value),
            prev_close=_as_float(ws.cell(r, 7).value),
            close_date=_as_date(ws.cell(r, 8).value),
            cost_date=_as_date(ws.cell(r, 13).value),
            isin_override=_manual(ws.cell(r, 2).value),
            fmv_used="FMV" in flag_parts,
            flag=" | ".join(p for p in flag_parts if p and p != "FMV"),
            ca_factor=_as_float(ws.cell(r, 19).value),            # S = Adj factor
            cost_factor=_as_float(ws.cell(r, 20).value),          # T = Cost factor
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
            balance_today=_as_float(ws.cell(r, 8).value),    # H (formula → None)
            interest_earned=_as_float(ws.cell(r, 9).value),  # I
            xirr=_as_float(ws.cell(r, 10).value),            # J
        ))

    if "PPF_Ledger" in wb.sheetnames:
        ws = wb["PPF_Ledger"]
        h = _header_row(ws, "Owner")
        for r in range(h + 1, ws.max_row + 1):
            owner = _as_str(ws.cell(r, 1).value)
            acct = _as_str(ws.cell(r, 2).value)
            if not owner and not acct:
                continue
            data.ppf_ledger.append(PPFLedgerRow(
                owner=owner, account_no=acct,
                txn_date=_as_date(ws.cell(r, 3).value),
                amount=_as_float(ws.cell(r, 4).value),
            ))

    if "EPF" in wb.sheetnames:
        ws = wb["EPF"]
        h = _header_row(ws, "Owner")
        for r in _data_rows(ws, h):
            owner = _as_str(ws.cell(r, 1).value)
            if not owner:
                continue
            data.epf.append(EPFRow(
                owner=owner,
                establishment=_as_str(ws.cell(r, 2).value),
                member_id=_as_str(ws.cell(r, 3).value),
                balance=_as_float(ws.cell(r, 4).value),
                as_on=_as_date(ws.cell(r, 5).value),
                rate=_as_float(ws.cell(r, 6).value),
                notes=_as_str(ws.cell(r, 7).value),
            ))

    if "Gold_Silver" in wb.sheetnames:
        ws = wb["Gold_Silver"]
        h = _header_row(ws, "Owner")
        data.bullion_rate_asof = _as_date(ws["I2"].value)
        for r in _data_rows(ws, h):
            owner = _as_str(ws.cell(r, 1).value)
            if not owner:
                continue
            data.bullion.append(BullionRow(
                owner=owner,
                metal_type=_as_str(ws.cell(r, 2).value),
                description=_as_str(ws.cell(r, 3).value),
                isin=_as_str(ws.cell(r, 4).value),
                qty=_as_float(ws.cell(r, 5).value),
                purity=_as_float(ws.cell(r, 6).value),
                buy_price=_as_float(ws.cell(r, 7).value),
                buy_date=_as_date(ws.cell(r, 8).value),
                rate_auto=_as_float(ws.cell(r, 9).value),
                rate_override=_as_float(ws.cell(r, 10).value),
                maturity=_as_date(ws.cell(r, 14).value),
            ))

    if "NPS" in wb.sheetnames:
        ws = wb["NPS"]
        h = _header_row(ws, "Owner")
        for r in _data_rows(ws, h):
            owner = _as_str(ws.cell(r, 1).value)
            scheme = _manual(ws.cell(r, 3).value)
            if not owner and not scheme:
                continue
            data.nps.append(NPSRow(
                owner=owner,
                pran=_as_str(ws.cell(r, 2).value),
                scheme=scheme,
                units=_as_float(ws.cell(r, 5).value),
                current_nav=_as_float(ws.cell(r, 6).value),
                total_contributed=_as_float(ws.cell(r, 8).value),
                first_contribution=_as_date(ws.cell(r, 9).value),
                xirr=_as_float(ws.cell(r, 10).value),
                scheme_code_override=_manual(ws.cell(r, 4).value),
            ))

    if "Manual_Assets" in wb.sheetnames:
        ws = wb["Manual_Assets"]
        h = _header_row(ws, "Owner")
        for r in _data_rows(ws, h):
            owner = _as_str(ws.cell(r, 1).value)
            if not owner:
                continue
            raw_class = _as_str(ws.cell(r, 2).value)
            data.manual_assets.append(ManualAssetRow(
                owner=owner,
                asset_class=_CANON_CLASS.get(raw_class.casefold(), raw_class),
                description=_as_str(ws.cell(r, 3).value),
                institution=_as_str(ws.cell(r, 4).value),
                invested=_as_float(ws.cell(r, 5).value),
                cost_date=_as_date(ws.cell(r, 6).value),
                value=_as_float(ws.cell(r, 7).value),
                as_on=_as_date(ws.cell(r, 8).value),
                notes=_as_str(ws.cell(r, 10).value),
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

    if "Corporate_Actions" in wb.sheetnames:
        ws = wb["Corporate_Actions"]
        h = _header_row(ws, "Symbol")
        for r in _data_rows(ws, h):
            symbol = _as_str(ws.cell(r, 1).value)
            isin = _as_str(ws.cell(r, 2).value)
            if not symbol and not isin:
                continue
            data.corporate_actions.append(CorporateAction(
                symbol=symbol, isin=isin,
                type=_as_str(ws.cell(r, 3).value).upper(),
                ex_date=_as_date(ws.cell(r, 4).value),
                ratio_from=_as_float(ws.cell(r, 5).value),
                ratio_to=_as_float(ws.cell(r, 6).value),
                source=_as_str(ws.cell(r, 8).value) or "Manual",
                details=_as_str(ws.cell(r, 9).value),
                new_isin=_as_str(ws.cell(r, 10).value),
                cost_pct=_as_float(ws.cell(r, 11).value),
                applied=_as_date(ws.cell(r, 12).value),
            ))

    if "Dividends" in wb.sheetnames:
        ws = wb["Dividends"]
        h = _header_row(ws, "FY")
        for r in _data_rows(ws, h):
            scrip = _as_str(ws.cell(r, 3).value)
            isin = _as_str(ws.cell(r, 4).value)
            if not scrip and not isin:
                continue                     # skips the by-month block too
            data.dividends.append(DividendRow(
                fy=_as_str(ws.cell(r, 1).value),
                owner=_as_str(ws.cell(r, 2).value),
                scrip=scrip, isin=isin,
                div_type=_as_str(ws.cell(r, 5).value).capitalize(),
                ex_date=_as_date(ws.cell(r, 6).value),
                rate=_as_float(ws.cell(r, 7).value),
                qty=_as_float(ws.cell(r, 8).value),
                source=_as_str(ws.cell(r, 10).value) or "Manual",
                details=_as_str(ws.cell(r, 11).value),
            ))

    def master_rows(sheet: str) -> list[tuple[str, str, str]]:
        m = wb[sheet]
        out = []
        for r in range(4, m.max_row + 1):
            isin = _as_str(m.cell(r, 3).value)
            if isin:
                out.append((_as_str(m.cell(r, 1).value), _as_str(m.cell(r, 2).value), isin))
        return out

    if "History" in wb.sheetnames:
        ws = wb["History"]
        h = _header_row(ws, "Date")
        # columns are label-keyed (SPEC §6.11): a class column may be absent
        # (never enabled) or in any position; unknown labels are ignored
        key_by_label = {c.label: c.key for c in ASSET_CLASSES}
        col_key: dict[int, str] = {}
        for c in range(2, 40):
            key = key_by_label.get(_as_str(ws.cell(h, c).value))
            if key:
                col_key[c] = key
        for r in range(h + 1, ws.max_row + 1):
            d = _as_date(ws.cell(r, 1).value)
            if d is None:
                continue
            snap = HistorySnapshot(snap_date=d)
            for c, key in col_key.items():
                setattr(snap, key, _as_float(ws.cell(r, c).value) or 0.0)
            data.history.append(snap)

    sm = wb["Stock_Master"]
    stock_status: dict[str, tuple[str, object]] = {}
    for r in range(4, sm.max_row + 1):
        isin = _as_str(sm.cell(r, 3).value)
        st = _as_str(sm.cell(r, 4).value)
        if isin and st:
            stock_status[isin] = (st, _as_date(sm.cell(r, 5).value))

    data.masters = Masters(
        mf_rows=master_rows("MF_Master"),
        stock_rows=master_rows("Stock_Master"),
        nps_rows=(master_rows("NPS_Master")
                  if "NPS_Master" in wb.sheetnames else []),
        mf_refreshed=_as_str(wb["MF_Master"]["E2"].value),
        stock_refreshed=_as_str(wb["Stock_Master"]["E2"].value),
        nps_refreshed=(_as_str(wb["NPS_Master"]["E2"].value)
                       if "NPS_Master" in wb.sheetnames else ""),
        stock_status=stock_status,
    )
    wb.close()
    return data
