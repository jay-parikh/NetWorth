"""Workbook generator — builds the complete xlsx from code (SPEC §3).

Everything is written through xlsxwriter: sheets, formulas, type-ahead
dropdowns, charts, comments, formats. Never edit the output by hand and never
save it through openpyxl (charts/comments would be lost) — change this module
and rebuild.
"""

from __future__ import annotations

import argparse
from datetime import date

import xlsxwriter

from . import model as M
from .guide_text import GUIDE_ROWS
from .model import PortfolioData


# ---------------------------------------------------------------- formats ---

BLUE = "#1F4E9C"       # input font
YELLOW = "#FFF2CC"     # "fill me" dashboard inputs
GREY_BG = "#F2F2F2"    # computed cells
HDR_BG = "#D9D9D9"
HINT = "#808080"
KEY = "#BFBFBF"

DATE_FMT = "dd-mm-yyyy"


def _formats(wb: xlsxwriter.Workbook) -> dict:
    def f(**kw):
        return wb.add_format(kw)

    return {
        "title": f(bold=True, font_size=14),
        "hint": f(font_color=HINT, italic=True, font_size=9),
        "header": f(bold=True, bg_color=HDR_BG, bottom=1),
        "label": f(bold=True),
        # inputs
        "in_text": f(font_color=BLUE),
        "in_num": f(font_color=BLUE, num_format="#,##0.###"),
        "in_money": f(font_color=BLUE, num_format="#,##0"),
        "in_price": f(font_color=BLUE, num_format="#,##0.00"),
        "in_rate": f(font_color=BLUE, num_format="0.0#"),
        "in_date": f(font_color=BLUE, num_format=DATE_FMT),
        "in_yellow": f(font_color=BLUE, bg_color=YELLOW, border=1, border_color="#D6C089"),
        # computed
        "c_text": f(bg_color=GREY_BG),
        "c_money": f(bg_color=GREY_BG, num_format="#,##0"),
        "c_price": f(bg_color=GREY_BG, num_format="#,##0.00"),
        "c_units": f(bg_color=GREY_BG, num_format="#,##0.000"),
        "c_pct": f(bg_color=GREY_BG, num_format="0.0%"),
        # updater-written values
        "u_price": f(num_format="#,##0.00"),
        "u_pct": f(num_format="0.0%"),
        "u_pct_bold": f(bold=True, num_format="0.0%"),
        # misc
        "money_bold": f(bold=True, num_format="#,##0"),
        "total": f(bold=True, bg_color=HDR_BG, num_format="#,##0"),
        "total_label": f(bold=True, bg_color=HDR_BG),
        "date_disp": f(num_format=DATE_FMT),
        "key": f(font_color=KEY, font_size=8),
        "section": f(bold=True, font_size=11),
        # conditional-format overlays (SPEC §6.9)
        "cf_green": f(font_color="#1F7A1F"),
        "cf_red": f(font_color="#C00000"),
    }


def _redgreen(ws, F, rng: str) -> None:
    """Green ≥ 0 / red < 0 on gain-loss and return cells (SPEC §6.9)."""
    ws.conditional_format(rng, {"type": "cell", "criteria": ">", "value": 0,
                                "format": F["cf_green"]})
    ws.conditional_format(rng, {"type": "cell", "criteria": "<", "value": 0,
                                "format": F["cf_red"]})


def _widths(ws, spec: dict[str, float]) -> None:
    for col, w in spec.items():
        ws.set_column(f"{col}:{col}", w)


def _typeahead(anchor_sheet: str, name_list: str, col: str = "C",
               anchor_col: str = "B") -> str:
    """SPEC §3.12 begins-with dropdown window formula."""
    return (f"=OFFSET({anchor_sheet}!${anchor_col}$3,"
            f"IFERROR(MATCH(${col}4&\"*\",{name_list},0),1),0,"
            f"MAX(1,COUNTIF({name_list},${col}4&\"*\")),1)")


_DROPDOWN_TIP = ("Type the first letters, then open the dropdown to see only "
                 "matching entries. Free text is allowed (a warning you can "
                 "accept) - then fill the ISIN yourself.")


def _add_dropdown(ws, rng: str, source: str, title: str) -> None:
    ws.data_validation(rng, {
        "validate": "list",
        "source": source,
        "show_error": False,
        "input_title": title,
        "input_message": _DROPDOWN_TIP,
    })


# ------------------------------------------------------------ sheet parts ---

def _sheet_head(ws, F, title: str, hint: str = "") -> None:
    ws.write("A1", title, F["title"])
    ws.set_row(0, 18)
    if hint:
        ws.write("A2", hint, F["hint"])


def _key_formula(r: int) -> str:
    return f'=IF($A{r}="","",$A{r}&"#"&COUNTIF($A$4:$A{r},$A{r}))'


def _master_lookup(scheme_cell: str, master_col: str) -> str:
    return (f'=IF({scheme_cell}="","",IFERROR(INDEX(MF_Master!${master_col}:${master_col},'
            f'MATCH({scheme_cell},MF_Master!$B:$B,0)),""))')


# ----------------------------------------------------------------- sheets ---

def _fy_end_label() -> str:
    from datetime import date as _date
    today = _date.today()
    fy_year = today.year if today <= _date(today.year, 3, 31) else today.year + 1
    return f"Expected @ 31-Mar-{fy_year}"


def _write_dashboard(wb, F, data: PortfolioData):
    ws = wb.add_worksheet("Dashboard")
    _widths(ws, {"A": 16, "B": 15, "C": 15, "D": 15, "E": 14, "F": 15, "G": 16, "H": 18})
    ws.write("A1", "FAMILY PORTFOLIO — NET WORTH TRACKER", F["title"])
    ws.set_row(0, 18)
    ws.write("A2", "As on", F["label"])
    ws.write_formula("B2", "=TODAY()", F["date_disp"])
    ws.write("D2", "Expected return % p.a.", F["label"])
    ws.write_number("E2", data.expected_return_pct, F["in_yellow"])
    ws.write_comment("E2", "Your assumed annual return for Equity and Mutual Funds - "
                           "drives only the 'Expected @ FY-end' estimate. Fixed income "
                           "uses each row's own rate.")
    ws.write("A3", "Family net worth", F["label"])
    ws.write_formula("B3", "=G16", F["money_bold"])
    ws.write("A4", "Portfolio XIRR", F["label"])
    if data.xirr.portfolio is not None:
        ws.write_number("B4", data.xirr.portfolio, F["u_pct_bold"])
    else:
        ws.write_blank("B4", None, F["u_pct_bold"])
    ws.write("D3", "Inflation % p.a.", F["label"])
    ws.write_number("E3", data.inflation_pct, F["in_yellow"])
    ws.write("D4", "Real return", F["label"])
    ws.write_formula("E4", '=IF(B4="","",(1+B4)/(1+E3/100)-1)', F["u_pct"])
    ws.write_formula(
        "F4", '=IF(B4="","",IF(B4>E3/100,"Beats inflation ✓","Below inflation ✗"))')

    headers = ["Person", "Equity", "Mutual Funds", "Fixed Deposits", "PPF", "Bonds",
               "Total", _fy_end_label()]
    ws.write_row("A5", headers, F["header"])
    ws.write_comment("H5", "Estimate of each person's total at the financial-year end: "
                           "FDs/PPF/Bonds accrue at their own rates; Equity and Mutual "
                           "Funds grow at the 'Expected return %' input (E2). Written "
                           "by the updater.")
    class_cols = [("B", "Equity!$I:$I", "Equity!$A:$A"),
                  ("C", "MutualFunds!$I:$I", "MutualFunds!$A:$A"),
                  ("D", "FixedDeposits!$I:$I", "FixedDeposits!$A:$A"),
                  ("E", "PPF!$D:$D", "PPF!$A:$A"),
                  ("F", "Bonds!$K:$K", "Bonds!$A:$A")]
    for r in range(M.DASH_PERSON_FIRST, M.DASH_PERSON_LAST + 1):
        idx = r - M.DASH_PERSON_FIRST
        name = data.persons[idx] if idx < len(data.persons) else ""
        ws.write(f"A{r}", name, F["in_yellow"])
        for col, val_rng, own_rng in class_cols:
            ws.write_formula(f"{col}{r}",
                             f'=IF($A{r}="","",SUMIFS({val_rng},{own_rng},$A{r}))',
                             F["c_money"])
        ws.write_formula(f"G{r}", f'=IF($A{r}="","",SUM(B{r}:F{r}))', F["c_money"])
        fy = data.fy_expected.get(name)
        if fy is not None:
            ws.write_number(f"H{r}", fy, F["c_money"])
    tr = M.DASH_TOTAL_ROW
    ws.write(f"A{tr}", "TOTAL", F["total_label"])
    for col in "BCDEFG":
        ws.write_formula(f"{col}{tr}", f"=SUM({col}6:{col}15)", F["total"])
    ws.write_formula(f"H{tr}", f"=IF(SUM(H6:H15)=0,\"\",SUM(H6:H15))", F["total"])

    ws.write("A18", "Allocation by asset class", F["section"])
    ws.write_row("A19", ["Asset class", "Value", "XIRR"], F["header"])
    classes = [("Equity", "B16", data.xirr.equity),
               ("Mutual Funds", "C16", data.xirr.mutual_funds),
               ("Fixed Deposits", "D16", data.xirr.fixed_deposits),
               ("PPF", "E16", data.xirr.ppf),
               ("Bonds", "F16", data.xirr.bonds)]
    for i, (label, ref, x) in enumerate(classes):
        r = 20 + i
        ws.write(f"A{r}", label)
        ws.write_formula(f"B{r}", f"={ref}", F["c_money"])
        if x is not None:
            ws.write_number(f"C{r}", x, F["u_pct"])
        else:
            ws.write_blank(f"C{r}", None, F["u_pct"])

    pie = wb.add_chart({"type": "pie"})
    pie.add_series({
        "categories": "=Dashboard!$A$20:$A$24",
        "values": "=Dashboard!$B$20:$B$24",
        "data_labels": {"percentage": True},
    })
    pie.set_title({"name": "Allocation by asset class"})
    ws.insert_chart("I4", pie, {"x_scale": 1.1, "y_scale": 1.1})

    bar = wb.add_chart({"type": "column"})
    bar.add_series({
        "name": "=Dashboard!$G$5",
        "categories": "=Dashboard!$A$6:$A$15",
        "values": "=Dashboard!$G$6:$G$15",
    })
    bar.set_title({"name": "Net worth by person"})
    bar.set_legend({"none": True})
    ws.insert_chart("I21", bar, {"x_scale": 1.1, "y_scale": 1.1})

    _redgreen(ws, F, "B4")
    _redgreen(ws, F, "E4")
    _redgreen(ws, F, "C20:C24")
    return ws


def _write_projection(wb, F):
    ws = wb.add_worksheet("Projection")
    _widths(ws, {"A": 8, "B": 24, "C": 26, "D": 24})
    _sheet_head(ws, F, "CORPUS PROJECTION — NEXT 20 YEARS",
                "Today's corpus compounding at the portfolio XIRR vs plain inflation. "
                "Assumes a constant return and no future contributions. Edit inflation "
                "on the Dashboard (yellow cell); the updaters refresh the XIRR.")
    ws.write_row("A3", ["Year", "Corpus @ portfolio XIRR",
                        "Corpus growing at inflation", "Real corpus (today's money)"],
                 F["header"])
    for n in range(M.PROJECTION_YEARS + 1):
        r = M.FIRST_DATA_ROW + n
        ws.write_formula(f"A{r}", f"=YEAR(TODAY())+{n}", F["c_text"])
        ws.write_formula(f"B{r}", f"=Dashboard!$B$3*(1+Dashboard!$B$4)^{n}", F["c_money"])
        ws.write_formula(f"C{r}", f"=Dashboard!$B$3*(1+Dashboard!$E$3/100)^{n}", F["c_money"])
        ws.write_formula(f"D{r}", f"=B{r}/(1+Dashboard!$E$3/100)^{n}", F["c_money"])

    line = wb.add_chart({"type": "line"})
    last = M.FIRST_DATA_ROW + M.PROJECTION_YEARS
    for col in "BCD":
        line.add_series({
            "name": f"=Projection!${col}$3",
            "categories": f"=Projection!$A$4:$A${last}",
            "values": f"=Projection!${col}$4:${col}${last}",
        })
    line.set_title({"name": "Corpus trajectory — portfolio return vs inflation (20 years)"})
    line.set_x_axis({"name": "Year"})
    ws.insert_chart("F3", line, {"x_scale": 1.6, "y_scale": 1.4})
    return ws


_PERSON_BLOCKS = [
    # (title, block const, headers, source sheet, source cols, key col, fmts)
    ("EQUITY", M.PERSON_EQ_BLOCK,
     ["ISIN", "Scrip", "Qty", "Avg cost", "Cur. val", "Net chg.", "XIRR"],
     "Equity", ["B", "C", "D", "E", "I", "K", "N"], "P",
     ["c_text", "c_text", "c_text", "c_price", "c_money", "c_money", "c_pct"]),
    ("MUTUAL FUNDS", M.PERSON_MF_BLOCK,
     ["Fund House", "Scheme", "Units", "Cur NAV", "Cur. val", "Net chg.", "XIRR"],
     "MutualFunds", ["B", "C", "E", "G", "I", "J", "L"], "N",
     ["c_text", "c_text", "c_units", "c_price", "c_money", "c_money", "c_pct"]),
    ("FIXED DEPOSITS", M.PERSON_FD_BLOCK,
     ["Bank", "Principal", "Rate %", "Maturity", "Value today"],
     "FixedDeposits", ["B", "D", "E", "G", "I"], "L",
     ["c_text", "c_money", "c_text", "c_text", "c_money"]),
    ("PPF", M.PERSON_PPF_BLOCK,
     ["Institution", "Balance", "As-on"],
     "PPF", ["B", "D", "E"], "I",
     ["c_text", "c_money", "c_text"]),
    ("BONDS", M.PERSON_BOND_BLOCK,
     ["Issuer / Bond", "Qty", "Buy Price", "Cur Price", "Cur. val", "Net chg."],
     "Bonds", ["B", "D", "F", "G", "K", "L"], "N",
     ["c_text", "c_text", "c_price", "c_price", "c_money", "c_money"]),
]


def _write_person(wb, F, name: str):
    ws = wb.add_worksheet(name)
    _widths(ws, {"A": 26, "B": 30, "C": 12, "D": 12, "E": 14, "F": 13, "G": 10})
    ws.write("A1", f"{name} — PORTFOLIO", F["title"])
    ws.set_row(0, 18)
    ws.write("A2", "Owner", F["label"])
    ws.write("B2", name)
    ws.write("A3", "Net worth", F["label"])
    ws.write_formula("B3", "=B11", F["money_bold"])

    ws.write_row("A5", ["Asset class", "Value", "# holdings"], F["header"])
    rows = [("Equity", "Equity!$I:$I", "Equity!$A:$A"),
            ("Mutual Funds", "MutualFunds!$I:$I", "MutualFunds!$A:$A"),
            ("Fixed Deposits", "FixedDeposits!$I:$I", "FixedDeposits!$A:$A"),
            ("PPF", "PPF!$D:$D", "PPF!$A:$A"),
            ("Bonds", "Bonds!$K:$K", "Bonds!$A:$A")]
    for i, (label, val_rng, own_rng) in enumerate(rows):
        r = 6 + i
        ws.write(f"A{r}", label)
        ws.write_formula(f"B{r}", f"=SUMIFS({val_rng},{own_rng},$B$2)", F["c_money"])
        ws.write_formula(f"C{r}", f"=COUNTIF({own_rng},$B$2)", F["c_text"])
    ws.write("A11", "Total", F["total_label"])
    ws.write_formula("B11", "=SUM(B6:B10)", F["total"])
    ws.write_formula("C11", "=SUM(C6:C10)", F["total"])

    pie = wb.add_chart({"type": "pie"})
    pie.add_series({
        "categories": f"='{name}'!$A$6:$A$10",
        "values": f"='{name}'!$B$6:$B$10",
        "data_labels": {"percentage": True},
    })
    pie.set_title({"name": f"{name} — allocation"})
    ws.insert_chart("E4", pie)

    for title, (title_row, first, last), headers, src, cols, key_col, fmts in _PERSON_BLOCKS:
        ws.write(f"A{title_row}", title, F["section"])
        ws.write_row(f"A{title_row + 1}", headers, F["header"])
        for r in range(first, last + 1):
            slot = r - first + 1
            for ci, (col, fmt) in enumerate(zip(cols, fmts)):
                out = chr(ord("A") + ci)
                ws.write_formula(
                    f"{out}{r}",
                    f'=IFERROR(INDEX({src}!${col}:${col},'
                    f'MATCH($B$2&"#"&{slot},{src}!${key_col}:${key_col},0)),"")',
                    F[fmt])
        if title in ("EQUITY", "MUTUAL FUNDS"):
            _redgreen(ws, F, f"F{first}:G{last}")      # Net chg. + XIRR
        elif title == "BONDS":
            _redgreen(ws, F, f"F{first}:F{last}")      # Net chg.
    ws.freeze_panes("A5")
    return ws


def _write_equity(wb, F, data: PortfolioData):
    ws = wb.add_worksheet("Equity")
    _widths(ws, {"A": 12, "B": 15, "C": 34, "D": 10, "E": 12, "F": 13, "G": 12,
                 "H": 16, "I": 14, "J": 14, "K": 13, "L": 12, "M": 13, "N": 9, "P": 13})
    _sheet_head(ws, F, "EQUITY HOLDINGS",
                "Yellow-ish/blue cells are inputs. Pick the Scrip from the dropdown — "
                "ISIN fills itself. Prices refresh via the updater.")
    ws.write_row("A3", ["Owner", "ISIN", "Scrip", "Quantity", "Avg. cost",
                        "Closing Price", "Prev. close", "Closing Price Date",
                        "Cur. val", "Invested", "Net chg.", "Day chg.",
                        "Cost date", "XIRR"], F["header"])
    ws.write("P3", "Key", F["header"])

    by_row = {M.FIRST_DATA_ROW + i: row for i, row in enumerate(data.equity)}
    for r in range(M.FIRST_DATA_ROW, M.EQUITY_LAST_ROW + 1):
        row = by_row.get(r)
        if row and row.isin_override:
            ws.write(f"B{r}", row.isin_override, F["in_text"])
        else:
            ws.write_formula(
                f"B{r}",
                f'=IF($C{r}="","",IFERROR(INDEX(Stock_Master!$C:$C,'
                f'MATCH($C{r},Stock_Master!$B:$B,0)),""))', F["c_text"])
        if row:
            ws.write(f"A{r}", row.owner, F["in_text"])
            ws.write(f"C{r}", row.scrip, F["in_text"])
            if row.qty is not None:
                ws.write_number(f"D{r}", row.qty, F["in_num"])
            if row.avg_cost is not None:
                ws.write_number(f"E{r}", row.avg_cost, F["in_price"])
            if row.close is not None:
                ws.write_number(f"F{r}", row.close, F["u_price"])
            if row.prev_close is not None:
                ws.write_number(f"G{r}", row.prev_close, F["u_price"])
            if row.close_date:
                ws.write(f"H{r}", row.close_date)
            if row.cost_date is not None:
                ws.write_datetime(f"M{r}", row.cost_date, F["in_date"])
        ws.write_formula(f"I{r}", f'=IF($D{r}="","",$D{r}*$F{r})', F["c_money"])
        ws.write_formula(f"J{r}", f'=IF(OR($D{r}="",$E{r}=""),"",$D{r}*$E{r})', F["c_money"])
        ws.write_formula(f"K{r}", f'=IF(OR($E{r}="",$D{r}=""),"",$I{r}-$J{r})', F["c_money"])
        ws.write_formula(f"L{r}", f'=IF(OR($G{r}="",$D{r}=""),"",$D{r}*($F{r}-$G{r}))',
                         F["c_money"])
        ws.write_formula(
            f"N{r}",
            f'=IF(OR($M{r}="",N($J{r})=0,$I{r}="",TODAY()<=$M{r}),"",'
            f'($I{r}/$J{r})^(365/(TODAY()-$M{r}))-1)', F["c_pct"])
        ws.write_formula(f"P{r}", _key_formula(r), F["key"])

    tr = M.EQUITY_TOTAL_ROW
    ws.write(f"C{tr}", "TOTAL", F["total_label"])
    for col in "IJKL":
        ws.write_formula(f"{col}{tr}", f"=SUM({col}4:{col}{M.EQUITY_LAST_ROW})", F["total"])
    ws.write(f"M{tr}", "XIRR →", F["label"])
    if data.xirr.equity is not None:
        ws.write_number(f"N{tr}", data.xirr.equity, F["u_pct_bold"])
    else:
        ws.write_blank(f"N{tr}", None, F["u_pct_bold"])
    ws.write_comment(f"N{tr}",
                     "Portfolio equity XIRR - written by the updater (plain value).")

    _add_dropdown(ws, f"C4:C{M.EQUITY_LAST_ROW}",
                  _typeahead("Stock_Master", "Stock_NameList"), "Scrip")
    _redgreen(ws, F, f"K4:L{M.EQUITY_LAST_ROW}")
    _redgreen(ws, F, f"N4:N{M.EQUITY_LAST_ROW}")
    _redgreen(ws, F, f"K{tr}:L{tr}")
    _redgreen(ws, F, f"N{tr}")
    ws.freeze_panes("A4")
    return ws


def _write_mutualfunds(wb, F, data: PortfolioData):
    ws = wb.add_worksheet("MutualFunds")
    _widths(ws, {"A": 12, "B": 26, "C": 46, "D": 15, "E": 12, "F": 13, "G": 12,
                 "H": 14, "I": 14, "J": 13, "K": 10, "L": 9, "N": 13})
    _sheet_head(ws, F, "MUTUAL FUNDS",
                "One row per fund. Units and Invested come from the MF_SIP ledger — "
                "enter purchases there. Pick the Scheme from the dropdown.")
    ws.write_row("A3", ["Owner", "Fund House", "Scheme Name", "ISIN", "Units",
                        "Avg cost NAV", "Current NAV", "Invested", "Cur. val",
                        "Net chg.", "Return %", "XIRR"], F["header"])
    ws.write("N3", "Key", F["header"])
    ws.write_comment("L3", "XIRR (annualised, cashflow-dated return) is computed and "
                           "written by the updater each time you run it. Run it after "
                           "adding or changing purchases.")

    by_row = {M.FIRST_DATA_ROW + i: row for i, row in enumerate(data.mutual_funds)}
    for r in range(M.FIRST_DATA_ROW, M.MF_LAST_ROW + 1):
        row = by_row.get(r)
        if row and row.fund_house_override:
            ws.write(f"B{r}", row.fund_house_override, F["in_text"])
        else:
            ws.write_formula(f"B{r}", _master_lookup(f"$C{r}", "A"), F["c_text"])
        if row and row.isin_override:
            ws.write(f"D{r}", row.isin_override, F["in_text"])
        else:
            ws.write_formula(f"D{r}", _master_lookup(f"$C{r}", "C"), F["c_text"])
        if row:
            ws.write(f"A{r}", row.owner, F["in_text"])
            ws.write(f"C{r}", row.scheme, F["in_text"])
            if row.current_nav is not None:
                ws.write_number(f"G{r}", row.current_nav, F["u_price"])
            if row.xirr is not None:
                ws.write_number(f"L{r}", row.xirr, F["u_pct"])
        ws.write_formula(f"E{r}",
                         f'=IF($A{r}="","",SUMIFS(MF_SIP!$H:$H,MF_SIP!$A:$A,$A{r},'
                         f'MF_SIP!$D:$D,$D{r}))', F["c_units"])
        ws.write_formula(f"F{r}", f'=IF(OR($A{r}="",N($E{r})=0),"",$H{r}/$E{r})',
                         F["c_price"])
        ws.write_formula(f"H{r}",
                         f'=IF($A{r}="","",SUMIFS(MF_SIP!$F:$F,MF_SIP!$A:$A,$A{r},'
                         f'MF_SIP!$D:$D,$D{r}))', F["c_money"])
        ws.write_formula(f"I{r}", f'=IF(OR(N($E{r})=0,$G{r}=""),"",$E{r}*$G{r})',
                         F["c_money"])
        ws.write_formula(f"J{r}", f'=IF(N($E{r})=0,"",$I{r}-$H{r})', F["c_money"])
        ws.write_formula(f"K{r}", f'=IF(OR(N($H{r})=0,N($J{r})=0),"",$J{r}/$H{r})',
                         F["c_pct"])
        ws.write_formula(f"N{r}", _key_formula(r), F["key"])

    tr = M.MF_TOTAL_ROW
    ws.write(f"C{tr}", "TOTAL", F["total_label"])
    for col in "IJ":
        ws.write_formula(f"{col}{tr}", f"=SUM({col}4:{col}{M.MF_LAST_ROW})", F["total"])
    ws.write_formula(f"K{tr}", f'=IF(N(H{tr})=0,"",J{tr}/H{tr})', F["u_pct"])
    if data.xirr.mutual_funds is not None:
        ws.write_number(f"L{tr}", data.xirr.mutual_funds, F["u_pct_bold"])
    else:
        ws.write_blank(f"L{tr}", None, F["u_pct_bold"])
    ws.write_formula(f"H{tr}", f"=SUM(H4:H{M.MF_LAST_ROW})", F["total"])
    ws.write_formula(f"I{tr}", f"=SUM(I4:I{M.MF_LAST_ROW})", F["total"])

    _add_dropdown(ws, f"C4:C{M.MF_LAST_ROW}",
                  _typeahead("MF_Master", "MF_SchemeList"), "Scheme Name")
    _redgreen(ws, F, f"J4:L{M.MF_LAST_ROW}")
    _redgreen(ws, F, f"J{tr}:L{tr}")
    ws.freeze_panes("A4")
    return ws


def _write_mf_sip(wb, F, data: PortfolioData):
    ws = wb.add_worksheet("MF_SIP")
    _widths(ws, {"A": 12, "B": 26, "C": 46, "D": 15, "E": 12, "F": 12, "G": 12,
                 "H": 12, "J": 12})
    _sheet_head(ws, F, "MUTUAL FUND SIP / PURCHASE LEDGER",
                "One row per purchase — SIP instalment or lump sum. "
                "Redemption = negative Amount.")
    ws.write("J1", "Portfolio MF XIRR", F["label"])
    ws.write_comment("J1", "Written by the updater (plain value). Run it after any change.")
    if data.xirr.mutual_funds is not None:
        ws.write_number("J2", data.xirr.mutual_funds, F["u_pct_bold"])
    else:
        ws.write_blank("J2", None, F["u_pct_bold"])
    ws.write_row("A3", ["Owner", "Fund House", "Scheme Name", "ISIN", "Date",
                        "Amount", "NAV on date", "Units"], F["header"])
    ws.write_comment("H3", "Units = Amount / NAV on date (auto). If you know units but "
                           "not the NAV, you may overtype Units with the number from "
                           "your statement.")

    by_row = {M.FIRST_DATA_ROW + i: row for i, row in enumerate(data.sip)}
    for r in range(M.FIRST_DATA_ROW, M.SIP_LAST_ROW + 1):
        row = by_row.get(r)
        if row and row.fund_house_override:
            ws.write(f"B{r}", row.fund_house_override, F["in_text"])
        else:
            ws.write_formula(f"B{r}", _master_lookup(f"$C{r}", "A"), F["c_text"])
        if row and row.isin_override:
            ws.write(f"D{r}", row.isin_override, F["in_text"])
        else:
            ws.write_formula(f"D{r}", _master_lookup(f"$C{r}", "C"), F["c_text"])
        if row:
            ws.write(f"A{r}", row.owner, F["in_text"])
            ws.write(f"C{r}", row.scheme, F["in_text"])
            if row.txn_date is not None:
                ws.write_datetime(f"E{r}", row.txn_date, F["in_date"])
            if row.amount is not None:
                ws.write_number(f"F{r}", row.amount, F["in_money"])
            if row.nav is not None:
                ws.write_number(f"G{r}", row.nav, F["in_price"])
        if row and row.units_override is not None:
            ws.write_number(f"H{r}", row.units_override, F["in_num"])
        else:
            ws.write_formula(f"H{r}", f'=IF(OR($F{r}="",$G{r}=""),"",$F{r}/$G{r})',
                             F["c_units"])

    _add_dropdown(ws, f"C4:C{M.SIP_LAST_ROW}",
                  _typeahead("MF_Master", "MF_SchemeList"), "Scheme Name")
    ws.freeze_panes("A4")
    return ws


def _write_master(wb, F, name: str, title: str, hint: str,
                  headers: list[str], rows: list[tuple], refreshed: str):
    ws = wb.add_worksheet(name)
    _widths(ws, {"A": 34, "B": 60, "C": 16, "D": 12, "E": 12})
    ws.write("A1", title, F["title"])
    ws.set_row(0, 18)
    ws.write("A2", hint, F["hint"])
    ws.write("D2", "Refreshed:", F["label"])
    ws.write("E2", refreshed, F["in_text"])
    ws.write_row("A3", headers, F["header"])
    r = M.FIRST_DATA_ROW - 1  # 0-based row index for write_row
    for tup in rows:
        ws.write_row(r, 0, tup)
        r += 1
    ws.freeze_panes("A4")
    return ws


def _write_fd(wb, F, data: PortfolioData):
    ws = wb.add_worksheet("FixedDeposits")
    _widths(ws, {"A": 12, "B": 24, "C": 16, "D": 13, "E": 11, "F": 12, "G": 14,
                 "H": 10, "I": 17, "J": 15, "L": 13})
    _sheet_head(ws, F, "FIXED DEPOSITS",
                "Enter Owner, Bank, FD no., Principal, Rate % p.a., Start & Maturity "
                "dates and compounding/yr. Value columns compute themselves.")
    ws.write_row("A3", ["Owner", "Bank / Institution", "FD No.", "Principal",
                        "Rate % p.a.", "Start Date", "Maturity Date", "Comp./yr",
                        "Value as on today", "Maturity Value"], F["header"])
    ws.write("L3", "Key", F["header"])

    by_row = {M.FIRST_DATA_ROW + i: row for i, row in enumerate(data.fixed_deposits)}
    for r in range(M.FIRST_DATA_ROW, M.FD_LAST_ROW + 1):
        row = by_row.get(r)
        if row:
            ws.write(f"A{r}", row.owner, F["in_text"])
            ws.write(f"B{r}", row.bank, F["in_text"])
            ws.write(f"C{r}", row.fd_no, F["in_text"])
            if row.principal is not None:
                ws.write_number(f"D{r}", row.principal, F["in_money"])
            if row.rate is not None:
                ws.write_number(f"E{r}", row.rate, F["in_rate"])
            if row.start is not None:
                ws.write_datetime(f"F{r}", row.start, F["in_date"])
            if row.maturity is not None:
                ws.write_datetime(f"G{r}", row.maturity, F["in_date"])
            if row.comp_per_year is not None:
                ws.write_number(f"H{r}", row.comp_per_year, F["in_text"])
        ws.write_formula(
            f"I{r}",
            f'=IF($D{r}="","",$D{r}*(1+($E{r}/100)/$H{r})^'
            f'($H{r}*YEARFRAC($F{r},MIN(TODAY(),$G{r}))))', F["c_money"])
        ws.write_formula(
            f"J{r}",
            f'=IF($D{r}="","",$D{r}*(1+($E{r}/100)/$H{r})^($H{r}*YEARFRAC($F{r},$G{r})))',
            F["c_money"])
        ws.write_formula(f"L{r}", _key_formula(r), F["key"])

    tr = M.FD_TOTAL_ROW
    ws.write(f"C{tr}", "TOTAL", F["total_label"])
    for col in "IJ":
        ws.write_formula(f"{col}{tr}", f"=SUM({col}4:{col}{M.FD_LAST_ROW})", F["total"])
    _add_dropdown(ws, f"B4:B{M.FD_LAST_ROW}",
                  _typeahead("Bank_Master", "Bank_NameList", col="B", anchor_col="A"),
                  "Bank / Institution")
    ws.freeze_panes("A4")
    return ws


def _write_bank_master(wb, F):
    from .model import load_banks
    ws = wb.add_worksheet("Bank_Master")
    _widths(ws, {"A": 36, "B": 22})
    ws.write("A1", "BANK MASTER (India)", F["title"])
    ws.set_row(0, 18)
    ws.write("A2", "Feeds the Bank dropdown on the FixedDeposits sheet. Bundled with "
                   "the app and refreshed by new releases. Free text is always allowed "
                   "on the FD sheet for anything not listed.", F["hint"])
    ws.write_row("A3", ["Bank Name", "Type"], F["header"])
    for i, (name, kind) in enumerate(load_banks()):
        ws.write(3 + i, 0, name)
        ws.write(3 + i, 1, kind)
    ws.freeze_panes("A4")
    return ws


def _write_ppf(wb, F, data: PortfolioData):
    ws = wb.add_worksheet("PPF")
    _widths(ws, {"A": 12, "B": 20, "C": 16, "D": 16, "E": 14, "F": 12, "G": 24, "I": 13})
    _sheet_head(ws, F, "PPF ACCOUNTS",
                "Enter Owner, Institution, Account no., Current balance and the date "
                "it was true. XIRR treats the balance as growing at Rate% from that date.")
    ws.write_row("A3", ["Owner", "Institution", "Account No.", "Current Balance",
                        "Balance as-on", "Rate % (ref)", "Notes"], F["header"])
    ws.write("I3", "Key", F["header"])

    by_row = {M.FIRST_DATA_ROW + i: row for i, row in enumerate(data.ppf)}
    for r in range(M.FIRST_DATA_ROW, M.PPF_LAST_ROW + 1):
        row = by_row.get(r)
        if row:
            ws.write(f"A{r}", row.owner, F["in_text"])
            ws.write(f"B{r}", row.institution, F["in_text"])
            ws.write(f"C{r}", row.account_no, F["in_text"])
            if row.balance is not None:
                ws.write_number(f"D{r}", row.balance, F["in_money"])
            if row.as_on is not None:
                ws.write_datetime(f"E{r}", row.as_on, F["in_date"])
            if row.rate is not None:
                ws.write_number(f"F{r}", row.rate, F["in_rate"])
            if row.notes:
                ws.write(f"G{r}", row.notes, F["in_text"])
        ws.write_formula(f"I{r}", _key_formula(r), F["key"])

    tr = M.PPF_TOTAL_ROW
    ws.write(f"C{tr}", "TOTAL", F["total_label"])
    ws.write_formula(f"D{tr}", f"=SUM(D4:D{M.PPF_LAST_ROW})", F["total"])
    ws.freeze_panes("A4")
    return ws


def _write_bonds(wb, F, data: PortfolioData):
    ws = wb.add_worksheet("Bonds")
    _widths(ws, {"A": 12, "B": 26, "C": 16, "D": 8, "E": 12, "F": 11, "G": 13,
                 "H": 13, "I": 14, "J": 13, "K": 13, "L": 12, "M": 12, "N": 13,
                 "O": 15, "P": 19})
    _sheet_head(ws, F, "CORPORATE / OTHER BONDS",
                "Fill Buy Date — rows without it stay out of XIRR. Current Price is "
                "refreshed by the updater when the ISIN trades.")
    ws.write_row("A3", ["Owner", "Issuer / Bond", "ISIN", "Qty", "Face Value",
                        "Buy Price", "Current Price", "Coupon % p.a.",
                        "Maturity Date", "Invested", "Cur. val", "Net chg.",
                        "Buy Date"], F["header"])
    ws.write("N3", "Key", F["header"])
    ws.write("O3", "Maturity Value", F["header"])
    ws.write("P3", "Coupons till maturity", F["header"])
    ws.write_comment("O3", "Redemption at face value (Qty x Face). For cumulative/"
                           "zero-coupon bonds set Face Value to the redemption amount "
                           "and Coupon to 0.")
    ws.write_comment("P3", "Simple sum of remaining annual coupons to maturity "
                           "(not reinvested).")

    by_row = {M.FIRST_DATA_ROW + i: row for i, row in enumerate(data.bonds)}
    for r in range(M.FIRST_DATA_ROW, M.BOND_LAST_ROW + 1):
        row = by_row.get(r)
        if row:
            ws.write(f"A{r}", row.owner, F["in_text"])
            ws.write(f"B{r}", row.issuer, F["in_text"])
            ws.write(f"C{r}", row.isin, F["in_text"])
            if row.qty is not None:
                ws.write_number(f"D{r}", row.qty, F["in_num"])
            if row.face is not None:
                ws.write_number(f"E{r}", row.face, F["in_money"])
            if row.buy_price is not None:
                ws.write_number(f"F{r}", row.buy_price, F["in_price"])
            if row.cur_price is not None:
                ws.write_number(f"G{r}", row.cur_price, F["u_price"])
            if row.coupon is not None:
                ws.write_number(f"H{r}", row.coupon, F["in_rate"])
            if row.maturity is not None:
                ws.write_datetime(f"I{r}", row.maturity, F["in_date"])
            if row.buy_date is not None:
                ws.write_datetime(f"M{r}", row.buy_date, F["in_date"])
        ws.write_formula(f"J{r}", f'=IF(OR($D{r}="",$F{r}=""),"",$D{r}*$F{r})', F["c_money"])
        ws.write_formula(f"K{r}", f'=IF(OR($D{r}="",$G{r}=""),"",$D{r}*$G{r})', F["c_money"])
        ws.write_formula(f"L{r}", f'=IF(OR($F{r}="",$D{r}=""),"",$K{r}-$J{r})', F["c_money"])
        ws.write_formula(f"N{r}", _key_formula(r), F["key"])
        ws.write_formula(f"O{r}", f'=IF(OR($D{r}="",$E{r}=""),"",$D{r}*$E{r})', F["c_money"])
        ws.write_formula(
            f"P{r}",
            f'=IF(OR($D{r}="",$E{r}="",$H{r}="",$I{r}="",$I{r}<=TODAY()),"",'
            f'$D{r}*$E{r}*($H{r}/100)*YEARFRAC(TODAY(),$I{r}))', F["c_money"])

    tr = M.BOND_TOTAL_ROW
    ws.write(f"C{tr}", "TOTAL", F["total_label"])
    for col in "JKLOP":
        ws.write_formula(f"{col}{tr}", f"=SUM({col}4:{col}{M.BOND_LAST_ROW})", F["total"])
    _redgreen(ws, F, f"L4:L{M.BOND_LAST_ROW}")
    _redgreen(ws, F, f"L{tr}")
    ws.freeze_panes("A4")
    return ws


def _write_by_scrip(wb, F, data: PortfolioData):
    ws = wb.add_worksheet("By Scrip")
    _widths(ws, {"A": 16, "B": 34, "C": 11})
    for i, p in enumerate(data.persons[:10]):
        ws.set_column(3 + i, 3 + i, 11)
    _sheet_head(ws, F, "FAMILY EXPOSURE BY SCRIP",
                "Total family quantity and current value per scrip, split by person. "
                "Live from the Equity sheet. To add a scrip not listed, type its ISIN "
                "and name in a blank row.")
    person_cols = [chr(ord("D") + i) for i in range(len(data.persons))]
    curval_col = chr(ord("D") + len(data.persons))
    ws.write_row("A3", ["ISIN", "Scrip", "Total Qty"] + list(data.persons) + ["Cur. val"],
                 F["header"])

    by_row = {M.FIRST_DATA_ROW + i: row for i, row in enumerate(data.by_scrip)}
    for r in range(M.FIRST_DATA_ROW, M.BYSCRIP_LAST_ROW + 1):
        row = by_row.get(r)
        if row:
            ws.write(f"A{r}", row.isin, F["in_text"])
            ws.write(f"B{r}", row.name, F["in_text"])
        ws.write_formula(f"C{r}",
                         f'=IF($A{r}="","",SUMIF(Equity!$B:$B,$A{r},Equity!$D:$D))',
                         F["c_text"])
        for col, person in zip(person_cols, data.persons):
            ws.write_formula(f"{col}{r}",
                             f'=IF($A{r}="","",SUMIFS(Equity!$D:$D,Equity!$B:$B,$A{r},'
                             f'Equity!$A:$A,"{person}"))', F["c_text"])
        ws.write_formula(f"{curval_col}{r}",
                         f'=IF($A{r}="","",SUMIF(Equity!$B:$B,$A{r},Equity!$I:$I))',
                         F["c_money"])
    ws.freeze_panes("A4")
    return ws


def _write_guide(wb, F):
    ws = wb.add_worksheet("Guide")
    ws.set_column("A:A", 100)
    for i, row in enumerate(GUIDE_ROWS):
        fmt = F["title"] if i == 0 else None
        for c, text in enumerate(row):
            if text:
                ws.write(i, c, text, fmt)
    return ws


# ------------------------------------------------------------------ build ---

def build_workbook(data: PortfolioData, out_path: str) -> None:
    wb = xlsxwriter.Workbook(out_path, {"default_date_format": DATE_FMT})
    F = _formats(wb)

    wb.define_name(
        "MF_SchemeList",
        "=MF_Master!$B$4:INDEX(MF_Master!$B:$B,COUNTA(MF_Master!$B:$B)+2)")
    wb.define_name(
        "Stock_NameList",
        "=Stock_Master!$B$4:INDEX(Stock_Master!$B:$B,COUNTA(Stock_Master!$B:$B)+2)")
    wb.define_name(
        "Bank_NameList",
        "=Bank_Master!$A$4:INDEX(Bank_Master!$A:$A,COUNTA(Bank_Master!$A:$A)+2)")

    # tab order = SPEC §3.1
    _write_dashboard(wb, F, data)
    _write_projection(wb, F)
    for person in data.persons:
        _write_person(wb, F, person)
    _write_equity(wb, F, data)
    _write_mutualfunds(wb, F, data)
    _write_mf_sip(wb, F, data)
    _write_master(wb, F, "MF_Master", "MUTUAL FUND MASTER (AMFI)",
                  "Every AMFI scheme with an ISIN — feeds the Scheme dropdowns on "
                  "MutualFunds and MF_SIP. Kept sorted by scheme name (the dropdown "
                  "filter needs that). Do not edit by hand: the updater refreshes it.",
                  ["Fund Name", "Scheme Name", "ISIN"],
                  data.masters.mf_rows, data.masters.mf_refreshed)
    _write_master(wb, F, "Stock_Master", "STOCK MASTER (BSE/NSE bhavcopy)",
                  "Feeds the Scrip dropdown on the Equity sheet. Auto-refreshed by the "
                  "updater: newly listed stocks are added, existing names are kept "
                  "stable. Do not edit by hand.",
                  ["Symbol", "Stock Name", "ISIN"],
                  data.masters.stock_rows, data.masters.stock_refreshed)
    _write_bank_master(wb, F)
    _write_fd(wb, F, data)
    _write_ppf(wb, F, data)
    _write_bonds(wb, F, data)
    _write_by_scrip(wb, F, data)
    _write_guide(wb, F)

    wb.close()


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Build the template workbook from code.")
    parser.add_argument("-o", "--out", default=M.TEMPLATE_FILENAME,
                        help=f"output path (default: {M.TEMPLATE_FILENAME})")
    args = parser.parse_args(argv)

    from .sample_data import sample_portfolio
    data = sample_portfolio()
    build_workbook(data, args.out)
    print(f"Wrote {args.out}: {len(data.masters.mf_rows)} MF schemes, "
          f"{len(data.masters.stock_rows)} stocks, {len(data.persons)} persons.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
