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
from .model import (ASSET_CLASSES, PortfolioData, class_has_data,
                    effective_enabled, enabled_classes)


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
        # amber = degraded data: stale price, suspended/delisted, FMV-estimated cost
        "cf_amber": f(bg_color="#FFE8C4"),
        "amber_price": f(bg_color="#FFE8C4", num_format="#,##0.00"),
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


# Excel silently drops a validation whose input message exceeds 255 chars —
# keep this tip well under that.
_DROPDOWN_TIP = ("FILTER IN 2 STEPS: type the first letters, press ENTER, "
                 "then re-open the dropdown (arrow or Alt+Down) - only "
                 "matching names remain. Excel does not suggest while "
                 "typing. Free text is allowed for names not in the list.")


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


def _class_sumifs(cls, owner_ref: str) -> str:
    """SUMIFS for a class total; shared-sheet classes add their Class filter."""
    extra = (f',{cls.class_filter[0]},"{cls.class_filter[1]}"'
             if cls.class_filter else "")
    return f"SUMIFS({cls.value_col},{cls.owner_col},{owner_ref}{extra})"


def _class_countifs(cls, owner_ref: str) -> str:
    extra = (f',{cls.class_filter[0]},"{cls.class_filter[1]}"'
             if cls.class_filter else "")
    return f"COUNTIFS({cls.owner_col},{owner_ref}{extra})"


def _history_classes(data: PortfolioData):
    """History columns: classes that are enabled OR carry nonzero history —
    an old trend line never disappears because a class was switched off."""
    return [c for c in ASSET_CLASSES
            if effective_enabled(data, c)
            or any(getattr(s, c.key, 0) for s in data.history)]


def _write_settings(wb, F, data: PortfolioData):
    ws = wb.add_worksheet("Settings")
    _widths(ws, {"A": 18, "B": 10, "C": 10, "D": 14, "E": 44})
    _sheet_head(ws, F, "SETTINGS",
                "Yes/No shows or hides each asset class - a class holding data "
                "is never hidden. Target % feeds the Dashboard drift view. "
                "Run the update after changing anything here.")
    ws.write_row("A3", ["Asset class", "Enabled", "Target %", "Status", "Notes"],
                 F["header"])
    ws.write_comment("B3", "Yes = show this class's sheets and Dashboard "
                           "presence. No = hide them. Data is never deleted.")
    ws.write_comment("C3", "Your target share of the family total, in percent. "
                           "Leave blank for no target. Drives the Dashboard "
                           "drift view.")
    for i, cls in enumerate(ASSET_CLASSES):
        r = M.SETTINGS_FIRST_ROW + i
        setting = data.class_settings.get(cls.key)
        enabled = setting.enabled if setting else cls.default_enabled
        target = setting.target_pct if setting else None
        ws.write(f"A{r}", cls.label)
        ws.write(f"B{r}", "Yes" if enabled else "No", F["in_text"])
        if target is not None:
            ws.write_number(f"C{r}", target, F["in_rate"])
        else:
            ws.write_blank(f"C{r}", None, F["in_rate"])
        if enabled:
            status = "On"
        elif class_has_data(data, cls.key):
            status = "On (has data)"
        else:
            status = "Off"
        ws.write(f"D{r}", status, F["c_text"])
        if status == "On (has data)":
            ws.write(f"E{r}", "Holds rows, so it stays visible - delete or "
                              "move them to hide it.", F["hint"])
    ws.data_validation(
        f"B{M.SETTINGS_FIRST_ROW}:B{M.SETTINGS_LAST_ROW}",
        {"validate": "list", "source": ["Yes", "No"], "show_error": False,
         "input_title": "Show this class?",
         "input_message": "Yes shows its sheets, No hides them. "
                          "Data is never deleted."})
    tol = M.SETTINGS_TOL_ROW
    ws.write(f"A{tol}", "Drift tolerance (± % points)", F["label"])
    ws.write_number(f"B{tol}", data.drift_tolerance_pct, F["in_rate"])
    ws.write_comment(f"B{tol}", "How far a class may drift from its target "
                                "before the Dashboard flags it red. 5 means "
                                "±5 percentage points.")
    tot = M.SETTINGS_SUM_ROW
    ws.write(f"A{tot}", "Targets total", F["label"])
    ws.write_formula(
        f"B{tot}",
        f"=SUM(C{M.SETTINGS_FIRST_ROW}:C{M.SETTINGS_LAST_ROW})", F["c_text"])
    ws.conditional_format(f"B{tot}", {
        "type": "formula",
        "criteria": f"=AND(B{tot}>0,B{tot}<>100)",
        "format": F["cf_amber"]})
    ws.freeze_panes("A4")
    return ws


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
    ws.write_formula(
        "B3", f'={chr(ord("B") + len(enabled_classes(data)))}{M.DASH_TOTAL_ROW}',
        F["money_bold"])
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

    enabled = enabled_classes(data)
    n = len(enabled)
    col_of = {c.key: chr(ord("B") + i) for i, c in enumerate(enabled)}
    total_col = chr(ord("B") + n)
    fy_col = chr(ord("B") + n + 1)

    ws.write_row("A5", ["Person"] + [c.label for c in enabled]
                 + ["Total", _fy_end_label()], F["header"])
    ws.write_comment(f"{fy_col}5",
                     "Estimate of each person's total at the financial-year end: "
                     "FDs/PPF/Bonds accrue at their own rates; Equity and Mutual "
                     "Funds grow at the 'Expected return %' input (E2). Written "
                     "by the updater.")
    for r in range(M.DASH_PERSON_FIRST, M.DASH_PERSON_LAST + 1):
        idx = r - M.DASH_PERSON_FIRST
        name = data.persons[idx] if idx < len(data.persons) else ""
        ws.write(f"A{r}", name, F["in_yellow"])
        for cls in enabled:
            ws.write_formula(
                f"{col_of[cls.key]}{r}",
                f'=IF($A{r}="","",{_class_sumifs(cls, f"$A{r}")})',
                F["c_money"])
        if n:
            ws.write_formula(f"{total_col}{r}",
                             f'=IF($A{r}="","",SUM(B{r}:{chr(ord("B") + n - 1)}{r}))',
                             F["c_money"])
        fy = data.fy_expected.get(name)
        if fy is not None:
            ws.write_number(f"{fy_col}{r}", fy, F["c_money"])
    tr = M.DASH_TOTAL_ROW
    ws.write(f"A{tr}", "TOTAL", F["total_label"])
    for i in range(n + 1):                       # class columns + Total
        col = chr(ord("B") + i)
        ws.write_formula(f"{col}{tr}", f"=SUM({col}6:{col}15)", F["total"])
    ws.write_formula(f"{fy_col}{tr}",
                     f'=IF(SUM({fy_col}6:{fy_col}15)=0,"",SUM({fy_col}6:{fy_col}15))',
                     F["total"])

    if "equity" in col_of:
        fy_now = _fy_label_today()
        ws.write("A17", f"Dividends FY {fy_now}", F["label"])
        ws.write_formula(
            "B17", f'=SUMIFS(Dividends!$I:$I,Dividends!$A:$A,"{fy_now}")',
            F["money_bold"])
        ws.write_comment("B17", "Cash your shares declared this financial year - "
                                "details and a month-by-month chart on the "
                                "Dividends tab. Estimated from your current rows.")

    ws.write("A18", "Allocation by asset class", F["section"])
    ws.write_row("A19", ["Asset class", "Value", "XIRR", "Actual %",
                         "Target %", "Drift", "Rebalance hint"], F["header"])
    ws.write_comment("E19", "Set targets on the Settings tab (Target % column). "
                            "Blank = no target for that class.")
    ws.write_comment("G19", "Indicative pre-tax amount to reach your target - "
                            "not lot-level selling advice.")
    alloc_last = 19 + n
    total_cell = f"${total_col}${M.DASH_TOTAL_ROW}"
    tol_cell = f"Settings!$B${M.SETTINGS_TOL_ROW}"
    settings_row = {c.key: M.SETTINGS_FIRST_ROW + i
                    for i, c in enumerate(ASSET_CLASSES)}
    for i, cls in enumerate(enabled):
        r = 20 + i
        ws.write(f"A{r}", cls.label)
        ws.write_formula(f"B{r}", f"={col_of[cls.key]}16", F["c_money"])
        x = getattr(data.xirr, cls.key, None) if cls.has_xirr else None
        if x is not None:
            ws.write_number(f"C{r}", x, F["u_pct"])
        else:
            ws.write_blank(f"C{r}", None, F["u_pct"])
        sr = settings_row[cls.key]
        ws.write_formula(f"D{r}",
                         f'=IF({total_cell}=0,"",B{r}/{total_cell})', F["c_pct"])
        ws.write_formula(f"E{r}",
                         f'=IF(Settings!$C${sr}="","",Settings!$C${sr}/100)',
                         F["c_pct"])
        ws.write_formula(f"F{r}", f'=IF($E{r}="","",$D{r}-$E{r})', F["c_pct"])
        ws.write_formula(
            f"G{r}",
            f'=IF($E{r}="","",IF(ABS($F{r})<={tol_cell}/100,"On target",'
            f'"Move ₹"&TEXT(ABS($F{r})*{total_cell},"#,##0")'
            f'&IF($F{r}>0," out"," in")))',
            F["c_text"])
    if n:
        ws.conditional_format(f"B20:B{alloc_last}", {
            "type": "data_bar", "bar_color": "#9DB9E3",
            "bar_solid": True, "bar_no_border": True})
        # drift band: green inside ±tolerance, red outside (blank = no target)
        ws.conditional_format(f"F20:F{alloc_last}", {
            "type": "formula",
            "criteria": f'=AND($F20<>"",ABS($F20)<={tol_cell}/100)',
            "format": F["cf_green"]})
        ws.conditional_format(f"F20:F{alloc_last}", {
            "type": "formula",
            "criteria": f'=AND($F20<>"",ABS($F20)>{tol_cell}/100)',
            "format": F["cf_red"]})
        target_chart = wb.add_chart({"type": "column"})
        target_chart.add_series({
            "name": "Actual %",
            "categories": f"=Dashboard!$A$20:$A${alloc_last}",
            "values": f"=Dashboard!$D$20:$D${alloc_last}",
        })
        target_chart.add_series({
            "name": "Target %",
            "categories": f"=Dashboard!$A$20:$A${alloc_last}",
            "values": f"=Dashboard!$E$20:$E${alloc_last}",
        })
        target_chart.set_title({"name": "Actual vs Target %"})
        ws.insert_chart("Q4", target_chart, {"x_scale": 1.1, "y_scale": 1.1})

    if n:
        pie = wb.add_chart({"type": "pie"})
        pie.add_series({
            "categories": f"=Dashboard!$A$20:$A${alloc_last}",
            "values": f"=Dashboard!$B$20:$B${alloc_last}",
            "data_labels": {"percentage": True},
        })
        pie.set_title({"name": "Allocation by asset class"})
        ws.insert_chart("I4", pie, {"x_scale": 1.1, "y_scale": 1.1})

    bar = wb.add_chart({"type": "column"})
    bar.add_series({
        "name": f"=Dashboard!${total_col}$5",
        "categories": "=Dashboard!$A$6:$A$15",
        "values": f"=Dashboard!${total_col}$6:${total_col}$15",
    })
    bar.set_title({"name": "Net worth by person"})
    bar.set_legend({"none": True})
    ws.insert_chart("I21", bar, {"x_scale": 1.1, "y_scale": 1.1})

    hist_classes = _history_classes(data)
    hist_total_col = chr(ord("B") + len(hist_classes))
    trend = wb.add_chart({"type": "line"})
    trend.add_series({
        "name": "Net worth",
        "categories": f"=History!$A$4:$A${M.HISTORY_LAST_ROW}",
        "values": f"=History!${hist_total_col}$4:${hist_total_col}${M.HISTORY_LAST_ROW}",
    })
    trend.set_title({"name": "Net worth over time"})
    trend.set_legend({"none": True})
    trend.set_x_axis({"num_format": "dd-mmm-yy"})
    ws.insert_chart("I38", trend, {"x_scale": 1.6, "y_scale": 1.2})

    if hist_classes:
        area = wb.add_chart({"type": "area", "subtype": "stacked"})
        for i, cls in enumerate(hist_classes):
            col = chr(ord("B") + i)
            area.add_series({
                "name": cls.label,
                "categories": f"=History!$A$4:$A${M.HISTORY_LAST_ROW}",
                "values": f"=History!${col}$4:${col}${M.HISTORY_LAST_ROW}",
            })
        area.set_title({"name": "Net worth by class over time"})
        area.set_x_axis({"num_format": "dd-mmm-yy"})
        ws.insert_chart("I56", area, {"x_scale": 1.6, "y_scale": 1.2})

    _redgreen(ws, F, "B4")
    _redgreen(ws, F, "E4")
    if n:
        _redgreen(ws, F, f"C20:C{alloc_last}")
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


# Per-class person-sheet holding block: headers, source sheet columns, key
# column, formats, and the red/green column range within the block (or None).
_PERSON_BLOCK_SPECS = {
    "equity": ("EQUITY",
               ["ISIN", "Scrip", "Qty", "Avg cost", "Cur. val", "Net chg.", "XIRR"],
               "Equity", ["B", "C", "O", "P", "I", "K", "N"], "Q",  # O/P = demat view
               ["c_text", "c_text", "c_text", "c_price", "c_money", "c_money", "c_pct"],
               "F:G"),
    "mutual_funds": ("MUTUAL FUNDS",
                     ["Fund House", "Scheme", "Units", "Cur NAV", "Cur. val",
                      "Net chg.", "XIRR"],
                     "MutualFunds", ["B", "C", "E", "G", "I", "J", "L"], "N",
                     ["c_text", "c_text", "c_units", "c_price", "c_money",
                      "c_money", "c_pct"],
                     "F:G"),
    "fixed_deposits": ("FIXED DEPOSITS",
                       ["Bank", "Principal", "Rate %", "Maturity", "Value today"],
                       "FixedDeposits", ["B", "D", "E", "G", "I"], "L",
                       ["c_text", "c_money", "c_text", "c_text", "c_money"],
                       None),
    "ppf": ("PPF",
            ["Institution", "Balance", "As-on"],
            "PPF", ["B", "H", "E"], "L",        # H = Balance today, L = Key
            ["c_text", "c_money", "c_text"],
            None),
    "epf": ("EPF",
            ["Establishment / UAN", "Balance today", "As-on"],
            "EPF", ["B", "H", "E"], "J",
            ["c_text", "c_money", "c_text"],
            None),
    "bonds": ("BONDS",
              ["Issuer / Bond", "Qty", "Buy Price", "Cur Price", "Cur. val",
               "Net chg."],
              "Bonds", ["B", "D", "F", "G", "K", "L"], "N",
              ["c_text", "c_text", "c_price", "c_price", "c_money", "c_money"],
              "F:F"),
    "gold_silver": ("GOLD & SILVER",
                    ["Type", "Description", "Qty (g)", "Cur. val", "Net chg."],
                    "Gold_Silver", ["B", "C", "E", "K", "M"], "O",
                    ["c_text", "c_text", "c_text", "c_money", "c_money"],
                    "E:E"),
    "nps": ("NPS",
            ["Scheme", "Units", "Cur NAV", "Cur. val"],
            "NPS", ["C", "E", "F", "G"], "K",
            ["c_text", "c_units", "c_price", "c_money"],
            None),
}


def _write_person(wb, F, name: str, data: PortfolioData):
    ws = wb.add_worksheet(name)
    _widths(ws, {"A": 26, "B": 30, "C": 12, "D": 12, "E": 14, "F": 13, "G": 10})
    ws.write("A1", f"{name} — PORTFOLIO", F["title"])
    ws.set_row(0, 18)
    ws.write("A2", "Owner", F["label"])
    ws.write("B2", name)

    enabled = enabled_classes(data)
    n = len(enabled)
    total_row = 6 + n
    ws.write("A3", "Net worth", F["label"])
    ws.write_formula("B3", f"=B{total_row}", F["money_bold"])

    ws.write_row("A5", ["Asset class", "Value", "# holdings"], F["header"])
    for i, cls in enumerate(enabled):
        r = 6 + i
        ws.write(f"A{r}", cls.label)
        ws.write_formula(f"B{r}", f"={_class_sumifs(cls, '$B$2')}", F["c_money"])
        ws.write_formula(f"C{r}", f"={_class_countifs(cls, '$B$2')}", F["c_text"])
    ws.write(f"A{total_row}", "Total", F["total_label"])
    ws.write_formula(f"B{total_row}", f"=SUM(B6:B{total_row - 1})", F["total"])
    ws.write_formula(f"C{total_row}", f"=SUM(C6:C{total_row - 1})", F["total"])

    if n:
        pie = wb.add_chart({"type": "pie"})
        pie.add_series({
            "categories": f"='{name}'!$A$6:$A${total_row - 1}",
            "values": f"='{name}'!$B$6:$B${total_row - 1}",
            "data_labels": {"percentage": True},
        })
        pie.set_title({"name": f"{name} — allocation"})
        ws.insert_chart("E4", pie)

    title_row = max(M.PERSON_BLOCKS_START, total_row + 3)
    for cls in enabled:
        spec = _PERSON_BLOCK_SPECS.get(cls.key)
        if spec is None:
            continue
        title, headers, src, cols, key_col, fmts, rg = spec
        first = title_row + 2
        last = first + cls.person_rows - 1
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
        if rg:
            c1, c2 = rg.split(":")
            _redgreen(ws, F, f"{c1}{first}:{c2}{last}")
        title_row = last + 2
    ws.freeze_panes("A5")
    return ws


def _write_equity(wb, F, data: PortfolioData):
    ws = wb.add_worksheet("Equity")
    _widths(ws, {"A": 12, "B": 15, "C": 34, "D": 10, "E": 12, "F": 13, "G": 12,
                 "H": 16, "I": 14, "J": 14, "K": 13, "L": 12, "M": 13, "N": 9,
                 "O": 11, "P": 15, "Q": 13, "R": 14, "S": 10, "T": 11})
    _sheet_head(ws, F, "EQUITY HOLDINGS",
                "Yellow-ish/blue cells are inputs. Pick the Scrip from the dropdown — "
                "ISIN fills itself. Prices refresh via the updater.")
    ws.write_row("A3", ["Owner", "ISIN", "Scrip", "Quantity", "Avg. cost",
                        "Closing Price", "Prev. close", "Closing Price Date",
                        "Cur. val", "Invested", "Net chg.", "Day chg.",
                        "Cost date", "XIRR"], F["header"])
    ws.write("O3", "Qty today", F["header"])
    ws.write("P3", "Avg cost today", F["header"])
    ws.write("Q3", "Key", F["header"])
    ws.write("R3", "Flags", F["key"])
    ws.write("S3", "Adj factor", F["header"])
    ws.write("T3", "Cost factor", F["header"])
    ws.write_comment("T3", "Demerger cost apportionment (written by the "
                           "updater): Invested = Quantity x Avg. cost x this. "
                           "Blank = 1. Your typed Avg. cost never changes; "
                           "the spun-off share of the cost moves to the new "
                           "company's row.")
    ws.write_comment("D3", "As bought. After a split/bonus, 'Qty today' shows the "
                           "post-action share count - matching your demat.")
    ws.write_comment("E3", "As bought (per share). Leave blank for pre-Feb-2018 "
                           "purchases whose price you don't know - the updater "
                           "fills the 31-Jan-2018 FMV (LTCG grandfathering value) "
                           "and marks the cell amber.")
    ws.write_comment("O3", "Your holding after splits/bonuses (Quantity x Adj "
                           "factor) - should match your demat. Updates "
                           "automatically from the Corporate_Actions sheet.")
    ws.write_comment("P3", "Cost per share today (Avg. cost / Adj factor) - "
                           "comparable to your broker app after splits/bonuses.")
    ws.write_comment("S3", "Split/bonus multiplier since your Cost date, from the "
                           "Corporate_Actions sheet (written by the updater). "
                           "Valuation uses Quantity x Adj factor; your typed "
                           "Quantity and Avg. cost stay as-purchased.")

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
                if row.fmv_used:
                    ws.write_number(f"E{r}", row.avg_cost, F["amber_price"])
                    ws.write_comment(f"E{r}", "Cost unknown - using the 31-01-2018 "
                                              "FMV (LTCG grandfathering value). "
                                              "Overtype with the real cost if you "
                                              "find it.")
                else:
                    ws.write_number(f"E{r}", row.avg_cost, F["in_price"])
            flag_text = row.flag or ("FMV" if row.fmv_used else "")
            if flag_text:
                ws.write(f"R{r}", flag_text, F["key"])
            if row.close is not None:
                ws.write_number(f"F{r}", row.close, F["u_price"])
            if row.prev_close is not None:
                ws.write_number(f"G{r}", row.prev_close, F["u_price"])
            if row.close_date:
                ws.write_datetime(f"H{r}", row.close_date, F["date_disp"])
            if row.cost_date is not None:
                ws.write_datetime(f"M{r}", row.cost_date, F["in_date"])
        if row and row.ca_factor is not None:
            ws.write_number(f"S{r}", row.ca_factor, F["c_units"])
        if row and row.cost_factor is not None:
            ws.write_number(f"T{r}", row.cost_factor, F["c_units"])
        ws.write_formula(f"I{r}",
                         f'=IF($D{r}="","",$D{r}*IF($S{r}="",1,$S{r})*$F{r})',
                         F["c_money"])
        ws.write_formula(f"J{r}",
                         f'=IF(OR($D{r}="",$E{r}=""),"",'
                         f'$D{r}*$E{r}*IF($T{r}="",1,$T{r}))', F["c_money"])
        ws.write_formula(f"K{r}", f'=IF(OR($E{r}="",$D{r}=""),"",$I{r}-$J{r})', F["c_money"])
        ws.write_formula(f"L{r}",
                         f'=IF(OR($G{r}="",$D{r}=""),"",'
                         f'$D{r}*IF($S{r}="",1,$S{r})*($F{r}-$G{r}))',
                         F["c_money"])
        ws.write_formula(f"O{r}",
                         f'=IF($D{r}="","",$D{r}*IF($S{r}="",1,$S{r}))',
                         F["c_units"])
        ws.write_formula(f"P{r}",
                         f'=IF(OR($D{r}="",$E{r}=""),"",$E{r}/IF($S{r}="",1,$S{r}))',
                         F["c_price"])
        ws.write_formula(
            f"N{r}",
            f'=IF(OR($M{r}="",N($J{r})=0,$I{r}="",TODAY()<=$M{r}),"",'
            f'($I{r}/$J{r})^(365/(TODAY()-$M{r}))-1)', F["c_pct"])
        ws.write_formula(f"Q{r}", _key_formula(r), F["key"])

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
    # ▲/▼ arrows on Day chg. — up above zero, down below (absolute thresholds)
    ws.conditional_format(f"L4:L{M.EQUITY_LAST_ROW}", {
        "type": "icon_set", "icon_style": "3_arrows",
        "icons": [{"criteria": ">", "type": "number", "value": 0},
                  {"criteria": ">=", "type": "number", "value": 0}]})
    _redgreen(ws, F, f"K{tr}:L{tr}")
    _redgreen(ws, F, f"N{tr}")
    # amber flags (SPEC §6.5): stale price date; suspended/delisted scrip
    ws.conditional_format(f"F4:H{M.EQUITY_LAST_ROW}", {
        "type": "formula",
        "criteria": '=AND($H4<>"",TODAY()-$H4>7)',
        "format": F["cf_amber"]})
    ws.conditional_format(f"C4:C{M.EQUITY_LAST_ROW}", {
        "type": "formula",
        "criteria": ('=OR(IFERROR(INDEX(Stock_Master!$D:$D,MATCH($B4,'
                     'Stock_Master!$C:$C,0)),"")="Suspended",'
                     'IFERROR(INDEX(Stock_Master!$D:$D,MATCH($B4,'
                     'Stock_Master!$C:$C,0)),"")="Delisted")'),
        "format": F["cf_amber"]})
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
                  headers: list[str], rows: list[tuple], refreshed: str,
                  status: dict | None = None):
    ws = wb.add_worksheet(name)
    _widths(ws, {"A": 34, "B": 60, "C": 16, "D": 12, "E": 12})
    ws.write("A1", title, F["title"])
    ws.set_row(0, 18)
    ws.write("A2", hint, F["hint"])
    ws.write("D2", "Refreshed:", F["label"])
    ws.write("E2", refreshed, F["in_text"])
    ws.write_row("A3", headers, F["header"])
    if status is not None:
        ws.write("D3", "Status", F["header"])
        ws.write("E3", "Last Traded", F["header"])
    r = M.FIRST_DATA_ROW - 1  # 0-based row index for write_row
    for tup in rows:
        ws.write_row(r, 0, tup)
        if status and tup[2] in status:
            st, last = status[tup[2]]
            ws.write(r, 3, st)
            if last:
                ws.write_datetime(r, 4, last, F["date_disp"])
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
    _widths(ws, {"A": 12, "B": 18, "C": 16, "D": 15, "E": 13, "F": 11, "G": 20,
                 "H": 14, "I": 14, "J": 9, "L": 13})
    _sheet_head(ws, F, "PPF ACCOUNTS",
                "Enter Owner, Institution, Account no. For accuracy, log deposits on "
                "the PPF_Ledger sheet - then Balance today, Interest and XIRR are "
                "computed exactly. No ledger? Just type a Current Balance and it is "
                "used as-is.")
    ws.write_row("A3", ["Owner", "Institution", "Account No.", "Current Balance",
                        "Balance as-on", "Rate % (ref)", "Notes", "Balance today",
                        "Interest earned", "XIRR"], F["header"])
    ws.write("L3", "Key", F["header"])
    ws.write_comment("D3", "Only used when this account has NO rows on PPF_Ledger. "
                           "With a ledger, Balance today is computed from the deposits.")
    ws.write_comment("F3", "Auto-filled with the current PPF rate by the updater if "
                           "left blank; overtype to pin a specific rate.")
    ws.write_comment("H3", "With a PPF_Ledger: deposits + interest to date (official "
                           "monthly-minimum-balance rule). Without: your Current "
                           "Balance. This is what the Dashboard totals.")
    ws.write_comment("I3", "Total interest earned to date (ledger accounts only), "
                           "written by the updater.")

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
            if row.balance_today is not None:      # ledger account, updater-computed
                ws.write_number(f"H{r}", row.balance_today, F["c_money"])
            if row.interest_earned is not None:
                ws.write_number(f"I{r}", row.interest_earned, F["c_money"])
            if row.xirr is not None:
                ws.write_number(f"J{r}", row.xirr, F["u_pct"])
        if not (row and row.balance_today is not None):
            # no ledger (or not yet updated) → Balance today = Current Balance
            ws.write_formula(f"H{r}", f'=IF($D{r}="","",$D{r})', F["c_money"])
        ws.write_formula(f"L{r}", _key_formula(r), F["key"])

    tr = M.PPF_TOTAL_ROW
    ws.write(f"C{tr}", "TOTAL", F["total_label"])
    ws.write_formula(f"H{tr}", f"=SUM(H4:H{M.PPF_LAST_ROW})", F["total"])
    _redgreen(ws, F, f"J4:J{M.PPF_LAST_ROW}")
    ws.freeze_panes("A4")
    return ws


def _write_ppf_ledger(wb, F, data: PortfolioData):
    ws = wb.add_worksheet("PPF_Ledger")
    _widths(ws, {"A": 12, "B": 18, "C": 13, "D": 14})
    _sheet_head(ws, F, "PPF DEPOSIT LEDGER (optional)",
                "One row per PPF deposit. Match Owner + Account No. to a row on the "
                "PPF sheet. Fill this in and the updater computes exact balance, "
                "interest and XIRR (monthly-minimum-balance rule). Leave empty to "
                "keep using the Current Balance you typed on the PPF sheet.")
    ws.write_row("A3", ["Owner", "Account No.", "Date", "Amount"], F["header"])
    by_row = {M.FIRST_DATA_ROW + i: row for i, row in enumerate(data.ppf_ledger)}
    for r in range(M.FIRST_DATA_ROW, M.PPF_LEDGER_LAST_ROW + 1):
        row = by_row.get(r)
        if row:
            ws.write(f"A{r}", row.owner, F["in_text"])
            ws.write(f"B{r}", row.account_no, F["in_text"])
            if row.txn_date is not None:
                ws.write_datetime(f"C{r}", row.txn_date, F["in_date"])
            if row.amount is not None:
                ws.write_number(f"D{r}", row.amount, F["in_money"])
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


def _write_gold_silver(wb, F, data: PortfolioData):
    ws = wb.add_worksheet("Gold_Silver")
    _widths(ws, {"A": 12, "B": 8, "C": 26, "D": 14, "E": 12, "F": 8, "G": 13,
                 "H": 12, "I": 14, "J": 13, "K": 14, "L": 13, "M": 13,
                 "N": 12, "O": 12})
    _sheet_head(ws, F, "GOLD & SILVER",
                "Type grams (and purity for jewellery, e.g. 0.916 for 22K) - "
                "today's value appears at the daily bullion rate. Prefer your "
                "jeweller's rate? Type it in Rate override - it wins. SGBs: "
                "just fill the ISIN, they price like shares.")
    ws.write_row("A3", ["Owner", "Type", "Description / Series", "ISIN",
                        "Qty (g / units)", "Purity", "Buy Price ₹/unit",
                        "Buy Date", "Rate today (auto)", "Rate override",
                        "Cur. val", "Invested", "Net chg.", "Maturity"],
                 F["header"])
    ws.write("O3", "Key", F["key"])
    ws.write("H2", "Rates as on", F["label"])
    if data.bullion_rate_asof:
        ws.write_datetime("I2", data.bullion_rate_asof, F["date_disp"])
    ws.write_comment("I3", "Filled by the updater: SGBs get their exchange "
                           "close; gold/silver get the daily benchmark ₹/gram "
                           "(IBJA; market-implied fallback). Amber when the "
                           "rate is over a week old.")
    ws.write_comment("J3", "Your own ₹/gram (e.g. the jeweller's board rate). "
                           "When set, it always wins over the auto rate.")
    ws.write_comment("F3", "Fine-metal fraction: 24K/SGB = blank or 1, "
                           "22K = 0.916, 18K = 0.75.")
    by_row = {M.FIRST_DATA_ROW + i: b for i, b in enumerate(data.bullion)}
    for r in range(M.FIRST_DATA_ROW, M.GS_LAST_ROW + 1):
        b = by_row.get(r)
        if b:
            ws.write(f"A{r}", b.owner, F["in_text"])
            ws.write(f"B{r}", b.metal_type, F["in_text"])
            ws.write(f"C{r}", b.description, F["in_text"])
            ws.write(f"D{r}", b.isin, F["in_text"])
            if b.qty is not None:
                ws.write_number(f"E{r}", b.qty, F["in_num"])
            if b.purity is not None:
                ws.write_number(f"F{r}", b.purity, F["in_num"])
            if b.buy_price is not None:
                ws.write_number(f"G{r}", b.buy_price, F["in_price"])
            if b.buy_date:
                ws.write_datetime(f"H{r}", b.buy_date, F["in_date"])
            if b.rate_auto is not None:
                ws.write_number(f"I{r}", b.rate_auto, F["u_price"])
            if b.rate_override is not None:
                ws.write_number(f"J{r}", b.rate_override, F["in_price"])
            if b.maturity:
                ws.write_datetime(f"N{r}", b.maturity, F["in_date"])
        ws.write_formula(
            f"K{r}",
            f'=IF(OR($E{r}="",AND($I{r}="",$J{r}="")),"",'
            f'$E{r}*IF($F{r}="",1,$F{r})*IF($J{r}="",$I{r},$J{r}))',
            F["c_money"])
        ws.write_formula(f"L{r}",
                         f'=IF(OR($E{r}="",$G{r}=""),"",$E{r}*$G{r})',
                         F["c_money"])
        ws.write_formula(f"M{r}",
                         f'=IF(OR($K{r}="",$L{r}=""),"",$K{r}-$L{r})',
                         F["c_money"])
        ws.write_formula(f"O{r}", _key_formula(r), F["key"])
    tr = M.GS_LAST_ROW + 2
    ws.write(f"C{tr}", "TOTAL", F["total_label"])
    for col in "KL":
        ws.write_formula(f"{col}{tr}", f"=SUM({col}4:{col}{M.GS_LAST_ROW})",
                         F["total"])
    ws.data_validation(f"B4:B{M.GS_LAST_ROW}", {
        "validate": "list", "source": ["SGB", "Gold", "Silver"],
        "show_error": False, "input_title": "Type",
        "input_message": "SGB (fill the ISIN too) / Gold / Silver"})
    _redgreen(ws, F, f"M4:M{M.GS_LAST_ROW}")
    ws.conditional_format(f"I4:I{M.GS_LAST_ROW}", {
        "type": "formula",
        "criteria": '=AND($I4<>"",$I$2<>"",TODAY()-$I$2>7)',
        "format": F["cf_amber"]})
    ws.freeze_panes("A4")
    return ws


def _write_nps(wb, F, data: PortfolioData):
    ws = wb.add_worksheet("NPS")
    _widths(ws, {"A": 12, "B": 15, "C": 44, "D": 12, "E": 12, "F": 12,
                 "G": 14, "H": 16, "I": 15, "J": 10, "K": 12})
    _sheet_head(ws, F, "NPS — NATIONAL PENSION SYSTEM",
                "Pick the scheme from the dropdown and type your units (both "
                "are on your CRA statement) - the NAV and value fill in on "
                "update. Add what you've contributed so far for a return "
                "figure.")
    ws.write_row("A3", ["Owner", "PRAN", "Scheme", "Scheme Code", "Units",
                        "Current NAV", "Cur. val", "Total contributed",
                        "First contribution", "XIRR"], F["header"])
    ws.write("K3", "Key", F["key"])
    ws.write_comment("J3", "Approximate: one flow in (Total contributed at "
                           "First contribution) vs today's value. A dated "
                           "contribution ledger is on the roadmap.")
    by_row = {M.FIRST_DATA_ROW + i: n for i, n in enumerate(data.nps)}
    for r in range(M.FIRST_DATA_ROW, M.NPS_LAST_ROW + 1):
        n = by_row.get(r)
        if n:
            ws.write(f"A{r}", n.owner, F["in_text"])
            ws.write(f"B{r}", n.pran, F["in_text"])
            ws.write(f"C{r}", n.scheme, F["in_text"])
            if n.scheme_code_override:
                ws.write(f"D{r}", n.scheme_code_override, F["in_text"])
            if n.units is not None:
                ws.write_number(f"E{r}", n.units, F["in_num"])
            if n.current_nav is not None:
                ws.write_number(f"F{r}", n.current_nav, F["u_price"])
            if n.total_contributed is not None:
                ws.write_number(f"H{r}", n.total_contributed, F["in_money"])
            if n.first_contribution:
                ws.write_datetime(f"I{r}", n.first_contribution, F["in_date"])
            if n.xirr is not None:
                ws.write_number(f"J{r}", n.xirr, F["u_pct"])
        if not (n and n.scheme_code_override):
            ws.write_formula(
                f"D{r}",
                f'=IF($C{r}="","",IFERROR(INDEX(NPS_Master!$A:$A,'
                f'MATCH($C{r},NPS_Master!$B:$B,0)),""))', F["c_text"])
        ws.write_formula(f"G{r}",
                         f'=IF(OR($E{r}="",$F{r}=""),"",$E{r}*$F{r})',
                         F["c_money"])
        ws.write_formula(f"K{r}", _key_formula(r), F["key"])
    tr = M.NPS_LAST_ROW + 2
    ws.write(f"C{tr}", "TOTAL", F["total_label"])
    ws.write_formula(f"G{tr}", f"=SUM(G4:G{M.NPS_LAST_ROW})", F["total"])
    _add_dropdown(ws, f"C4:C{M.NPS_LAST_ROW}",
                  _typeahead("NPS_Master", "NPS_SchemeList"), "NPS scheme")
    _redgreen(ws, F, f"J4:J{M.NPS_LAST_ROW}")
    ws.freeze_panes("A4")
    return ws


def _write_epf(wb, F, data: PortfolioData):
    ws = wb.add_worksheet("EPF")
    _widths(ws, {"A": 12, "B": 24, "C": 16, "D": 15, "E": 12, "F": 10,
                 "G": 24, "H": 15, "I": 2, "J": 12})
    _sheet_head(ws, F, "EPF — EMPLOYEES' PROVIDENT FUND",
                "Copy the balance and its date from your EPFO passbook - "
                "that's all. It grows at the current EPF rate (filled in for "
                "you) until you paste a newer balance.")
    ws.write_row("A3", ["Owner", "Establishment / UAN", "Member ID",
                        "Current Balance", "Balance as-on", "Rate %", "Notes",
                        "Balance today"], F["header"])
    ws.write("J3", "Key", F["key"])
    ws.write_comment("D3", "The closing balance shown on your EPFO passbook "
                           "(employee + employer share).")
    ws.write_comment("F3", "The EPFO annual rate. Left blank, the updater "
                           "fills the current declared rate.")
    ws.write_comment("H3", "Passbook balance grown at Rate % from the as-on "
                           "date - an estimate until a contribution ledger "
                           "lands (roadmap). Exact figure: your passbook.")
    by_row = {M.FIRST_DATA_ROW + i: e for i, e in enumerate(data.epf)}
    for r in range(M.FIRST_DATA_ROW, M.EPF_LAST_ROW + 1):
        e = by_row.get(r)
        if e:
            ws.write(f"A{r}", e.owner, F["in_text"])
            ws.write(f"B{r}", e.establishment, F["in_text"])
            ws.write(f"C{r}", e.member_id, F["in_text"])
            if e.balance is not None:
                ws.write_number(f"D{r}", e.balance, F["in_money"])
            if e.as_on:
                ws.write_datetime(f"E{r}", e.as_on, F["in_date"])
            if e.rate is not None:
                ws.write_number(f"F{r}", e.rate, F["in_rate"])
            if e.notes:
                ws.write(f"G{r}", e.notes, F["in_text"])
        ws.write_formula(
            f"H{r}",
            f'=IF($D{r}="","",IF(OR($E{r}="",$F{r}=""),$D{r},'
            f'$D{r}*(1+$F{r}/100)^YEARFRAC($E{r},TODAY())))',
            F["c_money"])
        ws.write_formula(f"J{r}", _key_formula(r), F["key"])
    tr = M.EPF_LAST_ROW + 2
    ws.write(f"C{tr}", "TOTAL", F["total_label"])
    ws.write_formula(f"H{tr}", f"=SUM(H4:H{M.EPF_LAST_ROW})", F["total"])
    ws.freeze_panes("A4")
    return ws


def _write_manual_assets(wb, F, data: PortfolioData):
    ws = wb.add_worksheet("Manual_Assets")
    _widths(ws, {"A": 12, "B": 13, "C": 28, "D": 20, "E": 14, "F": 13,
                 "G": 16, "H": 13, "I": 14, "J": 30, "K": 12})
    _sheet_head(ws, F, "OTHER ASSETS — VALUED BY YOU",
                "The house, savings accounts, insurance surrender values, "
                "anything else. Type today's value and its date - refresh it "
                "now and then (the date turns amber after 90 days).")
    ws.write_row("A3", ["Owner", "Class", "Description", "Institution / Ref",
                        "Invested / Cost", "Cost date", "Current value ₹",
                        "Value as-on", "Net chg.", "Notes"], F["header"])
    ws.write("K3", "Key", F["key"])
    ws.write_comment("B3", "Real Estate / Cash / Insurance / Other - each is "
                           "its own line on the Dashboard.")
    ws.write_comment("E3", "Optional. Real estate: purchase cost. Insurance: "
                           "premiums paid so far. Enables Net chg. and a "
                           "return figure.")
    ws.write_comment("G3", "What it is worth today: market estimate, account "
                           "balance, or the insurer's surrender value.")
    ws.write_comment("H3", "When you last checked the value. Amber after 90 "
                           "days - hand-typed values rot silently.")
    by_row = {M.FIRST_DATA_ROW + i: a for i, a in enumerate(data.manual_assets)}
    for r in range(M.FIRST_DATA_ROW, M.MA_LAST_ROW + 1):
        a = by_row.get(r)
        if a:
            ws.write(f"A{r}", a.owner, F["in_text"])
            ws.write(f"B{r}", a.asset_class, F["in_text"])
            ws.write(f"C{r}", a.description, F["in_text"])
            ws.write(f"D{r}", a.institution, F["in_text"])
            if a.invested is not None:
                ws.write_number(f"E{r}", a.invested, F["in_money"])
            if a.cost_date:
                ws.write_datetime(f"F{r}", a.cost_date, F["in_date"])
            if a.value is not None:
                ws.write_number(f"G{r}", a.value, F["in_money"])
            if a.as_on:
                ws.write_datetime(f"H{r}", a.as_on, F["in_date"])
            if a.notes:
                ws.write(f"J{r}", a.notes, F["in_text"])
        ws.write_formula(f"I{r}",
                         f'=IF(OR($E{r}="",$G{r}=""),"",$G{r}-$E{r})',
                         F["c_money"])
        ws.write_formula(f"K{r}", _key_formula(r), F["key"])
    tr = M.MA_LAST_ROW + 2
    ws.write(f"C{tr}", "TOTAL", F["total_label"])
    ws.write_formula(f"G{tr}", f"=SUM(G4:G{M.MA_LAST_ROW})", F["total"])
    ws.data_validation(f"B4:B{M.MA_LAST_ROW}", {
        "validate": "list",
        "source": M.MANUAL_CLASS_LABELS,
        "show_error": False,
        "input_title": "Asset class",
        "input_message": "Real Estate / Cash / Insurance / Other",
    })
    _redgreen(ws, F, f"I4:I{M.MA_LAST_ROW}")
    ws.conditional_format(f"H4:H{M.MA_LAST_ROW}", {
        "type": "formula",
        "criteria": f'=AND($H4<>"",TODAY()-$H4>90)',
        "format": F["cf_amber"]})
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
                         f'=IF($A{r}="","",SUMIF(Equity!$B:$B,$A{r},Equity!$O:$O))',
                         F["c_text"])
        for col, person in zip(person_cols, data.persons):
            ws.write_formula(f"{col}{r}",
                             f'=IF($A{r}="","",SUMIFS(Equity!$O:$O,Equity!$B:$B,$A{r},'
                             f'Equity!$A:$A,"{person}"))', F["c_text"])
        ws.write_formula(f"{curval_col}{r}",
                         f'=IF($A{r}="","",SUMIF(Equity!$B:$B,$A{r},Equity!$I:$I))',
                         F["c_money"])
    ws.freeze_panes("A4")
    return ws


def _write_corporate_actions(wb, F, data: PortfolioData):
    ws = wb.add_worksheet("Corporate_Actions")
    _widths(ws, {"A": 14, "B": 16, "C": 16, "D": 12, "E": 11, "F": 11, "G": 9,
                 "H": 9, "I": 46, "J": 16, "K": 9, "L": 12})
    _sheet_head(ws, F, "CORPORATE ACTIONS (SPLITS / BONUSES / RESTRUCTURES)",
                "Auto rows are fetched for your held stocks on every update; "
                "Curated rows (mergers/demergers) ship with each release. Add "
                "Manual rows for anything missed. Your Equity rows are never "
                "rewritten - factors apply at valuation time.")
    ws.write_row("A3", ["Symbol", "ISIN", "Type", "Ex-Date", "Ratio From",
                        "Ratio To", "Factor", "Source", "Details", "New ISIN",
                        "Cost %", "Applied"], F["header"])
    ws.write_comment("E3", "SPLIT/CONSOLIDATION: old face value. BONUS/MERGER/"
                           "DEMERGER: A of A:B (A new shares per B held).")
    ws.write_comment("F3", "SPLIT/CONSOLIDATION: new face value. BONUS/MERGER/"
                           "DEMERGER: B of A:B.")
    ws.write_comment("J3", "MERGER/ISIN_CHANGE: the surviving security's ISIN. "
                           "DEMERGER: the resulting security of this row (one "
                           "row per outcome; the parent-retention row repeats "
                           "the old ISIN).")
    ws.write_comment("K3", "Cost-basis share in percent. A demerger's rows "
                           "must sum to 100 across the event.")
    ws.write_comment("L3", "When the updater applied this event (wrote child "
                           "rows / cost factors). Managed automatically - a "
                           "demerger is applied ONCE, so deleting a spun-off "
                           "row later is respected.")

    by_row = {M.FIRST_DATA_ROW + i: a for i, a in enumerate(data.corporate_actions)}
    for r in range(M.FIRST_DATA_ROW, M.CA_LAST_ROW + 1):
        a = by_row.get(r)
        if a:
            fmt = F["c_text"] if a.source in ("Auto", "Curated") else F["in_text"]
            dfmt = (F["date_disp"] if a.source in ("Auto", "Curated")
                    else F["in_date"])
            ws.write(f"A{r}", a.symbol, fmt)
            ws.write(f"B{r}", a.isin, fmt)
            ws.write(f"C{r}", a.type, fmt)
            if a.ex_date:
                ws.write_datetime(f"D{r}", a.ex_date, dfmt)
            if a.ratio_from is not None:
                ws.write_number(f"E{r}", a.ratio_from, fmt)
            if a.ratio_to is not None:
                ws.write_number(f"F{r}", a.ratio_to, fmt)
            ws.write(f"H{r}", a.source, F["c_text"])
            if a.details:
                ws.write(f"I{r}", a.details, fmt)
            if a.new_isin:
                ws.write(f"J{r}", a.new_isin, fmt)
            if a.cost_pct is not None:
                ws.write_number(f"K{r}", a.cost_pct, fmt)
            if a.applied:
                ws.write_datetime(f"L{r}", a.applied, F["date_disp"])
        ws.write_formula(f"G{r}",
                         f'=IF(OR($C{r}="",$E{r}="",$F{r}=""),"",'
                         f'IF($C{r}="BONUS",1+$E{r}/$F{r},'
                         f'IF(OR($C{r}="DEMERGER",$C{r}="ISIN_CHANGE"),1,'
                         f'$E{r}/$F{r})))',
                         F["c_units"])
    ws.data_validation(f"C4:C{M.CA_LAST_ROW}", {
        "validate": "list",
        "source": ["SPLIT", "BONUS", "CONSOLIDATION", "MERGER", "DEMERGER",
                   "ISIN_CHANGE"],
        "show_error": False,
        "input_title": "Action type",
        "input_message": "SPLIT / BONUS / CONSOLIDATION / MERGER / DEMERGER / "
                         "ISIN_CHANGE",
    })
    ws.freeze_panes("A4")
    return ws


def _fy_label_today() -> str:
    from .model import fy_label
    from datetime import date as _date
    return fy_label(_date.today())


_FY_MONTHS = [4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3]
_FY_MONTH_NAMES = ["Apr", "May", "Jun", "Jul", "Aug", "Sep",
                   "Oct", "Nov", "Dec", "Jan", "Feb", "Mar"]


def _write_dividends(wb, F, data: PortfolioData):
    ws = wb.add_worksheet("Dividends")
    _widths(ws, {"A": 9, "B": 10, "C": 30, "D": 14, "E": 9, "F": 12, "G": 11,
                 "H": 14, "I": 12, "J": 8, "K": 46, "L": 2, "M": 7, "N": 12})
    fy_now = _fy_label_today()
    _sheet_head(ws, F, "DIVIDENDS — YOUR SHARES' CASH INCOME",
                "Filled in for you on every update: one row per dividend your "
                "stocks declared this financial year. Older years stay as a "
                "record. Qty and amount are estimates from your current rows "
                "(amber) - add a Manual row for anything the feed missed.")
    ws.write_row("A3", ["FY", "Owner", "Scrip", "ISIN", "Type", "Ex-Date",
                        "Rate ₹/share", "Qty @ ex-date (est.)", "Est. amount",
                        "Source", "Details"], F["header"])
    ws.write_comment("H3", "Estimated shares held on the ex-date: your rows "
                           "bought before that date, adjusted for splits/bonuses. "
                           "If you sold in between, correct it here by hand.")
    ws.write_comment("I3", "Rate x Qty - an estimate, marked amber. The exact "
                           "credit is on your bank statement.")

    by_row = {M.FIRST_DATA_ROW + i: d for i, d in enumerate(data.dividends)}
    for r in range(M.FIRST_DATA_ROW, M.DIV_LAST_ROW + 1):
        d = by_row.get(r)
        if d:
            auto = d.source == "Auto"
            fmt = F["c_text"] if auto else F["in_text"]
            ws.write(f"A{r}", d.fy, fmt)
            ws.write(f"B{r}", d.owner, fmt)
            ws.write(f"C{r}", d.scrip, fmt)
            ws.write(f"D{r}", d.isin, fmt)
            ws.write(f"E{r}", d.div_type, fmt)
            if d.ex_date:
                ws.write_datetime(f"F{r}", d.ex_date,
                                  F["date_disp"] if auto else F["in_date"])
            if d.rate is not None:
                ws.write_number(f"G{r}", d.rate,
                                F["u_price"] if auto else F["in_price"])
            if d.qty is not None:
                ws.write_number(f"H{r}", d.qty, F["amber_price"])
            ws.write(f"J{r}", d.source, F["c_text"])
            if d.details:
                ws.write(f"K{r}", d.details, fmt)
        ws.write_formula(f"I{r}",
                         f'=IF(OR($G{r}="",$H{r}=""),"",$G{r}*$H{r})',
                         F["amber_price"])
    ws.data_validation(f"E4:E{M.DIV_LAST_ROW}", {
        "validate": "list",
        "source": ["Interim", "Final", "Special"],
        "show_error": False,
        "input_title": "Dividend type",
        "input_message": "Interim / Final / Special",
    })

    # "Dividends by month" — income the user can SEE arriving
    ws.write("M3", "", F["header"])
    ws.write("N3", f"FY {fy_now} ₹", F["header"])
    for i, (m, label) in enumerate(zip(_FY_MONTHS, _FY_MONTH_NAMES)):
        r = 4 + i
        ws.write(f"M{r}", label, F["c_text"])
        ws.write_formula(
            f"N{r}",
            f'=SUMPRODUCT(($G$4:$G${M.DIV_LAST_ROW})*($H$4:$H${M.DIV_LAST_ROW})'
            f'*(MONTH($F$4:$F${M.DIV_LAST_ROW})={m})'
            f'*($A$4:$A${M.DIV_LAST_ROW}="{fy_now}"))',
            F["c_money"])
    chart = wb.add_chart({"type": "column"})
    chart.add_series({
        "name": f"Dividends by month — FY {fy_now}",
        "categories": "=Dividends!$M$4:$M$15",
        "values": "=Dividends!$N$4:$N$15",
    })
    chart.set_title({"name": f"Dividends by month — FY {fy_now}"})
    chart.set_legend({"none": True})
    ws.insert_chart("M18", chart, {"x_scale": 1.2, "y_scale": 1.1})
    ws.freeze_panes("A4")
    return ws


def _write_history(wb, F, data: PortfolioData):
    ws = wb.add_worksheet("History")
    classes = _history_classes(data)
    n = len(classes)
    _widths(ws, {"A": 13, **{chr(ord("B") + i): 14 for i in range(n + 1)}})
    _sheet_head(ws, F, "NET WORTH HISTORY",
                "One snapshot per day, written by the updater (the same run that "
                "refreshes prices). The Dashboard trend chart reads this sheet. "
                "Machine-managed - no need to edit it.")
    ws.write_row("A3", ["Date"] + [c.label for c in classes] + ["Total"],
                 F["header"])
    last_cls_col = chr(ord("B") + n - 1) if n else "B"
    total_col = chr(ord("B") + n)
    by_row = {M.FIRST_DATA_ROW + i: s for i, s in enumerate(data.history)}
    for r in range(M.FIRST_DATA_ROW, M.HISTORY_LAST_ROW + 1):
        snap = by_row.get(r)
        if snap and snap.snap_date is not None:
            ws.write_datetime(f"A{r}", snap.snap_date, F["date_disp"])
            for i, cls in enumerate(classes):
                ws.write_number(f"{chr(ord('B') + i)}{r}",
                                getattr(snap, cls.key, 0.0), F["c_money"])
            if n:
                ws.write_formula(f"{total_col}{r}",
                                 f"=SUM(B{r}:{last_cls_col}{r})", F["c_money"])
    ws.freeze_panes("A4")
    return ws


def _write_guide(wb, F):
    ws = wb.add_worksheet("Guide")
    ws.hide_gridlines(2)                       # page-like look
    ws.set_column("A:A", 2)                    # left margin
    ws.set_column("B:B", 4)                    # icon / badge / swatch
    ws.set_column("C:C", 96)                   # content

    def fmt(**kw):
        return wb.add_format(kw)

    NAVY = "#1F3864"
    PALETTE = ["#2E5A9C", "#2E7D5B", "#9C5A2E", "#6A4C93", "#1F6F78", "#B5522E", "#4C6A2E"]
    f_title = fmt(bold=True, font_size=18, font_color="white", bg_color=NAVY,
                  valign="vcenter", indent=1)
    f_sub = fmt(font_size=11, italic=True, font_color="#D6E0F0", bg_color=NAVY,
                valign="vcenter", indent=1)
    f_body = fmt(font_size=11, font_color="#333333", valign="vcenter")
    f_tip = fmt(font_size=10.5, italic=True, font_color="#6A6A6A", valign="vcenter")
    f_footer = fmt(font_size=10.5, font_color="#8A8A8A", top=1, top_color="#D0D0D0",
                   valign="vcenter")
    f_badge = fmt(bold=True, font_color="white", bg_color="#2E5A9C",
                  align="center", valign="vcenter")
    sw_yellow = fmt(bg_color="#FFF2CC", border=1, border_color="#D6C089")
    sw_grey = fmt(bg_color="#F2F2F2", border=1, border_color="#C9C9C9")
    # font-only fragments for rich strings
    rf_body = fmt(font_size=11, font_color="#333333")
    rf_key = fmt(bold=True, font_size=11, font_color=NAVY)
    rf_bul = fmt(bold=True, font_size=11, font_color="#2E5A9C")
    rf_green = fmt(bold=True, font_size=11, font_color="#1F7A1F")
    rf_red = fmt(bold=True, font_size=11, font_color="#C00000")
    rf_amber = fmt(bold=True, font_size=11, font_color="#B8860B")
    band_cache: dict[str, object] = {}

    def band(color):
        if color not in band_cache:
            band_cache[color] = fmt(bold=True, font_size=12, font_color="white",
                                    bg_color=color, valign="vcenter", indent=1)
        return band_cache[color]

    r, sec = 0, 0
    for row in GUIDE_ROWS:
        kind = row[0]
        if kind == "title":
            ws.merge_range(r, 0, r, 2, row[1], f_title); ws.set_row(r, 34); r += 1
            ws.merge_range(r, 0, r, 2, row[2], f_sub); ws.set_row(r, 22); r += 1
        elif kind == "section":
            ws.merge_range(r, 0, r, 2, f"{row[1]}   {row[2]}",
                           band(PALETTE[sec % len(PALETTE)])); sec += 1
            ws.set_row(r, 24); r += 1
        elif kind == "legend":
            ws.write_blank(r, 1, None, sw_yellow)
            ws.write(r, 2, "You type into the blue / yellow cells — only these.", f_body)
            ws.set_row(r, 17); r += 1
            ws.write_blank(r, 1, None, sw_grey)
            ws.write(r, 2, "Grey cells are worked out for you — leave them alone.", f_body)
            ws.set_row(r, 17); r += 1
            ws.write_rich_string(r, 2, rf_green, "Green", rf_body, " = gain      ",
                                 rf_red, "Red", rf_body, " = loss      ",
                                 rf_amber, "Amber", rf_body, " = take a look")
            ws.set_row(r, 17); r += 1
        elif kind == "step":
            ws.write(r, 1, row[1], f_badge)
            ws.write(r, 2, row[2], f_body); ws.set_row(r, 20); r += 1
        elif kind == "kv":
            ws.write_rich_string(r, 2, rf_key, row[1] + "   ", rf_body, "— " + row[2])
            ws.set_row(r, 16); r += 1
        elif kind == "bullet":
            ws.write_rich_string(r, 2, rf_bul, "•   ", rf_body, row[1])
            ws.set_row(r, 16); r += 1
        elif kind == "tip":
            ws.write(r, 2, "Tip ·  " + row[1], f_tip); ws.set_row(r, 16); r += 1
        elif kind == "text":
            ws.write(r, 2, row[1], f_body); ws.set_row(r, 16); r += 1
        elif kind == "footer":
            ws.write(r, 2, row[1], f_footer); ws.set_row(r, 24); r += 1
        elif kind == "space":
            ws.set_row(r, 7); r += 1
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
    wb.define_name(
        "NPS_SchemeList",
        "=NPS_Master!$B$4:INDEX(NPS_Master!$B:$B,COUNTA(NPS_Master!$B:$B)+2)")

    # tab order = SPEC §3.1
    _write_dashboard(wb, F, data)
    _write_projection(wb, F)
    _write_settings(wb, F, data)
    for person in data.persons:
        _write_person(wb, F, person, data)
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
                  data.masters.stock_rows, data.masters.stock_refreshed,
                  status=data.masters.stock_status)
    _write_bank_master(wb, F)
    _write_fd(wb, F, data)
    _write_ppf(wb, F, data)
    _write_ppf_ledger(wb, F, data)
    _write_epf(wb, F, data)
    _write_bonds(wb, F, data)
    _write_gold_silver(wb, F, data)
    _write_nps(wb, F, data)
    _write_master(wb, F, "NPS_Master", "NPS SCHEME MASTER (NPS Trust)",
                  "Every NPS scheme with a daily NAV — feeds the Scheme "
                  "dropdown on the NPS sheet. Kept sorted by scheme name. "
                  "Do not edit by hand: the updater refreshes it.",
                  ["Scheme Code", "Scheme Name", "PFM"],
                  data.masters.nps_rows, data.masters.nps_refreshed)
    _write_manual_assets(wb, F, data)
    _write_by_scrip(wb, F, data)
    _write_corporate_actions(wb, F, data)
    _write_dividends(wb, F, data)
    _write_history(wb, F, data)
    _write_guide(wb, F)

    # hide (never omit) the sheets of switched-off classes — data survives,
    # formulas keep resolving, flipping Yes brings everything back (SPEC §3.14).
    # A shared sheet (e.g. Manual_Assets) hides only when EVERY class on it
    # is off, hence the subtraction.
    off = {sheet
           for cls in ASSET_CLASSES if not effective_enabled(data, cls)
           for sheet in cls.sheets}
    on = {sheet
          for cls in ASSET_CLASSES if effective_enabled(data, cls)
          for sheet in cls.sheets}
    for ws in wb.worksheets():
        if ws.get_name() in off - on:
            ws.hide()

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
