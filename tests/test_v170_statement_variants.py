"""v1.7 — statement-variant torture battery (SPEC §6.17).

Every layout quirk the wild produces must either import PROVEN-correct or
refuse with a plain reason — never a wrong number (the never-garbage
contract). Fixtures are synthetic but shaped on real CAMS/KFintech and
broker exports.
"""

from datetime import date

from networth import model as M
from networth.importers.brokers import parse_equity_csv, sniff_csv
from networth.importers.cams import parse as parse_cas
from networth.importers.merge import (merge_equity_batches,
                                      merge_sip_batches)

TODAY = date(2026, 7, 18)


def _merge(text, owner_map=None, persons=("Amit", "Priya")):
    d = M.PortfolioData(persons=list(persons))
    b = parse_cas(text, TODAY)
    rep = merge_sip_batches(d, [b],
                            owner_map or {a: "Amit" for a, _ in b.accounts},
                            TODAY)
    return d, b, rep


# ---- KFin multi-line scheme header ------------------------------------------

KFIN_MULTILINE = """\
CONSOLIDATED ACCOUNT STATEMENT
KFintech
PRIYA SHARMA

Axis Bluechip Fund - Direct Growth
ISIN: INF846K01EW2   Registrar : KFINTECH
Folio No: 91099/22
Opening Unit Balance: 0.000
10-Apr-2024      Systematic Purchase                     4,999.75       99.9950      50.0000        99.995
Closing Unit Balance: 99.995
"""


def test_scheme_name_on_its_own_line_is_captured():
    d, b, rep = _merge(KFIN_MULTILINE)
    assert rep.sip_added == 1 and rep.funds[0].reconciled
    assert d.sip[0].scheme == "Axis Bluechip Fund - Direct Growth"
    assert d.sip[0].isin_override == "INF846K01EW2"


# ---- multi-folio, multi-investor, page noise --------------------------------

MULTI = """\
Consolidated Account Statement
CAMS + KFintech
01-Jan-2015 To 18-Jul-2026
AMIT KUMAR
amit@example.com

Fund One - Growth ISIN: INF111A01011
Folio No: 111/11
Opening Unit Balance: 0.000
05-Jan-2024      Purchase - SIP                          10,000.00      200.0000      50.0000       200.000
05-Jan-2024      *** Stamp Duty ***                          0.50
05-Jan-2024      STT Paid                                    0.10
Closing Unit Balance: 200.000
                       Page 1 of 2
Consolidated Account Statement
PRIYA SHARMA

Fund Two - Growth ISIN: INF222B01022
Folio No: 222/22
Opening Unit Balance: 0.000
07-Feb-2024      Purchase                              1,23,456.78    1,234.5678    100.0000     1,234.568
10-Mar-2024      Dividend Reinvestment                      500.00        5.0000    100.0000     1,239.568
Closing Unit Balance: 1239.568
"""


def test_multi_folio_multi_investor_with_page_noise():
    d, b, rep = _merge(MULTI, owner_map={"111/11": "Amit",
                                         "222/22": "Priya"})
    assert rep.sip_added == 3                            # noise rows skipped
    assert all(l.ok and l.reconciled for l in rep.funds)
    assert {a for a, _h in b.accounts} == {"111/11", "222/22"}
    hints = dict(b.accounts)
    assert hints["111/11"] == "Amit Kumar"            # per-folio hint
    assert hints["222/22"] == "Priya Sharma"
    lakh = next(r for r in d.sip if r.owner == "Priya")
    assert lakh.amount == 123456.78                      # lakh commas
    assert lakh.units_override == 1234.5678


# ---- sign conventions: type wins over print style ---------------------------

SIGNS = """\
Consolidated Account Statement CAMS KFintech
Fund One - Growth ISIN: INF111A01011
Folio No: 111/11
Opening Unit Balance: 0.000
05-Jan-2024      Purchase                                10,000.00      200.0000      50.0000       200.000
10-Feb-2024      Redemption                               5,100.00      100.0000      51.0000       100.000
Closing Unit Balance: 100.000
"""


def test_redemption_printed_positive_is_normalised_negative():
    d, b, rep = _merge(SIGNS)
    assert rep.sip_added == 2 and rep.funds[0].reconciled
    red = next(r for r in d.sip if (r.amount or 0) < 0)
    assert red.amount == -5100.0 and red.units_override == -100.0


SWITCH = """\
Consolidated Account Statement CAMS KFintech
Fund One - Growth ISIN: INF111A01011
Folio No: 111/11
Opening Unit Balance: 0.000
05-Jan-2024      Purchase                                10,000.00      200.0000      50.0000       200.000
10-Feb-2024      Switch Out - To Fund Two                (5,100.00)    (100.0000)     51.0000       100.000
Closing Unit Balance: 100.000

Fund Two - Growth ISIN: INF222B01022
Folio No: 111/11
Opening Unit Balance: 0.000
10-Feb-2024      Switch In - From Fund One                5,100.00       51.0000     100.0000        51.000
Closing Unit Balance: 51.000
"""


def test_switch_pair_reconciles_both_funds():
    d, b, rep = _merge(SWITCH)
    assert rep.sip_added == 3
    assert all(l.ok and l.reconciled for l in rep.funds)
    out = next(r for r in d.sip if (r.amount or 0) < 0)
    assert out.amount == -5100.0                          # money left fund 1


BONUS = """\
Consolidated Account Statement CAMS KFintech
Fund One - Growth ISIN: INF111A01011
Folio No: 111/11
Opening Unit Balance: 0.000
05-Jan-2024      Purchase                                10,000.00      200.0000      50.0000       200.000
10-Jun-2024      Bonus Units Allotted                         0.00       20.0000       0.0000       220.000
Closing Unit Balance: 220.000
"""


def test_bonus_units_row_reconciles():
    d, b, rep = _merge(BONUS)
    assert rep.sip_added == 2 and rep.funds[0].reconciled
    bonus = d.sip[1]
    assert bonus.amount == 0.0 and bonus.units_override == 20.0


# ---- one bad fund never poisons its siblings --------------------------------

def test_mangled_fund_refused_sibling_imports_end_to_end():
    text = MULTI.replace(
        "07-Feb-2024      Purchase                              "
        "1,23,456.78    1,234.5678    100.0000     1,234.568",
        "07-Feb-2024      Purchase                              gibberish")
    d, b, rep = _merge(text, owner_map={"111/11": "Amit",
                                        "222/22": "Priya"})
    ok = [l for l in rep.funds if l.ok]
    bad = [l for l in rep.funds if not l.ok]
    assert len(ok) == 1 and len(bad) == 1
    assert "statement says" in bad[0].reason              # balance caught it
    assert all(r.owner == "Amit" for r in d.sip)          # sibling landed


# ---- broker CSV variants -----------------------------------------------------

GROWW_HOLDINGS = ("﻿Stock Name,ISIN,Quantity,Average Buy Price\n"
                  "Reliance Industries,INE002A01018,10,\"2,400.50\"\n")

UPSTOX_TB = ("Upstox Tradebook Export\n"
             "Generated on 18-07-2026\n"
             "Scrip Name,ISIN,Trade Date,Buy/Sell,Qty,Price\n"
             "RELIANCE,INE002A01018,10/01/2024,Buy,10,\"2,400.50\"\n"
             "RELIANCE,INE002A01018,12/03/2024,Sell,4,\"2,700.00\"\n")

ICICI_TB = ("Security Name,ISIN Code,Transaction Type,Deal Date,"
            "Quantity,Price per share\n"
            "INFOSYS LTD,INE009A01021,B,15-Jan-2024,12,1500\n")


def test_groww_style_holdings_with_bom_and_quoted_commas():
    assert sniff_csv(GROWW_HOLDINGS) == ("broker holdings", "holdings")
    b = parse_equity_csv(GROWW_HOLDINGS, TODAY)
    assert b.holdings[0].qty == 10 and b.holdings[0].avg_cost == 2400.50


def test_upstox_style_tradebook_with_banner_rows():
    assert sniff_csv(UPSTOX_TB) == ("broker tradebook", "trades")
    b = parse_equity_csv(UPSTOX_TB, TODAY)
    assert len(b.trades) == 2 and not b.warnings
    d = M.PortfolioData(persons=["Amit"])
    d.masters.stock_rows = [("RELIANCE", "RELIANCE INDUSTRIES",
                             "INE002A01018")]
    rep = merge_equity_batches(d, [b], {"": "Amit"}, TODAY)
    assert rep.eq_added == 1 and d.equity[0].qty == 6


def test_icici_style_single_letter_side_and_deal_date():
    b = parse_equity_csv(ICICI_TB, TODAY)
    assert len(b.trades) == 1 and not b.warnings
    t = b.trades[0]
    assert t.side == "BUY" and t.trade_date == date(2024, 1, 15)
    assert t.qty == 12 and t.price == 1500


def test_crlf_and_windows_export():
    b = parse_equity_csv(UPSTOX_TB.replace("\n", "\r\n"), TODAY)
    assert len(b.trades) == 2


# ---- new gates: sign-vs-type, bonus needs an anchor, multi-folio anchors ----

from networth.importers.common import ImportBatch, ImportedSipTxn  # noqa: E402


def _sip(day, amount, nav, units, folio="F1", isin="INF111A01011",
         ttype="PURCHASE"):
    return ImportedSipTxn(folio=folio, isin=isin, scheme_name="Fund One",
                          txn_date=day, amount=amount, nav=nav, units=units,
                          txn_type=ttype)


def test_sign_type_disagreement_refuses_fund():
    d = M.PortfolioData(persons=["Amit"])
    txns = [_sip(date(2024, 1, 5), 10000, 50.0, 200.0),
            # a "redemption" with positive money — parser bug or layout
            # drift; the MERGE gate must catch it, not the parser
            _sip(date(2024, 2, 5), 5100, 51.0, 100.0, ttype="REDEMPTION")]
    rep = merge_sip_batches(d, [ImportBatch(source="cas", sip_txns=txns)],
                            {"F1": "Amit"}, TODAY)
    assert rep.sip_added == 0 and not d.sip
    assert "reads as money coming in" in rep.funds[0].reason


def test_bonus_without_closing_balance_is_refused():
    d = M.PortfolioData(persons=["Amit"])
    txns = [_sip(date(2024, 1, 5), 10000, 50.0, 200.0),
            _sip(date(2024, 6, 5), 0.0, None, 20.0, ttype="BONUS")]
    rep = merge_sip_batches(d, [ImportBatch(source="cas", sip_txns=txns)],
                            {"F1": "Amit"}, TODAY)   # no closing_units
    assert rep.sip_added == 0 and "can't be checked" in rep.funds[0].reason


def test_two_folios_one_fund_anchors_sum():
    d = M.PortfolioData(persons=["Amit"])
    txns = [_sip(date(2024, 1, 5), 10000, 50.0, 200.0, folio="F1"),
            _sip(date(2024, 2, 5), 10000, 50.0, 200.0, folio="F2")]
    b = ImportBatch(source="cas", sip_txns=txns,
                    closing_units={("F1", "INF111A01011"): 200.0,
                                   ("F2", "INF111A01011"): 200.0})
    rep = merge_sip_batches(d, [b], {"F1": "Amit", "F2": "Amit"}, TODAY,
                            replace=True)
    assert rep.sip_added == 2
    # the combined sheet total (400) must NOT trigger the single-folio
    # false alarm — anchors sum across folios
    assert not any("check for doubled history" in w for w in rep.warnings)


# ---- the pre-2018 opening-lot answer (missing buys) -------------------------

def test_pre2018_opening_covers_missing_buy():
    from networth.importers.common import ImportedTrade
    d = M.PortfolioData(persons=["Amit"])
    d.masters.stock_rows = [("RELIANCE", "RELIANCE INDUSTRIES",
                             "INE002A01018")]
    b = ImportBatch(source="broker tradebook", trades=[ImportedTrade(
        account="", isin="INE002A01018", symbol="RELIANCE",
        trade_date=date(2024, 3, 12), qty=40, price=2700.0, side="SELL")])
    # user said yes: these shares predate Feb 2018
    rep = merge_equity_batches(d, [b], {"": "Amit"}, TODAY, cg_on=True,
                               pre2018_openings={("Amit", "INE002A01018")})
    assert rep.stocks[0].ok
    assert rep.sells_added == 1
    s = d.equity_sells[0]
    assert s.buy_date == date(2018, 1, 31)               # FMV convention
    assert s.buy_price is None                           # official value fills
    assert not d.equity                                  # fully sold
    assert any("before Feb 2018" in w for w in rep.warnings)


def test_missing_buy_without_answer_still_refuses():
    from networth.importers.common import ImportedTrade
    d = M.PortfolioData(persons=["Amit"])
    d.masters.stock_rows = [("RELIANCE", "RELIANCE INDUSTRIES",
                             "INE002A01018")]
    b = ImportBatch(source="broker tradebook", trades=[ImportedTrade(
        account="", isin="INE002A01018", symbol="RELIANCE",
        trade_date=date(2024, 3, 12), qty=40, price=2700.0, side="SELL")])
    rep = merge_equity_batches(d, [b], {"": "Amit"}, TODAY)
    assert not rep.stocks[0].ok and "sells more" in rep.stocks[0].reason


# ---- traditional back-office (MoneyMaker-style, e.g. Shah Investors) --------

SIHL_REGISTER = ("Client Code,Scrip Name,ISIN,Date,Buy Qty,Buy Rate,"
                 "Sell Qty,Sell Rate\n"
                 "SIHL123,RELIANCE INDUSTRIES,INE002A01018,05-01-2024,"
                 "10,\"2,400.50\",,\n"
                 "SIHL123,RELIANCE INDUSTRIES,INE002A01018,12-03-2024,"
                 ",,4,\"2,700.00\"\n"
                 "SIHL123,RELIANCE INDUSTRIES,INE002A01018,20-03-2024,"
                 "2,2500,1,2650\n")

SIHL_HOLDINGS = ("Scrip Name,ISIN,Qty,Rate,Value\n"
                 "RELIANCE INDUSTRIES,INE002A01018,7,\"2,420.00\","
                 "\"16,940.00\"\n")


def test_split_column_register_parses_both_legs():
    assert sniff_csv(SIHL_REGISTER) == ("broker transaction register",
                                        "trades_split")
    b = parse_equity_csv(SIHL_REGISTER, TODAY)
    assert not b.warnings
    sides = sorted((t.side, t.qty) for t in b.trades)
    assert sides == [("BUY", 2), ("BUY", 10), ("SELL", 1), ("SELL", 4)]
    d = M.PortfolioData(persons=["Amit"])
    d.masters.stock_rows = [("RELIANCE", "RELIANCE INDUSTRIES",
                             "INE002A01018")]
    rep = merge_equity_batches(d, [b], {"SIHL123": "Amit"}, TODAY)
    assert rep.stocks[0].ok
    assert round(sum(r.qty for r in d.equity), 3) == 7    # 12 bought − 5 sold


def test_backoffice_holdings_with_rate_header():
    assert sniff_csv(SIHL_HOLDINGS) == ("broker holdings", "holdings")
    b = parse_equity_csv(SIHL_HOLDINGS, TODAY)
    assert b.holdings[0].qty == 7 and b.holdings[0].avg_cost == 2420.0


def test_sihl_shaped_xlsx_holdings_with_zero_average(tmp_path):
    # shaped on a real MoneyMaker export (2026-07-18): banner + summary
    # block, header at row 6, symbols with series suffix, Average 0 on
    # demat-converted paper holdings. Values here are synthetic.
    import xlsxwriter
    p = tmp_path / "Holdings_test.xlsx"
    wbk = xlsxwriter.Workbook(str(p))
    ws = wbk.add_worksheet("Sheet1")
    ws.write_row(0, 0, ["Holdings (2)"])
    ws.write_row(2, 0, ["Invested", "Current", "Overall Profit/Loss",
                        "Today's Profit/Loss"])
    ws.write_row(3, 0, ["281", "4173", "4173 (100%)", "-25 (-8%)"])
    ws.write_row(5, 0, ["Symbol", "Quantity", "Average", "Current",
                        "Invested", "LTP", "Change In Percentage",
                        "P & L", "P & L (%)", "MTM"])
    ws.write_row(6, 0, ["AJMERA EQ", 15, 0, 1902, 0, 126.8, -1.78,
                        1902, 100, -34])
    ws.write_row(7, 0, ["ONGC EQ", 1, 281, 247, 281, 247.29, 0.22,
                        -33, -12, 0])
    wbk.close()

    from networth.importers import detect_format
    from networth.importers.brokers import parse_equity_csv, rows_from_xlsx
    assert detect_format(p) == "equity_xlsx"
    b = parse_equity_csv(rows_from_xlsx(p), TODAY)
    assert len(b.holdings) == 2
    ajmera = next(h for h in b.holdings if "AJMERA" in h.name)
    assert ajmera.avg_cost is None            # 0 = unknown, never ₹0
    assert next(h for h in b.holdings if "ONGC" in h.name).avg_cost == 281

    d = M.PortfolioData(persons=["Amit"])
    d.masters.stock_rows = [("AJMERA", "AJMERA REALTY", "INE298G01027"),
                            ("ONGC", "OIL AND NATURAL GAS", "INE213A01029")]
    rep = merge_equity_batches(d, [b], {"": "Amit"}, TODAY,
                               pre2018_openings={("Amit", "AJMERA EQ")})
    assert rep.eq_added == 2
    aj = next(r for r in d.equity if r.scrip == "AJMERA REALTY")
    assert aj.avg_cost is None                # blank cost →
    assert aj.cost_date == date(2018, 1, 31)  # official 2018 value fills
    on = next(r for r in d.equity if "OIL" in r.scrip)
    assert on.avg_cost == 281 and on.cost_date is None
    assert any("official\n2018 value stands in".replace("\n", " ") in w
               or "official 2018 value stands in" in w
               for w in rep.warnings)


def test_own_workbook_is_never_an_import_candidate(tmp_path):
    from networth.generate import build_workbook
    from networth.importers import detect_format
    from networth.sample_data import sample_portfolio
    p = tmp_path / "Family_Portfolio_Tracker.xlsx"
    build_workbook(sample_portfolio(), str(p))
    assert detect_format(p) is None           # Dashboard ≠ broker headers


ZERODHA_NEW_HOLDINGS = (
    '"Instrument","Qty.","Avg. cost","LTP","Invested","Cur. val","P&L",'
    '"Net chg.","Day chg.",""\n'
    '"RELIANCE",5,2400.50,2700,12002.5,13500,1497.5,12.4,0.5,""\n'
    '"SOMENCD028",10,0,9746.01,0,97460.1,0,0,0.46,""\n'
    '"Some Flexi Cap Fund",3499.825,10,9.9772,34998.25,34918.45,'
    '-79.8,-0.23,0.2,""\n')


def test_zerodha_current_holdings_headers_and_mixed_rows():
    # the 2026 Kite export: "Qty." with a period, no ISIN column, fund
    # units and NCDs mixed in with the shares
    assert sniff_csv(ZERODHA_NEW_HOLDINGS) == ("broker holdings",
                                               "holdings")
    b = parse_equity_csv(ZERODHA_NEW_HOLDINGS, TODAY)
    assert len(b.holdings) == 3
    ncd = next(h for h in b.holdings if h.name == "SOMENCD028")
    assert ncd.avg_cost is None                # 0 = unknown, never ₹0
    d = M.PortfolioData(persons=["Amit"])
    d.masters.stock_rows = [("RELIANCE", "RELIANCE INDUSTRIES",
                             "INE002A01018")]
    rep = merge_equity_batches(d, [b], {"": "Amit"}, TODAY)
    assert rep.eq_added == 1                   # the share row
    assert d.equity[0].avg_cost == 2400.50
    # the fund row and the unknown NCD are left out LOUDLY, never silently
    assert sum(1 for w in rep.warnings
               if "isn't in the stock list" in w) == 2
    assert any("Fund units come via" in w for w in rep.warnings)
    assert any("no buy date" in w for w in rep.warnings)  # aggregated
