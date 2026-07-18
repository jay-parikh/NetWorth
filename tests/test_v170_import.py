"""v1.7 import pipeline — merge engine + never-garbage gates (SPEC §6.17)."""

from datetime import date

from networth import model as M
from networth.importers.common import (ImportBatch, ImportedSipTxn,
                                       parse_date_any, parse_inr,
                                       triangle_ok)
from networth.importers.merge import merge_sip_batches

TODAY = date(2026, 7, 18)
ISIN = "INF879O01027"           # in no master — exercises the override path


def _data(persons=("Amit",), masters=()):
    d = M.PortfolioData(persons=list(persons))
    d.masters.mf_rows = list(masters)
    return d


def _txn(day, amount, nav, units=None, folio="F1", isin=ISIN,
         scheme="Parag Parikh Flexi Cap", ttype="PURCHASE"):
    if units is None and nav:
        units = round(amount / nav, 4)
    return ImportedSipTxn(folio=folio, isin=isin, scheme_name=scheme,
                          fund_house="PPFAS", txn_date=day, amount=amount,
                          nav=nav, units=units, txn_type=ttype)


def _batch(txns, closing=None):
    b = ImportBatch(source="cas", path="x.pdf",
                    accounts=[("F1", "AMIT KUMAR")], sip_txns=txns)
    if closing is not None:
        b.closing_units = closing
    return b


# ---- hardened parsing helpers ----------------------------------------------

def test_parse_inr_lakh_commas_and_negatives():
    assert parse_inr("1,23,456.78") == 123456.78
    assert parse_inr("₹ 10,000") == 10000
    assert parse_inr("(5,000.50)") == -5000.50
    assert parse_inr("-152.30") == -152.30
    for garbage in ("", None, "10-01-2024", "12 Jan", "1,2,3,4x",
                    "9,99,99,99,99,999"):        # > ₹100 crore = misparse
        assert parse_inr(garbage) is None


def test_parse_date_bounds_and_formats():
    for s in ("10-Jan-2024", "10-01-2024", "10/01/2024", "2024-01-10"):
        assert parse_date_any(s, TODAY) == date(2024, 1, 10)
    assert parse_date_any("31-12-1989", TODAY) is None   # pre-1990
    assert parse_date_any("01-01-2027", TODAY) is None   # future
    assert parse_date_any("not a date", TODAY) is None


def test_triangle_identity():
    assert triangle_ok(10000, 152.371, 65.629)           # ≈ within 0.1%
    assert not triangle_ok(10000, 152.371, 6.5629)       # shifted decimal
    assert not triangle_ok(10000, None, 65.0)            # unprovable
    assert not triangle_ok(10000, 152.0, 0)              # NAV must be > 0


# ---- merge: the happy path + idempotency -----------------------------------

def test_import_adds_rows_and_mf_summary():
    d = _data()
    txns = [_txn(date(2024, m, 5), 10000, 50.0) for m in (1, 2, 3)]
    rep = merge_sip_batches(d, [_batch(txns, {("F1", ISIN): 600.0})],
                            {"F1": "Amit"}, TODAY)
    assert rep.sip_added == 3 and not rep.deferred
    assert len(d.sip) == 3
    assert d.sip[0].units_override == 200.0              # verbatim units
    assert d.sip[0].isin_override == ISIN                # no master match
    assert [r for r in d.mutual_funds
            if (r.owner, r.scheme) == ("Amit", "Parag Parikh Flexi Cap")]
    line = rep.funds[0]
    assert line.ok and line.reconciled and line.invested == 30000


def test_reimport_adds_zero():
    d = _data()
    txns = [_txn(date(2024, m, 5), 10000, 50.0) for m in (1, 2, 3)]
    batch = _batch(txns, {("F1", ISIN): 600.0})
    merge_sip_batches(d, [batch], {"F1": "Amit"}, TODAY)
    rep2 = merge_sip_batches(d, [batch], {"F1": "Amit"}, TODAY)
    assert rep2.sip_added == 0 and rep2.sip_skipped == 3
    assert len(d.sip) == 3
    rep3 = merge_sip_batches(d, [batch], {"F1": "Amit"}, TODAY,
                             replace=True)               # replace-mode too
    assert rep3.sip_added == 0 and len(d.sip) == 3


def test_master_name_wins_when_isin_known():
    d = _data(masters=[("PPFAS", "Parag Parikh Flexi Cap Fund-Growth", ISIN)])
    rep = merge_sip_batches(d, [_batch([_txn(date(2024, 1, 5), 10000, 50.0)])],
                            {"F1": "Amit"}, TODAY)
    assert rep.sip_added == 1
    assert d.sip[0].scheme == "Parag Parikh Flexi Cap Fund-Growth"
    assert d.sip[0].isin_override == ""                  # master resolves it


# ---- statement-wins vs append-only -----------------------------------------

def test_statement_wins_replaces_typed_rows():
    d = _data(masters=[("PPFAS", "Parag Parikh Flexi Cap", ISIN)])
    d.sip = [M.SIPRow(owner="Amit", scheme="Parag Parikh Flexi Cap",
                      txn_date=date(2024, 1, 7), amount=10000, nav=49.0)]
    txns = [_txn(date(2024, 1, 5), 10000, 50.0),
            _txn(date(2024, 2, 5), 10000, 50.0)]
    rep = merge_sip_batches(d, [_batch(txns, {("F1", ISIN): 400.0})],
                            {"F1": "Amit"}, TODAY, replace=True)
    assert rep.sip_replaced == 1 and rep.sip_added == 2
    assert len(d.sip) == 2 and d.sip[0].nav == 50.0      # exact history


def test_append_only_keeps_typed_rows():
    d = _data(masters=[("PPFAS", "Parag Parikh Flexi Cap", ISIN)])
    typed = M.SIPRow(owner="Amit", scheme="Parag Parikh Flexi Cap",
                     txn_date=date(2024, 1, 7), amount=10000, nav=49.0)
    d.sip = [typed]
    txns = [_txn(date(2024, 1, 5), 10000, 50.0)]
    rep = merge_sip_batches(d, [_batch(txns)], {"F1": "Amit"}, TODAY)
    assert rep.sip_added == 1 and typed in d.sip         # untouched


def test_multiset_absorbs_same_day_twin_sips():
    d = _data(masters=[("PPFAS", "Parag Parikh Flexi Cap", ISIN)])
    d.sip = [M.SIPRow(owner="Amit", scheme="Parag Parikh Flexi Cap",
                      txn_date=date(2024, 1, 5), amount=10000, nav=50.0)]
    txns = [_txn(date(2024, 1, 5), 10000, 50.0),
            _txn(date(2024, 1, 5), 10000, 50.0)]         # genuine twins
    rep = merge_sip_batches(d, [_batch(txns)], {"F1": "Amit"}, TODAY)
    assert rep.sip_skipped == 1 and rep.sip_added == 1   # multiplicity
    assert len(d.sip) == 2


# ---- never-garbage gates ---------------------------------------------------

def test_triangle_failure_refuses_whole_fund_atomically():
    d = _data()
    good = [_txn(date(2025, m, 5), 5000, 40.0, folio="F2",
                 isin="INF000000012", scheme="Other Fund") for m in (1, 2)]
    bad = [_txn(date(2024, 1, 5), 10000, 50.0),
           _txn(date(2024, 2, 5), 10000, 50.0, units=2000.0)]  # 10× shift
    rep = merge_sip_batches(
        d, [_batch(bad + good)], {"F1": "Amit", "F2": "Amit"}, TODAY)
    refused = next(l for l in rep.funds if l.isin == ISIN)
    assert not refused.ok and "doesn't add up" in refused.reason
    assert all(r.scheme == "Other Fund" for r in d.sip)  # sibling imported
    assert len(d.sip) == 2


def test_balance_mismatch_refuses_fund():
    d = _data()
    txns = [_txn(date(2024, m, 5), 10000, 50.0) for m in (1, 2)]
    rep = merge_sip_batches(d, [_batch(txns, {("F1", ISIN): 999.0})],
                            {"F1": "Amit"}, TODAY)
    assert rep.sip_added == 0 and not d.sip
    assert "statement says 999.000" in rep.funds[0].reason


def test_negative_running_units_refuses_fund():
    d = _data()
    txns = [_txn(date(2024, 1, 5), 10000, 50.0),
            _txn(date(2024, 2, 5), -25000, 50.0, units=-500.0,
                 ttype="REDEMPTION")]
    rep = merge_sip_batches(d, [_batch(txns)], {"F1": "Amit"}, TODAY)
    assert rep.sip_added == 0 and "more units sold" in rep.funds[0].reason


def test_unmapped_folio_skips_with_reason():
    d = _data()
    rep = merge_sip_batches(d, [_batch([_txn(date(2024, 1, 5), 5000, 50.0)])],
                            {}, TODAY)
    assert rep.sip_added == 0 and not d.sip
    assert "not matched to a person" in rep.funds[0].reason


def test_bonus_units_only_row_is_accepted():
    d = _data()
    txns = [_txn(date(2024, 1, 5), 10000, 50.0),
            _txn(date(2024, 6, 5), 0.0, None, units=20.0, ttype="BONUS")]
    rep = merge_sip_batches(d, [_batch(txns, {("F1", ISIN): 220.0})],
                            {"F1": "Amit"}, TODAY)
    assert rep.sip_added == 2 and rep.funds[0].reconciled
    assert d.sip[1].amount == 0.0 and d.sip[1].units_override == 20.0


def test_import_map_round_trips(tmp_path):
    from networth.generate import build_workbook
    from networth.reader import read_workbook
    from networth.sample_data import sample_portfolio
    d = sample_portfolio()
    d.import_map = [M.ImportMapRow(source="fund statement (CAS)",
                                   account="12345678/90",
                                   name_hint="AMIT KUMAR", owner="Amit")]
    d.imported_files = [M.ImportedFileRow(
        file="CAS_Amit.pdf", fingerprint="a1b2c3d4e5f6",
        imported_on=date(2026, 7, 18), decision="imported")]
    p = tmp_path / "wb.xlsx"
    build_workbook(d, str(p))
    back = read_workbook(str(p))
    assert back.import_map == d.import_map
    assert back.imported_files == d.imported_files


def test_capacity_defers_whole_import_untouched():
    d = _data()
    cap = M.SIP_LAST_ROW - M.FIRST_DATA_ROW + 1
    d.sip = [M.SIPRow(owner="Amit", scheme="X", txn_date=date(2024, 1, 5),
                      amount=1000, nav=10.0) for _ in range(cap - 1)]
    txns = [_txn(date(2025, 1, d_), 1000, 10.0) for d_ in (5, 6, 7)]
    rep = merge_sip_batches(d, [_batch(txns)], {"F1": "Amit"}, TODAY)
    assert rep.deferred and rep.sip_added == 0
    assert len(d.sip) == cap - 1                          # untouched
    assert all(not l.ok for l in rep.funds)


# ---- broker CSV parsing (I3) -----------------------------------------------

from networth.importers.brokers import parse_equity_csv, sniff_csv  # noqa: E402
from networth.importers.common import ImportedTrade  # noqa: E402
from networth.importers.merge import merge_equity_batches  # noqa: E402

RELIANCE = "INE002A01018"
INFY = "INE009A01021"

ZERODHA_TB = (
    "symbol,isin,trade_date,exchange,segment,series,trade_type,quantity,"
    "price,trade_id,order_id,order_execution_time\n"
    f"RELIANCE,{RELIANCE},2024-01-10,NSE,EQ,EQ,buy,5,2500.00,t1,o1,"
    "2024-01-10T10:00:01\n"
    f"RELIANCE,{RELIANCE},2024-01-10,NSE,EQ,EQ,buy,5,2510.00,t2,o1,"
    "2024-01-10T10:00:02\n"
    f"RELIANCE,{RELIANCE},2024-03-12,NSE,EQ,EQ,sell,4,2700.00,t3,o2,"
    "2024-03-12T11:00:00\n")

GENERIC_TB = (
    "Company Name,ISIN Code,Txn Date,Action,No. of Shares,Rate\n"
    f"Infosys,{INFY},10-01-2024,Bought,10,1500.50\n")

HOLDINGS = (
    "symbol,isin,quantity available,average price\n"
    f"RELIANCE,{RELIANCE},6,2505.00\n")


def _eq_data(persons=("Amit",), stock=(("RELIANCE", "RELIANCE INDUSTRIES",
                                        RELIANCE),)):
    d = M.PortfolioData(persons=list(persons))
    d.masters.stock_rows = list(stock)
    return d


def test_zerodha_signature_and_fill_collapse():
    assert sniff_csv(ZERODHA_TB) == ("Zerodha tradebook", "trades")
    b = parse_equity_csv(ZERODHA_TB, TODAY)
    assert len(b.trades) == 2                             # fills collapsed
    buy = next(t for t in b.trades if t.side == "BUY")
    assert buy.qty == 10 and buy.price == 2505.0          # weighted


def test_generic_headers_parse():
    assert sniff_csv(GENERIC_TB) == ("broker tradebook", "trades")
    b = parse_equity_csv(GENERIC_TB, TODAY)
    assert b.trades[0].isin == INFY and b.trades[0].side == "BUY"
    assert b.trades[0].trade_date == date(2024, 1, 10)


def test_unrecognisable_csv_fails_politely():
    import pytest as _pytest
    with _pytest.raises(ValueError, match="couldn't recognise"):
        parse_equity_csv("a,b,c\n1,2,3\n", TODAY)


def test_buys_and_netted_sell_land_as_lots_and_sells():
    d = _eq_data()
    b = parse_equity_csv(ZERODHA_TB, TODAY)
    rep = merge_equity_batches(d, [b], {"": "Amit"}, TODAY, cg_on=True)
    assert rep.eq_added == 1               # fills collapsed → one lot, net 6
    assert [round(r.qty, 3) for r in d.equity] == [6]
    assert d.equity[0].scrip == "RELIANCE INDUSTRIES"     # master name
    assert d.equity[0].flag.startswith("IMPORTED:")
    assert rep.sells_added == 1
    assert d.equity_sells[0].qty == 4
    assert d.equity_sells[0].buy_price == 2505.0          # FIFO lot cost


def test_equity_reimport_is_stateless_idempotent():
    d = _eq_data()
    b = parse_equity_csv(ZERODHA_TB, TODAY)
    merge_equity_batches(d, [b], {"": "Amit"}, TODAY, cg_on=True)
    before = [
        (r.scrip, r.qty, r.avg_cost, r.cost_date) for r in d.equity]
    rep2 = merge_equity_batches(d, [b], {"": "Amit"}, TODAY, cg_on=True)
    assert rep2.eq_added == 0 and rep2.sells_added == 0
    assert [(r.scrip, r.qty, r.avg_cost, r.cost_date)
            for r in d.equity] == before


def test_uncovered_sell_refuses_isin():
    d = _eq_data()
    b = ImportBatch(source="broker tradebook", trades=[ImportedTrade(
        account="", isin=RELIANCE, symbol="RELIANCE",
        trade_date=date(2024, 3, 12), qty=4, price=2700.0, side="SELL")])
    rep = merge_equity_batches(d, [b], {"": "Amit"}, TODAY)
    assert rep.eq_added == 0 and not d.equity
    assert "sells more than it buys" in rep.stocks[0].reason


def test_replace_mode_consumes_typed_lots():
    d = _eq_data()
    d.equity = [M.EquityRow(owner="Amit", scrip="RELIANCE INDUSTRIES",
                            qty=10, avg_cost=2000.0,
                            cost_date=date(2023, 1, 1))]
    b = ImportBatch(source="broker tradebook", trades=[ImportedTrade(
        account="", isin=RELIANCE, symbol="RELIANCE",
        trade_date=date(2024, 3, 12), qty=4, price=2700.0, side="SELL")])
    rep = merge_equity_batches(d, [b], {"": "Amit"}, TODAY,
                               cg_on=True, replace=True)
    assert rep.eq_reduced == 1 and d.equity[0].qty == 6
    assert rep.sells_added == 1
    assert d.equity_sells[0].buy_price == 2000.0          # typed lot cost


def test_ca_inside_trade_window_refuses_isin():
    d = _eq_data()
    d.corporate_actions = [M.CorporateAction(
        symbol="RELIANCE", isin=RELIANCE, type="SPLIT",
        ex_date=date(2024, 2, 1), ratio_from=10, ratio_to=5, source="Auto")]
    b = parse_equity_csv(ZERODHA_TB, TODAY)
    rep = merge_equity_batches(d, [b], {"": "Amit"}, TODAY)
    assert rep.eq_added == 0 and not d.equity
    assert "split" in rep.stocks[0].reason


def test_holdings_fallback_and_crosscheck():
    d = _eq_data()
    b = parse_equity_csv(HOLDINGS, TODAY)
    rep = merge_equity_batches(d, [b], {"": "Amit"}, TODAY)
    assert rep.eq_added == 1 and d.equity[0].qty == 6
    assert d.equity[0].cost_date is None                  # dated by the user
    assert any("no buy date" in w for w in rep.warnings)
    # second run: rows exist now → cross-check path, no duplicate
    rep2 = merge_equity_batches(d, [b], {"": "Amit"}, TODAY)
    assert rep2.eq_added == 0 and len(d.equity) == 1


def test_unknown_symbol_without_isin_warns():
    d = _eq_data()
    tb = ("Company Name,Txn Date,Action,No. of Shares,Rate\n"
          "Mystery Co,10-01-2024,Bought,10,100\n")
    b = parse_equity_csv(tb, TODAY)
    rep = merge_equity_batches(d, [b], {"": "Amit"}, TODAY)
    assert rep.eq_added == 0 and not d.equity
    assert any("isn't in the stock list" in w for w in rep.warnings)


# ---- CAS PDF text parser (I4) ----------------------------------------------

import pytest  # noqa: E402

from networth.importers import detect_format  # noqa: E402
from networth.importers.cams import looks_like_cas, parse as parse_cas  # noqa: E402
from networth.importers.merge import condense_txns  # noqa: E402

CAS_TEXT = """\
Consolidated Account Statement
CAMS + KFintech
01-Jan-2015 To 18-Jul-2026

AMIT KUMAR
email@example.com

XYZ123 - Parag Parikh Flexi Cap Fund - Direct Plan - Growth (Advisor: ARN-0000)  ISIN: INF879O01027
Folio No: 12345678 / 90
PAN: ABCDE1234F   KYC: OK
Opening Unit Balance: 0.000
01-Jan-2024      Purchase - Systematic                 10,000.00      200.0000      50.0000       200.000
01-Jan-2024      *** Stamp Duty ***                         0.50
05-Feb-2024      Purchase                               10,000.00      196.0784      51.0000       396.078
10-Mar-2024      Redemption                            (5,100.00)     (100.0000)     51.0000       296.078
Closing Unit Balance: 296.078   NAV on 18-Jul-2026: INR 95.50
"""

KFIN_TEXT = """\
CONSOLIDATED ACCOUNT STATEMENT
KFintech / CAMS
PRIYA SHARMA

Some Debt Fund - Growth ISIN: INF999X01010
Folio No: 555/44
Opening Unit Balance: 0.000
07-Jun-2023      SIP Purchase                            5,000.00      100.0000      50.0000       100.000
Closing Unit Balance: 100.000
"""


def test_cas_parse_and_merge_end_to_end():
    assert looks_like_cas(CAS_TEXT)
    b = parse_cas(CAS_TEXT, TODAY)
    assert len(b.sip_txns) == 3                          # stamp duty skipped
    assert b.closing_units[("12345678/90", ISIN)] == 296.078
    assert b.accounts == [("12345678/90", "Amit Kumar")]
    red = b.sip_txns[2]
    assert red.amount == -5100.0 and red.units == -100.0
    d = _data()
    rep = merge_sip_batches(d, [b], {"12345678/90": "Amit"}, TODAY)
    assert rep.sip_added == 3 and rep.funds[0].reconciled
    assert d.sip[0].scheme.startswith("XYZ123 - Parag Parikh")


def test_kfin_variant_parses():
    b = parse_cas(KFIN_TEXT, TODAY)
    assert len(b.sip_txns) == 1
    assert b.sip_txns[0].isin == "INF999X01010"
    assert b.closing_units[("555/44", "INF999X01010")] == 100.0


def test_summary_cas_refused_with_teaching_message():
    summary = ("Consolidated Account Statement\nCAMS KFintech\n"
               "Fund X ISIN: INF879O01027\nFolio No: 111\n"
               "Closing Unit Balance: 100.000\n")
    with pytest.raises(ValueError, match="DETAILED"):
        parse_cas(summary, TODAY)


def test_mid_history_statement_fund_is_dropped_with_warning():
    text = CAS_TEXT.replace("Opening Unit Balance: 0.000",
                            "Opening Unit Balance: 50.000")
    b = parse_cas(text, TODAY)
    assert not b.sip_txns and not b.closing_units
    assert any("Since inception" in w for w in b.warnings)


def test_mangled_line_warns_and_balance_check_refuses():
    text = CAS_TEXT.replace(
        "05-Feb-2024      Purchase                               10,000.00"
        "      196.0784      51.0000       396.078",
        "05-Feb-2024      Purchase                               10,000.00")
    b = parse_cas(text, TODAY)
    assert any("couldn't read the line" in w for w in b.warnings)
    d = _data()
    rep = merge_sip_batches(d, [b], {"12345678/90": "Amit"}, TODAY)
    assert rep.sip_added == 0 and not d.sip              # refused whole
    assert "statement says" in rep.funds[0].reason


def test_condense_conserves_totals_and_reconciles():
    txns = [_txn(date(2016 + i, 1, 5), 1000, 10.0 + i) for i in range(7)]
    txns.append(_txn(date(2024, 1, 5), 1000, 20.0))      # recent — kept
    out = condense_txns(txns, date(2023, 4, 1))
    old = [t for t in out if t.txn_type == "OPENING"]
    assert len(old) == 1 and len(out) == 2               # 7 rolled + 1 kept
    assert old[0].amount == 7000.0
    assert old[0].units == round(sum(
        round(1000 / (10.0 + i), 4) for i in range(7)), 4)
    assert abs(old[0].amount - old[0].units * old[0].nav) <= 1.0  # triangle
    # totals conserved → the closing-balance anchor still proves the fund
    total = round(sum(t.units for t in out), 3)
    d = _data()
    b = _batch(out, {("F1", ISIN): total})
    rep = merge_sip_batches(d, [b], {"F1": "Amit"}, TODAY)
    assert rep.sip_added == 2 and rep.funds[0].reconciled


# ---- pdftext: extraction + password behaviour (I4) --------------------------

def _mini_pdf(lines, path):
    """A tiny valid one-page PDF carrying `lines` as text — built by hand
    so no PDF binary ever lands in the repo."""
    stream = ("BT /F1 10 Tf 40 800 Td " + " ".join(
        f"({ln.replace('(', '[').replace(')', ']')}) Tj 0 -14 Td"
        for ln in lines) + " ET").encode()
    objs = [
        b"<< /Type /Catalog /Pages 2 0 R >>",
        b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 842] "
        b"/Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>",
        b"<< /Length " + str(len(stream)).encode() + b" >>\nstream\n"
        + stream + b"\nendstream",
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
    ]
    out, offsets = b"%PDF-1.4\n", []
    for i, body in enumerate(objs, start=1):
        offsets.append(len(out))
        out += f"{i} 0 obj\n".encode() + body + b"\nendobj\n"
    xref = len(out)
    out += f"xref\n0 {len(objs) + 1}\n0000000000 65535 f \n".encode()
    for off in offsets:
        out += f"{off:010d} 00000 n \n".encode()
    out += (f"trailer\n<< /Size {len(objs) + 1} /Root 1 0 R >>\n"
            f"startxref\n{xref}\n%%EOF").encode()
    path.write_bytes(out)


def test_pdftext_extracts_and_enforces_password(tmp_path):
    from pypdf import PdfReader, PdfWriter

    from networth.importers.pdftext import (NeedsPassword, WrongPassword,
                                            extract_text)
    plain = tmp_path / "cas.pdf"
    _mini_pdf(["Consolidated Account Statement", "CAMS KFintech",
               "Folio No: 111"], plain)
    text = extract_text(plain)
    assert "Consolidated Account Statement" in text
    assert detect_format(plain) == "cas"

    locked = tmp_path / "locked.pdf"
    w = PdfWriter(clone_from=PdfReader(str(plain)))
    w.encrypt("secret", algorithm="AES-256")
    with open(locked, "wb") as fh:
        w.write(fh)
    with pytest.raises(NeedsPassword):
        extract_text(locked)
    with pytest.raises(WrongPassword):
        extract_text(locked, "wrong")
    assert "Folio No: 111" in extract_text(locked, "secret")


def test_detect_format_csv_and_garbage(tmp_path):
    good = tmp_path / "tradebook.csv"
    good.write_text(ZERODHA_TB)
    assert detect_format(good) == "equity_csv"
    bad = tmp_path / "random.csv"
    bad.write_text("a,b,c\n1,2,3\n")
    assert detect_format(bad) is None
    txt = tmp_path / "notes.txt"
    txt.write_text("hello")
    assert detect_format(txt) is None


# ---- end-to-end: run() with import batches (I5) -----------------------------

def test_run_imports_and_persists_map_and_nevernag(tmp_path):
    from networth.fetch.amfi import AmfiData
    from networth.generate import build_workbook
    from networth.reader import read_workbook
    from networth.sample_data import sample_portfolio
    from networth.update import run

    d = sample_portfolio()
    path = tmp_path / "wb.xlsx"
    build_workbook(d, str(path))
    b = parse_cas(CAS_TEXT, TODAY)
    b.fingerprint = "cafe12345678"
    summary = run(path, price_data=None, amfi_data=AmfiData(), ca_data=[],
                  div_data=[], today=TODAY,
                  import_batches=[b],
                  import_owner_map={"12345678/90": "Amit"},
                  import_decisions=[("cas.pdf", "cafe12345678", "imported")])
    rep = summary["import_report"]
    assert rep.sip_added == 3 and rep.funds[0].reconciled
    back = read_workbook(str(path))
    # the sample master knows this ISIN, so the imported rows take the
    # master's scheme name; they are the only rows with verbatim units
    got = [r for r in back.sip if r.units_override is not None]
    assert len(got) == 3 and got[0].units_override == 200.0
    assert back.import_map[0].account == "12345678/90"
    assert back.import_map[0].owner == "Amit"
    assert back.imported_files[0].fingerprint == "cafe12345678"
    # second run: the stored Import_Map answers the owner question, and
    # the merge adds nothing — idempotent end-to-end
    b2 = parse_cas(CAS_TEXT, TODAY)
    s2 = run(path, price_data=None, amfi_data=AmfiData(), ca_data=[],
             div_data=[], today=TODAY, import_batches=[b2])
    rep2 = s2["import_report"]
    assert rep2.sip_added == 0 and rep2.sip_skipped == 3
    assert len([r for r in read_workbook(str(path)).sip
                if r.units_override is not None]) == 3
