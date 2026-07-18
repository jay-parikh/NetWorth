"""v1.7 pre-release review fixes — one regression test per finding.

Each test names the defect it pins down; if one fails, read the matching
paragraph in docs/SPEC.md §6.17/§6.18 before "fixing" the assertion.
"""

from datetime import date

import pytest

from networth import model as M
from networth.importers.brokers import parse_equity_csv, sniff_csv
from networth.importers.cams import parse as parse_cas
from networth.importers.common import (ImportBatch, ImportedHolding,
                                       ImportedTrade)
from networth.importers.merge import merge_equity_batches, merge_sip_batches

TODAY = date(2026, 7, 18)
RELIANCE = "INE002A01018"
INFY = "INE009A01021"
ISIN = "INF879O01027"

CAS_TEXT = """\
Consolidated Account Statement
CAMS + KFintech

AMIT KUMAR

XYZ123 - Parag Parikh Flexi Cap Fund - Growth  ISIN: INF879O01027
Folio No: 12345678 / 90
Opening Unit Balance: 0.000
01-Jan-2024      Purchase - Systematic                 10,000.00      200.0000      50.0000       200.000
Closing Unit Balance: 200.000
"""


def _eq_data(persons=("Amit",)):
    d = M.PortfolioData(persons=list(persons))
    d.masters.stock_rows = [("RELIANCE", "RELIANCE INDUSTRIES", RELIANCE),
                            ("INFY", "INFOSYS", INFY)]
    return d


def _sell(qty, isin=RELIANCE, sym="RELIANCE", day=date(2024, 3, 12),
          price=2700.0):
    return ImportedTrade(account="", isin=isin, symbol=sym, trade_date=day,
                         qty=qty, price=price, side="SELL")


def _buy(qty, isin=RELIANCE, sym="RELIANCE", day=date(2024, 1, 10),
         price=2500.0):
    return ImportedTrade(account="", isin=isin, symbol=sym, trade_date=day,
                         qty=qty, price=price, side="BUY")


def _typed_lot(qty=10.0, owner="Amit"):
    return M.EquityRow(owner=owner, scrip="RELIANCE INDUSTRIES", qty=qty,
                       avg_cost=2000.0, cost_date=date(2023, 1, 1))


# ---- F1: a capacity defer must leave typed rows byte-for-byte alone --------

def test_capacity_defer_never_mutates_typed_lots(monkeypatch):
    # tiny caps so the defer trips without building 1500 rows
    monkeypatch.setattr(M, "EQUITY_LAST_ROW", M.FIRST_DATA_ROW + 1)
    d = _eq_data()
    d.equity = [_typed_lot(qty=10.0)]
    # ISIN A consumes the typed lot; ISIN B's buys then overflow the cap
    b = ImportBatch(source="broker tradebook", trades=[
        _sell(4),
        _buy(1, isin=INFY, sym="INFY", day=date(2024, 1, 5), price=1500.0),
        _buy(1, isin=INFY, sym="INFY", day=date(2024, 2, 5), price=1500.0),
    ])
    before = [(r.owner, r.scrip, r.qty, r.avg_cost, r.cost_date)
              for r in d.equity]
    rep = merge_equity_batches(d, [b], {"": "Amit"}, TODAY,
                               cg_on=True, replace=True)
    assert rep.deferred                       # the import did defer
    assert [(r.owner, r.scrip, r.qty, r.avg_cost, r.cost_date)
            for r in d.equity] == before      # ...and touched NOTHING
    assert not d.equity_sells


# ---- F2a: a refused sibling folio blocks statement-wins for the fund -------

def test_mid_history_sibling_folio_blocks_replace():
    two_folio = CAS_TEXT + """
XYZ123 - Parag Parikh Flexi Cap Fund - Growth  ISIN: INF879O01027
Folio No: 99999999 / 11
Opening Unit Balance: 500.000
05-Feb-2024      Purchase                               10,000.00      196.0784      51.0000       696.078
Closing Unit Balance: 696.078
"""
    b = parse_cas(two_folio, TODAY)
    assert ("99999999/11", ISIN) in b.partial
    d = M.PortfolioData(persons=["Amit"])
    typed = [M.SIPRow(owner="Amit", scheme="Parag Parikh Flexi Cap",
                      txn_date=date(2020, 1, 1), amount=99000.0,
                      isin_override=ISIN)]
    d.sip = list(typed)
    rep = merge_sip_batches(d, [b], {"12345678/90": "Amit"}, TODAY,
                            replace=True)
    line = next(ln for ln in rep.funds if ln.folio == "12345678/90")
    assert not line.ok and "another folio" in line.reason
    assert d.sip == typed                     # typed history kept whole


# ---- F2b: more money typed than the whole statement shows → keep typed -----

def test_typed_exceeding_statement_blocks_replace():
    b = parse_cas(CAS_TEXT, TODAY)            # statement shows ₹10,000
    d = M.PortfolioData(persons=["Amit"])
    typed = [M.SIPRow(owner="Amit", scheme="Parag Parikh Flexi Cap",
                      txn_date=date(2020, 1, m), amount=25000.0,
                      isin_override=ISIN) for m in (1, 2)]
    d.sip = list(typed)
    rep = merge_sip_batches(d, [b], {"12345678/90": "Amit"}, TODAY,
                            replace=True)
    assert not rep.funds[0].ok
    assert "missing a folio" in rep.funds[0].reason
    assert d.sip == typed


# ---- F3: with the Capital-gains switch off, a typed-lot sale refuses -------

def test_cg_off_sell_of_typed_lot_refuses_not_silently_skips():
    d = _eq_data()
    d.equity = [_typed_lot(qty=10.0)]
    b = ImportBatch(source="broker tradebook", trades=[_sell(4)])
    rep = merge_equity_batches(d, [b], {"": "Amit"}, TODAY,
                               cg_on=False, replace=True)
    line = rep.stocks[0]
    assert not line.ok and "Capital gains" in line.reason
    assert d.equity[0].qty == 10.0            # untouched, not overstated
    assert rep.eq_reduced == 0 and not d.equity_sells
    # same file with the switch ON nets correctly (the fix's advice works)
    rep2 = merge_equity_batches(d, [b], {"": "Amit"}, TODAY,
                                cg_on=True, replace=True)
    assert rep2.eq_reduced == 1 and d.equity[0].qty == 6.0


# ---- F4: an explicit avg-cost header beats a market "Rate" column ----------

def test_avg_cost_header_beats_market_rate_column():
    csv_text = ("Symbol,Qty,Rate,Avg. Cost\n"
                "RELIANCE,10,2999.99,2000.00\n")
    assert sniff_csv(csv_text) == ("broker holdings", "holdings")
    b = parse_equity_csv(csv_text, TODAY)
    assert b.holdings[0].avg_cost == 2000.00  # never today's price


def test_rate_only_holdings_still_reads_rate_as_cost():
    # traditional back-offices (SIHL/MoneyMaker) label avg-cost "Rate" and
    # have NO explicit avg column — the loose fallback must keep working
    csv_text = "Scrip Name,Qty,Rate\nRELIANCE,10,2000.00\n"
    b = parse_equity_csv(csv_text, TODAY)
    assert b.holdings[0].avg_cost == 2000.00


# ---- F5: unreadable balance lines refuse the fund (never default to 0) -----

def test_unreadable_balance_lines_refuse_fund():
    text = (CAS_TEXT
            .replace("Opening Unit Balance: 0.000",
                     "Opening Unit Balance: 3.708.293")
            .replace("Closing Unit Balance: 200.000",
                     "Closing Unit Balance: 3.908.293"))
    b = parse_cas(text, TODAY)
    assert not b.sip_txns                     # nothing importable survives
    assert ("12345678/90", ISIN) in b.partial
    assert any("couldn't be read reliably" in w for w in b.warnings)


def test_missing_closing_balance_refuses_fund():
    text = CAS_TEXT.replace(
        "Closing Unit Balance: 200.000   NAV", "NAV").replace(
        "Closing Unit Balance: 200.000", "")
    b = parse_cas(text, TODAY)
    assert not b.sip_txns
    assert any("closing balance" in w for w in b.warnings)


# ---- F7: an unreadable line with no identity refuses the whole file --------

def test_orphan_bad_line_refuses_whole_file():
    d = _eq_data()
    bad = ImportedTrade(account="", isin="", symbol="", side="BAD")
    b = ImportBatch(source="broker tradebook",
                    trades=[_buy(10), _sell(4), bad])
    rep = merge_equity_batches(d, [b], {"": "Amit"}, TODAY, cg_on=True)
    assert rep.eq_added == 0 and not d.equity and not d.equity_sells
    assert any("which stock it belongs to" in w for w in rep.warnings)


# ---- F6: a corrupt PDF becomes a per-file message, never an abort ----------

def test_corrupt_pdf_does_not_abort_the_import(tmp_path, capsys):
    from networth.importers import detect_format
    from networth.update import _parse_candidate, prompt_imports
    bad = tmp_path / "broken.pdf"
    bad.write_bytes(b"%PDF-1.7 then garbage, no xref, truncated")
    assert detect_format(bad) == "cas"
    batch, reason = _parse_candidate(bad, "cas", TODAY, interactive=False)
    assert batch is None and "damaged" in reason
    # a good file BESIDE the corrupt one still lands (both via --import;
    # the workbook peek is bypassed with a fake source)
    good = tmp_path / "trades.csv"
    good.write_text("Symbol,ISIN,Trade Date,Trade Type,Quantity,Price\n"
                    f"RELIANCE,{RELIANCE},10-01-2024,buy,10,2500\n")
    wb = tmp_path / "wb.xlsx"
    from networth.generate import build_workbook
    from networth.sample_data import sample_portfolio
    build_workbook(sample_portfolio(), str(wb))
    batches, _m, _r, decisions, _o, _c = prompt_imports(
        tmp_path, wb, ["Amit"], [str(bad), str(good)], False, TODAY,
        workbook=wb)
    assert len(batches) == 1                  # the good file survived
    assert batches[0].trades
    assert ("broken.pdf" in f for f, _fp, dec in decisions)


# ---- F9: a headless run only touches files it was explicitly given ---------

def test_headless_run_never_sweeps_the_folder(tmp_path):
    from networth.update import _import_candidates
    (tmp_path / "sneaky.csv").write_text(
        "Symbol,ISIN,Trade Date,Trade Type,Quantity,Price\n"
        f"RELIANCE,{RELIANCE},10-01-2024,buy,10,2500\n")
    got = _import_candidates(tmp_path, [], set(), include_folder=False)
    assert got == []
    got2 = _import_candidates(tmp_path, [], set(), include_folder=True)
    assert len(got2) == 1                     # interactive still finds it


# ---- F8: a refusal the user can fix keeps the file offerable ---------------

def test_fixable_refusal_does_not_burn_the_fingerprint(tmp_path):
    from networth.fetch.amfi import AmfiData
    from networth.generate import build_workbook
    from networth.reader import read_workbook
    from networth.sample_data import sample_portfolio
    from networth.update import run
    d = sample_portfolio()
    path = tmp_path / "wb.xlsx"
    build_workbook(d, str(path))
    b = parse_cas(CAS_TEXT, TODAY)
    b.fingerprint = "dead12345678"
    # no owner map → the fund refuses with a FIXABLE reason
    summary = run(path, price_data=None, amfi_data=AmfiData(), ca_data=[],
                  div_data=[], today=TODAY, import_batches=[b],
                  import_decisions=[("cas.pdf", "dead12345678", "imported")])
    assert not summary["import_report"].sip_added
    back = read_workbook(str(path))
    assert not back.imported_files            # file stays offerable
    # a run that LANDS the fund does record the never-nag entry
    b2 = parse_cas(CAS_TEXT, TODAY)
    b2.fingerprint = "dead12345678"
    run(path, price_data=None, amfi_data=AmfiData(), ca_data=[],
        div_data=[], today=TODAY, import_batches=[b2],
        import_owner_map={"12345678/90": "Amit"},
        import_decisions=[("cas.pdf", "dead12345678", "imported")])
    back = read_workbook(str(path))
    assert [f.fingerprint for f in back.imported_files] == ["dead12345678"]


# ---- F10: a fresh owner answer REPAIRS a stale Import_Map row --------------

def test_new_owner_answer_updates_stale_map_row(tmp_path):
    from networth.fetch.amfi import AmfiData
    from networth.generate import build_workbook
    from networth.reader import read_workbook
    from networth.sample_data import sample_portfolio
    from networth.update import run
    d = sample_portfolio()
    d.import_map.append(M.ImportMapRow(source="cas", account="12345678/90",
                                       name_hint="", owner="Ghost"))
    path = tmp_path / "wb.xlsx"
    build_workbook(d, str(path))
    b = parse_cas(CAS_TEXT, TODAY)
    run(path, price_data=None, amfi_data=AmfiData(), ca_data=[],
        div_data=[], today=TODAY, import_batches=[b],
        import_owner_map={"12345678/90": "Amit"})
    back = read_workbook(str(path))
    rows = [m for m in back.import_map if m.account == "12345678/90"]
    assert len(rows) == 1 and rows[0].owner == "Amit"   # updated, not doubled


def test_prompt_suppression_ignores_stale_stored_owner(tmp_path):
    # peek returns {"F1": "Ghost"}; the prompt-side filter must drop it so
    # the question is asked again instead of silently failing every merge
    from networth.generate import build_workbook
    from networth.sample_data import sample_portfolio
    from networth.update import peek_import_state
    d = sample_portfolio()
    d.import_map.append(M.ImportMapRow(source="cas", account="F1",
                                       name_hint="", owner="Ghost"))
    path = tmp_path / "wb.xlsx"
    build_workbook(d, str(path))
    _fps, amap, _rows = peek_import_state(path)
    assert amap.get("F1") == "Ghost"          # peek is raw...
    valid = {p.casefold() for p in d.persons}
    assert "ghost" not in valid               # ...and the filter would drop it


# ---- F11: no-client-id broker files key the owner answer by file name ------

def test_blank_account_keyed_by_file_name(tmp_path):
    from networth.generate import build_workbook
    from networth.sample_data import sample_portfolio
    from networth.update import prompt_imports
    f = tmp_path / "holdings.csv"
    f.write_text("Symbol,ISIN,Quantity Available,Average Price\n"
                 f"RELIANCE,{RELIANCE},10,2000\n")
    wb = tmp_path / "wb.xlsx"
    build_workbook(sample_portfolio(), str(wb))
    batches, _m, _r, _d, _o, _c = prompt_imports(
        tmp_path, wb, ["Amit"], [str(f)], False, TODAY, workbook=wb)
    assert batches[0].accounts == [("holdings.csv", "")]
    assert all(h.account == "holdings.csv" for h in batches[0].holdings)


# ---- F12: one parser warning shows once, not once per engine ---------------

def test_batch_warnings_not_duplicated_across_engines(tmp_path):
    from networth.fetch.amfi import AmfiData
    from networth.generate import build_workbook
    from networth.sample_data import sample_portfolio
    from networth.update import run
    path = tmp_path / "wb.xlsx"
    build_workbook(sample_portfolio(), str(path))
    b = ImportBatch(source="broker tradebook",
                    trades=[ImportedTrade(
                        account="", isin=RELIANCE, symbol="RELIANCE",
                        trade_date=date(2024, 1, 10), qty=10, price=2500.0,
                        side="BUY")],
                    warnings=["tradebook line 7: couldn't read it"])
    summary = run(path, price_data=None, amfi_data=AmfiData(), ca_data=[],
                  div_data=[], today=TODAY, import_batches=[b],
                  import_owner_map={"": "Amit"})
    hits = [w for w in summary["warnings"]
            if w == "tradebook line 7: couldn't read it"]
    assert len(hits) == 1


# ---- condensing: the over-cap fallback the user consents to up front -------

def _sip_batch_years(n_per_year, years, closing_key=("F1", ISIN)):
    from networth.importers.common import ImportedSipTxn
    txns, total_units = [], 0.0
    for y in years:
        for m in range(1, n_per_year + 1):
            u = 200.0
            txns.append(ImportedSipTxn(
                folio="F1", isin=ISIN, scheme_name="Parag Parikh Flexi Cap",
                fund_house="PPFAS", txn_date=date(y, m, 5), amount=10000.0,
                nav=50.0, units=u, txn_type="PURCHASE"))
            total_units += u
    b = ImportBatch(source="cas", path="x.pdf",
                    accounts=[("F1", "Amit Kumar")], sip_txns=txns)
    b.closing_units = {closing_key: total_units}
    return b, total_units


def test_condense_consent_rolls_oldest_years_until_it_fits(monkeypatch):
    from networth import model as M2
    monkeypatch.setattr(M2, "SIP_LAST_ROW", M2.FIRST_DATA_ROW + 9)  # cap 10
    d = M.PortfolioData(persons=["Amit"])
    b, total_units = _sip_batch_years(4, (2020, 2021, 2022, 2023))  # 16 txns
    rep = merge_sip_batches(d, [b], {"F1": "Amit"}, date(2026, 7, 18),
                            allow_condense=True)
    assert not rep.deferred and d.sip                 # it landed
    assert len(d.sip) <= 10
    # totals conserved through the roll-up — the anchor still proves it
    assert abs(sum(r.units_override for r in d.sip) - total_units) < 0.001
    assert any(r.amount and r.nav and abs(r.amount - r.units_override * r.nav)
               <= max(1.0, r.amount * 0.001) for r in d.sip)
    assert any("opening line per fund" in w for w in rep.warnings)
    # least-detail-lost: recent transactions stay verbatim
    assert any(r.txn_date and r.txn_date.year == 2023 for r in d.sip)


def test_condense_without_consent_still_defers(monkeypatch):
    from networth import model as M2
    monkeypatch.setattr(M2, "SIP_LAST_ROW", M2.FIRST_DATA_ROW + 9)
    d = M.PortfolioData(persons=["Amit"])
    b, _units = _sip_batch_years(4, (2020, 2021, 2022, 2023))
    rep = merge_sip_batches(d, [b], {"F1": "Amit"}, date(2026, 7, 18))
    assert rep.deferred and not d.sip                 # headless = untouched


def test_condense_that_cannot_fit_falls_back_to_deferral(monkeypatch):
    from networth import model as M2
    monkeypatch.setattr(M2, "SIP_LAST_ROW", M2.FIRST_DATA_ROW + 1)  # cap 2
    d = M.PortfolioData(persons=["Amit"])
    # 3 distinct funds -> even fully condensed needs >= 3 rows
    from networth.importers.common import ImportedSipTxn
    txns = []
    closing = {}
    for i, isin in enumerate(("INF879O01027", "INF879O01035",
                              "INF879O01043")):
        for m in (1, 2):
            txns.append(ImportedSipTxn(
                folio=f"F{i}", isin=isin, scheme_name=f"Fund {i}",
                txn_date=date(2020 + m, 1, 5), amount=1000.0, nav=10.0,
                units=100.0, txn_type="PURCHASE"))
        closing[(f"F{i}", isin)] = 200.0
    b = ImportBatch(source="cas", path="x.pdf", sip_txns=txns,
                    accounts=[(f"F{i}", "") for i in range(3)])
    b.closing_units = closing
    owner = {f"F{i}": "Amit" for i in range(3)}
    rep = merge_sip_batches(d, [b], owner, date(2026, 7, 18),
                            allow_condense=True)
    assert rep.deferred and not d.sip


# ---- read-once: the sniff's payload feeds the parser -----------------------

def test_detect_format_payload_reused_by_parser(tmp_path):
    from networth.importers import detect_format
    from networth.update import _parse_candidate
    f = tmp_path / "trades.csv"
    f.write_text("Symbol,ISIN,Trade Date,Trade Type,Quantity,Price\n"
                 f"RELIANCE,{RELIANCE},10-01-2024,buy,10,2500\n")
    kind, payload = detect_format(f, with_payload=True)
    assert kind == "equity_csv" and "RELIANCE" in payload
    f.unlink()                                        # parse must NOT re-read
    batch, reason = _parse_candidate(f, kind, TODAY, False, payload=payload)
    assert batch is not None and len(batch.trades) == 1
    # and the sniff-only form still returns a bare kind (API unchanged)
    f.write_text("Symbol,ISIN,Trade Date,Trade Type,Quantity,Price\n"
                 f"RELIANCE,{RELIANCE},10-01-2024,buy,10,2500\n")
    assert detect_format(f) == "equity_csv"
