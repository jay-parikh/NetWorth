"""v1.7.5 — ETFs on the Equity sheet, and taxed the way they really are.

From an end-user report (2026-08-10): a silver ETF's price never updated and
the ETF wasn't in Stock_Master. Fixing that surfaced a bigger problem — every
Equity row was taxed as EQUITY, so a gold/silver/debt ETF wrongly received the
₹1.25L §112A allowance that belongs to shares and equity funds.

Covered here:
  1. ETFs reach Stock_Master from the AMFI master, whatever the exchanges did
  2. an ETF with no exchange quote falls back to its NAV; a real quote wins
  3. the Equity 'Tax type' column routes rows to the right tax bucket
  4. no realised rupee can vanish from the FY summary (the invariant that
     makes adding a bucket safe)
"""

from datetime import date

import pytest

from networth import model as M
from networth.compute.capital_gains import (NON_EQUITY_BUCKETS,
                                            capital_gains_report)
from networth.fetch.amfi import AmfiData
from networth.fetch.bhavcopy import PriceData
from networth.generate import build_workbook
from networth.model import equity_tax_bucket, load_tax_rules, tax_rule_for
from networth.reader import read_workbook
from networth.sample_data import sample_portfolio
from networth.update import _etf_master_rows, run

TODAY = date(2026, 8, 10)
SILVER = "INF204KC1402"          # Nippon India Silver ETF (NSE-only)
NIFTY = "INF204KB14I2"
RELIANCE = "INE002A01018"


# ---- 1: ETFs always reachable on the Equity sheet ---------------------------

def test_etf_rows_derived_from_the_amfi_master():
    mf = [("Nippon", "Nippon India ETF Silver BeES", SILVER),
          ("SBI", "SBI Nifty 50 ETF", "INF200KA1FS1"),
          ("Motilal", "Motilal Oswal NASDAQ 100 Exchange Traded Fund",
           "INF247L01AT9"),
          ("Kotak", "Kotak Gold ETF-Growth", "INF174K01LS6"),
          ("HDFC", "HDFC Top 100 Fund", "INF179K01BE2"),      # plain fund
          ("Mirae", "Mirae Asset Gold ETF Fund of Fund", "INF769K01OS1"),
          ("X", "Weird ETFO Fund", "INF999X01011"),           # not an ETF
          ("Y", "No ISIN ETF", "")]
    got = {isin for _s, _n, isin in _etf_master_rows(mf)}
    assert got == {SILVER, "INF200KA1FS1", "INF247L01AT9", "INF174K01LS6"}
    # a FUND OF an ETF is not traded — it must never reach the equity dropdown
    assert "INF769K01OS1" not in got


def test_etf_lands_in_stock_master_on_a_bse_only_day(tmp_path):
    d = sample_portfolio()
    d.equity.append(M.EquityRow(owner="Amit",
                                scrip="Nippon India ETF Silver BeES",
                                qty=100, avg_cost=95.0,
                                cost_date=date(2025, 5, 5)))
    path = tmp_path / "wb.xlsx"
    build_workbook(d, str(path))
    px = PriceData(prices={RELIANCE: {"close": 1500.0, "prev": 1490.0}},
                   master_rows=[("RELIANCE", "RELIANCE INDUSTRIES LTD.",
                                 RELIANCE)],
                   trade_date=TODAY, source="BSE", sources=["BSE"])
    amfi = AmfiData(nav_by_isin={SILVER: 104.5},
                    master_rows=[("Nippon India MF",
                                  "Nippon India ETF Silver BeES", SILVER)])
    s = run(path, price_data=px, amfi_data=amfi, ca_data=[], div_data=[],
            today=TODAY)
    back = read_workbook(str(path))
    assert SILVER in {isin for _s, _n, isin in back.masters.stock_rows}
    # ...and the row's name now resolves to that ISIN, so it prices from here on
    by_name = {n: i for _s, n, i in back.masters.stock_rows}
    assert by_name["Nippon India ETF Silver BeES"] == SILVER
    # an ETF we cannot query (no exchange symbol) is never nagged about
    assert "Nippon India ETF Silver BeES" not in s.get("ca_unverified", [])


def test_nav_fills_in_when_no_exchange_quotes_it_and_a_quote_wins(tmp_path):
    d = sample_portfolio()
    d.equity.append(M.EquityRow(owner="Amit",
                                scrip="Nippon India ETF Silver BeES",
                                qty=100, avg_cost=95.0,
                                cost_date=date(2025, 5, 5)))
    path = tmp_path / "wb.xlsx"
    build_workbook(d, str(path))
    px = PriceData(prices={RELIANCE: {"close": 1500.0, "prev": 1490.0}},
                   master_rows=[("RELIANCE", "RELIANCE INDUSTRIES LTD.",
                                 RELIANCE)],
                   trade_date=TODAY, source="BSE", sources=["BSE"])
    amfi = AmfiData(nav_by_isin={SILVER: 104.5},
                    master_rows=[("Nippon India MF",
                                  "Nippon India ETF Silver BeES", SILVER)])
    s1 = run(path, price_data=px, amfi_data=amfi, ca_data=[], div_data=[],
             today=TODAY)
    assert s1.get("nav_priced") == 1
    etf = next(r for r in read_workbook(str(path)).equity
               if "Silver" in r.scrip)
    assert etf.close == pytest.approx(104.5)

    # the exchange starts quoting it: the market price must win outright
    px.prices[SILVER] = {"close": 101.2, "prev": 100.9}
    px.master_rows.append(("SILVERBEES", "NIPPON INDIA ETF SILVER BEES",
                           SILVER))
    s2 = run(path, price_data=px, amfi_data=amfi, ca_data=[], div_data=[],
             today=TODAY)
    assert not s2.get("nav_priced")
    etf2 = next(r for r in read_workbook(str(path)).equity
                if "Silver" in r.scrip)
    assert etf2.close == pytest.approx(101.2)


# ---- 2: the Tax type column -------------------------------------------------

def test_tax_type_words_map_to_buckets():
    assert equity_tax_bucket("") == "equity"
    assert equity_tax_bucket("Equity") == "equity"
    for word in ("Gold-Silver", "gold", "SILVER", "bullion", "Overseas",
                 "international"):
        assert equity_tax_bucket(word) == "mf_other", word
    for word in ("Debt", "bond", "DEBT "):
        assert equity_tax_bucket(word) == "mf_debt", word
    # anything unrecognised falls back to what the sheet always did
    assert equity_tax_bucket("something else") == "equity"


def test_bundled_rules_cover_the_new_bucket():
    rules = load_tax_rules()
    r = tax_rule_for(rules, "mf_other", date(2026, 6, 1))
    assert r is not None
    assert r.lt_days == 365            # listed units: long-term after a year
    assert r.ltcg_pct == pytest.approx(12.5)
    assert r.ltcg_exempt == 0          # the 1.25L allowance is equity-only
    assert r.stcg_pct is None          # short-term = at your slab


def test_tax_type_roundtrips_through_the_workbook(tmp_path):
    d = sample_portfolio()
    d.equity.append(M.EquityRow(owner="Amit", scrip="Silver ETF",
                                isin_override=SILVER, qty=10, avg_cost=90.0,
                                cost_date=date(2025, 1, 1),
                                tax_type="Gold-Silver", manual_price=99.0))
    path = tmp_path / "wb.xlsx"
    build_workbook(d, str(path))
    back = read_workbook(str(path))
    row = next(r for r in back.equity if r.isin_override == SILVER)
    assert row.tax_type == "Gold-Silver"
    assert row.manual_price == pytest.approx(99.0)


def _two_etfs(tax_type_for_silver="Gold-Silver"):
    d = M.PortfolioData(persons=["Jay"])
    d.masters.stock_rows = [("", "Silver ETF", SILVER), ("", "Nifty ETF", NIFTY)]
    d.show_capital_gains = True
    for name, tt in (("Silver ETF", tax_type_for_silver), ("Nifty ETF", "")):
        d.equity.append(M.EquityRow(owner="Jay", scrip=name, qty=1000,
                                    avg_cost=100.0, close=150.0,
                                    cost_date=date(2024, 8, 1), tax_type=tt))
        d.equity_sells.append(M.EquitySellRow(
            owner="Jay", scrip=name, isin_override="", qty=500,
            buy_date=date(2024, 8, 1), buy_price=100.0,
            sell_date=date(2026, 6, 1), sell_price=150.0))
    return d


def test_gold_silver_etf_keeps_out_of_the_equity_allowance():
    rep = capital_gains_report(_two_etfs(), TODAY)
    buckets = {r.name: r.bucket for r in rep.realised}
    assert buckets["Silver ETF"] == "mf_other"
    assert buckets["Nifty ETF"] == "equity"
    s = next(x for x in rep.summaries if x.fy == "2026-27")
    # only the Nifty ETF's 25,000 may consume the 1.25L allowance
    assert s.ltcg == pytest.approx(25_000)
    assert s.debt_gain == pytest.approx(25_000)
    assert s.exemption_used == pytest.approx(25_000)
    # unrealised too — and the sale inherited the holding's Tax type
    un = {r.name: r.bucket for r in rep.unrealised}
    assert un["Silver ETF"] == "mf_other" and un["Nifty ETF"] == "equity"


def test_without_the_marker_it_behaves_exactly_as_before():
    rep = capital_gains_report(_two_etfs(tax_type_for_silver=""), TODAY)
    s = next(x for x in rep.summaries if x.fy == "2026-27")
    assert s.ltcg == pytest.approx(50_000)     # both in the equity family
    assert s.debt_gain == pytest.approx(0)


def test_grandfathering_is_equity_only():
    """The 31-01-2018 relief is a §112A benefit — a bullion ETF must not
    get it even when the lot predates the cutoff."""
    d = M.PortfolioData(persons=["Jay"])
    d.masters.stock_rows = [("", "Silver ETF", SILVER)]
    d.show_capital_gains = True
    d.equity.append(M.EquityRow(owner="Jay", scrip="Silver ETF", qty=100,
                                avg_cost=10.0, close=200.0,
                                cost_date=date(2015, 1, 1),
                                tax_type="Gold-Silver"))
    rep = capital_gains_report(d, TODAY)
    row = next(r for r in rep.unrealised if r.name == "Silver ETF")
    assert row.bucket == "mf_other"
    assert "grandfathered" not in row.note
    assert row.gain_today == pytest.approx(100 * (200.0 - 10.0))


def test_debt_marked_lot_after_2023_is_slab_taxed():
    d = M.PortfolioData(persons=["Jay"])
    d.masters.stock_rows = [("", "Bond ETF", "INF204KB17I5")]
    d.show_capital_gains = True
    d.equity.append(M.EquityRow(owner="Jay", scrip="Bond ETF", qty=100,
                                avg_cost=100.0, close=120.0,
                                cost_date=date(2024, 5, 1), tax_type="Debt"))
    rep = capital_gains_report(d, TODAY)
    row = next(r for r in rep.unrealised if r.name == "Bond ETF")
    assert row.bucket == "slab" and row.term == "At your slab"
    assert row.lt_on is None


# ---- 3: the invariant that makes a new bucket safe --------------------------

def test_no_realised_gain_can_vanish_from_the_fy_summary():
    """Every realised row must land in exactly one summary figure. Without
    this, adding a bucket could silently drop a gain from the totals."""
    d = _two_etfs()
    d.masters.stock_rows.append(("", "Bond ETF", "INF204KB17I5"))
    d.equity.append(M.EquityRow(owner="Jay", scrip="Bond ETF", qty=100,
                                avg_cost=100.0, close=90.0,
                                cost_date=date(2024, 5, 1), tax_type="Debt"))
    d.equity_sells.append(M.EquitySellRow(
        owner="Jay", scrip="Bond ETF", isin_override="", qty=100,
        buy_date=date(2024, 5, 1), buy_price=100.0,
        sell_date=date(2026, 6, 1), sell_price=90.0))
    rep = capital_gains_report(d, TODAY)
    for s in rep.summaries:
        rows = [r for r in rep.realised if r.fy == s.fy]
        counted = (s.stcg + s.ltcg + s.slab_gain + s.debt_gain + s.spec_gain)
        assert counted == pytest.approx(sum(r.gain for r in rows), abs=0.01), (
            f"FY {s.fy}: rows sum to {sum(r.gain for r in rows)} but the "
            f"summary figures account for {counted}")


def test_every_bucket_a_row_can_carry_is_summarised():
    """A guard for the NEXT bucket someone adds: each one must be reachable
    from an equity Tax type AND be counted somewhere in the summary."""
    reachable = {equity_tax_bucket(t) for t in M.EQUITY_TAX_TYPES}
    assert reachable == {"equity", "mf_debt", "mf_other"}
    # equity is the §112A family; the rest must be in the non-equity figure
    # (mf_debt can also become "slab", which has its own figure)
    assert set(NON_EQUITY_BUCKETS) >= reachable - {"equity"}
