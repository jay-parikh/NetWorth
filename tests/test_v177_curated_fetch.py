"""v1.7.7 — curated merger/demerger data arrives without a new release.

Splits and bonuses come from the exchanges; mergers and demergers do not
(no feed publishes a swap ratio, and the cost apportionment lives only in
a company filing). Until now the only way to deliver one was a whole app
release, so a user who hit a demerger waited for a version.

The updater now refreshes the curated list from the project on every run:
an event added today reaches every user on their next Update Portfolio,
with nothing for them to do. The bundled file stays the offline baseline,
and a bad download can never replace it.
"""

from datetime import date

import pytest

from networth.fetch import curated as C
from networth.model import (load_restructures, parse_restructures,
                            restructure_key)

HDR = ("ex_date,type,old_isin,old_name,old_symbol,new_isin,new_name,"
       "new_symbol,ratio_from,ratio_to,cost_pct,details\n")
NEW = ("2026-08-01,DEMERGER,INE999Z01011,NEWCO LTD,NEWCO,INE999Z01011,"
       "NEWCO LTD,NEWCO,1,1,70,parent keeps 70%\n"
       "2026-08-01,DEMERGER,INE999Z01011,NEWCO LTD,NEWCO,INE888Y01012,"
       "SPUNOFF LTD,SPUN,1,1,30,child takes 30%\n")


_REAL_FETCH = C.fetch_restructures


@pytest.fixture(autouse=True)
def _use_the_real_fetcher(monkeypatch):
    """conftest stubs the curated fetch off for the whole suite (so no test
    waits on the network). THIS module is the one that exercises it, always
    through a stub session, so put the real function back."""
    monkeypatch.setattr(C, "fetch_restructures", _REAL_FETCH)


class _Resp:
    def __init__(self, text, code=200):
        self.text, self.status_code = text, code


class _Sess:
    """A stub that answers exactly once, like requests.Session.get."""
    def __init__(self, text, code=200):
        self._r = _Resp(text, code)
        self.calls = 0

    def get(self, *a, **k):
        self.calls += 1
        return self._r


def _bundled_text():
    from networth.model import DATA_DIR
    return (DATA_DIR / "restructures.csv").read_text(encoding="utf-8")


# ---- the point of the release: new events, no new app version -------------

def test_a_newly_published_event_reaches_the_user():
    bundled = load_restructures()
    got = C.fetch_restructures(session=_Sess(_bundled_text() + NEW))
    merged, added = C.merge_restructures(bundled, got)
    assert added == 2
    assert len(merged) == len(bundled) + 2
    keys = {restructure_key(a) for a in merged}
    assert ("INE999Z01011", "DEMERGER", date(2026, 8, 1),
            "INE888Y01012") in keys
    # and every event the app already shipped is still there
    assert {restructure_key(a) for a in bundled} <= keys


def test_a_moved_child_replaces_the_event_and_never_over_apportions():
    """The review's case: an upstream correction moves a demerger child to
    a different ISIN. A row-by-row union would keep BOTH children — 130%
    of the cost apportioned and a phantom holding appended. Events swap
    whole, so the merged list still sums to exactly 100."""
    from networth.model import validate_restructures
    bundled = parse_restructures(HDR + NEW)
    moved = parse_restructures(
        HDR + NEW.replace("INE888Y01012", "INE777X01013"))
    merged, changed = C.merge_restructures(bundled, moved)

    assert len(merged) == 2                      # not 3
    assert {a.new_isin for a in merged} == {"INE999Z01011", "INE777X01013"}
    validate_restructures(merged)                # must not raise
    assert sum(a.cost_pct for a in merged) == pytest.approx(100.0)
    assert changed  # the summary must not claim nothing happened


def test_a_merge_that_would_break_the_hundred_rule_is_refused():
    """Two individually-valid lists can still merge into an invalid one;
    when that happens the bundled list stands."""
    bundled = parse_restructures(HDR + NEW)
    # same event, but only the child row is republished (retention dropped)
    partial = parse_restructures(
        HDR + "2026-08-01,DEMERGER,INE999Z01011,NEWCO LTD,NEWCO,"
              "INE888Y01012,SPUNOFF LTD,SPUN,1,1,100,child only\n")
    merged, changed = C.merge_restructures(bundled, partial)
    # the replacement IS coherent on its own (100 in one row), so it stands
    assert sum(a.cost_pct for a in merged) == pytest.approx(100.0)

    # but a genuinely incoherent replacement is refused wholesale
    broken = list(parse_restructures(HDR + NEW))
    broken[1].cost_pct = 55.0                    # 70 + 55 = 125
    assert C.merge_restructures(bundled, broken) == (bundled, 0)


def test_a_corrected_percentage_supersedes_the_bundled_row():
    bundled = parse_restructures(HDR + NEW)
    fixed = parse_restructures(HDR + NEW.replace(",1,1,70,", ",1,1,60,")
                               .replace(",1,1,30,", ",1,1,40,"))
    merged, added = C.merge_restructures(bundled, fixed)
    assert added == 2 and len(merged) == 2      # replaced, not duplicated
    by_new = {a.new_isin: a.cost_pct for a in merged}
    assert by_new["INE999Z01011"] == pytest.approx(60)
    assert by_new["INE888Y01012"] == pytest.approx(40)


# ---- a bad download must never cost the user data -------------------------

def test_offline_or_error_keeps_the_bundled_copy():
    bundled = load_restructures()
    for sess in (_Sess("", 599), _Sess("", 404), _Sess("<html>nope</html>")):
        assert not C.fetch_restructures(session=sess)
        assert C.refresh_restructures(bundled, session=sess) == (bundled, 0)


def test_a_raised_exception_is_swallowed():
    class Boom:
        def get(self, *a, **k):
            raise RuntimeError("proxy exploded")
    assert C.fetch_restructures(session=Boom()) is None
    assert C.refresh_restructures([], session=Boom()) == ([], 0)


def test_truncated_or_short_answer_is_refused():
    assert not C.fetch_restructures(session=_Sess(HDR))          # no rows
    assert not C.fetch_restructures(session=_Sess(HDR + NEW[:40]))


def test_fewer_events_than_the_app_ships_is_refused():
    bundled = load_restructures()
    # the floor is inside refresh_restructures, so no caller can skip it
    assert C.refresh_restructures(bundled,
                                  session=_Sess(HDR + NEW)) == (bundled, 0)


def test_a_bad_cost_split_is_refused_wholesale():
    """The shared parser's 100% rule guards the fetched file too — a
    downloaded row can never be trusted on weaker terms than a shipped one."""
    bad = HDR + NEW.replace(",1,1,30,", ",1,1,20,")
    assert C.fetch_restructures(session=_Sess(bad)) is None
    with pytest.raises(ValueError, match="sums to"):
        parse_restructures(HDR + NEW.replace(",1,1,30,", ",1,1,20,"))


def test_the_bundled_file_is_parsed_by_the_same_function():
    from networth.model import DATA_DIR
    assert load_restructures() == parse_restructures(
        (DATA_DIR / "restructures.csv").read_text(encoding="utf-8"))


# ---- wiring: the run must never break, and can be switched off ------------

def test_run_survives_a_dead_network_and_can_opt_out(tmp_path, monkeypatch):
    from networth.fetch.amfi import AmfiData
    from networth.fetch.bhavcopy import PriceData
    from networth.generate import build_workbook
    from networth.sample_data import sample_portfolio
    from networth.update import run

    p = tmp_path / "wb.xlsx"
    build_workbook(sample_portfolio(), str(p))

    def boom(*a, **k):
        raise RuntimeError("no network here")
    monkeypatch.setattr(C, "fetch_restructures", boom)

    # opted out: the fetch is not even attempted
    s = run(p, price_data=PriceData(), amfi_data=AmfiData(), ca_data=[], div_data=[],
            today=date(2026, 8, 11), restructures=load_restructures())
    assert s["ca_rows"] is not None and not s.get("curated_added")


def test_run_uses_a_fetched_event_end_to_end(tmp_path, monkeypatch):
    """The path that matters: the fetch succeeds and a newly published
    demerger reaches the sheet in the SAME run, with no user action."""
    from networth import model as M
    from networth.fetch.amfi import AmfiData
    from networth.fetch.bhavcopy import PriceData
    from networth.generate import build_workbook
    from networth.reader import read_workbook
    from networth.sample_data import sample_portfolio
    from networth.update import run

    d = sample_portfolio()
    held = M.EquityRow(owner="Amit", scrip="RELIANCE INDUSTRIES LTD.",
                       qty=100, avg_cost=900.0, cost_date=date(2020, 1, 1))
    d.equity = [held]
    p = tmp_path / "wb.xlsx"
    build_workbook(d, str(p))

    isin = "INE002A01018"                                   # Reliance
    published = parse_restructures(
        HDR
        + f"2026-06-01,DEMERGER,{isin},RELIANCE,RELIANCE,{isin},"
          "RELIANCE INDUSTRIES LTD.,RELIANCE,1,1,80,parent keeps 80%\n"
        + f"2026-06-01,DEMERGER,{isin},RELIANCE,RELIANCE,INE555Q01019,"
          "NEWSPUN LTD,NEWSPUN,1,1,20,child takes 20%\n")
    monkeypatch.setattr(C, "fetch_restructures",
                        lambda *a, **k: load_restructures() + published)

    s = run(p, price_data=PriceData(), amfi_data=AmfiData(), ca_data=[], div_data=[],
            today=date(2026, 8, 11))
    assert s.get("curated_added") == 2
    back = read_workbook(str(p))
    kids = [r for r in back.equity if r.isin_override == "INE555Q01019"]
    assert len(kids) == 1 and kids[0].qty == pytest.approx(100.0)
    assert kids[0].cost_date == date(2020, 1, 1)      # holding period kept
