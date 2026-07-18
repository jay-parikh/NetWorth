"""Capital-gains engine (v1.6, SPEC §6.16) — realised & unrealised STCG/LTCG.

Everything here is derived from persisted inputs (Equity_Sells, MF_SIP, the
bundled FMV and tax tables, Corporate_Actions), so the report is computed at
BUILD time and never stored — regeneration always reproduces it and the
round-trip identity invariant is untouched.

Honesty rules: figures are INDICATIVE (planning, not filing); when something
can't be computed correctly (nav-less redemption, oversold fund, unknown
scrip) it is skipped with a warning — never guessed. Same-FY set-off follows
Sec 70 across the capital-gains buckets (v1.6.2, in FYs whose rules are
known): the whole short-term head nets first — equity-family, debt-fund
and slab (Sec 50AA) together, so debt/slab ST gains absorb equity ST
losses and vice versa — then a genuine all-ST net loss and debt-fund LT
losses (Sec 70(3)) reduce LTCG, all before the §112A exemption. Every
set-off shows in the set-off column, the console bits or the tax figures —
never folded into the raw gain columns (the sheet subtitle says so).
Leftover losses never widen headroom and never carry forward; other income
heads are never touched.

Units: an Equity_Sells row is in SELL-TIME share units (the contract note's
view) and is never CA-adjusted. Only the bundled per-share FMV of 31-01-2018
is normalised across later corporate actions before use (§6.6).
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, timedelta

from ..model import (
    EQSELL_LAST_ROW, FIRST_DATA_ROW, PortfolioData, TaxRule,
    chained_adjustment_factor, effective_tax_rules, fy_label, load_fmv,
    tax_rule_for,
)

GRANDFATHER_DATE = date(2018, 1, 31)   # FMV valuation day (§6.6)
GRANDFATHER_CUTOFF = date(2018, 2, 1)  # bought before this → grandfathering
DEBT_MF_SLAB_FROM = date(2023, 4, 1)   # debt-MF lots bought on/after → slab
DUST_INR = 1.0                         # FIFO residuals worth less than this are
                                       # float rounding (amounts are paise-
                                       # rounded, units NAV-derived), not
                                       # holdings or oversells


@dataclass
class RealisedRow:
    fy: str = ""
    owner: str = ""
    name: str = ""                     # scrip or scheme
    bucket: str = ""                   # equity | mf_equity | mf_debt | slab
    qty: float | None = None
    buy_date: date | None = None
    sell_date: date | None = None
    held_days: int = 0
    term: str = ""                     # Short-term | Long-term | At your slab
    proceeds: float = 0.0
    taxable_cost: float = 0.0
    gain: float = 0.0
    note: str = ""


@dataclass
class UnrealisedRow:
    owner: str = ""
    name: str = ""
    bucket: str = ""
    qty: float | None = None
    value_today: float = 0.0
    gain_today: float = 0.0            # vs the TAXABLE (grandfathered) cost
    term: str = ""
    lt_on: date | None = None          # the day it turns long-term (blank = already)
    note: str = ""


@dataclass
class FYSummary:
    fy: str = ""
    stcg: float = 0.0                  # equity bucket (equity + mf_equity)
    ltcg: float = 0.0
    exemption: float = 0.0             # §112A allowance for that FY
    exemption_used: float = 0.0
    headroom: float = 0.0
    tax_stcg: float | None = None      # None = not computable (e.g. pre-2018 era)
    tax_ltcg: float | None = None
    slab_gain: float = 0.0             # gains taxed at the user's slab (no amount)
    debt_gain: float = 0.0             # old-regime debt gains (rate era dependent)
    st_setoff: float = 0.0             # total Sec 70 set-off vs this FY's
                                       # LTCG before the §112A exemption
                                       # (§6.16: ST excess + debt LT losses)
    st_sheltered: float = 0.0          # debt/slab ST losses absorbed by the
                                       # equity ST figure (console surface)

    @property
    def ltcg_eff(self) -> float:
        """The long-term figure the exemption and tax apply to: raw LTCG
        after the Sec 70 set-offs, floored at 0. ONE derivation shared by
        the FY row and headroom_now so the two surfaces can never drift."""
        return max(self.ltcg - self.st_setoff, 0.0)
    spec_gain: float = 0.0             # intraday = speculative income (Sec 43(5)),
                                       # slab-taxed business income, NOT capital
                                       # gains — shown so nothing is hidden


@dataclass
class CGReport:
    realised: list[RealisedRow] = field(default_factory=list)
    unrealised: list[UnrealisedRow] = field(default_factory=list)
    summaries: list[FYSummary] = field(default_factory=list)   # newest first
    warnings: list[str] = field(default_factory=list)
    headroom_now: float | None = None  # current-FY LTCG still tax-free
    fy_now: str = ""


def _term(days: int, lt_days: int) -> str:
    return "Long-term" if days > lt_days else "Short-term"


def _mf_units(row) -> float | None:
    if row.units_override is not None:
        return abs(row.units_override)
    if row.amount and row.nav and row.nav > 0:      # a negative NAV is noise
        return abs(row.amount) / row.nav
    return None


def capital_gains_report(data: PortfolioData, today: date,
                         rules: list[TaxRule] | None = None,
                         fmv: tuple[dict, dict] | None = None) -> CGReport:
    rep = CGReport(fy_now=fy_label(today))
    if rules is None:
        # bundled defaults upserted with the workbook's Tax_Rules rows
        # (§3.22) — a Budget change is an Excel edit, not an app release
        rules, _invalid, rule_warnings = effective_tax_rules(data.tax_rules)
        rep.warnings.extend(rule_warnings)
    try:
        fmv_by_isin, fmv_by_symbol = fmv if fmv is not None else load_fmv()
    except (OSError, ValueError):
        fmv_by_isin, fmv_by_symbol = {}, {}

    cap = EQSELL_LAST_ROW - FIRST_DATA_ROW + 1
    if len(data.equity_sells) > cap:
        rep.warnings.append(
            f"Equity_Sells holds {len(data.equity_sells)} sale rows but the "
            f"sheet fits {cap} - the extra rows still count in this run's "
            "figures but are DROPPED from the saved file. Move old years to "
            "another file to keep the record.")

    isin_by_name = {name: isin for _s, name, isin in data.masters.stock_rows}
    symbol_by_isin = {isin: sym for sym, _n, isin in data.masters.stock_rows}
    actions = data.corporate_actions

    def fmv_per_share(isin: str, at: date) -> float | None:
        """31-01-2018 FMV expressed in the share units of `at` (§6.6): the
        bundled value is per 2018 share; splits/bonuses since then multiply
        the count, so the per-share value divides by the same factor."""
        raw = fmv_by_isin.get(isin) or fmv_by_symbol.get(
            symbol_by_isin.get(isin, ""))
        if not raw:
            return None
        f = chained_adjustment_factor(isin, GRANDFATHER_DATE, at, actions)
        return raw / f if f else raw

    # ---- equity realised: self-contained Equity_Sells rows (§3.20) --------
    for s in data.equity_sells:
        what = s.scrip or s.isin_override or "row"
        # a typed 0 price is real data (a delisted write-off, a bonus-share
        # cost), not a missing value — only None means "not entered"
        if not (s.qty and s.sell_date and s.buy_date) or s.sell_price is None:
            if s.owner or s.scrip:     # ignore fully blank lines
                rep.warnings.append(
                    f"Equity_Sells: {what} is missing qty/dates/prices - "
                    "left out of the tax report")
            continue
        if (s.qty < 0 or s.sell_price < 0
                or (s.buy_price is not None and s.buy_price < 0)):
            rep.warnings.append(
                f"Equity_Sells: {what} has a negative quantity or price - "
                "check the row; left out of the tax report")
            continue
        if s.sell_date > today:
            rep.warnings.append(
                f"Equity_Sells: {what} has a sell date in the future - "
                "left out of the tax report")
            continue
        if s.sell_date == s.buy_date:
            # intraday = speculative income (Sec 43(5)): slab-taxed BUSINESS
            # income, not capital gains — shown in its own bucket so nothing
            # is hidden, but never mixed into STCG/LTCG and no tax computed
            if s.buy_price is None:
                rep.warnings.append(
                    f"Equity_Sells: {what} is an intraday trade with no "
                    "buy price - left out of the tax report")
                continue
            rep.realised.append(RealisedRow(
                fy=fy_label(s.sell_date), owner=s.owner,
                name=s.scrip or s.isin_override, bucket="speculative",
                qty=s.qty, buy_date=s.buy_date, sell_date=s.sell_date,
                held_days=0, term="Intraday",
                proceeds=s.qty * s.sell_price,
                taxable_cost=s.qty * s.buy_price,
                gain=s.qty * (s.sell_price - s.buy_price),
                note="speculative income (intraday) - taxed at your slab "
                     "as business income, not as capital gains"))
            continue
        if s.sell_date < s.buy_date:
            rep.warnings.append(
                f"Equity_Sells: {what} shows it was sold before it was "
                "bought - check the two dates; left out for now")
            continue
        isin = s.isin_override or isin_by_name.get(s.scrip, "")
        note = ""
        cost_sh = s.buy_price
        if s.buy_date < GRANDFATHER_CUTOFF:
            f = fmv_per_share(isin, s.sell_date)
            if f is not None:
                # §6.6: taxable cost = higher of actual cost vs
                # min(FMV, sale price) — per share
                cost_sh = max(s.buy_price or 0.0, min(f, s.sell_price))
                note = "grandfathered (31-Jan-2018 value)"
                if s.buy_price is None:
                    note = ("buy price blank - the 31-Jan-2018 market value "
                            "(FMV) was used, an estimate")
            elif cost_sh is None:
                rep.warnings.append(
                    f"Equity_Sells: {what} has no buy price and no "
                    "31-Jan-2018 value was found - left out")
                continue
        elif cost_sh is None:
            rep.warnings.append(
                f"Equity_Sells: {what} needs a buy price (bought after "
                "Jan 2018) - left out")
            continue
        held = (s.sell_date - s.buy_date).days
        rule = tax_rule_for(rules, "equity", s.sell_date)
        lt_days = rule.lt_days if rule else 365
        rep.realised.append(RealisedRow(
            fy=fy_label(s.sell_date), owner=s.owner, name=s.scrip or isin,
            bucket="equity", qty=s.qty, buy_date=s.buy_date,
            sell_date=s.sell_date, held_days=held, term=_term(held, lt_days),
            proceeds=s.qty * s.sell_price, taxable_cost=s.qty * cost_sh,
            gain=s.qty * (s.sell_price - cost_sh), note=note))

    # ---- MF realised: FIFO over the MF_SIP ledger (§6.16) -----------------
    # Tax type is a property of the SCHEME — a fund is equity or debt for
    # every owner. First non-blank row wins; conflicting rows warn instead of
    # silently taking whichever row happens to be read last.
    tax_type_by_scheme: dict[str, str] = {}
    for m in data.mutual_funds:
        tt = (m.tax_type or "").strip().casefold()
        if not (m.scheme and tt):
            continue
        prev = tax_type_by_scheme.setdefault(m.scheme, tt)
        if prev != tt:
            rep.warnings.append(
                f"MutualFunds: {m.scheme} is marked both Equity and Debt on "
                f"different rows - using {prev.capitalize()}; please make "
                "the rows match")
    ledgers: dict[tuple[str, str], list] = {}
    for r in data.sip:
        if (r.owner and r.scheme and r.txn_date and r.amount
                and r.txn_date <= today):   # future rows wait for their date
            ledgers.setdefault((r.owner, r.scheme), []).append(r)

    open_lots: dict[tuple[str, str], list] = {}   # (owner, scheme) → [lot...]
    for key, rows in ledgers.items():
        owner, scheme = key
        rows.sort(key=lambda r: r.txn_date)
        tt = tax_type_by_scheme.get(scheme, "")
        is_debt = tt == "debt"
        assumed = tt == ""             # blank or scheme not on MutualFunds
        lots: list[dict] = []          # {date, units, cost_per_unit}
        for r in rows:
            units = _mf_units(r)
            if r.amount > 0:
                if units:
                    lots.append({"date": r.txn_date, "units": units,
                                 "cpu": r.amount / units})
                else:
                    # v1.6.2: as loud as the redemption twin below — a
                    # silent NAV-less buy vanished from FIFO and XIRR
                    rep.warnings.append(
                        f"MF_SIP: a {scheme} purchase on "
                        f"{r.txn_date:%d-%m-%Y} has no NAV - left out of "
                        "the tax report and the return figure")
                continue
            # redemption row (a zero/negative NAV counts as "no NAV")
            if not units and not (r.nav and r.nav > 0):
                rep.warnings.append(
                    f"MF_SIP: a {scheme} redemption on "
                    f"{r.txn_date:%d-%m-%Y} has no NAV - left out of the "
                    "tax report")
                continue
            sell_nav = (r.nav if r.nav and r.nav > 0
                        else abs(r.amount) / units)
            remaining = units if units else abs(r.amount) / sell_nav
            while remaining * sell_nav > DUST_INR and lots:
                lot = lots[0]
                take = min(lot["units"], remaining)
                held = (r.txn_date - lot["date"]).days
                if is_debt and lot["date"] >= DEBT_MF_SLAB_FROM:
                    bucket, term, rule = "slab", "At your slab", None
                else:
                    bucket = "mf_debt" if is_debt else "mf_equity"
                    rule = tax_rule_for(rules, bucket, r.txn_date)
                    term = _term(held, rule.lt_days if rule else 365)
                note = ""
                if not is_debt and assumed:
                    note = "assumed Equity - set Tax type on MutualFunds"
                if bucket == "mf_equity" and lot["date"] < GRANDFATHER_CUTOFF:
                    note = (note + "; " if note else "") + \
                        "pre-2018 fund - grandfathering not applied, " \
                        "gain may be overstated"
                rep.realised.append(RealisedRow(
                    fy=fy_label(r.txn_date), owner=owner, name=scheme,
                    bucket=bucket, qty=take, buy_date=lot["date"],
                    sell_date=r.txn_date, held_days=held, term=term,
                    proceeds=take * sell_nav, taxable_cost=take * lot["cpu"],
                    gain=take * (sell_nav - lot["cpu"]), note=note))
                lot["units"] -= take
                remaining -= take
                if lot["units"] * sell_nav <= DUST_INR:
                    lots.pop(0)
            if remaining * sell_nav > DUST_INR:
                rep.warnings.append(
                    f"MF_SIP: {scheme} shows more units redeemed than "
                    "bought - the extra part is left out of the tax report")
        if lots:
            open_lots[key] = lots

    # ---- unrealised / sell-planning (§6.16) --------------------------------
    eq_rule_now = tax_rule_for(rules, "equity", today)
    eq_lt = eq_rule_now.lt_days if eq_rule_now else 365
    for r in data.equity:
        if not (r.qty and r.avg_cost and r.close and r.cost_date):
            continue
        if r.cost_date >= today:
            continue
        isin = r.isin_override or isin_by_name.get(r.scrip, "")
        # identical arithmetic to equity_flows (§6.2) so the two surfaces
        # can never drift
        value = r.qty * (r.ca_factor or 1.0) * r.close
        cf = r.cost_factor if r.cost_factor is not None else 1.0
        cost = r.qty * r.avg_cost * cf
        note = ""
        if r.cost_date < GRANDFATHER_CUTOFF:
            f = fmv_per_share(isin, today)
            if f is not None:
                qty_today = r.qty * (r.ca_factor or 1.0)
                px_today = r.close
                gf_sh = max((cost / qty_today) if qty_today else 0.0,
                            min(f, px_today))
                cost = qty_today * gf_sh
                note = "grandfathered (31-Jan-2018 value)"
        held = (today - r.cost_date).days
        rep.unrealised.append(UnrealisedRow(
            owner=r.owner, name=r.scrip, bucket="equity", qty=r.qty,
            value_today=value, gain_today=value - cost,
            term=_term(held, eq_lt),
            lt_on=(None if held > eq_lt
                   else r.cost_date + timedelta(days=eq_lt + 1)),
            note=note))
    nav_by_scheme = {m.scheme: m.current_nav for m in data.mutual_funds
                     if m.scheme and m.current_nav}
    for (owner, scheme), lots in open_lots.items():
        nav = nav_by_scheme.get(scheme)
        if not nav:
            continue
        tt = tax_type_by_scheme.get(scheme, "")
        is_debt = tt == "debt"
        assumed = tt == ""
        for lot in lots:
            held = (today - lot["date"]).days
            if is_debt and lot["date"] >= DEBT_MF_SLAB_FROM:
                bucket, term, lt_on = "slab", "At your slab", None
            else:
                bucket = "mf_debt" if is_debt else "mf_equity"
                rule = tax_rule_for(rules, bucket, today)
                ltd = rule.lt_days if rule else 365
                term = _term(held, ltd)
                lt_on = (None if held > ltd
                         else lot["date"] + timedelta(days=ltd + 1))
            # same caveats as the realised rows (§3.21: every caveat lives in
            # the row's Note) — a planning figure must not hide its guesses
            note = ""
            if not is_debt and assumed:
                note = "assumed Equity - set Tax type on MutualFunds"
            if bucket == "mf_equity" and lot["date"] < GRANDFATHER_CUTOFF:
                note = (note + "; " if note else "") + \
                    "pre-2018 fund - grandfathering not applied, " \
                    "gain may be overstated"
            rep.unrealised.append(UnrealisedRow(
                owner=owner, name=scheme, bucket=bucket, qty=lot["units"],
                value_today=lot["units"] * nav,
                gain_today=lot["units"] * (nav - lot["cpu"]),
                term=term, lt_on=lt_on, note=note))

    # ---- FY summaries (newest first) ---------------------------------------
    by_fy: dict[str, list[RealisedRow]] = {}
    for row in rep.realised:
        by_fy.setdefault(row.fy, []).append(row)
    fy_end_year = {fy: int(fy.split("-")[0]) + 1 for fy in by_fy}
    for fy in sorted(by_fy, reverse=True):
        rows = by_fy[fy]
        s = FYSummary(fy=fy)
        # equity bucket = equity + mf_equity: they share the §112A exemption
        eq_rows = [r for r in rows if r.bucket in ("equity", "mf_equity")]
        s.stcg = sum(r.gain for r in eq_rows if r.term == "Short-term")
        s.ltcg = sum(r.gain for r in eq_rows if r.term == "Long-term")
        s.slab_gain = sum(r.gain for r in rows if r.bucket == "slab")
        s.debt_gain = sum(r.gain for r in rows if r.bucket == "mf_debt")
        s.spec_gain = sum(r.gain for r in rows if r.bucket == "speculative")
        fy_end = date(fy_end_year[fy], 3, 31)
        rule_end = tax_rule_for(rules, "equity", fy_end)
        if rule_end:
            # Sec 70 same-FY set-off, across buckets (§6.16 is normative).
            # Era-gated: rule-less pre-2018 §10(38) FYs (LTCG exempt,
            # losses carry-forward only) keep st_setoff 0 like their blank
            # tax cells. Derived from NETTED figures, not the tax loop's
            # leftover (that loop skips unknown-rate rows and would
            # overstate the excess). Dust below ₹0.005 clamps at each
            # source so no consumer can shift microscopically. The whole
            # SHORT-TERM head nets first — a debt/slab ST GAIN absorbs an
            # equity ST loss (Sec 50AA deems slab lots short-term), so only
            # a genuine all-ST net loss spills to LTCG.
            def _dust(x: float) -> float:
                return x if abs(x) >= 0.005 else 0.0
            debt_st_net = _dust(
                sum(r.gain for r in rows if r.bucket == "mf_debt"
                    and r.term == "Short-term") + s.slab_gain)
            debt_lt_loss = _dust(max(0.0, -sum(
                r.gain for r in rows if r.bucket == "mf_debt"
                and r.term == "Long-term")))
            debt_st_loss = max(0.0, -debt_st_net)     # feeds the tax loop
            excess_st = _dust(max(0.0, -(s.stcg + debt_st_net)))
            s.st_setoff = _dust(min(debt_lt_loss + excess_st,
                                    max(s.ltcg, 0.0)))
            # debt losses absorbed by the equity ST figure — surfaced on
            # the console so "STCG ₹1L, tax ₹0" never looks impossible
            s.st_sheltered = _dust(min(debt_st_loss, max(0.0, s.stcg)))
            s.exemption = rule_end.ltcg_exempt
            s.exemption_used = min(s.ltcg_eff, s.exemption)
            s.headroom = max(0.0, s.exemption - s.ltcg_eff)
            # STCG tax per sale-date rule (the 2024-07-23 mid-FY switch):
            # ST losses — equity-family and debt/slab — offset the
            # highest-taxed gains first (the taxpayer-favourable order)
            st_rows = [r for r in eq_rows if r.term == "Short-term"]
            st_loss = (-sum(r.gain for r in st_rows if r.gain < 0)
                       + debt_st_loss)
            gains: list[tuple[float, float]] = []      # (gain, rate %)
            computable = True
            for r in st_rows:
                if r.gain <= 0:
                    continue
                # each row at ITS OWN asset's rate — equity and mf_equity
                # ship identical bundled rates, but Tax_Rules lets a user
                # diverge them, and then an MF gain must not tax at the
                # share rate (only the §112A exemption bucket is shared)
                rr = tax_rule_for(rules, r.bucket, r.sell_date)
                if rr is None or rr.stcg_pct is None:
                    computable = False
                    continue
                gains.append((r.gain, rr.stcg_pct))
            tax_st = 0.0
            for gain, pct in sorted(gains, key=lambda g: -g[1]):
                offset = min(gain, st_loss)
                st_loss -= offset
                tax_st += (gain - offset) * pct / 100
            s.tax_stcg = tax_st if computable else None
            if rule_end.ltcg_pct is not None:
                s.tax_ltcg = (max(0.0, s.ltcg_eff - s.exemption)
                              * rule_end.ltcg_pct / 100)
        rep.summaries.append(s)

    # current-FY headroom even when nothing was sold this year yet
    rule_now = tax_rule_for(rules, "equity", today)
    if rule_now:
        ltcg_now = next((s.ltcg_eff for s in rep.summaries
                         if s.fy == rep.fy_now), 0.0)
        rep.headroom_now = max(0.0, rule_now.ltcg_exempt - ltcg_now)
    return rep
