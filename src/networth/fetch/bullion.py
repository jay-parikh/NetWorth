"""Daily gold/silver reference rate, ₹ per gram of fine metal (SPEC §5.7).

Layered by design — this is the flakiest data in the product, so it must
never block a run:

1. PRIMARY: the IBJA daily benchmark (ibjarates.com) — the rate the bullion
   trade itself quotes from; RBI uses IBJA 999 for SGB redemption. The page
   carries stable span ids: lblGold999_PM (₹/10 g) and lblSilver999_PM
   (₹/kg), with _AM fallbacks earlier in the day.
2. FALLBACK: a market-implied rate from the bhavcopy we already fetched —
   the median ₹/g over the proxies in data/bullion_proxies.csv (SGB tranches
   trade at ₹/gram; GoldBeES ≈ 0.01 g/unit; SilverBeES ≈ 1 g/unit). Usually
   2–4 % below the IBJA retail benchmark; the Guide says so.
3. The updater keeps the previous rate (amber after 7 days) when both fail,
   and the sheet's Rate-override column always wins.
"""

from __future__ import annotations

import csv
import re
from statistics import median

from ..model import DATA_DIR

IBJA_URL = "https://www.ibjarates.com/"
_HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
            "Accept": "text/html,*/*", "Accept-Language": "en-US,en;q=0.9"}

# span id → (metal, divisor to ₹/gram); PM preferred, AM as fallback
_IBJA_SPANS = [("lblGold999_PM", "gold", 10.0), ("lblGold999_AM", "gold", 10.0),
               ("lblSilver999_PM", "silver", 1000.0),
               ("lblSilver999_AM", "silver", 1000.0)]


def parse_ibja(html: str) -> dict[str, float]:
    """→ {"gold": ₹/g, "silver": ₹/g} for the ids present and positive."""
    out: dict[str, float] = {}
    for span, metal, div in _IBJA_SPANS:
        if metal in out:
            continue
        m = re.search(rf'id="{span}"[^>]*>\s*([\d,]+(?:\.\d+)?)\s*<', html)
        if m:
            try:
                v = float(m.group(1).replace(",", ""))
            except ValueError:
                continue
            if v > 0:
                out[metal] = round(v / div, 2)
    return out


def fetch_ibja(session=None, timeout: int = 30) -> dict[str, float]:
    """Never raises — {} on any failure (the caller falls back)."""
    try:
        import requests
        sess = session or requests.Session()
        resp = sess.get(IBJA_URL, headers=_HEADERS, timeout=timeout)
        if resp.status_code != 200:
            return {}
        return parse_ibja(resp.text)
    except Exception:  # noqa: BLE001 — graceful degradation by contract
        return {}


def load_proxies(data_dir=DATA_DIR) -> list[dict]:
    with open(data_dir / "bullion_proxies.csv", newline="",
              encoding="utf-8") as f:
        return list(csv.DictReader(f))


def derive_from_bhavcopy(price_data, proxies: list[dict] | None = None
                         ) -> dict[str, float]:
    """Market-implied ₹/g per metal: median close/grams over the quoted
    proxies. Empty dict per metal with no quoted proxy."""
    if price_data is None:
        return {}
    proxies = proxies if proxies is not None else load_proxies()
    isin_by_symbol = {sym: isin for sym, _n, isin in price_data.master_rows}
    samples: dict[str, list[float]] = {}
    for p in proxies:
        grams = float(p["grams_per_unit"])
        if p["match"] == "isin":
            isins = [p["key"]]
        else:                              # symbol_prefix
            isins = [i for s, i in isin_by_symbol.items()
                     if s.startswith(p["key"])]
        for isin in isins:
            quote = price_data.prices.get(isin)
            if quote and quote["close"] > 0 and grams > 0:
                samples.setdefault(p["metal"], []).append(
                    quote["close"] / grams)
    return {metal: round(median(vals), 2) for metal, vals in samples.items()}
