"""Shared test plumbing.

The shipped sample portfolio populates EVERY asset class (v1.4 onboarding:
delete what you don't own). Two consequences for tests:

1. `run()` would try to fetch live NPS NAVs and IBJA bullion rates for the
   sample rows — the autouse fixture below stubs both with deterministic
   values so no test ever touches the network.
2. Tests that assert the CLASSIC five-class layout use `classic()` to strip
   and disable the newer classes first.
"""

import pytest

from networth.fetch.nps import NpsData
from networth.model import ClassSetting
from networth.sample_data import sample_portfolio

STUB_BULLION = {"gold": 14167.9, "silver": 217.43}
STUB_NPS = NpsData(
    nav_by_code={"SM001003": 56.7834},
    master_rows=[("SM001003", "SBI PENSION FUND SCHEME E - TIER I",
                  "SBI Pension Funds Pvt. Ltd.")],
)


@pytest.fixture(autouse=True)
def _no_network_side_feeds(monkeypatch):
    import networth.fetch.bullion as bullion_mod
    import networth.fetch.nps as nps_mod
    monkeypatch.setattr(bullion_mod, "fetch_ibja",
                        lambda *a, **k: dict(STUB_BULLION))
    monkeypatch.setattr(nps_mod, "fetch",
                        lambda *a, **k: NpsData(
                            nav_by_code=dict(STUB_NPS.nav_by_code),
                            master_rows=list(STUB_NPS.master_rows)))


NEW_CLASS_KEYS = ("epf", "gold_silver", "nps", "real_estate", "cash",
                  "insurance", "other_assets")


def classic(data=None):
    """A sample portfolio slimmed to the classic five classes — for tests
    that assert the classic Dashboard/person/History layout."""
    data = data or sample_portfolio()
    data.epf = []
    data.bullion = []
    data.nps = []
    data.manual_assets = []
    data.bullion_rate_asof = None
    for key in NEW_CLASS_KEYS:
        data.class_settings[key] = ClassSetting(enabled=False)
    return data
