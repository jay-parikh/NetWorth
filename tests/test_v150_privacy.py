"""v1.5.0 privacy: the ••• Mask (curtain) + the encryption Lock (safe).

Four legal Mask×Lock states, one password. Values must survive every state
untouched (the workbook is the data store); a masked build may not leak a
figure through any visual channel; the Lock path must never write plaintext.
"""

import re
import zipfile
from dataclasses import asdict
from datetime import date
from pathlib import Path

import pytest
from openpyxl import load_workbook

from networth import crypto
from networth.fetch.amfi import AmfiData
from networth.fetch.bhavcopy import PriceData
from networth.generate import MASK_NUM, build_workbook
from networth.reader import read_workbook
from networth.sample_data import sample_portfolio
from networth.update import relock, run

TODAY = date(2026, 7, 15)
PW = "Fam!2026"


def _priv_data(mask=True, lock=False, hash_=True):
    d = sample_portfolio()
    d.privacy_enabled = mask
    d.lock_enabled = lock
    if hash_:
        d.privacy_hash = crypto.hash_password(PW)
    return d


def _upd(path, **kw):
    return run(Path(path), price_data=PriceData(trade_date=TODAY, source="T"),
               amfi_data=AmfiData(), ca_data=[], div_data=[], today=TODAY, **kw)


def _sheet_xmls(path):
    with zipfile.ZipFile(path) as z:
        return {n: z.read(n).decode() for n in z.namelist()
                if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", n)}


# ---- the mask itself -------------------------------------------------------

def test_masked_build_keeps_values_and_dates(tmp_path):
    path = tmp_path / "wb.xlsx"
    build_workbook(_priv_data(), str(path), masked=True)
    back = read_workbook(str(path))
    assert back.equity[0].qty == 50.0                 # value intact under •••
    assert back.equity[0].cost_date == date(2018, 1, 31)   # dates stay visible
    assert back.masked_at_rest is True
    # openpyxl must NOT see masked numbers as dates (the numFmtId collision
    # regression: a default num_format made every number render as a date)
    wb = load_workbook(path)
    assert wb["Equity"]["D4"].value == 50


def test_mask_format_has_three_explicit_sections(tmp_path):
    # a single-section mask renders negatives as "-•••" — a sign leak
    assert MASK_NUM.count("•••") == 3 and MASK_NUM.endswith("@")
    path = tmp_path / "wb.xlsx"
    build_workbook(_priv_data(), str(path), masked=True)
    with zipfile.ZipFile(path) as z:
        styles = z.read("xl/styles.xml").decode()
    assert styles.count("&quot;•••&quot;;&quot;•••&quot;;&quot;•••&quot;;@") >= 1


def test_masked_build_suppresses_every_leaky_visual(tmp_path):
    path = tmp_path / "wb.xlsx"
    build_workbook(_priv_data(), str(path), masked=True)
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        charts = [n for n in names if n.startswith("xl/charts/chart")]
        drawings = " ".join(z.read(n).decode() for n in names
                            if re.fullmatch(r"xl/drawings/drawing\d+\.xml", n))
    sheets = "".join(_sheet_xmls(path).values())
    assert charts == []                               # charts redraw values
    assert drawings.count("Charts are hidden") == 10  # a note per chart
    assert "dataBar" not in sheets                    # bar length ∝ value
    assert "iconSet" not in sheets                    # ▲/▼ leaks the sign
    assert "<conditionalFormatting" not in sheets     # red/green leaks too


def test_masked_build_protects_every_sheet(tmp_path):
    path = tmp_path / "wb.xlsx"
    build_workbook(_priv_data(), str(path), masked=True)
    for name, xml in _sheet_xmls(path).items():
        assert "<sheetProtection" in xml and 'selectLockedCells="1"' in xml, name
    build_workbook(_priv_data(), str(path), masked=False)
    assert all("<sheetProtection" not in xml
               for xml in _sheet_xmls(path).values())


def test_roundtrip_identity_through_a_masked_build(tmp_path):
    data = _priv_data()
    p1, p2 = tmp_path / "a.xlsx", tmp_path / "b.xlsx"
    build_workbook(data, str(p1), masked=True)
    back = read_workbook(str(p1))
    build_workbook(back, str(p2), masked=True)
    assert asdict(read_workbook(str(p2))) == asdict(back)
    assert back.privacy_enabled and back.privacy_hash == data.privacy_hash


# ---- the lock --------------------------------------------------------------

def test_lock_encrypt_decrypt_roundtrip_and_wrong_password(tmp_path):
    path = tmp_path / "wb.xlsx"
    build_workbook(_priv_data(lock=True), str(path))
    plain = path.read_bytes()
    cipher = crypto.encrypt_workbook(plain, PW)       # self-verifies inside
    assert crypto.is_encrypted(cipher) and not crypto.is_encrypted(plain)
    assert crypto.decrypt_workbook(cipher, PW).getvalue() == plain
    with pytest.raises(crypto.WrongPassword):
        crypto.decrypt_workbook(cipher, "nope")


def test_password_hash_survives_the_workbook_and_verifies(tmp_path):
    path = tmp_path / "wb.xlsx"
    data = _priv_data()
    build_workbook(data, str(path))
    back = read_workbook(str(path))
    assert back.privacy_hash == data.privacy_hash     # defined-name roundtrip
    assert crypto.verify_password(PW, back.privacy_hash)
    assert not crypto.verify_password("guess", back.privacy_hash)


# ---- run() lifecycles ------------------------------------------------------

def test_mask_lifecycle_enable_view_remask_reset(tmp_path):
    path = tmp_path / "wb.xlsx"
    d = sample_portfolio()
    d.privacy_enabled = True                          # user's Settings Yes
    build_workbook(d, str(path))

    s = _upd(path, password=PW)                       # first run sets the pw
    assert s["privacy"] == "masked"
    assert read_workbook(str(path)).masked_at_rest is True

    s = _upd(path, password=PW, reveal=True)          # view run
    assert s["privacy"] == "open (viewing)"
    assert read_workbook(str(path)).masked_at_rest is False

    s = _upd(path)                                    # headless re-mask
    assert s["privacy"] == "masked"
    back = read_workbook(str(path))
    assert back.masked_at_rest is True
    assert back.equity[0].close == pytest.approx(1520)  # prices still refresh
    # v1.6.2 policy: OLDER readable backups are purged once the mask
    # returns, but this run's own pre-run copy survives one cycle as the
    # rollback of last resort (the next masked run removes it)
    unmasked = [b for b in (tmp_path / "backups").glob("*.xlsx")
                if ".unmasked-backup-" in b.name]
    assert len(unmasked) == 1
    assert any("readable backup" in w for w in s["warnings"])

    s = _upd(path, reset_privacy=True)                # RESET escape hatch
    back = read_workbook(str(path))
    assert back.privacy_enabled is False and back.privacy_hash == ""
    assert s["privacy"] == ""


def test_lock_lifecycle_and_headless_safety(tmp_path):
    path = tmp_path / "wb.xlsx"
    build_workbook(_priv_data(mask=False, lock=True, hash_=False), str(path))

    s = _upd(path)                # headless: lock wanted, password unconfirmed
    assert not crypto.is_encrypted(path)              # NEVER encrypt blind
    assert any("password" in w for w in s["warnings"])

    s = _upd(path, password=PW)                       # confirmed → locked
    assert s["privacy"] == "locked" and crypto.is_encrypted(path)

    frozen = path.read_bytes()
    with pytest.raises(SystemExit):                   # headless can't update
        _upd(path)
    with pytest.raises(SystemExit):                   # wrong pw can't either
        _upd(path, password="wrong")
    assert path.read_bytes() == frozen                # file untouched

    s = _upd(path, password=PW)                       # normal locked update
    assert s["privacy"] == "locked" and crypto.is_encrypted(path)
    inner = read_workbook(crypto.decrypt_workbook(path, PW))
    assert inner.equity[0].qty == 50.0
    assert inner.masked_at_rest is False              # lock-only: no mask


def test_both_layers_inner_mask_and_reveal(tmp_path):
    path = tmp_path / "wb.xlsx"
    d = sample_portfolio()
    d.privacy_enabled = d.lock_enabled = True
    build_workbook(d, str(path))
    s = _upd(path, password=PW)
    assert s["privacy"] == "locked + masked" and crypto.is_encrypted(path)
    assert read_workbook(crypto.decrypt_workbook(path, PW)).masked_at_rest

    s = _upd(path, password=PW, reveal=True)          # show numbers this run
    assert s["privacy"] == "locked"                   # still encrypted!
    assert crypto.is_encrypted(path)
    assert not read_workbook(crypto.decrypt_workbook(path, PW)).masked_at_rest


def test_relock_is_offline(tmp_path, monkeypatch):
    path = tmp_path / "wb.xlsx"
    build_workbook(_priv_data(), str(path))           # mask on, at rest OPEN
    import networth.update as U

    def _boom(*a, **k):
        raise AssertionError("relock must not fetch anything")
    monkeypatch.setattr(U.bhav_mod, "fetch", _boom)
    monkeypatch.setattr(U.amfi_mod, "fetch", _boom)
    monkeypatch.setattr(U.ca_mod, "fetch", _boom)

    res = relock(path)
    assert res["privacy"] == "masked"
    assert read_workbook(str(path)).masked_at_rest is True


# ---- the interactive mask prompt -------------------------------------------

def test_mask_prompt_staged_flow(monkeypatch):
    """Enter alone keeps the mask (no password asked); a wrong password is
    caught at the prompt with retries; RESET needs its own YES."""
    import getpass
    from networth.update import _prompt_mask_password
    h = crypto.hash_password(PW)

    def feed(answers, typed):
        a, t = iter(answers), iter(typed)
        monkeypatch.setattr("builtins.input", lambda *_: next(a))
        monkeypatch.setattr(getpass, "getpass", lambda *_: next(t))

    feed([""], [])                                    # Enter → stay masked,
    assert _prompt_mask_password(h) == (None, False)  # password never asked
    feed(["y"], [PW])                                 # y + right password
    assert _prompt_mask_password(h) == (PW, False)
    feed(["y"], ["typo", PW])                         # typo, then right
    assert _prompt_mask_password(h) == (PW, False)
    feed(["y"], ["a", "b", "c"])                      # 3 wrong → stay masked
    assert _prompt_mask_password(h) == (None, False)
    feed(["y", "YES"], ["RESET"])                     # forgot → RESET → YES
    assert _prompt_mask_password(h) == (None, True)
    feed(["y", ""], ["RESET"])                        # RESET, then back out
    assert _prompt_mask_password(h) == (None, False)


def test_wrong_password_is_never_silent(tmp_path):
    path = tmp_path / "wb.xlsx"
    build_workbook(_priv_data(), str(path), masked=True)
    s = _upd(path, password="wrong", reveal=True)
    assert s["privacy"] == "masked"
    assert any("didn't match" in w for w in s["warnings"])
