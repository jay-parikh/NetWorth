"""Parser fixtures for the AMFI and bhavcopy data contracts (SPEC §5)."""

from networth.fetch.amfi import parse as amfi_parse
from networth.fetch.bhavcopy import parse as bhav_parse

AMFI_SAMPLE = """Scheme Code;ISIN Div Payout/ ISIN Growth;ISIN Div Reinvestment;Scheme Name;Net Asset Value;Date

Open Ended Schemes ( Equity Scheme - Flexi Cap Fund )

PPFAS Mutual Fund

122639;INF879O01027;-;Parag Parikh Flexi Cap Fund - Direct Plan - Growth;91.2000;14-Jul-2026
122640;INF879O01035;INF879O01043;Parag Parikh Flexi Cap Fund - Regular Plan - Growth;84.1000;14-Jul-2026

SBI Mutual Fund

103504;INF200K01QX4;-;SBI Large Cap FUND-DIRECT PLAN -GROWTH;96.4000;14-Jul-2026
999999;INF200K99999;-;Broken NAV Scheme;N.A.;14-Jul-2026
"""

BHAV_SAMPLE = """TradDt,BizDt,Sgmt,Src,FinInstrmTp,FinInstrmId,ISIN,TckrSymb,SctySrs,FinInstrmNm,OpnPric,HghPric,LwPric,ClsPric,LastPric,PrvsClsgPric,TtlTradgVol
2026-07-14,2026-07-14,CM,BSE,STK,500325,INE002A01018,RELIANCE,A,RELIANCE INDUSTRIES LTD.,1510,1531,1502,1520.5,1519,1512.3,100000
2026-07-14,2026-07-14,CM,BSE,STK,532540,INE467B01029,TCS,A,TATA CONSULTANCY SERVICES LTD.,3440,3462,3421,3455.0,3454,3462.1,50000
2026-07-14,2026-07-14,CM,BSE,STK,999999,INE999X01010,ZEROCO,Z,ZERO PRICE CO,0,0,0,0,0,0,0
"""


def test_amfi_parse():
    out = amfi_parse(AMFI_SAMPLE)
    assert out.nav_by_isin["INF879O01027"] == 91.2
    assert out.nav_by_isin["INF200K01QX4"] == 96.4
    # both ISIN columns of one scheme are captured
    assert out.nav_by_isin["INF879O01035"] == out.nav_by_isin["INF879O01043"] == 84.1
    # N.A. NAV skipped
    assert "INF200K99999" not in out.nav_by_isin
    # fund house = most recent section header
    rows = {isin: (fund, scheme) for fund, scheme, isin in out.master_rows}
    assert rows["INF879O01027"][0] == "PPFAS Mutual Fund"
    assert rows["INF200K01QX4"][0] == "SBI Mutual Fund"


def test_bhavcopy_parse():
    out = bhav_parse(BHAV_SAMPLE)
    q = out.prices["INE002A01018"]
    assert q["close"] == 1520.5 and q["prev"] == 1512.3
    # zero close dropped
    assert "INE999X01010" not in out.prices
    # master rows carry symbol + name + isin
    assert ("RELIANCE", "RELIANCE INDUSTRIES LTD.", "INE002A01018") in out.master_rows


def test_bhavcopy_never_confuses_prev_for_close():
    csv_text = ("ISIN,PrvsClsgPric,ClsPric\n"
                "INE0TEST0001,90,100\n")
    out = bhav_parse(csv_text)
    assert out.prices["INE0TEST0001"]["close"] == 100
    assert out.prices["INE0TEST0001"]["prev"] == 90
