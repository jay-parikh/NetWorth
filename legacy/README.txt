FAMILY PORTFOLIO TRACKER - TEMPLATE (with sample data)
======================================================
One Excel workbook to track a family's net worth across Equity, Mutual Funds
(incl. SIPs), Fixed Deposits, PPF and Bonds - with automatic price/NAV updates
and XIRR (annualised returns).

The workbook is pre-filled with FICTIONAL sample data for three people
(Amit, Priya, Rahul) using real ISINs, so everything works out of the box.
Replace the sample rows with your own data.

GET STARTED
  1. Keep all files of this folder together. Open Family_Portfolio_Tracker.xlsx.
  2. Read the Guide sheet (2 minutes).
  3. Replace the sample rows with your holdings:
       - your people: type names in the Owner columns AND in the yellow Person
         cells on the Dashboard (rows 6-15)
       - Equity / MutualFunds / MF_SIP / FixedDeposits / PPF / Bonds sheets:
         yellow-ish cells are inputs, grey columns calculate themselves
       - Mutual funds: PICK the Scheme Name from the dropdown - Fund House and
         ISIN fill themselves from the MF_Master sheet (ships with the full
         AMFI list, ~14,000 schemes)
       - Equity: PICK the Scrip from the dropdown - ISIN fills itself from the
         Stock_Master sheet (ships with ~4,500 BSE-listed stocks)
  4. Double-click UpdatePrices.bat    -> refreshes share prices (BSE/NSE, by ISIN)
                                         and merges new listings into Stock_Master
     Double-click UpdateNAV.bat       -> refreshes fund NAVs (AMFI, by ISIN)
     Double-click UpdateFundMaster.bat-> refreshes the MF_Master fund list (AMFI)
     The first two also compute and write the XIRR columns. All need desktop
     Microsoft Excel (no Python or other installs). Each writes a log and pauses.

PORTFOLIO XIRR, INFLATION & 20-YEAR PROJECTION
  - The Dashboard shows a single 'Portfolio XIRR' across ALL asset classes
    (Equity + Mutual Funds + FDs + PPF + Bonds) plus a per-class XIRR column.
    Both updaters recompute these every run.
  - PPF keeps no contribution history, so its return is estimated at the
    'Rate %' you enter, from the balance as-on date.
  - Bonds: fill the new 'Buy Date' column - rows without it stay out of XIRR.
    Coupons are not tracked, so bond XIRR is price change only.
  - 'Inflation % p.a.' on the Dashboard is yours to edit (default 7). The
    'Real return' cell and the Projection sheet - a 20-year corpus trajectory
    chart of portfolio return vs inflation - update instantly, no script run.

NOTES
  - Mutual funds: one row per fund on MutualFunds; one row per purchase on
    MF_SIP (SIP installment or lump sum - same thing, one row each).
    Redemption = negative Amount.
  - MF_Master and Stock_Master are machine-managed - don't edit them by hand.
    If a scheme/scrip is missing there (delisted/merged), Excel shows a warning
    you can accept; then type the ISIN (and Fund House) yourself for that row.
  - Stock_Master updates are add-only: new listings appear after each
    UpdatePrices.bat run, existing names stay stable so your rows keep working.
  - Equity 'Cost date' drives per-stock XIRR; set it to your buy date.
  - Deleting/inserting/sorting rows is safe - there are no hidden helpers.
  - XIRR columns are plain values written by the updaters; re-run after changes.
