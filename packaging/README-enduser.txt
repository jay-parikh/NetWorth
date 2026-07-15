FAMILY PORTFOLIO TRACKER
========================
One Excel workbook for your family's entire net worth - Equity, Mutual Funds
(incl. SIPs), Fixed Deposits, PPF and Bonds - with one-click price/NAV updates
and XIRR (annualised returns). Everything runs on YOUR computer; nothing is
uploaded anywhere. The only internet access is downloading public price data
(AMFI fund NAVs, BSE/NSE stock prices).

GET STARTED (5 minutes)
  1. Keep the files of this folder together, anywhere you like.
  2. Open Family_Portfolio_Tracker.xlsx and read its Guide sheet.
  3. It ships with FICTIONAL sample data for three people (Amit, Priya,
     Rahul). Replace it with your own:
       - your people: type names in the Owner columns AND in the yellow
         Person cells on the Dashboard (rows 6-15)
       - blue/yellow cells are inputs, grey columns calculate themselves
       - Mutual funds & stocks: PICK from the dropdown (type the first
         letters, then open the list) - ISINs fill in automatically
  4. Save and CLOSE the workbook, then run the updater:
       Windows : double-click "Update Portfolio.exe"
       macOS   : right-click "Update Portfolio.command" -> Open
                 (only the first time; afterwards double-click works)
     It backs up your file (backups folder), fetches the latest prices and
     NAVs, recomputes all XIRR figures and rewrites the workbook.

GOOD TO KNOW
  - The workbook must be CLOSED while the updater runs.
  - A dated backup of your previous file is kept in backups/ (last 10).
  - Rows you add, delete or sort are all fine. Only don't rearrange the
    column layout or edit the two Master sheets by hand.
  - If a scheme/stock is missing from a dropdown (delisted/merged), type
    its name anyway, accept the warning, and fill the ISIN yourself.
  - Windows SmartScreen may warn on first run: More info -> Run anyway.
  - No Excel? The workbook also opens fine in LibreOffice.

PRIVACY
  Your data never leaves your machine. The updater only downloads:
    - https://www.amfiindia.com/spages/NAVAll.txt        (fund NAVs)
    - https://www.bseindia.com/...  /  nsearchives.nseindia.com/...
                                                          (stock prices)

Project home, source code and new versions:
  https://github.com/jay-parikh/NetWorth
