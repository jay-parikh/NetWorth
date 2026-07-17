FAMILY PORTFOLIO TRACKER
========================
One Excel workbook for your family's entire net worth - shares, mutual funds
(incl. SIPs), fixed deposits, PPF, EPF, NPS, gold & silver, bonds, property,
cash, insurance - with one-click updates for prices, fund NAVs, the daily
bullion rate and annualised returns (XIRR). Everything runs on YOUR
computer; nothing is uploaded anywhere. The only internet access is
downloading public data (prices, NAVs, company announcements).

GET STARTED (5 minutes)
  1. Keep the files of this folder together, anywhere you like.
  2. Open Family_Portfolio_Tracker.xlsx and read its Guide tab (the gold
     one, last). The file opens calm on purpose: you see the five everyday
     tabs. Gold, EPF, NPS, property and more are one switch away on the
     Settings tab - each has a worked example waiting inside.
  3. It ships with FICTIONAL sample data for three people (Amit, Priya,
     Rahul). Replace it with your own:
       - your people: type names in the yellow Person cells on the
         Dashboard, then use those names in the Owner columns
       - blue/yellow cells are inputs, grey columns calculate themselves
       - mutual funds & stocks: PICK from the dropdown (type the first
         letters, press Enter, then open the list) - IDs fill in themselves
       - own something the file isn't showing? Settings tab -> set that
         class to Yes. Don't own something? Set it to No - it hides, its
         rows are kept, and it simply isn't counted.
  4. Save and CLOSE the workbook, then run the updater:
       Windows : double-click "Update Portfolio.exe"
       macOS   : right-click "Update Portfolio.command" -> Open
                 (only the first time; afterwards double-click works)
     It backs up your file (backups folder), fetches the latest data,
     recomputes every number and rewrites the workbook. It will also offer
     to add a family member or show/hide asset classes - just type a number.

GOOD TO KNOW
  - The workbook must be CLOSED while the updater runs.
  - A dated backup of your previous file is kept in backups/ (last 10).
  - Rows you add, delete or sort are all fine. Only don't rearrange the
    column layout or rename the tabs.
  - Some tabs are hidden on purpose (the stock/fund/bank/pension name lists
    that feed the dropdowns). Curious? Settings tab -> Reference lists ->
    Yes.
  - If a scheme/stock is missing from a dropdown (delisted/merged), type
    its name anyway, accept the warning, and fill the ISIN yourself.
  - Windows SmartScreen may warn on first run: More info -> Run anyway.
  - No Excel? The workbook also opens fine in LibreOffice.

PRIVACY
  Your data never leaves your machine. The updater only downloads public
  data: AMFI fund NAVs, BSE/NSE prices and announcements, NPS Trust NAVs,
  and the IBJA bullion rate.

  Want the file itself protected? Two optional switches on the Settings
  tab (one password for both):
    - Privacy mask : every number shows as ... until you type your
      password in the update window. A curtain, not a safe (RESET works
      if you forget).
    - Lock file    : real encryption - Excel asks for the password just
      to open the file, and updates need it too. NO recovery if the
      password is forgotten: write it down somewhere safe.

Project home, source code and new versions:
  https://github.com/jay-parikh/NetWorth
