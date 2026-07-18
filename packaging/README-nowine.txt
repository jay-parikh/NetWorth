FAMILY PORTFOLIO TRACKER  (Windows)
===================================
Everything runs on YOUR computer. The only internet use is downloading public
price data (AMFI fund NAVs, BSE/NSE stock prices, corporate actions).

TO RUN
  1. Keep this whole folder together (don't move files out of it).
  2. Open Family_Portfolio_Tracker.xlsx, enter your holdings, then SAVE and
     CLOSE it. (Or skip the typing: save your fund statement - the CAS PDF
     from camsonline.com - or your broker's tradebook/holdings file in this
     folder, and the updater offers to type the history in for you. Funds
     held in demat come in from the holdings file too - one opening line
     each - and splits/bonuses are never counted twice.)
  3. Double-click  "Update Portfolio.bat".
     A black window opens, refreshes prices/NAVs/XIRR, and pauses so you can
     read the summary. Press a key to close it.

FIRST-RUN WARNINGS
  Windows SmartScreen may warn about the .bat (it is not code-signed):
  click "More info" -> "Run anyway". Nothing is installed; nothing is uploaded.

WHAT'S INSIDE
  Family_Portfolio_Tracker.xlsx   your workbook
  Update Portfolio.bat            the updater (double-click this)
  app\                            a private copy of Python + this tool (leave it)

No separate Python install is needed — a private copy is bundled in app\.

Project home & new versions:  https://github.com/jay-parikh/NetWorth
