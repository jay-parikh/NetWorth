"""Guide sheet content — the friendly, in-workbook manual for end users.

Each row is (style, text). Styles the generator understands:
  "title"   the big heading (once, at the top)
  "h"       a section heading
  "legend"  a sentinel: the generator draws the two coloured colour-legend lines
  ""        ordinary body text (blank text = spacer row)

Plain language on purpose — keep it simple, keep the meaning. Update it when a
feature changes what the user does.
"""

GUIDE_ROWS: list[tuple[str, str]] = [
    ("title", "How to use your Family Portfolio Tracker"),
    ("", ""),
    ("", "This one file keeps track of your whole family's net worth in one place."),
    ("", "You type in what you own — the file does all the maths for you. Every so"),
    ('', 'often you double-click "Update Portfolio" to pull in fresh prices. That\'s it.'),
    ("", ""),

    ("h", "What the colours mean"),
    ("legend", ""),
    ("", "In short: only ever type into the blue/yellow cells. Leave the grey ones alone."),
    ("", ""),

    ("h", "Do this first"),
    ("", "1.  On the Dashboard, delete the sample family (Amit, Priya, Rahul) and type"),
    ("", "    your own people's names into the yellow boxes."),
    ("", "2.  Delete the sample holdings and add your own (see the next section)."),
    ('', '3.  Save, close the file, and double-click "Update Portfolio".'),
    ("", ""),

    ("h", "Where each thing goes"),
    ("", "    Shares .............  the Equity tab"),
    ("", "    Mutual funds .......  the MutualFunds tab (one row per fund), and log"),
    ("", "                          each purchase on the MF_SIP tab"),
    ("", "    Fixed deposits .....  the FixedDeposits tab"),
    ("", "    PPF ................  the PPF tab (list deposits on PPF_Ledger for exact"),
    ("", "                          interest — optional)"),
    ("", "    Bonds ..............  the Bonds tab"),
    ("", ""),
    ("", "Tip: pick the fund, share or bank from the dropdown and its ID fills in for"),
    ("", "you. To search a long list: type the first few letters, press Enter, then"),
    ("", "open the dropdown — it now shows only the matching names."),
    ("", "(Selling a fund? Log it on MF_SIP as a purchase with a minus amount.)"),
    ("", ""),

    ("h", "Add a family member"),
    ('', 'When you run "Update Portfolio" it asks "Add a new person?" — just type a'),
    ("", "name and their tab is built for you. (Or type the name into a yellow box on"),
    ("", "the Dashboard yourself.) Then use that same name in the Owner column of any"),
    ("", "holding, and their totals roll into the family total and charts."),
    ("", ""),

    ("h", "Keep it up to date"),
    ('', 'Close the file and double-click "Update Portfolio". One run backs up your'),
    ("", "file first, then refreshes share prices, fund values, splits & bonuses, and"),
    ("", "every return figure. Do it whenever you like — once a week is plenty."),
    ("", ""),

    ("h", "What it quietly handles for you (nothing to do)"),
    ("", "    Splits & bonuses  —  your share count updates in the 'Qty today' column;"),
    ("", "                         the numbers you typed are never changed."),
    ("", "    PPF interest      —  worked out exactly if you list deposits on PPF_Ledger,"),
    ("", "                         otherwise it uses the balance you typed."),
    ("", "    An old share whose —  bought before Feb 2018? Leave 'Avg. cost' blank and"),
    ("", "    price you forget       it fills in the official 2018 value for you."),
    ("", "    A delisted stock  —  kept at its last known price and shaded amber."),
    ("", ""),

    ("h", "Good to know"),
    ("", "    •  Return figures (XIRR) are filled in by the updater — run it again after"),
    ("", "       you add or change holdings."),
    ("", "    •  Add, delete or sort rows freely. Just don't rename the tabs — the"),
    ("", "       Dashboard finds them by name."),
    ('', '    •  Change "Inflation %" on the Dashboard to watch the 20-year Projection'),
    ("", "       react instantly."),
    ("", "    •  The Master tabs (the fund and stock lists) look after themselves —"),
    ("", "       don't edit them by hand."),
    ("", "    •  Nothing about your money ever leaves your computer."),
    ("", ""),
    ("", "Questions, help and new versions:   github.com/jay-parikh/NetWorth"),
]
