"""Guide sheet content — the friendly, in-workbook manual for end users.

Each row is a small tuple whose first item is its KIND; the generator
(`_write_guide`) turns these into a designed page (title banner, coloured
section bars with icons, numbered step badges, a colour legend with swatches,
key/value rows). Kinds:

    ("title", heading, subtitle)
    ("section", emoji, heading)     coloured header bar
    ("legend",)                     the colour key (swatches)
    ("step", n, text)               numbered badge + text
    ("kv", key, value)              bold key — value
    ("bullet", text)                • text
    ("tip", text)                   muted "Tip · …"
    ("text", text)                  plain line
    ("footer", text)                muted footer
    ("space",)                      breathing room

Plain language on purpose — keep it simple, keep the meaning.
"""

GUIDE_ROWS: list[tuple] = [
    ("title", "Your Family Portfolio Tracker",
     "One file for your whole family's net worth. You type in what you own — it does the maths."),
    ("space",),

    ("section", "🎨", "What the colours mean"),
    ("legend",),
    ("text", "Rule of thumb: only ever type into the blue / yellow cells."),
    ("space",),

    ("section", "🚀", "Do this first"),
    ("step", "1", "On the Dashboard, replace the sample names (Amit, Priya, Rahul) with your family."),
    ("step", "2", "Delete the sample holdings and add your own — see 'Where each holding goes' below."),
    ("step", "3", 'Save, close the file, and double-click "Update Portfolio".'),
    ("space",),

    ("section", "📒", "Where each holding goes"),
    ("kv", "Shares", "the Equity tab"),
    ("kv", "Mutual funds", "the MutualFunds tab (one row per fund) + log each purchase on MF_SIP"),
    ("kv", "Fixed deposits", "the FixedDeposits tab"),
    ("kv", "PPF", "the PPF tab   (list deposits on PPF_Ledger for exact interest — optional)"),
    ("kv", "EPF", "the EPF tab — copy the balance and date from your EPFO passbook, done"),
    ("kv", "Bonds", "the Bonds tab"),
    ("kv", "House, savings, insurance", "the Manual_Assets tab — type today's value, pick the Class"),
    ("space",),
    ("tip", "Pick funds, shares and banks from the dropdown — the ID fills in for you."),
    ("tip", "Searching a long list? Type a few letters, press Enter, then open the dropdown."),
    ("tip", "Selling a fund? Log it on MF_SIP as a purchase with a minus amount."),
    ("space",),

    ("section", "🗂️", "Show only what you own"),
    ("text", "On the Settings tab, set any asset class to No and its tabs disappear —"),
    ("text", "a tidy file with only what your family actually owns. Flip it back to Yes anytime;"),
    ("text", "nothing is ever deleted, and a class holding rows always stays visible."),
    ("space",),

    ("section", "🎯", "Keep your money balanced"),
    ("text", "Type a Target % beside each class on Settings (say Equity 55, Gold 10)."),
    ("text", "The Dashboard then shows green “On target” or a red hint like “Move ₹1,20,000 out” —"),
    ("text", "live, the moment you edit a holding. The tolerance (±5 points) is yours to change."),
    ("space",),

    ("section", "👥", "Add a family member"),
    ("text", 'When you run "Update Portfolio" it asks “Add a new person?” — type a name and'),
    ("text", "their tab is built for you. Then use that same name in the Owner column of any holding."),
    ("space",),

    ("section", "🔄", "Keep it up to date"),
    ("text", 'Close the file and double-click "Update Portfolio". It backs up your file first, then'),
    ("text", "refreshes prices, fund values, splits & bonuses and every return figure. Weekly is plenty."),
    ("space",),

    ("section", "✨", "What it quietly handles for you"),
    ("kv", "Splits & bonuses", "your share count updates in 'Qty today' — your typed numbers never change"),
    ("kv", "Dividends", "each one your shares declare this year is logged on the Dividends tab, with a month-by-month chart"),
    ("kv", "PPF interest", "exact if you list deposits on PPF_Ledger, otherwise your typed balance is used"),
    ("kv", "A forgotten old cost", "bought before Feb 2018? leave 'Avg. cost' blank — it fills in the 2018 value"),
    ("kv", "Delisted stocks", "kept at their last known price and shaded amber"),
    ("space",),

    ("section", "💡", "Good to know"),
    ("bullet", "Return figures (XIRR) are filled in by the updater — run it again after you make changes."),
    ("bullet", "Add, delete or sort rows freely. Just don't rename the tabs — the Dashboard finds them by name."),
    ("bullet", 'Change "Inflation %" on the Dashboard to watch the 20-year Projection react instantly.'),
    ("bullet", "The Master tabs (fund & stock lists) look after themselves — don't edit them by hand."),
    ("bullet", "Nothing about your money ever leaves your computer."),
    ("space",),

    ("footer", "Questions, help and new versions:    github.com/jay-parikh/NetWorth"),
]
