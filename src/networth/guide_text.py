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
    ("step", "2", "Every tab has a few sample rows showing how it works. Replace them with your own,"),
    ("step", "3", "and delete the samples for things you don't own — those tabs then tidy themselves away."),
    ("step", "4", 'Save, close the file, and double-click "Update Portfolio".'),
    ("space",),

    ("section", "📒", "Where each holding goes"),
    ("kv", "Shares", "the Equity tab"),
    ("kv", "Mutual funds", "the MutualFunds tab (one row per fund) + log each purchase on MF_SIP"),
    ("kv", "Fixed deposits", "the FixedDeposits tab"),
    ("kv", "PPF", "the PPF tab   (list deposits on PPF_Ledger for exact interest — optional)"),
    ("kv", "EPF", "the EPF tab — copy the balance and date from your EPFO passbook, done"),
    ("kv", "Bonds", "the Bonds tab"),
    ("kv", "House, savings, insurance", "the Manual_Assets tab — type today's value, pick the Class"),
    ("kv", "Gold & silver", "the Gold_Silver tab — type grams, today's bullion rate fills in (your jeweller's rate wins if you type it)"),
    ("kv", "NPS", "the NPS tab — pick the scheme, type units from your CRA statement"),
    ("space",),
    ("tip", "Pick funds, shares and banks from the dropdown — the ID fills in for you."),
    ("tip", "Searching a long list? Type a few letters, press Enter, then open the dropdown."),
    ("tip", "Selling a fund? Log it on MF_SIP as a purchase with a minus amount."),
    ("space",),

    ("section", "🗂️", "Show only what you own"),
    ("text", 'Easiest way: "Update Portfolio" asks you — pick a number to show or hide any asset class.'),
    ("text", "(Or set Yes/No on the Settings tab; either way it takes effect on the next update run.)"),
    ("text", "Nothing is ever deleted — and a tab that still holds rows stays visible until you clear them,"),
    ("text", "so a hidden number can never lurk inside your totals."),
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
    ("kv", "Mergers & demergers", "old shares price as the new company, spun-off shares appear as their own row — cost and dates carried correctly"),
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
