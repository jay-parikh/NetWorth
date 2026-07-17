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
    ("text", "The tab strip is colour-coded too: navy = overview, teal = family members, "
             "blue = you type here, grey = automatic."),
    ("space",),

    ("section", "🚀", "Do this first"),
    ("step", "1", "On the Dashboard, replace the sample names (Amit, Priya, Rahul) with your family."),
    ("step", "2", "Every tab has a few sample rows showing how it works. Replace them with your own."),
    ("step", "3", "Own gold, EPF, NPS, property…? Switch those tabs on (Settings tab, or the update window asks) — a worked example is waiting inside each."),
    ("step", "4", 'Save, close the file, and double-click "Update Portfolio".'),
    ("space",),

    ("section", "📒", "Where each holding goes"),
    ("kv", "Shares", "the Equity tab"),
    ("kv", "Mutual funds", "the MutualFunds tab (one row per fund) + log each purchase on MF_SIP"),
    ("kv", "Fixed deposits", "the FixedDeposits tab"),
    ("kv", "PPF", "the PPF tab   (list deposits on PPF_Ledger for exact interest — optional)"),
    ("kv", "EPF", "the EPF tab — copy the balance and date from your EPFO passbook, done"),
    ("kv", "Bonds", "the Bonds tab"),
    ("kv", "House, savings, insurance", "the Manual_Assets tab — pick the Class (Property / Cash / Insurance / Other), type today's value"),
    ("kv", "Gold & silver", "the Gold_Silver tab — see the three steps below"),
    ("kv", "NPS", "the NPS tab — pick the scheme, type units from your CRA statement"),
    ("space",),
    ("tip", "Pick funds, shares and banks from the dropdown — the ID fills in for you."),
    ("tip", "Searching a long list? Type a few letters, press Enter, then open the dropdown."),
    ("tip", "Selling a fund? Log it on MF_SIP as a purchase with a minus amount."),
    ("space",),

    ("section", "🥇", "Adding gold or silver — three steps"),
    ("step", "1", 'Weigh it in grams and describe it: "Gold coins, 2 x 10 g", "Silver bar, 1 kg", "Jewellery, 40 g".'),
    ("step", "2", "Pick the Type (Gold / Silver). Jewellery? Add its purity: 22K = 0.916. Coins and bars: leave purity blank."),
    ("step", "3", "Done — today's value appears at the daily bullion rate on the next update. Prefer your jeweller's rate? Type it in Rate override; it wins."),
    ("text", "Sovereign Gold Bonds are even easier: pick Type SGB and fill the ISIN from your statement — they price like shares."),
    ("space",),

    ("section", "🗂️", "Show only what you own"),
    ("text", 'Easiest way: "Update Portfolio" asks you — pick a number to show or hide any asset class.'),
    ("text", "(Or set Yes/No in the Show? column of the Settings tab; either way it takes effect on the next update.)"),
    ("text", "Your choice wins: a hidden tab keeps its rows but is left out of every total, and one amber"),
    ("text", "line on the Dashboard reminds you what's hidden — nothing is ever deleted."),
    ("space",),
    ("text", "A few reference tabs stay tucked away: the stock, fund, bank and pension name lists that"),
    ("text", 'feed the dropdowns, and the actions history. Curious? Set "Reference lists" to Yes on Settings.'),
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

    ("section", "🔒", "Keep the numbers private (optional)"),
    ("text", "Two switches on the Settings tab, both off until you want them, sharing one password:"),
    ("kv", "Privacy mask", "every number shows as ••• (names and dates stay). A curtain against people "
           "glancing at your screen — not a safe. Forgot the password? Type RESET in the update window."),
    ("kv", "Lock file", "the file itself is encrypted — Excel asks for the password just to open it, "
           "and updates need it too. Real protection for a lost laptop or a synced folder."),
    ("text", 'Turn either on, run "Update Portfolio", and it walks you through choosing a password.'),
    ("text", "To see masked numbers: run the update and type the password — press Enter next time"),
    ("text", "(or run with --lock) and the mask goes straight back on."),
    ("text", "⚠  The Lock has no reset and no recovery — that is what makes it real. Write the"),
    ("text", "password down somewhere safe. Backups made before locking stay readable; delete the"),
    ("text", "backups folder if that matters. Scheduled hands-free updates can't run while locked."),
    ("space",),

    ("section", "💡", "Good to know"),
    ("bullet", "Return figures (XIRR) are filled in by the updater — run it again after you make changes."),
    ("bullet", "Add, delete or sort rows freely. Just don't rename the tabs — the Dashboard finds them by name."),
    ("bullet", 'Change "Inflation %" on the Dashboard to watch the 20-year Projection react instantly.'),
    ("bullet", "The hidden reference tabs (fund & stock lists) look after themselves — don't edit them by hand."),
    ("bullet", "Nothing about your money ever leaves your computer."),
    ("space",),

    ("footer", "Questions, help and new versions:    github.com/jay-parikh/NetWorth"),
]
