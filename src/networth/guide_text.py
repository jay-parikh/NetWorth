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
     "Everything your family owns, in one file. You type what you own — it works out what it's worth."),
    ("space",),

    ("section", "🎨", "First, the colours"),
    ("legend",),
    ("text", "One rule, that's it: type only in the coloured cells. Grey looks after itself."),
    ("text", "The tabs are colour-coded too: navy = your totals, teal = each person, blue = you type, grey = automatic, gold = this Guide."),
    ("space",),

    ("section", "🚀", "Start in 4 steps"),
    ("step", "1", "On the Dashboard, type your family's names over the samples (Amit, Priya, Rahul)."),
    ("step", "2", "Every tab has a sample row or two. Type your own holdings straight over them."),
    ("step", "3", "Own gold, EPF, NPS, property…? Switch it on (Settings tab, or the update just asks) — an example is waiting inside."),
    ("step", "4", 'Save, close the file, and double-click "Update Portfolio".'),
    ("tip", "Picking a share, fund or bank? Choose it from the dropdown — its code fills in for you. Long list? Type a few letters, press Enter, then open the dropdown."),
    ("space",),

    ("section", "📒", "Where does each thing go?"),
    ("kv", "Shares", "the Equity tab. Bought the same share again? Just add another row — same name, its own quantity, price and date; each purchase keeps its own tax date"),
    ("kv", "Mutual funds", "the MutualFunds tab (one row per fund) — log each purchase on MF_SIP. Sold some? Add a row with a minus amount."),
    ("kv", "Fixed deposits", "the FixedDeposits tab"),
    ("kv", "PPF", "the PPF tab (want exact interest? list deposits on PPF_Ledger — optional)"),
    ("kv", "EPF", "the EPF tab — copy the balance and date from your EPFO passbook, done"),
    ("kv", "Bonds", "the Bonds tab"),
    ("kv", "Gold & silver", "the Gold_Silver tab — weigh it in grams, pick Gold or Silver. Jewellery? Add purity (22K = 0.916); coins and bars leave it blank."),
    ("kv", "NPS", "the NPS tab — pick the scheme, type your units from your statement"),
    ("kv", "House, savings, insurance", "the Manual_Assets tab — pick the Class (Property / Cash / Insurance / Other), type what it's worth today"),
    ("tip", 'Gold & silver price at the daily bullion rate. Prefer your jeweller\'s rate? Type it in "Rate override" — it wins. A gold bond (SGB)? Fill the ISIN from your statement — it prices like a share.'),
    ("space",),

    ("section", "📥", "Ten years of SIPs, no typing"),
    ("step", 1, 'Email yourself your fund statement: camsonline.com > Statements > "CAS – CAMS+KFintech". Pick "Detailed" and "Since inception", and choose a password you\'ll remember.'),
    ("step", 2, "Save the PDF from the email into the same folder as this file."),
    ("step", 3, 'Double-click "Update Portfolio". It spots the statement, asks for its password, shows you what it read — every fund checked against the statement\'s own balance — and types it all in for you. Sales and switches too.'),
    ("kv", "Shares too", "save your broker's tradebook or holdings file (CSV or Excel) in the same folder — buys, sales and holdings come in the same way, any broker whose file names its columns. Old paper shares at cost 0? One question dates them and the official 2018 value fills in."),
    ("tip", "Run it again anytime — nothing gets added twice. A newer statement just adds the new months. A fund that can't be read reliably is left out and says why — never guessed."),
    ("space",),

    ("section", "✨", "It quietly does the boring bits"),
    ("text", "You don't have to learn this file — you get to notice it. A stock splits and your share count just updates. A dividend lands and it's logged before your bank's SMS."),
    ("kv", "Splits & bonuses", '"Qty today" keeps your share count up to date on its own — your typed numbers never change'),
    ("kv", "Mergers & demergers", "old shares price as the new company at the right ratio; spun-off shares appear as their own row — costs and dates carried correctly"),
    ("kv", "Dividends", "every one your shares declare this year is logged on the Dividends tab, with a month-by-month chart — and it counts in your return figure (XIRR)"),
    ("kv", "PPF interest", "worked out exactly when you list deposits on PPF_Ledger — otherwise your typed balance is used"),
    ("kv", "A forgotten old cost", 'bought before Feb 2018 and don\'t know the price? Leave "Avg. cost" blank — the official 2018 value fills in'),
    ("kv", "Old paper shares", 'know only the company and how many? Type the quantity, leave "Avg. cost" blank, and put 31-01-2018 as the Buy date — the official 2018 value fills in and the tax maths is right'),
    ("kv", "Delisted shares", "kept at their last known price and shaded amber so you can spot them"),
    ("text", "To refresh it all: close the file and double-click \"Update Portfolio\". It backs up first, then updates every price and figure. Once a week is plenty."),
    ("space",),

    ("section", "🎯", "Make it yours"),
    ("kv", "Show only what you own", "the update asks you, or set Yes/No on Settings (it takes effect on the next update). Your choice wins: a hidden tab keeps its rows but stays out of every total — an amber line on the Dashboard reminds you; nothing is ever deleted."),
    ("kv", "Add a family member", 'the update asks "Add a new person?" — type a name and their tab is built. Use that name in the Owner column.'),
    ("kv", "Stay on track", 'give each class a Target % on Settings, and the Dashboard answers in colour — green "On target", or a red hint like "Move ₹1,20,000 out" — live, the moment you edit. The tolerance (±5 points) is yours to change.'),
    ("kv", "Reference lists", 'the stock, fund, bank and pension name lists behind the dropdowns stay tucked away — they look after themselves (don\'t edit them by hand). Curious? Set "Reference lists" to Yes on Settings.'),
    ("space",),

    ("section", "🔒", "Keep it private (optional)"),
    ("text", "Two switches on the Settings tab, both off until you want them, sharing one password:"),
    ("kv", "Privacy mask", "every number shows as ••• (names and dates stay). A curtain for nosy glances — not a safe. Forgot the password? Type RESET when the update asks."),
    ("kv", "Lock file", "the file itself needs the password to open, in Excel too. Real protection for a lost laptop or a shared folder."),
    ("text", 'Turn either on and run "Update Portfolio" — it walks you through choosing a password.'),
    ("text", "To see masked numbers: run the update, type y when it asks, then your password. The next update masks them again by itself (sooner? run with --lock)."),
    ("text", "⚠  The Lock has no reset and no recovery — that is what makes it real. Write the password down somewhere safe."),
    ("text", "Backups made before locking stay readable (delete the backups folder if that matters). Scheduled hands-free updates can't run while the file is locked."),
    ("space",),

    ("section", "🧾", "Selling & tax (optional)"),
    ("step", "1", "Sold shares? Type the sale on the Equity_Sells tab — one row per sale, straight from your contract note."),
    ("step", "2", "Then reduce the Quantity on the Equity tab — that tab is always what you own now."),
    ("step", "3", 'Switch "Capital gains report" to Yes on Settings and run the update. The Capital Gains tab shows STCG & LTCG per year, what is still tax-free, and when each holding turns long-term.'),
    ("text", "Fund sales need no extra typing — a minus Amount row on MF_SIP is enough."),
    ("text", "Indicative only — for planning, not for filing. Recorded sales also count in your return figure (XIRR)."),
    ("text", "Government changed a rate or the tax-free allowance? Edit the Tax_Rules tab (it ships filled in with the current law) — no new app version needed."),
    ("text", "Bought and sold the same day? That is speculative income — shown separately at your slab, never mixed into STCG / LTCG."),
    ("text", "Sold at a loss? It reduces that year's taxable gains the way the law allows — debt-fund losses count too, and a leftover short-term loss even counts against long-term gains."),
    ("space",),

    ("section", "🔤", "Words you'll see"),
    ("kv", "Return a year (XIRR)", "your yearly return — it counts when you invested, not just how much"),
    ("kv", "ISIN", "the code on your statement — it fills in from the dropdown, you rarely type it"),
    ("kv", "NAV", "the price of one unit of a fund"),
    ("kv", "PRAN / UAN", "your NPS number / your PF number — both are on your statements"),
    ("kv", "SGB", "Sovereign Gold Bond — a government bond that prices like a share"),
    ("kv", "Ex-date", "the date that decides who gets a dividend"),
    ("kv", "STCG / LTCG", "short- / long-term capital gains — profit on things you sold; long-term (held over a year) is taxed less, with a tax-free slice each year"),
    ("kv", "Grandfathering", "for shares bought before Feb 2018, tax counts the official 31-Jan-2018 value as your cost — usually kinder to you"),
    ("space",),

    ("section", "💡", "Good to know"),
    ("bullet", "Return figures (XIRR) are filled in by the updater — run it again after you change a holding."),
    ("bullet", "Add, delete or sort rows freely. Just don't rename the tabs — the Dashboard finds them by name."),
    ("bullet", 'Change "Inflation %" on the Dashboard to watch the 20-year Projection tab react instantly.'),
    ("bullet", "Nothing about your money ever leaves your computer."),
    ("space",),

    ("footer", "Questions, help and new versions:    github.com/jay-parikh/NetWorth"),
]
