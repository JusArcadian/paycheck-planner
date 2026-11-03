# Patch Notes

## v2.12.4 — 2026-10-22 (small fixes from the first real test)

### Bugs fixed
- Spending → Budget check said **"Over by $0"** when you were over by less than a dollar. Amounts under $1 now
  show cents ("Over by $0.26", "Close: $0.10 left").
- Setup: "Custom (I'll type my own %)" was cut off in the plan box. Column C is wider (20 → 29).

## v2.12.3 — 2026-10-22 (import fix, take 2)

### Bugs fixed
- The v2.12.2 fix didn't work: the import still stopped with "Not on the list". Google checks dropdowns when it
  SAVES the sheet, not when the script writes, so the try/catch never saw the error. Now the importer always lifts
  the dropdown rules first, writes, saves right away (`SpreadsheetApp.flush()`), then puts the rules back.

## v2.12.2 — 2026-10-22 (import no longer stops on "Not on the list")

### Bugs fixed
- **Import my data stopped halfway with "Exception: Not on the list - Please pick from the list."** when an old
  value wasn't on the new copy's dropdown list (a renamed category, an old choice name...). Now it keeps your
  value, puts the dropdown back (the cell gets a small red corner), and the "Done!" message lists those cells
  (up to 12) so you can pick again.

## v2.12.1 — 2026-10-22 (sort Money in by date)

### Features
- **Budget Tools → Sort Money in by date**: puts Money in (Date received → Notes) oldest first, blank rows last.
  "Sort Spending by date" now sorts just Spending. Both still sort by themselves every time the sheet opens.
- "How to use" tab: a "Sort by date" line in the Money in and Spending sections, and a "Got a refund?" line
  (type it as a minus amount in the same category, not under Money in).

## v2.12.0 — 2026-10-22 (bank holidays)

### Features
- **Bank holidays move paydays too.** Setup C15 is now "Payday on a weekend or bank holiday?":
  Workday before / Workday after / No change. Christmas on a Friday payday → paid Thursday Dec 24.
  Labor Day Monday → paid the Friday before.
- Every-week and every-2-weeks pay only moves for holidays (it always lands on the weekday you picked).
- **Engine → Bank holidays** (columns E:G): the 11 Federal Reserve holidays for 2026-2030, worked out by
  formula (like "3rd Monday of January"). Sunday holidays move to Monday; Saturday ones stay (banks open Friday).
  The list is in config/product.json → bank_holidays. Checked against the Fed's published dates.

### Buffs
- Copies made with v2.11 that say "Friday before" still work (treated as "Workday before").
- Budget Tools → Send my data to the app now carries "Which days?" and the weekend/holiday answer (app v0.7.0).

## v2.11.0 — 2026-10-22 (paydays on set days + weekend rule)

### Features
- **"Paid twice a month? Which days?"** (Setup C14): From first payday (old way: 15 days apart), 1st & 15th,
  15th & 30th, or 15th & last day. "The 30th" in February becomes Feb 28 (or 29).
- **"Payday lands on a weekend?"** (Setup C15): Friday before, Monday after, or No change.
  Used for Semi-monthly and Monthly pay (weekly and every-2-weeks pay always lands on the same weekday).
- The Engine payday list now has two columns: **Official payday** (pay stub date) and **Money lands**
  (after the weekend rule and the bank-early days).

### Buffs
- "Paychecks per year" moved into the tip next to "How often" ("That's 24 paychecks a year"), freeing the row.
- Reports → Paycheck by paycheck reads the Engine payday list instead of working out each date again (lighter).
- The smaller first check is matched to the first payday in the list, so the weekend rule applies to it too.
- Import my data copies the two new answers (older copies don't have them, so they're skipped).

### Nerfs
- "Does your bank pay you early?" moved from row 15 to row 16 (the importer finds it by its label).

## v2.10.0 — 2026-10-22 ("How to use" guide tab)

### Features
- **New "How to use" tab** (first tab): a plain-English guide in 6 short sections: set it up once, every time
  money comes in, every time you spend, where each number comes from (Expected, Received, Bills & needs, Flexible,
  Saved, Left over), what each tab is for, and quick fixes when something looks wrong.
- The sheet still opens on Setup; the guide is one click to the left.

### Buffs
- Clearer Setup tips: "First paycheck" = your very first check if it was only part of a full one;
  "first payday" = the day your FIRST check landed, even if it was a half check.

## v2.9.0 — 2026-10-22 (every month list runs Jan 2026 to Dec 2028)

### Features
- **Month lists run Jan 2026 to Dec 2028 (36 months)** instead of "12 months from your first payday".
  Set once in `config/product.json` → `calendar_start` ("2026-01") and `calendar_months` (36).
  - Dashboard month picker: all 36 months.
  - Reports → Month by month: 36 rows plus a **2026 total, 2027 total, 2028 total** row.
  - Goals → Savings scorecard: 36 rows plus one total per year. "Saved this year" / "Goal this year" start over
    every January. Months before your first payday (or before "Start tracking from") say "Before you started"
    and don't count.
  - Accounts → Net worth log: 36 months.
- **Goals cards are for THIS year** ("SAVED THIS YEAR", "Goal for 2026: ...", "by Dec 31, 2026").
- **This month is highlighted** (soft color) in Reports, the Goals scorecard and the net worth log.
- **Reports → Paycheck by paycheck** lists every paycheck up to Dec 2028 (it was one year).
- **Reports → Each category by month** shows Jan to Dec of the year you pick on the Dashboard.

### Buffs
- The Engine's payday list covers 4 years (was 2), so paydays keep counting through Dec 2028.
- **Import my data** copies the net worth log **by month** (Aug 2026 goes to Aug 2026), so older copies whose
  list started on your first payday land in the right rows.

### Nerfs
- Reports sections C and D swapped: **Day by day** is now C, **Paycheck by paycheck** is D (it's the longest list,
  so it goes last).

### Bugs fixed
- Bills expected and Savings goal no longer count months before you started (they showed in the year total).

## v2.8.0 — 2026-10-22 (budget check on every purchase, OK it or cover it, sort by date)

### Features
- **Budget check on every purchase** (Spending, column F): how that budget is doing after this purchase:
  "✔ $120 left in Groceries", "Close: $15 left" (yellow at 90%), "⚠ Over by $19 in Fun: OK it or cover it" (red).
  Categories with a monthly goal check themselves; the rest check their group (Wants, Fun...).
- **OK'd?** (column G): pick a reason (Planned it, Emergency, Sale / good deal, Bill was higher, Special day / gift,
  Other). The row and the Dashboard alert turn into a calm "Noted: $29 over · Special day / gift".
- **Cover with** (column H): "Last month's leftover" adds what that budget didn't spend last month.
  "Extra money in" = log it under Money in and pick the budget under the new **For** column.
- **Money in → For (optional):** extra money for one budget (groups or categories) adds to it that month.
  The Dashboard group "Plan (+ extra)" includes it; safe-to-spend adds it on top of your paycheck.
- **Sort categories:** Budget Tools → "Sort categories A to Z" or "Sort categories by group" (Setup's categories
  table; empty rows go to the bottom). Every category dropdown was already A to Z.
- **Sort by date:** Budget Tools → "Sort Spending by date". It also sorts itself every time the sheet opens
  (Spending and Money in), never while you're typing.
- Over-budget rows light up their OK'd? / Cover with boxes until you fill one in.

- **Budget Tools → Send my data to the app:** makes one file with everything you typed in the sheet, in the app's
  format. In the app: Setup → Your data → Load backup. Nothing is uploaded anywhere.
- **New Phone view tab (phone + tablet friendly)**, right after the Dashboard: one narrow column that fits a phone screen:
  the alert banner, SAFE TO SPEND + per day, a "Can I afford it?" box, this month's pay/spent/left, money left in
  each group (colored), this month's bills, and the top 5 alerts. (A spreadsheet can't tell what device it's on,
  so this is its own tab. The app changes its layout by itself.)
- Spending keeps the header rows AND the Date column frozen while you swipe sideways.

### Bugs fixed
- A blank copy said "On track: $0 left", "✔ Goal reached", "✔ Ahead", "All set ✔", "On track ($0 under)" and
  "✔ You own more than you owe". Now it says what to do: "Add your pay on Setup", "No bills yet"...
  The Dashboard banner says "👋 Start here: type your pay and first payday on the Setup tab."
- Status colors now fill the whole merged box (like "How it's going" K:M), so they line up with the header.

### Buffs
- "Import my data" finds Spending and Money in columns by their header names (columns moved this version).
- Anything starting with ✔ is always green (before, "✔ Covered" turned red because it contains "over").
- "✔ Right on budget" when a bill is paid exactly.

## v2.7.0 — 2026-10-22 (paycheck setup, paid early, safe to spend, payment arrangements)

### Features
- **Your pay, in a clearer order:** How often → First paycheck (if smaller) → **Full paycheck (ONE paycheck)**
  → **Monthly income** (automatic, with the after-tax amount) → before-tax deductions → first payday.
- **"Does your bank pay you early?"** (on payday, or 1–5 days early). Every payday, bill and report uses
  the day the money actually lands.
- **Bills tab, new section A: Before your next paycheck.** Money from the paycheck you're on − what you've
  spent since it landed − bills and plan payments due before the next one = **SAFE TO SPEND**, plus a
  per-day amount.
- **"Can I afford it?"** box: type what you want to spend (or plan to). It says "✔ Yes, you'd still have $X"
  or "✖ Not safe: you'd be $Y short for Rent (due Nov 1)".
- **Spending tab, row 3:** the safe-to-spend line, so you see it while you log.
- **Payment arrangements** (Bills, section C): who/what, total owed, each payment, how often, first payment,
  paid so far → next payment and status (On track / Due in N days / Late: $X behind / ✔ Paid off).
  Plan payments count in "safe to spend" and in "Which paycheck pays which bill".
- **New alerts:** spending leaves you short for upcoming bills (red), a payment plan is late (red),
  a plan payment is due soon (yellow).
- **Every pick-list is A to Z** (categories, groups, types, how often...). Lists with a natural order
  ("on payday, 1 day early, 2 days early...") keep it.

### Buffs
- **"Import my data" finds each section by its heading**, not fixed cells, so old copies import correctly
  even though Setup moved down 3 rows.
- Bills tab sections are now: A Before your next paycheck · B Bills this month · C Payment arrangements ·
  D Which paycheck pays which bill · E Yearly costs.

## v2.6.0 — 2026-10-22 (bills, debts, first paycheck, net worth + polish)

### Features
- **First paycheck, if it was smaller** (Setup C13). Started mid pay period? Type that first check
  (e.g. $500 of a $1,000 week). Paycheck #1 uses it; every check after is the full amount. Flows into
  Pay this month, Left this month, spend per day, alerts, Reports and the Goals scorecard.
- **Bill due dates** (Setup → Categories → new "Due day" column, 1–31).
- **New Bills tab:**
  - Cards: bills this month, paid so far, still to pay (and how many overdue), next bill.
  - A. Bills sorted by due date: amount (the category's monthly goal), paid, left, which paycheck pays it,
    status: ✔ Paid / Partly paid / Due in N days / Due today / Overdue by N days.
  - B. **Which paycheck pays which bill:** each paycheck this month (and last month's last one), its
    take-home, the bills it covers, and what's left. Red if a paycheck can't cover its bills.
  - C. **Yearly costs:** car registration, holidays, renewals... cost ÷ months left = set aside per month.
- **New Debts tab (payoff planner):** list debts (balance, APR, minimum), add an extra amount, pick
  Snowball (smallest balance first) or Avalanche (highest interest first). Shows payoff order, months,
  payoff date and interest per debt; debt-free date; interest saved vs. minimums only; what you owe at the
  end of each year. Warns when a minimum is too low to ever pay a debt off.
- **Net worth** (Accounts tab, section D): savings & investing + other things you own − debts − other money
  you owe. Monthly log with change and bars. Budget Tools → "Save this month's net worth" fills it in
  (and the monthly auto-update now saves it on the 28th).
- **New alerts:** Overdue bill (red), a paycheck short for its bills (red), bill due within 5 days (yellow).
- "Import my data" also copies the first paycheck, due days, yearly costs, debts and net worth.

### Buffs (look + feel)
- Every Total row has the same tinted band. Money that goes negative in "Left"/"Left over" turns red.
- Dashboard sections are lettered A/B/C like the other tabs. Small indent on text in tables.
- Status colors cover more words (overdue, partly paid, due in, ready, too low...); first match wins.
- The Alerts list is colored by how serious each alert is (not by its first word).
- Tabs now: Setup · Spending · Dashboard · Bills · Goals · Accounts · Debts · Reports.

### Nerfs
- **The master template starts empty.** No example pay, first payday, tax rates, spending, goals, accounts,
  debts or yearly costs: every buyer types their own numbers. Group and category names stay as a
  starting point. With no first payday yet, the plan year is this January and no paydays are counted.

### Buffs (lined up)
- Dashboard: the alert banner runs the full width; an **ALERTS card** sits beside the 4 cards on the same rows
  (red / yellow / green); the full alert list sits beside the groups table, one alert per row.
- Cards on Goals and Accounts stretch the full width of their tables. White lines only BETWEEN cards,
  so outer edges line up with everything above and below.
- Every label is vertically centered (Excel's default is bottom), so text next to each other lines up.
- Grey tips and yellow boxes have a small indent so they never touch. Debts inputs are the same width.
- An overflow scan found and fixed every spot where text ran into the next cell.

### Tools
- **Auto-cleanup of output/:** every build keeps the newest 3 versions and deletes older spreadsheets
  (owner's standing permission). Change the number with `keep_output_versions` in config/product.json.
- New `tools/clean_output.py` (`--preview` shows what it would delete) and menu option 7.

## v2.5.0 — 2026-10-22 (overspending alerts)

### Features
- **Alerts on the Dashboard.** Row 3 is a banner: green "No alerts", yellow "heads up", or red "over budget",
  with the most important alert in it. The full list (up to 12) sits to the right of the cards.
- **What triggers an alert** (for the month picked at the top):
  - Red: you spent more than you expect to get; a bills/flexible group is over its plan
    ("That's money your savings goals don't get"); a category is over its own monthly goal.
  - Yellow: a group is 90%+ used; flexible spending is running ahead of the calendar
    (e.g. 80% used with 50% of the month gone); not enough left to finish this month's savings;
    a savings goal on the Goals tab is behind (after its first month).
- **Google Sheets pop-up:** log a purchase on Spending that puts you over and a pop-up appears
  (bottom right, on a computer).
- **Optional morning emails:** Budget Tools → "Email me alerts every morning". Only sends on days with an alert.
  Also "Check my alerts now" and "Stop alert emails".
- Settings in `config/product.json`: `alert_almost_pct` (0.9) and `alert_pace_buffer` (0.10).

### Buffs
- The "Almost out" status on the groups table now uses the same 90% setting.

## v2.4.1 — 2026-10-22 (dropdowns in Google Sheets + uniform statuses)

### Bugs fixed
- **Dropdowns that read a list from the same tab disappeared in Google Sheets.** That was Category on
  Spending, Group on Setup, the month picker on Dashboard, "Counts money logged to" on Goals and
  "Deposits logged to" on Accounts. The file wrote the list as `=$J$5:$J$34`; Excel allows the "=",
  Google Sheets quietly throws the whole dropdown away. Now it's written as `$J$5:$J$34`.
- **"How it's going" didn't look uniform.** "Still to put away", "Coming up" and similar had no color
  while "On track" was green. They now get a calm blue, so every row is colored.

### Features
- **Date pickers.** Date cells (Spending date, Money in date, first payday, goal dates, start tracking,
  account "as of") only accept real dates. In Google Sheets, double-click the cell for a calendar.

## v2.4.0 — 2026-10-22 (Accounts: high-yield savings + investing)

### Features
- **New Accounts tab** (between Goals and Reports) for high-yield savings (SoFi, Ally, Capital One...) and
  investing accounts (Fidelity, Robinhood, Schwab...). Up to 8 accounts.
- Type each account's balance, the date you checked it, its yearly rate and how much you add each month.
- **Now (estimate)** grows the balance at its rate since that date, plus deposits you logged on Spending
  to the category picked in "Deposits logged to".
- **Earns per month**, and what it could be worth in **1, 5, 10 and 30 years** with your monthly deposits.
- 4 cards: total now, earning per month, worth in 10 years (put in vs. growth), and the
  **high-yield bonus**: how much more your HYSA earns per year than the 0.38% national average.
- **How your money grows** table: 1–30 years, what you put in vs. what it's worth, with a bar.
- Switch: "Show future amounts in today's dollars" (takes inflation out).
- "Good to know" notes per account type (FDIC, investing can lose money, retirement penalties).
- "Import my data" also copies the Accounts tab.

### Buffs
- Example accounts: High-yield savings 4% + $100/mo, Index fund 7% + $50/mo (balances start at $0).

## v2.3.2 — 2026-10-22 (start tracking from)

### Features
- Goals tab: **Start tracking savings from (month)** (cell G9). Started mid-year? Months before it show
  "Before you started" and don't count against you. Cards, scorecard, yearly goal and pace all use it.
  Blank = the whole plan year (same as before). Matches the Paycheck Planner app.
- The example goals now start the month you start tracking.
- "Import my data" also copies the new tracking month.

### Bugs fixed
- Someone starting in October looked "Behind by $3,700" because January to September counted as missed savings

### Buffs / Nerfs
- Card text "Whole-year goal" is now "Goal for the year" (it only counts the months you track)

## v2.3.1 — 2026-10-22 (easy upgrades)

### Features
- Google Sheets **Budget Tools → Import my data from an old copy**: buyers move everything they typed
  (settings, groups, categories, spending, Money in, goals) into a new version in one step.
  Steps for you and buyers: `google-sheets/SETUP.md` → Steps 5 and "How buyers move to a new version"

- `menu.py`: option 4 now reads **Build + push to Google Sheets (Update)**
- Setup questions shortened so each fits on one line (before-tax deductions, own tax rate, own inflation guess);
  the extra detail moved to the grey hint next to the box
- Goals tab tidied: every header fits on 2 lines, even column widths, taller rows so nothing gets cut off,
  shorter "Behind by $X · save $Y/mo" status

### Bugs fixed
- Bold style: the Ahead/Behind card showed red text on a purple block (hard to read). Red now only shows in the card style
- Building while a template is open in Excel crashed with a long `PermissionError`.
  Now it says: "Can't save … it's open in Excel. Close that file, then run this again."

### Buffs / Nerfs
- None

## v2.3.0 — 2026-10-22 (Goals: am I on track?)

### Features
- New **Goals** tab (between Dashboard and Reports):
  - **4 big cards**: Saved this plan year · Your goal so far · Ahead or behind (+/− $) · On pace to save by year end
  - **Your savings goals** (up to 8): target, start + target date, already saved, which category counts.
    Shows saved so far, a progress bar, % done, what you *should* have by today, ahead/behind by how much,
    and how much to save each month to finish on time
  - **Savings scorecard, month by month**: money in (uses split payments from Money in), bills expected vs paid,
    flexible spending, put into savings, left over, savings goal, over/under the goal, savings rate,
    and a running "saved this year vs goal this year" total
  - Switch: **count leftover money as saved** (only after the month ends)
- Setup's how-it-works line now says "Check your Dashboard and Goals"

### Bugs fixed
- `backup.py` now removes characters Windows can't use in file names (like `:`) from the backup note

### Buffs / Nerfs
- Status colors understand more words: "missed" = red, "behind" / "to go" = amber, "met" = green

## v2.2.3 — 2026-10-22 (owner's picks)

### Features
- Google Sheets now gets the **bold color blocks** look (`"push_style": "bold"`)
- The example purchase on Spending is dated **today** (`=TODAY()`), so new buyers see it on this month's Dashboard right away

### Bugs fixed
- The example purchase was dated Nov 3, 2026, so it didn't show on the Dashboard in other months

### Buffs / Nerfs
- None

## v2.2.2 — 2026-10-22 (shortcuts)

### Features
- `menu.py`: type `python menu.py` and pick a number (backup, update inflation, build, build + push, practice push, listing copy)
- `docs/SHORTCUTS.md`: every link (master Sheet, Google Cloud keys, Gumroad, Etsy, GitHub, Bitwarden…) and keyboard shortcut (Windows, PyCharm, Claude Code, Sheets/Excel)
- Study guide Chapter 13 + the `.json.json` rename trap in "When things break"

### Bugs fixed
- None

### Buffs / Nerfs
- None

## v2.2.1 — 2026-10-22 (ready for Git + GitHub)

### Features
- `.gitignore`: Git will never save secrets (`.env`, Google key files), your real budget (`reference/`), buyer emails, `.venv`, `output/` or `backups/`
- `CLAUDE.md`: commit after each change; only push to GitHub when the owner says so

### Bugs fixed
- Caught before it shipped: comments on the same line as a `.gitignore` rule silently break it (it would have uploaded `reference/`). All comments now sit on their own lines

### Buffs / Nerfs
- None

## v2.2.0 — 2026-10-22 (auto-update Google Sheets)

### Features
- `tools/push_to_sheets.py`: puts the newest build INTO the master Google Sheet. Same file, same link, buyers' copies untouched. `--dry-run` shows what it would do without changing anything
- `python tools/build_template.py --push` = build everything, then push in one command
- `docs/GOOGLE_PUSH_SETUP.md`: one-time Google Cloud setup, step by step
- `docs/TOOLKIT.md`: programs, APIs and automation ideas for this and future projects
- `config/product.json` → `"google"`: master Sheet ID, its name, and which theme/style gets pushed

### Bugs fixed
- Reports: the frozen-panes line ran down through every section title. Now only the title rows are frozen

### Buffs / Nerfs
- `backup.py` and `CLAUDE.md` now also protect `google_credentials.json` and `google_token.json`

## v2.1.0 — 2026-10-22 (money in + real monthly pay)

### Features
- **Real pay for every month, on any schedule.** The sheet counts how many paydays land in each month. Paid every 2 weeks? It knows 2 months a year have 3 paychecks (e.g. $3,000 instead of $2,000)
- **Money in log (optional)** on the right side of Spending: log each deposit when it actually lands, even part of a paycheck or side money
- Dashboard: first card is now **PAY THIS MONTH** (paydays × take-home per check), with "average month" underneath
- Dashboard: **Left this month** now also shows **Cash in hand** (received − spent) once you log money in
- Dashboard: new line **"Money in: $X received of $Y expected · $Z still coming"**
- Reports → Month by month: new **Paydays**, **Expected pay** and **Actually received** columns
- Reports → Paycheck by paycheck: new **Received** column, and **"Partly paid: $X still coming"** status for split paychecks
- Setup: hint under Your Pay pointing to Money in
- Engine: a list of every payday for 2 years (feeds the payday counts)

### Bugs fixed
- `backup.py` would have zipped PyCharm's new `.venv` folder (thousands of library files). Now skipped

### Buffs / Nerfs
- Buff: "Left over" in Reports uses money you actually received when you log it; otherwise it uses expected pay

## v2.0.0 — 2026-10-22 (the big redesign)

### Features
- **4 tabs instead of 6:** Setup, Spending, Dashboard, Reports. The math lives on a hidden Engine tab
- **Setup is a simple question form** with a one-line hint beside every question
- **Budget plan menu:** 50/30/20, 60/20/20, 70/20/10, 80/10/10, Starter, or Custom (your own %)
- **Up to 6 groups**, each with a kind (Bills & needs / Flexible spending / Saving & investing) and a "grows with inflation" switch
- **Up to 30 categories** with optional monthly goals (pre-filled with common ones like Rent, Groceries, Eating out)
- **Dashboard:** 4 big numbers (take-home, spent, left, you can spend per day), spent today, days left, group table with progress bars and plain-English status, a "left to spend" donut for every group, category cards per group (red when over a goal), and taxes & inflation explained in plain sentences
- **Reports:** month by month, each category by month, paycheck by paycheck (weekly, biweekly, semi-monthly or monthly paydays), and day by day pace
- **Two looks:** clean card style (default) and bold color blocks (`_bold` files)
- Ideas borrowed from the owner's old planner: buckets of categories, donut charts, daily allowance, spent today, paycheck/monthly/category reports

### Bugs fixed
- Dropdowns could vanish when converting to Google Sheets. Every dropdown now has its choices typed in or reads a list on its own tab
- Dashboard label "Every $1 you spend costs…" was cut off. Replaced by full plain-English sentences
- Donut rings now show money LEFT in color, so they "drain" as you spend

### Buffs / Nerfs
- Buff: categories 10 → 30, spending rows 500 → 1000
- Nerf: removed the visible Tax Tables and Start Here tabs (moved into Engine and the top of Setup)

## v1.0.1 — 2026-10-22 (project setup)

### Features
- `CLAUDE.md`: Claude in PyCharm (and Remote Control on your phone) now follows the same rules every time: backup → change → patch notes → study guide → sources → backup, then a summary, code walkthrough and test checklist in chat
- `tools/SOURCES.md`: one page listing every tax number, inflation figure, API, platform rule and library, with links and re-check dates
- Study Guide Chapter 8: working with Claude in PyCharm and from your phone

### Bugs fixed
- `backup.py` was zipping PyCharm's `.idea` settings folder. Now skipped.

### Checked
- 2026 tax brackets for all 3 filing statuses re-verified against the IRS release (single, married filing jointly) and NerdWallet (head of household). All match.

### Buffs / Nerfs
- None

## v1.0.0 — 2026-10-21 (first build)

### Features
- Budget template with custom categories (up to 10), your own % split, and a check that warns if it isn't 100%
- Tax: Auto mode (2026 US federal brackets + Social Security/Medicare + state %) or Manual mode (your own rate)
- Inflation: Auto (3.7%, US CPI for Sep 2026) or Manual; per-category "grow with inflation" switch and next year's targets
- "Real cost" column: what each category's spending cost in pre-tax pay
- Expense log (500 rows) with category dropdowns; monthly dashboard with chart and Over/Almost/On track status
- 3 themes: Sage, Midnight, Blush
- Google Sheets script: Budget Tools menu, one-click and monthly inflation updates from the BLS API
- Tools: template builder, inflation updater, listing kit generator (optional ChatGPT/Claude), backup script
- Delivery: Gumroad sale webhook (secret key check, sales log, welcome email) and sales report
- Docs: README, launch guide, study guide, prompt pack

### Bugs fixed
- Dashboard month cell showed `###` (column too narrow). Widened it and shortened the date format.

### Buffs / Nerfs
- None yet

### Known limits
- Social Security wage cap, tax credits and itemized deductions are not modeled (fine for most first-jobbers).
- The inflation scripts couldn't be tested live from the build machine. The math was tested with the real Sep 2026 numbers (gives 3.7%).
