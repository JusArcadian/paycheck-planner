# Study Guide

Everything in this kit, explained simply. Read a chapter when you need it, not all at once.
New things get added here every time the kit changes.

**Chapters**
1. The big picture
2. The money math inside the spreadsheet
3. How `build_template.py` works
4. APIs, and the inflation scripts
5. Webhooks, `.env` files and secrets
6. When things break
7. Glossary
8. Working with Claude in PyCharm and from your phone
9. How v2 works (the redesign)
10. Money in and real monthly pay (v2.1)
11. Pushing to Google Sheets (v2.2)
12. Git and GitHub
13. menu.py and shortcuts

---

## Chapter 1 — The big picture

**Like you're 5:** You're running a lemonade stand. The lemonade is the spreadsheet. The scripts are robot helpers: one makes the lemonade (build), one writes the sign (listing), one checks lemon prices (inflation), and one says "thank you!" to every customer (webhook).

```
config/*.json ──► build_template.py ──► output/*.xlsx ──► Gumroad / Etsy
      │                                                        │
      └──► make_listing.py ──► listing_kit.md          sale happens
                                                                ▼
                                          gumroad_webhook.py ──► welcome email + sales log
```

**Why configs?** The product's settings (name, colors, tax brackets) live in `.json` files. Scripts read them. To change something, you edit one small file and rebuild. You don't dig through code.

---

## Chapter 2 — The money math inside the spreadsheet

### 2.1 Take-home pay
**Like you're 5:** Your paycheck is a pizza. Before you get a slice, the government takes some slices (taxes), and your 401k takes some (pre-tax deductions). Take-home is what's left on your plate.

```
Take-home = Gross pay − Pre-tax deductions − Total taxes
```
Example: $1,200 every 2 weeks = $31,200/year. Taxes ≈ $5,199. Take-home ≈ $26,001/year ≈ $2,167/month.

### 2.2 How the federal tax formula works (the tricky one)
US tax is *progressive*: you pay 10% on the first chunk, 12% on the next chunk, and so on. Only the dollars **inside** each bracket get that bracket's rate.

The Dashboard (cell B10) does it in one formula:
```
=SUMPRODUCT( (Taxable > Floors) * (Taxable − Floors) * RateSteps )
```
- **Floors** = where each bracket starts (0, 12,400, 50,400…)
- **RateSteps** = how much the rate *jumps* at each floor (10%, then +2%, then +10%…)

**Walkthrough with real numbers** (taxable income = $15,100, Single):

| Floor | Above it? | Dollars above floor | Rate step | Adds |
|---|---|---|---|---|
| 0 | yes | 15,100 | 10% | $1,510 |
| 12,400 | yes | 2,700 | +2% | $54 |
| 50,400 | no | — | — | $0 |
| **Total** | | | | **$1,564** |

Same as doing it by hand: 10% × 12,400 + 12% × 2,700 = 1,240 + 324 = $1,564. ✔

`(Taxable > Floors)` gives TRUE/FALSE, which act like 1/0, so brackets you haven't reached add nothing.

### 2.3 "Every $1 you spend costs this much of your pay"
If 17% of your pay goes to tax, you keep 83 cents of every dollar earned. So spending $1 means you had to earn `1 ÷ 0.83 = $1.20`. That's the **Real cost** column: a $64 grocery trip really took $77 of your pay.

### 2.4 Inflation
**Like you're 5:** If a candy bar costs $1 today and inflation is 3.7%, next year it costs about $1.04. Same candy, more money.
- **Next year's target** = this year's target × (1 + inflation). Only for categories you set to "Yes."
- Savings/Investing are usually set to "No": you're choosing that amount, prices don't force it up.
- **Buying power next year** = take-home ÷ (1 + inflation). That's what your same paycheck will feel like.

### 2.5 Monthly spending (SUMIFS)
```
=SUMIFS(Amounts, Categories, "Needs", Dates, ">="&MonthStart, Dates, "<"&EDATE(MonthStart,1))
```
In plain words: "Add up the amounts where the category is Needs AND the date is on or after the 1st AND before the 1st of next month." `EDATE(date, 1)` means "same day, one month later."

### 2.6 Why so many `IF(A24="","", …)`
Empty category slots would otherwise show `0` or errors. This says: "if there's no category name, show nothing."

---

## Chapter 3 — How `build_template.py` works

The script uses **openpyxl**, a Python library that writes Excel files. One function builds each tab:

```python
def build(theme, style):
    look = Look(theme, style)      # 1. pick colors for this theme + style
    wb = Workbook()                # 2. a blank Excel file
    build_setup(wb, look)          # 3. one function per tab
    build_spending(wb, look)
    build_engine(wb, look)         # hidden math tab (made before Dashboard: its charts read it)
    build_dashboard(wb, look)
    build_reports(wb, look)
    wb.save(path)                  # 4. save it
```

**Writing a formula** is just putting text that starts with `=` in a cell:
```python
ws["B8"] = "=B4*B6"      # gross per year = pay per check × checks per year
```
Excel calculates it when the buyer opens the file. We never type the answer ourselves, so it updates when inputs change.

**Dropdowns** are "data validation":
```python
dropdown(ws, "C9", ["Weekly", "Biweekly", "Semi-monthly", "Monthly"])   # C9 can only be one of these
```

**Themes and styles:** the loop at the bottom builds one file per theme × style in `product.json`:
```python
for t in themes:
    for s in styles:
        print("Built", build(t, s))
```
To add a theme, add one line to `"themes"` in the config (pick 3 hex colors) and rebuild. No code changes.

---

## Chapter 4 — APIs, and the inflation scripts

**Like you're 5:** An API is a drive-through window for data. You pull up (send a request to a web address), ask for something specific, and they hand you a bag (usually JSON).

The BLS (US Bureau of Labor Statistics) has a free API. We ask for series `CUUR0000SA0`, the main price index (CPI). It sends back numbers like:
```json
{"year": "2026", "period": "M09", "periodName": "September", "value": "336.940"}
```
Inflation = this month ÷ same month last year − 1:
```
336.940 ÷ 324.800 − 1 = 0.037  →  3.7%
```

- **Python version** (`tools/update_inflation.py`) is for you, the seller. It updates `product.json` so new Excel files ship with a fresh number.
- **Apps Script version** (`google-sheets/inflation_autoupdate.gs`) runs *inside* each buyer's Google Sheet. It's JavaScript. `UrlFetchApp.fetch` is Google's way of calling an API. `onOpen()` runs automatically when the sheet opens and adds the menu.

Excel files can't call APIs without macros (which scare buyers), so Excel buyers get your monthly number plus a Manual option.

---

## Chapter 5 — Webhooks, `.env` files and secrets

**API vs webhook, like you're 5:**
- API = **you** call the pizza place to ask "is my pizza ready?"
- Webhook = the pizza place **calls you** when it's ready.

Gumroad's "Ping" is a webhook. When someone buys, Gumroad sends their info to your web address, and `gumroad_webhook.py` reacts:
```python
@app.post("/gumroad")                       # this function runs when Gumroad calls
def gumroad_ping():
    if request.args.get("key") != os.environ.get("WEBHOOK_SECRET"):
        abort(403)                          # wrong password → go away
    sale = request.form.to_dict()           # buyer info → a Python dictionary
    log_sale(sale)                          # add a row to sales_log.csv
    send(welcome_email(sale["email"], ...)) # send the welcome email
    return "ok", 200                        # tell Gumroad "got it"
```

**Why the `?key=` secret?** Your web address is public. Without the secret, anyone could send fake "sales" and make your server email random people.

**`.env` files:** A text file of `NAME=value` lines for passwords. The code reads them with `os.environ["NAME"]`, so passwords are never written *inside* the code. That's why you can share code or ask ChatGPT about it safely. Just never share `.env`.

**Gmail "App password":** Google won't let scripts use your real password. You make a special 16-letter one in Google Account → Security (2-Step Verification must be on).

**`DRY_RUN=true`:** Emails are printed instead of sent. Leave it on until you've tested.

---

## Chapter 6 — When things break

| You see | It means | Fix |
|---|---|---|
| `'python' is not recognized` | Python isn't on PATH | Reinstall and tick "Add to PATH", or try `py` instead of `python` |
| `ModuleNotFoundError: No module named 'openpyxl'` | Library not installed | `pip install -r requirements.txt` |
| `FileNotFoundError: ...product.json` | Ran from the wrong folder | `cd` into the `budget-kit` folder first |
| A dropdown is missing in Google Sheets | The dropdown read its list from ANOTHER tab, and Google dropped it while converting | Fixed in v2: lists are typed in or live on the same tab. To fix by hand: click the cell → **Data → Data validation** → add a dropdown |
| `URLError: [Errno 11004] getaddrinfo failed` | Your PC couldn't look up the website's address: Wi-Fi dropped, or the internet blipped | Check you're online, then run it again. Nothing is broken |
| PyCharm: "No Python interpreter configured" / "Please select a valid Python interpreter" | PyCharm doesn't know which Python to use for the ▶ button | Click **Create a virtual environment using requirements.txt** in the yellow bar → OK. It makes a private `.venv` folder for this project |
| Apps Script: not sure whether to click **Deploy** | Deploy is for web apps and add-ons | Don't. Pick the function (e.g. `updateInflation`) from the dropdown and click **Run** |
| `ModuleNotFoundError: No module named 'google'` / "Google libraries missing" | The new Google libraries aren't installed in your `.venv` | In PyCharm's Terminal: `pip install -r requirements.txt` |
| `Missing delivery\google_credentials.json` but the file IS there | Windows hid the ending, so your rename made `google_credentials.json.json` | Turn on **View → Show → File name extensions** in File Explorer, then rename it. Or in the Terminal: `Rename-Item delivery\google_credentials.json.json google_credentials.json` |
| `json.decoder.JSONDecodeError` | Typo in a config file | Usually a missing comma or an extra comma before `}`. Paste the file into jsonlint.com |
| A same-tab dropdown (Category, Group, month) is missing in Google Sheets | The file wrote the list as `=$J$5:$J$34`; Google drops it because of the "=" | Fixed in v2.4.1 (Chapter 18). Rebuild + `--push` |
| PermissionError when building | The .xlsx is open in Excel | Close Excel, run it again |
| A brand-new copy shows bills "overdue" | Example bills had due days but nothing logged | v2.6 ships with no example due days; add your own on Setup |
| `###` in a cell | Column too narrow | Drag the column wider |
| `#NAME?` in a cell | A formula uses a function that app doesn't know | Tell Claude which cell. Probably a newer Excel-only function |
| Dashboard "Spent" stays 0 | Expense date isn't in the "Month to view", or category is misspelled | Use the dropdown for category; check Dashboard B4 |
| Webhook returns 403 | `?key=` in Gumroad doesn't match `WEBHOOK_SECRET` | Copy-paste both again, no spaces |
| `SMTPAuthenticationError` | Gmail rejected the login | Use an App password, not your normal password |

---

## Chapter 7 — Glossary

- **Gross pay**: pay before anything is taken out.
- **Take-home / net pay**: what actually lands in your bank.
- **FICA**: Social Security (6.2%) + Medicare (1.45%) = 7.65%.
- **Standard deduction**: income the IRS doesn't tax at all ($16,100 for Single in 2026).
- **Effective tax rate**: total taxes ÷ gross pay. Your "real" overall rate.
- **CPI**: Consumer Price Index, the government's measure of prices. Inflation is how fast it rises.
- **Config file**: a settings file the code reads.
- **API**: a web address that hands out data on request.
- **Webhook**: a web address *you* own that another service calls when something happens.
- **JSON**: a text format for data: `{"name": "value"}`.
- **Library**: someone else's code you install and reuse (openpyxl, Flask).
- **CLAUDE.md**: a rules file Claude reads automatically when it starts in this folder.
- **Remote Control**: lets your phone drive the Claude session that's running on your PC.

---

## Chapter 8 — Working with Claude in PyCharm and from your phone

### 8.1 What CLAUDE.md is
**Like you're 5:** It's the sticky note on the fridge for a babysitter. Every time a new babysitter (a new Claude session) shows up, they read the note first: "back up before you change anything, write down what you did, explain it simply."

Claude Code looks for a file named exactly `CLAUDE.md` in the folder it starts in and reads it before doing anything. So your rules (backups, patch notes, study guide, test checklists) apply in PyCharm and on your phone without you repeating them.

The most important part is the checklist:
```markdown
## After EVERY change (this checklist is required)
1. Back up first: python tools/backup.py "what you're changing"
2. Make the change. Prefer editing config/*.json over changing the scripts.
3. If the change touches the template: run python tools/build_template.py ...
4. Update PATCH_NOTES.md ...
5. Update docs/STUDY_GUIDE.md ...
6. Update tools/SOURCES.md if you used a new website, doc, data source or number.
7. Back up again ...
8. In chat: summary + the code that changed + a test checklist
```
To change a rule, just edit the file. It's plain text.

### 8.2 What SOURCES.md is for
Every number in this product came from somewhere: the IRS, the BLS, Etsy's rules. `tools/SOURCES.md` lists each one with its link and **when to re-check it**. The tax brackets change every fall, inflation changes every month. If a buyer asks "where did 3.7% come from?", the answer is on that page.

### 8.3 The backup.py change (v1.0.1)
PyCharm makes a hidden `.idea` folder of its own settings. It isn't part of the product, so backups now skip it:
```python
SKIP_DIRS = {"backups", "output", "__pycache__", ".git", ".idea"}  # .idea = PyCharm settings
```
A **set** (the `{ }`) is a list of names Python can check very fast. Later in the file, this line throws out any file whose path goes through one of those folders:
```python
if any(p in SKIP_DIRS for p in rel.parts):
    continue          # "continue" = skip this file, go to the next one
```
`rel.parts` splits a path into pieces: `.idea/workspace.xml` → `(".idea", "workspace.xml")`. If any piece is in the skip list, the file is skipped.

### 8.4 Starting Claude in PyCharm
1. Open the BudgetTool folder in PyCharm.
2. Press **Ctrl+Esc** (or type `claude` in the Terminal tab).
3. Ask for a change in plain English. Claude shows each edit in a side-by-side diff. Click **Accept** or **Reject**.

### 8.5 Working from your phone
**Option A, this chat in the Claude app:** works whenever your PC is on with the Claude desktop app open.
**Option B, Remote Control:** in the PyCharm Claude terminal, type `/remote-control`, scan the QR code with your phone, and keep going there.

Both need your PC awake: **Settings → System → Power → Screen and sleep → "When plugged in, put my device to sleep after: Never."**

| Problem | Fix |
|---|---|
| `'claude' is not recognized` | Close and reopen PyCharm (it needs to see the new install), or run the install command again |
| "Cannot launch Claude Code" in PyCharm | **Settings → Tools → Claude Code [Beta] → Claude command**: put the full path from `where claude` |
| Esc doesn't stop Claude | **Settings → Tools → Terminal** → untick "Move focus to the editor with Escape" |
| Phone says the session is offline | Your PC went to sleep or PyCharm closed. Wake the PC; it reconnects by itself |


---

## Chapter 9 — How v2 works (the redesign)

### 9.1 The new layout
**Like you're 5:** v1 was a toolbox with 6 drawers. v2 has 4 drawers with labels anyone can read, plus a locked drawer (Engine) where the grown-up math lives.

| Tab | Job |
|---|---|
| **Setup** | Answer questions: pay, taxes, inflation, budget plan, groups, categories |
| **Spending** | Type each purchase: date, what, category, amount |
| **Dashboard** | See this month: big numbers, groups, donuts, categories, plain-English tax sentences |
| **Reports** | Month by month, category by month, paycheck by paycheck, day by day |
| *Engine (hidden)* | Tax brackets, plan presets, donut chart data |

### 9.2 Groups vs. categories
- **Group** = a big bucket with a % of your take-home (Needs 50%, Wants 20%…). Each group has a **kind**:
  *Bills & needs*, *Flexible spending*, or *Saving & investing*. The kind decides what counts toward your daily allowance and what counts as "saved."
- **Category** = a specific thing you spend on (Rent, Groceries). Every category belongs to one group.
- When you log "Groceries $64", the sheet looks up Groceries → Needs → Bills & needs, using two hidden helper columns on Spending (G and H):
```
G5:  =IFERROR(INDEX(Setup!C57:C86, MATCH(D5, Setup!B57:B86, 0)), "(no group)")
```
**In plain words:** "Find this category's row in the Setup list (MATCH), then grab the group from that row (INDEX)." `IFERROR` catches typos.

### 9.3 The budget plan menu
The plan names and their % live on Engine (rows 29–33). Each group's "% in use" is:
```
=IF(plan="Custom (I'll type my own %)", MyOwn%, INDEX(that group's column of presets, MATCH(plan, plan names, 0)))
```
"If you picked Custom, use what you typed. Otherwise look up the plan's row and take this group's number."

### 9.4 The daily allowance ("You can spend / day")
```
(Flexible plan − Flexible spent this month) ÷ days left in the month
```
Only *Flexible spending* groups count. Rent isn't something you spend a bit of each day.
Example from testing: $650.03 plan − $386.50 spent = $263.53 left ÷ 10 days = **$26.35 a day**.
Days left counts today: on Oct 22, that's Nov 1 − Oct 22 = 10 days.

### 9.5 Listing each group's categories on the Dashboard
Setup has a hidden column J that labels every category like `Needs#1`, `Needs#2`, `Fun#1`:
```
J57:  =C57&"#"&COUNTIFS($C$57:C57, C57, $B$57:B57, "<>")
```
`COUNTIFS` with a growing range (`$C$57:C57`) counts "how many Needs have I seen so far?", giving each one a number.
The Needs card then asks for `"Needs#1"`, `"Needs#2"`… with INDEX/MATCH. That's how categories sort themselves into cards with no special functions.

### 9.6 Progress bars made of text
```
=REPT("█", ROUND(G11*10,0)) & REPT("░", 10 - ROUND(G11*10,0))
```
`REPT("█", 7)` repeats a block 7 times. 71% used → 7 full blocks + 3 empty ones: `███████░░░`. It works in Excel AND Google Sheets. Fancier "data bars" don't survive conversion.

### 9.7 Paydays for every pay schedule
- Weekly / biweekly: first payday + 7 or 14 days × (paycheck # − 1)
- Monthly: `EDATE(first, n−1)` = same day, n−1 months later
- Semi-monthly (like the 15th and 30th): odd paychecks are `EDATE(first, …)`, even ones are `EDATE(first+15, …)`.
  `EDATE` is smart about short months: "Jan 30 + 1 month" gives **Feb 28**, not March 2.

### 9.8 Why dropdowns now survive Google Sheets
When Google converts an Excel file, a dropdown that reads its list from a **different tab** can get dropped. That's what happened to v1's Setup B5.
v2 rule: fixed lists are typed straight into the dropdown (`"Weekly,Biweekly,Semi-monthly,Monthly"`), and changing lists read from the **same tab**. For example, Spending has a hidden column J that copies your category names, so the Category dropdown reads from Spending itself.

### 9.9 Card vs. bold look
`build_template.py` builds every theme in both styles from `"styles": ["card", "bold"]` in `product.json`.
The only difference is colors: the `Look` class picks the primary color (card) or a different color per group (bold, from `"group_colors"`). Same formulas, same layout.
```bash
python tools/build_template.py sage bold    # just the sage theme, bold style
```


---

## Chapter 10 — Money in and real monthly pay (v2.1)

### 10.1 Why "average month" isn't the whole story
**Like you're 5:** If you get a cookie every 2 weeks, most months you get 2 cookies. But months are a bit longer than 4 weeks, so the extra days pile up, and twice a year a month gets **3 cookies**.

Biweekly = 26 paychecks a year. 26 ÷ 12 = 2.17, so most months have 2, and 2 months have 3.
Semi-monthly (1st & 15th) = always 2. Monthly = always 1. Weekly = 4 or 5.

### 10.2 How the sheet counts paydays in a month
The hidden Engine tab lists every payday for 2 years (rows 60+), using the same payday formula as Reports. Then:
```
=COUNTIFS(Engine!B60:B165, ">="&MonthStart, Engine!B60:B165, "<"&NextMonthStart)
```
"Count the paydays on or after the 1st of this month and before the 1st of next month."
**Pay this month** = that count × take-home per paycheck. Test: biweekly from Jan 9, 2026 → May and October get 3 paydays ($3,000.14), and the 12 months add up to exactly $26,001.20.

### 10.3 The Money in log (optional)
On the right side of Spending: **Date received · From · Amount · Notes**.
Log a deposit when it actually hits your bank: a full paycheck, half of one, or side money.
```
Received this month = SUMIFS(Money-in amounts, dates >= MonthStart, dates < NextMonthStart)
```
- If you **never** use it, everything works off expected pay, same as before.
- If you **do**, you get "Cash in hand" (received − spent) and "still coming" (expected − received).

### 10.4 Split paychecks
Example from testing: paycheck #21 was expected to be $1,000.05, but only $500 came in. The Paycheck by paycheck report says:
```
=IF(AND(Received>0, Received<Expected), "Partly paid: "&TEXT(Expected-Received,"$#,##0")&" still coming", ...)
```
→ **Partly paid: $500 still coming**. When the rest lands, log it, and the status flips to ✔ On track.

### 10.5 Why the group budgets still use the average month
Your groups (Needs 50%, Wants 20%…) are based on your **average** month ($2,167 in the example), so your plan stays steady. In a 3-paycheck month, the extra paycheck shows up as more "Left this month." That's a good moment to save it.


---

## Chapter 11 — Pushing to Google Sheets (v2.2)

### 11.1 What OAuth is (the "Allow" screen)
**Like you're 5:** Your script wants to borrow your Google Drive. Instead of giving it your password, Google gives it a **visitor badge** after you say "yes, allow it."
- `google_credentials.json` = the script's **ID card** (which app it is). You download it once from Google Cloud.
- `google_token.json` = the **visitor badge** you approved. It's saved after the first sign-in, so the browser doesn't pop up every time.
- Badges for "Testing" apps expire after 7 days. Then the browser pops up again and you click Allow.
Both files are secrets. Anyone with them could edit your Drive.

### 11.2 How the push works
```python
media = MediaFileUpload(str(xlsx), mimetype=XLSX_MIME)
service.files().update(fileId=sheet_id, body={"name": new_name}, media_body=media).execute()
```
- `files().update(fileId=...)` means "change THIS existing file." It doesn't create a new one, so the link stays the same.
- `media_body` = the new contents (your .xlsx). Google converts it into Sheet format on the way in.
- `body={"name": ...}` renames it to "Paycheck Planner", so buyers' copies are called "Copy of Paycheck Planner."

Safety checks in the script:
1. It stops if the ID isn't a Google Sheet (so you can't overwrite the wrong file).
2. It stops if Google stored the upload as Excel instead of converting it.
3. `--dry-run` prints what it would do and changes nothing.

### 11.3 Getting the Sheet ID
From the link `https://docs.google.com/spreadsheets/d/`**`1xNLyDCm...F9Zg`**`/edit`, the long code between `/d/` and `/edit` is the ID. It's stored in `config/product.json` → `"google"` → `"master_sheet_id"`.

### 11.4 Oops, undo
Google keeps every version: **File → Version history → See version history** → pick one → **Restore this version**.

### 11.5 The frozen-panes fix
Reports used `ws.freeze_panes = "C3"`, which freezes rows 1–2 **and** columns A–B. That's why a line ran through every title. Now it's `"A3"`: freeze rows only.


---

## Chapter 12 — Git and GitHub

### 12.1 Like you're 5
- **Git** = a save-game system for your project. Every "commit" is a save point with a note. You can load any old save.
- **GitHub** = a cloud locker for those saves. It's an off-site backup and a portfolio people can see (if you make it public).
- **Commit** = make a save point. **Push** = upload your save points to GitHub. **Pull** = download them.

### 12.2 .gitignore: the "never save this" list
```gitignore
# your real budget: never upload
reference/
delivery/google_credentials.json
```
Each line is a file or folder Git pretends doesn't exist. A `/` at the end means "this whole folder."
⚠ **Comments must be on their own line.** `reference/   # private` on ONE line does NOT work: Git reads it as a strange filename, and your real budget would get uploaded. We caught exactly this in testing.

### 12.3 How to check nothing private is going up
In PyCharm's Terminal:
```bash
git status --ignored
```
Files under **"Ignored files"** stay private. Make sure `reference/`, `delivery/.env` and the Google key files are listed there BEFORE your first push.

### 12.4 Everyday use in PyCharm
1. **Commit:** press **Ctrl+K**, tick the files, type a message like `v2.2.1: add .gitignore`, click **Commit**.
2. **Push:** press **Ctrl+Shift+K** → **Push**.
3. **See history:** the **Git** tab at the bottom → **Log**.
4. **Undo a file:** right-click the file → **Git → Rollback**.


---

## Chapter 13 — menu.py and shortcuts

### 13.1 Why a menu
**Like you're 5:** instead of remembering the name of every toy box, you get one big box with numbered drawers.
Type `python menu.py`, pick a number, and it runs the right script for you.

### 13.2 How it works
```python
CHOICES = [
    ("Back up the project",           ["tools/backup.py", "{note}"]),
    ("Build + push to Google Sheets (Update)", ["tools/build_template.py", "--push"]),
    ...
]
result = subprocess.run([PY, *args], cwd=ROOT)
```
- `CHOICES` is a list of (label shown to you, command to run) pairs. Adding a menu item = adding one line.
- `subprocess.run` starts another Python script, as if you'd typed it in the Terminal.
- `PY = sys.executable` means "use the same Python that's running the menu", so it always uses your `.venv`.
- `{note}` is a blank the menu fills in by asking you "What are you changing?" before a backup.
- `result.returncode == 0` means the script finished without errors. Anything else means it failed, and the menu tells you to send the red text to Claude.

### 13.3 Why the example row says =TODAY()
`=TODAY()` is a formula that always shows today's date. Buyers might open the sheet months from now, so a fixed example date would hide the example from their Dashboard. Typing over it with a real date (or deleting the row) is fine.

### 13.4 Every shortcut and link
They're all in `docs/SHORTCUTS.md`. The most useful: **↑** in the Terminal repeats your last command, and **Win + Shift + S** takes a screenshot to paste to Claude.


## Chapter 14 — The Goals tab: am I on track? (v2.3)

**What it answers:** "Once my bills are paid, how much did I actually save this month, this year,
and am I ahead or behind my goal, by how much?"

### The 3 parts
1. **Cards** (top): saved this plan year, goal so far, ahead/behind, on pace for year end.
2. **Your savings goals**: one row per goal (Emergency fund, Car, Trip…).
3. **Savings scorecard**: one row per month of your plan year.

### How "saved" is counted
Saving = a row on the Spending tab whose category is in a **Saving & investing** group
(e.g. you move $300 to savings → log it as "Emergency fund", $300).
If you turn on **"Count money left over as saved"**, unspent money also counts, but only after the month is over
(this month isn't finished, so its leftover could still get spent).

```python
as_calc(ws, f"I{r}", f'=G{r}+IF(AND($G$8="Yes",{nxt}<=TODAY()),MAX(0,H{r}),0)', MONEY, bold=True)
```
ELI5: "Counted as saved" = what you put into savings (G) **plus**, if the switch says Yes **and** the month
has ended (`next month's start <= today`), whatever was left over (H). `MAX(0, …)` stops a negative leftover
(overspending) from *subtracting* from your savings.

### "Should have by today" (the on-track line)
```python
=IF(TODAY()<=Start, Already, IF(TODAY()>=Target date, Target,
   Already + (Target - Already) * (TODAY() - Start) / (Target date - Start)))
```
ELI5: draw a straight line from what you had on the start date to your target on the target date.
Halfway through the time = you should be halfway there. **Ahead/behind = saved so far − that line.**
`+$200` means ahead, `−$150` means behind.

### "Needed each month now"
`(Target − Saved so far) ÷ months left until the target date` (at least 1 month, so we never divide by 0).

### Ahead or behind for the whole year
```python
=SUMIFS(counted, months, "<="&TODAY())    ' saved in months that have started
=SUMIFS(goal,    months, "<="&TODAY())    ' goal for those same months
```
`SUMIFS(add these, look here, "<="&TODAY())` adds up only the months up to today, so future months
don't make you look behind. The difference is the big **Ahead or behind** card.

### New number format
`'+$#,##0;-$#,##0;$0'` = three formats split by `;`: positive (with a +), negative (with a −), and zero.

### Split paychecks (payment 1 + payment 2)
The scorecard's **Money in** uses what you logged under *Money in* on the Spending tab when there is any
(so $800 + $900 = $1,700), otherwise the paychecks you expect.

### Errors and gotchas
- **A dropdown pointing to another tab can vanish in Google Sheets**, so the "Counts money logged to"
  list lives in a hidden column (S) on the Goals tab itself, copied from Setup.
- Category → kind needed a new hidden helper on Setup (column K), because a category only knows its *group*,
  and the group knows its *kind* (bills / flexible / saving).


## Chapter 15 — Updating buyers' copies without touching their data (v2.3.1)

**The problem:** when someone taps *Make a copy*, Google gives them a brand-new, separate file.
You can't see it (good) and you can't push changes into it (annoying).

**The fix:** an "Import my data from an old copy" menu item. The buyer makes a NEW copy of the new version,
then the script copies only the yellow boxes (what they typed) from their OLD copy.

```javascript
var INPUT_RANGES = [ ['Setup', 'C8:C11'], ['Spending', 'B5:F1004'], ['Goals', 'B12:G19'] /* ... */ ];
...
var values = src.getValues();      // what you SEE in each cell
var formulas = src.getFormulas();  // the formula behind it, or "" if it was typed
var merged = values.map(function (row, r) {
  return row.map(function (v, c) { return formulas[r][c] ? formulas[r][c] : v; });
});
to.getRange(pair[1]).setValues(merged);
```
ELI5: for every yellow box, if it had a formula (like `=TODAY()`) copy the formula; otherwise copy what was typed.
`setValues` with a text that starts with `=` becomes a formula again.

- `SpreadsheetApp.openByUrl(url)` opens the old copy. Google asks the buyer for permission the first time,
  because the script reads a different file.
- If the old copy has no Goals tab (older version), that part is skipped, and nothing breaks.
- **Rule for the future:** never move a yellow box to a new address without updating `INPUT_RANGES`,
  or imports would put data in the wrong place.

### Error explained: `PermissionError: [Errno 13] Permission denied: ...xlsx`
Windows **locks** a file while Excel has it open, so Python can't save over it.
Fix: close the file in Excel and run again. The build now catches this and prints a friendly message:
```python
except PermissionError as e:
    print(f"Can't save {Path(e.filename).name}: it's open in Excel (or another app).")
```

### Why a web app solves this for good
In the web app, the *app code* lives on your website and each person's *data* lives on their own phone.
When you publish an update, everyone gets the new code next time they open it, and their data stays put.


## Chapter 16 — "Start tracking savings from" (v2.3.2)

**Problem:** your plan year starts with your first payday (say January). If you start using the planner in
October, January to September look like 9 months of missed savings.

**Fix:** a hidden helper cell U4 = the month you started (cell G9), or the plan year's start if G9 is blank:
```python
put(ws, "U4", '=IF(N($G$9)=0,U2,MAX(U2,DATE(YEAR($G$9),MONTH($G$9),1)))')
since = f'{months},">="&$U$4'      # added to every SUMIFS: "only months since you started"
```
- `N($G$9)=0` means "G9 is empty" (N turns blank into 0).
- `DATE(YEAR(..),MONTH(..),1)` snaps any date to the 1st of its month.
- `MAX(U2, ...)` stops a tracking date earlier than the plan year from breaking things.
Every month before U4 shows **"Before you started"** and is left out of the cards and running totals.

---
## Chapter 17 — The Accounts tab: high-yield savings + investing (v2.4)

**ELI5:** a high-yield savings account (HYSA) is a piggy bank that pays you rent for keeping money in it.
A regular bank pays almost nothing (about 0.38% a year). A HYSA pays about 4%. Investing (index funds) can
pay more over many years, but it can also drop. This tab shows what your money does in each one.

### What each column does
| Column | Formula idea | Plain English |
|---|---|---|
| J Now | `Balance × (1+rate)^(years since "as of") + deposits logged since` | Your balance, grown a bit since you checked it |
| K Earns/month | `Now × ((1+rate)^(1/12) − 1)` | Turns a yearly rate into a monthly one |
| L–O | `FV(monthly rate, months, −deposit, −now)` | Future Value: what it grows to if you keep adding |

```
=FV(((1+G12)^(1/12)-1), 120, -H12, -J12)
```
- `FV` = Excel's "future value" function. It asks: rate per month, how many months, how much you add each
  month, how much you start with. The minus signs are Excel's way of saying "money going IN".
- 120 months = 10 years.
- **Compound interest:** you earn interest on your interest. That's why the 30-year number is so much
  bigger than what you put in.

### Today's dollars switch
`(1+rate)/(1+inflation) − 1` = the "real" rate. $100,000 in 30 years won't buy what $100,000 buys today,
so "Yes" shows what the future money is worth in today's prices.

### High-yield bonus card
```
=SUMPRODUCT((C12:C19="High-yield savings")*J12:J19*(G12:G19-$U$5))
```
For each HYSA row: balance × (your rate − national average). `SUMPRODUCT` adds it all up. The `(C="...")`
part is 1 for HYSA rows and 0 for others, so other accounts don't count.

### Hidden helpers
- Columns W:AF: worth after 1, 2, 3, 4, 5, 10, 15, 20, 25, 30 years (feed the growth table).
- U5 = national average savings rate (from `config/product.json`), U6 = the inflation rate Setup uses.

### Not financial advice
The sheet shows estimates. Rates change; investing can lose money. Many people plan with 6–7% for long-term
investing to be careful.

---
## Chapter 18 — Why the dropdowns vanished in Google Sheets (v2.4.1)

**ELI5:** Excel and Google Sheets read the same file, but Google is pickier. We wrote the dropdown list as
`=$J$5:$J$34`. Excel ignores the extra "=". Google sees something it doesn't understand and silently
deletes the dropdown. No error, it's just gone.

**How we found it:** we asked Google which cells have dropdowns. The typed lists ("Weekly, Biweekly...")
were there; every list that points at cells was missing. Then we uploaded a tiny test file with both
versions. Only the one without "=" survived.

```python
# before
src = '"' + ",".join(choices) + '"' if choices else "=" + same_sheet_range
# after: no "=" in front of a range
src = '"' + ",".join(choices) + '"' if choices else same_sheet_range.lstrip("=")
```
`.lstrip("=")` = "remove any = from the start of this text".

### Date pickers
```python
dv = DataValidation(type="date", operator="greaterThan", formula1="36526", allow_blank=True)
```
- A "date" rule means only dates are allowed. Google Sheets adds a calendar when you double-click.
- 36526 is how spreadsheets store Jan 1, 2000 (days since 1900). Any real date after that is OK.

### Every status gets a color
Conditional formatting rules are checked top to bottom. Red (over), amber (almost), green (on track),
and now a last catch-all: `LEN(cell)>0` → blue. "LEN > 0" means "the cell has any text", so anything the
first three rules didn't catch still gets a color.

**Lesson:** after `--push`, open the Google copy and click a dropdown. If one is missing, the file did
something Google doesn't like.

---
## Chapter 19 — Overspending alerts (v2.5)

**ELI5:** the sheet is a smoke detector. Hidden columns check every way you could be overspending.
The banner and the Alerts list are the beeper. The Google script can also beep in a pop-up or an email.

### Step 1: one row per possible alert (hidden columns X, Y, Z on the Dashboard)
```
X12 = IF(AND(B13<>"", Setup!$C$47<>"Saving & investing", N(D13)>0, N(E13)>N(D13)),
         "Over budget: "&B13&" is "&TEXT(E13-D13,"$#,##0")&" over. ...", "")
Y12 = 1                       ← 1 = red, 2 = yellow
Z12 = IF(X12="","",SUMPRODUCT(--(LEN($X$2:X12)>0)))
```
- X is the message, or blank if it isn't happening.
- Z numbers only the real alerts: 1, 2, 3... `LEN(...)>0` is "has text"; `--` turns TRUE/FALSE into 1/0;
  SUMPRODUCT adds them up from the top down to this row.
- There are 53 possible alerts: 1 overall, 6 groups over, 30 categories over, 6 groups almost out,
  1 spending pace, 1 savings squeeze, 8 goals behind. Red ones come first, so they're listed first.

### Step 2: show alert #1, #2, #3... with no gaps
```
=IFERROR(INDEX($X$2:$X$54, MATCH(3, $Z$2:$Z$54, 0)), "")
```
"Find the row numbered 3 in Z, and give me its message from X." No FILTER needed (FILTER breaks in
older Excel; see the coding rules).

### Spending pace, in plain words
`used = flexible spent ÷ flexible plan`, `gone = days passed ÷ days in month`.
If `used > gone + 10%`, you're spending faster than the calendar. Halfway through the month with 70%
spent → heads up.

### The Google script beeper
```javascript
function onEdit(e) {               // Google runs this by itself every time you type
  if (sheet is not Spending) return;
  SpreadsheetApp.flush();          // let the formulas finish first
  var a = getAlerts_();            // reads Dashboard X:Y
  if (a.red.length) e.source.toast(a.red[0], '⚠ Over budget', 10);
}
```
- `onEdit` is a "simple trigger": no setup, but it can't send email.
- Emails need permission, so they run from a timer the buyer turns on in the menu
  (`ScriptApp.newTrigger('checkAlertsAndEmail').timeBased().everyDays(1).atHour(8)`).
- The email goes only to the sheet owner (`Session.getEffectiveUser()`), never anyone else.
- The Google Sheets phone/iPad app doesn't show pop-ups. There, the red banner and the emails do the job.

### Grace period
A savings goal doesn't nag in its first month. Otherwise someone who just started would see
"behind" on day one.

---
## Chapter 20 — A smaller first paycheck (v2.6)
**ELI5:** your first check is often small because you only worked part of the pay period.
```
Engine!B20 = IF(first paycheck typed, per_check × first ÷ normal pay, per_check)
expected pay for a month = paydays × per_check − (per_check − Engine!B20) if your first payday is in that month
```
Taxes take the same share of a small check as a big one (roughly), so we shrink the take-home by the
same fraction. `expected_pay()` in build_template.py writes that formula everywhere pay is expected.

---
## Chapter 21 — Bills: due dates and "which paycheck pays it" (v2.6)
- Due date this month: `DATE(year, month, MIN(due day, last day of the month))`. A bill due on the 31st is
  due on the 30th in a 30-day month.
- **Which paycheck pays it** = the last payday on or before the due date:
```
INDEX(paydays, COUNTIF(paydays, "<=" & due date))
```
  Paydays are in order, so "how many paydays are on or before the due date" is the position of the right one.
- **Sorting by due date without SORT():** each bill gets a number `due date + row/1000` (so two bills on
  the same day don't tie), then `COUNTIF(numbers, "<" & mine) + 1` = its place in line. The visible list
  shows place 1, 2, 3... with INDEX/MATCH (same trick as the alerts).
- **Joining names without TEXTJOIN** (TEXTJOIN isn't in older Excel): `IF(match, name & ", ", "") & ...`
  for every bill, then cut off the last ", " with `LEFT(text, LEN(text) - 2)`.

---
## Chapter 22 — The debt payoff planner (v2.6)
**ELI5:** every debt gets its minimum. Your extra money is a snowball you roll onto ONE debt. When that debt
is gone, its minimum joins the snowball, so the next debt goes faster.

The hidden grid (columns AN onward) has one row per month for 30 years. For each debt, in payoff order:
```
owed     = last month's balance × (1 + APR ÷ 12)        ← interest gets added
minimum  = MIN(owed, minimum payment)                  ← never pay more than what's owed
pool     = (all minimums + extra) − minimums paid this month
extra    = MAX(0, MIN(owed − minimum, pool − extra already used by debts ahead of it))
balance  = owed − minimum − extra
```
- Months to pay off = how many months still had a balance + 1. If it's still not paid after 360 months,
  it says **never** (the minimum doesn't even cover the interest).
- Interest = rate × the sum of every month's starting balance.
- A second grid does "minimums only", so the sheet can show what your plan saves.
- Snowball vs. Avalanche only changes the ORDER: smallest balance first, or highest APR first.
- The app (`Calc.debts`) runs the same loop in JavaScript and matches the sheet to the cent.

---
## Chapter 23 — Net worth and the "photo" script (v2.6)
Net worth = what you own − what you owe. A formula always shows TODAY, so it can't remember last month.
The script `saveNetWorth()` finds the cell next to "NET WORTH", then writes that number into this month's
row of the log. ELI5: taking a photo of your piggy bank once a month.

---
## Chapter 24 — Making it look the same everywhere (v2.6)
- `big_cards()`: one function draws the 4 cards on every tab, so they always match.
- `total_row()`: every Total row gets the same tinted band and bold text.
- `red_if_negative()`: a conditional format, "if this number is below 0, make it red".
- `status_colors()` now uses `words(cell, "over", "never", ...)` to build the OR(...) formula, and
  `stopIfTrue` so only the first matching color applies.

---
## Chapter 25 — Cleaning up old versions automatically
```python
NAME = re.compile(r"_v(\d+)\.(\d+)\.(\d+)_.+\.xlsx$")      # finds "2.6.0" in the file name
versions = sorted(by_version, reverse=True)                   # (2,6,0) comes before (2,5,0)
keep = versions[:KEEP]                                        # the newest 3
```
- **Why numbers, not text?** As text, "2.10.0" sorts before "2.9.0" (because "1" < "9").
  Turning each part into a number fixes that: (2,10,0) > (2,9,0).
- `f.unlink()` deletes one file. If Excel has it open, Windows says PermissionError, and we just tell you.
- Only files that look like `Paycheck_Planner_vX.Y.Z_....xlsx` are ever touched.
- Old versions can always be rebuilt from a zip in `backups/`.

---
## Chapter 26 — Making it look professional: lining things up
- **Same Y axis:** things side by side must live on the SAME rows. The Dashboard alerts used to be one row
  per alert starting at row 5, so they drifted away from the cards. Now: banner = row 3, ALERTS card =
  rows 4–6 (like the other cards), alert list = rows 11–16 (like the 6 groups).
- **Vertical centering:** Excel puts text at the bottom of a cell unless you say otherwise. `put()` now
  centers it by default:
```python
elif value is not None and not c.alignment.vertical:
    c.alignment = Alignment(vertical="center")
```
- **No overlap:** a small checker (not shipped) estimates each text's width (letters × font size) and
  flags any that spill into a cell that already has something in it. Fixes: wider column, shorter text,
  or wrap.
- **Empty template:** `plan_start()` = "the month of your first payday, or this January if it's blank",
  so a blank copy never shows the year 1900.

---
## Chapter 27 — Safe to spend, "Can I afford it?", paid early, payment plans (v2.7)
**ELI5:** you just got paid. Some of that money already has a job (bills due before your next paycheck).
Whatever is left after what you've already spent is safe to spend.
```
this paycheck  = the last payday on or before today          INDEX(paydays, COUNTIF(paydays, "<=" & TODAY()))
next paycheck  = the one after it                             INDEX(paydays, that count + 1)
safe to spend  = take-home − spent since this paycheck − (bills + plan payments due before the next one)
can I afford X = X <= safe to spend
```
- A bill counts if its next due date falls before the next paycheck AND you haven't logged enough to it
  since this paycheck landed.
- **Paid early:** every payday in the hidden list is `payday − days early`. One change, and everything
  (months, bills, reports) uses the day the money lands.
- **Payment plans:** "how many payments should be done by date X" = for every 2 weeks:
  `INT((X − first) / 14) + 1`; monthly: count the months. What you should have paid = payments × amount
  (never more than the total). Should-have − paid so far > 0 means **Late**.
- **A-to-Z dropdowns without SORT():** each category gets its place in line with
  `COUNTIF(names, "<" & name) + 1` (text compares alphabetically), then the list shows place 1, 2, 3...
- **Import by headings:** `findRow_()` in the script looks for "Group name", "Category", "Who / what"...
  in column B of the old AND the new copy, so rows can move between versions.

---
## Chapter 28 — Budget check, OK it or cover it, sort by date (v2.8)
**ELI5:** every purchase asks "how much is left in my bucket after this?" If the bucket is empty, you can
say why it's OK, or pour more money in (extra money or last month's leftover).
```
which bucket  = the category if it has a monthly goal, otherwise its group          (hidden column L)
bucket size   = plan + extra Money in "For" it this month + last month's leftover (if a row picked it)   (X)
spent so far  = everything in that bucket this month up to this purchase's date                       (Y)
status        = over? → OK'd? "Noted" : "⚠ Over by $X";   90%+ used → "Close";   else "✔ $X left"
```
- "Last month's leftover" = MAX(0, last month's plan − last month's spending) for that bucket.
- The Dashboard alert checks hidden keys like "Fun|202610" (made only when a row is OK'd): if one exists,
  the alert becomes yellow "Noted" instead of red.
- **Sorting with formulas around:** the script sorts only the columns you type in plus column F. Every
  formula only looks at its own row, so after a sort each row's formulas just recalculate.

---
## Chapter 29 — Phone view tab + "blank" wording (v2.8)
- A phone shows about 2 normal columns at once. The Today tab uses just B (25 wide) and C (15 wide), one
  thing per row, so nothing needs zooming. It only points at other tabs (`=Dashboard!$B$5`), so it's
  always in sync.
- Colors on Today come from a hidden column D holding the status text from the other tab, because
  conditional formatting in Excel can't look at a different tab directly.
- Merged cells: color rules must cover the WHOLE merged box (K11:M16), not just K11:K16, or the color can
  look shifted next to the header.
- Blank template: every status first checks "is there any pay yet?" (`N(Setup!C10)=0`) before saying
  "On track". Otherwise $0 of $0 looks like success.

---
## Chapter 30 — Sending your sheet data to the app (v2.8)
**ELI5:** packing your notebook into a box the app knows how to unpack.
- `buildAppData_()` in the script reads every place you type (found by their labels and headers, like the
  importer), and builds one object shaped exactly like the app's saved data (`planner.v1`).
- Dates become text like "2026-10-22"; "Last month's leftover" becomes `cover: "leftover"`; "2 days early"
  becomes `early: 2`.
- The dialog turns it into a file with `new Blob([...])` and a download link. Nothing leaves your device.
- **How we tested it without Google:** a fake spreadsheet (`getRange(...).getValues()`) filled with a real
  test copy's values ran the SAME script code, then the app loaded the file and showed the same numbers
  ($2,413 pay, $156.34 safe to spend, net worth $4,338.71).
- **Why the sheet can't change its layout by device:** a spreadsheet has one fixed grid for everyone. The
  app is a web page, and CSS (`@media` rules) lets it rearrange itself for a phone, tablet or computer.

---
## Chapter 31 — Sorting the categories table
`range.sort([{column: 3, ascending: true}, {column: 2, ascending: true}])` = sort by Group (column C), and
inside each group by name (column B). It's safe because every other tab finds a category by its NAME
(MATCH / SUMIFS), never by which row it's on. The groups table is NOT sortable: the Dashboard reads groups
by row (group 1 = Dashboard row 11).


## Chapter 32 — One calendar for every month list (v2.9)

**What changed:** every list of months (Dashboard picker, Reports, Goals scorecard, net worth log) now runs
Jan 2026 to Dec 2028. The start and length live in `config/product.json`:

```json
"calendar_start": "2026-01",
"calendar_months": 36,
```

The builder turns that into one formula and reuses it everywhere:

```python
def cal_start():
    y, m = CONFIG["calendar_start"].split("-")
    return f"DATE({int(y)},{int(m)},1)"          # -> DATE(2026,1,1)

for i in range(CAL_MONTHS):                     # 36 rows
    as_calc(ws, f"B{r}", f"=EDATE({cal_start()},{i})", "mmmm yyyy")
```

**ELI5:** `EDATE(start, 5)` means "the start date, 5 months later". Row 1 is +0 months (Jan 2026),
row 2 is +1 (Feb 2026)... row 36 is +35 (Dec 2028). Want 2029 too? Change 36 to 48 and rebuild.

**Year totals** use SUMIFS with two date limits, "on or after Jan 1" and "before next Jan 1":

```python
in_year = f'{months},">="&{year_start(k)},{months},"<"&{year_start(k + 1)}'
f"=SUMIFS(C$24:C$59,{in_year})"
```

**Goals cards start over each January.** Hidden helpers on Goals: U2 = Jan 1 of this year, U5 = first month
that counts at all (the latest of Jan 2026, your first payday's month, and "Start tracking from"),
U4 = the later of U2 and U5. The cards only add months between U4 and next January.

**Importing by month, not by position.** Old copies' net worth log started on your first payday; the new one
starts in January. Copying "row 1 to row 1" would put August's number in January. So the importer reads each
old row's month and looks for the same month in the new list:

```js
if (d instanceof Date && d.getFullYear() === o[0].getFullYear() && d.getMonth() === o[0].getMonth()) {
  to.getRange(hN + 1 + k, 3).setValue(o[1]);    // same month -> copy the number
}
```

**When things break (this chapter):**
- *Reports showed 0 paydays from Aug 2028 on.* The Engine payday list only held 2 years of paychecks.
  Fix: it now holds 4 years (`PAYDAY_YEARS`), enough for a first payday late in 2026 plus 3 more years.
- *Reading past the bottom of a tab.* `getRange` throws an error if you ask for rows that don't exist. The
  importer and `table_` now cap the number of rows with `getMaxRows()`.
- *"Money in" looks too low for a month.* A Money in row only counts in a month if its **Date received** is a
  real date in that month (it turns right-aligned when Google sees a date). Text like "Aug 27th" or a date in
  the wrong year doesn't count.


## Chapter 33 — The "How to use" tab (v2.10)

The guide is just a Python list of sections, and each section is a list of (label, text) pairs:

```python
GUIDE = [
    ("SET IT UP ONCE (about 5 minutes, on the Setup tab)", [
        ("Step 1", "How often do you get paid? ..."),
        ...
    ]),
]
```

`build_guide()` walks through that list and writes one row per pair: the label in column B (bold, colored),
the text in column C (wrapped). It guesses the row height from the text length:

```python
lines = max(1, -(-len(text) // 105))     # about 105 letters fit on one line
ws.row_dimensions[r].height = 6 + 16 * lines
```

**ELI5:** `-(-a // b)` is a trick for "divide and round UP". 210 letters ÷ 105 = 2 lines; 211 letters = 3 lines.
Each line is about 16 points tall, plus 6 for breathing room.

**Why a list?** To change the guide, edit the words in `GUIDE` and rebuild. No cell addresses to fix.

**Expected vs Received (the question that started this):** Expected comes from Setup (your plan). Received comes
only from Money in rows. If Received shows "-", no Money in row has a real date inside that paycheck's dates.


## Chapter 34 — Paydays on set days (15th & 30th) and the weekend rule (v2.11)

Some jobs pay on **set days** (the 15th and the 30th) instead of "every 15 days". If that day is a Saturday or
Sunday, most pay the **Friday before**.

**Step 1: find the official payday.** Think of each month as having two "slots": slot 0 = the 15th, slot 1 = the 30th.
Paycheck #1 is whichever slot your first payday was in; every paycheck after it is the next slot:

```python
s_ = f"({slot0}+{n}-1)"            # slot number, counting from your first payday's month
mo = f"INT({s_}/2)"                 # every 2 slots = 1 month later
day = f"IF(MOD({s_},2)=0,{d1},{d2})"  # even slot = 15th, odd slot = 30th
f"MIN(DATE(YEAR(base),MONTH(base)+{mo},{day}),EOMONTH(base,{mo}))"
```

**ELI5:** paychecks go 15th, 30th, 15th, 30th... like left foot, right foot. `INT(slot/2)` counts how many
months you've walked, `MOD(slot,2)` says which foot. `MIN(..., EOMONTH(...))` stops "Feb 30" from becoming
March 2: it can never go past the last day of that month.

**Step 2: the weekend rule.** `WEEKDAY(date,2)` gives 1 for Monday ... 6 for Saturday, 7 for Sunday:

```python
f'{cell}+IF(OR(freq="Weekly",freq="Biweekly",{wd}<6),0,'
f'IF(rule="Friday before",5-{wd},IF(rule="Monday after",8-{wd},0)))'
```

Saturday: 5 - 6 = -1 day (Friday). Sunday: 5 - 7 = -2 days (Friday). Monday after: 8 - 6 = +2, 8 - 7 = +1.

**Step 3: bank early days** come off last: Money lands = official payday, moved off the weekend, minus 0-5 days.

**Why two columns in the Engine?** The first-payday formula is long. Writing it once in column C and pointing
column B (and Reports) at it keeps the file small and fast instead of copying a 5,000-letter formula everywhere.

**Tested:** 15th & 30th from Aug 15, 2026 → Fri Aug 14, Fri Aug 28 (30th was Sunday), Tue Sep 15, Wed Sep 30,
... Fri Nov 13, ... Fri Feb 26, 2027 (Feb "30th" = Feb 28, a Sunday).

**When things break:** a dropdown answer that's wider than its box spills into the tip next to it
("Paid the Friday before"). Fix: shorter choices ("Friday before").


## Chapter 35 — Bank holidays (v2.12)

**Working out a holiday with a formula.** "3rd Monday of January" (MLK Day):

```python
f"DATE(y,1,1)+MOD(1-WEEKDAY(DATE(y,1,1),2),7)+14"
```

**ELI5:** start on Jan 1. `MOD(1 - weekday, 7)` is "how many days until the first Monday". Add 14 more days
(2 weeks) to get the 3rd Monday. "Last Monday of May" (Memorial Day) works backwards from May 31.
If a holiday lands on Sunday, banks close Monday: `+IF(WEEKDAY(d,2)=7,1,0)`.

**Moving a payday off weekends AND holidays with one function:**

```python
WORKDAY(payday+1, -1, holidays)    # this day if it's a business day, otherwise the one before
WORKDAY(payday-1,  1, holidays)    # this day if it's a business day, otherwise the one after
```

`WORKDAY(date, -1, list)` jumps back 1 business day, skipping Saturdays, Sundays and every date in `list`.
Starting from the day AFTER payday and going back 1 lands on payday itself when payday is a normal workday.

**Every-2-weeks pay** only moves when payday is in the holiday list (`COUNTIF(holidays, payday)>0`), because
it's supposed to land on the same weekday every time.

**When things break:** renaming a dropdown choice ("Friday before" → "Workday before") breaks older copies that
still have the old word in the cell. Fix: the formula accepts both: `OR(C15="Workday before", C15="Friday before")`.


## Chapter 36 — Sort Money in by date (v2.12.1)

One sorting function now serves two menu items. Each table is described once:

```js
var SORT_TABLES = {
  spend: ['Date', 'Notes (optional)', 1000, 'Spending is in date order.'],
  'in': ['Date received', 'Notes (optional)', 300, 'Money in is in date order.']
};
function sortMoneyIn() { sortTables_(['in'], false); }
```

`sortTables_` finds the table's first header ("Date received") and its last one ("Notes (optional)") by name,
then sorts every column in between by the first column:

```js
sh.getRange(h.row + 1, first, t[2], width).sort({ column: first, ascending: true });
```

**ELI5:** grab the whole Money in block (date, from, amount, for, notes) like one stack of cards, and sort the
stack by the date on each card. Because the whole row moves together, an amount never gets separated from its date.
Blank rows go to the bottom.

**Why find headers by name?** If a column is added later, the script still finds the right block.


## Chapter 37 — When Import my data hits "Not on the list" (v2.12.2)

**The error:** `Exception: Not on the list - Please pick from the list.`
Dropdowns in this template REJECT anything that isn't on their list. Apps Script obeys that rule too, so when the
importer tried to write an old value that the new copy's list doesn't have, Google threw an error and the whole
import stopped halfway.

**The fix:** try normally; if Google says no, lift the dropdown rules, write, and put them back.

```js
try {
  toRange.setValues(out);
} catch (e) {
  var rules = toRange.getDataValidations();   // 1. remember the dropdowns
  toRange.clearDataValidations();             // 2. take them off for a moment
  toRange.setValues(out);                     // 3. write your old values
  toRange.setDataValidations(rules);          // 4. put the dropdowns back
  findOffList_(toRange, out, rules);          // 5. note which cells don't match their list
}
```

**ELI5:** a bouncer (the dropdown) won't let a guest in whose name isn't on the list. We walk the guest in through
the side door, put the bouncer back, and hand you a note: "these guests aren't on the list, check them".
Google marks those cells with a small red corner until you pick a value from the list.

`try { ... } catch (e) { ... }` means: try this; if it throws an error, run the catch part instead of crashing.


### Chapter 37b — Why the try/catch didn't work (v2.12.3)

Apps Script saves your changes in batches. `setValues()` only puts the values in a waiting line; Google writes the
waiting line later (when the script reads something, or ends). The dropdown said "no" at THAT moment, after the
script had already left the `try { }` block, so `catch` never ran.

```js
if (hasRules) toRange.clearDataValidations();   // take the dropdowns off FIRST (no waiting to fail)
toRange.setValues(out);
SpreadsheetApp.flush();                          // "save now", while nothing can refuse the values
if (hasRules) toRange.setDataValidations(rules); // dropdowns back on; off-list cells get a red corner
```

**ELI5:** instead of sneaking the guest past the bouncer after he says no, we send the bouncer on a coffee break,
seat everyone, lock the doors (flush), and then bring the bouncer back.

**Lesson:** a `try/catch` only catches errors that happen INSIDE it. If something is delayed, call `flush()` inside
the try, or avoid the error in the first place.


## Chapter 38 — Showing cents only when they matter (v2.12.4)

```python
money = lambda x: f'TEXT({x},IF(ABS({x})<1,"$0.00","$#,##0"))'
```

`TEXT(number, format)` turns a number into words. The format can itself be a formula: under $1 we use "$0.00"
(shows cents), otherwise "$#,##0" (whole dollars). `ABS()` drops the minus sign so -0.26 counts as "under $1" too.

**When things break:** "$#,##0" rounds to whole dollars, so 0.26 became "$0", and "Over by $0" looked like a bug.
