# Making the Google Sheets version (one time, ~10 minutes)

You build this once. Every buyer then gets their own copy from a link.

## Step 1 — Upload the Excel file
1. Go to drive.google.com and click **New → File upload**.
2. Pick `output/Paycheck_Planner_v2.0.0_sage.xlsx` (or another theme/style).
3. Right-click the uploaded file → **Open with → Google Sheets**.
4. Click **File → Save as Google Sheets**. You now have a real Google Sheet.

## Step 2 — Add the inflation script
1. In the Google Sheet, click **Extensions → Apps Script**.
2. Delete everything in the editor (the empty `myFunction`).
3. Paste in everything from `google-sheets/inflation_autoupdate.gs`.
4. Click the 💾 save icon. Name the project `Budget Tools`.
5. Pick `updateInflation` in the function dropdown (it starts on `onOpen`) and click **Run**.
   **Don't click Deploy.** Deploy is for web apps and add-ons. This script is attached to the sheet and travels with every copy.
6. Google asks for permission ("This app isn't verified"). Click **Advanced → Go to Budget Tools**, then **Allow**. This is normal for your own scripts.
7. Go back to the sheet. On the Setup tab, "Latest US rate" (cell C23) now shows the newest number.

## Step 3 — Test it like a buyer
1. Reload the sheet. A **Budget Tools** menu appears at the top.
2. Click **Budget Tools → Update inflation now**. A pop-up should say "Inflation set to X%".
3. Check the dropdowns still work: Setup C9 (pay frequency), Setup C43 (budget plan), Setup C57 (category group), Spending D5 (category), Dashboard C2 (month).

## Step 4 — Make the "copy" link buyers get
1. Click **Share → General access → Anyone with the link → Viewer**.
2. Copy the link. It looks like `https://docs.google.com/spreadsheets/d/LONG_ID/edit?usp=sharing`.
3. Swap the end: replace `/edit?usp=sharing` with `/copy`.
4. Paste that `/copy` link into `config/product.json` → `"sheets_copy_link"`.

When a buyer opens a `/copy` link, Google makes them their own private copy, script included. They'll see the same one-time permission screen when they first use the menu. The buyer PDF should tell them it's expected.

## Step 5 — Update the script when it changes (v2.3+: "Import my data")
Do this once in your MASTER sheet whenever `inflation_autoupdate.gs` changes:
1. Open the master Google Sheet → **Extensions → Apps Script**.
2. Select everything in the editor (Ctrl+A) and delete it.
3. Paste in everything from `google-sheets/inflation_autoupdate.gs`. Click 💾 save.
4. Reload the sheet. **Budget Tools** now also shows **Import my data from an old copy**.
New `/copy` copies get the new script automatically. Pushing a new build (`menu.py` → 4) keeps the script,
because the push updates the same file instead of making a new one.

## How buyers move to a new version (send them this)
1. Open the new `/copy` link → **Make a copy**.
2. In the new copy: **Budget Tools → Import my data from an old copy**.
3. Allow the one-time permission screen (it needs to read your old copy).
4. Paste the link to your old copy → **OK** → **Yes**.
5. Done: all your settings, spending, deposits and goals are in the new version. Your old copy isn't changed.

## Never do this
- Don't give buyers the `/edit` link. They'd all be typing into YOUR sheet.
- Don't put your real budget in the master sheet. Keep it blank except for the example row.
