# Auto-update your Google Sheet from PyCharm (one-time setup, ~15 minutes)

After this, one command builds the template AND updates your master Google Sheet:
```bash
python tools/build_template.py --push
```
Same Sheet, same link, new content. Buyers' existing copies are never touched.

---

## Part 1: Make a Google Cloud project (free)

**Like you're 5:** Google won't let just any program change your files. A "project" is a name tag for your script, so Google knows who's asking and can ask YOU for permission.

1. Go to **console.cloud.google.com** and sign in with the **same Google account** that owns your Paycheck Planner sheet.
2. If it asks you to agree to terms, tick the box and click **Agree and continue**. There's no credit card and no charge for this.
3. At the top left, click the project picker (it may say "Select a project") → **New project**.
4. Name: `BudgetTool`. Click **Create**. Wait a few seconds, then make sure **BudgetTool** is selected in the project picker.

## Part 2: Turn on the Google Drive API

5. In the search bar at the top, type **Google Drive API** and click it.
6. Click **Enable**.

## Part 3: The permission screen (who's allowed to use the script)

7. Click **☰ Menu → Google Auth platform → Branding**. If you see **Get started**, click it.
8. **App name:** `BudgetTool`. **User support email:** your Gmail. Click **Next**.
9. **Audience:** pick **External**. (Internal only exists for business Google accounts.) Click **Next**.
10. **Contact information:** your Gmail. Click **Next**.
11. Tick the agreement box → **Continue** → **Create**.
12. Click **Audience** in the left menu → under **Test users**, click **+ Add users** → type your Gmail → **Save**.

## Part 4: Download your key file

13. Click **Clients** in the left menu → **+ Create client**.
14. **Application type:** **Desktop app**. **Name:** `BudgetTool PyCharm`. Click **Create**.
15. Click **Download JSON** in the pop-up (or the ⬇ icon next to the client in the list).
16. Rename the downloaded file to exactly **`google_credentials.json`**.
17. Move it into **`BudgetTool\delivery\`**.

⚠ **This file is a secret key.** Never upload it, share it, or paste it into ChatGPT or Claude.
`backup.py` already skips it, so it never ends up in a backup zip.

## Part 5: First run (in PyCharm)

18. Open PyCharm's **Terminal** tab (the normal one, not Claude) and install the new libraries:
    ```bash
    pip install -r requirements.txt
    ```
19. Do a practice run that changes nothing:
    ```bash
    python tools/push_to_sheets.py --dry-run
    ```
20. Your browser opens. Pick your Google account.
    - "Google hasn't verified this app" → **Continue**. (It's your own app, and only you can use it.)
    - Tick **See, edit, create, and delete all your Google Drive files** → **Continue**.
    - The browser says "The authentication flow has completed." Close that tab.
21. ✅ The terminal should print:
    ```
    Pushing  Paycheck_Planner_v2.1.0_sage.xlsx
       into  Paycheck_Planner_v2.0.0_sage  ->  renamed 'Paycheck Planner'
       link  https://docs.google.com/spreadsheets/d/...
    Dry run: nothing was changed.
    ```
22. Now the real push:
    ```bash
    python tools/push_to_sheets.py
    ```
    ✅ It should end with: `Done. Same link, new content (...)`

## Part 6: Check it worked

- [ ] Reload your Google Sheet. It's now named **Paycheck Planner**.
- [ ] Reports → Month by month has **Paydays / Expected pay / Actually received** columns (that's v2.1).
- [ ] Spending has the **Money in** table on the right.
- [ ] The **Budget Tools** menu is still there. If it's gone, tell Claude. You'd just re-paste the Apps Script once.
- [ ] Dropdowns still work (Setup C9, C43 · Spending D5 · Dashboard C2).

---

## Every time after this
```bash
python tools/build_template.py --push
```
No browser and no clicking: it remembers you in `delivery\google_token.json` (also secret, also skipped by backups).

## Good to know
- **Signed out every 7 days?** Google gives "Testing" apps a 7-day pass. When it runs out, the browser just pops up again; click through and you're done. To stop that: **Google Auth platform → Audience → Publish app**. It stays unverified, which is fine for an app only you use.
- **Want to push a different look?** Change `"push_theme"` and `"push_style"` in `config/product.json` → `"google"`.
- **Pushed by mistake?** In the Google Sheet, go to **File → Version history → See version history**, pick the older version, and click **Restore**.

| Error | Means | Fix |
|---|---|---|
| `Missing delivery\google_credentials.json` | Key file isn't in place | Steps 15–17. Check the exact name: Windows may hide `.json`, so it might really be `google_credentials.json.json` |
| `Access blocked: BudgetTool has not completed the Google verification process` | Your Gmail isn't a test user | Step 12 |
| `HttpError 403 ... Drive API has not been used` | API not turned on | Step 6, then wait a minute |
| `HttpError 404 ... File not found` | Wrong Sheet ID, or a different Google account | Check `"master_sheet_id"`, and sign in with the account that owns the sheet |
| `push_to_sheets did not convert` | Google stored it as Excel instead of a Sheet | Tell Claude: there's a fallback plan |
