# Launch Guide: zero to first sale

Do these in order. Tick each box as you go. Don't skip Phase 1: everything else depends on it.

---

## Phase 1 — Get the kit running on your computer (30 min)

- [ ] **1. Install Python.** Go to python.org → Downloads → install the latest 3.x.
  On Windows, tick **"Add Python to PATH"** on the first screen.
- [ ] **2. Unzip the kit** somewhere easy, like `Documents/budget-kit`.
- [ ] **3. Open a terminal in that folder.**
  Windows: open the folder, click the address bar, type `cmd`, press Enter.
  Mac: right-click the folder → **New Terminal at Folder**.
- [ ] **4. Install the 3 libraries the scripts need:**
  ```bash
  pip install -r requirements.txt
  ```
- [ ] **5. Build the templates:**
  ```bash
  python tools/build_template.py
  ```
  ✅ You should see 3 "Built output/..." lines.
- [ ] **6. Make your first backup:**
  ```bash
  python tools/backup.py "first setup"
  ```

## Phase 2 — Test it like a buyer (30 min)

Open `output/Paycheck_Planner_v2.0.0_sage.xlsx` and check each one:

| Test | Where | What should happen |
|---|---|---|
| Change paycheck to 1500 | Setup C8 | "Take-home pay" goes up, and so does the Dashboard's first big number |
| Pick "Use my own rate" | Setup C15 | Federal/state lines become $0; "Taxes at your own rate" fills in |
| Pick "70/20/10 (tight)" | Setup C43 | "% in use" changes; message still says ✔ 100% |
| Pick "Custom" and set Needs to 60% | Setup C43, D45 | Warning says your groups add up to 110% |
| Log a purchase dated today | Spending row 6 | Dashboard "Spent so far", "Spent today" and its category card update |
| Spend more than a category goal | Setup D column + Spending | That category turns red on the Dashboard card |
| Change the month | Dashboard C2 | Every number switches to that month |
| Look at Reports → Month by month | Reports | Biweekly pay: 2 months show **3** paydays; the year totals your full take-home |
| Log a deposit smaller than one paycheck | Spending L5:N5 (Money in) | Dashboard says "$X still coming"; that paycheck in Reports says "Partly paid" |

- [ ] All 9 tests pass. (If one doesn't, check `STUDY_GUIDE.md` Chapter 6, "When things break".)

## Phase 3 — Make the Google Sheets version (15 min)

- [ ] Follow `google-sheets/SETUP.md` step by step.
- [ ] Paste your `/copy` link into `config/product.json`.

## Phase 4 — Make the buyer PDF (20 min)

Buyers get a PDF from Etsy/Gumroad. The PDF holds the Google Sheets link.
- [ ] In Canva (free), make a 1–2 page PDF: product name, the big **"Get your Google Sheet"** button linking to your `/copy` link, the 4 steps from the Start Here tab, and "Estimates only, not tax advice."
- [ ] Click the link in the exported PDF to make sure it works.

## Phase 5 — Listing photos (45 min)

Etsy buyers decide from the photos. You need at least 5:
- [ ] 1. Hero image: dashboard screenshot on a laptop mockup, with the title in big text
- [ ] 2. "See your REAL take-home" (zoom on the tax breakdown)
- [ ] 3. "Your categories, your %" (Categories tab)
- [ ] 4. "Inflation built in" (next year's targets)
- [ ] 5. "What's included" (Sheets + Excel + 3 themes, show all three colors)

Tip: put `$1,200 biweekly` example numbers in before screenshots. Never use your real budget.

## Phase 6 — List it (30 min)

- [ ] Run `python tools/make_listing.py` and open `output/listing_kit.md`.
- [ ] **Gumroad** (easiest to start): New product → Digital product → upload the 3 xlsx files + PDF → paste the Gumroad description → price $9.
- [ ] **Etsy:** Add listing → Digital → upload the PDF + xlsx files → paste title, tags, description. Check Etsy's current listing and transaction fees first.
- [ ] Buy your own product once (or use Gumroad's test purchase) to see what buyers see.

## Phase 7 — Turn on the welcome email (optional, 45 min)

Gumroad already sends the files. This adds a personal welcome email and your own sales log.
- [ ] Copy `delivery/.env.example` to `delivery/.env` and fill it in (see the Study Guide, Chapter 5).
- [ ] Test locally with `DRY_RUN=true`: `python delivery/gumroad_webhook.py`
- [ ] Host it free on Render.com: New → Web Service → connect your GitHub repo → start command `gunicorn delivery.gumroad_webhook:app`. Add the `.env` values in Render's **Environment** tab, NOT in the code.
- [ ] Gumroad → Settings → Advanced → **Ping endpoint**: `https://YOUR-APP.onrender.com/gumroad?key=YOUR_WEBHOOK_SECRET`
- [ ] Click "Send test ping". A row appears in your sales log.
- [ ] Switch `DRY_RUN=false` when you're happy.

## Phase 8 — Get eyes on it (ongoing)

- [ ] Open `docs/PROMPT_PACK.md` → "30-day posting plan". Post 1 short video a day.
- [ ] Every month: `update_inflation.py` → `build_template.py` → re-upload files → backup → patch notes.
- [ ] Every November: check the new IRS brackets and update `config/product.json`.
