# Sources

Every outside fact, number, API and service this project relies on, with where it came from
and when to check it again. **Add a row whenever you use something new.**

Last reviewed: 2026-10-22

---

## 1. Tax numbers (used in `config/product.json` → Tax Tables tab)

| What | Value we use | Source | Re-check |
|---|---|---|---|
| 2026 federal brackets & standard deduction | $16,100 single / $32,200 joint / $24,150 head of household; 10% up to $12,400 single; 37% over $640,600 single | [IRS newsroom: 2026 inflation adjustments](https://www.irs.gov/newsroom/irs-releases-tax-inflation-adjustments-for-tax-year-2026-including-amendments-from-the-one-big-beautiful-bill) · [Accounting Today summary](https://www.accountingtoday.com/news/irs-announces-inflation-adjustments-for-2026) · [CPA Practice Advisor](https://www.cpapracticeadvisor.com/2025/10/09/irs-adjusts-tax-brackets-standard-deduction-for-2026/170661/) | Every Oct/Nov when the IRS releases the next year |
| Middle bracket floors (12%–35%) | Single and joint: see `product.json` | IRS page above (single + married filing jointly). Head of household: [NerdWallet 2026 brackets](https://www.nerdwallet.com/taxes/learn/2026-tax-brackets) (IRS Rev. Proc. 2025-32 has the official table) | Same as above |
| FICA rate | 7.65% (6.2% Social Security + 1.45% Medicare) | [IRS Topic 751: Social Security and Medicare withholding](https://www.irs.gov/taxtopics/tc751) | Rarely changes. Check yearly |

**Not modeled (known limits):** Social Security wage cap, tax credits, itemized deductions, local taxes beyond the single state % field.

## 2. Inflation

| What | Value we use | Source | Re-check |
|---|---|---|---|
| Default "Auto" inflation rate | 3.7% (US CPI-U, Sep 2026, released Oct 14 2026) | [BLS CPI home](https://www.bls.gov/cpi/) · [CNN report](https://www.cnn.com/2026/10/14/economy/us-cpi-inflation-september) · [CNBC report](https://www.cnbc.com/2026/10/14/cpi-inflation-report-september-2026.html) | Monthly. BLS releases mid-month |
| CPI data series | `CUUR0000SA0` = CPI-U, all items, US city average, not seasonally adjusted | [BLS CPI home](https://www.bls.gov/cpi/) | Doesn't change |

## 3. APIs and services the scripts call

| Used by | Service | Docs | Needs a key? |
|---|---|---|---|
| `tools/update_inflation.py`, `google-sheets/inflation_autoupdate.gs` | BLS Public Data API v1 | [BLS API: Getting Started](https://www.bls.gov/developers/) | No (v1 needs no registration; v2 does) |
| `delivery/gumroad_webhook.py` | Gumroad Ping (webhook) | [Gumroad Help Center](https://gumroad.com/help) (search "Ping") | No, but uses your own `WEBHOOK_SECRET` |
| `delivery/sales_report.py` | Gumroad API v2 `/sales` | [Gumroad API docs](https://gumroad.com/api) | Yes: `GUMROAD_ACCESS_TOKEN` |
| `tools/make_listing.py --ai` | OpenAI Chat Completions | [OpenAI API docs](https://platform.openai.com/docs/api-reference/chat) | Yes: `OPENAI_API_KEY` |
| `tools/make_listing.py --ai` | Anthropic Messages API | [Claude API docs](https://docs.claude.com/en/api/overview) | Yes: `ANTHROPIC_API_KEY` |
| `tools/push_to_sheets.py` | Google Drive API v3 (`files.update`) | [Drive API Python quickstart](https://developers.google.com/workspace/drive/api/quickstart/python) | Yes: `google_credentials.json` (OAuth desktop client) |
| Welcome emails | Gmail SMTP + App password | [Google: sign in with app passwords](https://support.google.com/accounts/answer/185833) | Yes: app password in `.env` |
| Google Sheets script | Google Apps Script (`UrlFetchApp`, triggers) | [Apps Script docs](https://developers.google.com/apps-script) | No |

## 4. Selling platforms

| Platform | Rules / limits we follow | Source |
|---|---|---|
| Etsy | Title max 140 characters; 13 tags, 20 characters each; digital items need a "nothing shipped" note | [Etsy Seller Handbook](https://www.etsy.com/seller-handbook) · [Etsy fees](https://www.etsy.com/legal/fees/) |
| Gumroad | Digital products, Ping endpoint, access tokens | [Gumroad Help Center](https://gumroad.com/help) |
| Render.com | Free hosting for the webhook | [Render docs](https://render.com/docs) |

## 5. Code libraries

| Library | Used for | Docs |
|---|---|---|
| openpyxl | Writing the Excel templates | [openpyxl docs](https://openpyxl.readthedocs.io/) |
| Flask | The Gumroad webhook web server | [Flask docs](https://flask.palletsprojects.com/) |
| gunicorn | Running Flask on Render | [gunicorn docs](https://docs.gunicorn.org/) |

## 6. Design references

| What | Used for | Where |
|---|---|---|
| Owner's old "My Budget Dashboard" | Feature ideas for v2: category buckets, donut charts, daily allowance, spent today, paycheck / monthly / category reports. Only ideas were reused, no data | `reference/My_Budget_Dashboard_OLD.xlsx` (private: has real transactions, never ship it) |

## 8. Savings and investing rates (Accounts tab, v2.4.0)

| What | Value used | Source | Re-check |
|---|---|---|---|
| National average savings rate | 0.38% (`national_avg_savings_rate`) | [NerdWallet: best high-yield savings accounts (Oct 21, 2026)](https://www.nerdwallet.com/banking/best/high-yield-online-savings-accounts?trk=nww) | Every 3 months |
| Top HYSA rates | about 4–5% (example account uses 4%) | Same NerdWallet page; SoFi up to 4.40% promo, 3.30% standard | Every 3 months |
| S&P 500 average return | about 9.92%/yr since 1928; about 6.64% after inflation | [Motley Fool: S&P 500 annual returns](https://www.fool.com/research/sp-500-annual-returns/) | Once a year |
| FDIC insurance limit | $250,000 per depositor, per bank | [FDIC: deposit insurance](https://www.fdic.gov/resources/deposit-insurance) | Once a year |

## 7. Tools for building this project

| Tool | Used for | Docs |
|---|---|---|
| Claude Code + PyCharm plugin | Coding with Claude inside PyCharm | [Claude Code: JetBrains IDEs](https://code.claude.com/docs/en/jetbrains) |
| Claude Code Remote Control | Driving the PyCharm session from your phone | [Claude Code: Remote Control](https://code.claude.com/docs/en/remote-control) |
| Python | Running the scripts | [python.org downloads](https://www.python.org/downloads/) |

---

### How to add a source
1. Pick the right section (or make a new one).
2. Add a row: **what** it is, the **value** you used, the **link**, and **when to re-check** it.
3. Update "Last reviewed" at the top.

## 9. Bank holidays (paydays that land on one move to the business day before/after, v2.12.0)

| What | Value | Source | How often to check |
|---|---|---|---|
| Which days are bank holidays | New Year's, MLK Day, Presidents' Day, Memorial Day, Juneteenth, Independence Day, Labor Day, Columbus Day, Veterans Day, Thanksgiving, Christmas | [Federal Reserve: Holidays Observed (K.8)](https://www.federalreserve.gov/aboutthefed/k8.htm) | Once a year (new holidays are rare) |
| Sunday / Saturday rule | Sunday holiday → banks closed Monday. Saturday holiday → Reserve Banks stay OPEN Friday | Same page | Rarely changes |
| Checked | Engine list 2026-2030 matched the Fed's published dates exactly (2026-10-22) | Same page | — |
