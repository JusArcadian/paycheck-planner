# Toolkit: programs and APIs that make everything work together

What's worth adding, why, and when. **★ = do this soon.** The rest are for when you need them.
These work for Paycheck Planner AND your other projects (Roblox games, the Minecraft mod, apps).

---

## 1. Safety net and history

| Tool | What it does for you | Cost |
|---|---|---|
| ★ **Git + GitHub** | Every change saved with a note, and you can rewind to any point. Built into PyCharm (**VCS → Enable Version Control**). A free private GitHub repo is an off-site backup, and a public one is a portfolio | Free |
| ★ **Password manager** (Bitwarden, 1Password…) | Stores your Google, Gumroad, Etsy and GitHub logins plus 2FA codes. You'll have a lot of accounts soon | Bitwarden has a free tier |
| ★ **2-step verification everywhere** | Google, GitHub, Gumroad, Etsy, bank. Your store is money: lock it | Free |
| **`.gitignore` file** | Tells Git to never save secrets (`.env`, `google_credentials.json`, `google_token.json`) or junk (`.venv`, `output/`) | Free |

> `backup.py` zips are your safety net today. Git is the upgrade: it shows exactly *what* changed, line by line.

## 2. Glue: making apps talk to each other

| Tool | Best for | Example for you |
|---|---|---|
| **Python scripts** (what you have) | Anything custom | `build_template.py --push`, `update_inflation.py` |
| **Windows Task Scheduler** | Running a script on a schedule on your PC | Run `update_inflation.py` on the 15th of every month |
| **GitHub Actions** | Running scripts in the cloud on a schedule, even with your PC off | Monthly: update inflation → build → push to Google Sheets |
| **n8n / Zapier / Make** | Drag-and-drop "when X happens, do Y" between apps, with no server to manage | New Gumroad sale → add a row to a sales Sheet → ping your phone |
| **Webhooks** | One app pinging another the moment something happens | Your `gumroad_webhook.py` already does this |
| **Discord webhooks** | Instant notifications to your own Discord channel. Easy, no login code | "💸 New sale: Paycheck Planner $9" |

## 3. Claude features you can lean on

| Feature | What it does |
|---|---|
| **Claude Code in PyCharm** | Edits code, runs scripts, follows `CLAUDE.md` |
| **Remote Control** (`/remote-control`) | Drive that PyCharm session from your phone |
| **Connectors** (you have Google Drive + Sheets) | Claude reads and checks your Google files directly. Others worth connecting: Gmail, Google Calendar, Notion, GitHub |
| **Scheduled tasks** | Claude runs a job on a schedule, e.g. "every month, check new inflation data and tell me if the template needs rebuilding" |
| **CLAUDE.md per project** | Make one for each project (Roblox, mod) so Claude follows your rules everywhere |

## 4. APIs worth knowing

| API | Use | Key needed? |
|---|---|---|
| BLS (inflation) | Already used | No |
| Google Drive / Sheets | Already used (push) | Yes (done in setup) |
| Gumroad | Sales, products, webhooks | Yes |
| Etsy Open API | Manage listings from scripts. Requires applying for access | Yes |
| OpenAI / Anthropic (Claude) | AI text inside your own scripts (`make_listing.py --ai`) | Yes, pay per use |
| Discord webhooks | Notifications | Just a URL |
| Stripe | Taking payments on your own site later | Yes |

## 5. Making and selling

| Tool | Use |
|---|---|
| **Canva** | Buyer PDF, Etsy listing images, TikTok covers |
| **Gumroad / Etsy** | Where the template sells |
| **Notion or Obsidian** | One home for ideas, to-dos and notes across all projects |
| **Google Drive for desktop** | Makes Drive show up as a folder on your PC |

## 6. For your other projects

| Project | Tool | Why |
|---|---|---|
| Roblox games | **Rojo** + VS Code | Write Roblox scripts in a real editor and keep them in Git |
| Minecraft mod | IntelliJ IDEA + Gradle (Fabric/Forge MDK) | You already have IntelliJ, the standard for Java mods |
| Apps (Localy, CramCal, budgeting app) | GitHub + a hosting service (Render, Vercel…) | Put ideas online fast |

---

## The "it just runs" pipeline (what full automation could look like)

```
On the 15th of each month (GitHub Actions or Task Scheduler)
  └─ update_inflation.py   → new inflation number?
       └─ yes → build_template.py --push   → Google Sheet updated (same buyer link)
              → backup.py                   → zip saved
              → Discord ping                → "Template rebuilt: inflation 3.7% → 3.5%"

Every sale (Gumroad webhook)
  └─ gumroad_webhook.py → welcome email + sales log → Discord ping "💸 New sale"
```
You don't touch anything. You just get the pings. Ask Claude to build any piece of this, one step at a time.
