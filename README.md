# Paycheck Planner — Product Kit

A budget template for students and first-jobbers, plus the tools to build, list, sell and support it.
**This file is the map. Start here every time.**

## What's in the box

| Folder / file | What it is | When you touch it |
|---|---|---|
| `config/product.json` | Name, price, colors, default categories, tax brackets, inflation default | To change the product |
| `config/listing.json` | Keywords, tags, features, FAQ for your listing | To change listing copy |
| `tools/build_template.py` | Builds the Excel templates (all 3 themes) | After any config change |
| `tools/update_inflation.py` | Pulls the latest US inflation rate into the config | Once a month |
| `tools/make_listing.py` | Writes Etsy + Gumroad copy and checks Etsy limits (`--ai` for ChatGPT/Claude ideas) | Before listing |
| `tools/backup.py` | Zips the project into `backups/` | Before and after every change |
| `tools/push_to_sheets.py` | Updates the master Google Sheet with the newest build (same link) | After building (`build_template.py --push`) |
| `tools/SOURCES.md` | Every outside number, API and link, with when to re-check it | When you use a new source |
| `CLAUDE.md` | Rules Claude follows when it works in this folder (PyCharm or phone) | When your rules change |
| `google-sheets/` | Inflation auto-update script + how to make the Sheets version | One-time setup |
| `delivery/` | Gumroad sale webhook (welcome email + sales log) and sales report | After you're selling |
| `docs/LAUNCH_GUIDE.md` | Step-by-step from zero to first sale | Your to-do list |
| `docs/STUDY_GUIDE.md` | Every piece explained simply, with code walkthroughs | When something's confusing |
| `docs/PROMPT_PACK.md` | Copy-paste prompts for ChatGPT/Claude | Marketing, support, ideas |
| `docs/GOOGLE_PUSH_SETUP.md` | One-time setup so pushing to Google Sheets works | Once |
| `docs/SHORTCUTS.md` | Every link and keyboard shortcut in one place | Bookmark it |
| `menu.py` | Numbered menu that runs the scripts for you | Every day |
| `docs/TOOLKIT.md` | Programs, APIs and automation ideas | When you want to automate more |
| `reference/` | Your old planner (private, has real data) | Ideas only |
| `PATCH_NOTES.md` | Every change, by version | Update after every change |
| `output/` | Built files (xlsx + listing_kit.md). Safe to delete, just rebuild | — |

## Easiest way: the menu

```bash
python menu.py
```
Pick a number. All links and keyboard shortcuts: `docs/SHORTCUTS.md`.

## The 4 commands you'll use most

```bash
python tools/backup.py "what I'm about to change"   # 1. back up first
python tools/update_inflation.py                    # 2. (monthly) refresh inflation
python tools/build_template.py --push               # 3. rebuild the templates + update Google Sheet
python tools/make_listing.py                        # 4. regenerate listing copy
```

## Golden rules
1. Edit **config files**, not the scripts, for everyday changes.
2. Back up before and after changes, and add a line to `PATCH_NOTES.md`.
3. Never share or upload `delivery/.env`, `delivery/google_credentials.json` or `delivery/google_token.json`. They're keys.
4. Tax numbers are estimates. Re-check the IRS brackets every October/November when new ones come out.
