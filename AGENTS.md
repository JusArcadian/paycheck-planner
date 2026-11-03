# AGENTS.md — rules for Codex in this project

Codex reads this file automatically every time it starts in this folder.
These are the owner's standing rules. Follow all of them on every change.

## About the owner
- Beginner programmer. Explain things in plain English, and "like I'm 5" when something is new or confusing.
- Wants step-by-step instructions for anything they have to do themselves (numbered, one action per step).
- Works on Windows 11 in PyCharm. Project folder: `C:\Users\terre\OneDrive\Documents\Portfolio\BudgetTool`

## About the project
Paycheck Planner: a sellable budget template (Excel + Google Sheets) for students and first-jobbers,
plus scripts to build, list and sell it. Start with `README.md` for the map of every file.

## After EVERY change (this checklist is required)
1. **Back up first:** `python tools/backup.py "what you're changing"`
2. Make the change. Prefer editing `config/*.json` over changing the scripts.
3. If the change touches the template: run `python tools/build_template.py` and check every "Built" line appears.
   Only add `--push` (updates the live Google Sheet buyers copy) when the owner says the change is ready.
4. **Update `PATCH_NOTES.md`:** bump the version and list the change under Features / Bugs fixed / Buffs / Nerfs.
5. **Update `docs/STUDY_GUIDE.md`:** add or extend the right chapter with the new code and a simple explanation.
   Explain any error that came up in Chapter 6 ("When things break").
6. **Update `tools/SOURCES.md`** if you used a new website, doc, data source or number.
7. **Back up again** with a note like `"after: added X"`.
7b. **If the project uses Git:** commit with a short message (e.g. `v2.2.1: add .gitignore`). Never commit anything `.gitignore` blocks.
    Only push to GitHub when the owner says so.
8. **In chat, give the owner:**
   - a short summary of what changed
   - the actual code that changed, with a plain-English explanation of how it works
   - a test checklist: what to click or run, which command or key to use, and what they should see

## Coding rules
- Python 3: standard library + the libraries in `requirements.txt` only. Ask before adding a new library.
- Spreadsheet formulas: Excel-2007-era functions only (SUMIFS, INDEX, MATCH, SUMPRODUCT, IFERROR).
  No XLOOKUP, FILTER, SORT, UNIQUE, SEQUENCE: they break in Google Sheets/LibreOffice.
- Dropdowns: type the choices into the dropdown, or point at a list on the SAME tab. Never another tab (Google Sheets can drop it).
- Never hardcode a calculated number in the spreadsheet. Always use a formula.
- Comment code so a beginner can follow it.

## Safety rules
- Never read, print, copy or commit `delivery/.env`, `delivery/google_credentials.json` or `delivery/google_token.json`. They are secrets.
- Never put real buyer emails or personal budgets in the project files.
- `reference/` holds the owner's real budget. Borrow ideas only; never copy its data into the product.
- Tax and inflation numbers are estimates. Every number needs a source in `tools/SOURCES.md`.
- Ask before deleting any file.
