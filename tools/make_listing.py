"""
make_listing.py — writes ready-to-paste Etsy + Gumroad listing copy from
config/listing.json and config/product.json, and checks Etsy's limits for you.

Run:
    python tools/make_listing.py          # template-based copy, no internet needed
    python tools/make_listing.py --ai     # ALSO asks ChatGPT or Claude for fresh variations

For --ai, set ONE of these first (never paste keys into the code):
    OPENAI_API_KEY      (ChatGPT)   optional OPENAI_MODEL, default below
    ANTHROPIC_API_KEY   (Claude)    optional ANTHROPIC_MODEL, default below
"""
import json
import os
import sys
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
PRODUCT = json.loads((ROOT / "config" / "product.json").read_text())["product"]
L = json.loads((ROOT / "config" / "listing.json").read_text())
OUT = ROOT / "output" / "listing_kit.md"

ETSY_TITLE_MAX = 140
ETSY_TAG_MAX_COUNT = 13
ETSY_TAG_MAX_LEN = 20


def build_title():
    title = f'{L["main_keyword"]}'
    for kw in L["title_keywords"]:
        candidate = f"{title} | {kw}"
        if len(candidate) > ETSY_TITLE_MAX:
            break
        title = candidate
    return title


def check_tags(tags):
    problems = []
    if len(tags) > ETSY_TAG_MAX_COUNT:
        problems.append(f"Etsy allows {ETSY_TAG_MAX_COUNT} tags; you have {len(tags)}.")
    for t in tags:
        if len(t) > ETSY_TAG_MAX_LEN:
            problems.append(f'Tag too long ({len(t)} chars, max {ETSY_TAG_MAX_LEN}): "{t}"')
    return problems


def description(platform):
    lines = [f'{PRODUCT["name"]}: {PRODUCT["tagline"]}.', ""]
    lines.append("Most budgets pretend your whole paycheck is yours. This one shows what "
                 "you ACTUALLY take home after taxes, then helps you give every dollar a job.")
    lines += ["", "WHAT IT DOES"] + [f"✔ {f}" for f in L["features"]]
    lines += ["", "WHAT YOU GET"] + [f"• {w}" for w in L["whats_included"]]
    lines += ["", "FAQ"]
    for q, a in L["faq"]:
        lines += [f"Q: {q}", f"A: {a}"]
    if platform == "etsy":
        lines += ["", "This is a DIGITAL download. Nothing physical will be shipped."]
    return "\n".join(lines)


# ---------------------------------------------------------------- optional AI
def ask_ai(prompt):
    if os.getenv("OPENAI_API_KEY"):
        req = urllib.request.Request(
            "https://api.openai.com/v1/chat/completions",
            data=json.dumps({"model": os.getenv("OPENAI_MODEL", "gpt-4o-mini"),
                             "messages": [{"role": "user", "content": prompt}]}).encode(),
            headers={"Authorization": f'Bearer {os.environ["OPENAI_API_KEY"]}',
                     "Content-Type": "application/json"})
        with urllib.request.urlopen(req, timeout=60) as r:
            return json.load(r)["choices"][0]["message"]["content"]
    if os.getenv("ANTHROPIC_API_KEY"):
        req = urllib.request.Request(
            "https://api.anthropic.com/v1/messages",
            data=json.dumps({"model": os.getenv("ANTHROPIC_MODEL", "claude-sonnet-5-5"),
                             "max_tokens": 1500,
                             "messages": [{"role": "user", "content": prompt}]}).encode(),
            headers={"x-api-key": os.environ["ANTHROPIC_API_KEY"],
                     "anthropic-version": "2023-06-01",
                     "content-type": "application/json"})
        with urllib.request.urlopen(req, timeout=60) as r:
            return json.load(r)["content"][0]["text"]
    return None


def main():
    title = build_title()
    tags = L["tags"]
    problems = check_tags(tags)

    out = [f'# Listing kit — {PRODUCT["name"]} v{PRODUCT["version"]}', "",
           f'Suggested price: ${PRODUCT["price_usd"]}', "",
           f"## Etsy title ({len(title)}/{ETSY_TITLE_MAX} chars)", "", title, "",
           f"## Etsy tags ({len(tags)}/{ETSY_TAG_MAX_COUNT})", "", ", ".join(tags), ""]
    if problems:
        out += ["### ⚠ Fix these before posting"] + [f"- {p}" for p in problems] + [""]
    out += ["## Etsy description", "", "```", description("etsy"), "```", "",
            "## Gumroad description", "", "```", description("gumroad"), "```", "",
            "## Short social captions", "",
            f'- Your $1,200 paycheck isn\'t really $1,200. {PRODUCT["name"]} shows what you actually keep after taxes. Link in bio.',
            "- I built the budget sheet I wish I had at my first job: custom %, real tax math, inflation built in.",
            "- POV: you finally know where your paycheck goes. (Google Sheets + Excel)", ""]

    if "--ai" in sys.argv:
        prompt = ("You write Etsy listings for digital spreadsheet products. Product:\n"
                  f"{json.dumps({'product': PRODUCT, 'listing': L}, indent=2)}\n\n"
                  "Give me: 3 alternative Etsy titles under 140 characters, 5 scroll-stopping "
                  "opening lines for the description, and 5 TikTok hook ideas for students and "
                  "first-jobbers. Plain text, no tax promises.")
        ai = ask_ai(prompt)
        out += ["## AI variations", ""] + ([ai] if ai else
               ["(No OPENAI_API_KEY or ANTHROPIC_API_KEY set, so this part was skipped.)"])

    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text("\n".join(out) + "\n")
    print(f"Wrote {OUT.relative_to(ROOT)}")
    for p in problems:
        print("WARNING:", p)


if __name__ == "__main__":
    main()
