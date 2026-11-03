"""
menu.py — one command instead of remembering script names.

In PyCharm's Terminal type:   python menu.py
Then press a number and Enter.
"""
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
PY = sys.executable  # the same Python that runs this menu (your .venv)

CHOICES = [
    ("Back up the project",                        ["tools/backup.py", "{note}"]),
    ("Update inflation (monthly)",                 ["tools/update_inflation.py"]),
    ("Build the templates",                        ["tools/build_template.py"]),
    ("Build + push to Google Sheets (Update)",     ["tools/build_template.py", "--push"]),
    ("Push to Google Sheets: practice (no change)", ["tools/push_to_sheets.py", "--dry-run"]),
    ("Write the Etsy/Gumroad listing copy",         ["tools/make_listing.py"]),
    ("Clean up old spreadsheets (keep newest 3)",   ["tools/clean_output.py"]),
]


def run(args):
    print(f"\n> python {' '.join(args)}\n")
    result = subprocess.run([PY, *args], cwd=ROOT)
    print("\n✔ Done." if result.returncode == 0 else "\n✖ That didn't work. Copy the red text above and send it to Claude.")


def main():
    while True:
        print("\n=== Paycheck Planner menu ===")
        for i, (label, _) in enumerate(CHOICES, start=1):
            print(f"  {i}. {label}")
        print("  0. Quit")
        pick = input("Pick a number: ").strip()
        if pick in ("0", "q", ""):
            return
        if not pick.isdigit() or not 1 <= int(pick) <= len(CHOICES):
            print("Type one of the numbers above.")
            continue
        label, args = CHOICES[int(pick) - 1]
        if "{note}" in args:
            note = input("What are you changing? (a few words): ").strip() or "manual backup"
            args = [a.replace("{note}", note) for a in args]
        run(args)


if __name__ == "__main__":
    main()
