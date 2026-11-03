"""
clean_output.py - keeps only the newest versions of the built spreadsheets in output/.

The owner gave standing permission (2026-10-22): delete old files in output/, keep the
newest 3 versions (the current one + the 2 before it). Change the number in
config/product.json -> "keep_output_versions".

Run:  python tools/clean_output.py            (asks before deleting)
      python tools/clean_output.py --preview  (only shows what it would delete)
      python tools/clean_output.py --yes      (no question; build_template.py uses this)

ELI5: the output folder is a shelf. We keep the 3 newest boxes and throw out the rest.
Nothing else is touched: only .xlsx files named like Paycheck_Planner_v2.6.0_sage.xlsx.
Old versions can always be rebuilt from a zip in backups/.
"""
import json
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "output"
CONFIG = json.loads((ROOT / "config" / "product.json").read_text(encoding="utf-8"))
KEEP = int(CONFIG.get("keep_output_versions", 3))
NAME = re.compile(r"_v(\d+)\.(\d+)\.(\d+)_.+\.xlsx$")   # finds the "2.6.0" in the file name


def old_files():
    """Returns (versions kept, files to delete)."""
    by_version = {}
    for f in OUT.glob("*.xlsx"):
        m = NAME.search(f.name)
        if m:
            by_version.setdefault(tuple(int(x) for x in m.groups()), []).append(f)
    versions = sorted(by_version, reverse=True)            # newest first: (2,6,0) > (2,5,0) > (2,4,1)
    keep = versions[:KEEP]
    delete = [f for v in versions[KEEP:] for f in sorted(by_version[v])]
    return keep, delete


def main(args):
    keep, delete = old_files()
    show = lambda v: "v" + ".".join(map(str, v))
    print(f"Keeping the newest {KEEP}: {', '.join(show(v) for v in keep) or 'nothing built yet'}")
    if not delete:
        print("Nothing old to delete. ✔")
        return
    print(f"Old files ({len(delete)}):")
    for f in delete:
        print("   -", f.name)
    if "--preview" in args:
        print("(Preview only. Nothing was deleted.)")
        return
    if "--yes" not in args and input("Delete these? Type Y and press Enter: ").strip().lower() != "y":
        print("Okay, nothing deleted.")
        return
    gone = 0
    for f in delete:
        try:
            f.unlink()
            gone += 1
        except PermissionError:
            print(f"✖ Couldn't delete {f.name}: it's open in Excel. Close it and run this again.")
    print(f"Deleted {gone} old file(s). ✔")


if __name__ == "__main__":
    main(sys.argv[1:])
