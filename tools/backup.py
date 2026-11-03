"""
backup.py — saves a dated zip of the whole project into backups/.

Run before AND after any change:   python tools/backup.py "added dark theme"
Skips: backups/, output/ (rebuildable), .idea/ and .venv/ (PyCharm stuff), and .env (secrets never go in backups).
"""
import sys
import zipfile
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SKIP_DIRS = {"backups", "output", "__pycache__", ".git", ".idea", ".venv"}  # .idea/.venv = PyCharm settings + libraries
SKIP_FILES = {".env", "sales_log.csv", "google_credentials.json", "google_token.json"}  # secrets


def main():
    note = "-".join(sys.argv[1:]).replace(" ", "-")[:40] or "backup"
    note = "".join(ch for ch in note if ch not in '<>:"/\\|?*')   # Windows can't use these in file names
    stamp = datetime.now().strftime("%Y-%m-%d_%H%M")
    dest = ROOT / "backups" / f"budget-kit_{stamp}_{note}.zip"
    dest.parent.mkdir(exist_ok=True)
    count = 0
    with zipfile.ZipFile(dest, "w", zipfile.ZIP_DEFLATED) as z:
        for path in ROOT.rglob("*"):
            rel = path.relative_to(ROOT)
            if path.is_dir() or rel.parts[0] in SKIP_DIRS or path.name in SKIP_FILES:
                continue
            if any(p in SKIP_DIRS for p in rel.parts):
                continue
            z.write(path, rel)
            count += 1
    print(f"Backed up {count} files -> {dest.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
