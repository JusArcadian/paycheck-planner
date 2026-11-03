"""
push_to_sheets.py — puts the newest built template INTO your master Google Sheet.

It swaps the Sheet's contents for the new .xlsx but keeps the SAME file, so:
  - the link (and your buyer /copy link) never changes
  - buyers who already made a copy are never touched

Run:
    python tools/push_to_sheets.py             # push the newest build
    python tools/push_to_sheets.py --dry-run   # just show what it WOULD do
    python tools/build_template.py --push      # build everything, then push

Setup (one time): docs/GOOGLE_PUSH_SETUP.md
Secrets it uses (never share these, never upload them):
    delivery/google_credentials.json   (downloaded from Google Cloud)
    delivery/google_token.json         (created after your first sign-in)
"""
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
CONFIG = json.loads((ROOT / "config" / "product.json").read_text())
CREDENTIALS = ROOT / "delivery" / "google_credentials.json"
TOKEN = ROOT / "delivery" / "google_token.json"
# Full Drive access is needed because the master Sheet was created by you, not by this script.
SCOPES = ["https://www.googleapis.com/auth/drive"]
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
SHEET_MIME = "application/vnd.google-apps.spreadsheet"


def build_path():
    """Which .xlsx to push: the theme + style named in config/product.json -> "google"."""
    p, g = CONFIG["product"], CONFIG["google"]
    suffix = "" if g["push_style"] == "card" else f"_{g['push_style']}"
    return ROOT / "output" / f"{p['name'].replace(' ', '_')}_v{p['version']}_{g['push_theme']}{suffix}.xlsx"


def google_service():
    """Sign in to Google (browser pops up the first time) and return a Drive connection."""
    try:
        from google.auth.exceptions import RefreshError
        from google.auth.transport.requests import Request
        from google.oauth2.credentials import Credentials
        from google_auth_oauthlib.flow import InstalledAppFlow
        from googleapiclient.discovery import build
    except ImportError:
        sys.exit("Google libraries missing. In the PyCharm terminal run:  pip install -r requirements.txt")

    creds = None
    if TOKEN.exists():
        creds = Credentials.from_authorized_user_file(str(TOKEN), SCOPES)
    if creds and not creds.valid and creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())          # quietly get a fresh pass
        except RefreshError:
            creds = None                      # pass expired (Testing apps: every 7 days) -> sign in again
    if not creds or not creds.valid:
        if not CREDENTIALS.exists():
            sys.exit(f"Missing {CREDENTIALS.relative_to(ROOT)}. Follow docs/GOOGLE_PUSH_SETUP.md, step 4.")
        flow = InstalledAppFlow.from_client_secrets_file(str(CREDENTIALS), SCOPES)
        creds = flow.run_local_server(port=0)   # opens your browser to sign in
        TOKEN.write_text(creds.to_json())
    return build("drive", "v3", credentials=creds)


def push(dry_run=False, service=None):
    g = CONFIG["google"]
    sheet_id = g.get("master_sheet_id", "")
    if not sheet_id or "PASTE" in sheet_id:
        sys.exit('Put your Google Sheet ID in config/product.json -> "google" -> "master_sheet_id".')
    xlsx = build_path()
    if not xlsx.exists():
        sys.exit(f"{xlsx.relative_to(ROOT)} doesn't exist yet. Run: python tools/build_template.py")

    service = service or google_service()
    meta = service.files().get(fileId=sheet_id, fields="name,mimeType,webViewLink").execute()
    if meta["mimeType"] != SHEET_MIME:
        sys.exit(f"'{meta['name']}' isn't a Google Sheet (it's {meta['mimeType']}). Check master_sheet_id.")
    new_name = g.get("master_sheet_name") or CONFIG["product"]["name"]
    print(f"Pushing  {xlsx.name}")
    print(f"   into  {meta['name']}  ->  renamed '{new_name}'")
    print(f"   link  {meta['webViewLink']}")
    if dry_run:
        print("Dry run: nothing was changed.")
        return meta

    from googleapiclient.http import MediaFileUpload
    media = MediaFileUpload(str(xlsx), mimetype=XLSX_MIME, resumable=False)
    result = service.files().update(fileId=sheet_id, body={"name": new_name}, media_body=media,
                                    fields="id,name,mimeType,modifiedTime").execute()
    if result["mimeType"] != SHEET_MIME:
        sys.exit("Google kept the upload as an Excel file instead of converting it. Tell Claude: "
                 "'push_to_sheets did not convert'.")
    print(f"Done. Same link, new content ({result['modifiedTime']}). Reload the Google Sheet to see it.")
    return result


if __name__ == "__main__":
    push(dry_run="--dry-run" in sys.argv)
