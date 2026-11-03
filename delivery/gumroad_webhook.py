"""
gumroad_webhook.py — runs after every Gumroad sale.

What happens on each sale:
  1. Gumroad "pings" this little web server with the buyer's info.
  2. We check the secret key in the URL (so strangers can't fake sales).
  3. We log the sale to sales_log.csv.
  4. We email the buyer a welcome note with their Google Sheets copy link.

Gumroad ALREADY delivers your Excel/PDF files by itself. This adds the
personal welcome email + your own sales log.

Run locally:   python delivery/gumroad_webhook.py
Settings come from delivery/.env (copy .env.example and fill it in).
"""
import csv
import json
import os
import smtplib
from datetime import datetime, timezone
from email.message import EmailMessage
from pathlib import Path

from flask import Flask, abort, request

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
LOG = HERE / "sales_log.csv"
PRODUCT = json.loads((ROOT / "config" / "product.json").read_text())["product"]


def load_env():
    """Tiny .env reader: KEY=value lines become environment variables."""
    env = HERE / ".env"
    if env.exists():
        for line in env.read_text().splitlines():
            if "=" in line and not line.strip().startswith("#"):
                k, v = line.split("=", 1)
                os.environ.setdefault(k.strip(), v.strip())


load_env()
app = Flask(__name__)


def log_sale(sale):
    new_file = not LOG.exists()
    with LOG.open("a", newline="") as f:
        w = csv.writer(f)
        if new_file:
            w.writerow(["logged_at_utc", "sale_id", "product", "price_cents", "email", "test"])
        w.writerow([datetime.now(timezone.utc).isoformat(timespec="seconds"),
                    sale.get("sale_id"), sale.get("product_name"), sale.get("price"),
                    sale.get("email"), sale.get("test", "false")])


def welcome_email(to_addr, name):
    msg = EmailMessage()
    msg["Subject"] = f'Your {PRODUCT["name"]} is ready 🎉'
    msg["From"] = os.environ["SMTP_FROM"]
    msg["To"] = to_addr
    first = (name or "there").split()[0]
    msg.set_content(
        f"Hi {first},\n\n"
        f'Thanks for grabbing {PRODUCT["name"]}!\n\n'
        f"Google Sheets version (makes your own private copy):\n"
        f'{PRODUCT["sheets_copy_link"]}\n\n'
        "The Excel files are in your Gumroad receipt.\n\n"
        "Start on the 'Start Here' tab. It takes about 5 minutes to set up.\n"
        "Questions? Just reply to this email.\n")
    return msg


def send(msg):
    if os.getenv("DRY_RUN", "true").lower() == "true":
        print("DRY_RUN on, not sending. Would send:\n", msg)
        return
    with smtplib.SMTP_SSL(os.environ.get("SMTP_HOST", "smtp.gmail.com"), 465) as s:
        s.login(os.environ["SMTP_USER"], os.environ["SMTP_APP_PASSWORD"])
        s.send_message(msg)


@app.post("/gumroad")
def gumroad_ping():
    if request.args.get("key") != os.environ.get("WEBHOOK_SECRET"):
        abort(403)                                   # wrong or missing secret
    sale = request.form.to_dict()
    log_sale(sale)
    if sale.get("email"):
        send(welcome_email(sale["email"], sale.get("full_name")))
    return "ok", 200


@app.get("/")
def health():
    return "Paycheck Planner webhook is running", 200


if __name__ == "__main__":
    app.run(port=int(os.environ.get("PORT", 5000)))
