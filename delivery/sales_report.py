"""
sales_report.py — pulls your sales from the Gumroad API and prints a summary.

Run:  python delivery/sales_report.py
Needs GUMROAD_ACCESS_TOKEN in delivery/.env
(Gumroad: Settings → Advanced → Applications → create app → "Generate access token").
"""
import json
import os
import urllib.parse
import urllib.request
from collections import Counter, defaultdict
from pathlib import Path

HERE = Path(__file__).resolve().parent
API = "https://api.gumroad.com/v2/sales"


def load_env():
    env = HERE / ".env"
    if env.exists():
        for line in env.read_text().splitlines():
            if "=" in line and not line.strip().startswith("#"):
                k, v = line.split("=", 1)
                os.environ.setdefault(k.strip(), v.strip())


def fetch_all_sales(token):
    sales, page_key = [], None
    while True:
        params = {"access_token": token}
        if page_key:
            params["page_key"] = page_key
        with urllib.request.urlopen(f"{API}?{urllib.parse.urlencode(params)}", timeout=30) as r:
            data = json.load(r)
        sales += data.get("sales", [])
        page_key = data.get("next_page_key")
        if not page_key:
            return sales


def summarize(sales):
    by_month, by_product = defaultdict(int), Counter()
    for s in sales:
        month = s.get("created_at", "")[:7]
        by_month[month] += int(s.get("price", 0))
        by_product[s.get("product_name", "?")] += 1
    print(f"Total sales: {len(sales)}")
    for month in sorted(by_month):
        print(f"  {month}: ${by_month[month] / 100:,.2f}")
    print("By product:")
    for name, n in by_product.most_common():
        print(f"  {name}: {n}")


if __name__ == "__main__":
    load_env()
    token = os.environ.get("GUMROAD_ACCESS_TOKEN")
    if not token:
        raise SystemExit("Add GUMROAD_ACCESS_TOKEN to delivery/.env first.")
    summarize(fetch_all_sales(token))
