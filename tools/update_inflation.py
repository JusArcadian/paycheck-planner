"""
update_inflation.py — refreshes the "Auto" inflation rate in config/product.json
from the US Bureau of Labor Statistics (BLS), so newly built Excel files ship
with the latest number.

Run (then rebuild):
    python tools/update_inflation.py
    python tools/build_template.py

No API key needed. Run it once a month (BLS publishes mid-month).
"""
import json
import urllib.request
from pathlib import Path

BLS_URL = "https://api.bls.gov/publicAPI/v1/timeseries/data/CUUR0000SA0"
CONFIG_PATH = Path(__file__).resolve().parent.parent / "config" / "product.json"


def yoy_rate(series_data):
    """Year-over-year inflation from BLS data (newest month first)."""
    latest = series_data[0]
    year_ago = next(d for d in series_data
                    if d["period"] == latest["period"]
                    and int(d["year"]) == int(latest["year"]) - 1)
    rate = float(latest["value"]) / float(year_ago["value"]) - 1
    return round(rate, 3), f'{latest["periodName"][:3]} {latest["year"]}'


def main():
    with urllib.request.urlopen(BLS_URL, timeout=20) as resp:
        payload = json.load(resp)
    if payload.get("status") != "REQUEST_SUCCEEDED":
        raise SystemExit(f"BLS API error: {payload.get('message')}")
    rate, label = yoy_rate(payload["Results"]["series"][0]["data"])

    config = json.loads(CONFIG_PATH.read_text())
    old = config["inflation"]["auto_rate"]
    config["inflation"] = {"auto_rate": rate, "as_of": f"{label} (US CPI-U, BLS)"}
    CONFIG_PATH.write_text(json.dumps(config, indent=2) + "\n")
    print(f"Inflation updated: {old:.1%} -> {rate:.1%} ({label}). Now run build_template.py.")


if __name__ == "__main__":
    main()
