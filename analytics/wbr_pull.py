# TODO (2026-04-24): verify MEASURES dict against Cube.dev schema when endpoint
# is accessible. Run GET /cubejs-api/v1/meta or the list_cubes MCP tool
# (server: https://casinetto.staging.cibos.dev/api/mcp) to confirm all
# cube and measure names below before using in production.

"""
wbr_pull.py — Weekly Business Review data pull from Cube.dev staging.

Fetches 14 WBR metrics for a given week-ending date and the same week
prior year, then prints a comparison table and saves a CSV.

Usage:
    python analytics/wbr_pull.py                         # defaults to last Sunday
    python analytics/wbr_pull.py --week-ending 2026-04-20

Output:
    stdout  — formatted table
    file    — analytics/output/wbr_<YYYY-MM-DD>.csv

Requires:
    pip install requests python-dotenv tabulate
    .env with CUBE_API_URL and CUBE_API_TOKEN
"""

import argparse
import csv
import os
import sys
from datetime import date, timedelta
from pathlib import Path

import requests
from dotenv import load_dotenv

# ---------------------------------------------------------------------------
# Cube.dev measure / dimension names
# Update these keys if the staging schema changes.
# Confirm names via: GET /cubejs-api/v1/meta  (or the list_cubes MCP tool)
# ---------------------------------------------------------------------------

MEASURES = {
    # Measure name in Cube.dev          : human label
    "sales_overview.revenue"            : "Revenue (B2C)",
    "sales_overview.gross_profit"       : "Gross Profit",
    "sales_overview.order_count"        : "Orders",
    "customers.new_count"               : "New Customers",
    "customers.returning_count"         : "Returning Customers",
    "customers.conversion_rate"         : "Conversion Rate",
    "marketing.total_spend"             : "Marketing Spend",
    "operations.total_cost"             : "Ops Cost",
    "sales_overview.cogs"               : "COGS",
    "inventory.total_value"             : "Inventory Value",
    "inventory.stockout_count"          : "Stockout Count",
}

# Derived metrics computed after querying (not fetched directly from Cube)
DERIVED = ["GP%", "AOV", "CAC"]

# Time dimension used across all cubes
TIME_DIM = {
    "sales_overview" : "sales_overview.order_date",
    "customers"      : "customers.order_date",
    "marketing"      : "marketing.date",
    "operations"     : "operations.date",
    "inventory"      : "inventory.snapshot_date",
}

# B2C revenue filter
B2C_FILTER = {
    "member"   : "sales_overview.order_source",
    "operator" : "equals",
    "values"   : ["B2C-Website"],
}

# ---------------------------------------------------------------------------
# Date helpers
# ---------------------------------------------------------------------------

def last_sunday(ref: date = None) -> date:
    """Return the most recent Sunday on or before ref (default: today)."""
    d = ref or date.today()
    return d - timedelta(days=(d.weekday() + 1) % 7)


def week_range(week_ending: date):
    """Return (Monday, Sunday) for the week ending on week_ending."""
    return week_ending - timedelta(days=6), week_ending


def ly_week_range(week_ending: date):
    """Same week-ending date, prior year (52 weeks back = 364 days)."""
    ly_end = week_ending - timedelta(weeks=52)
    return week_range(ly_end)


# ---------------------------------------------------------------------------
# Cube.dev query helpers
# ---------------------------------------------------------------------------

def cube_query(session: requests.Session, api_url: str, measures: list,
               date_start: str, date_end: str, cube_name: str,
               extra_filters: list = None) -> dict:
    """
    Execute a single Cube.dev /load query and return the first data row
    as a flat dict {measure_name: value}.
    """
    time_dim_key = TIME_DIM.get(cube_name)
    if not time_dim_key:
        raise ValueError(f"No TIME_DIM entry for cube '{cube_name}'")

    payload = {
        "query": {
            "measures"       : measures,
            "timeDimensions" : [{
                "dimension" : time_dim_key,
                "dateRange" : [date_start, date_end],
            }],
            "filters"        : extra_filters or [],
        }
    }

    resp = session.post(f"{api_url}/cubejs-api/v1/load", json=payload, timeout=30)
    resp.raise_for_status()
    data = resp.json().get("data", [])
    return data[0] if data else {}


def fetch_measure(session: requests.Session, api_url: str, measure: str,
                  date_start: str, date_end: str) -> float:
    """Fetch a single measure value for a date range. Returns 0.0 on miss."""
    cube_name = measure.split(".")[0]
    extra = [B2C_FILTER] if cube_name == "sales_overview" else []
    row = cube_query(session, api_url, [measure], date_start, date_end,
                     cube_name, extra_filters=extra)
    val = row.get(measure, 0)
    try:
        return float(val) if val is not None else 0.0
    except (TypeError, ValueError):
        return 0.0


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def parse_args():
    p = argparse.ArgumentParser(description="Pull WBR metrics from Cube.dev staging.")
    p.add_argument(
        "--week-ending",
        metavar="YYYY-MM-DD",
        default=None,
        help="Week-ending Sunday date (default: last Sunday)",
    )
    return p.parse_args()


def fmt_currency(v: float) -> str:
    return f"AED {v:,.0f}" if v != 0 else "—"


def fmt_pct(v: float) -> str:
    return f"{v:.1f}%" if v != 0 else "—"


def fmt_count(v: float) -> str:
    return f"{v:,.0f}" if v != 0 else "—"


def fmt_yoy(this: float, ly: float) -> str:
    if ly == 0:
        return "N/A"
    pct = (this - ly) / abs(ly) * 100
    sign = "+" if pct >= 0 else ""
    return f"{sign}{pct:.1f}%"


def main():
    load_dotenv()

    api_url   = os.getenv("CUBE_API_URL", "").rstrip("/")
    api_token = os.getenv("CUBE_API_TOKEN", "")

    if not api_url or not api_token:
        sys.exit("ERROR: CUBE_API_URL and CUBE_API_TOKEN must be set in .env")

    args = parse_args()
    if args.week_ending:
        try:
            week_end = date.fromisoformat(args.week_ending)
        except ValueError:
            sys.exit(f"ERROR: Invalid date '{args.week_ending}'. Use YYYY-MM-DD.")
    else:
        week_end = last_sunday()

    tw_start, tw_end = week_range(week_end)
    ly_start, ly_end = ly_week_range(week_end)

    print(f"\nWBR pull  ·  This week: {tw_start} → {tw_end}  |  LY week: {ly_start} → {ly_end}\n")

    session = requests.Session()
    session.headers.update({
        "Authorization" : api_token,
        "Content-Type"  : "application/json",
    })

    def fetch(measure, start, end):
        try:
            return fetch_measure(session, api_url, measure, str(start), str(end))
        except requests.HTTPError as e:
            print(f"  WARNING: HTTP {e.response.status_code} fetching {measure}: {e}", file=sys.stderr)
            return 0.0
        except Exception as e:
            print(f"  WARNING: Error fetching {measure}: {e}", file=sys.stderr)
            return 0.0

    # ── Fetch all base measures ──────────────────────────────────────────────
    tw_vals, ly_vals = {}, {}
    for measure in MEASURES:
        tw_vals[measure] = fetch(measure, tw_start, tw_end)
        ly_vals[measure] = fetch(measure, ly_start, ly_end)

    # ── Derive computed metrics ──────────────────────────────────────────────
    tw_revenue  = tw_vals["sales_overview.revenue"]
    ly_revenue  = ly_vals["sales_overview.revenue"]
    tw_gp       = tw_vals["sales_overview.gross_profit"]
    ly_gp       = ly_vals["sales_overview.gross_profit"]
    tw_orders   = tw_vals["sales_overview.order_count"]
    ly_orders   = ly_vals["sales_overview.order_count"]
    tw_new_cust = tw_vals["customers.new_count"]
    ly_new_cust = ly_vals["customers.new_count"]
    tw_mktg     = tw_vals["marketing.total_spend"]
    ly_mktg     = ly_vals["marketing.total_spend"]

    tw_gp_pct  = (tw_gp / tw_revenue * 100) if tw_revenue else 0.0
    ly_gp_pct  = (ly_gp / ly_revenue * 100) if ly_revenue else 0.0
    tw_aov     = (tw_revenue / tw_orders)    if tw_orders  else 0.0
    ly_aov     = (ly_revenue / ly_orders)    if ly_orders  else 0.0
    tw_cac     = (tw_mktg / tw_new_cust)     if tw_new_cust else 0.0
    ly_cac     = (ly_mktg / ly_new_cust)     if ly_new_cust else 0.0

    # ── Build rows for display ───────────────────────────────────────────────
    rows = [
        ("Revenue (B2C)",     fmt_currency(tw_revenue),
                              fmt_currency(ly_revenue),
                              fmt_yoy(tw_revenue, ly_revenue)),
        ("Gross Profit",      fmt_currency(tw_gp),
                              fmt_currency(ly_gp),
                              fmt_yoy(tw_gp, ly_gp)),
        ("GP%",               fmt_pct(tw_gp_pct),
                              fmt_pct(ly_gp_pct),
                              fmt_yoy(tw_gp_pct, ly_gp_pct)),
        ("Orders",            fmt_count(tw_orders),
                              fmt_count(ly_orders),
                              fmt_yoy(tw_orders, ly_orders)),
        ("AOV",               fmt_currency(tw_aov),
                              fmt_currency(ly_aov),
                              fmt_yoy(tw_aov, ly_aov)),
        ("New Customers",     fmt_count(tw_new_cust),
                              fmt_count(ly_new_cust),
                              fmt_yoy(tw_new_cust, ly_new_cust)),
        ("Returning Customers", fmt_count(tw_vals["customers.returning_count"]),
                              fmt_count(ly_vals["customers.returning_count"]),
                              fmt_yoy(tw_vals["customers.returning_count"],
                                      ly_vals["customers.returning_count"])),
        ("Conversion Rate",   fmt_pct(tw_vals["customers.conversion_rate"]),
                              fmt_pct(ly_vals["customers.conversion_rate"]),
                              fmt_yoy(tw_vals["customers.conversion_rate"],
                                      ly_vals["customers.conversion_rate"])),
        ("Marketing Spend",   fmt_currency(tw_mktg),
                              fmt_currency(ly_mktg),
                              fmt_yoy(tw_mktg, ly_mktg)),
        ("CAC",               fmt_currency(tw_cac),
                              fmt_currency(ly_cac),
                              fmt_yoy(tw_cac, ly_cac)),
        ("Ops Cost",          fmt_currency(tw_vals["operations.total_cost"]),
                              fmt_currency(ly_vals["operations.total_cost"]),
                              fmt_yoy(tw_vals["operations.total_cost"],
                                      ly_vals["operations.total_cost"])),
        ("COGS",              fmt_currency(tw_vals["sales_overview.cogs"]),
                              fmt_currency(ly_vals["sales_overview.cogs"]),
                              fmt_yoy(tw_vals["sales_overview.cogs"],
                                      ly_vals["sales_overview.cogs"])),
        ("Inventory Value",   fmt_currency(tw_vals["inventory.total_value"]),
                              fmt_currency(ly_vals["inventory.total_value"]),
                              fmt_yoy(tw_vals["inventory.total_value"],
                                      ly_vals["inventory.total_value"])),
        ("Stockout Count",    fmt_count(tw_vals["inventory.stockout_count"]),
                              fmt_count(ly_vals["inventory.stockout_count"]),
                              fmt_yoy(tw_vals["inventory.stockout_count"],
                                      ly_vals["inventory.stockout_count"])),
    ]

    # ── Print table ──────────────────────────────────────────────────────────
    try:
        from tabulate import tabulate
        header = ["Metric", f"This Week\n({tw_start})", f"LY Week\n({ly_start})", "YoY %"]
        print(tabulate(rows, headers=header, tablefmt="rounded_outline"))
    except ImportError:
        # Fallback: plain text table without tabulate
        col_w = [24, 18, 18, 10]
        def _row(cells):
            return "  ".join(str(c).ljust(w) for c, w in zip(cells, col_w))
        print(_row(["Metric", f"This Week ({tw_start})", f"LY Week ({ly_start})", "YoY %"]))
        print("  ".join("-" * w for w in col_w))
        for r in rows:
            print(_row(r))

    # ── Save CSV ─────────────────────────────────────────────────────────────
    output_dir = Path(__file__).parent / "output"
    output_dir.mkdir(exist_ok=True)
    csv_path = output_dir / f"wbr_{week_end}.csv"

    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["metric", "this_week", "ly_week", "yoy_pct"])
        writer.writerow(["week_ending", str(tw_end), str(ly_end), ""])
        writer.writerow(["week_start",  str(tw_start), str(ly_start), ""])
        writer.writerow([])
        writer.writerows(rows)

    print(f"\nSaved: {csv_path}\n")


if __name__ == "__main__":
    main()
