import requests
import csv
import os
from datetime import datetime

SHOP = "casinetto"
ACCESS_TOKEN = os.environ["SHOPIFY_TOKEN"]
API_VERSION = "2026-01"

SINCE = "2026-03-01"
UNTIL = "2026-03-12"
COMPARE_SINCE = "2025-03-03"
COMPARE_UNTIL = "2025-03-13"

SHOPIFY_QL = (
    "FROM sales\n"
    "  SHOW customers, orders, average_order_value, net_items_sold,\n"
    "    quantity_ordered_per_order\n"
    "  TIMESERIES day WITH TOTALS\n"
    f"  SINCE {SINCE} UNTIL {UNTIL}\n"
    f"  COMPARE TO {COMPARE_SINCE} UNTIL {COMPARE_UNTIL}\n"
    "  ORDER BY day ASC"
)

GRAPHQL_QUERY = """
query analyticsReport($query: String!) {
  analyticsReport(queryString: $query) {
    parseErrors { code message }
    tableData {
      unformattedData { rowData columnNames }
    }
  }
}
"""

url = f"https://{SHOP}.myshopify.com/admin/api/{API_VERSION}/graphql.json"
headers = {
    "Content-Type": "application/json",
    "X-Shopify-Access-Token": ACCESS_TOKEN,
}

response = requests.post(url, headers=headers, json={
    "query": GRAPHQL_QUERY,
    "variables": {"query": SHOPIFY_QL}
})

data = response.json()
errors = data["data"]["analyticsReport"].get("parseErrors", [])
if errors:
    print("ShopifyQL errors:", errors)
    exit(1)

report = data["data"]["analyticsReport"]["tableData"]["unformattedData"]
columns = report["columnNames"]
rows = report["rowData"]

today = datetime.today().strftime("%Y-%m-%d")
filename = f"war_room_daily_{today}.csv"

with open(filename, "w", newline="") as f:
    writer = csv.writer(f)
    writer.writerow(columns)
    for row in rows:
        writer.writerow(row)

print(f"Saved: {filename}")
