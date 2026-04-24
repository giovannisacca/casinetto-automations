# casinetto-automations

Automation toolkit for Casinetto e-commerce operations (Dubai/UAE, Italian specialty food B2C).

## Purpose

Convert recurring manual workflows from the Head of E-commerce role into committed, reusable code. Built iteratively via Claude Code.

## Modules (planned and in-progress)

- **analytics/** — Cube.dev staging queries (WBR, nationality YoY, Oro tier refresh, variable weight audit)
- **scrapers/** — competitor catalog scrapers (Euromercato, Eataly Arabia, Longino & Cardenal, Maison Duffour)
- **reports/** — Power BI refresh helpers, Overstock dashboard
- **utils/** — SAP B1 + Shopify data reconciliation

## Environment

- Windows local dev
- Python primary language
- Cube.dev staging MCP for analytics
- Ahrefs, Slack, Confluence APIs for orchestration

## Usage

### WBR data pull (`analytics/wbr_pull.py`)

Fetches 14 Weekly Business Review metrics from Cube.dev staging for a given
week-ending date and the same week prior year.

```bash
# Install dependencies
pip install requests python-dotenv tabulate

# Copy and fill in credentials
cp .env.example .env

# Run for last Sunday (default)
python analytics/wbr_pull.py

# Run for a specific week-ending date
python analytics/wbr_pull.py --week-ending 2026-04-20
```

**Output:**
- Formatted table printed to stdout with columns: Metric | This Week | LY Week | YoY %
- CSV saved to `analytics/output/wbr_<YYYY-MM-DD>.csv` (gitignored)

**Metrics covered:**
Revenue (B2C), Gross Profit, GP%, Orders, AOV, New Customers,
Returning Customers, Conversion Rate, Marketing Spend, CAC,
Ops Cost, COGS, Inventory Value, Stockout Count.

> **Note:** Cube measure names in `analytics/wbr_pull.py` under the `MEASURES`
> dict may need updating to match the actual staging schema. Run
> `GET /cubejs-api/v1/meta` against the endpoint to confirm names.

---

## Owner

Giovanni Saccà, Head of E-commerce, Casinetto.
