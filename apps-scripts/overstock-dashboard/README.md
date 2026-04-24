# Overstock Dashboard — Apps Script

Populates the `Ecom_Short_Expiry_Overstock_daily` tab of the **war-room-report** Google Sheet with a
live clearance-progress dashboard for Christmas/Easter overstock SKUs, rebuilt from scratch on
every run.

---

## What it does

1. **Locates today's source tab** in the *Active SKUs – Short Expiry, overstock, re-order* sheet.
   The source file has one tab per day named `DD.MM.YY` (e.g. `23.04.26`).
   If today's tab does not exist yet, it falls back to the most recent dated tab and prepends a
   staleness warning to the dashboard.

2. **Filters** every data row where **Order Interval = "Christmas/Easter"** (column `DK`).

3. **Computes all metrics** using the snapshot-delta approach:
   - *Daily cleared AED* = yesterday's sum of col `DI` minus today's sum of col `DI`
   - *Daily cleared Units* = same using col `DH`
   - *Cumulative cleared since 14 Apr 2026* = baseline (`14.04.26` tab) minus today
   - *Required daily burn* = remaining overstock AED ÷ days left to 31 May 2026
   - *Trailing 7-day avg burn* = average of last 7 tab-to-tab deltas
   - *Week-over-week* totals for last complete week and current partial week
   - *Top-risk SKU* = Christmas/Easter SKU with highest remaining overstock AED

4. **Clears** the target tab completely (content + formats + merges), then **writes native
   Google Sheets cells** — no images, no screenshots.

5. **Logs** a Stackdriver summary line on every run (visible in Apps Script → Executions).

---

## Source and target file IDs

| Role | Name | File ID |
|------|------|---------|
| Source | Active SKUs – Short Expiry, overstock, re-order | `12TOeabCNl2YUImUCoyDYOmpi-_CCCFos2n6gukLNEm4` |
| Target | war-room-report | `1Z3Sn4_zPV_VSYSPjZgMEoyYBnzrHpEI35xJB8XAQYu0` |
| Target tab | — | `Ecom_Short_Expiry_Overstock_daily` |

---

## Step-by-step: paste into Google Apps Script editor

1. Open the **war-room-report** Google Sheet.
2. Click **Extensions → Apps Script**.
3. If a default `Code.gs` file is shown, delete all its contents.
4. Paste the entire contents of `overstock-dashboard.gs` into the editor.
5. Click **Save** (floppy-disk icon or `Ctrl+S`).
6. Click **Run → runDashboard** to test once manually.
   - The first run will ask for OAuth permissions — click *Review permissions* and approve.
7. Verify the `Ecom_Short_Expiry_Overstock_daily` tab has been populated.

> **Important:** The target tab `Ecom_Short_Expiry_Overstock_daily` must already exist in the
> war-room-report file before you run the script. The script will refuse to run if the tab is
> missing (safety guard against writing to the wrong place).

---

## How to set up the daily time trigger

1. In the Apps Script editor, click the **clock icon** (Triggers) in the left sidebar.
2. Click **+ Add Trigger** (bottom right).
3. Fill in:
   | Field | Value |
   |-------|-------|
   | Function to run | `runDashboard` |
   | Deployment | Head |
   | Event source | Time-driven |
   | Type of time-based trigger | Day timer |
   | Time of day | Choose a time after the source file is refreshed (e.g. 07:00–08:00) |
4. Click **Save**.

The script will now run automatically each day and overwrite the dashboard tab with fresh data.

---

## Staleness handling

- If today's tab (`DD.MM.YY`) is missing in the source file, the script falls back to the most
  recent dated tab and adds a **red warning row** at the top of the dashboard.
- If the most recent tab is more than **2 days** behind today, a staleness banner is shown.
- If **zero Christmas/Easter SKUs** are found, a red error row is shown — this likely means the
  source sheet structure has changed; check the `Order Interval` column.

---

## Modifying the layout

The dashboard is written in `_writeDashboard()` using a simple `row(cells, opts)` builder.
Each call adds one row. To add, remove, or reorder sections:

- **Add a row:** call `row(['Label', value, ...], { labelBold: true })` in the right place.
- **Change a section header colour:** update the `bg` option on the relevant `row(...)` call.
- **Add a new metric:** compute it in `runDashboard()`, pass it into `_writeDashboard()` via the
  `d` object, then add a `row(...)` call in `_writeDashboard()`.
- **Change the deadline or baseline date:** update `CFG.DEADLINE_DATE` and `CFG.BASELINE_DATE`.
- **Change the filter value:** update `CFG.FILTER_VALUE` (currently `'Christmas/Easter'`).

Column positions are discovered dynamically from row 1 of the source tab — no hard-coded
column letters. The script is resilient to new columns being inserted to the left.
