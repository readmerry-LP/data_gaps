# IOR Data Gaps Trend Dashboard

An interactive HTML dashboard that consolidates all weekly IOR Data Gaps Master reports into a single view — tracking gap trends, warehouse performance, recurring dockets, client codes, and analyst observations over time.

Built and maintained by the Logistics Plus Data Integrity team.

---

## What It Does

Each Monday, the BI team generates a Master `.xlsx` report listing every shipment with missing or incorrect data fields across three report types: **Orders**, **Receipts**, and **WW ETAs**. This dashboard reads all archived Master reports and produces a self-contained HTML file with:

- Week-over-week gap volume trends (by report type)
- New vs. carried-over docket tracking and carryover rate
- Open gap age breakdown (new → chronic)
- Per-warehouse gap type breakdown and weekly trend
- Meeting Spotlight for one-on-one warehouse ops meetings (with auto-generated action items)
- Recurring docket tracker (dockets open 3+ consecutive weeks)
- Org code heat map (top 25 clients × warehouses)
- Auto-generated analyst Observations summary
- BI rules changelog with trend chart annotation

---

## Files

| File | Description |
|---|---|
| `generate_dashboard.py` | Main script — reads all Master reports, computes all metrics, writes the dashboard HTML |
| `template_before.html` | HTML head, styles, nav, and tab panel structure |
| `template_after.html` | All JavaScript — chart rendering, filters, interactivity |
| `exceptions_log.csv` | Changelog of BI reporting logic changes (see below) |
| `DataGaps_Dashboard.html` | **Generated output** — open this in a browser to view the dashboard |

---

## Requirements

```
Python 3.8+
openpyxl
```

Install dependency:

```bash
pip install openpyxl
```

No other dependencies. Chart.js is fetched from CDN at generation time and embedded inline, so the output HTML is fully self-contained (works in OneDrive browser preview and offline).

---

## Setup

1. Clone this repo into your local `DataGaps Dashboard` folder.
2. Open `generate_dashboard.py` and update `ARCHIVE_FOLDER` at the top to point to your local archive of Master reports:

```python
ARCHIVE_FOLDER = r'C:\Your\Path\To\Archived\Master Reports - 2026'
```

3. Optionally update `MIN_DATE` to control how far back the dashboard looks:

```python
MIN_DATE = '2026-05-01'
```

---

## Weekly Workflow

1. BI team runs the weekly report — Master `.xlsx` is saved to the archive folder.
2. Run the script:

```bash
python generate_dashboard.py
```

3. Open or share `DataGaps_Dashboard.html`.

The script reads every `*MASTER*.xlsx` file in the archive folder dated on or after `MIN_DATE`, computes all metrics, and regenerates the dashboard. It takes about 30–60 seconds depending on archive size.

---

## Reporting Rules Changes (`exceptions_log.csv`)

When the BI team updates their flagging logic (e.g., suppressing a gap type for a specific client or carrier), add a row to `exceptions_log.csv`. The dashboard will automatically annotate the trend chart with a dashed vertical line at the effective date and display a plain-language note explaining what changed and why.

**Columns:**

| Column | Description |
|---|---|
| `Date Reported` | When the warehouse flagged the issue |
| `Effective Date` | First Monday report the BI change appears in (format: `YYYY-MM-DD`) |
| `Warehouse` | Warehouse code (e.g., `VEW`), or blank for network-wide |
| `Org Code` | Client org code affected, or blank |
| `Gap Type` | Gap value being changed (e.g., `Missing POD`) |
| `Carrier Filter` | Carrier restriction if applicable (e.g., `DHL`), or blank |
| `Action` | `BI Rule Change` |
| `Reason` | Plain-language explanation |
| `Reported By` | Who flagged it |
| `Notes` | Any additional context |

**Example:** VEW informed us on 6/11 that two flags were invalid starting with the 6/15 report:
- `CISDENSEA` — Cisco doesn't ship on its own account, so `Missing POD` should never fire for them.
- `WORKWINWK` — `Missing POD` should not fire for DHL orders specifically.

The BI team updated their script. Those rows in `exceptions_log.csv` explain why the 6/15 gap count is lower than 6/8 — it's partially a rules correction, not purely warehouse improvement.

---

## Dashboard Tabs

| Tab | What It Shows |
|---|---|
| **Guide** | Presenter cheat sheet — brief talking points for each KPI and how to use them in warehouse meetings |
| **Overview** | Weekly KPIs, trend chart (filterable by warehouse), new vs. carried-over, gap age breakdown |
| **Warehouses** | Per-warehouse gap type chart, weekly trend, and meeting spotlight with action items |
| **Gap Types** | Network-wide gap column frequency by report type |
| **Recurrence** | All dockets open 3+ consecutive weeks |
| **Org Codes** | Client × warehouse heat map (top 25 org codes) |
| **Observations** | Auto-generated analyst summary — regenerated every run |

---

## Schema Drift

Master report column structures have changed over time. The script joins on **column names, not positions**, so new columns are picked up automatically and older reports with different schemas are still read correctly. Missing columns default to zero for earlier weeks.

---

## Notes

- The `DataGaps_Dashboard.html` output file can be committed to the repo or shared directly via OneDrive — it has no external dependencies once generated.
- The `exceptions_log.csv` should be committed alongside the templates so the changelog is version-controlled.
- To add a new year's archive, update `ARCHIVE_FOLDER` and `MIN_DATE` in `generate_dashboard.py`.
