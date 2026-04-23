# iiq-assets-to-sheets

Google Apps Script project that extracts Incident IQ (iiQ) asset data into Google Sheets for device reporting, replacement cycle planning, and dashboard consumption. Designed for school district IT teams.

## Quick Start

### 1. Create and Set Up Your Spreadsheet

**Option A: Copy the Template (Fastest)**

[**Make a copy of the Google Sheets template**](https://docs.google.com/spreadsheets/d/1ZDHUsL0Hqowe9xqS42fGHLtP2e0WYV6pM1EF_Tj9gXM/edit?usp=sharing)

This template includes all sheets, formulas, and scripts pre-configured. Skip to step 2.

**Option B: Build from Scratch**

1. Create a new Google Spreadsheet
2. Go to **Extensions > Apps Script**
3. Copy all `.gs` files from the `scripts/` directory
4. Save and refresh the spreadsheet
5. Run **iiQ Assets > Setup > Setup Spreadsheet**

### 2. Configure API Access

In the **Config** sheet, enter your Incident IQ credentials:

| Setting | Value | Where to Find It |
|---------|-------|------------------|
| `API_BASE_URL` | `https://yourdistrict.incidentiq.com` | Your iiQ URL (the `/api` is added automatically) |
| `BEARER_TOKEN` | Your API token | iiQ Admin > Integrations > API |
| `SITE_ID` | Your site UUID (optional) | Only needed for multi-site districts |

Then run **iiQ Assets > Setup > Verify Configuration** to confirm everything connects.

### 3. Load Your Data

1. Run **iiQ Assets > Asset Data > Load / Resume Assets**
   - Reference data (locations, status types) loads automatically on first run
   - Script runs ~5.5 minutes per batch, then pauses (Apps Script time limit)

> **Tip for large districts:** Instead of manually running "Load / Resume" repeatedly, set up triggers in step 4 and let automation finish the load.

### 4. Set Up Automated Refresh (Recommended)

Run **iiQ Assets > Setup > Setup Automated Triggers** to create all triggers automatically.

| Trigger | Schedule | What It Does |
|---------|----------|--------------|
| `triggerDataContinue` | Every 10 min | Continues any in-progress loading (no-op once complete) |
| `triggerDailyRefresh` | Daily 3 AM | Incremental refresh — only fetches changed assets |
| `triggerWeeklyFullRefresh` | Sunday 2 AM | Full reload including reference data |

> **About `triggerDataContinue`:** This is your "keep things moving" trigger. It automatically continues initial data loading, and once complete, applies formulas. You can leave it enabled permanently — when everything is loaded, it does nothing.

### Replacement Planning Config (Optional)

Two Config keys control the ReplacementPlanning and ReplacementForecast sheets:

- `REPLACEMENT_AGE_YEARS` — devices older than this are flagged for replacement (default: 4)
- `NEXT_SCHOOL_YEAR_START` — target date for planning (default: 2026-07-01, format YYYY-MM-DD)

## What You Get

**Data loaded automatically from your iiQ instance:**
- Complete asset inventory (28 columns) — identity, device model, location, owner, status, purchase info, storage, tickets, and more
- Location directory and asset status types
- Student enrollment and device coverage per school (optional)

**Calculated columns on every asset (ARRAYFORMULA — instant even at 300K+ rows):**
- Device age in days and years (falls back to CreatedDate if PurchasedDate is empty)
- Warranty status: Active / Expiring (< 90 days) / Expired / None

### Default Analytics (created by Setup)

| Sheet | What It Answers |
|-------|-----------------|
| ★ FleetSummary | Top-line KPIs: total assets, value, age, warranty, tickets, assignment |
| ★ LocationSummary | How many assets per school? How old? Warranty status? |
| ★ ModelBreakdown | Which device models do we have? How many active vs retired? |
| ★ AgingAnalysis | What's our fleet age distribution? When is the replacement cliff? |
| ★ BudgetPlanning | What's the replacement cost per location based on warranty/age? |
| ★ ServiceImpact | Which models generate the most support tickets? |
| ★ AssignmentOverview | How many devices are assigned vs idle per location? |
| ★ StatusOverview | What's the breakdown of active/retired/in-storage? |

### Additional Analytics (add via menu)

Use **iiQ Assets > Analytics Sheets** to add any of these 16 optional sheets:

| Category | Available Sheets |
|----------|------------------|
| Fleet Operations | Device Readiness, Spare Assets, Lost/Stolen Rate, Model Fragmentation, Unassigned Inventory |
| Service & Reliability | Break Rate, High Ticket Locations |
| Budget & Planning | Replacement Planning, Replacement Forecast, Warranty Timeline, Device Lifecycle |
| Fleet Composition | Location Model Breakdown, Location Model Filtered, Category Breakdown, Manufacturer Summary |
| People | Individual Lookup (dropdown-driven asset assignment history per user — calls the user activities API live on selection; works for districts that assign devices directly without formal checkouts) |

> **Flexible & Customizable:** Districts can delete any analytics sheet and recreate it later via the menu. Default sheets (marked with ★) can also be recreated if accidentally deleted.

## Menu Structure

```
iiQ Assets
├── Setup
│   ├── Setup Spreadsheet
│   ├── Verify Configuration
│   ├── Setup Automated Triggers
│   ├── View Trigger Status
│   └── Remove Automated Triggers
├── Load Reference Data
│   ├── Refresh Locations
│   ├── Refresh Status Types
│   ├── Refresh Location Enrollment
│   ├── View Available Roles
│   └── Refresh All Reference Data
├── Asset Data
│   ├── Load / Resume Assets
│   ├── Refresh Changed Assets
│   ├── Apply Formulas
│   ├── Show Status
│   ├── Remove Duplicates
│   ├── Clear Data + Reset Progress
│   └── Full Reload
├── Analytics Sheets
│   ├── Fleet Operations
│   │   ├── ★ Assignment Overview
│   │   ├── ★ Status Overview
│   │   ├── Device Readiness
│   │   ├── Spare Assets
│   │   ├── Lost/Stolen Rate
│   │   ├── Model Fragmentation
│   │   ├── Unassigned Inventory
│   │   └── Regenerate Fleet Operations
│   ├── Service & Reliability
│   │   ├── ★ Service Impact
│   │   ├── Break Rate
│   │   ├── High Ticket Locations
│   │   └── Regenerate Service & Reliability
│   ├── Budget & Planning
│   │   ├── ★ Budget Planning
│   │   ├── ★ Aging Analysis
│   │   ├── Replacement Planning
│   │   ├── Replacement Forecast
│   │   ├── Warranty Timeline
│   │   ├── Device Lifecycle
│   │   └── Regenerate Budget & Planning
│   ├── Fleet Composition
│   │   ├── ★ Fleet Summary
│   │   ├── ★ Location Summary
│   │   ├── ★ Model Breakdown
│   │   ├── Location Model Breakdown
│   │   ├── Location Model Filtered
│   │   ├── Category Breakdown
│   │   ├── Manufacturer Summary
│   │   └── Regenerate Fleet Composition
│   ├── People
│   │   ├── Individual Lookup
│   │   └── Regenerate People
│   ├── Regenerate All Default (★)
│   └── Regenerate All Analytics
```

Default sheets (★) are created by **Setup Spreadsheet**. All sheets — default and optional — appear in their category submenu for individual regeneration or installation.

**When to regenerate:** Analytics sheets use live Google Sheets formulas, so they update automatically whenever your data refreshes. Regeneration is only needed after a code update changes a formula definition. Use the per-category regenerate options, "Regenerate All Default" to rebuild the 8 starred sheets, or "Regenerate All Analytics" to rebuild all installed analytics sheets.

## Config Sheet

| Key | Required | Description |
|-----|----------|-------------|
| `API_BASE_URL` | Yes | Your iiQ instance URL (e.g., `https://yourdistrict.incidentiq.com`) |
| `BEARER_TOKEN` | Yes | API token from Admin > Integrations > API |
| `SITE_ID` | No | Site UUID (only for multi-site instances) |
| `PAGE_SIZE` | No | Records per API call (default 100) |
| `THROTTLE_MS` | No | Delay between requests in ms (default 1000) |
| `ASSET_BATCH_SIZE` | No | Assets per page for bulk load (default 500) |
| `STUDENT_ROLE_ID` | No | Role UUID for student enrollment counts |
| `REPLACEMENT_AGE_YEARS` | No | Device age threshold for replacement planning (default 4) |
| `NEXT_SCHOOL_YEAR_START` | No | Target date for replacement planning (default 2026-07-01, format YYYY-MM-DD) |

Progress-tracking keys (`ASSET_LAST_PAGE`, `ASSET_TOTAL_PAGES`, `ASSET_COMPLETE`, `LAST_REFRESH_DATE`) are managed automatically.

## Automated Refresh

After initial setup, the automated triggers handle everything:

| Trigger | Schedule | What It Does |
|---------|----------|-------------|
| `triggerDataContinue` | Every 10 min | Continues interrupted initial loading (no-op once complete) |
| `triggerDailyRefresh` | Daily 3 AM | Incremental refresh — only fetches assets modified since last refresh |
| `triggerWeeklyFullRefresh` | Sunday 2 AM | Full data refresh including reference data |

The daily refresh uses iiQ's ModifiedDate filter to only pull changed records. If your district has 100,000 assets but only 200 changed today, only those 200 are fetched and updated in-place.

The weekly full refresh starts Sunday at 2 AM. If the reload does not finish in a single Apps Script run, the 10-minute continue trigger keeps resuming it until the dataset is fully rebuilt.

Deleted assets in iiQ are automatically excluded by the API — they are never downloaded. The weekly full reload catches edge cases like un-deleted assets reappearing.

## Connecting to Looker Studio / Power BI

Connect your BI tool directly to this Google Spreadsheet:
- **AssetData** sheet for detailed device-level data
- **Analytics sheets** for pre-aggregated summaries
- Formula columns (AgeYears, WarrantyStatus) are ready for filtering and grouping
- **LocationEnrollment** for student device coverage metrics

Data is refreshed daily at 3 AM. The weekly full refresh starts Sunday at 2 AM and may continue in 10-minute batches for larger districts.

## Troubleshooting

| Issue | Solution |
|-------|----------|
| "API configuration missing" | Fill in `API_BASE_URL` and `BEARER_TOKEN` in Config sheet |
| Loading stops partway | Run "Load / Resume Assets" again — it resumes from where it left off |
| Analytics show #REF errors | Run "Regenerate All Default" or the specific category regenerate |
| Need fresh data now | Run "Refresh Changed Assets" for on-demand update |
| Need complete reset | Remove automated triggers first, then run "Full Reload" |
| "STUDENT_ROLE_ID not configured" | Run "View Available Roles" to find the ID, add it to Config |

## Documentation

- [**Implementation Guide**](GUIDE.md) — Detailed setup, data loading, formulas, and how everything works
- [**CLAUDE.md**](CLAUDE.md) — Technical reference for developers

## License

MIT — Free to use and modify for your district.
