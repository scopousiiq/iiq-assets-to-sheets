# iiq-assets-to-sheets

Google Apps Script project that extracts Incident IQ (iiQ) asset data into Google Sheets for device reporting, replacement cycle planning, and dashboard consumption. Designed for school district IT teams.

## What You Get

**Data loaded automatically from your iiQ instance:**
- Complete asset inventory (28 columns) — identity, device model, location, owner, status, purchase info, storage, tickets, and more
- Location directory and asset status types
- Student enrollment and device coverage per school (optional)

**23 analytics sheets organized by category (formula-driven, auto-updating):**

Sheets marked with a star (★) are created automatically by Setup Spreadsheet. Optional sheets can be installed individually from the **iiQ Assets > Analytics Sheets** menu.

### Fleet Operations

| Sheet | What It Answers |
|-------|-----------------|
| ★ AssignmentOverview | How many devices are assigned vs idle per location? |
| ★ StatusOverview | What's the breakdown of active/retired/in-storage? |
| DeviceReadiness | What's actually deployable at each school right now? |
| SpareAssets | Do I have enough working spares at each school? |
| LostStolenRate | Which schools are losing devices? |
| ModelFragmentation | Which schools are a patchwork of device models? |
| UnassignedInventory | Where is idle inventory sitting? |

### Service & Reliability

| Sheet | What It Answers |
|-------|-----------------|
| ★ ServiceImpact | Which models generate the most support tickets? |
| BreakRate | Which individual devices and models have the most tickets? |
| HighTicketLocations | Which schools have the most device problems? |

### Budget & Planning

| Sheet | What It Answers |
|-------|-----------------|
| ★ BudgetPlanning | What's the replacement cost per location based on warranty/age? |
| ★ AgingAnalysis | What's our fleet age distribution? When is the replacement cliff? |
| ReplacementPlanning | What do I need to buy before next school year? |
| ReplacementForecast | What does future replacement volume and cost look like by year? |
| WarrantyTimeline | When does warranty expire by cohort? |
| DeviceLifecycle | How long do devices actually last by model? |

### Fleet Composition

| Sheet | What It Answers |
|-------|-----------------|
| ★ FleetSummary | Top-line KPIs: total assets, value, age, warranty, tickets, assignment |
| ★ LocationSummary | How many assets per school? How old? Warranty status? |
| ★ ModelBreakdown | Which device models do we have? How many active vs retired? |
| LocationModelBreakdown | What models are at each school? (location/model breakdown) |
| LocationModelFiltered | Show me one school's model mix (dropdown-driven) |
| CategoryBreakdown | What types of devices? Chromebooks vs laptops vs tablets? |
| ManufacturerSummary | Which vendors are we invested in? |

**Calculated columns on every asset (ARRAYFORMULA — instant even at 300K+ rows):**
- Device age in days and years (falls back to CreatedDate if PurchasedDate is empty)
- Warranty status: Active / Expiring (< 90 days) / Expired / None

## Quick Start

1. Create a new Google Spreadsheet
2. Go to **Extensions > Apps Script**
3. Copy all `.gs` files from the `scripts/` directory
4. Save and refresh the spreadsheet
5. Run **iiQ Assets > Setup > Setup Spreadsheet**
6. Fill in the **Config** sheet:
   - `API_BASE_URL`: Your iiQ instance URL (e.g., `https://yourdistrict.incidentiq.com`)
   - `BEARER_TOKEN`: Your API token (Admin > Integrations > API)
   - `SITE_ID`: Optional — only needed for multi-site instances
7. Run **iiQ Assets > Setup > Verify Configuration**
8. Run **iiQ Assets > Asset Data > Load / Resume Assets**
   - Reference data (locations, status types) loads automatically on first run
   - Script runs ~5.5 minutes per batch, then pauses (Apps Script time limit)
9. Run **iiQ Assets > Setup > Setup Automated Triggers**
   - Automation finishes loading, applies formulas, and keeps data current going forward

### Replacement Planning Config (Optional)

Two Config keys control the ReplacementPlanning and ReplacementForecast sheets:

- `REPLACEMENT_AGE_YEARS` — devices older than this are flagged for replacement (default: 4)
- `NEXT_SCHOOL_YEAR_START` — target date for planning (default: 2026-07-01, format YYYY-MM-DD)

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

## Project Structure

```
scripts/
  Config.gs              - Configuration, logging, concurrency control
  ApiClient.gs           - HTTP client with retry/backoff
  AssetData.gs           - Asset loader (bulk + incremental refresh)
  ReferenceData.gs       - Locations, status types, student enrollment
  Setup.gs               - Sheet creation, formulas, default analytics sheets (★)
  OptionalAnalytics.gs   - Optional (non-default) analytics sheet setup
  Menu.gs                - Menu system, category submenus, UI entry points
  Triggers.gs            - Automated trigger management
  appsscript.json        - Apps Script manifest
```

## License

Open source — use as a base for your district's asset reporting needs.
