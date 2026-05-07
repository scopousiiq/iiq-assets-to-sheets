# Implementation Guide

This guide explains how the iiQ Assets to Sheets system works — the sheet structure, data loading, formulas, and technical details. It's written for IT staff who want to understand or customize the system.

**Just want to get started?** See the [README](README.md) for quick setup instructions.

### Architecture

```
┌─────────────┐     ┌──────────────────┐     ┌──────────────────┐
│  iiQ API    │────▶│  Google Sheets   │────▶│  Dashboards      │
│             │     │  + Apps Script   │     │                  │
│ /assets     │     │                  │     │ • Looker Studio  │
│ /locations  │     │ • Config sheet   │     │ • Power BI       │
│ /status     │     │ • Asset data     │     │ • Google Sheets  │
│ /users      │     │ • Analytics      │     │   (built-in)     │
└─────────────┘     └──────────────────┘     └──────────────────┘
```

---

## Part 1: Google Sheets Setup

### Step 1: Create the Spreadsheet

Create a new Google Sheet. The **Setup Spreadsheet** function creates all required sheets automatically.

**Data Sheets (created by Setup):**

| Sheet Name | Purpose |
|------------|---------|
| `Instructions` | Setup and usage guide (first tab) |
| `Config` | API credentials and settings |
| `AssetData` | Main asset inventory (33 columns: 30 API + 3 calculated) |
| `Locations` | Location directory |
| `StatusTypes` | Asset status types |
| `Logs` | Operation logs |

**Default Analytics Sheets (created by Setup):**

| Sheet Name | Purpose |
|------------|---------|
| `FleetSummary` | Executive KPIs: total assets, value, age, warranty, tickets |
| `LocationSummary` | Assets per school with enrollment data |
| `ModelBreakdown` | Device model inventory |
| `AgingAnalysis` | Fleet age distribution and replacement cliff |
| `BudgetPlanning` | Replacement cost by location |
| `ServiceImpact` | Models ranked by tickets-per-device |
| `AssignmentOverview` | Assigned vs idle devices per location |
| `StatusOverview` | Status breakdown (active, retired, storage) |

### Step 2: Install the Scripts

1. Open **Extensions > Apps Script** in your spreadsheet
2. Delete the default `Code.gs` file
3. Create one file per `.gs` file in the `scripts/` directory
4. Copy the contents of each file
5. Save all files, then close and refresh the spreadsheet

The **iiQ Assets** menu appears after refresh.

### Step 3: Configure

Run **iiQ Assets > Setup > Setup Spreadsheet**, then fill in the Config sheet:

| Key | Required | Description |
|-----|----------|-------------|
| `API_BASE_URL` | Yes | Your iiQ instance URL (e.g., `https://yourdistrict.incidentiq.com`) |
| `BEARER_TOKEN` | Yes | API token from Admin > Developer Tools |
| `SITE_ID` | No | Only for multi-site instances |

Run **iiQ Assets > Setup > Verify Configuration** to test your connection.

### Step 4: Load Data

Run **iiQ Assets > Asset Data > Load / Resume Assets**.

- Reference data (locations, status types) loads automatically on first run
- Each run processes ~5.5 minutes of data, then pauses (Apps Script limit)
- For small districts (< 2,000 assets), may finish in one run
- For large districts (10,000+ assets), set up triggers and let automation finish

### Step 5: Set Up Triggers

Run **iiQ Assets > Setup > Setup Automated Triggers**. This creates:

| Trigger | Schedule | Purpose |
|---------|----------|---------|
| `triggerDataContinue` | Every 10 min | Finishes initial load, then becomes a no-op |
| `triggerDailyRefresh` | Daily 3 AM | Incremental refresh (only changed assets) |
| `triggerWeeklyFullRefresh` | Sunday 2 AM | Full reload + reference data refresh |

After triggers are set up, you can close the spreadsheet. The triggers will finish loading and keep data current.

---

## Part 2: How Data Loading Works

### Initial Load (Bulk)

The initial load uses the iiQ asset search endpoint with pagination:

1. **POST** `/v1.0/assets?$p={page}&$s={batchSize}&$o=AssetCreatedDate asc`
2. Each page returns up to `ASSET_BATCH_SIZE` assets (default 500)
3. Progress is checkpointed to the Config sheet after each page
4. If the script times out (5.5-minute limit), `triggerDataContinue` resumes it

Sorting by `AssetCreatedDate asc` ensures a stable dataset during the paginated pull — new assets added during loading land at the end and don't shift earlier pages.

### Incremental Refresh (Daily)

The daily refresh fetches only assets modified since the last refresh:

1. **POST** `/v1.0/assets` with a `modifieddate` facet filter (`date>=YYYY-MM-DD`)
2. For each returned asset:
   - If the AssetId already exists in the sheet, the row is updated in-place
   - If it's a new asset, it's appended at the bottom
3. `LAST_REFRESH_DATE` is updated after each run

This is efficient — if your district has 50,000 assets but only 200 changed today, only those 200 are fetched.

### Weekly Full Refresh

The Sunday reload catches edge cases:

1. Refreshes reference data (locations, status types)
2. Clears all asset data and resets progress
3. Reloads everything from scratch
4. Reapplies formulas and continues enrollment if configured

### Deleted Assets

Deleted assets in iiQ are automatically excluded by the API — they are never downloaded. The weekly reload ensures any edge cases (like un-deleted assets) are handled.

---

## Part 3: AssetData Column Layout

The AssetData sheet has 33 columns: 30 from the API and 3 calculated by ARRAYFORMULA.

### API Columns (A-AD)

| Col | Header | API Source |
|-----|--------|------------|
| A | AssetId | `AssetId` |
| B | AssetTag | `AssetTag` |
| C | Name | `Name` |
| D | SerialNumber | `SerialNumber` |
| E | ModelName | `Model.ModelName` |
| F | ManufacturerName | `Model.Manufacturer.Name` |
| G | CategoryName | `Model.CategoryNameWithParent` |
| H | LocationId | `Location.LocationId` |
| I | LocationName | `Location.Name` |
| J | LocationType | `Location.LocationTypeName` |
| K | OwnerId | `Owner.UserId` |
| L | OwnerFullName | `Owner.FullName` (falls back to `Owner.Name`) |
| M | StatusName | `AssetStatus.Name` |
| N | PurchasedDate | `PurchasedDate` |
| O | WarrantyExpDate | `WarrantyExpirationDate` |
| P | PurchasePrice | `PurchasePrice` |
| Q | CreatedDate | `CreatedDate` |
| R | ModifiedDate | `ModifiedDate` |
| S | OwnerRoleName | `Owner.RoleName` |
| T | OwnerGrade | `Owner.Grade` |
| U | OwnerLocationId | `Owner.LocationId` |
| V | StorageLocationName | `StorageLocationName` |
| W | StorageUnitNumber | `StorageUnitNumber` |
| X | DeployedDate | `DeployedDate` |
| Y | OpenTickets | `OpenTicketCount` |
| Z | OwnerFirstName | `Owner.FirstName` |
| AA | OwnerLastName | `Owner.LastName` |
| AB | OwnerEmail | `Owner.Email` |
| AC | OwnerSchoolIdNumber | `Owner.SchoolIdNumber` |
| AD | LocationRoomName | `LocationRoom.Name` |

### Calculated Columns (AE-AG)

These are set as ARRAYFORMULAs in row 2 and spill down automatically:

| Col | Header | Formula Logic |
|-----|--------|---------------|
| AE | AgeDays | `TODAY() - PurchasedDate` (falls back to CreatedDate if empty) |
| AF | AgeYears | `AgeDays / 365.25` |
| AG | WarrantyStatus | "Active" / "Expiring" (< 90 days) / "Expired" / "None" |

**Note on device age:** If PurchasedDate is empty (common when districts don't track purchases in iiQ), CreatedDate is used as a fallback. CreatedDate is when the asset record was added to iiQ — a reasonable proxy for device age. All analytics sheets follow this same logic.

---

## Part 4: Analytics Sheets

All analytics sheets use Google Sheets formulas (LET, BYROW, UNIQUE, COUNTIFS, AVERAGEIFS, etc.) and auto-recalculate whenever AssetData changes. No scripts run to update them.

### Formula Architecture

Each analytics sheet is a single formula in cell A2 that spills into a full table:

```
=LET(
  locs, UNIQUE(FILTER(AssetData!I2:I, AssetData!I2:I<>"")),
  total, BYROW(locs, LAMBDA(loc, COUNTIF(AssetData!I:I, loc))),
  active, BYROW(locs, LAMBDA(loc, COUNTIFS(AssetData!I:I, loc, AssetData!M:M, "<>Retired"))),
  ...
  IFERROR(SORT(HSTACK(locs, total, active, ...), 2, FALSE), HSTACK(locs, total, active, ...))
)
```

Key patterns:
- **UNIQUE + FILTER** to get distinct entity values (locations, models, etc.)
- **BYROW + LAMBDA** to calculate per-row aggregates
- **HSTACK** to combine all columns into one output table
- **IFERROR(SORT(...), ...)** as a safety wrapper so data shows unsorted if sort fails

### Default vs Optional Sheets

**8 default sheets** (marked with ★) are created by Setup Spreadsheet. They cover the most common reporting needs.

**17 optional sheets** can be installed individually from the **iiQ Assets > Analytics Sheets** menu, organized into five category submenus (Fleet Operations, Service & Reliability, Budget & Planning, Fleet Composition, Lookups). Each submenu lists both default and optional sheets.

The **Lookups** category holds two interactive drill-down views that call the iiQ API live (rather than using a spreadsheet formula):

- `IndividualLookup` — dropdown-driven per-user asset assignment history. Pick a user from the dropdown in `B1`; the script fetches `/users/{userId}/activities` and writes the asset events into the sheet. Works for districts that assign devices directly without formal checkouts.
- `VerificationLookup` — paste-driven per-asset audit verification history. Type or paste an Asset Tag or Serial Number into `B1`; the script resolves it to an `AssetId` against `AssetData`, calls `/assets/{assetId}/verifications`, and writes the verification records (date, pass/fail, method, location, room, verifier name, comments). Verifier UUIDs are resolved to display names via `/users/{userId}` (one call per unique verifier, cached for the lookup).

### Regeneration

Analytics formulas are live — they update automatically when data refreshes. Regeneration is only needed after a code update changes a formula definition. Options:

- **Per-category:** e.g., "Regenerate Fleet Operations"
- **Regenerate All Default (★):** Rebuilds the 8 default sheets
- **Regenerate All Analytics:** Rebuilds all installed sheets

Regeneration refreshes the formula in-place (via `getOrCreateSheet`) — it does not delete and recreate the sheet.

---

## Part 5: Student Enrollment (Optional)

The LocationEnrollment feature counts students and device coverage per school.

### Setup

1. Run **iiQ Assets > Load Reference Data > View Available Roles** to see available roles
2. Copy the student role's `RoleId` into `STUDENT_ROLE_ID` on the Config sheet
3. Run **iiQ Assets > Load Reference Data > Refresh Location Enrollment**

### How It Works

For each location, two API calls are made:

1. `POST /v1.0/users?$s=1` with role + location filters → total students
2. Same + `hasassigneddevice` facet → students with assigned devices

This produces a `DeviceCoverage%` per location. The `LocationSummary` analytics sheet references this data.

Loading is checkpointed per-location. For large districts (200+ schools), it may take multiple runs — `triggerDataContinue` handles this automatically.

---

## Part 6: Code Structure

### File Overview

| File | Lines | Purpose |
|------|-------|---------|
| `Config.gs` | Configuration reader, type helpers, logging, LockService concurrency |
| `ApiClient.gs` | HTTP client with retry/backoff, asset search, location/status endpoints |
| `AssetData.gs` | Bulk loader with checkpoint resume + incremental refresh |
| `ReferenceData.gs` | Locations, status types, student enrollment |
| `Setup.gs` | Sheet creation, headers, formulas, Instructions sheet, default analytics |
| `OptionalAnalytics.gs` | Optional (non-default) analytics sheet setup functions |
| `Menu.gs` | "iiQ Assets" menu, category submenus, UI entry points |
| `Triggers.gs` | Time-driven trigger functions |
| `appsscript.json` | Apps Script manifest |

### Dependency Graph

```
Config.gs        ← foundation (every file uses this)
  ↑
ApiClient.gs     ← HTTP layer (retry, backoff, auth)
  ↑
AssetData.gs     ← data loading (bulk + incremental)
ReferenceData.gs ← reference loading (locations, status, enrollment)
  ↑
Setup.gs         ← sheet creation + formulas
OptionalAnalytics.gs
  ↑
Menu.gs          ← user-facing entry points
Triggers.gs      ← automated entry points
```

### Key Patterns

**Pagination with timeout handling:**
```javascript
const MAX_RUNTIME_MS = 5.5 * 60 * 1000; // Stay under 6min Apps Script limit
while (Date.now() - startTime < MAX_RUNTIME_MS) {
  // Process one page, checkpoint progress to Config sheet
  // triggerDataContinue resumes on next run
}
```

**Concurrency control (LockService):**
- Menu functions use `acquireScriptLock()` — waits briefly, shows "busy" if unavailable
- Trigger functions use `tryAcquireScriptLock()` — skips gracefully if locked
- Prevents data corruption from overlapping operations

**Trigger safety (destructive operations):**
- Full Reload and Clear Data require all triggers removed first
- `requireNoTriggers()` blocks the operation and shows instructions if triggers exist

**Rate limiting:**
- Exponential backoff: 2s base, doubles on retry (up to 3 retries)
- Configurable throttle via `THROTTLE_MS` in Config sheet (default 1000ms)

---

## Part 7: Customization

### Adding a New Column to AssetData

1. Add the header to `ASSET_HEADERS` in `AssetData.gs`
2. Update `ASSET_DATA_COLS` (if it's an API column) or `ASSET_TOTAL_COLS`
3. Add the extraction in `extractAssetRow()`
4. If it's a formula column, add the ARRAYFORMULA in `applyAssetFormulas()` in `Setup.gs`
5. Update the column layout comment at the top of `AssetData.gs`

### Adding a New Analytics Sheet

1. Write a `setup{SheetName}Sheet(ss)` function following the LET/BYROW/HSTACK pattern
2. Use `getOrCreateSheet()` for idempotent creation
3. Add a menu entry in `Menu.gs`
4. Add the function call to the appropriate regenerate function

### Changing the Replacement Age Threshold

Set `REPLACEMENT_AGE_YEARS` in the Config sheet (default 4). The `BudgetPlanning`, `ReplacementPlanning`, and `ReplacementForecast` sheets read this value via the formula.

---

## Part 8: API Reference

### Endpoints Used

| Endpoint | Method | Purpose |
|----------|--------|---------|
| `/v1.0/assets?$p={page}&$s={size}&$o={sort}` | POST | Bulk asset search with filters |
| `/v2.0/locations/all?$s=1000` | GET | Location directory |
| `/v1.0/assets/status/types?$s=100` | GET | Asset status types |
| `/v1.0/sites/roles` | GET | Site roles (for finding STUDENT_ROLE_ID) |
| `/v1.0/users?$s=1` | POST | User count by filters (enrollment) |
| `/v1.0/users/{userId}/activities` | GET | Per-user activity log (IndividualLookup) |
| `/v1.0/users/{userId}` | GET | Single user lookup (VerificationLookup verifier resolution) |
| `/v1.0/assets/{assetId}/verifications` | GET | Per-asset audit verification history (VerificationLookup) |

### Authentication

All requests include:
- `Authorization: Bearer {token}` — from Config `BEARER_TOKEN`
- `SiteId: {uuid}` header — from Config `SITE_ID` (if set)
- `Content-Type: application/json`

The `API_BASE_URL` is automatically suffixed with `/api` if not already present.

### Asset Search Filters

The asset search endpoint accepts a `Filters` array in the POST body:

```json
{
  "Filters": [
    { "Facet": "modifieddate", "Value": "date>=2026-01-01" }
  ]
}
```

For the initial load, an empty filter array is used (fetches all non-deleted assets). For incremental refresh, the `modifieddate` facet narrows the results.

### Sort Parameters

The `$o` query parameter controls result ordering:

- Initial load: `AssetCreatedDate asc` (stable dataset during pagination)
- Incremental refresh: `AssetModifiedDate asc` (oldest changes first)

Sorting is critical for reliable pagination — unsorted results can shift between pages during long-running loads.

---

## Part 9: Built-in Dashboard (Web App)

A native Apps Script Web App ships with this project (added in v1.5.0). Each district deploys their own instance bound to their sheet — no Looker Studio account needed for the basic view.

### Architecture

```
viewer browser → doGet(e) → DashboardApp.html (Chart.js 4.4.1)
                   ↓ google.script.run.getDashboardData()
                getDashboardData()
                   ↓ reads AssetData + iterates CHART_REGISTRY
                returns { kpis, badges, categoryGroups }
```

Files:
- `scripts/Dashboard.gs` — server entry points (`doGet`, `getDashboardData`, `showDashboardUrl`).
- `scripts/ChartRegistry.gs` — declarative `CHART_REGISTRY` array (19 chart specs across 5 categories) plus `CATEGORY_ORDER`.
- `scripts/DashboardApp.html` — single-page frontend (Chart.js 4.4.1 via CDN).

### Deployment

1. **Extensions → Apps Script** from the sheet.
2. **Deploy → New deployment** → Type: **Web app**.
3. Execute as: **Me**. Who has access: **Anyone within your domain**.
4. Click **Deploy**, authorize, copy the `/exec` URL.
5. Paste into the `DASHBOARD_URL` row in the Config sheet.
6. **iiQ Assets > Setup > Show Dashboard URL** confirms.

For code updates: **Deploy → Manage deployments → Edit → New version** publishes to the same URL.

### Dashboard Contents

**4 KPI cards** computed directly from `AssetData`:
- Total Assets — row count
- Avg Age — mean of column AF (AgeYears)
- Warranty Active % — share of column AG = "Active"
- Assignment Rate % — share of rows with non-empty `OwnerId`

**5 chart tabs + 2 lookup tabs** (lookup tabs only appear when their source sheet is installed):

| Tab | Charts |
|-----|--------|
| Composition | LocationSummary, ModelBreakdown, CategoryBreakdown, ManufacturerSummary |
| Status | StatusOverview, AssignmentOverview, DeviceReadiness, SpareAssets |
| Aging | AgingAnalysis, WarrantyTimeline, DeviceLifecycle |
| Budget | BudgetPlanning, ReplacementPlanning, ReplacementForecast |
| Service | ServiceImpact, BreakRate, HighTicketLocations, LostStolenRate, ModelFragmentation |
| Individual | IndividualLookup (interactive — dropdown of users, results as HTML table) |
| Verification | VerificationLookup (interactive — paste Asset Tag/Serial, results as HTML table) |

Charts only render if you've installed the source sheet via **Analytics Sheets** menu. Lookup tabs make live API calls (`/users/{userId}/activities` and `/assets/{assetId}/verifications`) — they run server-side as the deployer, so viewers don't need spreadsheet or iiQ access.

### Adding a New Chart

When you add a new analytics sheet (whether default or optional), it won't show up in the dashboard until you register it in `ChartRegistry.gs`:

```javascript
{
  sheetName: 'YourNewSheet',
  category: 'Fleet Composition',  // or any of the 5 categories
  tabLabel: 'Composition',
  charts: [{
    title: 'My New Chart',
    type: 'horizontalBar',  // or bar, line, stackedBar, stackedHorizontalBar
    labelCol: 0,            // 0-indexed
    series: [
      { header: 'Count', col: 1, color: 'darkBlue' }
    ],
    rowStart: 2,
    rowMode: 'fixed',
    rowCount: 20
  }]
}
```

After deploying, the new chart appears on the next page refresh.

### Why Not Looker Studio?

The built-in dashboard is **complementary** to Looker Studio — it's not a replacement. Use the built-in dashboard for:
- Quick at-a-glance views of fleet composition and status
- Easy sharing with non-technical stakeholders (no Looker account needed)
- Snapshot reporting on a known set of metrics

Use Looker Studio for:
- Custom filtering, parameterization, drill-downs
- Combining multiple data sources
- Scheduled email delivery
- More chart types (pie, geo, table, etc.)

