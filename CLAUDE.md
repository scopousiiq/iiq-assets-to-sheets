# CLAUDE.md

Google Apps Script project for extracting Incident IQ (iiQ) asset data into Google Sheets for reporting and replacement cycle planning. Designed for school district IT teams tracking Chromebooks, laptops, and other devices.

## Architecture

```
iiQ API  →  Google Apps Script  →  Google Sheets  →  Looker Studio / Power BI
             (scripts/*.gs)        (raw data +        (dashboards)
                                    formula analytics)
```

**Data Flow:**
1. Reference data loads first (locations, status types)
2. Bulk asset search (paginated, fast) with checkpoint resume
3. Formula columns calculate derived metrics (age, warranty status)
4. Analytics sheets auto-calculate via Google Sheets formulas
5. Daily incremental refresh keeps data current via ModifiedDate filter

## Code Structure (scripts/)

| File | Purpose |
|------|---------|
| `Config.gs` | Config reader, type helpers, logging, lock management, config caching |
| `ApiClient.gs` | HTTP client with retry/backoff, asset search, location/status endpoints |
| `AssetData.gs` | Bulk asset loader with checkpoint resume + incremental refresh |
| `ReferenceData.gs` | Locations and status types |
| `Setup.gs` | Sheet creation, headers, formulas, default analytics sheets (★) |
| `OptionalAnalytics.gs` | Optional (non-default) analytics sheet setup functions |
| `Menu.gs` | "iiQ Assets" menu, UI entry points, category submenus |
| `Triggers.gs` | Time-driven functions, trigger management |

**Key Dependencies:**
- `ApiClient.gs` → `Config.gs`
- `AssetData.gs` → `ApiClient.gs`, `Config.gs`
- `ReferenceData.gs` → `ApiClient.gs`, `Config.gs`

## Sheets Overview

### Data Sheets
| Sheet | Type | Purpose |
|-------|------|---------|
| Instructions | Static | Setup and usage guide |
| Config | Manual | API settings, progress tracking |
| AssetData | Data | Main asset data (28 columns: 25 API + 3 formula) |
| Locations | Reference | Location directory |
| StatusTypes | Reference | Asset status type directory |
| Logs | Data | Operation logs |

### Analytics Sheets (★ = default, created by Setup Spreadsheet)

**Fleet Operations**
| Sheet | Question Answered |
|-------|-------------------|
| ★ AssignmentOverview | "How many devices are assigned vs idle per location?" |
| ★ StatusOverview | "What's the breakdown of active/retired/storage?" |
| DeviceReadiness | "What's actually deployable at each school right now?" |
| SpareAssets | "Do I have enough working spares at each school?" |
| LostStolenRate | "Which schools are losing devices?" |
| ModelFragmentation | "Which schools are a patchwork of device models?" |
| UnassignedInventory | "Where is idle inventory sitting?" |

**Service & Reliability**
| Sheet | Question Answered |
|-------|-------------------|
| ★ ServiceImpact | "Which models generate the most support tickets? What's unreliable?" |
| BreakRate | "Which individual devices and models have the most tickets?" |
| HighTicketLocations | "Which schools have the most device problems?" |

**Budget & Planning**
| Sheet | Question Answered |
|-------|-------------------|
| ★ BudgetPlanning | "What's the replacement cost per location based on warranty/age?" |
| ★ AgingAnalysis | "What's our fleet age distribution? When is the replacement cliff?" |
| ReplacementPlanning | "What do I need to buy before next school year?" |
| ReplacementForecast | "How many devices need replacing in 1/2/3 years?" |
| WarrantyTimeline | "When does warranty expire by cohort?" |
| DeviceLifecycle | "How long do devices actually last by model?" |

**Fleet Composition**
| Sheet | Question Answered |
|-------|-------------------|
| ★ FleetSummary | "What are our top-line KPIs? Total assets, value, age, warranty, tickets, assignment?" |
| ★ LocationSummary | "How many assets per school? How old? Warranty status?" |
| ★ ModelBreakdown | "Which device models do we have? How many active vs retired?" |
| LocationModelBreakdown | "What models are at each school? (cross-tab)" |
| LocationModelFiltered | "Show me one school's model mix (dropdown-driven)" |
| CategoryBreakdown | "What types of devices do we have? Chromebooks vs laptops vs tablets?" |
| ManufacturerSummary | "Which vendors are we invested in?" |

**People**
| Sheet | Question Answered |
|-------|-------------------|
| IndividualLookup | "What's this person's device assignment history?" (dropdown-driven, live API fetch against `/users/{userId}/activities` — works for direct-assignment districts without formal checkouts) |

## Menu Structure

```
iiQ Assets
├── Setup
│   ├── Setup Spreadsheet
│   ├── Verify Configuration
│   └── Setup Automated Triggers
├── Asset Data
│   ├── Load / Resume Assets
│   ├── Refresh Updated Assets
│   └── Full Reload (All Assets)
├── Analytics Sheets
│   ├── Fleet Operations
│   │   ├── ★ AssignmentOverview
│   │   ├── ★ StatusOverview
│   │   ├── DeviceReadiness
│   │   ├── SpareAssets
│   │   ├── LostStolenRate
│   │   ├── ModelFragmentation
│   │   ├── UnassignedInventory
│   │   └── Regenerate Fleet Operations
│   ├── Service & Reliability
│   │   ├── ★ ServiceImpact
│   │   ├── BreakRate
│   │   ├── HighTicketLocations
│   │   └── Regenerate Service & Reliability
│   ├── Budget & Planning
│   │   ├── ★ BudgetPlanning
│   │   ├── ★ AgingAnalysis
│   │   ├── ReplacementPlanning
│   │   ├── ReplacementForecast
│   │   ├── WarrantyTimeline
│   │   ├── DeviceLifecycle
│   │   └── Regenerate Budget & Planning
│   ├── Fleet Composition
│   │   ├── ★ FleetSummary
│   │   ├── ★ LocationSummary
│   │   ├── ★ ModelBreakdown
│   │   ├── LocationModelBreakdown
│   │   ├── LocationModelFiltered
│   │   ├── CategoryBreakdown
│   │   ├── ManufacturerSummary
│   │   └── Regenerate Fleet Composition
│   ├── People
│   │   ├── Individual Lookup
│   │   └── Regenerate People
│   ├── Regenerate All Default (★)
│   └── Regenerate All Analytics
└── Reference Data
    ├── Reload Locations
    └── Reload Status Types
```

**★ = default sheets** created by Setup Spreadsheet. Non-starred sheets are optional.

**Regeneration:** Analytics setup functions use `getOrCreateSheet` -- on regeneration, only the formula is refreshed (no delete/create/reformat). Formulas are live and auto-recalculate when AssetData changes; regeneration is only needed after code updates.

## AssetData Column Layout (28 columns)

| Col | Header | Source |
|-----|--------|--------|
| A | AssetId | API |
| B | AssetTag | API |
| C | Name | API |
| D | SerialNumber | API |
| E | ModelName | API (Model.ModelName) |
| F | ManufacturerName | API (Model.Manufacturer.Name) |
| G | CategoryName | API (Model.CategoryNameWithParent) |
| H | LocationId | API |
| I | LocationName | API (Location.Name) |
| J | LocationType | API (Location.LocationTypeName) |
| K | OwnerId | API |
| L | OwnerName | API (Owner.Name) |
| M | StatusName | API (AssetStatus.Name) |
| N | PurchasedDate | API |
| O | WarrantyExpDate | API (WarrantyExpirationDate) |
| P | PurchasePrice | API |
| Q | CreatedDate | API |
| R | ModifiedDate | API |
| S | OwnerRoleName | API (Owner.RoleName) |
| T | OwnerGrade | API (Owner.Grade) |
| U | OwnerLocationId | API (Owner.LocationId) |
| V | StorageLocationName | API |
| W | StorageUnitNumber | API |
| X | DeployedDate | API |
| Y | OpenTickets | API (OpenTicketCount) |
| Z | AgeDays | ARRAYFORMULA: TODAY() - PurchasedDate (fallback CreatedDate) |
| AA | AgeYears | ARRAYFORMULA: AgeDays / 365.25 |
| AB | WarrantyStatus | ARRAYFORMULA: Active / Expiring / Expired / None |

### Analytics Formula Column Reference

| Entity | Name Column (use for UNIQUE/COUNTIFS) |
|--------|---------------------------------------|
| Location | **I (LocationName)** |
| Model | **E (ModelName)** |
| Manufacturer | **F (ManufacturerName)** |
| Status | **M (StatusName)** |
| Warranty Status | **AB (WarrantyStatus)** |
| Age (Years) | **AA (AgeYears)** |
| Open Tickets | **Y (OpenTickets)** |

## API Endpoints Used

| Endpoint | Method | Purpose |
|----------|--------|---------|
| `/v1.0/assets?$p={page}&$s={size}` | POST | Bulk asset search with filters (deleted assets excluded by default) |
| `/v2.0/locations/all?$s=1000` | GET | Location directory |
| `/v1.0/assets/status/types?$s=100` | GET | Asset status types |
| `/v1.0/sites/roles` | GET | Site roles (for STUDENT_ROLE_ID) |
| `/v1.0/users?$s=1` | POST | User count by filters (enrollment) |
| `/v1.0/users/{userId}/activities` | GET | Per-user activity log — filtered client-side for asset assignment events (IndividualLookup) |

## Config Sheet Keys

Required:
- `API_BASE_URL`: iiQ instance URL (e.g., `https://district.incidentiq.com`) — `/api` added automatically
- `BEARER_TOKEN`: JWT authentication token

Optional:
- `SITE_ID`: Site UUID (multi-site instances)
- `PAGE_SIZE`: Records per API call (default 100)
- `THROTTLE_MS`: Delay between requests (default 1000)
- `ASSET_BATCH_SIZE`: Assets per page for bulk load (default 500)
- `REPLACEMENT_AGE_YEARS`: Device age threshold for replacement planning (default 4)
- `NEXT_SCHOOL_YEAR_START`: Target date for replacement planning (default 2026-07-01, format YYYY-MM-DD)

Progress Tracking (auto-managed):
- `ASSET_LAST_PAGE`, `ASSET_TOTAL_PAGES`, `ASSET_COMPLETE`: Load progress
- `LAST_REFRESH_DATE`: ISO timestamp of last incremental refresh

Version Information (auto-managed):
- `SCRIPT_VERSION`: Installed script version (from `SCRIPT_VERSION` constant in `Config.gs`)
- `LATEST_VERSION`: Latest available version from GitHub, with status text and color-coded cell (green = current, yellow = update available)
- `VERSION_CHECK_DATE`: Date of last successful version check

**Version Check:** Scripts check GitHub daily for newer versions (piggybacked on `triggerDataContinue` — only fires if `isVersionCheckStale()` returns true, i.e. last check was >24h ago). Fetches `version.json` from the repo's `main` branch via raw.githubusercontent.com. Results land in the Config sheet with color coding — no pop-up dialogs. Manual check via **iiQ Assets > Setup > Check for Updates**. The version-check code fails silently if GitHub is unreachable — it must never break data operations.

## Data Loading

**Initial Load — Bulk Asset Search:**
- POST `/v1.0/assets` with empty filters, paginated
- 5.5-minute timeout with page-level checkpoints
- `triggerDataContinue` resumes across invocations

**Incremental Refresh — ModifiedDate Filter:**
- POST `/v1.0/assets` with `modifieddate` facet filter (`date>=YYYY-MM-DD`)
- Finds existing rows by AssetId and updates in-place
- Appends new assets not previously seen
- Runs daily at 3 AM via trigger, also available on-demand from menu

**After loading:** `applyAssetFormulas()` adds ARRAYFORMULA calculated columns Z-AB.

## Initial Setup

1. Create a new Google Spreadsheet
2. Extensions > Apps Script → copy all `.gs` files
3. Save and refresh
4. **iiQ Assets > Setup > Setup Spreadsheet**
5. Fill in Config sheet (API_BASE_URL, BEARER_TOKEN)
6. **iiQ Assets > Setup > Verify Configuration**
7. **iiQ Assets > Asset Data > Load / Resume Assets** (auto-loads reference data, then starts assets)
8. **iiQ Assets > Setup > Setup Automated Triggers** (automation finishes loading + applies formulas)

## Trigger Schedule

| Function | Schedule | Purpose |
|----------|----------|---------|
| `triggerDataContinue` | Every 10 min | Resume interrupted initial load |
| `triggerDailyRefresh` | Daily 3 AM | Incremental refresh (ModifiedDate filter) |
| `triggerWeeklyFullRefresh` | Sunday 2 AM | Full reload + reference data refresh |
