# CLAUDE.md

Google Apps Script project for extracting Incident IQ (iiQ) asset data into Google Sheets for reporting and replacement cycle planning. Designed for school district IT teams tracking Chromebooks, laptops, and other devices.

## Architecture

```
iiQ API  →  Google Apps Script  →  Google Sheets  →  Looker Studio / Power BI
             (scripts/*.gs)        (raw data +        (dashboards)
                                    formula analytics)
```

**Data Flow:**
1. Reference data loads first (locations, status types, custom fields)
2. Phase 1: Bulk asset search (paginated, fast)
3. Phase 2: Custom field enrichment (AUE date from custom fields, per-batch)
4. Formula columns calculate derived metrics (AUE status, age, warranty, replacement cycle)
5. Analytics sheets auto-calculate via Google Sheets formulas

## Code Structure (scripts/)

| File | Purpose |
|------|---------|
| `Config.gs` | Config reader, type helpers, logging, lock management, config caching |
| `ApiClient.gs` | HTTP client with retry/backoff, asset search, location/status/custom-field endpoints |
| `AssetData.gs` | Two-phase loader: bulk assets (Phase 1) + custom field enrichment (Phase 2) |
| `ReferenceData.gs` | Locations, status types, custom field discovery (AUE auto-detection) |
| `Setup.gs` | Sheet creation, headers, formulas, analytics sheets |
| `Menu.gs` | "iiQ Assets" menu, UI entry points |
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
| Config | Manual | API settings, progress tracking, AUE field config |
| AssetData | Data | Main asset data (25 columns: 20 API + 5 formula) |
| Locations | Reference | Location directory |
| StatusTypes | Reference | Asset status type directory |
| CustomFields | Reference | Discovered custom fields for assets |
| Logs | Data | Operation logs |

### Analytics Sheets (Default)
| Sheet | Question Answered |
|-------|-------------------|
| LocationSummary | "How many assets per school? How old? AUE status?" |
| ModelBreakdown | "Which device models do we have? How many active vs retired?" |
| AUEPlanning | "When do devices need replacing, by fiscal year?" |
| BudgetPlanning | "What's the replacement cost per location?" |
| StatusOverview | "What's the breakdown of active/retired/storage?" |

## AssetData Column Layout (25 columns)

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
| S | OpenTickets | API (OpenTicketCount) |
| T | AUEDate | Phase 2 (custom field) |
| U | AUEStatus | Formula: Expired / < 6 Months / < 1 Year / < 2 Years / OK |
| V | AgeDays | Formula: TODAY() - PurchasedDate |
| W | AgeYears | Formula: AgeDays / 365.25 |
| X | WarrantyStatus | Formula: Active / Expiring / Expired / None |
| Y | ReplacementCycle | Formula: Fiscal year when AUE expires (e.g., "2025-2026") |

### Analytics Formula Column Reference

| Entity | Name Column (use for UNIQUE/COUNTIFS) |
|--------|---------------------------------------|
| Location | **I (LocationName)** |
| Model | **E (ModelName)** |
| Manufacturer | **F (ManufacturerName)** |
| Status | **M (StatusName)** |
| AUE Status | **U (AUEStatus)** |
| Replacement Cycle | **Y (ReplacementCycle)** |

## API Endpoints Used

| Endpoint | Method | Purpose |
|----------|--------|---------|
| `/v1.0/assets?$p={page}&$s={size}` | POST | Bulk asset search with filters |
| `/v2.0/locations/all?$s=1000` | GET | Location directory |
| `/v1.0/assets/status/types?$s=100` | GET | Asset status types |
| `/v1.0/custom-fields/for/asset` | POST | Discover custom field definitions |
| `/v1.0/custom-fields/values/for/assets` | POST | Batch custom field values |

## Config Sheet Keys

Required:
- `API_BASE_URL`: iiQ instance URL (e.g., `https://district.incidentiq.com`) — `/api` added automatically
- `BEARER_TOKEN`: JWT authentication token

Optional:
- `SITE_ID`: Site UUID (multi-site instances)
- `PAGE_SIZE`: Records per API call (default 100)
- `THROTTLE_MS`: Delay between requests (default 1000)
- `ASSET_BATCH_SIZE`: Assets per page for bulk load (default 500)

Custom Field Config (auto-detected or manual):
- `AUE_CUSTOM_FIELD_ID`: UUID of the AUE date custom field
- `AUE_CUSTOM_FIELD_NAME`: Display name of the AUE field

Progress Tracking (auto-managed):
- `ASSET_LAST_PAGE`, `ASSET_TOTAL_PAGES`, `ASSET_COMPLETE`: Phase 1 progress
- `ENRICH_LAST_IDX`, `ENRICH_COMPLETE`: Phase 2 progress

## Two-Phase Loading

**Phase 1 — Bulk Asset Search:**
- POST `/v1.0/assets` with empty filters, paginated
- 5.5-minute timeout with page-level checkpoints
- `triggerDataContinue` resumes across invocations

**Phase 2 — Custom Field Enrichment (AUE):**
- Reads asset IDs from column A, batches of 50
- POST `/v1.0/custom-fields/values/for/assets` per batch
- Writes AUE date to column T
- Only runs if `AUE_CUSTOM_FIELD_ID` is configured

**After both phases:** `applyAssetFormulas()` adds calculated columns U-Y.

## Initial Setup

1. Create a new Google Spreadsheet
2. Extensions > Apps Script → copy all `.gs` files
3. Save and refresh
4. **iiQ Assets > Setup > Setup Spreadsheet**
5. Fill in Config sheet (API_BASE_URL, BEARER_TOKEN)
6. **iiQ Assets > Setup > Verify Configuration**
7. **iiQ Assets > Load Reference Data > Refresh All Reference Data**
8. Check CustomFields sheet — verify AUE field detected (or set manually)
9. **iiQ Assets > Asset Data > Continue Loading** (repeat until complete)
10. **iiQ Assets > Asset Data > Enrich Custom Fields**
11. **iiQ Assets > Asset Data > Apply Formulas**
12. **iiQ Assets > Setup > Setup Automated Triggers**

## Trigger Schedule

| Function | Schedule | Purpose |
|----------|----------|---------|
| `triggerDataContinue` | Every 10 min | Resume interrupted loads (Phase 1 or 2) |
| `triggerWeeklyFullRefresh` | Sunday 2 AM | Full reload + reference data refresh |
