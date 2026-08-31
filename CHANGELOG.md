# Changelog — iiQ Assets to Sheets

All notable changes to this project are documented here.

---

## v1.6.0 — Asset custom fields (2026-08-31)

### Added
- **Up to five district-defined asset custom fields as AssetData columns AM-AQ** (`CustomField1`-`CustomField5`). Values are read out of the `CustomFieldValues` array already present on every item in the `/v1.0/assets` search response, so no per-asset request is added — the only extra traffic is a single `POST /v1.0/custom-fields/for/asset` per loading run.
- **`CustomFields` sheet** listing every asset custom field in the district (Name, CustomFieldTypeId, EditorType). Populated by the new **iiQ Assets > Setup > Refresh Custom Fields** menu item. Rows are deduplicated by `CustomFieldTypeId`, because `/custom-fields/for/asset` returns one row per field-to-filter-set mapping rather than one per field.
- **`CUSTOM_FIELD_1`-`CUSTOM_FIELD_5` Config keys.** Paste a `CustomFieldTypeId` from the `CustomFields` sheet. A field display name is also accepted, but ids are unambiguous — a district can define two different fields sharing one name.
- **`CUSTOM_FIELD_1_ID`-`_5_ID` Config keys (auto-managed)** showing what each slot resolved to, or `NOT_FOUND`. Written for diagnostics on every run and never read back as a cache, so editing a slot cannot leave it pulling the previously resolved field.
- **Verify Configuration now reports custom field slots**, checking each configured id against the district's actual field list rather than just its GUID shape.
- Dropdown, multi-select, and iiQ-location custom fields have their stored ids translated to display names. Multi-value fields are joined with a comma.

### Changed
- **AssetData widens from 38 to 43 columns.** The custom field block is appended *after* the formula columns, so `AgeDays`/`AgeYears`/`WarrantyStatus` stay at `AJ`/`AK`/`AL`. Every analytics formula and every dashboard column offset is unchanged.
- `Remove Duplicates` now carries custom field values through deduplication instead of dropping them.
- `scripts/Config.gs` `SCRIPT_VERSION` -> `1.6.0`.

### Fixed
- `Setup Spreadsheet` and `Clear Data + Reset Progress` addressed AssetData's full column width without first growing the sheet's grid, which throws on a grid narrower than the layout. Both now widen the grid first.

### Upgrade Notes
This release is additive — no destructive re-setup, and no `BEARER_TOKEN` to save and restore.

1. Update all `.gs` files from the `scripts/` directory (note the new `CustomFields.gs`).
2. Reload the spreadsheet so the menu picks up **Refresh Custom Fields**.
3. Run **iiQ Assets > Setup > Refresh Custom Fields**. This creates the `CustomFields` sheet and adds the `CUSTOM_FIELD_*` rows to your existing Config sheet.
4. Copy a `CustomFieldTypeId` from the `CustomFields` sheet into `CUSTOM_FIELD_1` (through `_5`) on the Config sheet.
5. Run **iiQ Assets > Asset Data > Full Reload** to backfill the new columns for the whole fleet. Without it, the columns fill in gradually as assets change.

Districts that configure no custom fields are unaffected: the columns stay empty and no extra API calls are made.

---

## v1.5.2 — Funding source and last verification columns (2026-05-13)

### Added
- **AssetData** gains five new API-sourced columns extracted directly from the existing `/v1.0/assets` search payload (no new endpoints):
  - `AE` — `FundingSourceName` from `FundingSource.Name`
  - `AF` — `LastVerificationDate` from `LastVerificationDateTime` (formatted yyyy-MM-dd)
  - `AG` — `LastVerificationType` from `LastVerificationTypeName`
  - `AH` — `LastVerificationLocation` from `LastVerificationLocationName`
  - `AI` — `LastVerificationSuccess` from `LastVerificationSuccessful` (boolean)

### Changed
- **AssetData column layout** (35 API cols + 3 formula cols = 38 total; was 30 + 3 = 33):
  - Formula columns shifted by 5: `AgeDays` is now `AJ` (was `AE`), `AgeYears` is now `AK` (was `AF`), `WarrantyStatus` is now `AL` (was `AG`).
- All analytics formulas updated accordingly (`AssetData!AF:AF` → `AssetData!AK:AK`, `AssetData!AG:AG` → `AssetData!AL:AL`, etc.) across `Setup.gs` and `OptionalAnalytics.gs`.
- `Dashboard.gs` `DASH_COL` offsets updated: `AGE_YEARS` 31 → 36, `WARRANTY_STATUS` 32 → 37.
- `scripts/Config.gs` `SCRIPT_VERSION` → `1.5.2`.
- Instructions sheet and CLAUDE.md / README.md / GUIDE.md column-layout tables updated to the 38-column layout.

### Upgrade Notes
1. **Save your `BEARER_TOKEN`** from the Config sheet before proceeding.
2. Remove automated triggers: **iiQ Assets > Setup > Remove Automated Triggers**.
3. Update all `.gs` files from the `scripts/` directory.
4. Run **iiQ Assets > Setup > Setup Spreadsheet** (destructive: recreates all sheets with the 38-column layout and correct headers).
5. Paste your `BEARER_TOKEN` back into the Config sheet.
6. Run **iiQ Assets > Asset Data > Full Reload** to repopulate AssetData in the new 38-column layout.
7. Run **iiQ Assets > Setup > Setup Automated Triggers** to restore automation.
8. For the dashboard: **Deploy → Manage deployments → Edit → New version** to publish the updated `DASH_COL` offsets.

---

## v1.5.1 — Email in Individual Lookup dropdown (2026-05-07)

### Changed
- **Individual Lookup dropdown** now displays `Name (email)` instead of just `Name` — disambiguates same-name users (which is common in larger districts). Applies to both the in-sheet `IndividualLookup` (column Z dropdown source) and the dashboard's **Individual** tab.
- `resolveOwnerIdByName` now parses the `Name (email)` format and matches by email first (more specific), falling back to name. Users without an email in AssetData fall back to the bare-name format and continue to work unchanged.

### Upgrade Notes
1. Update `scripts/OptionalAnalytics.gs` and `scripts/Dashboard.gs` from the repo (or copy the updated master template).
2. Regenerate the IndividualLookup sheet via **iiQ Assets > Analytics Sheets > Lookups > Individual Lookup** so column Z picks up the new formula.
3. For the dashboard: **Deploy → Manage deployments → Edit → New version** to publish, then refresh the Individual tab.

---

## v1.5.0 — Web app dashboard, audit verification, room granularity (2026-05-07)

A major release combining a new built-in web-app dashboard, asset audit verification history, room-level granularity in inventory and cross-tab analytics, and a handful of API extraction fixes.

### Added
- **Asset Dashboard Web App** — a deployable Apps Script Web App that renders a full-page tabbed dashboard powered by `AssetData` and registered analytics sheets. Each district deploys their own instance bound to their sheet.
  - **3 new files**: `scripts/Dashboard.gs`, `scripts/ChartRegistry.gs`, `scripts/DashboardApp.html`.
  - **4 KPI cards** computed directly from `AssetData`: Total Assets, Avg Age (years), Warranty Active %, Assignment Rate %.
  - **19 charts** across **5 category tabs**: Fleet Composition, Status & Operations, Aging & Warranty, Replacement Planning, Service & Risk. Charts only render if the source analytics sheet exists.
  - **2 interactive Lookup tabs** (when their source sheets are installed):
    - **Individual** — dropdown of owner names; selection runs `/v1.0/users/{userId}/activities` and renders the events as an HTML table.
    - **Verification** — text input accepting Asset Tag *or* Serial Number; runs `/v1.0/assets/{assetId}/verifications` and renders results as an HTML table. Verifier UUIDs resolve to names via `/v1.0/users/{userId}`; room UUIDs resolve to names via `/v2.0/locations/rooms/{roomId}` (one call per unique value, cached for the lookup).
  - **Chart.js 4.4.1** frontend (CDN-loaded), 2-column responsive grid, refresh button, deep-linkable tab hashes.
  - New menu item: **iiQ Assets > Setup > Show Dashboard URL** — modal with the deployed URL or, if not yet deployed, with deployment instructions.
  - New Config row: `DASHBOARD_URL` (paste the `/exec` URL after deploying).
- **VerificationLookup sheet** — new optional sheet for per-asset audit verification history. Paste or type an Asset Tag or Serial Number in `B1`; the sheet resolves the asset against `AssetData` and live-fetches verification records from `/v1.0/assets/{assetId}/verifications`.
  - Columns: Date, Result (Pass/Fail), Method, Location, Room, Verified By, Comments.
  - Install via **iiQ Assets > Analytics Sheets > Lookups > Verification Lookup**.
- **API client additions** in `scripts/ApiClient.gs`:
  - `getAssetVerifications(assetId)` — `/v1.0/assets/{assetId}/verifications`
  - `getUserById(userId)` — `/v1.0/users/{userId}` (for verifier name resolution)
  - `getLocationRoomById(roomId)` — `/v2.0/locations/rooms/{roomId}` (for room name resolution)

### Changed
- **UnassignedInventory** — now groups by **Location + Room** (was Location only). One row per unique `(LocationName, LocationRoomName)` pair, sorted Location A-Z then Room A-Z. Blank room cells display as `(no room)`.
- **LocationModelBreakdown** — gained a **Room** column (now Location → Room → Model grouping) plus a **Filter by Location** dropdown in `B1` (blank = show all). Row 2 holds column headers; data starts in row 3.
- **Analytics menu** — "People" submenu renamed to **Lookups** (now hosts both `IndividualLookup` and `VerificationLookup`).
- `scripts/Config.gs` `SCRIPT_VERSION` → `1.5.0`.
- `scripts/Setup.gs` Instructions sheet adds a **Built-in Asset Dashboard** section and updates the Lookups submenu reference.

### Fixed
- **VerificationLookup verifier names** — were displaying as UUIDs because `getUserById` returns a `{ Item: UserDetail }` envelope; the resolver now unwraps `.Item` and reads `Name` (with `FirstName`/`LastName` fallback).
- **VerificationLookup room** — Room column was showing the LocationRoom UUID; now resolved via `/v2.0/locations/rooms/{roomId}` (envelope unwrapped from `.Item.Name`). Falls back to UUID on API miss.
- **CategoryBreakdown #N/A error** — formula now wraps the `LET(...)` in an outer `IFERROR` and shows a friendly message when `AssetData` column G (CategoryName) is empty for all rows. Previously crashed with "No matches are found in FILTER evaluation."
- **AssetData CategoryName extraction** — added `Model.Category.Name` as a third fallback after `Model.CategoryNameWithParent` and `Model.CategoryName`. Many iiQ instances (including demo) return `CategoryNameWithParent: null` but populate `Category.Name`. Existing rows need a Full Reload (or wait for Sunday's weekly refresh) to backfill column G.

### Upgrade Notes
This release is **non-destructive** for AssetData layout. Two-step upgrade:

1. **Update files**:
   - Update all `.gs` files from the `scripts/` directory.
   - **Add** `Dashboard.gs`, `ChartRegistry.gs`, `DashboardApp.html` (new files).
   - Save and refresh the spreadsheet.
2. **Deploy the web app** (optional — only if you want the dashboard):
   - Open **Extensions → Apps Script**.
   - **Deploy → New deployment** → Type: **Web app**.
   - Execute as: **Me**; Access: **Anyone within your domain**.
   - Click **Deploy**, authorize, copy the `/exec` URL.
   - Paste into the `DASHBOARD_URL` row in the Config sheet.
   - Confirm via **iiQ Assets > Setup > Show Dashboard URL**.

For future code updates, use **Deploy → Manage deployments → Edit → New version** so the URL remains stable.

To populate column G (CategoryName) for existing rows, run **iiQ Assets > Asset Data > Full Reload** (requires triggers off first) or wait for the Sunday 2 AM weekly refresh trigger.

To regenerate analytics sheets with the new column layouts: **iiQ Assets > Analytics Sheets > Fleet Operations > Unassigned Inventory** and **iiQ Assets > Analytics Sheets > Fleet Composition > Location Model Breakdown**.

---

## v1.4.1 — Location room name (2026-05-07)

### Changed
- **AssetData column layout** (30 API cols + 3 formula cols = 33 total; was 29 + 3 = 32):
  - `AD` (new) — `LocationRoomName` from `LocationRoom.Name`.
  - Formula columns shifted by 1: `AgeDays` is now `AE` (was `AD`), `AgeYears` is now `AF` (was `AE`), `WarrantyStatus` is now `AG` (was `AF`).
- All analytics formulas updated accordingly (`AE:AE` → `AF:AF`, `AF:AF` → `AG:AG`, etc.).

### Upgrade Notes
1. **Save your `BEARER_TOKEN`** from the Config sheet before proceeding.
2. Remove automated triggers: **iiQ Assets > Setup > Remove Automated Triggers**.
3. Update all `.gs` files from the `scripts/` directory.
4. Run **iiQ Assets > Setup > Setup Spreadsheet** (destructive: recreates all sheets with the 33-column layout and correct headers).
5. Paste your `BEARER_TOKEN` back into the Config sheet.
6. Run **iiQ Assets > Asset Data > Full Reload** to repopulate AssetData in the new 33-column layout.
7. Run **iiQ Assets > Setup > Setup Automated Triggers** to restore automation.

---

## v1.4.0 — Owner email and school ID fields (2026-05-06)

### Changed
- **AssetData column layout** (29 API cols + 3 formula cols = 32 total; was 27 + 3 = 30):
  - `AB` (new) — `OwnerEmail` from `Owner.Email`.
  - `AC` (new) — `OwnerSchoolIdNumber` from `Owner.SchoolIdNumber`.
  - Formula columns shifted by 2: `AgeDays` is now `AD` (was `AB`), `AgeYears` is now `AE` (was `AC`), `WarrantyStatus` is now `AF` (was `AD`).
- All analytics formulas that reference the calculated columns updated accordingly (`AC:AC` → `AE:AE`, `AD:AD` → `AF:AF`, etc.).

### Upgrade Notes
1. **Save your `BEARER_TOKEN`** from the Config sheet before proceeding.
2. Remove automated triggers: **iiQ Assets > Setup > Remove Automated Triggers**.
3. Update all `.gs` files from the `scripts/` directory.
4. Run **iiQ Assets > Setup > Setup Spreadsheet** (destructive: recreates all sheets with the 32-column layout and correct headers).
5. Paste your `BEARER_TOKEN` back into the Config sheet.
6. Run **iiQ Assets > Asset Data > Full Reload** to repopulate AssetData in the new 32-column layout.
7. Run **iiQ Assets > Setup > Setup Automated Triggers** to restore automation.

**Why not just "Full Reload"?** The `Full Reload` operation clears data rows (row 2+) but does not rewrite row 1 headers. Running Full Reload without Setup Spreadsheet first would leave stale 30-column headers while data is in 32-column layout, misaligning header labels with data. Setup Spreadsheet deletes and recreates the AssetData sheet, ensuring headers and data layout match.

---

## v1.3.0 — Canonical telemetry client + polling-requires-opt-in (2026-04-24)

### Changed
- **Telemetry rewired to the shared client** (`scripts/Telemetry.gs`, copied from `iiq-sheets-telemetry/client/`). Replaces the inline `reportTelemetry`/`sha256Hex` block that lived in `Config.gs` through v1.2.0.
- **District identification**: `instanceUrl` (hostname derived from `API_BASE_URL`, e.g. `demo.incidentiq.com`) replaces the SHA-256 `districtHash`. iiQ owns the customer relationship, so direct identification is the correct model.
- **Payload** (`schemaVersion: 1`): `{schemaVersion, installId, project, version, instanceUrl, installedAt, scriptTimeZone, sentAt, rowCount, primarySheet, analyticsSheets}`. No asset content, no PII, no credentials, no config values beyond the enable flag.
- **`installId`** Script Property key renamed from `INSTALL_ID` to `TELEMETRY_INSTALL_ID`. Existing installs get a fresh UUID on first v1.3.0 ping — the server treats them as new installs from that point.
- **`TELEMETRY_LAST_SENT`** Config row is no longer written or used. Client-side throttling is replaced by server-side dedupe (~6h per install/project/day). Existing rows are harmless and can be left or deleted.

### Added — Policy: automated polling requires telemetry opt-in
- **Runtime gate**: `enforceTelemetryGate()` runs as the first line of every time-based trigger handler (`triggerDataContinue`, `triggerDailyRefresh`, `triggerWeeklyFullRefresh`). If `TELEMETRY_ENABLED != TRUE`, the handler uninstalls every time-based trigger in the project and returns. Edit / open triggers (e.g. IndividualLookup) are left alone.
- **Install-time gate**: `assertTelemetryEnabledForTriggers()` runs at the top of `setupAutomatedTriggers()`. Setup Automated Triggers now throws a user-visible error and installs nothing when `TELEMETRY_ENABLED != TRUE`.
- **Instructions sheet** gains an "iiQ Telemetry" section documenting what's sent, how to opt out, and the consequence that opting out also disables automated polling. Manual menu actions continue to work regardless.

### Upgrade Notes
1. Pull all `.gs` files — in particular, the new `scripts/Telemetry.gs`.
2. Existing installs keep `TELEMETRY_ENABLED=TRUE` and continue polling with no action required. The legacy `TELEMETRY_LAST_SENT` row is ignored and can be deleted or left.
3. **To opt out**: set `TELEMETRY_ENABLED=FALSE`. The next scheduled trigger fire will uninstall all time-based triggers automatically. Manual refreshes from the menu continue to work.
4. **To re-enable after opting out**: set `TELEMETRY_ENABLED=TRUE`, then run **iiQ Assets > Setup > Setup Automated Triggers** to reinstall the triggers.
5. **Maintainer note**: `TELEMETRY_URL` is still empty in the shipped code. Paste the deployed `/exec` URL from the Telemetry Master before pushing to distribution.

---

## v1.2.0 — Owner name split (2026-04-24)

### Changed
- **AssetData column layout** (27 API cols + 3 formula cols = 30 total; was 25 + 3 = 28):
  - `L` — renamed `OwnerName` → `OwnerFullName`. Populated from `Owner.FullName` (falls back to `Owner.Name`).
  - `Z` (new) — `OwnerFirstName` from `Owner.FirstName`.
  - `AA` (new) — `OwnerLastName` from `Owner.LastName`.
  - Formula columns shifted by 2: `AgeDays` is now `AB` (was `Z`), `AgeYears` is now `AC` (was `AA`), `WarrantyStatus` is now `AD` (was `AB`).
- All analytics formulas that reference the calculated columns updated accordingly (`AA:AA` → `AC:AC`, `AB:AB` → `AD:AD`, etc.).

### Upgrade Notes
1. Update all `.gs` files from the `scripts/` directory.
2. Run **iiQ Assets > Analytics Sheets > Regenerate All Analytics** so installed analytics pick up the new formula column positions.
3. Existing AssetData rows still carry old values in the old columns — to repopulate the new owner-name columns and move the formula columns, either:
   - Run **iiQ Assets > Setup > Setup Spreadsheet** (destructive: recreates all sheets, clears data), or
   - Run **iiQ Assets > Asset Data > Full Reload** after removing triggers (reloads all asset data into the new layout).
4. After the next daily refresh (or any incremental refresh), any modified assets will be rewritten with the new column layout in place.

---

## v1.1.0 — Telemetry ping (2026-04-23)

### Added
- **Telemetry ping** — `reportTelemetry()` in `Config.gs` POSTs a small JSON payload (installId, version, district hash, asset count, installed analytics sheet names, trigger count) once per 24 hours to a hardcoded `TELEMETRY_URL` endpoint constant. Piggybacks on `triggerDataContinue` alongside the version check. **On by default for new installs; disable by setting `TELEMETRY_ENABLED=FALSE` in Config.**
- Config rows: `TELEMETRY_ENABLED` (defaults to `TRUE` on new Setup Spreadsheet runs) and `TELEMETRY_LAST_SENT` (auto-managed transparency field). The endpoint URL itself is not in Config — it's a maintainer-controlled constant in `Config.gs`.
- Server counterpart lives in the sibling `iiq-sheets-telemetry/` directory — a Google Apps Script Web App that accepts pings and appends to a `Pings` sheet in a private master spreadsheet.

### Privacy
- No PII or asset content is sent. District identity is a SHA-256 hash of `API_BASE_URL`, allowing CSMs to intersect their own hashed distribution list against the master for attribution without the master storing raw district domains.
- `installId` is a stable UUID per installed copy, generated on first ping via `PropertiesService`.
- All failures are silent and must never affect data operations.

### Upgrade Notes
1. Update all `.gs` files. **Existing pre-1.1.0 installs that upgrade without re-running Setup Spreadsheet will NOT auto-enable telemetry** — their Config sheet has no `TELEMETRY_ENABLED` row, which is read as disabled. Telemetry only kicks in if you either (a) run Setup Spreadsheet fresh (destructive — recreates all sheets) or (b) manually add `TELEMETRY_ENABLED=TRUE` and `TELEMETRY_URL=<endpoint>` to your Config.
2. To disable telemetry on a new install, set `TELEMETRY_ENABLED` to `FALSE` in the Config sheet. No restart needed — the next scheduled run will honor the change.

---

## v1.0.0 — IndividualLookup + versioning baseline (2026-04-23)

First tagged release. Establishes the version-check infrastructure going forward.

### Added
- **IndividualLookup analytics sheet** (new **People** submenu under Analytics Sheets). Dropdown-driven per-user asset assignment history — select a user in B1, the sheet fetches their full assignment/unassignment history live from `GET /v1.0/users/{userId}/activities` and writes it as rows. Works for districts that assign devices directly without formal checkout transactions.
  - Columns: Date, Action, Asset Tag, Serial Number, Model, Location, Currently With.
  - Dropdown source is a sorted unique list of current owners from `AssetData!L` (hidden helper column Z on the sheet).
  - Installable onEdit trigger wired during setup — the same trigger survives across `Remove Automated Triggers` (time-based only).
- **Version check** — `SCRIPT_VERSION` constant in `Config.gs`. Scripts check the repo's `version.json` on GitHub once per 24h (piggybacked on `triggerDataContinue`) and write results to the Config sheet with color-coded cells (green = current, yellow = update available). Manual check via **iiQ Assets > Setup > Check for Updates**.

### Changed
- **Trigger-safety refactor** — `requireNoTriggers` and `removeAllProjectTriggers` now operate on time-based (CLOCK) triggers only. The installable onEdit trigger for IndividualLookup is preserved across destructive operations and "Remove Automated Triggers". `View Trigger Status` surfaces time-based vs other triggers separately.

### Upgrade Notes
1. Update all `.gs` files from the `scripts/` directory.
2. Run **iiQ Assets > Setup > Setup Spreadsheet** *only if* starting fresh. Existing sheets can skip this — the new Config rows (`SCRIPT_VERSION`, `LATEST_VERSION`, `VERSION_CHECK_DATE`) will appear on the next Setup Spreadsheet run or can be added manually.
3. Add the new analytics sheet via **iiQ Assets > Analytics Sheets > People > Individual Lookup**. Authorize the onEdit trigger when prompted.
