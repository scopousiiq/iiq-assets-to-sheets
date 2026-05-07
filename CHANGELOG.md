# Changelog — iiQ Assets to Sheets

All notable changes to this project are documented here.

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
