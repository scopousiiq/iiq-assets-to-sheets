# Changelog — iiQ Assets to Sheets

All notable changes to this project are documented here.

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
