# Changelog — iiQ Assets to Sheets

All notable changes to this project are documented here.

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
