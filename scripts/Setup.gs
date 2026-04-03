/**
 * iiQ Asset Reporting - Setup
 * Creates all sheets with headers, formulas, and formatting.
 */

function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  setupConfigSheet(ss);
  setupAssetDataSheet(ss);
  setupLocationsSheet(ss);
  setupStatusTypesSheet(ss);
  setupLocationEnrollmentSheet(ss);
  setupLogsSheet(ss);
  setupInstructionsSheet(ss);

  // Default analytics sheets
  setupFleetSummarySheet(ss);
  setupLocationSummarySheet(ss);
  setupModelBreakdownSheet(ss);
  setupAgingAnalysisSheet(ss);
  setupBudgetPlanningSheet(ss);
  setupServiceImpactSheet(ss);
  setupAssignmentOverviewSheet(ss);
  setupStatusOverviewSheet(ss);

  SpreadsheetApp.getUi().alert('Setup complete! Fill in the Config sheet, then load reference data.');
}

function deleteSheetIfExists(ss, name) {
  const sheet = ss.getSheetByName(name);
  if (sheet) ss.deleteSheet(sheet);
}

/**
 * Returns an existing sheet after clearing its data, or creates a new one.
 * On regeneration (sheet exists): clears data below headers, skips delete/create overhead.
 * On first install (sheet missing): creates fresh sheet with headers, formatting, and tab color.
 *
 * @param {Spreadsheet} ss
 * @param {string} name - Sheet name
 * @param {string[][]} headers - Header row values, e.g. [['Col A', 'Col B']]
 * @param {string} tabColor - Hex color for tab
 * @param {Object} [opts] - Optional: { frozenRows: 1, columnWidths: {1: 280, 2: 180} }
 * @returns {{ sheet: SpreadsheetApp.Sheet, isNew: boolean }}
 */
function getOrCreateSheet(ss, name, headers, tabColor, opts) {
  const existing = ss.getSheetByName(name);
  if (existing) {
    // Regenerate path: clear data below header row, preserving structure
    if (existing.getLastRow() > 1) {
      existing.getRange(2, 1, existing.getLastRow() - 1, existing.getLastColumn()).clearContent();
    }
    return { sheet: existing, isNew: false };
  }
  // Create path: full setup
  const sheet = ss.insertSheet(name);
  sheet.getRange(1, 1, 1, headers[0].length).setValues(headers).setFontWeight('bold');
  const frozen = (opts && opts.frozenRows) || 1;
  sheet.setFrozenRows(frozen);
  if (opts && opts.columnWidths) {
    Object.entries(opts.columnWidths).forEach(([col, width]) => sheet.setColumnWidth(Number(col), width));
  }
  sheet.setTabColor(tabColor);
  return { sheet, isNew: true };
}

// =============================================================================
// DATA SHEETS
// =============================================================================

function setupConfigSheet(ss) {
  deleteSheetIfExists(ss, 'Config');
  const sheet = ss.insertSheet('Config');
  sheet.getRange(1, 1, 1, 2).setValues([['Key', 'Value']]).setFontWeight('bold');

  const configRows = [
    ['API_BASE_URL', 'https://YOUR-DISTRICT.incidentiq.com'],
    ['BEARER_TOKEN', ''],
    ['SITE_ID', ''],
    ['PAGE_SIZE', '100'],
    ['THROTTLE_MS', '1000'],
    ['ASSET_BATCH_SIZE', '500'],
    ['STUDENT_ROLE_ID', ''],
    ['REPLACEMENT_AGE_YEARS', '4'],
    ['NEXT_SCHOOL_YEAR_START', '2026-07-01'],
    ['', ''],
    ['--- Progress (auto-managed) ---', ''],
    ['ASSET_LAST_PAGE', '-1'],
    ['ASSET_TOTAL_PAGES', '-1'],
    ['ASSET_COMPLETE', 'FALSE'],
    ['LAST_REFRESH_DATE', ''],
  ];

  sheet.getRange(2, 1, configRows.length, 2).setValues(configRows);
  sheet.setColumnWidth(1, 250);
  sheet.setColumnWidth(2, 400);
  sheet.setTabColor('#febb12');
}

function setupAssetDataSheet(ss) {
  deleteSheetIfExists(ss, 'AssetData');
  const sheet = ss.insertSheet('AssetData');

  // Headers
  sheet.getRange(1, 1, 1, ASSET_TOTAL_COLS).setValues([ASSET_HEADERS]).setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);
  sheet.setTabColor('#22b2a3');
}

/**
 * Apply ARRAYFORMULA-based formulas to AssetData sheet.
 * Uses a single formula per column that spills to all data rows automatically.
 * Call after data loading completes.
 */
function applyAssetFormulas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('AssetData');
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  // Clear any existing per-row formulas in the formula columns before applying ARRAYFORMULAs
  const numRows = lastRow - 1;
  sheet.getRange(2, 26, numRows, 3).clearContent();

  // Column Z (26): AgeDays — days since PurchasedDate (col N), falls back to CreatedDate (col Q)
  // Uses (N="")*(Q="") instead of AND() which doesn't work in ARRAYFORMULA
  sheet.getRange(2, 26).setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",IF((N2:N="")*(Q2:Q=""),"",INT(TODAY()-IF(N2:N<>"",N2:N,Q2:Q)))))'
  );

  // Column AA (27): AgeYears — AgeDays / 365.25
  sheet.getRange(2, 27).setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",IF(Z2:Z="","",ROUND(Z2:Z/365.25,1))))'
  );

  // Column AB (28): WarrantyStatus — based on WarrantyExpDate (col O)
  sheet.getRange(2, 28).setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",IF(O2:O="","None",IF(O2:O<TODAY(),"Expired",IF(O2:O<TODAY()+90,"Expiring","Active")))))'
  );

  logOperation('Formulas', 'COMPLETE', `Applied ARRAYFORMULA to ${numRows} rows`);
}

function setupLocationsSheet(ss) {
  deleteSheetIfExists(ss, 'Locations');
  const sheet = ss.insertSheet('Locations');
  sheet.getRange(1, 1, 1, 5).setValues([['LocationId', 'Name', 'Abbreviation', 'LocationType', 'Address']]).setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.setTabColor('#22b2a3');
}

function setupStatusTypesSheet(ss) {
  deleteSheetIfExists(ss, 'StatusTypes');
  const sheet = ss.insertSheet('StatusTypes');
  sheet.getRange(1, 1, 1, 4).setValues([['StatusTypeId', 'Name', 'IsRetired', 'SortOrder']]).setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.setTabColor('#22b2a3');
}

function setupLocationEnrollmentSheet(ss) {
  deleteSheetIfExists(ss, 'LocationEnrollment');
  const sheet = ss.insertSheet('LocationEnrollment');
  sheet.getRange(1, 1, 1, 6).setValues([[
    'LocationId', 'LocationName', 'LocationType', 'TotalStudents', 'StudentsWithDevices', 'DeviceCoverage%'
  ]]).setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.getRange('F:F').setNumberFormat('0.0%');
  sheet.setTabColor('#22b2a3');
}

function setupLogsSheet(ss) {
  deleteSheetIfExists(ss, 'Logs');
  const sheet = ss.insertSheet('Logs');
  sheet.getRange(1, 1, 1, 4).setValues([['Timestamp', 'Operation', 'Status', 'Details']]).setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(4, 500);
  sheet.setTabColor('#febb12');
}

function setupInstructionsSheet(ss) {
  deleteSheetIfExists(ss, 'Instructions');
  const sheet = ss.insertSheet('Instructions');
  ss.setActiveSheet(sheet);
  ss.moveActiveSheet(1);

  sheet.setColumnWidth(1, 800);

  const content = [
    ['iiQ ASSET REPORTING - SETUP AND USAGE GUIDE'],                                      // 1
    [''],                                                                                   // 2
    ['═══════════════════════════════════════════════════════════════════════════════'],      // 3
    ['OVERVIEW'],                                                                           // 4
    ['═══════════════════════════════════════════════════════════════════════════════'],      // 5
    [''],                                                                                   // 6
    ['This spreadsheet extracts asset inventory data from Incident IQ (iiQ) for'],          // 7
    ['reporting, device lifecycle planning, and dashboard consumption. Data is loaded'],     // 8
    ['via Google Apps Script and kept current with automated daily refresh.'],               // 9
    [''],                                                                                   // 10
    ['Use Cases:'],                                                                         // 11
    ['  • Asset inventory by school/location'],                                             // 12
    ['  • Device model and manufacturer breakdown'],                                        // 13
    ['  • Fleet age analysis and replacement planning'],                                    // 14
    ['  • Warranty tracking and budget forecasting'],                                       // 15
    ['  • Service impact analysis (which models create the most tickets)'],                 // 16
    ['  • Device assignment and utilization tracking'],                                     // 17
    ['  • Student enrollment and device coverage per school'],                              // 18
    [''],                                                                                   // 19
    ['Data Flow: iiQ API → Google Apps Script → This Spreadsheet → Looker Studio / Power BI'], // 20
    [''],                                                                                   // 21
    ['NOTE: Deleted assets in iiQ are automatically excluded. Only active (non-deleted)'],  // 22
    ['assets are loaded. The iiQ API filters them out by default.'],                        // 23
    [''],                                                                                   // 24
    [''],                                                                                   // 25
    ['═══════════════════════════════════════════════════════════════════════════════'],      // 26
    ['INITIAL SETUP'],                                                                      // 27
    ['═══════════════════════════════════════════════════════════════════════════════'],      // 28
    [''],                                                                                   // 29
    ['1. SETUP SPREADSHEET'],                                                               // 30
    ['   • Menu: iiQ Assets > Setup > Setup Spreadsheet'],                                  // 31
    ['   • Creates all data sheets, 8 default analytics sheets, and configures formulas'],   // 32
    ['   • WARNING: This is a full reset — all existing data and config will be deleted'],   // 33
    [''],                                                                                   // 34
    ['2. CONFIGURE API CREDENTIALS (Config sheet)'],                                        // 35
    ['   • API_BASE_URL: Your iiQ instance (e.g., https://district.incidentiq.com)'],       // 36
    ['   • BEARER_TOKEN: JWT token from iiQ (Admin > Integrations > API)'],                 // 37
    ['   • SITE_ID: Optional — only needed for multi-site instances'],
    ['   • REPLACEMENT_AGE_YEARS: Device age threshold in years (default 4)'],
    ['   • NEXT_SCHOOL_YEAR_START: Target date for replacement planning (default 2026-07-01)'],
    [''],
    ['3. VERIFY CONFIGURATION'],                                                            // 40
    ['   • Menu: iiQ Assets > Setup > Verify Configuration'],                               // 41
    ['   • Confirms API connection and shows total available assets'],                       // 42
    ['   • Fix any issues reported before proceeding'],                                     // 43
    [''],                                                                                   // 44
    ['4. START LOADING DATA'],                                                              // 45
    ['   • Menu: iiQ Assets > Asset Data > Load / Resume Assets'],                          // 46
    ['   • Reference data (locations, status types) loads automatically on first run'],      // 47
    ['   • Script runs for ~5.5 minutes then pauses (Apps Script time limit)'],              // 48
    ['   • You will see "Loading Complete" when all assets are loaded'],                     // 49
    ['   • For small districts (< 2,000 assets): may finish in one run'],                   // 50
    ['   • For large districts (10,000+): continue to step 5 and let automation finish'],    // 51
    [''],                                                                                   // 52
    ['5. SETUP AUTOMATED TRIGGERS'],                                                        // 53
    ['   • Menu: iiQ Assets > Setup > Setup Automated Triggers'],                           // 54
    ['   • The 10-minute trigger will automatically continue loading if not yet complete'],   // 55
    ['   • Once loading finishes, formulas are applied automatically'],                      // 56
    ['   • Daily refresh and weekly reload keep data current going forward'],                // 57
    [''],                                                                                   // 58
    ['   After triggers are set up, you can close the spreadsheet and come back later.'],    // 59
    ['   Check progress anytime: iiQ Assets > Asset Data > Show Status'],                   // 60
    [''],                                                                                   // 61
    ['   APPLY FORMULAS (only if you completed loading manually without triggers):'],        // 62
    ['   • Menu: iiQ Assets > Asset Data > Apply Formulas'],                                // 63
    ['   • Adds calculated columns: AgeDays, AgeYears, WarrantyStatus'],                    // 64
    [''],                                                                                   // 65
    ['6. LOAD STUDENT ENROLLMENT (Optional)'],                                              // 66
    ['   • Menu: iiQ Assets > Load Reference Data > View Available Roles'],                  // 67
    ['   • Find the student role and copy its RoleId into STUDENT_ROLE_ID on the Config sheet'], // 68
    ['   • Menu: iiQ Assets > Load Reference Data > Refresh Location Enrollment'],           // 69
    ['   • Counts students and device coverage per location (may take several runs for large districts)'], // 70
    ['   • Enrollment data feeds into LocationSummary analytics'],                           // 71
    [''],                                                                                   // 72
    [''],                                                                                   // 73
    ['═══════════════════════════════════════════════════════════════════════════════'],      // 74
    ['AUTOMATED TRIGGERS (Recommended)'],                                                   // 75
    ['═══════════════════════════════════════════════════════════════════════════════'],      // 76
    [''],                                                                                   // 77
    ['EASY SETUP: Use iiQ Assets > Setup > Setup Automated Triggers'],                      // 78
    ['This creates all recommended triggers automatically.'],                                // 79
    [''],                                                                                   // 80
    ['MANUAL SETUP: Extensions > Apps Script > Triggers (clock icon)'],                      // 81
    [''],                                                                                   // 82
    ['| Function                    | Schedule         | Purpose                              |'], // 83
    ['|---------------------------- |------------------|--------------------------------------|'], // 84
    ['| triggerDataContinue         | Every 10 min     | Continue any in-progress initial load |'], // 85
    ['| triggerDailyRefresh         | Daily 3 AM       | Incremental refresh (changed assets)  |'], // 86
    ['| triggerWeeklyFullRefresh    | Weekly Sun 2 AM  | Full reload + reference data refresh  |'], // 87
    [''],                                                                                   // 88
    ['About triggerDataContinue (the "keep things moving" trigger):'],                       // 89
    ['• If initial load is not complete → continues loading the next batch of assets'],      // 90
    ['• If enrollment loading is incomplete → continues counting locations'],                // 91
    ['• If everything is complete → does nothing (safe to leave enabled permanently)'],      // 92
    [''],                                                                                   // 93
    ['About triggerDailyRefresh (the "keep data fresh" trigger):'],                          // 94
    ['• Queries iiQ for assets modified since the last refresh'],                            // 95
    ['• Updates existing rows in-place by AssetId (no duplicate data)'],                     // 96
    ['• Appends any new assets not previously seen'],                                        // 97
    ['• Formula columns auto-expand to cover new rows (no reapply needed)'],                 // 98
    ['• Typical run: fetches only the 100-500 assets that changed, not the full inventory'], // 99
    [''],                                                                                   // 100
    ['Data Freshness with this schedule:'],                                                  // 101
    ['• Asset changes: reflected by next morning (3 AM refresh)'],                           // 102
    ['• New assets: appear by next morning'],                                                // 103
    ['• Deleted assets: automatically excluded by iiQ API (never downloaded)'],              // 104
    ['• Edge cases (un-deleted assets, data drift): caught by weekly full reload (Sunday 2 AM)'], // 105
    ['• On-demand: Use "Refresh Changed Assets" for immediate update'],                      // 106
    [''],                                                                                   // 107
    [''],                                                                                   // 108
    ['═══════════════════════════════════════════════════════════════════════════════'],      // 109
    ['SHEETS REFERENCE'],                                                                   // 110
    ['═══════════════════════════════════════════════════════════════════════════════'],      // 111
    [''],                                                                                   // 112
    ['DATA SHEETS (populated by scripts):'],                                                 // 113
    [''],                                                                                   // 114
    ['• AssetData (28 columns: 25 from API + 3 calculated)'],                                // 115
    ['  Main asset inventory. Columns:'],                                                    // 116
    ['  - Identity: AssetId, AssetTag, Name, SerialNumber'],                                 // 117
    ['  - Device: ModelName, ManufacturerName, CategoryName'],                               // 118
    ['  - Location: LocationId, LocationName, LocationType'],                                // 119
    ['  - Assignment: OwnerId, OwnerName'],                                                  // 120
    ['  - Status: StatusName (Active, Retired, In Storage, etc.)'],                          // 121
    ['  - Purchase: PurchasedDate, WarrantyExpDate, PurchasePrice'],                          // 122
    ['  - Tracking: CreatedDate, ModifiedDate'],                                             // 123
    ['  - Owner Detail: OwnerRoleName, OwnerGrade, OwnerLocationId'],                        // 124
    ['  - Storage: StorageLocationName, StorageUnitNumber, DeployedDate'],                    // 125
    ['  - Service: OpenTickets'],                                                            // 126
    ['  - Calculated: AgeDays, AgeYears, WarrantyStatus (ARRAYFORMULA columns)'],            // 127
    [''],                                                                                   // 128
    ['  NOTE ON DEVICE AGE: AgeDays and AgeYears are calculated from PurchasedDate when'],   // 129
    ['  available. If PurchasedDate is empty (common when districts do not track purchase'],  // 130
    ['  dates in iiQ), CreatedDate is used as a fallback. CreatedDate reflects when the'],    // 131
    ['  asset record was added to iiQ, which is a reasonable proxy for device age. All'],     // 132
    ['  analytics sheets that reference device age use this same fallback logic.'],           // 133
    [''],                                                                                   // 134
    ['• Locations'],                                                                         // 135
    ['  Location directory loaded from iiQ (ID, Name, Abbreviation, Type, Address).'],       // 136
    [''],                                                                                   // 137
    ['• StatusTypes'],                                                                       // 138
    ['  Asset status type directory (ID, Name, IsRetired, SortOrder).'],                     // 139
    [''],                                                                                   // 140
    ['• LocationEnrollment (optional — requires STUDENT_ROLE_ID in Config)'],                // 141
    ['  Student count and device coverage per location.'],                                   // 142
    ['  Columns: LocationId, LocationName, LocationType, TotalStudents,'],                   // 143
    ['  StudentsWithDevices, DeviceCoverage%.'],                                             // 144
    [''],                                                                                   // 145
    ['• Logs'],                                                                              // 146
    ['  Operation logs for troubleshooting. Auto-pruned to 500 rows.'],                      // 147
    [''],                                                                                   // 148
    [''],                                                                                   // 149
    ['ANALYTICS SHEETS (formula-based, auto-calculate from AssetData):'],
    [''],
    ['All analytics sheets are organized into four categories. Default sheets (marked ★)'],
    ['are created automatically by Setup Spreadsheet. Optional sheets can be installed'],
    ['individually from the categorized Analytics Sheets menu.'],
    [''],
    [''],
    ['  FLEET OPERATIONS — Assignment, status, and device readiness'],
    [''],
    ['  ★ AssignmentOverview — Device utilization'],
    ['    Per-location: total assets, assigned, unassigned, assignment rate, active.'],
    ['    Identifies where surplus/idle inventory is sitting.'],
    [''],
    ['  ★ StatusOverview — Status breakdown'],
    ['    Count and percentage of total by each asset status (Active, Retired, etc.).'],
    [''],
    ['  • DeviceReadiness — What is actually deployable at each school right now?'],
    ['    Per-location totals for deployable, in repair, lost/stolen, retired, and readiness rate.'],
    [''],
    ['  • SpareAssets — Do I have enough working spares at each school?'],
    ['    Spare counts vs total assigned, with spare ratio per location.'],
    [''],
    ['  • LostStolenRate — Which schools are losing devices?'],
    ['    Lost/stolen counts and rates by location.'],
    [''],
    ['  • ModelFragmentation — Which schools are a patchwork of device models?'],
    ['    Count of distinct models per location. High fragmentation = harder support.'],
    [''],
    ['  • UnassignedInventory — Where is idle inventory sitting?'],
    ['    Unassigned counts by location with active count, average age, and estimated value.'],
    [''],
    [''],
    ['  SERVICE & RELIABILITY — Tickets, break rates, and problem models'],
    [''],
    ['  ★ ServiceImpact — Model reliability analysis'],
    ['    Models ranked by tickets-per-device ratio. Shows device count, total open'],
    ['    tickets, tickets/device, and average age. Identifies unreliable models'],
    ['    that should be phased out.'],
    [''],
    ['  • BreakRate — Which individual devices and models have the most tickets?'],
    ['    Per-device and per-model ticket counts for identifying chronic problem units.'],
    [''],
    ['  • HighTicketLocations — Schools ranked by tickets-per-device ratio'],
    ['    Locations with the highest service burden relative to device count.'],
    [''],
    [''],
    ['  BUDGET & PLANNING — Replacement costs, aging, and lifecycle forecasting'],
    [''],
    ['  ★ BudgetPlanning — Replacement cost estimates'],
    ['    Per-location: warranty expired, warranty expiring, devices older than'],
    ['    REPLACEMENT_AGE_YEARS (default 4), average purchase price, and estimated'],
    ['    replacement cost. Excludes retired devices (already out of service).'],
    [''],
    ['  ★ AgingAnalysis — Fleet age distribution'],
    ['    Year cohorts showing device counts, active/retired split, warranty expired,'],
    ['    average open tickets, and total value per year. Uses PurchasedDate when'],
    ['    available, falls back to CreatedDate. Answers "when is the replacement cliff?"'],
    [''],
    ['  • ReplacementPlanning — What do I need to buy before next school year?'],
    ['    Per-location counts of devices currently over the age threshold, projected to'],
    ['    cross it by NEXT_SCHOOL_YEAR_START, and estimated replacement cost.'],
    [''],
    ['  • ReplacementForecast — Projected replacement volume/cost by year'],
    ['    Future replacement years, device counts, average price, and estimated cost.'],
    [''],
    ['  • WarrantyTimeline — Upcoming warranty expirations by quarter'],
    ['    Warranty expiration schedule for proactive planning.'],
    [''],
    ['  • DeviceLifecycle — Average lifespan at retirement by model'],
    ['    How long devices actually last, by model.'],
    [''],
    [''],
    ['  FLEET COMPOSITION — Inventory breakdown by model, location, and category'],
    [''],
    ['  ★ FleetSummary — Executive KPI dashboard'],
    ['    Top-line metrics: total assets, fleet value, average age, warranty coverage,'],
    ['    open tickets, assignment rate. The "board meeting" sheet.'],
    [''],
    ['  ★ LocationSummary — Assets per school'],
    ['    Total, active, retired, warranty expired, average age, and enrollment data'],
    ['    by location. Includes student count and device coverage if enrollment is loaded.'],
    ['    Sorted by total assets descending.'],
    [''],
    ['  ★ ModelBreakdown — Device model inventory'],
    ['    Total, active, retired, and average age by model/manufacturer.'],
    ['    Identifies which models dominate the fleet.'],
    [''],
    ['  • LocationModelBreakdown — What models are at each school?'],
    ['    Flat location/model breakdown with counts, active vs retired, and average age.'],
    [''],
    ['  • LocationModelFiltered — Show me one school\'s model mix (dropdown-driven)'],
    ['    Single-location model breakdown with a dropdown selector.'],
    [''],
    ['  • CategoryBreakdown — Inventory by device category (Chromebook, laptop, etc.)'],
    ['    Counts, active vs retired, average age, and total value by category.'],
    [''],
    ['  • ManufacturerSummary — Device count, age, warranty, tickets by manufacturer'],
    ['    Which vendors are you invested in, and how are their devices performing?'],
    [''],
    [''],
    ['All analytics sheets can be regenerated from the menu (per-category, all default,'],
    ['or all analytics). Regeneration refreshes formulas without deleting the sheet.'],
    [''],
    [''],                                                                                   // 200
    ['═══════════════════════════════════════════════════════════════════════════════'],      // 201
    ['MENU REFERENCE'],                                                                     // 202
    ['═══════════════════════════════════════════════════════════════════════════════'],      // 203
    [''],                                                                                   // 204
    ['iiQ Assets > Setup'],                                                                  // 205
    ['  • Setup Spreadsheet — FULL RESET: deletes all sheets and recreates from scratch'],   // 206
    ['  • Verify Configuration — Check API settings and test connection'],                   // 207
    ['  • Setup Automated Triggers — Create all recommended triggers'],                      // 208
    ['  • View Trigger Status — Show installed triggers'],                                   // 209
    ['  • Remove Automated Triggers — Remove all triggers (required before destructive ops)'], // 210
    [''],                                                                                   // 211
    ['iiQ Assets > Load Reference Data'],                                                    // 212
    ['  • Refresh Locations — Reload location directory'],                                   // 213
    ['  • Refresh Status Types — Reload asset status types'],                                // 214
    ['  • Refresh Location Enrollment — Count students and device coverage per location'],    // 215
    ['  • View Available Roles — Show iiQ roles (find your STUDENT_ROLE_ID)'],               // 216
    ['  • Refresh All Reference Data — Reload all reference data'],                          // 217
    [''],                                                                                   // 218
    ['iiQ Assets > Asset Data'],                                                             // 219
    ['  • Load / Resume Assets — Load or continue loading asset data (runs ~5.5 min per batch)'], // 220
    ['  • Refresh Changed Assets — On-demand incremental update (ModifiedDate filter)'],     // 221
    ['  • Apply Formulas — Add/refresh calculated columns (AgeDays, AgeYears, WarrantyStatus)'], // 222
    ['  • Show Status — Display loading progress and last refresh date'],                    // 223
    ['  • Remove Duplicates — Keep the newest row per AssetId and reapply formulas'],
    ['  • Clear Data + Reset Progress — Clear all asset data (requires triggers removed)'],
    ['  • Full Reload — Clear and reload from scratch (requires triggers removed)'],
    [''],
    ['iiQ Assets > Analytics Sheets'],
    ['  Organized into four category submenus:'],
    [''],
    ['  Fleet Operations'],
    ['    • ★ AssignmentOverview, ★ StatusOverview (default)'],
    ['    • DeviceReadiness, SpareAssets, LostStolenRate, ModelFragmentation,'],
    ['      UnassignedInventory (optional — install individually)'],
    ['    • Regenerate Fleet Operations — Refresh all installed Fleet Operations sheets'],
    [''],
    ['  Service & Reliability'],
    ['    • ★ ServiceImpact (default)'],
    ['    • BreakRate, HighTicketLocations (optional — install individually)'],
    ['    • Regenerate Service & Reliability — Refresh all installed sheets in this category'],
    [''],
    ['  Budget & Planning'],
    ['    • ★ BudgetPlanning, ★ AgingAnalysis (default)'],
    ['    • ReplacementPlanning, ReplacementForecast, WarrantyTimeline,'],
    ['      DeviceLifecycle (optional — install individually)'],
    ['    • Regenerate Budget & Planning — Refresh all installed sheets in this category'],
    [''],
    ['  Fleet Composition'],
    ['    • ★ FleetSummary, ★ LocationSummary, ★ ModelBreakdown (default)'],
    ['    • LocationModelBreakdown, LocationModelFiltered, CategoryBreakdown,'],
    ['      ManufacturerSummary (optional — install individually)'],
    ['    • Regenerate Fleet Composition — Refresh all installed sheets in this category'],
    [''],
    ['  Regenerate All Default (★) — Rebuild all 8 default analytics sheets'],
    ['  Regenerate All Analytics — Rebuild all installed analytics sheets (default + optional)'],
    [''],                                                                                   // 234
    [''],                                                                                   // 235
    ['═══════════════════════════════════════════════════════════════════════════════'],      // 236
    ['TROUBLESHOOTING'],                                                                    // 237
    ['═══════════════════════════════════════════════════════════════════════════════'],      // 238
    [''],                                                                                   // 239
    ['"API configuration missing"'],                                                         // 240
    ['  → Check Config sheet has API_BASE_URL and BEARER_TOKEN filled in'],                  // 241
    [''],                                                                                   // 242
    ['"Rate limited" or 429 errors'],                                                        // 243
    ['  → Increase THROTTLE_MS in Config (default 1000ms)'],                                 // 244
    ['  → Script automatically retries with exponential backoff'],                           // 245
    [''],                                                                                   // 246
    ['Loading seems stuck or stops partway'],                                                // 247
    ['  → Check Logs sheet for errors'],                                                     // 248
    ['  → Run "Show Status" to see progress'],                                               // 249
    ['  → Run "Load / Resume Assets" again — it resumes from where it left off'],            // 250
    ['  → Large districts may need 3-5 runs to complete initial load'],                      // 251
    [''],                                                                                   // 252
    ['Formula errors (#REF, #VALUE) in analytics sheets'],                                   // 253
    ['  → Ensure AssetData has data loaded'],                                                // 254
    ['  → Run "Apply Formulas" after loading completes'],                                    // 255
    ['  → Try "Regenerate All Analytics" to rebuild analytics sheets'],                      // 256
    [''],                                                                                   // 257
    ['Trigger not running'],                                                                 // 258
    ['  → Check Apps Script > Triggers for errors'],                                         // 259
    ['  → Use iiQ Assets > Setup > View Trigger Status for diagnostics'],                    // 260
    ['  → Remove and re-add triggers if needed'],                                            // 261
    [''],                                                                                   // 262
    ['"Another operation is in progress"'],                                                  // 263
    ['  → A script is already running (menu action or trigger)'],                            // 264
    ['  → Wait a few minutes for it to complete, then try again'],                           // 265
    [''],                                                                                   // 266
    ['"Remove triggers first"'],                                                             // 267
    ['  → Destructive operations (Full Reload, Clear Data) require triggers removed'],       // 268
    ['  → Use iiQ Assets > Setup > Remove Automated Triggers first'],                        // 269
    ['  → Re-add triggers after the operation completes'],                                   // 270
    [''],                                                                                   // 271
    ['"STUDENT_ROLE_ID not configured"'],                                                    // 272
    ['  → Use iiQ Assets > Load Reference Data > View Available Roles'],                     // 273
    ['  → Copy the student role ID into the Config sheet under STUDENT_ROLE_ID'],            // 274
    [''],                                                                                   // 275
    ['Refresh shows 0 updated, 0 new'],                                                      // 276
    ['  → Normal if no assets have been modified since last refresh'],                       // 277
    ['  → Check LAST_REFRESH_DATE in Config sheet'],                                         // 278
    ['  → If date looks wrong, run "Full Reload" to reset'],                                 // 279
    [''],                                                                                   // 280
    [''],                                                                                   // 281
    ['═══════════════════════════════════════════════════════════════════════════════'],      // 282
    ['DASHBOARD INTEGRATION'],                                                               // 283
    ['═══════════════════════════════════════════════════════════════════════════════'],      // 284
    [''],                                                                                   // 285
    ['LOOKER STUDIO (Recommended):'],                                                        // 286
    ['1. Go to lookerstudio.google.com > Create > Report'],                                  // 287
    ['2. Add Google Sheets connector > Select this spreadsheet > AssetData sheet'],           // 288
    ['3. Add analytics sheets as additional data sources if needed'],                         // 289
    ['4. Formula columns (AgeYears, WarrantyStatus) are ready for filtering and grouping'],  // 290
    [''],                                                                                   // 291
    ['Recommended Looker Studio visualizations:'],                                           // 292
    ['  • Fleet age distribution (bar chart from AgeYears)'],                                // 293
    ['  • Assets by location (table or map from LocationName)'],                             // 294
    ['  • Warranty status breakdown (pie chart from WarrantyStatus)'],                       // 295
    ['  • Model inventory (table from ModelName with counts)'],                              // 296
    ['  • Budget exposure (table from BudgetPlanning sheet)'],                               // 297
    ['  • Student device coverage by school (from LocationEnrollment or LocationSummary)'],  // 298
    [''],                                                                                   // 299
    ['POWER BI:'],                                                                           // 300
    ['1. In Power BI Desktop: Get Data > Web'],                                              // 301
    ['2. Use the shareable link for each sheet (File > Share > Publish to web)'],             // 302
    ['3. Or use the Google Sheets connector if available'],                                   // 303
    ['4. Set up scheduled refresh in Power BI Service'],                                     // 304
    [''],                                                                                   // 305
    ['Data is refreshed daily at 3 AM, so dashboards can refresh on a similar schedule.'],   // 306
    ['Weekly full refresh starts Sunday at 2 AM and may continue in 10-minute batches.'],
    [''],                                                                                   // 308
    [''],                                                                                   // 309
    ['═══════════════════════════════════════════════════════════════════════════════'],      // 310
    ['SUPPORT'],                                                                             // 311
    ['═══════════════════════════════════════════════════════════════════════════════'],      // 312
    [''],                                                                                   // 313
    ['GitHub: https://github.com/scopousiiq/iiq-assets-to-sheets'],                          // 314
    ['For issues or feature requests, check the Logs sheet first for error details.'],       // 315
    [''],                                                                                   // 316
    ['Last updated: ' + new Date().toISOString().split('T')[0]],                             // 317
  ];

  // Write content + wrapping in one pass
  sheet.getRange(1, 1, content.length, 1).setValues(content).setWrap(true);

  // Format title
  sheet.getRange(1, 1).setFontSize(16).setFontWeight('bold').setFontColor('#1a73e8');

  const sectionTitles = new Set([
    'OVERVIEW',
    'INITIAL SETUP',
    'AUTOMATED TRIGGERS (Recommended)',
    'SHEETS REFERENCE',
    'MENU REFERENCE',
    'TROUBLESHOOTING',
    'DASHBOARD INTEGRATION',
    'SUPPORT'
  ]);
  const sectionRanges = [];
  const dividerRanges = [];

  content.forEach((row, index) => {
    const value = row[0];
    if (sectionTitles.has(value)) sectionRanges.push('A' + (index + 1));
    if (typeof value === 'string' && /^═+$/.test(value)) dividerRanges.push('A' + (index + 1));
  });

  if (sectionRanges.length) {
    sheet.getRangeList(sectionRanges).setFontWeight('bold').setFontColor('#1a73e8');
  }
  if (dividerRanges.length) {
    sheet.getRangeList(dividerRanges).setFontColor('#dadce0');
  }

  sheet.setFrozenRows(1);
  sheet.setTabColor('#365c96');
}

// =============================================================================
// DEFAULT ANALYTICS SHEETS
// =============================================================================

function setupLocationSummarySheet(ss) {
  const { sheet, isNew } = getOrCreateSheet(ss, 'LocationSummary', [[
    'Location', 'Total Assets', 'Active', 'Retired', 'Warranty Expired', 'Avg Age (Years)',
    'Assigned', 'Total Students', 'Students w/ Devices', 'Device Coverage %'
  ]], '#365c96');

  const formula = '=LET(\n' +
    '  locs, UNIQUE(FILTER(AssetData!I2:I, AssetData!I2:I<>"")),\n' +
    '  total, BYROW(locs, LAMBDA(loc, COUNTIF(AssetData!I:I, loc))),\n' +
    '  active, BYROW(locs, LAMBDA(loc, COUNTIFS(AssetData!I:I, loc, AssetData!M:M, "<>Retired"))),\n' +
    '  retired, BYROW(locs, LAMBDA(loc, COUNTIFS(AssetData!I:I, loc, AssetData!M:M, "Retired"))),\n' +
    '  warr_exp, BYROW(locs, LAMBDA(loc, COUNTIFS(AssetData!I:I, loc, AssetData!AB:AB, "Expired"))),\n' +
    '  avg_age, BYROW(locs, LAMBDA(loc, IFERROR(AVERAGEIFS(AssetData!AA:AA, AssetData!I:I, loc), 0))),\n' +
    '  assigned, BYROW(locs, LAMBDA(loc, COUNTIFS(AssetData!I:I, loc, AssetData!K:K, "<>"))),\n' +
    '  students, BYROW(locs, LAMBDA(loc, IFERROR(INDEX(FILTER(LocationEnrollment!D:D, LocationEnrollment!B:B=loc), 1), 0))),\n' +
    '  with_dev, BYROW(locs, LAMBDA(loc, IFERROR(INDEX(FILTER(LocationEnrollment!E:E, LocationEnrollment!B:B=loc), 1), 0))),\n' +
    '  coverage, BYROW(locs, LAMBDA(loc, IFERROR(\n' +
    '    INDEX(FILTER(LocationEnrollment!E:E, LocationEnrollment!B:B=loc), 1)\n' +
    '    / INDEX(FILTER(LocationEnrollment!D:D, LocationEnrollment!B:B=loc), 1), 0))),\n' +
    '  IFERROR(SORT(HSTACK(locs, total, active, retired, warr_exp, avg_age, assigned, students, with_dev, coverage), 2, FALSE),\n' +
    '    HSTACK(locs, total, active, retired, warr_exp, avg_age, assigned, students, with_dev, coverage))\n' +
    ')';

  sheet.getRange(2, 1).setFormula(formula);
  if (isNew) sheet.getRange('J:J').setNumberFormat('0.0%');
}

function setupModelBreakdownSheet(ss) {
  const { sheet } = getOrCreateSheet(ss, 'ModelBreakdown', [[
    'Model', 'Manufacturer', 'Total', 'Active', 'Retired', 'Avg Age (Years)'
  ]], '#365c96');

  const formula = '=LET(\n' +
    '  models, UNIQUE(FILTER(AssetData!E2:E, AssetData!E2:E<>"")),\n' +
    '  mfr, BYROW(models, LAMBDA(m, IFERROR(INDEX(FILTER(AssetData!F:F, AssetData!E:E=m), 1), ""))),\n' +
    '  total, BYROW(models, LAMBDA(m, COUNTIF(AssetData!E:E, m))),\n' +
    '  active, BYROW(models, LAMBDA(m, COUNTIFS(AssetData!E:E, m, AssetData!M:M, "<>Retired"))),\n' +
    '  retired, BYROW(models, LAMBDA(m, COUNTIFS(AssetData!E:E, m, AssetData!M:M, "Retired"))),\n' +
    '  avg_age, BYROW(models, LAMBDA(m, IFERROR(AVERAGEIFS(AssetData!AA:AA, AssetData!E:E, m), 0))),\n' +
    '  IFERROR(SORT(HSTACK(models, mfr, total, active, retired, avg_age), 3, FALSE),\n' +
    '    HSTACK(models, mfr, total, active, retired, avg_age))\n' +
    ')';

  sheet.getRange(2, 1).setFormula(formula);
}

function setupBudgetPlanningSheet(ss) {
  const { sheet } = getOrCreateSheet(ss, 'BudgetPlanning', [[
    'Location', 'Warranty Expired', 'Warranty Expiring', 'Older Than 4 Years', 'Avg Purchase Price', 'Est. Replacement Cost'
  ]], '#365c96');

  const formula = '=LET(\n' +
    '  locs, UNIQUE(FILTER(AssetData!I2:I, AssetData!I2:I<>"")),\n' +
    '  warr_expired, BYROW(locs, LAMBDA(loc, COUNTIFS(AssetData!I:I, loc, AssetData!AB:AB, "Expired", AssetData!M:M, "<>Retired"))),\n' +
    '  warr_expiring, BYROW(locs, LAMBDA(loc, COUNTIFS(AssetData!I:I, loc, AssetData!AB:AB, "Expiring", AssetData!M:M, "<>Retired"))),\n' +
    '  old_devices, BYROW(locs, LAMBDA(loc, COUNTIFS(AssetData!I:I, loc, AssetData!AA:AA, ">="&4, AssetData!M:M, "<>Retired"))),\n' +
    '  avg_price, BYROW(locs, LAMBDA(loc, IFERROR(AVERAGEIFS(AssetData!P:P, AssetData!I:I, loc, AssetData!P:P, ">0"), 0))),\n' +
    '  est_cost, BYROW(locs, LAMBDA(loc, IFERROR(\n' +
    '    (COUNTIFS(AssetData!I:I, loc, AssetData!AB:AB, "Expired", AssetData!M:M, "<>Retired")\n' +
    '     + COUNTIFS(AssetData!I:I, loc, AssetData!AB:AB, "Expiring", AssetData!M:M, "<>Retired"))\n' +
    '    * AVERAGEIFS(AssetData!P:P, AssetData!I:I, loc, AssetData!P:P, ">0"), 0))),\n' +
    '  IFERROR(SORT(HSTACK(locs, warr_expired, warr_expiring, old_devices, avg_price, est_cost), 6, FALSE),\n' +
    '    HSTACK(locs, warr_expired, warr_expiring, old_devices, avg_price, est_cost))\n' +
    ')';

  sheet.getRange(2, 1).setFormula(formula);
}

function setupFleetSummarySheet(ss) {
  const { sheet, isNew } = getOrCreateSheet(ss, 'FleetSummary', [['Metric', 'Value']], '#365c96',
    { columnWidths: { 1: 280, 2: 180 } });

  const kpis = [
    ['--- Inventory ---', ''],
    ['Total Assets', '=COUNTA(AssetData!A2:A)'],
    ['Active Assets', '=COUNTIFS(AssetData!A2:A,"<>",AssetData!M2:M,"<>Retired")'],
    ['Retired Assets', '=COUNTIFS(AssetData!A2:A,"<>",AssetData!M2:M,"Retired")'],
    ['Active Rate', '=IFERROR(COUNTIFS(AssetData!A2:A,"<>",AssetData!M2:M,"<>Retired")/COUNTA(AssetData!A2:A),0)'],
    ['Unique Models', '=COUNTA(UNIQUE(FILTER(AssetData!E2:E,AssetData!E2:E<>"")))'],
    ['Unique Locations', '=COUNTA(UNIQUE(FILTER(AssetData!I2:I,AssetData!I2:I<>"")))'],
    ['', ''],
    ['--- Financial ---', ''],
    ['Total Fleet Value', '=SUM(AssetData!P2:P)'],
    ['Avg Purchase Price', '=IFERROR(AVERAGEIF(AssetData!P2:P,">0"),0)'],
    ['', ''],
    ['--- Fleet Age ---', ''],
    ['Avg Age (Years)', '=IFERROR(AVERAGE(AssetData!AA2:AA),0)'],
    ['Active Devices > 4 Years', '=COUNTIFS(AssetData!AA2:AA,">="&4,AssetData!M2:M,"<>Retired")'],
    ['', ''],
    ['--- Warranty ---', ''],
    ['Warranty Active', '=COUNTIF(AssetData!AB2:AB,"Active")'],
    ['Warranty Expiring (< 90 days)', '=COUNTIF(AssetData!AB2:AB,"Expiring")'],
    ['Warranty Expired', '=COUNTIF(AssetData!AB2:AB,"Expired")'],
    ['No Warranty Data', '=COUNTIF(AssetData!AB2:AB,"None")'],
    ['Warranty Coverage Rate', '=IFERROR((COUNTIF(AssetData!AB2:AB,"Active")+COUNTIF(AssetData!AB2:AB,"Expiring"))/(COUNTA(AssetData!A2:A)-COUNTIF(AssetData!AB2:AB,"None")),0)'],
    ['', ''],
    ['--- Service Load ---', ''],
    ['Total Open Tickets', '=SUM(AssetData!Y2:Y)'],
    ['Assets With Open Tickets', '=COUNTIF(AssetData!Y2:Y,">0")'],
    ['', ''],
    ['--- Assignment ---', ''],
    ['Assigned Devices', '=COUNTIFS(AssetData!A2:A,"<>",AssetData!K2:K,"<>")'],
    ['Unassigned Devices', '=COUNTA(AssetData!A2:A)-COUNTIFS(AssetData!A2:A,"<>",AssetData!K2:K,"<>")'],
    ['Assignment Rate', '=IFERROR(COUNTIFS(AssetData!A2:A,"<>",AssetData!K2:K,"<>")/COUNTA(AssetData!A2:A),0)'],
  ];

  // Write labels and formulas in two batch calls (instead of ~20 individual setFormula calls)
  sheet.getRange(2, 1, kpis.length, 1).setValues(kpis.map(k => [k[0]]));
  const formulaRows = [];
  kpis.forEach((kpi, i) => {
    if (kpi[1] && kpi[1].startsWith('=')) {
      formulaRows.push({ row: i + 2, formula: kpi[1] });
    }
  });
  // Write all formulas in one setFormulas call via a contiguous range
  const valCol = kpis.map(k => [k[1] && k[1].startsWith('=') ? k[1] : '']);
  sheet.getRange(2, 2, kpis.length, 1).setValues(valCol);

  // Batch format section headers, percentages, and currency via getRangeList
  const boldRows = [];
  const pctRows = [];
  const currRows = [];
  kpis.forEach((kpi, i) => {
    if (kpi[0].startsWith('---')) boldRows.push('A' + (i + 2));
    if (['Active Rate', 'Warranty Coverage Rate', 'Assignment Rate'].includes(kpi[0])) pctRows.push('B' + (i + 2));
    if (['Total Fleet Value', 'Avg Purchase Price'].includes(kpi[0])) currRows.push('B' + (i + 2));
  });
  if (isNew) {
    if (boldRows.length) sheet.getRangeList(boldRows).setFontWeight('bold');
    if (pctRows.length) sheet.getRangeList(pctRows).setNumberFormat('0.0%');
    if (currRows.length) sheet.getRangeList(currRows).setNumberFormat('$#,##0');
  }
}

function setupAgingAnalysisSheet(ss) {
  const { sheet, isNew } = getOrCreateSheet(ss, 'AgingAnalysis', [[
    'Year (Purchase/Created)', 'Total Devices', 'Active', 'Retired', 'Warranty Expired', 'Avg Open Tickets', 'Total Value'
  ]], '#365c96');

  const yr = 'IFERROR(YEAR(IF(AssetData!N2:N<>"", AssetData!N2:N, AssetData!Q2:Q)), 0)';
  const formula = '=LET(\n' +
    '  all_status, AssetData!M2:M,\n' +
    '  all_warranty, AssetData!AB2:AB,\n' +
    '  all_tickets, AssetData!Y2:Y,\n' +
    '  all_price, AssetData!P2:P,\n' +
    '  years, SORT(UNIQUE(FILTER(' + yr + ', ' + yr + '>0)), 1, TRUE),\n' +
    '  total, BYROW(years, LAMBDA(y, SUMPRODUCT((' + yr + '=y)*1))),\n' +
    '  active, BYROW(years, LAMBDA(y, SUMPRODUCT((' + yr + '=y)*(all_status<>"Retired")*(all_status<>"")*1))),\n' +
    '  retired, BYROW(years, LAMBDA(y, SUMPRODUCT((' + yr + '=y)*(all_status="Retired")*1))),\n' +
    '  warr_exp, BYROW(years, LAMBDA(y, SUMPRODUCT((' + yr + '=y)*(all_warranty="Expired")*1))),\n' +
    '  avg_tix, BYROW(years, LAMBDA(y, IFERROR(SUMPRODUCT((' + yr + '=y)*IF(ISNUMBER(all_tickets),all_tickets,0))/SUMPRODUCT((' + yr + '=y)*1), 0))),\n' +
    '  tot_val, BYROW(years, LAMBDA(y, SUMPRODUCT((' + yr + '=y)*IF(ISNUMBER(all_price),all_price,0)))),\n' +
    '  IFERROR(SORT(HSTACK(years, total, active, retired, warr_exp, avg_tix, tot_val), 1, TRUE),\n' +
    '    HSTACK(years, total, active, retired, warr_exp, avg_tix, tot_val))\n' +
    ')';

  sheet.getRange(2, 1).setFormula(formula);
  if (isNew) sheet.getRange('G:G').setNumberFormat('$#,##0');
}

function setupServiceImpactSheet(ss) {
  const { sheet, isNew } = getOrCreateSheet(ss, 'ServiceImpact', [[
    'Model', 'Manufacturer', 'Device Count', 'Total Open Tickets', 'Tickets / Device', 'Avg Age (Years)'
  ]], '#365c96');

  const formula = '=LET(\n' +
    '  models, UNIQUE(FILTER(AssetData!E2:E, AssetData!E2:E<>"")),\n' +
    '  mfr, BYROW(models, LAMBDA(m, IFERROR(INDEX(FILTER(AssetData!F:F, AssetData!E:E=m), 1), ""))),\n' +
    '  dev_count, BYROW(models, LAMBDA(m, COUNTIF(AssetData!E:E, m))),\n' +
    '  tot_tickets, BYROW(models, LAMBDA(m, SUMIF(AssetData!E:E, m, AssetData!Y:Y))),\n' +
    '  tix_per, BYROW(models, LAMBDA(m, IFERROR(SUMIF(AssetData!E:E, m, AssetData!Y:Y)/COUNTIF(AssetData!E:E, m), 0))),\n' +
    '  avg_age, BYROW(models, LAMBDA(m, IFERROR(AVERAGEIFS(AssetData!AA:AA, AssetData!E:E, m), 0))),\n' +
    '  IFERROR(SORT(HSTACK(models, mfr, dev_count, tot_tickets, tix_per, avg_age), 5, FALSE),\n' +
    '    HSTACK(models, mfr, dev_count, tot_tickets, tix_per, avg_age))\n' +
    ')';

  sheet.getRange(2, 1).setFormula(formula);
  if (isNew) sheet.getRange('E:E').setNumberFormat('0.00');
}

function setupAssignmentOverviewSheet(ss) {
  const { sheet, isNew } = getOrCreateSheet(ss, 'AssignmentOverview', [[
    'Location', 'Total Assets', 'Assigned', 'Unassigned', 'Assignment Rate', 'Active Assets'
  ]], '#365c96');

  const formula = '=LET(\n' +
    '  locs, UNIQUE(FILTER(AssetData!I2:I, AssetData!I2:I<>"")),\n' +
    '  total, BYROW(locs, LAMBDA(loc, COUNTIF(AssetData!I:I, loc))),\n' +
    '  assigned, BYROW(locs, LAMBDA(loc, COUNTIFS(AssetData!I:I, loc, AssetData!K:K, "<>"))),\n' +
    '  unassigned, BYROW(locs, LAMBDA(loc, COUNTIF(AssetData!I:I, loc) - COUNTIFS(AssetData!I:I, loc, AssetData!K:K, "<>"))),\n' +
    '  pct, BYROW(locs, LAMBDA(loc, IFERROR(COUNTIFS(AssetData!I:I, loc, AssetData!K:K, "<>")/COUNTIF(AssetData!I:I, loc), 0))),\n' +
    '  active, BYROW(locs, LAMBDA(loc, COUNTIFS(AssetData!I:I, loc, AssetData!M:M, "<>Retired"))),\n' +
    '  IFERROR(SORT(HSTACK(locs, total, assigned, unassigned, pct, active), 2, FALSE),\n' +
    '    HSTACK(locs, total, assigned, unassigned, pct, active))\n' +
    ')';

  sheet.getRange(2, 1).setFormula(formula);
  if (isNew) sheet.getRange('E:E').setNumberFormat('0.0%');
}

function setupStatusOverviewSheet(ss) {
  const { sheet, isNew } = getOrCreateSheet(ss, 'StatusOverview', [['Status', 'Count', '% of Total']], '#365c96');

  const formula = '=LET(\n' +
    '  statuses, UNIQUE(FILTER(AssetData!M2:M, AssetData!M2:M<>"")),\n' +
    '  counts, BYROW(statuses, LAMBDA(s, COUNTIF(AssetData!M:M, s))),\n' +
    '  total, SUM(counts),\n' +
    '  pct, BYROW(counts, LAMBDA(c, IFERROR(c / total, 0))),\n' +
    '  IFERROR(SORT(HSTACK(statuses, counts, pct), 2, FALSE),\n' +
    '    HSTACK(statuses, counts, pct))\n' +
    ')';

  sheet.getRange(2, 1).setFormula(formula);
  if (isNew) sheet.getRange('C:C').setNumberFormat('0.0%');
}
