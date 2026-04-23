/**
 * iiQ Asset Reporting - Menu System
 * Creates "iiQ Assets" menu with all operations.
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('iiQ Assets')
    .addSubMenu(ui.createMenu('Setup')
      .addItem('Setup Spreadsheet', 'menuSetupSpreadsheet')
      .addItem('Verify Configuration', 'menuVerifyConfig')
      .addSeparator()
      .addItem('Setup Automated Triggers', 'menuSetupTriggers')
      .addItem('View Trigger Status', 'menuViewTriggerStatus')
      .addItem('Remove Automated Triggers', 'menuRemoveTriggers')
      .addSeparator()
      .addItem('Check for Updates', 'menuCheckForUpdates')
    )
    .addSeparator()
    .addSubMenu(ui.createMenu('Load Reference Data')
      .addItem('Refresh Locations', 'menuLoadLocations')
      .addItem('Refresh Status Types', 'menuLoadStatusTypes')
      .addItem('Refresh Location Enrollment', 'menuLoadLocationEnrollment')
      .addSeparator()
      .addItem('View Available Roles', 'menuViewRoles')
      .addItem('Refresh All Reference Data', 'menuLoadAllReferenceData')
    )
    .addSubMenu(ui.createMenu('Asset Data')
      .addItem('Load / Resume Assets', 'menuContinueLoading')
      .addItem('Refresh Changed Assets', 'menuRefreshAssets')
      .addItem('Apply Formulas', 'menuApplyFormulas')
      .addItem('Show Status', 'menuShowStatus')
      .addItem('Remove Duplicates', 'menuDeduplicateAssets')
      .addSeparator()
      .addItem('Clear Data + Reset Progress', 'menuClearAndReset')
      .addItem('Full Reload', 'menuFullReload')
    )
    .addSeparator()
    .addSubMenu(ui.createMenu('Analytics Sheets')
      .addSubMenu(ui.createMenu('Fleet Operations')
        .addItem('\u2605 Assignment Overview', 'menuAddAssignmentOverview')
        .addItem('\u2605 Status Overview', 'menuAddStatusOverview')
        .addSeparator()
        .addItem('Device Readiness', 'menuAddDeviceReadiness')
        .addItem('Spare Assets', 'menuAddSpareAssets')
        .addItem('Lost/Stolen Rate', 'menuAddLostStolenRate')
        .addItem('Model Fragmentation', 'menuAddModelFragmentation')
        .addItem('Unassigned Inventory', 'menuAddUnassignedInventory')
        .addSeparator()
        .addItem('Regenerate Fleet Operations', 'menuRegenerateFleetOperations')
      )
      .addSubMenu(ui.createMenu('Service & Reliability')
        .addItem('\u2605 Service Impact', 'menuAddServiceImpact')
        .addSeparator()
        .addItem('Break Rate', 'menuAddBreakRate')
        .addItem('High Ticket Locations', 'menuAddHighTicketLocations')
        .addSeparator()
        .addItem('Regenerate Service & Reliability', 'menuRegenerateServiceReliability')
      )
      .addSubMenu(ui.createMenu('Budget & Planning')
        .addItem('\u2605 Budget Planning', 'menuAddBudgetPlanning')
        .addItem('\u2605 Aging Analysis', 'menuAddAgingAnalysis')
        .addSeparator()
        .addItem('Replacement Planning', 'menuAddReplacementPlanning')
        .addItem('Replacement Forecast', 'menuAddReplacementForecast')
        .addItem('Warranty Timeline', 'menuAddWarrantyTimeline')
        .addItem('Device Lifecycle', 'menuAddDeviceLifecycle')
        .addSeparator()
        .addItem('Regenerate Budget & Planning', 'menuRegenerateBudgetPlanning')
      )
      .addSubMenu(ui.createMenu('Fleet Composition')
        .addItem('\u2605 Fleet Summary', 'menuAddFleetSummary')
        .addItem('\u2605 Location Summary', 'menuAddLocationSummary')
        .addItem('\u2605 Model Breakdown', 'menuAddModelBreakdown')
        .addSeparator()
        .addItem('Location Model Breakdown', 'menuAddLocationModelBreakdown')
        .addItem('Location Model Filtered', 'menuAddLocationModelFiltered')
        .addItem('Category Breakdown', 'menuAddCategoryBreakdown')
        .addItem('Manufacturer Summary', 'menuAddManufacturerSummary')
        .addSeparator()
        .addItem('Regenerate Fleet Composition', 'menuRegenerateFleetComposition')
      )
      .addSubMenu(ui.createMenu('People')
        .addItem('Individual Lookup', 'menuAddIndividualLookup')
        .addSeparator()
        .addItem('Regenerate People', 'menuRegeneratePeople')
      )
      .addSeparator()
      .addItem('Regenerate All Default (\u2605)', 'menuRegenerateAllDefault')
      .addItem('Regenerate All Analytics', 'menuRegenerateAllAnalytics')
    )
    .addToUi();
}

// =============================================================================
// SETUP
// =============================================================================

function menuSetupSpreadsheet() {
  if (!requireNoTriggers('Setup Spreadsheet')) return;
  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert('Confirm Full Reset',
    'This will delete ALL sheets and recreate them from scratch.\n\n' +
    'All data, configuration, and logs will be lost.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return;
  setupSpreadsheet();
}

function menuVerifyConfig() {
  const config = getConfig();
  const ui = SpreadsheetApp.getUi();
  const issues = [];

  if (!config.baseUrl || config.baseUrl.includes('YOUR-DISTRICT')) issues.push('API_BASE_URL not configured — replace the placeholder with your iiQ instance URL');
  if (!config.bearerToken) issues.push('BEARER_TOKEN is empty');

  if (issues.length > 0) {
    ui.alert('Configuration Issues', issues.join('\n'), ui.ButtonSet.OK);
  } else {
    // Try a test API call
    try {
      const response = makeApiRequest('/v1.0/assets?$p=0&$s=1', 'POST', { Filters: [] });
      const total = response.Paging ? response.Paging.TotalRows : '?';
      ui.alert('Configuration OK',
        `API connection successful.\nTotal assets available: ${total}`,
        ui.ButtonSet.OK);
    } catch (e) {
      ui.alert('API Connection Failed', e.message, ui.ButtonSet.OK);
    }
  }
}

function menuSetupTriggers() {
  setupAutomatedTriggers();
  SpreadsheetApp.getUi().alert('Automated triggers have been set up.');
}

function menuViewTriggerStatus() {
  const status = checkForTriggers();
  const otherCount = status.allCount - status.count;
  const others = status.allTriggers.filter(n => status.triggers.indexOf(n) === -1);
  const lines = [];
  if (status.hasTriggers) {
    lines.push(`${status.count} time-based trigger(s):`);
    status.triggers.forEach(n => lines.push('  • ' + n));
  } else {
    lines.push('No time-based triggers installed.');
  }
  if (otherCount > 0) {
    lines.push('');
    lines.push(`${otherCount} other trigger(s) (edit/open/change):`);
    others.forEach(n => lines.push('  • ' + n));
  }
  SpreadsheetApp.getUi().alert('Trigger Status', lines.join('\n'), SpreadsheetApp.getUi().ButtonSet.OK);
}

function menuRemoveTriggers() {
  removeAllProjectTriggers();
  SpreadsheetApp.getUi().alert('All automated triggers have been removed.');
}

// =============================================================================
// REFERENCE DATA
// =============================================================================

function menuLoadLocations() {
  const lock = acquireScriptLock();
  if (!lock) { showOperationBusyMessage('Refresh Locations'); return; }
  try {
    loadLocations();
    SpreadsheetApp.getUi().alert('Locations loaded.');
  } finally { releaseScriptLock(lock); }
}

function menuLoadStatusTypes() {
  const lock = acquireScriptLock();
  if (!lock) { showOperationBusyMessage('Refresh Status Types'); return; }
  try {
    loadStatusTypes();
    SpreadsheetApp.getUi().alert('Status types loaded.');
  } finally { releaseScriptLock(lock); }
}

function menuViewRoles() {
  const ui = SpreadsheetApp.getUi();
  try {
    const roles = getSiteRoles();
    if (!roles || roles.length === 0) {
      ui.alert('No roles returned from API.');
      return;
    }
    const lines = roles.map(r => {
      const name = r.Name || 'Unknown';
      const category = r.CategoryName ? ` [${r.CategoryName}]` : '';
      return `${name}${category}  —  ${r.RoleId}  (${r.Users || 0} users)`;
    });
    ui.alert('Available Roles',
      'Copy the RoleId for your student role into the STUDENT_ROLE_ID config key.\n\n' +
      lines.join('\n'),
      ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', e.message, ui.ButtonSet.OK);
  }
}

function menuLoadLocationEnrollment() {
  const lock = acquireScriptLock();
  if (!lock) { showOperationBusyMessage('Refresh Location Enrollment'); return; }
  try {
    const result = loadLocationEnrollment();
    const ui = SpreadsheetApp.getUi();
    if (result === 'already_complete') {
      ui.alert('Location enrollment is already complete.\nTo reload, clear the LocationEnrollment sheet first.');
    } else if (result === 'paused') {
      ui.alert('Loading paused (timeout).\nRun "Refresh Location Enrollment" again to continue.');
    } else {
      ui.alert('Location enrollment loaded.');
    }
  } finally { releaseScriptLock(lock); }
}

function menuLoadAllReferenceData() {
  const lock = acquireScriptLock();
  if (!lock) { showOperationBusyMessage('Load All Reference Data'); return; }
  try {
    loadLocations();
    loadStatusTypes();

    // Enrollment requires STUDENT_ROLE_ID — skip gracefully if not configured
    const config = getConfig();
    if (config.studentRoleId) {
      loadLocationEnrollment();
    } else {
      logOperation('ReferenceData', 'SKIP', 'Skipping enrollment — STUDENT_ROLE_ID not configured');
    }

    SpreadsheetApp.getUi().alert('All reference data loaded.');
  } finally { releaseScriptLock(lock); }
}

// =============================================================================
// ASSET DATA
// =============================================================================

function menuContinueLoading() {
  const lock = acquireScriptLock();
  if (!lock) { showOperationBusyMessage('Load / Resume Assets'); return; }
  try {
    const result = loadAssetData(true);
    const status = getLoadingStatus();
    const ui = SpreadsheetApp.getUi();

    if (result === 'complete') {
      ui.alert('Loading Complete', `All assets loaded. ${status.rowCount} total rows.`, ui.ButtonSet.OK);
    } else if (result === 'paused') {
      ui.alert('Loading Paused', `Progress: ${status.phase1}\n${status.rowCount} rows so far.\nRun "Load / Resume Assets" again to resume.`, ui.ButtonSet.OK);
    } else if (result === 'already_complete') {
      ui.alert('Already Complete', `Asset loading is already complete. ${status.rowCount} rows.\nUse "Full Reload" to start over.`, ui.ButtonSet.OK);
    }
  } finally { releaseScriptLock(lock); }
}

function menuRefreshAssets() {
  const lock = acquireScriptLock();
  if (!lock) { showOperationBusyMessage('Refresh Changed Assets'); return; }
  try {
    const result = refreshAssetData(true);
    const ui = SpreadsheetApp.getUi();

    if (result === 'initial_load_incomplete') {
      ui.alert('Initial load must complete before refreshing.\nRun "Load / Resume Assets" first.');
    } else if (result === 'no_refresh_date') {
      ui.alert('No previous refresh date found.\nRun "Load / Resume Assets" to complete the initial load, or use "Full Reload".');
    } else if (result && typeof result === 'object') {
      ui.alert('Refresh Complete',
        `Updated: ${result.updated} assets\nNew: ${result.added} assets`,
        ui.ButtonSet.OK);
    }
  } finally { releaseScriptLock(lock); }
}

function menuApplyFormulas() {
  applyAssetFormulas();
  SpreadsheetApp.getUi().alert('Formulas applied to all data rows.');
}

function menuDeduplicateAssets() {
  const lock = acquireScriptLock();
  if (!lock) { showOperationBusyMessage('Remove Duplicates'); return; }
  try {
    const removed = deduplicateAssetData();
    const ui = SpreadsheetApp.getUi();
    if (removed > 0) {
      ui.alert('Deduplication Complete', `Removed ${removed} duplicate rows.`, ui.ButtonSet.OK);
    } else {
      ui.alert('No Duplicates', 'No duplicate AssetIds found.', ui.ButtonSet.OK);
    }
  } finally { releaseScriptLock(lock); }
}

function menuShowStatus() {
  const status = getLoadingStatus();
  SpreadsheetApp.getUi().alert('Loading Status',
    `Asset rows: ${status.rowCount}\n` +
    `Initial Load: ${status.phase1}\n` +
    `Last Refresh: ${status.lastRefresh}`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function menuClearAndReset() {
  if (!requireNoTriggers('Clear Data + Reset')) return;
  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert('Confirm', 'Clear all asset data and reset progress?', ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return;

  const lock = acquireScriptLock();
  if (!lock) { showOperationBusyMessage('Clear Data + Reset'); return; }
  try {
    clearAssetDataAndReset();
    ui.alert('Data cleared and progress reset.');
  } finally { releaseScriptLock(lock); }
}

function menuFullReload() {
  if (!requireNoTriggers('Full Reload')) return;
  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert('Confirm', 'Clear all data and start a full reload?', ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return;

  const lock = acquireScriptLock();
  if (!lock) { showOperationBusyMessage('Full Reload'); return; }
  try {
    clearAssetDataAndReset();
    const result = loadAssetData(true);
    const status = getLoadingStatus();
    if (result === 'complete') {
      ui.alert('Full reload complete. ' + status.rowCount + ' assets loaded.');
    } else {
      ui.alert('Full reload started. ' + status.rowCount + ' assets so far.\nRun "Load / Resume Assets" to resume.');
    }
  } finally { releaseScriptLock(lock); }
}

// =============================================================================
// ANALYTICS — FLEET OPERATIONS
// =============================================================================

function menuAddAssignmentOverview() {
  setupAssignmentOverviewSheet(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getUi().alert('Assignment Overview sheet added.');
}

function menuAddStatusOverview() {
  setupStatusOverviewSheet(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getUi().alert('Status Overview sheet added.');
}

function menuAddDeviceReadiness() {
  setupDeviceReadinessSheet(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getUi().alert('Device Readiness sheet added.');
}

function menuAddSpareAssets() {
  setupSpareAssetsSheet(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getUi().alert('Spare Assets sheet added.');
}

function menuAddLostStolenRate() {
  setupLostStolenRateSheet(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getUi().alert('Lost/Stolen Rate sheet added.');
}

function menuAddModelFragmentation() {
  setupModelFragmentationSheet(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getUi().alert('Model Fragmentation sheet added.');
}

function menuAddUnassignedInventory() {
  setupUnassignedInventorySheet(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getUi().alert('Unassigned Inventory sheet added.');
}

function menuRegenerateFleetOperations() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const count = regenerateFleetOperations(ss);
  SpreadsheetApp.getUi().alert(`Regenerated ${count} Fleet Operations sheet(s).`);
}

// =============================================================================
// ANALYTICS — SERVICE & RELIABILITY
// =============================================================================

function menuAddServiceImpact() {
  setupServiceImpactSheet(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getUi().alert('Service Impact sheet added.');
}

function menuAddBreakRate() {
  setupBreakRateSheet(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getUi().alert('Break Rate sheet added.');
}

function menuAddHighTicketLocations() {
  setupHighTicketLocationsSheet(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getUi().alert('High Ticket Locations sheet added.');
}

function menuRegenerateServiceReliability() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const count = regenerateServiceReliability(ss);
  SpreadsheetApp.getUi().alert(`Regenerated ${count} Service & Reliability sheet(s).`);
}

// =============================================================================
// ANALYTICS — BUDGET & PLANNING
// =============================================================================

function menuAddBudgetPlanning() {
  setupBudgetPlanningSheet(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getUi().alert('Budget Planning sheet added.');
}

function menuAddAgingAnalysis() {
  setupAgingAnalysisSheet(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getUi().alert('Aging Analysis sheet added.');
}

function menuAddReplacementPlanning() {
  setupReplacementPlanningSheet(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getUi().alert('Replacement Planning sheet added.');
}

function menuAddReplacementForecast() {
  setupReplacementForecastSheet(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getUi().alert('Replacement Forecast sheet added.');
}

function menuAddWarrantyTimeline() {
  setupWarrantyTimelineSheet(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getUi().alert('Warranty Timeline sheet added.');
}

function menuAddDeviceLifecycle() {
  setupDeviceLifecycleSheet(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getUi().alert('Device Lifecycle sheet added.');
}

function menuRegenerateBudgetPlanning() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const count = regenerateBudgetPlanning(ss);
  SpreadsheetApp.getUi().alert(`Regenerated ${count} Budget & Planning sheet(s).`);
}

// =============================================================================
// ANALYTICS — FLEET COMPOSITION
// =============================================================================

function menuAddFleetSummary() {
  setupFleetSummarySheet(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getUi().alert('Fleet Summary sheet added.');
}

function menuAddLocationSummary() {
  setupLocationSummarySheet(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getUi().alert('Location Summary sheet added.');
}

function menuAddModelBreakdown() {
  setupModelBreakdownSheet(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getUi().alert('Model Breakdown sheet added.');
}

function menuAddLocationModelBreakdown() {
  setupLocationModelBreakdownSheet(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getUi().alert('Location Model Breakdown sheet added.');
}

function menuAddLocationModelFiltered() {
  setupLocationModelFilteredSheet(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getUi().alert('Location Model Filtered sheet added.');
}

function menuAddCategoryBreakdown() {
  setupCategoryBreakdownSheet(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getUi().alert('Category Breakdown sheet added.');
}

function menuAddManufacturerSummary() {
  setupManufacturerSummarySheet(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getUi().alert('Manufacturer Summary sheet added.');
}

function menuRegenerateFleetComposition() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const count = regenerateFleetComposition(ss);
  SpreadsheetApp.getUi().alert(`Regenerated ${count} Fleet Composition sheet(s).`);
}

// =============================================================================
// ANALYTICS — PEOPLE
// =============================================================================

function menuAddIndividualLookup() {
  setupIndividualLookupSheet(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getUi().alert('Individual Lookup sheet added.\n\nSelect a user from the dropdown in B1 to load their checkout history.');
}

function menuRegeneratePeople() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const count = regeneratePeople(ss);
  SpreadsheetApp.getUi().alert(`Regenerated ${count} People sheet(s).`);
}

// =============================================================================
// ANALYTICS — GLOBAL REGENERATION
// =============================================================================

function menuRegenerateAllDefault() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const count = regenerateAllDefault(ss);
  SpreadsheetApp.getUi().alert(`Regenerated ${count} default (\u2605) analytics sheets.`);
}

function menuRegenerateAllAnalytics() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const count = regenerateAllAnalytics(ss);
  SpreadsheetApp.getUi().alert(`Regenerated ${count} analytics sheet(s) (default + installed optional).`);
}
