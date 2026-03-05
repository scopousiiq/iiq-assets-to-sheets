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
    )
    .addSeparator()
    .addSubMenu(ui.createMenu('Load Reference Data')
      .addItem('Refresh Locations', 'menuLoadLocations')
      .addItem('Refresh Status Types', 'menuLoadStatusTypes')
      .addItem('Discover Custom Fields', 'menuDiscoverCustomFields')
      .addSeparator()
      .addItem('Refresh All Reference Data', 'menuLoadAllReferenceData')
    )
    .addSubMenu(ui.createMenu('Asset Data')
      .addItem('Continue Loading', 'menuContinueLoading')
      .addItem('Enrich Custom Fields', 'menuEnrichAssets')
      .addItem('Apply Formulas', 'menuApplyFormulas')
      .addItem('Show Status', 'menuShowStatus')
      .addSeparator()
      .addItem('Clear Data + Reset Progress', 'menuClearAndReset')
      .addItem('Full Reload', 'menuFullReload')
    )
    .addSeparator()
    .addSubMenu(ui.createMenu('Analytics Sheets')
      .addItem('Location Summary', 'menuAddLocationSummary')
      .addItem('Model Breakdown', 'menuAddModelBreakdown')
      .addItem('AUE Planning', 'menuAddAUEPlanning')
      .addItem('Budget Planning', 'menuAddBudgetPlanning')
      .addItem('Status Overview', 'menuAddStatusOverview')
      .addSeparator()
      .addItem('Regenerate All Analytics', 'menuRegenerateAllAnalytics')
    )
    .addToUi();
}

// =============================================================================
// SETUP
// =============================================================================

function menuSetupSpreadsheet() {
  setupSpreadsheet();
}

function menuVerifyConfig() {
  const config = getConfig();
  const ui = SpreadsheetApp.getUi();
  const issues = [];

  if (!config.baseUrl) issues.push('API_BASE_URL is empty');
  if (!config.bearerToken) issues.push('BEARER_TOKEN is empty');

  if (issues.length > 0) {
    ui.alert('Configuration Issues', issues.join('\n'), ui.ButtonSet.OK);
  } else {
    // Try a test API call
    try {
      const response = makeApiRequest('/v1.0/assets?$p=0&$s=1', 'POST', { Filters: [] });
      const total = response.Paging ? response.Paging.TotalRows : '?';
      ui.alert('Configuration OK',
        `API connection successful.\nTotal assets available: ${total}\nAUE Field: ${config.aueFieldId || 'Not configured'}`,
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
  const msg = status.hasTriggers
    ? `${status.count} trigger(s) active:\n${status.triggers.join('\n')}`
    : 'No automated triggers are installed.';
  SpreadsheetApp.getUi().alert('Trigger Status', msg, SpreadsheetApp.getUi().ButtonSet.OK);
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

function menuDiscoverCustomFields() {
  const lock = acquireScriptLock();
  if (!lock) { showOperationBusyMessage('Discover Custom Fields'); return; }
  try {
    discoverCustomFields();
    const config = getConfig();
    const msg = config.aueFieldId
      ? `Custom fields discovered. AUE field detected: "${config.aueFieldName}"`
      : 'Custom fields discovered. No AUE field auto-detected.\nSet AUE_CUSTOM_FIELD_ID in Config if you have one.';
    SpreadsheetApp.getUi().alert(msg);
  } finally { releaseScriptLock(lock); }
}

function menuLoadAllReferenceData() {
  const lock = acquireScriptLock();
  if (!lock) { showOperationBusyMessage('Load All Reference Data'); return; }
  try {
    loadLocations();
    loadStatusTypes();
    discoverCustomFields();
    SpreadsheetApp.getUi().alert('All reference data loaded.');
  } finally { releaseScriptLock(lock); }
}

// =============================================================================
// ASSET DATA
// =============================================================================

function menuContinueLoading() {
  const lock = acquireScriptLock();
  if (!lock) { showOperationBusyMessage('Continue Loading'); return; }
  try {
    const result = loadAssetData(true);
    const status = getLoadingStatus();
    const ui = SpreadsheetApp.getUi();

    if (result === 'complete') {
      ui.alert('Loading Complete', `All assets loaded. ${status.rowCount} total rows.`, ui.ButtonSet.OK);
    } else if (result === 'paused') {
      ui.alert('Loading Paused', `Progress: ${status.phase1}\n${status.rowCount} rows so far.\nRun "Continue Loading" again to resume.`, ui.ButtonSet.OK);
    } else if (result === 'already_complete') {
      ui.alert('Already Complete', `Asset loading is already complete. ${status.rowCount} rows.\nUse "Full Reload" to start over.`, ui.ButtonSet.OK);
    }
  } finally { releaseScriptLock(lock); }
}

function menuEnrichAssets() {
  const lock = acquireScriptLock();
  if (!lock) { showOperationBusyMessage('Enrich Custom Fields'); return; }
  try {
    const result = enrichAssetData(true);
    const ui = SpreadsheetApp.getUi();

    if (result === 'complete' || result === 'already_complete') {
      ui.alert('Enrichment complete. AUE dates populated.');
    } else if (result === 'paused') {
      ui.alert('Enrichment paused. Run again to continue.');
    } else if (result === 'no_aue_field') {
      ui.alert('No AUE custom field configured.\nRun "Discover Custom Fields" first, or set AUE_CUSTOM_FIELD_ID manually.');
    } else if (result === 'skipped') {
      ui.alert('Asset loading must complete before enrichment.');
    }
  } finally { releaseScriptLock(lock); }
}

function menuApplyFormulas() {
  applyAssetFormulas();
  SpreadsheetApp.getUi().alert('Formulas applied to all data rows.');
}

function menuShowStatus() {
  const status = getLoadingStatus();
  SpreadsheetApp.getUi().alert('Loading Status',
    `Asset rows: ${status.rowCount}\n` +
    `Phase 1 (Bulk Load): ${status.phase1}\n` +
    `Phase 2 (Enrichment): ${status.phase2}`,
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
      ui.alert('Full reload started. ' + status.rowCount + ' assets so far.\nRun "Continue Loading" to resume.');
    }
  } finally { releaseScriptLock(lock); }
}

// =============================================================================
// ANALYTICS
// =============================================================================

function menuAddLocationSummary() {
  setupLocationSummarySheet(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getUi().alert('Location Summary sheet added.');
}

function menuAddModelBreakdown() {
  setupModelBreakdownSheet(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getUi().alert('Model Breakdown sheet added.');
}

function menuAddAUEPlanning() {
  setupAUEPlanningSheet(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getUi().alert('AUE Planning sheet added.');
}

function menuAddBudgetPlanning() {
  setupBudgetPlanningSheet(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getUi().alert('Budget Planning sheet added.');
}

function menuAddStatusOverview() {
  setupStatusOverviewSheet(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getUi().alert('Status Overview sheet added.');
}

function menuRegenerateAllAnalytics() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupLocationSummarySheet(ss);
  setupModelBreakdownSheet(ss);
  setupAUEPlanningSheet(ss);
  setupBudgetPlanningSheet(ss);
  setupStatusOverviewSheet(ss);
  SpreadsheetApp.getUi().alert('All analytics sheets regenerated.');
}
