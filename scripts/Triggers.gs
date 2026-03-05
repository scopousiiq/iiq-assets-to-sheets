/**
 * iiQ Asset Reporting - Triggers
 * Time-driven functions for automated data refresh.
 */

// =============================================================================
// TRIGGER SETUP / REMOVAL
// =============================================================================

function setupAutomatedTriggers() {
  removeAllProjectTriggers();

  // Continue any in-progress loading (Phase 1 or Phase 2)
  ScriptApp.newTrigger('triggerDataContinue')
    .timeBased().everyMinutes(10).create();

  // Weekly full refresh to catch changes (Sunday 2 AM)
  ScriptApp.newTrigger('triggerWeeklyFullRefresh')
    .timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(2).create();

  logOperation('Triggers', 'SETUP', 'Automated triggers installed (DataContinue 10min, WeeklyRefresh Sun 2AM)');
}

function removeAllProjectTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => ScriptApp.deleteTrigger(t));
  if (triggers.length > 0) {
    logOperation('Triggers', 'REMOVED', `${triggers.length} trigger(s) removed`);
  }
}

// =============================================================================
// TRIGGER FUNCTIONS (headless, no UI)
// =============================================================================

/**
 * Continue any in-progress loading.
 * Phase 1: If asset load is incomplete, continue pagination.
 * Phase 2: If enrichment is incomplete, continue enrichment.
 * If both complete, applies formulas if not yet applied.
 */
function triggerDataContinue() {
  const lock = tryAcquireScriptLock();
  if (!lock) {
    logOperation('Trigger', 'SKIP', 'triggerDataContinue - another operation in progress');
    return;
  }

  try {
    const config = getConfig();

    // Phase 1: Continue asset loading
    if (!config.assetComplete) {
      logOperation('Trigger', 'START', 'triggerDataContinue - continuing asset load');
      loadAssetData(false);
      return;
    }

    // Phase 2: Continue enrichment
    if (config.aueFieldId && !config.enrichComplete) {
      logOperation('Trigger', 'START', 'triggerDataContinue - continuing enrichment');
      enrichAssetData(false);
      return;
    }

    // Both complete — nothing to do
    logOperation('Trigger', 'SKIP', 'triggerDataContinue - all loading complete');
  } finally {
    releaseScriptLock(lock);
  }
}

/**
 * Weekly full refresh: clear data and reload everything.
 */
function triggerWeeklyFullRefresh() {
  const lock = tryAcquireScriptLock();
  if (!lock) {
    logOperation('Trigger', 'SKIP', 'triggerWeeklyFullRefresh - another operation in progress');
    return;
  }

  try {
    logOperation('Trigger', 'START', 'triggerWeeklyFullRefresh');

    // Refresh reference data
    loadLocations();
    loadStatusTypes();
    discoverCustomFields();

    // Clear and reload asset data
    clearAssetDataAndReset();
    loadAssetData(false);

    // If load completed in one run, enrich and apply formulas
    const config = getConfig();
    if (config.assetComplete && config.aueFieldId) {
      enrichAssetData(false);
    }
    if (config.assetComplete) {
      applyAssetFormulas();
    }

    logOperation('Trigger', 'COMPLETE', 'triggerWeeklyFullRefresh');
  } finally {
    releaseScriptLock(lock);
  }
}
