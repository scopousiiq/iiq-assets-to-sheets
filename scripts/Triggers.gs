/**
 * iiQ Asset Reporting - Triggers
 * Time-driven functions for automated data refresh.
 */

// =============================================================================
// TRIGGER SETUP / REMOVAL
// =============================================================================

function setupAutomatedTriggers() {
  removeAllProjectTriggers();

  // Continue any in-progress initial loading (every 10 min)
  ScriptApp.newTrigger('triggerDataContinue')
    .timeBased().everyMinutes(10).create();

  // Daily incremental refresh (3 AM)
  ScriptApp.newTrigger('triggerDailyRefresh')
    .timeBased().everyDays(1).atHour(3).create();

  // Weekly full refresh to catch edge cases (Sunday 2 AM)
  ScriptApp.newTrigger('triggerWeeklyFullRefresh')
    .timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(2).create();

  logOperation('Triggers', 'SETUP', 'Automated triggers installed (DataContinue 10min, DailyRefresh 3AM, WeeklyRefresh Sun 2AM)');
}

function removeAllProjectTriggers() {
  // Only remove time-based (CLOCK) triggers. Installable edit triggers
  // (e.g., IndividualLookup) are user-driven and preserved across cycles.
  const triggers = ScriptApp.getProjectTriggers();
  const clockTriggers = triggers.filter(t => t.getEventType() === ScriptApp.EventType.CLOCK);
  clockTriggers.forEach(t => ScriptApp.deleteTrigger(t));
  if (clockTriggers.length > 0) {
    logOperation('Triggers', 'REMOVED', `${clockTriggers.length} time-based trigger(s) removed`);
  }
}

// =============================================================================
// TRIGGER FUNCTIONS (headless, no UI)
// =============================================================================

/**
 * Continue any in-progress initial loading.
 * Once initial load is complete, this trigger is a no-op.
 */
function triggerDataContinue() {
  const lock = tryAcquireScriptLock();
  if (!lock) {
    logOperation('Trigger', 'SKIP', 'triggerDataContinue - another operation in progress');
    return;
  }

  try {
    const config = getConfig();

    // Phase 1: Continue asset loading if incomplete
    if (!config.assetComplete) {
      logOperation('Trigger', 'START', 'triggerDataContinue - continuing asset load');
      const result = loadAssetData(false);

      if (result === 'complete') {
        applyAssetFormulas();
        logOperation('Trigger', 'COMPLETE', 'triggerDataContinue - initial load finished, formulas applied');
        // Fall through to check enrollment
      } else {
        return; // Still loading assets, enrollment waits
      }
    }

    // Phase 2: Continue enrollment loading if configured and incomplete
    if (config.studentRoleId) {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const enrollSheet = ss.getSheetByName('LocationEnrollment');
      const locSheet = ss.getSheetByName('Locations');
      const enrollRows = enrollSheet ? enrollSheet.getLastRow() - 1 : 0;
      const locRows = locSheet ? locSheet.getLastRow() - 1 : 0;

      if (locRows > 0 && enrollRows < locRows) {
        logOperation('Trigger', 'START', `triggerDataContinue - continuing enrollment (${enrollRows}/${locRows})`);
        loadLocationEnrollment();
        return;
      }
    }

    logOperation('Trigger', 'SKIP', 'triggerDataContinue - all loading complete');
  } finally {
    releaseScriptLock(lock);
  }
}

/**
 * Daily incremental refresh: fetch assets modified since last refresh.
 * Reapplies formulas after refresh to cover new rows.
 */
function triggerDailyRefresh() {
  const lock = tryAcquireScriptLock();
  if (!lock) {
    logOperation('Trigger', 'SKIP', 'triggerDailyRefresh - another operation in progress');
    return;
  }

  try {
    const config = getConfig();
    if (!config.assetComplete) {
      logOperation('Trigger', 'SKIP', 'triggerDailyRefresh - initial load not complete');
      return;
    }

    logOperation('Trigger', 'START', 'triggerDailyRefresh');
    const result = refreshAssetData(false);

    if (result && typeof result === 'object' && (result.updated > 0 || result.added > 0)) {
      applyAssetFormulas();
    }

    logOperation('Trigger', 'COMPLETE', `triggerDailyRefresh - ${result.updated || 0} updated, ${result.added || 0} new`);
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

    // Clear and reload asset data
    clearAssetDataAndReset();
    loadAssetData(false);

    // If load completed in one run, apply formulas
    const config = getConfig();
    if (config.assetComplete) {
      applyAssetFormulas();
    }

    // Enrollment requires STUDENT_ROLE_ID — skip if not configured
    if (config.studentRoleId) {
      const enrollSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('LocationEnrollment');
      if (enrollSheet && enrollSheet.getLastRow() > 1) {
        enrollSheet.getRange(2, 1, enrollSheet.getLastRow() - 1, 6).clearContent();
      }
      loadLocationEnrollment();
    }

    logOperation('Trigger', 'COMPLETE', 'triggerWeeklyFullRefresh');
  } finally {
    releaseScriptLock(lock);
  }
}
