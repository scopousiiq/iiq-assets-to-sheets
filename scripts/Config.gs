/**
 * iiQ Asset Reporting - Configuration
 * Reads settings from the Config sheet, provides logging, locking, and helpers.
 */

// =============================================================================
// CONFIG READER
// =============================================================================

function getConfig() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  const allData = sheet.getDataRange().getValues();
  const data = allData.slice(1); // Skip header row

  const rawConfig = {};
  data.forEach(row => {
    if (row[0]) {
      rawConfig[row[0]] = row[1];
    }
  });

  // Normalize base URL: strip trailing slash, ensure /api suffix
  let baseUrl = rawConfig['API_BASE_URL'] || '';
  if (baseUrl) {
    baseUrl = baseUrl.replace(/\/+$/, '');
    if (!baseUrl.endsWith('/api')) {
      baseUrl = baseUrl + '/api';
    }
  }

  return {
    // Required
    baseUrl: baseUrl,
    bearerToken: rawConfig['BEARER_TOKEN'] || '',
    siteId: rawConfig['SITE_ID'] || '',

    // Optional tuning
    pageSize: getIntValue(rawConfig['PAGE_SIZE'], 100),
    throttleMs: getIntValue(rawConfig['THROTTLE_MS'], 1000),
    assetBatchSize: getIntValue(rawConfig['ASSET_BATCH_SIZE'], 500),

    // Enrollment
    studentRoleId: getStringValue(rawConfig['STUDENT_ROLE_ID']),

    // Asset loading progress
    assetTotalPages: getIntValue(rawConfig['ASSET_TOTAL_PAGES'], -1),
    assetLastPage: getIntValue(rawConfig['ASSET_LAST_PAGE'], -1),
    assetComplete: getBoolValue(rawConfig['ASSET_COMPLETE']),

    // Incremental refresh
    lastRefreshDate: getStringValue(rawConfig['LAST_REFRESH_DATE']),
  };
}

// =============================================================================
// TYPE COERCION HELPERS
// =============================================================================

function getStringValue(val) {
  if (!val || val === '') return '';
  if (val instanceof Date) return val.toISOString();
  return String(val).trim();
}

function getIntValue(val, defaultVal) {
  if (val === '' || val === null || val === undefined) return defaultVal;
  const parsed = parseInt(val);
  return isNaN(parsed) ? defaultVal : parsed;
}

function getBoolValue(val) {
  if (val === true || val === 'TRUE' || val === 'true') return true;
  return false;
}

function getDateString(val) {
  if (!val) return '';
  if (val instanceof Date) {
    const year = val.getFullYear();
    const month = String(val.getMonth() + 1).padStart(2, '0');
    const day = String(val.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
  }
  const str = String(val).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(str)) return str;
  const parsed = new Date(str);
  if (!isNaN(parsed.getTime())) {
    const year = parsed.getFullYear();
    const month = String(parsed.getMonth() + 1).padStart(2, '0');
    const day = String(parsed.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
  }
  return '';
}

// =============================================================================
// CONFIG WRITE (with row cache for performance)
// =============================================================================

let configRowCache_ = null;

function cacheConfigRowPositions_() {
  if (configRowCache_) return;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  const data = sheet.getDataRange().getValues();
  configRowCache_ = {};
  data.forEach((row, i) => { configRowCache_[row[0]] = i + 1; });
}

function setConfigValue(key, value) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  if (!sheet) return;

  cacheConfigRowPositions_();
  const row = configRowCache_[key];
  if (row) {
    sheet.getRange(row, 2).setValue(String(value));
  } else {
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, 1, 2).setValues([[key, String(value)]]);
    configRowCache_[key] = lastRow + 1;
  }
}

function resetConfigCache() {
  configRowCache_ = null;
}

// =============================================================================
// LOGGING
// =============================================================================

function logOperation(operation, status, details) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Logs');
  if (!sheet) {
    console.log(`[${operation}] ${status}: ${details}`);
    return;
  }

  sheet.insertRowAfter(1);
  sheet.getRange(2, 1, 1, 4).setValues([[
    new Date().toISOString(),
    operation,
    status,
    details
  ]]);

  // Keep only last 500 entries
  const lastRow = sheet.getLastRow();
  if (lastRow > 501) {
    sheet.deleteRows(502, lastRow - 501);
  }
}

// =============================================================================
// CONCURRENCY CONTROL - LockService helpers
// =============================================================================

function acquireScriptLock(waitTimeMs = 2000) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(waitTimeMs);
    return lock;
  } catch (e) {
    return null;
  }
}

function tryAcquireScriptLock(waitTimeMs = 1000) {
  const lock = LockService.getScriptLock();
  const acquired = lock.tryLock(waitTimeMs);
  return acquired ? lock : null;
}

function releaseScriptLock(lock) {
  if (lock) {
    try { lock.releaseLock(); } catch (e) { /* already released */ }
  }
}

// =============================================================================
// TRIGGER SAFETY
// =============================================================================

function checkForTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  const triggerNames = triggers.map(t => t.getHandlerFunction());
  return { hasTriggers: triggers.length > 0, count: triggers.length, triggers: triggerNames };
}

function requireNoTriggers(operationName) {
  const ui = SpreadsheetApp.getUi();
  const triggerStatus = checkForTriggers();
  if (triggerStatus.hasTriggers) {
    ui.alert(
      'Triggers Must Be Removed',
      `Cannot run "${operationName}" while automated triggers are installed.\n\n` +
      `${triggerStatus.count} trigger(s) found: ${triggerStatus.triggers.join(', ')}\n\n` +
      `Go to: iiQ Assets > Setup > Remove Automated Triggers`,
      ui.ButtonSet.OK
    );
    return false;
  }
  return true;
}

function showOperationBusyMessage(operationName) {
  SpreadsheetApp.getUi().alert(
    'Operation In Progress',
    `Cannot start "${operationName}" because another operation is running.\nPlease wait and try again.`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}
