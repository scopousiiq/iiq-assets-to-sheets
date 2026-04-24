/**
 * iiQ Asset Reporting - Configuration
 * Reads settings from the Config sheet, provides logging, locking, and helpers.
 */

/** Current script version — update when releasing new versions */
const SCRIPT_VERSION = '1.2.0';

/**
 * Telemetry endpoint — the deployed iiq-sheets-telemetry Web App /exec URL.
 * Hardcoded here (not in the Config sheet) because this is a maintainer
 * decision, not a per-district setting. Blank until the server is deployed;
 * reportTelemetry() is a no-op while it's blank.
 */
const TELEMETRY_URL = '';

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
  const clockTriggers = triggers.filter(t => t.getEventType() === ScriptApp.EventType.CLOCK);
  const clockNames = clockTriggers.map(t => t.getHandlerFunction());
  const allNames = triggers.map(t => t.getHandlerFunction());
  return {
    hasTriggers: clockTriggers.length > 0,
    count: clockTriggers.length,
    triggers: clockNames,
    allCount: triggers.length,
    allTriggers: allNames
  };
}

function requireNoTriggers(operationName) {
  // Only time-based (CLOCK) triggers block destructive ops — those can race
  // with bulk loads. Installable edit/open triggers are user-driven and safe.
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

// =============================================================================
// VERSION CHECK
// =============================================================================

/**
 * Check for script updates from GitHub.
 * Fetches remote version.json and updates Config sheet display.
 * Fails silently on any error — version check must never break anything.
 */
function checkForUpdates() {
  try {
    const REMOTE_VERSION_URL =
      'https://raw.githubusercontent.com/scopousiiq/iiq-assets-to-sheets/main/version.json';

    const response = UrlFetchApp.fetch(REMOTE_VERSION_URL, {
      muteHttpExceptions: true,
      followRedirects: true
    });

    if (response.getResponseCode() !== 200) {
      logOperation('VersionCheck', 'WARN',
        'Could not reach GitHub (HTTP ' + response.getResponseCode() + ')');
      return;
    }

    const remote = JSON.parse(response.getContentText());
    const remoteVersion = remote.version;

    if (!remoteVersion) {
      logOperation('VersionCheck', 'WARN', 'Remote version.json missing version field');
      return;
    }

    const updateAvailable = isNewerVersion(remoteVersion, SCRIPT_VERSION);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Config');
    if (!sheet) return;

    setConfigValue('SCRIPT_VERSION', SCRIPT_VERSION);
    setConfigValue('VERSION_CHECK_DATE', new Date().toISOString().split('T')[0]);

    if (updateAvailable) {
      setConfigValue('LATEST_VERSION', remoteVersion + '  ← update available');
    } else {
      setConfigValue('LATEST_VERSION', remoteVersion + '  (up to date)');
    }

    // Color the LATEST_VERSION cell: green if current, yellow if update available.
    const data = sheet.getDataRange().getValues();
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === 'LATEST_VERSION') {
        const cell = sheet.getRange(i + 1, 2);
        cell.setBackground(updateAvailable ? '#fff2cc' : '#d9ead3');
        break;
      }
    }

    if (updateAvailable) {
      logOperation('VersionCheck', 'UPDATE_AVAILABLE',
        'v' + remoteVersion + ' available (current: v' + SCRIPT_VERSION + '). ' +
        (remote.releaseUrl || '') + ' — ' + (remote.message || ''));
    } else {
      logOperation('VersionCheck', 'CURRENT', 'v' + SCRIPT_VERSION + ' is up to date');
    }
  } catch (e) {
    logOperation('VersionCheck', 'ERROR', 'Version check failed: ' + e.message);
  }
}

/**
 * Compare two semver strings (e.g., "1.2.0" vs "1.3.0").
 * @return {boolean} True if remoteVer is newer than localVer
 */
function isNewerVersion(remoteVer, localVer) {
  const remote = remoteVer.split('.').map(Number);
  const local = localVer.split('.').map(Number);
  for (let i = 0; i < 3; i++) {
    const r = remote[i] || 0;
    const l = local[i] || 0;
    if (r > l) return true;
    if (r < l) return false;
  }
  return false;
}

/**
 * Check if version check is stale (>24 hours since last check).
 * Guard so triggers don't hit GitHub every invocation.
 */
function isVersionCheckStale() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
    if (!sheet) return true;
    const data = sheet.getDataRange().getValues();
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === 'VERSION_CHECK_DATE') {
        const val = data[i][1];
        if (!val) return true;
        const lastCheck = new Date(val);
        if (isNaN(lastCheck.getTime())) return true;
        const hoursSince = (Date.now() - lastCheck.getTime()) / (1000 * 60 * 60);
        return hoursSince > 24;
      }
    }
    return true;
  } catch (e) {
    return false;
  }
}

/**
 * Write a single Config row's value. Creates the row if missing.
 */
function setConfigValue(key, value) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sheet.appendRow([key, value]);
}

/**
 * Menu wrapper for Check for Updates. Shows a non-blocking toast.
 */
function menuCheckForUpdates() {
  checkForUpdates();
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Version check complete. See Config sheet for results.',
    'Version Check', 5);
}

// =============================================================================
// TELEMETRY (on by default for new installs — opt-out via TELEMETRY_ENABLED)
// =============================================================================
//
// When TELEMETRY_ENABLED is TRUE and TELEMETRY_URL is set, the daily
// triggerDataContinue run POSTs a small JSON payload to the configured Web App
// endpoint so the project maintainer can see install counts, version
// distribution, and API-traffic patterns. No PII is sent — the district is
// identified by a SHA-256 hash of API_BASE_URL plus a locally-generated UUID.
//
// Setup Spreadsheet writes TELEMETRY_ENABLED=TRUE for new installs. Existing
// pre-1.1.0 installs that upgrade without re-running Setup Spreadsheet will
// have no TELEMETRY_ENABLED row — reportTelemetry() reads that as FALSE and
// skips the ping, so the upgrade path doesn't auto-enable telemetry.
//
// All failures are silent.

/**
 * Telemetry ping. Honors TELEMETRY_ENABLED Config row. Fails silently.
 */
function reportTelemetry() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
    if (!sheet) return;
    const raw = readConfigMap(sheet);

    const enabled = String(raw.TELEMETRY_ENABLED || '').toUpperCase() === 'TRUE';
    if (!enabled) return;

    if (!TELEMETRY_URL || TELEMETRY_URL.indexOf('http') !== 0) return;

    const baseUrl = raw.API_BASE_URL || '';
    const districtHash = baseUrl ? sha256Hex(baseUrl.replace(/\/+$/, '').toLowerCase()) : '';
    const installId = getOrCreateInstallId();
    const assetCount = countDataRows('AssetData');
    const analyticsSheets = listInstalledAnalyticsSheets();
    const triggerCount = ScriptApp.getProjectTriggers().length;

    const payload = {
      installId: installId,
      project: 'iiq-assets-to-sheets',
      version: SCRIPT_VERSION,
      districtHash: districtHash,
      assetCount: assetCount,
      analyticsSheets: analyticsSheets,
      triggersEnabled: triggerCount,
      sentAt: new Date().toISOString()
    };

    const response = UrlFetchApp.fetch(TELEMETRY_URL, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
      followRedirects: true
    });

    const code = response.getResponseCode();
    if (code >= 200 && code < 300) {
      setConfigValue('TELEMETRY_LAST_SENT', new Date().toISOString());
      logOperation('Telemetry', 'OK', `Ping accepted (${code}) — ${analyticsSheets.length} analytics sheets, ${assetCount} assets`);
    } else {
      logOperation('Telemetry', 'WARN', `Endpoint returned ${code}`);
    }
  } catch (e) {
    logOperation('Telemetry', 'ERROR', 'Ping failed: ' + e.message);
  }
}

/**
 * True when telemetry hasn't been sent in the last 24 hours.
 */
function isTelemetryStale() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
    if (!sheet) return true;
    const raw = readConfigMap(sheet);
    const last = raw.TELEMETRY_LAST_SENT;
    if (!last) return true;
    const t = new Date(last);
    if (isNaN(t.getTime())) return true;
    return (Date.now() - t.getTime()) / 3600000 > 24;
  } catch (e) {
    return false;
  }
}

/**
 * Stable per-install UUID. Stored in Script Properties so it survives sheet
 * edits and Full Reload, but is distinct per Apps Script project (i.e. per
 * spreadsheet copy).
 */
function getOrCreateInstallId() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('INSTALL_ID');
  if (!id) {
    id = Utilities.getUuid();
    props.setProperty('INSTALL_ID', id);
  }
  return id;
}

function sha256Hex(str) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, str, Utilities.Charset.UTF_8);
  let hex = '';
  for (let i = 0; i < bytes.length; i++) {
    const b = bytes[i] < 0 ? bytes[i] + 256 : bytes[i];
    hex += (b < 16 ? '0' : '') + b.toString(16);
  }
  return hex;
}

function readConfigMap(sheet) {
  const data = sheet.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) map[data[i][0]] = data[i][1];
  }
  return map;
}

function countDataRows(name) {
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!s) return 0;
  const last = s.getLastRow();
  return last > 1 ? last - 1 : 0;
}

function listInstalledAnalyticsSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Canonical set of analytics sheet names — only report ones present.
  const known = [
    'FleetSummary', 'LocationSummary', 'ModelBreakdown', 'AgingAnalysis',
    'BudgetPlanning', 'ServiceImpact', 'AssignmentOverview', 'StatusOverview',
    'DeviceReadiness', 'SpareAssets', 'LostStolenRate', 'ModelFragmentation',
    'UnassignedInventory', 'BreakRate', 'HighTicketLocations',
    'ReplacementPlanning', 'ReplacementForecast', 'WarrantyTimeline',
    'DeviceLifecycle', 'LocationModelBreakdown', 'LocationModelFiltered',
    'CategoryBreakdown', 'ManufacturerSummary', 'IndividualLookup'
  ];
  return known.filter(n => ss.getSheetByName(n) != null);
}
