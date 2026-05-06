/**
 * iiQ Sheets Telemetry — Client (iiq-assets-to-sheets)
 *
 * Sends one telemetry ping per successful trigger fire to the iiQ-owned
 * Telemetry Master Web App, and enforces the policy that automated polling
 * in iiq-*-to-sheets projects requires TELEMETRY_ENABLED=TRUE.
 *
 * Three functions are called from project code:
 *
 *   reportTelemetry()                    — tail of every trigger-fired function
 *   enforceTelemetryGate()               — head of every trigger-fired function
 *   assertTelemetryEnabledForTriggers()  — head of any function that installs
 *                                          a time-based trigger
 *
 * Deployment model:
 *   - TELEMETRY_URL lives in Config.gs as a maintainer-managed constant.
 *   - TELEMETRY_ENABLED lives in the project's Config sheet; Setup Spreadsheet
 *     stamps it to TRUE for new installs.
 *
 * reportTelemetry() is a silent no-op if any is true:
 *   - TELEMETRY_URL missing/empty
 *   - Config TELEMETRY_ENABLED != TRUE
 *   - No time-based trigger installed
 *   - API_BASE_URL missing or unparseable
 */

// ---- Per-project constants ----

const TELEMETRY_PROJECT = 'iiq-assets-to-sheets';
const TELEMETRY_PRIMARY_SHEET = 'AssetData';
const TELEMETRY_CANONICAL_ANALYTICS = [
  'FleetSummary',
  'LocationSummary',
  'ModelBreakdown',
  'AgingAnalysis',
  'BudgetPlanning',
  'ServiceImpact',
  'AssignmentOverview',
  'StatusOverview',
  'DeviceReadiness',
  'SpareAssets',
  'LostStolenRate',
  'ModelFragmentation',
  'UnassignedInventory',
  'BreakRate',
  'HighTicketLocations',
  'ReplacementPlanning',
  'ReplacementForecast',
  'WarrantyTimeline',
  'DeviceLifecycle',
  'LocationModelBreakdown',
  'LocationModelFiltered',
  'CategoryBreakdown',
  'ManufacturerSummary',
  'IndividualLookup'
];

// ---- Wire protocol — do not edit without coordinating a server change ----

const TELEMETRY_SCHEMA_VERSION = 1;


/**
 * Send one telemetry ping. Call at the tail of every trigger-fired function,
 * after the refresh completes. Best-effort: never throws.
 */
function reportTelemetry() {
  try {
    const url = (typeof TELEMETRY_URL === 'string' ? TELEMETRY_URL : '').trim();
    if (!url) return;

    const cfg = telemetryReadConfig_();
    if (String(cfg.TELEMETRY_ENABLED || '').toUpperCase() !== 'TRUE') return;

    const hasTimeTrigger = ScriptApp.getProjectTriggers().some(function (t) {
      return t.getEventType() === ScriptApp.EventType.CLOCK;
    });
    if (!hasTimeTrigger) return;

    const instanceUrl = telemetryExtractHostname_(cfg.API_BASE_URL);
    if (!instanceUrl) return;

    const ss = SpreadsheetApp.getActive();
    const primary = ss.getSheetByName(TELEMETRY_PRIMARY_SHEET);
    const rowCount = primary ? Math.max(0, primary.getLastRow() - 1) : 0;

    const presentSheets = ss.getSheets().map(function (s) { return s.getName(); });
    const analyticsSheets = TELEMETRY_CANONICAL_ANALYTICS.filter(function (name) {
      return presentSheets.indexOf(name) !== -1;
    });

    const payload = {
      schemaVersion: TELEMETRY_SCHEMA_VERSION,
      installId: telemetryGetOrCreateInstallId_(),
      project: TELEMETRY_PROJECT,
      version: typeof SCRIPT_VERSION === 'string' ? SCRIPT_VERSION : '',
      instanceUrl: instanceUrl,
      installedAt: telemetryGetOrStampInstalledAt_(),
      scriptTimeZone: Session.getScriptTimeZone(),
      sentAt: new Date().toISOString(),
      rowCount: rowCount,
      primarySheet: TELEMETRY_PRIMARY_SHEET,
      analyticsSheets: analyticsSheets
    };

    UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
      followRedirects: true
    });
  } catch (err) {
    console.warn('telemetry failed: ' + (err && err.message || err));
  }
}


/**
 * Runtime policy gate. Call as the FIRST line of every trigger-fired function.
 * Returns true when TELEMETRY_ENABLED=TRUE. Otherwise deletes every CLOCK
 * trigger in the project (leaving edit/open/form-submit triggers alone) and
 * returns false.
 *
 * Fails closed on read error — returns false without uninstalling.
 */
function enforceTelemetryGate() {
  try {
    const cfg = telemetryReadConfig_();
    const enabled = String(cfg.TELEMETRY_ENABLED || '').toUpperCase() === 'TRUE';
    if (enabled) return true;

    let removed = 0;
    ScriptApp.getProjectTriggers().forEach(function (t) {
      if (t.getEventType() === ScriptApp.EventType.CLOCK) {
        ScriptApp.deleteTrigger(t);
        removed++;
      }
    });
    console.warn(
      'Telemetry disabled (TELEMETRY_ENABLED != TRUE). Uninstalled ' +
      removed + ' time-based trigger(s). Automated polling is only available ' +
      'to installs with telemetry enabled. Set TELEMETRY_ENABLED=TRUE in ' +
      'Config and re-run Setup Automated Triggers to reinstall.'
    );
    return false;
  } catch (err) {
    console.error('telemetry gate error: ' + (err && err.message || err));
    return false;
  }
}


/**
 * Install-time policy gate. Call at the top of any function that creates a
 * time-based trigger. Throws a user-readable error if telemetry is off —
 * Apps Script surfaces it in a dialog for menu-triggered callers.
 */
function assertTelemetryEnabledForTriggers() {
  const cfg = telemetryReadConfig_();
  const enabled = String(cfg.TELEMETRY_ENABLED || '').toUpperCase() === 'TRUE';
  if (!enabled) {
    throw new Error(
      'Cannot install automated triggers: TELEMETRY_ENABLED must be TRUE ' +
      'in the Config sheet. Automated API access in iiQ sheet projects ' +
      'requires telemetry opt-in.'
    );
  }
}


// Reads the Config sheet (Key in col A, Value in col B) into a plain object.
// Self-contained so this module can drop into any project without assuming
// the canonical getConfig() helper exists.
function telemetryReadConfig_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Config');
  if (!sheet) return {};
  const last = sheet.getLastRow();
  if (last < 1) return {};
  const values = sheet.getRange(1, 1, last, 2).getValues();
  const out = {};
  for (let i = 0; i < values.length; i++) {
    const key = String(values[i][0] || '').trim();
    if (key) out[key] = values[i][1];
  }
  return out;
}

function telemetryGetOrCreateInstallId_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('TELEMETRY_INSTALL_ID');
  if (!id) {
    id = Utilities.getUuid();
    props.setProperty('TELEMETRY_INSTALL_ID', id);
  }
  return id;
}

function telemetryGetOrStampInstalledAt_() {
  const props = PropertiesService.getScriptProperties();
  let ts = props.getProperty('TELEMETRY_INSTALLED_AT');
  if (!ts) {
    ts = new Date().toISOString();
    props.setProperty('TELEMETRY_INSTALLED_AT', ts);
  }
  return ts;
}

// "https://demo.incidentiq.com/api/v1.0" -> "demo.incidentiq.com"
function telemetryExtractHostname_(urlLike) {
  if (!urlLike) return '';
  let s = String(urlLike).trim().replace(/^https?:\/\//i, '');
  s = s.split('/')[0].split('?')[0].split('#')[0];
  return s.toLowerCase();
}


/**
 * Debug menu handler: force-send one telemetry ping and show diagnostics.
 * Bypasses the TELEMETRY_ENABLED and trigger-presence gates so the endpoint
 * can be exercised on a fresh install before triggers are set up.
 * Only the two checks that would crash the POST are enforced: TELEMETRY_URL
 * must be set, and API_BASE_URL must parse to a hostname.
 */
function menuSendTelemetryPing() {
  const ui = SpreadsheetApp.getUi();

  const url = (typeof TELEMETRY_URL === 'string' ? TELEMETRY_URL : '').trim();
  const cfg = telemetryReadConfig_();
  const enabled = String(cfg.TELEMETRY_ENABLED || '').toUpperCase() === 'TRUE';
  const hasTimeTrigger = ScriptApp.getProjectTriggers().some(function (t) {
    return t.getEventType() === ScriptApp.EventType.CLOCK;
  });
  const instanceUrl = telemetryExtractHostname_(cfg.API_BASE_URL);

  const diagnostics = [
    'TELEMETRY_URL: ' + (url ? 'set' : 'MISSING'),
    'TELEMETRY_ENABLED: ' + (enabled ? 'TRUE' : (cfg.TELEMETRY_ENABLED || '(unset — would gate out normal fires)')),
    'Time-based triggers installed: ' + (hasTimeTrigger ? 'yes' : 'no (would gate out normal fires)'),
    'instanceUrl (from API_BASE_URL): ' + (instanceUrl || 'MISSING')
  ];

  if (!url) {
    ui.alert('Cannot send telemetry ping',
      diagnostics.join('\n') + '\n\nTELEMETRY_URL is empty in Config.gs. ' +
      'This build has telemetry disabled at the source.',
      ui.ButtonSet.OK);
    return;
  }
  if (!instanceUrl) {
    ui.alert('Cannot send telemetry ping',
      diagnostics.join('\n') + '\n\nAPI_BASE_URL is empty or unparseable in ' +
      'the Config sheet. Fill it in and try again.',
      ui.ButtonSet.OK);
    return;
  }

  const ss = SpreadsheetApp.getActive();
  const primary = ss.getSheetByName(TELEMETRY_PRIMARY_SHEET);
  const rowCount = primary ? Math.max(0, primary.getLastRow() - 1) : 0;
  const presentSheets = ss.getSheets().map(function (s) { return s.getName(); });
  const analyticsSheets = TELEMETRY_CANONICAL_ANALYTICS.filter(function (name) {
    return presentSheets.indexOf(name) !== -1;
  });

  const payload = {
    schemaVersion: TELEMETRY_SCHEMA_VERSION,
    installId: telemetryGetOrCreateInstallId_(),
    project: TELEMETRY_PROJECT,
    version: typeof SCRIPT_VERSION === 'string' ? SCRIPT_VERSION : '',
    instanceUrl: instanceUrl,
    installedAt: telemetryGetOrStampInstalledAt_(),
    scriptTimeZone: Session.getScriptTimeZone(),
    sentAt: new Date().toISOString(),
    rowCount: rowCount,
    primarySheet: TELEMETRY_PRIMARY_SHEET,
    analyticsSheets: analyticsSheets
  };

  let responseCode, responseBody;
  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
      followRedirects: true
    });
    responseCode = response.getResponseCode();
    responseBody = response.getContentText();
  } catch (err) {
    try { logOperation('Telemetry', 'ERROR', 'Debug ping failed: ' + (err && err.message || err)); } catch (_) {}
    ui.alert('Telemetry ping failed',
      diagnostics.join('\n') + '\n\nError: ' + (err && err.message || err),
      ui.ButtonSet.OK);
    return;
  }

  try {
    logOperation('Telemetry', 'DEBUG',
      'Debug ping — HTTP ' + responseCode + ' — payload: ' + JSON.stringify(payload) +
      ' — response: ' + responseBody);
  } catch (_) {}

  ui.alert('Telemetry ping sent',
    diagnostics.join('\n') +
    '\n\n--- Response ---\nHTTP ' + responseCode + '\n' + responseBody +
    '\n\n(Full payload written to the Logs sheet.)',
    ui.ButtonSet.OK);
}
