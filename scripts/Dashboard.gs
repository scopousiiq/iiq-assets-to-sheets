/**
 * Dashboard — Native web-app dashboard for asset analytics.
 *
 * Deployment model: published as an Apps Script Web App via doGet().
 * The script is container-bound, so SpreadsheetApp.getActiveSpreadsheet()
 * resolves to the district's sheet when deployed with "Execute as: Me".
 *
 * Data flow:
 *   viewer browser → doGet(e) → HtmlOutput (DashboardApp.html)
 *                      ↓  google.script.run.getDashboardData()
 *   getDashboardData() reads AssetData + registered analytics sheets
 *                      ↓  returns { kpis, badges, categoryGroups }
 *   client renders header → KPI row → badge row → tab bar → active tab
 *
 * See scripts/ChartRegistry.gs for the declarative sheet→chart map.
 * See CLAUDE.md "Deploying the Dashboard (Web App)" for the playbook.
 */

/**
 * 0-indexed column positions in AssetData rows after getValues().
 * NOT for use with getRange (which is 1-indexed).
 *
 * Layout (post v1.4.1): A=AssetId, K=OwnerId, M=StatusName, P=PurchasePrice,
 * AF=AgeYears (formula), AG=WarrantyStatus (formula).
 */
const DASH_COL = {
  OWNER_ID: 10,         // K
  STATUS: 12,           // M
  PURCHASE_PRICE: 15,   // P
  AGE_YEARS: 31,        // AF
  WARRANTY_STATUS: 32   // AG
};

/**
 * Web-app entry point. Called when a viewer loads the /exec URL.
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('DashboardApp')
    .setTitle('iiQ Asset Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Menu handler: shows the deployed dashboard URL (or deployment instructions
 * if not set yet) in a small HtmlService modal inside the sheet.
 */
function showDashboardUrl() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Config');
  let url = '';
  if (configSheet) {
    const values = configSheet.getDataRange().getValues();
    for (let i = 0; i < values.length; i++) {
      if (values[i][0] === 'DASHBOARD_URL') {
        url = String(values[i][1] || '').trim();
        break;
      }
    }
  }

  const html = url
    ? buildUrlModalHtml_(url)
    : buildDeploymentInstructionsHtml_();

  const output = HtmlService.createHtmlOutput(html)
    .setWidth(500)
    .setHeight(420);
  SpreadsheetApp.getUi().showModalDialog(output, 'iiQ Asset Dashboard');
}

function buildUrlModalHtml_(url) {
  const safeUrl = String(url).replace(/"/g, '&quot;');
  return '<style>' +
    'body{font-family:"Proxima Nova","Open Sans",system-ui,sans-serif;padding:18px;color:#374151;}' +
    'h2{color:#365c96;font-size:16px;margin:0 0 8px 0;border-bottom:2px solid #febb12;padding-bottom:6px;}' +
    'p{font-size:13px;margin:10px 0;}' +
    'input{width:100%;padding:8px;font-family:monospace;font-size:12px;border:1px solid #E5E7EB;border-radius:4px;}' +
    '.row{display:flex;gap:8px;margin-top:10px;}' +
    'button,a.btn{display:inline-block;padding:8px 14px;font-size:13px;font-weight:600;border-radius:6px;border:none;cursor:pointer;text-decoration:none;}' +
    '.copy{background:#365c96;color:#fff;}' +
    '.open{background:#22b2a3;color:#fff;}' +
    '.hint{color:#6B7280;font-size:11px;}' +
    '</style>' +
    '<h2>Dashboard URL</h2>' +
    '<p>Share this link with anyone in your domain:</p>' +
    '<input id="u" readonly value="' + safeUrl + '" onclick="this.select();">' +
    '<div class="row">' +
      '<button class="copy" onclick="document.getElementById(\'u\').select();document.execCommand(\'copy\');this.textContent=\'Copied!\';">Copy</button>' +
      '<a class="btn open" href="' + safeUrl + '" target="_blank">Open</a>' +
    '</div>' +
    '<p class="hint">To update the dashboard, redeploy via <strong>Deploy → Manage deployments → Edit → New version</strong> (keeps this URL stable).</p>';
}

function buildDeploymentInstructionsHtml_() {
  return '<style>' +
    'body{font-family:"Proxima Nova","Open Sans",system-ui,sans-serif;padding:18px;color:#374151;}' +
    'h2{color:#365c96;font-size:16px;margin:0 0 8px 0;border-bottom:2px solid #febb12;padding-bottom:6px;}' +
    'p{font-size:13px;margin:8px 0;}' +
    'ol{font-size:12px;line-height:1.6;padding-left:20px;}' +
    'li{margin-bottom:4px;}' +
    'code{background:#F9FAFB;border:1px solid #E5E7EB;padding:1px 4px;border-radius:3px;font-size:11px;}' +
    '</style>' +
    '<h2>Dashboard Not Deployed Yet</h2>' +
    '<p>Publish the dashboard as a web app, then paste the URL into the <code>DASHBOARD_URL</code> row in the Config sheet.</p>' +
    '<ol>' +
      '<li>Open <strong>Extensions → Apps Script</strong>.</li>' +
      '<li>Click <strong>Deploy → New deployment</strong>. Type: <strong>Web app</strong>.</li>' +
      '<li>Execute as: <strong>Me</strong> (the deployer).</li>' +
      '<li>Who has access: <strong>Anyone within your domain</strong>.</li>' +
      '<li>Click <strong>Deploy</strong>, authorize, and copy the <code>/exec</code> URL.</li>' +
      '<li>Paste the URL into the <strong>DASHBOARD_URL</strong> row in the Config sheet.</li>' +
      '<li>Re-open this menu to see the URL and a Copy button.</li>' +
      '<li>For future updates: <strong>Deploy → Manage deployments → Edit → New version</strong> (keeps the same URL).</li>' +
    '</ol>';
}

/**
 * Primary data provider. Called from DashboardApp.html via google.script.run.
 *
 * Returns:
 *   { generatedAt, totalAssets,
 *     kpis:   { totalAssets, avgAgeYears, warrantyActivePct, assignmentRatePct },
 *     badges: [ { sheetName, label, color, count } ],
 *     categoryGroups: [
 *       { category, tabLabel, charts: [
 *         { sheetName, title, type, labels, datasets }
 *       ] }
 *     ] }
 *   or { error: string } on fatal failure.
 */
function getDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const assetSheet = ss.getSheetByName('AssetData');
  if (!assetSheet) {
    return { error: 'AssetData sheet not found. Run Setup first.' };
  }

  const lastRow = assetSheet.getLastRow();
  const rows = lastRow >= 2
    ? assetSheet.getRange(2, 1, lastRow - 1, ASSET_TOTAL_COLS).getValues()
    : [];

  const kpis = computeKpis_(rows);

  const existingSheets = {};
  ss.getSheets().forEach(s => { existingSheets[s.getName()] = s; });

  const badges = [];
  const chartsByCategory = {};

  CHART_REGISTRY.forEach(entry => {
    const sheet = existingSheets[entry.sheetName];
    if (!sheet) return;

    if (entry.kpiBadge) {
      const count = countBadgeRows_(sheet, entry.kpiBadge.rowStart);
      badges.push({
        sheetName: entry.sheetName,
        label: entry.kpiBadge.label,
        color: entry.kpiBadge.color,
        count: count
      });
      return;
    }

    if (entry.lookup) {
      // Lookup tab — no precomputed data; client form drives a live API call.
      chartsByCategory[entry.category] = {
        category: entry.category,
        tabLabel: entry.tabLabel || entry.category,
        charts: [],
        lookup: Object.assign({ sheetName: entry.sheetName }, entry.lookup)
      };
      return;
    }

    if (!entry.charts || entry.charts.length === 0) return;

    const values = sheet.getDataRange().getValues();
    entry.charts.forEach(spec => {
      const built = buildChartPayload_(values, spec);
      if (!built) return;
      if (!chartsByCategory[entry.category]) {
        chartsByCategory[entry.category] = {
          category: entry.category,
          tabLabel: entry.tabLabel || entry.category,
          charts: []
        };
      }
      chartsByCategory[entry.category].charts.push(Object.assign({
        sheetName: entry.sheetName,
        title: spec.title,
        type: spec.type
      }, built));
    });
  });

  const categoryGroups = [];
  CATEGORY_ORDER.forEach(cat => {
    if (chartsByCategory[cat]) categoryGroups.push(chartsByCategory[cat]);
  });

  return {
    generatedAt: new Date().toISOString(),
    totalAssets: rows.length,
    kpis: kpis,
    badges: badges,
    categoryGroups: categoryGroups
  };
}

/** Fixed KPI row — computed directly from AssetData rows. */
function computeKpis_(rows) {
  const total = rows.length;
  let ageSum = 0, ageCount = 0;
  let warrantyActive = 0;
  let assigned = 0;

  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    const age = Number(r[DASH_COL.AGE_YEARS]);
    if (isFinite(age) && age >= 0) { ageSum += age; ageCount++; }
    if (r[DASH_COL.WARRANTY_STATUS] === 'Active') warrantyActive++;
    const owner = r[DASH_COL.OWNER_ID];
    if (owner !== '' && owner !== null && owner !== undefined) assigned++;
  }

  return {
    totalAssets: total,
    avgAgeYears: ageCount > 0 ? Math.round((ageSum / ageCount) * 10) / 10 : 0,
    warrantyActivePct: total > 0 ? Math.round((warrantyActive / total) * 1000) / 10 : 0,
    assignmentRatePct: total > 0 ? Math.round((assigned / total) * 1000) / 10 : 0
  };
}

// =============================================================================
// LOOKUP TABS — server functions called from DashboardApp.html via
// google.script.run. Reuse existing helpers in OptionalAnalytics.gs and
// ApiClient.gs (those are global functions in Apps Script).
// =============================================================================

/** Returns sorted unique list of OwnerFullName from AssetData, for the
 *  Individual Lookup dropdown. */
function dashboardGetOwnerNames() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const assetData = ss.getSheetByName('AssetData');
  if (!assetData || assetData.getLastRow() < 2) return [];
  // Col 12 = OwnerFullName (1-indexed for getRange).
  const values = assetData.getRange(2, 12, assetData.getLastRow() - 1, 1).getValues();
  const seen = {};
  for (let i = 0; i < values.length; i++) {
    const name = values[i][0];
    if (name) seen[name] = true;
  }
  return Object.keys(seen).sort();
}

/** Live IndividualLookup, returns { status, rows, headers, error? } */
function dashboardLookupIndividual(ownerName) {
  const headers = ['Date', 'Action', 'Asset Tag', 'Serial Number', 'Model', 'Location', 'Currently With'];
  const name = (ownerName || '').toString().trim();
  if (!name) return { status: 'Select a user.', rows: [], headers: headers };
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ownerId = resolveOwnerIdByName(ss, name);
    if (!ownerId) {
      return { status: 'No OwnerId for "' + name + '" in AssetData.', rows: [], headers: headers };
    }
    const response = getUserActivities(ownerId);
    const items = (response && response.Items) || [];
    if (items.length === 0) {
      return { status: 'No asset assignment events returned for ' + name + '.', rows: [], headers: headers };
    }
    const assetMap = buildAssetTagMap(ss);
    const locationMap = buildLocationMap(ss);
    const rows = items
      .map(item => parseActivityRow(item, assetMap, locationMap))
      .filter(row => row !== null);
    rows.sort((a, b) => (b[0] > a[0] ? 1 : b[0] < a[0] ? -1 : 0));
    const scanned = items.length;
    const suffix = scanned >= 500 ? ' (scanned 500 most-recent activities)' : '';
    return {
      status: rows.length + ' asset event' + (rows.length === 1 ? '' : 's') + suffix,
      rows: serializeRows_(rows),
      headers: headers
    };
  } catch (err) {
    try { logOperation('IndividualLookup', 'ERROR', (err.message || err).toString().substring(0, 200)); } catch (_) {}
    return { status: '', rows: [], headers: headers, error: (err.message || err).toString() };
  }
}

/** Live VerificationLookup, returns { status, rows, headers, error? } */
function dashboardLookupVerification(input) {
  const headers = ['Date', 'Result', 'Method', 'Location', 'Room', 'Verified By', 'Comments'];
  const trimmed = (input || '').toString().trim();
  if (!trimmed) return { status: 'Enter an Asset Tag or Serial Number.', rows: [], headers: headers };
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const asset = resolveAssetByTagOrSerial(ss, trimmed);
    if (!asset) {
      return { status: 'No asset matched "' + trimmed + '" by AssetTag or SerialNumber.', rows: [], headers: headers };
    }
    const response = getAssetVerifications(asset.assetId);
    const items = (response && response.Items) || [];
    if (items.length === 0) {
      return { status: formatAssetSummary(asset) + ' — 0 verifications', rows: [], headers: headers };
    }
    const locationMap = buildLocationMap(ss);
    const userMap = resolveVerifierNames(items);
    const roomMap = resolveRoomNames(items);
    const rows = items
      .map(item => parseVerificationRow(item, locationMap, userMap, roomMap))
      .filter(row => row !== null);
    rows.sort((a, b) => (b[0] > a[0] ? 1 : b[0] < a[0] ? -1 : 0));
    return {
      status: formatAssetSummary(asset) + ' — ' + rows.length + ' verification' + (rows.length === 1 ? '' : 's'),
      rows: serializeRows_(rows),
      headers: headers
    };
  } catch (err) {
    try { logOperation('VerificationLookup', 'ERROR', (err.message || err).toString().substring(0, 200)); } catch (_) {}
    return { status: '', rows: [], headers: headers, error: (err.message || err).toString() };
  }
}

/** Coerce row cells into JSON-safe values. Date → ISO string, else String(). */
function serializeRows_(rows) {
  return rows.map(row => row.map(cell => {
    if (cell instanceof Date) return cell.toISOString();
    if (cell === null || cell === undefined) return '';
    return cell;
  }));
}

/** Count non-empty label cells past rowStart, skipping section markers. */
function countBadgeRows_(sheet, rowStart) {
  const lastRow = sheet.getLastRow();
  if (lastRow < rowStart) return 0;
  const values = sheet.getRange(rowStart, 1, lastRow - rowStart + 1, 1).getValues();
  let count = 0;
  for (let i = 0; i < values.length; i++) {
    const v = values[i][0];
    if (v === '' || v === null || v === undefined) continue;
    const s = String(v);
    if (s.indexOf('---') === 0) continue;
    if (s.indexOf('No ') === 0) continue;
    count++;
  }
  return count;
}

/**
 * Build a Chart.js-ready payload from a sheet's raw values using a chart spec.
 * Returns null if no valid data rows remain after dirty-data filtering.
 */
function buildChartPayload_(values, spec) {
  const rowStart = spec.rowStart || 2;
  const startIdx = rowStart - 1;

  // 1. Slice the data window.
  let window;
  if (spec.rowMode === 'fixed') {
    window = values.slice(startIdx, startIdx + (spec.rowCount || 0));
  } else {
    window = [];
    for (let i = startIdx; i < values.length; i++) {
      const label = values[i][spec.labelCol];
      if (label === '' || label === null || label === undefined) break;
      const s = String(label);
      if (s.indexOf('---') === 0) break;
      window.push(values[i]);
    }
  }
  if (window.length === 0) return null;

  // 2. Resolve series.
  const series = spec.series;
  if (!series || series.length === 0) return null;

  // 3. Extract labels + numeric series, dropping rows with any non-finite value.
  const labels = [];
  const dataCols = series.map(() => []);

  for (let i = 0; i < window.length; i++) {
    const row = window[i];
    const label = row[spec.labelCol];
    const values_ = series.map(s => Number(row[s.col]));
    if (values_.some(v => !isFinite(v))) continue;
    labels.push(label instanceof Date ? label.toISOString().slice(0, 10) : String(label));
    values_.forEach((v, idx) => dataCols[idx].push(v));
  }

  if (labels.length === 0) return null;

  const datasets = series.map((s, idx) => ({
    label: s.header,
    color: s.color,
    data: dataCols[idx],
    percent: !!s.percent
  }));

  return { labels: labels, datasets: datasets };
}
