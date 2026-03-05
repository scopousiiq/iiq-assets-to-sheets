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
  setupCustomFieldsSheet(ss);
  setupLogsSheet(ss);
  setupInstructionsSheet(ss);

  // Default analytics sheets
  setupLocationSummarySheet(ss);
  setupModelBreakdownSheet(ss);
  setupAUEPlanningSheet(ss);
  setupBudgetPlanningSheet(ss);
  setupStatusOverviewSheet(ss);

  SpreadsheetApp.getUi().alert('Setup complete! Fill in the Config sheet, then load reference data.');
}

function deleteSheetIfExists(ss, name) {
  const sheet = ss.getSheetByName(name);
  if (sheet) ss.deleteSheet(sheet);
}

// =============================================================================
// DATA SHEETS
// =============================================================================

function setupConfigSheet(ss) {
  let sheet = ss.getSheetByName('Config');
  if (sheet) return; // Don't overwrite existing config

  sheet = ss.insertSheet('Config');
  sheet.getRange(1, 1, 1, 2).setValues([['Key', 'Value']]).setFontWeight('bold');

  const configRows = [
    ['API_BASE_URL', ''],
    ['BEARER_TOKEN', ''],
    ['SITE_ID', ''],
    ['PAGE_SIZE', '100'],
    ['THROTTLE_MS', '1000'],
    ['ASSET_BATCH_SIZE', '500'],
    ['', ''],
    ['--- Progress (auto-managed) ---', ''],
    ['ASSET_LAST_PAGE', '-1'],
    ['ASSET_TOTAL_PAGES', '-1'],
    ['ASSET_COMPLETE', 'FALSE'],
    ['ENRICH_LAST_IDX', '-1'],
    ['ENRICH_COMPLETE', 'FALSE'],
    ['', ''],
    ['--- Custom Fields (auto or manual) ---', ''],
    ['AUE_CUSTOM_FIELD_ID', ''],
    ['AUE_CUSTOM_FIELD_NAME', ''],
  ];

  sheet.getRange(2, 1, configRows.length, 2).setValues(configRows);
  sheet.setColumnWidth(1, 250);
  sheet.setColumnWidth(2, 400);
}

function setupAssetDataSheet(ss) {
  deleteSheetIfExists(ss, 'AssetData');
  const sheet = ss.insertSheet('AssetData');

  // Headers
  sheet.getRange(1, 1, 1, ASSET_TOTAL_COLS).setValues([ASSET_HEADERS]).setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);

  // Formula columns (U-Y) — placed in row 2, will auto-extend via arrays if needed
  // These are per-row formulas, applied when data is present via setup helper
  // For now, leave empty — formulas will be applied after data loads
}

/**
 * Apply per-row formulas to AssetData sheet for all data rows.
 * Call after data loading completes.
 */
function applyAssetFormulas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('AssetData');
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  const numRows = lastRow - 1;

  // Column U: AUEStatus — based on AUEDate (col T)
  // "Expired", "< 6 Months", "< 1 Year", "< 2 Years", "OK", or blank
  const aueFormulas = [];
  for (let r = 2; r <= lastRow; r++) {
    aueFormulas.push([
      `=IF(T${r}="","",IF(DATEVALUE(T${r})<TODAY(),"Expired",IF(DATEVALUE(T${r})<TODAY()+180,"< 6 Months",IF(DATEVALUE(T${r})<TODAY()+365,"< 1 Year",IF(DATEVALUE(T${r})<TODAY()+730,"< 2 Years","OK")))))`
    ]);
  }
  sheet.getRange(2, 21, numRows, 1).setFormulas(aueFormulas);

  // Column V: AgeDays — days since PurchasedDate (col N)
  const ageDayFormulas = [];
  for (let r = 2; r <= lastRow; r++) {
    ageDayFormulas.push([`=IF(N${r}="","",TODAY()-DATEVALUE(N${r}))`]);
  }
  sheet.getRange(2, 22, numRows, 1).setFormulas(ageDayFormulas);

  // Column W: AgeYears — AgeDays / 365.25
  const ageYearFormulas = [];
  for (let r = 2; r <= lastRow; r++) {
    ageYearFormulas.push([`=IF(V${r}="","",ROUND(V${r}/365.25,1))`]);
  }
  sheet.getRange(2, 23, numRows, 1).setFormulas(ageYearFormulas);

  // Column X: WarrantyStatus — based on WarrantyExpDate (col O)
  const warrantyFormulas = [];
  for (let r = 2; r <= lastRow; r++) {
    warrantyFormulas.push([
      `=IF(O${r}="","None",IF(DATEVALUE(O${r})<TODAY(),"Expired",IF(DATEVALUE(O${r})<TODAY()+90,"Expiring","Active")))`
    ]);
  }
  sheet.getRange(2, 24, numRows, 1).setFormulas(warrantyFormulas);

  // Column Y: ReplacementCycle — Fiscal year when AUE expires (July-June)
  const replFormulas = [];
  for (let r = 2; r <= lastRow; r++) {
    replFormulas.push([
      `=IF(T${r}="","",IF(MONTH(DATEVALUE(T${r}))>=7,YEAR(DATEVALUE(T${r}))&"-"&YEAR(DATEVALUE(T${r}))+1,YEAR(DATEVALUE(T${r}))-1&"-"&YEAR(DATEVALUE(T${r}))))`
    ]);
  }
  sheet.getRange(2, 25, numRows, 1).setFormulas(replFormulas);

  logOperation('Formulas', 'COMPLETE', `Applied formulas to ${numRows} rows`);
}

function setupLocationsSheet(ss) {
  deleteSheetIfExists(ss, 'Locations');
  const sheet = ss.insertSheet('Locations');
  sheet.getRange(1, 1, 1, 5).setValues([['LocationId', 'Name', 'Abbreviation', 'LocationType', 'Address']]).setFontWeight('bold');
  sheet.setFrozenRows(1);
}

function setupStatusTypesSheet(ss) {
  deleteSheetIfExists(ss, 'StatusTypes');
  const sheet = ss.insertSheet('StatusTypes');
  sheet.getRange(1, 1, 1, 4).setValues([['StatusTypeId', 'Name', 'IsRetired', 'SortOrder']]).setFontWeight('bold');
  sheet.setFrozenRows(1);
}

function setupCustomFieldsSheet(ss) {
  deleteSheetIfExists(ss, 'CustomFields');
  const sheet = ss.insertSheet('CustomFields');
  sheet.getRange(1, 1, 1, 5).setValues([['FieldId', 'Name', 'Type', 'Required', 'Description']]).setFontWeight('bold');
  sheet.setFrozenRows(1);
}

function setupLogsSheet(ss) {
  deleteSheetIfExists(ss, 'Logs');
  const sheet = ss.insertSheet('Logs');
  sheet.getRange(1, 1, 1, 4).setValues([['Timestamp', 'Operation', 'Status', 'Details']]).setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(4, 500);
}

function setupInstructionsSheet(ss) {
  deleteSheetIfExists(ss, 'Instructions');
  const sheet = ss.insertSheet('Instructions');
  ss.setActiveSheet(sheet);
  ss.moveActiveSheet(1);

  const instructions = [
    ['iiQ Asset Reporting - Setup Guide', ''],
    ['', ''],
    ['Step 1:', 'Fill in the Config sheet with your API credentials'],
    ['Step 2:', 'Run "iiQ Assets > Load Reference Data > Refresh Locations"'],
    ['Step 3:', 'Run "iiQ Assets > Load Reference Data > Refresh Status Types"'],
    ['Step 4:', 'Run "iiQ Assets > Load Reference Data > Discover Custom Fields"'],
    ['Step 5:', 'Check CustomFields sheet — verify AUE field was detected (or set AUE_CUSTOM_FIELD_ID manually)'],
    ['Step 6:', 'Run "iiQ Assets > Asset Data > Continue Loading" (repeat until complete)'],
    ['Step 7:', 'Run "iiQ Assets > Asset Data > Enrich Custom Fields" (for AUE dates)'],
    ['Step 8:', 'Run "iiQ Assets > Asset Data > Apply Formulas" to generate calculated columns'],
    ['Step 9:', 'Set up automated triggers via "iiQ Assets > Setup > Setup Automated Triggers"'],
    ['', ''],
    ['Analytics sheets are pre-built with formulas that auto-calculate from AssetData.', ''],
    ['', ''],
    ['For Looker Studio dashboards, connect directly to this spreadsheet.', ''],
  ];

  sheet.getRange(1, 1, instructions.length, 2).setValues(instructions);
  sheet.getRange(1, 1).setFontSize(14).setFontWeight('bold');
  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 600);
}

// =============================================================================
// DEFAULT ANALYTICS SHEETS
// =============================================================================

function setupLocationSummarySheet(ss) {
  deleteSheetIfExists(ss, 'LocationSummary');
  const sheet = ss.insertSheet('LocationSummary');
  sheet.getRange(1, 1, 1, 7).setValues([[
    'Location', 'Total Assets', 'Active', 'Retired', 'AUE Expired', 'AUE < 1 Year', 'Avg Age (Years)'
  ]]).setFontWeight('bold');
  sheet.setFrozenRows(1);

  // Single formula in A2 that builds the entire table
  const formula = '=LET(\n' +
    '  locs, UNIQUE(FILTER(AssetData!I2:I, AssetData!I2:I<>"")),\n' +
    '  total, BYROW(locs, LAMBDA(loc, COUNTIF(AssetData!I:I, loc))),\n' +
    '  active, BYROW(locs, LAMBDA(loc, COUNTIFS(AssetData!I:I, loc, AssetData!M:M, "<>Retired"))),\n' +
    '  retired, BYROW(locs, LAMBDA(loc, COUNTIFS(AssetData!I:I, loc, AssetData!M:M, "Retired"))),\n' +
    '  aue_exp, BYROW(locs, LAMBDA(loc, COUNTIFS(AssetData!I:I, loc, AssetData!U:U, "Expired"))),\n' +
    '  aue_soon, BYROW(locs, LAMBDA(loc, COUNTIFS(AssetData!I:I, loc, AssetData!U:U, "< 6 Months")+COUNTIFS(AssetData!I:I, loc, AssetData!U:U, "< 1 Year"))),\n' +
    '  avg_age, BYROW(locs, LAMBDA(loc, IFERROR(AVERAGEIFS(AssetData!W:W, AssetData!I:I, loc), 0))),\n' +
    '  SORT(HSTACK(locs, total, active, retired, aue_exp, aue_soon, avg_age), 2, FALSE)\n' +
    ')';

  sheet.getRange(2, 1).setFormula(formula);
}

function setupModelBreakdownSheet(ss) {
  deleteSheetIfExists(ss, 'ModelBreakdown');
  const sheet = ss.insertSheet('ModelBreakdown');
  sheet.getRange(1, 1, 1, 6).setValues([[
    'Model', 'Manufacturer', 'Total', 'Active', 'Retired', 'Avg Age (Years)'
  ]]).setFontWeight('bold');
  sheet.setFrozenRows(1);

  const formula = '=LET(\n' +
    '  models, UNIQUE(FILTER(AssetData!E2:E, AssetData!E2:E<>"")),\n' +
    '  mfr, BYROW(models, LAMBDA(m, IFERROR(INDEX(FILTER(AssetData!F:F, AssetData!E:E=m), 1), ""))),\n' +
    '  total, BYROW(models, LAMBDA(m, COUNTIF(AssetData!E:E, m))),\n' +
    '  active, BYROW(models, LAMBDA(m, COUNTIFS(AssetData!E:E, m, AssetData!M:M, "<>Retired"))),\n' +
    '  retired, BYROW(models, LAMBDA(m, COUNTIFS(AssetData!E:E, m, AssetData!M:M, "Retired"))),\n' +
    '  avg_age, BYROW(models, LAMBDA(m, IFERROR(AVERAGEIFS(AssetData!W:W, AssetData!E:E, m), 0))),\n' +
    '  SORT(HSTACK(models, mfr, total, active, retired, avg_age), 3, FALSE)\n' +
    ')';

  sheet.getRange(2, 1).setFormula(formula);
}

function setupAUEPlanningSheet(ss) {
  deleteSheetIfExists(ss, 'AUEPlanning');
  const sheet = ss.insertSheet('AUEPlanning');
  sheet.getRange(1, 1, 1, 5).setValues([[
    'Replacement Cycle', 'Total Devices', 'By Location (Top)', 'By Model (Top)', 'Est. Replace Cost'
  ]]).setFontWeight('bold');
  sheet.setFrozenRows(1);

  // Summary by replacement cycle year
  const formula = '=LET(\n' +
    '  cycles, SORT(UNIQUE(FILTER(AssetData!Y2:Y, AssetData!Y2:Y<>""))),\n' +
    '  total, BYROW(cycles, LAMBDA(c, COUNTIF(AssetData!Y:Y, c))),\n' +
    '  top_loc, BYROW(cycles, LAMBDA(c, IFERROR(INDEX(SORTN(UNIQUE(FILTER(AssetData!I2:I, AssetData!Y2:Y=c)), 1, 0, BYROW(UNIQUE(FILTER(AssetData!I2:I, AssetData!Y2:Y=c)), LAMBDA(loc, COUNTIFS(AssetData!I:I, loc, AssetData!Y:Y, c))), FALSE), 1), ""))),\n' +
    '  top_model, BYROW(cycles, LAMBDA(c, IFERROR(INDEX(SORTN(UNIQUE(FILTER(AssetData!E2:E, AssetData!Y2:Y=c)), 1, 0, BYROW(UNIQUE(FILTER(AssetData!E2:E, AssetData!Y2:Y=c)), LAMBDA(m, COUNTIFS(AssetData!E:E, m, AssetData!Y:Y, c))), FALSE), 1), ""))),\n' +
    '  est_cost, BYROW(cycles, LAMBDA(c, IFERROR(SUMIFS(AssetData!P:P, AssetData!Y:Y, c), 0))),\n' +
    '  HSTACK(cycles, total, top_loc, top_model, est_cost)\n' +
    ')';

  sheet.getRange(2, 1).setFormula(formula);
}

function setupBudgetPlanningSheet(ss) {
  deleteSheetIfExists(ss, 'BudgetPlanning');
  const sheet = ss.insertSheet('BudgetPlanning');
  sheet.getRange(1, 1, 1, 7).setValues([[
    'Location', 'AUE Expired', 'AUE < 1 Year', 'AUE < 2 Years', 'Total Needing Replace', 'Avg Purchase Price', 'Est. Replacement Cost'
  ]]).setFontWeight('bold');
  sheet.setFrozenRows(1);

  const formula = '=LET(\n' +
    '  locs, UNIQUE(FILTER(AssetData!I2:I, AssetData!I2:I<>"")),\n' +
    '  expired, BYROW(locs, LAMBDA(loc, COUNTIFS(AssetData!I:I, loc, AssetData!U:U, "Expired"))),\n' +
    '  yr1, BYROW(locs, LAMBDA(loc, COUNTIFS(AssetData!I:I, loc, AssetData!U:U, "< 6 Months")+COUNTIFS(AssetData!I:I, loc, AssetData!U:U, "< 1 Year"))),\n' +
    '  yr2, BYROW(locs, LAMBDA(loc, COUNTIFS(AssetData!I:I, loc, AssetData!U:U, "< 2 Years"))),\n' +
    '  need_replace, expired + yr1,\n' +
    '  avg_price, BYROW(locs, LAMBDA(loc, IFERROR(AVERAGEIFS(AssetData!P:P, AssetData!I:I, loc, AssetData!P:P, ">0"), 0))),\n' +
    '  est_cost, need_replace * avg_price,\n' +
    '  SORT(HSTACK(locs, expired, yr1, yr2, need_replace, avg_price, est_cost), 5, FALSE)\n' +
    ')';

  sheet.getRange(2, 1).setFormula(formula);
}

function setupStatusOverviewSheet(ss) {
  deleteSheetIfExists(ss, 'StatusOverview');
  const sheet = ss.insertSheet('StatusOverview');
  sheet.getRange(1, 1, 1, 3).setValues([['Status', 'Count', '% of Total']]).setFontWeight('bold');
  sheet.setFrozenRows(1);

  const formula = '=LET(\n' +
    '  statuses, UNIQUE(FILTER(AssetData!M2:M, AssetData!M2:M<>"")),\n' +
    '  counts, BYROW(statuses, LAMBDA(s, COUNTIF(AssetData!M:M, s))),\n' +
    '  total, SUM(counts),\n' +
    '  pct, counts / total,\n' +
    '  SORT(HSTACK(statuses, counts, pct), 2, FALSE)\n' +
    ')';

  sheet.getRange(2, 1).setFormula(formula);
  sheet.getRange('C:C').setNumberFormat('0.0%');
}
