/**
 * iiQ Asset Reporting - Optional Analytics Sheets
 * Additional analytics sheets installed individually via menu.
 * NOT created by Setup Spreadsheet or included in "Regenerate All Analytics".
 */

// =============================================================================
// WARRANTY TIMELINE
// "When does warranty expire by cohort?"
// Monthly/quarterly view of upcoming warranty expirations for procurement planning.
// =============================================================================

function setupWarrantyTimelineSheet(ss) {
  const { sheet, isNew } = getOrCreateSheet(ss, 'WarrantyTimeline', [[
    'Expiration Quarter', 'Devices Expiring', 'Active Devices', 'Total Value', 'Avg Age (Years)'
  ]], '#f1663c');

  const qtr = 'IF(AssetData!O2:O<>"",TEXT(AssetData!O2:O,"YYYY")&"-Q"&(INT((MONTH(AssetData!O2:O)-1)/3)+1),"")';

  const formula = '=LET(\n' +
    '  quarters, SORT(UNIQUE(FILTER(' + qtr + ', ' + qtr + '<>"")),1,TRUE),\n' +
    '  total, BYROW(quarters, LAMBDA(q, SUMPRODUCT((' + qtr + '=q)*1))),\n' +
    '  active, BYROW(quarters, LAMBDA(q, SUMPRODUCT((' + qtr + '=q)*(AssetData!M2:M<>"Retired")*(AssetData!M2:M<>"")*1))),\n' +
    '  value, BYROW(quarters, LAMBDA(q, SUMPRODUCT((' + qtr + '=q)*IF(ISNUMBER(AssetData!P2:P),AssetData!P2:P,0)))),\n' +
    '  avg_age, BYROW(quarters, LAMBDA(q, IFERROR(SUMPRODUCT((' + qtr + '=q)*IF(ISNUMBER(AssetData!AA2:AA),AssetData!AA2:AA,0))/SUMPRODUCT((' + qtr + '=q)*1),0))),\n' +
    '  IFERROR(SORT(HSTACK(quarters, total, active, value, avg_age),1,TRUE),\n' +
    '    HSTACK(quarters, total, active, value, avg_age))\n' +
    ')';

  sheet.getRange(2, 1).setFormula(formula);
  if (isNew) sheet.getRange('D:D').setNumberFormat('$#,##0');
}

// =============================================================================
// REPLACEMENT FORECAST
// "How many devices need replacing in the next 1/2/3 years?"
// Projects future replacement volume and cost based on 5-year lifecycle.
// =============================================================================

function setupReplacementForecastSheet(ss) {
  const { sheet, isNew } = getOrCreateSheet(ss, 'ReplacementForecast', [[
    'Projected Replacement Year', 'Device Count', 'Avg Purchase Price', 'Est. Replacement Cost'
  ]], '#f1663c');

  const replYr = 'IFERROR(YEAR(IF(AssetData!N2:N<>"",AssetData!N2:N,AssetData!Q2:Q))+5,0)';
  const nr = '(AssetData!M2:M<>"Retired")*(AssetData!M2:M<>"")';

  const formula = '=LET(\n' +
    '  years, SORT(UNIQUE(FILTER(' + replYr + ', (' + replYr + '>0)*' + nr + ')),1,TRUE),\n' +
    '  counts, BYROW(years, LAMBDA(y, SUMPRODUCT((' + replYr + '=y)*' + nr + '*1))),\n' +
    '  avg_price, BYROW(years, LAMBDA(y, IFERROR(\n' +
    '    SUMPRODUCT((' + replYr + '=y)*' + nr + '*IF(ISNUMBER(AssetData!P2:P),AssetData!P2:P,0))\n' +
    '    /SUMPRODUCT((' + replYr + '=y)*' + nr + '*(AssetData!P2:P>0)*1),0))),\n' +
    '  est_cost, BYROW(years, LAMBDA(y, IFERROR(\n' +
    '    SUMPRODUCT((' + replYr + '=y)*' + nr + '*1)\n' +
    '    *SUMPRODUCT((' + replYr + '=y)*' + nr + '*IF(ISNUMBER(AssetData!P2:P),AssetData!P2:P,0))\n' +
    '    /SUMPRODUCT((' + replYr + '=y)*' + nr + '*(AssetData!P2:P>0)*1),0))),\n' +
    '  IFERROR(SORT(HSTACK(years, counts, avg_price, est_cost),1,TRUE),\n' +
    '    HSTACK(years, counts, avg_price, est_cost))\n' +
    ')';

  sheet.getRange(2, 1).setFormula(formula);
  if (isNew) {
    sheet.getRange('C:C').setNumberFormat('$#,##0');
    sheet.getRange('D:D').setNumberFormat('$#,##0');
  }
}

// =============================================================================
// UNASSIGNED INVENTORY
// "Where is idle inventory sitting?"
// Devices not assigned to anyone, broken out by location.
// =============================================================================

function setupUnassignedInventorySheet(ss) {
  const { sheet, isNew } = getOrCreateSheet(ss, 'UnassignedInventory', [[
    'Location', 'Unassigned Devices', 'Active Unassigned', 'Avg Age (Years)', 'Est. Value'
  ]], '#f1663c');

  const formula = '=LET(\n' +
    '  locs, UNIQUE(FILTER(AssetData!I2:I, (AssetData!I2:I<>"")*(AssetData!K2:K=""))),\n' +
    '  total, BYROW(locs, LAMBDA(loc, COUNTIFS(AssetData!I:I, loc, AssetData!K:K, ""))),\n' +
    '  active, BYROW(locs, LAMBDA(loc, COUNTIFS(AssetData!I:I, loc, AssetData!K:K, "", AssetData!M:M, "<>Retired"))),\n' +
    '  avg_age, BYROW(locs, LAMBDA(loc, IFERROR(AVERAGEIFS(AssetData!AA:AA, AssetData!I:I, loc, AssetData!K:K, ""), 0))),\n' +
    '  value, BYROW(locs, LAMBDA(loc, SUMIFS(AssetData!P:P, AssetData!I:I, loc, AssetData!K:K, ""))),\n' +
    '  IFERROR(SORT(HSTACK(locs, total, active, avg_age, value), 2, FALSE),\n' +
    '    HSTACK(locs, total, active, avg_age, value))\n' +
    ')';

  sheet.getRange(2, 1).setFormula(formula);
  if (isNew) sheet.getRange('E:E').setNumberFormat('$#,##0');
}

// =============================================================================
// DEVICE LIFECYCLE
// "How long do devices actually last by model?"
// Average age at retirement per model/manufacturer.
// =============================================================================

function setupDeviceLifecycleSheet(ss) {
  const { sheet } = getOrCreateSheet(ss, 'DeviceLifecycle', [[
    'Model', 'Manufacturer', 'Retired Count', 'Avg Lifespan (Years)', 'Still Active Count'
  ]], '#f1663c');

  const formula = '=LET(\n' +
    '  models, UNIQUE(FILTER(AssetData!E2:E, (AssetData!E2:E<>"")*(AssetData!M2:M="Retired"))),\n' +
    '  mfr, BYROW(models, LAMBDA(m, IFERROR(INDEX(FILTER(AssetData!F:F, AssetData!E:E=m), 1), ""))),\n' +
    '  retired, BYROW(models, LAMBDA(m, COUNTIFS(AssetData!E:E, m, AssetData!M:M, "Retired"))),\n' +
    '  avg_life, BYROW(models, LAMBDA(m, IFERROR(AVERAGEIFS(AssetData!AA:AA, AssetData!E:E, m, AssetData!M:M, "Retired"), 0))),\n' +
    '  active, BYROW(models, LAMBDA(m, COUNTIFS(AssetData!E:E, m, AssetData!M:M, "<>Retired"))),\n' +
    '  IFERROR(SORT(HSTACK(models, mfr, retired, avg_life, active), 4, FALSE),\n' +
    '    HSTACK(models, mfr, retired, avg_life, active))\n' +
    ')';

  sheet.getRange(2, 1).setFormula(formula);
}

// =============================================================================
// CATEGORY BREAKDOWN
// "What types of devices do we have?"
// Chromebooks vs laptops vs tablets vs peripherals by category.
// =============================================================================

function setupCategoryBreakdownSheet(ss) {
  const { sheet, isNew } = getOrCreateSheet(ss, 'CategoryBreakdown', [[
    'Category', 'Total', 'Active', 'Retired', 'Avg Age (Years)', 'Total Value'
  ]], '#f1663c');

  const formula = '=LET(\n' +
    '  cats, UNIQUE(FILTER(AssetData!G2:G, AssetData!G2:G<>"")),\n' +
    '  total, BYROW(cats, LAMBDA(c, COUNTIF(AssetData!G:G, c))),\n' +
    '  active, BYROW(cats, LAMBDA(c, COUNTIFS(AssetData!G:G, c, AssetData!M:M, "<>Retired"))),\n' +
    '  retired, BYROW(cats, LAMBDA(c, COUNTIFS(AssetData!G:G, c, AssetData!M:M, "Retired"))),\n' +
    '  avg_age, BYROW(cats, LAMBDA(c, IFERROR(AVERAGEIFS(AssetData!AA:AA, AssetData!G:G, c), 0))),\n' +
    '  value, BYROW(cats, LAMBDA(c, SUMIF(AssetData!G:G, c, AssetData!P:P))),\n' +
    '  IFERROR(SORT(HSTACK(cats, total, active, retired, avg_age, value), 2, FALSE),\n' +
    '    HSTACK(cats, total, active, retired, avg_age, value))\n' +
    ')';

  sheet.getRange(2, 1).setFormula(formula);
  if (isNew) sheet.getRange('F:F').setNumberFormat('$#,##0');
}

// =============================================================================
// MANUFACTURER SUMMARY
// "Which vendors are we invested in?"
// Device count, avg age, warranty coverage, and ticket rate by manufacturer.
// =============================================================================

function setupManufacturerSummarySheet(ss) {
  const { sheet, isNew } = getOrCreateSheet(ss, 'ManufacturerSummary', [[
    'Manufacturer', 'Device Count', 'Avg Age (Years)', 'Warranty Active', 'Warranty Expired', 'Total Open Tickets', 'Tickets / Device'
  ]], '#f1663c');

  const formula = '=LET(\n' +
    '  mfrs, UNIQUE(FILTER(AssetData!F2:F, AssetData!F2:F<>"")),\n' +
    '  total, BYROW(mfrs, LAMBDA(m, COUNTIF(AssetData!F:F, m))),\n' +
    '  avg_age, BYROW(mfrs, LAMBDA(m, IFERROR(AVERAGEIFS(AssetData!AA:AA, AssetData!F:F, m), 0))),\n' +
    '  warr_active, BYROW(mfrs, LAMBDA(m, COUNTIFS(AssetData!F:F, m, AssetData!AB:AB, "Active"))),\n' +
    '  warr_expired, BYROW(mfrs, LAMBDA(m, COUNTIFS(AssetData!F:F, m, AssetData!AB:AB, "Expired"))),\n' +
    '  tickets, BYROW(mfrs, LAMBDA(m, SUMIF(AssetData!F:F, m, AssetData!Y:Y))),\n' +
    '  tix_per, BYROW(mfrs, LAMBDA(m, IFERROR(SUMIF(AssetData!F:F, m, AssetData!Y:Y)/COUNTIF(AssetData!F:F, m), 0))),\n' +
    '  IFERROR(SORT(HSTACK(mfrs, total, avg_age, warr_active, warr_expired, tickets, tix_per), 2, FALSE),\n' +
    '    HSTACK(mfrs, total, avg_age, warr_active, warr_expired, tickets, tix_per))\n' +
    ')';

  sheet.getRange(2, 1).setFormula(formula);
  if (isNew) sheet.getRange('G:G').setNumberFormat('0.00');
}

// =============================================================================
// HIGH TICKET LOCATIONS
// "Which schools have the most device problems?"
// Combines ticket counts with device counts to find disproportionate issues.
// =============================================================================

function setupHighTicketLocationsSheet(ss) {
  const { sheet, isNew } = getOrCreateSheet(ss, 'HighTicketLocations', [[
    'Location', 'Total Devices', 'Active Devices', 'Open Tickets', 'Tickets / Device', 'Avg Age (Years)'
  ]], '#f1663c');

  const formula = '=LET(\n' +
    '  locs, UNIQUE(FILTER(AssetData!I2:I, AssetData!I2:I<>"")),\n' +
    '  total, BYROW(locs, LAMBDA(loc, COUNTIF(AssetData!I:I, loc))),\n' +
    '  active, BYROW(locs, LAMBDA(loc, COUNTIFS(AssetData!I:I, loc, AssetData!M:M, "<>Retired"))),\n' +
    '  tickets, BYROW(locs, LAMBDA(loc, SUMIF(AssetData!I:I, loc, AssetData!Y:Y))),\n' +
    '  tix_per, BYROW(locs, LAMBDA(loc, IFERROR(SUMIF(AssetData!I:I, loc, AssetData!Y:Y)/COUNTIF(AssetData!I:I, loc), 0))),\n' +
    '  avg_age, BYROW(locs, LAMBDA(loc, IFERROR(AVERAGEIFS(AssetData!AA:AA, AssetData!I:I, loc), 0))),\n' +
    '  IFERROR(SORT(HSTACK(locs, total, active, tickets, tix_per, avg_age), 5, FALSE),\n' +
    '    HSTACK(locs, total, active, tickets, tix_per, avg_age))\n' +
    ')';

  sheet.getRange(2, 1).setFormula(formula);
  if (isNew) sheet.getRange('E:E').setNumberFormat('0.00');
}

// =============================================================================
// LOST/STOLEN RATE
// "Which schools are losing devices?"
// Per-location lost/stolen counts and rate, sorted by rate descending.
// =============================================================================

function setupLostStolenRateSheet(ss) {
  const { sheet, isNew } = getOrCreateSheet(ss, 'LostStolenRate', [[
    'Location', 'Total Devices', 'Lost', 'Stolen', 'Lost + Stolen', 'Rate (%)'
  ]], '#f1663c');

  const formula = '=LET(\n' +
    '  locs, UNIQUE(FILTER(AssetData!I2:I, AssetData!I2:I<>"")),\n' +
    '  total, BYROW(locs, LAMBDA(loc, COUNTIF(AssetData!I:I, loc))),\n' +
    '  lost, BYROW(locs, LAMBDA(loc, COUNTIFS(AssetData!I:I, loc, AssetData!M:M, "*Lost*"))),\n' +
    '  stolen, BYROW(locs, LAMBDA(loc, COUNTIFS(AssetData!I:I, loc, AssetData!M:M, "*Stolen*"))),\n' +
    '  combined, BYROW(locs, LAMBDA(loc, COUNTIFS(AssetData!I:I, loc, AssetData!M:M, "*Lost*") + COUNTIFS(AssetData!I:I, loc, AssetData!M:M, "*Stolen*"))),\n' +
    '  rate, BYROW(locs, LAMBDA(loc, IFERROR((COUNTIFS(AssetData!I:I, loc, AssetData!M:M, "*Lost*") + COUNTIFS(AssetData!I:I, loc, AssetData!M:M, "*Stolen*")) / COUNTIF(AssetData!I:I, loc), 0))),\n' +
    '  IFERROR(SORT(HSTACK(locs, total, lost, stolen, combined, rate), 6, FALSE),\n' +
    '    HSTACK(locs, total, lost, stolen, combined, rate))\n' +
    ')';

  sheet.getRange(2, 1).setFormula(formula);
  if (isNew) sheet.getRange('F:F').setNumberFormat('0.0%');
}

// =============================================================================
// SPARE ASSETS
// "Do I have enough working spares at each school?"
// Unassigned devices by location: deployable vs non-deployable, in storage
// vs unaccounted, spare ratio relative to assigned devices.
// =============================================================================

function setupSpareAssetsSheet(ss) {
  const { sheet, isNew } = getOrCreateSheet(ss, 'SpareAssets', [[
    'Location', 'Assigned Devices', 'Total Unassigned', 'Deployable Spares',
    'Non-Deployable', 'In Storage', 'Unaccounted', 'Spare Ratio (%)', 'Top Spare Model'
  ]], '#f1663c');

  const formula = '=LET(\n' +
    '  locs, UNIQUE(FILTER(AssetData!I2:I, AssetData!I2:I<>"")),\n' +
    '  assigned, BYROW(locs, LAMBDA(loc, COUNTIFS(AssetData!I:I, loc, AssetData!K:K, "<>"))),\n' +
    '  unassigned, BYROW(locs, LAMBDA(loc, COUNTIFS(AssetData!I:I, loc, AssetData!K:K, ""))),\n' +
    '  deployable, BYROW(locs, LAMBDA(loc, COUNTIFS(AssetData!I:I, loc, AssetData!K:K, "", AssetData!M:M, "<>Retired"))),\n' +
    '  non_deploy, BYROW(locs, LAMBDA(loc, COUNTIFS(AssetData!I:I, loc, AssetData!K:K, "") - COUNTIFS(AssetData!I:I, loc, AssetData!K:K, "", AssetData!M:M, "<>Retired"))),\n' +
    '  in_storage, BYROW(locs, LAMBDA(loc, SUMPRODUCT((AssetData!I2:I=loc)*(AssetData!K2:K="")*((AssetData!V2:V<>"")+((AssetData!W2:W<>"")*1)>0)*1))),\n' +
    '  unaccounted, BYROW(locs, LAMBDA(loc,\n' +
    '    COUNTIFS(AssetData!I:I, loc, AssetData!K:K, "", AssetData!M:M, "<>Retired")\n' +
    '    - SUMPRODUCT((AssetData!I2:I=loc)*(AssetData!K2:K="")*(AssetData!M2:M<>"Retired")*((AssetData!V2:V<>"")+((AssetData!W2:W<>"")*1)>0)*1))),\n' +
    '  ratio, BYROW(locs, LAMBDA(loc, IFERROR(\n' +
    '    COUNTIFS(AssetData!I:I, loc, AssetData!K:K, "", AssetData!M:M, "<>Retired")\n' +
    '    / COUNTIFS(AssetData!I:I, loc, AssetData!K:K, "<>"), 0))),\n' +
    '  top_model, BYROW(locs, LAMBDA(loc, IFERROR(\n' +
    '    INDEX(FILTER(AssetData!E2:E, (AssetData!I2:I=loc)*(AssetData!K2:K="")*(AssetData!M2:M<>"Retired")*(AssetData!E2:E<>"")),\n' +
    '      MODE(MATCH(FILTER(AssetData!E2:E, (AssetData!I2:I=loc)*(AssetData!K2:K="")*(AssetData!M2:M<>"Retired")*(AssetData!E2:E<>"")),\n' +
    '        FILTER(AssetData!E2:E, (AssetData!I2:I=loc)*(AssetData!K2:K="")*(AssetData!M2:M<>"Retired")*(AssetData!E2:E<>"")), 0))),\n' +
    '    ""))),\n' +
    '  IFERROR(SORT(HSTACK(locs, assigned, unassigned, deployable, non_deploy, in_storage, unaccounted, ratio, top_model), 8, TRUE),\n' +
    '    HSTACK(locs, assigned, unassigned, deployable, non_deploy, in_storage, unaccounted, ratio, top_model))\n' +
    ')';

  sheet.getRange(2, 1).setFormula(formula);
  if (isNew) sheet.getRange('H:H').setNumberFormat('0.0%');
}

// =============================================================================
// BREAK RATE
// "Which devices and models have the most tickets?"
// Two views: device-level top 100 by ticket count + model-level avg tickets.
// =============================================================================

function setupBreakRateSheet(ss) {
  // BreakRate has two header regions, so handle creation manually
  let sheet = ss.getSheetByName('BreakRate');
  const isNew = !sheet;
  if (sheet) {
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
    }
  } else {
    sheet = ss.insertSheet('BreakRate');
    sheet.getRange(1, 1, 1, 7).setValues([[
      'Asset Tag', 'Serial Number', 'Model', 'Manufacturer', 'Location', 'Open Tickets', 'Status'
    ]]).setFontWeight('bold');
    sheet.getRange(1, 9, 1, 5).setValues([[
      'Model', 'Device Count', 'Total Open Tickets', 'Avg Tickets/Device', 'Max Tickets'
    ]]).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.getRange('L:L').setNumberFormat('0.00');
    sheet.setTabColor('#f1663c');
  }

  const deviceFormula = '=IFERROR(ARRAY_CONSTRAIN(SORT(\n' +
    '  FILTER(HSTACK(AssetData!B2:B, AssetData!D2:D, AssetData!E2:E, AssetData!F2:F, AssetData!I2:I, AssetData!Y2:Y, AssetData!M2:M),\n' +
    '    ISNUMBER(AssetData!Y2:Y)*(AssetData!Y2:Y>0)),\n' +
    '  6, FALSE), 100, 7), "")';

  const modelFormula = '=LET(\n' +
    '  models, UNIQUE(FILTER(AssetData!E2:E, AssetData!E2:E<>"")),\n' +
    '  dev_count, BYROW(models, LAMBDA(m, COUNTIF(AssetData!E:E, m))),\n' +
    '  tot_tix, BYROW(models, LAMBDA(m, SUMIF(AssetData!E:E, m, AssetData!Y:Y))),\n' +
    '  avg_tix, BYROW(models, LAMBDA(m, IFERROR(SUMIF(AssetData!E:E, m, AssetData!Y:Y)/COUNTIF(AssetData!E:E, m), 0))),\n' +
    '  max_tix, BYROW(models, LAMBDA(m, MAXIFS(AssetData!Y:Y, AssetData!E:E, m))),\n' +
    '  IFERROR(SORT(HSTACK(models, dev_count, tot_tix, avg_tix, max_tix), 4, FALSE),\n' +
    '    HSTACK(models, dev_count, tot_tix, avg_tix, max_tix))\n' +
    ')';

  sheet.getRange(2, 1).setFormula(deviceFormula);
  sheet.getRange(2, 9).setFormula(modelFormula);
}

// =============================================================================
// DEVICE READINESS
// "What's actually deployable at each school right now?"
// Per-location breakdown: deployable vs in-repair vs lost/stolen vs retired.
// =============================================================================

function setupDeviceReadinessSheet(ss) {
  const { sheet, isNew } = getOrCreateSheet(ss, 'DeviceReadiness', [[
    'Location', 'Total Devices', 'Deployable', 'In Repair', 'Lost/Stolen',
    'Retired', 'Readiness Rate (%)'
  ]], '#f1663c');

  const formula = '=LET(\n' +
    '  locs, UNIQUE(FILTER(AssetData!I2:I, AssetData!I2:I<>"")),\n' +
    '  total, BYROW(locs, LAMBDA(loc, COUNTIF(AssetData!I:I, loc))),\n' +
    '  repair, BYROW(locs, LAMBDA(loc, COUNTIFS(AssetData!I:I, loc, AssetData!M:M, "*Repair*"))),\n' +
    '  lost_stolen, BYROW(locs, LAMBDA(loc,\n' +
    '    COUNTIFS(AssetData!I:I, loc, AssetData!M:M, "*Lost*")\n' +
    '    + COUNTIFS(AssetData!I:I, loc, AssetData!M:M, "*Stolen*"))),\n' +
    '  retired, BYROW(locs, LAMBDA(loc, COUNTIFS(AssetData!I:I, loc, AssetData!M:M, "Retired"))),\n' +
    '  deployable, BYROW(locs, LAMBDA(loc,\n' +
    '    COUNTIF(AssetData!I:I, loc)\n' +
    '    - COUNTIFS(AssetData!I:I, loc, AssetData!M:M, "*Repair*")\n' +
    '    - COUNTIFS(AssetData!I:I, loc, AssetData!M:M, "*Lost*")\n' +
    '    - COUNTIFS(AssetData!I:I, loc, AssetData!M:M, "*Stolen*")\n' +
    '    - COUNTIFS(AssetData!I:I, loc, AssetData!M:M, "Retired"))),\n' +
    '  rate, BYROW(locs, LAMBDA(loc, IFERROR(\n' +
    '    (COUNTIF(AssetData!I:I, loc)\n' +
    '    - COUNTIFS(AssetData!I:I, loc, AssetData!M:M, "*Repair*")\n' +
    '    - COUNTIFS(AssetData!I:I, loc, AssetData!M:M, "*Lost*")\n' +
    '    - COUNTIFS(AssetData!I:I, loc, AssetData!M:M, "*Stolen*")\n' +
    '    - COUNTIFS(AssetData!I:I, loc, AssetData!M:M, "Retired"))\n' +
    '    / COUNTIF(AssetData!I:I, loc), 0))),\n' +
    '  IFERROR(SORT(HSTACK(locs, total, deployable, repair, lost_stolen, retired, rate), 7, TRUE),\n' +
    '    HSTACK(locs, total, deployable, repair, lost_stolen, retired, rate))\n' +
    ')';

  sheet.getRange(2, 1).setFormula(formula);
  if (isNew) sheet.getRange('G:G').setNumberFormat('0.0%');
}

// =============================================================================
// MODEL FRAGMENTATION
// "Which schools are a patchwork of device models?"
// Distinct model count per location + top model share. More models = harder
// to manage spares, imaging, and teacher training.
// =============================================================================

function setupModelFragmentationSheet(ss) {
  const { sheet, isNew } = getOrCreateSheet(ss, 'ModelFragmentation', [[
    'Location', 'Total Devices', 'Distinct Models', 'Top Model',
    'Top Model Count', 'Top Model Share (%)'
  ]], '#f1663c');

  const formula = '=LET(\n' +
    '  locs, UNIQUE(FILTER(AssetData!I2:I, AssetData!I2:I<>"")),\n' +
    '  total, BYROW(locs, LAMBDA(loc, COUNTIF(AssetData!I:I, loc))),\n' +
    '  distinct, BYROW(locs, LAMBDA(loc, IFERROR(ROWS(UNIQUE(FILTER(AssetData!E2:E, (AssetData!I2:I=loc)*(AssetData!E2:E<>"")))), 0))),\n' +
    '  top_model, BYROW(locs, LAMBDA(loc, IFERROR(\n' +
    '    INDEX(FILTER(AssetData!E2:E, (AssetData!I2:I=loc)*(AssetData!E2:E<>"")),\n' +
    '      MODE(MATCH(FILTER(AssetData!E2:E, (AssetData!I2:I=loc)*(AssetData!E2:E<>"")),\n' +
    '        FILTER(AssetData!E2:E, (AssetData!I2:I=loc)*(AssetData!E2:E<>"")), 0))),\n' +
    '    ""))),\n' +
    '  top_count, BYROW(SEQUENCE(ROWS(locs)), LAMBDA(i,\n' +
    '    IFERROR(COUNTIFS(AssetData!I:I, INDEX(locs, i), AssetData!E:E, INDEX(top_model, i)), 0))),\n' +
    '  top_share, BYROW(SEQUENCE(ROWS(locs)), LAMBDA(i,\n' +
    '    IFERROR(INDEX(top_count, i) / INDEX(total, i), 0))),\n' +
    '  IFERROR(SORT(HSTACK(locs, total, distinct, top_model, top_count, top_share), 3, FALSE),\n' +
    '    HSTACK(locs, total, distinct, top_model, top_count, top_share))\n' +
    ')';

  sheet.getRange(2, 1).setFormula(formula);
  if (isNew) sheet.getRange('F:F').setNumberFormat('0.0%');
}

// =============================================================================
// REPLACEMENT PLANNING
// "What do I need to buy before next school year?"
// Per-location devices crossing the age threshold by the target date.
// References Config keys: REPLACEMENT_AGE_YEARS, NEXT_SCHOOL_YEAR_START.
// =============================================================================

function setupReplacementPlanningSheet(ss) {
  const { sheet, isNew } = getOrCreateSheet(ss, 'ReplacementPlanning', [[
    'Location', 'Total Active', 'Currently Over Threshold', 'Over by Target Date',
    'New Replacements Needed', 'Currently Over (%)', 'By Target (%)', 'Est. Replacement Cost'
  ]], '#f1663c');

  const cfgAge = 'IFERROR(VLOOKUP("REPLACEMENT_AGE_YEARS",Config!A:B,2,FALSE)*1, 4)';
  const cfgDate = 'IFERROR(DATEVALUE(""&VLOOKUP("NEXT_SCHOOL_YEAR_START",Config!A:B,2,FALSE)), TODAY()+90)';

  const formula = '=LET(\n' +
    '  age_yrs, ' + cfgAge + ',\n' +
    '  target_date, ' + cfgDate + ',\n' +
    '  days_delta, (target_date - TODAY()) / 365.25,\n' +
    '  locs, UNIQUE(FILTER(AssetData!I2:I, (AssetData!I2:I<>"")*(AssetData!M2:M<>"Retired"))),\n' +
    '  active, BYROW(locs, LAMBDA(loc, COUNTIFS(AssetData!I:I, loc, AssetData!M:M, "<>Retired"))),\n' +
    '  curr_over, BYROW(locs, LAMBDA(loc, SUMPRODUCT(\n' +
    '    (AssetData!I2:I=loc)*(AssetData!M2:M<>"Retired")*(AssetData!M2:M<>"")\n' +
    '    *(ISNUMBER(AssetData!AA2:AA))*(AssetData!AA2:AA>=age_yrs)*1))),\n' +
    '  future_over, BYROW(locs, LAMBDA(loc, SUMPRODUCT(\n' +
    '    (AssetData!I2:I=loc)*(AssetData!M2:M<>"Retired")*(AssetData!M2:M<>"")\n' +
    '    *(ISNUMBER(AssetData!AA2:AA))*((AssetData!AA2:AA+days_delta)>=age_yrs)*1))),\n' +
    '  new_repl, BYROW(SEQUENCE(ROWS(locs)), LAMBDA(i, INDEX(future_over, i) - INDEX(curr_over, i))),\n' +
    '  curr_pct, BYROW(SEQUENCE(ROWS(locs)), LAMBDA(i, IFERROR(INDEX(curr_over, i) / INDEX(active, i), 0))),\n' +
    '  future_pct, BYROW(SEQUENCE(ROWS(locs)), LAMBDA(i, IFERROR(INDEX(future_over, i) / INDEX(active, i), 0))),\n' +
    '  est_cost, BYROW(SEQUENCE(ROWS(locs)), LAMBDA(i, IFERROR(\n' +
    '    INDEX(new_repl, i) * AVERAGEIFS(AssetData!P:P, AssetData!I:I, INDEX(locs, i), AssetData!P:P, ">0"), 0))),\n' +
    '  IFERROR(SORT(HSTACK(locs, active, curr_over, future_over, new_repl, curr_pct, future_pct, est_cost), 5, FALSE),\n' +
    '    HSTACK(locs, active, curr_over, future_over, new_repl, curr_pct, future_pct, est_cost))\n' +
    ')';

  sheet.getRange(2, 1).setFormula(formula);
  if (isNew) {
    sheet.getRange('F:F').setNumberFormat('0.0%');
    sheet.getRange('G:G').setNumberFormat('0.0%');
    sheet.getRange('H:H').setNumberFormat('$#,##0');
  }
}

// =============================================================================
// LOCATION MODEL BREAKDOWN
// "What models are at each school?"
// Flat cross-tab: every Location × Model combination with counts.
// =============================================================================

function setupLocationModelBreakdownSheet(ss) {
  const { sheet } = getOrCreateSheet(ss, 'LocationModelBreakdown', [[
    'Location', 'Model', 'Manufacturer', 'Total', 'Active', 'Retired', 'Avg Age (Years)'
  ]], '#f1663c');

  const formula = '=LET(\n' +
    '  pairs, UNIQUE(FILTER(HSTACK(AssetData!I2:I, AssetData!E2:E), (AssetData!I2:I<>"")*(AssetData!E2:E<>""))),\n' +
    '  loc_col, INDEX(pairs,,1),\n' +
    '  model_col, INDEX(pairs,,2),\n' +
    '  mfr, BYROW(model_col, LAMBDA(m, IFERROR(INDEX(FILTER(AssetData!F:F, AssetData!E:E=m), 1), ""))),\n' +
    '  total, BYROW(SEQUENCE(ROWS(pairs)), LAMBDA(i,\n' +
    '    COUNTIFS(AssetData!I:I, INDEX(loc_col, i), AssetData!E:E, INDEX(model_col, i)))),\n' +
    '  active, BYROW(SEQUENCE(ROWS(pairs)), LAMBDA(i,\n' +
    '    COUNTIFS(AssetData!I:I, INDEX(loc_col, i), AssetData!E:E, INDEX(model_col, i), AssetData!M:M, "<>Retired"))),\n' +
    '  retired, BYROW(SEQUENCE(ROWS(pairs)), LAMBDA(i,\n' +
    '    COUNTIFS(AssetData!I:I, INDEX(loc_col, i), AssetData!E:E, INDEX(model_col, i), AssetData!M:M, "Retired"))),\n' +
    '  avg_age, BYROW(SEQUENCE(ROWS(pairs)), LAMBDA(i,\n' +
    '    IFERROR(AVERAGEIFS(AssetData!AA:AA, AssetData!I:I, INDEX(loc_col, i), AssetData!E:E, INDEX(model_col, i)), 0))),\n' +
    '  IFERROR(SORT(HSTACK(loc_col, model_col, mfr, total, active, retired, avg_age), 1, TRUE, 4, FALSE),\n' +
    '    HSTACK(loc_col, model_col, mfr, total, active, retired, avg_age))\n' +
    ')';

  sheet.getRange(2, 1).setFormula(formula);
}

// =============================================================================
// LOCATION MODEL FILTERED
// "Show me one school's model mix"
// Interactive dropdown-driven view — select a location, see its model breakdown.
// =============================================================================

function setupLocationModelFilteredSheet(ss) {
  // Custom layout: selector in row 1, headers in row 2, data from row 3
  let sheet = ss.getSheetByName('LocationModelFiltered');
  const isNew = !sheet;
  if (sheet) {
    if (sheet.getLastRow() > 2) {
      sheet.getRange(3, 1, sheet.getLastRow() - 2, sheet.getLastColumn()).clearContent();
    }
  } else {
    sheet = ss.insertSheet('LocationModelFiltered');
    sheet.getRange(1, 1).setValue('Select Location:').setFontWeight('bold');
    const locSheet = ss.getSheetByName('Locations');
    if (locSheet) {
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInRange(locSheet.getRange('B2:B'), true)
        .setAllowInvalid(false)
        .build();
      sheet.getRange(1, 2).setDataValidation(rule);
    }
    sheet.getRange(2, 1, 1, 6).setValues([[
      'Model', 'Manufacturer', 'Total', 'Active', 'Retired', 'Avg Age (Years)'
    ]]).setFontWeight('bold');
    sheet.setFrozenRows(2);
    sheet.setTabColor('#f1663c');
  }

  const formula = '=IF(B1="", "\u2190 Select a location",\n' +
    '  LET(\n' +
    '    sel, B1,\n' +
    '    models, UNIQUE(FILTER(AssetData!E2:E, (AssetData!I2:I=sel)*(AssetData!E2:E<>""))),\n' +
    '    mfr, BYROW(models, LAMBDA(m, IFERROR(INDEX(FILTER(AssetData!F:F, AssetData!E:E=m), 1), ""))),\n' +
    '    total, BYROW(models, LAMBDA(m, COUNTIFS(AssetData!I:I, sel, AssetData!E:E, m))),\n' +
    '    active, BYROW(models, LAMBDA(m, COUNTIFS(AssetData!I:I, sel, AssetData!E:E, m, AssetData!M:M, "<>Retired"))),\n' +
    '    retired, BYROW(models, LAMBDA(m, COUNTIFS(AssetData!I:I, sel, AssetData!E:E, m, AssetData!M:M, "Retired"))),\n' +
    '    avg_age, BYROW(models, LAMBDA(m, IFERROR(AVERAGEIFS(AssetData!AA:AA, AssetData!I:I, sel, AssetData!E:E, m), 0))),\n' +
    '    IFERROR(SORT(HSTACK(models, mfr, total, active, retired, avg_age), 3, FALSE),\n' +
    '      HSTACK(models, mfr, total, active, retired, avg_age))\n' +
    '  )\n' +
    ')';

  sheet.getRange(3, 1).setFormula(formula);
}

// =============================================================================
// INDIVIDUAL LOOKUP
// "Show me device checkout history for one person"
// Interactive dropdown-driven view — select a user, onEdit fetches history live.
// =============================================================================

const INDIVIDUAL_LOOKUP_SHEET = 'IndividualLookup';
const INDIVIDUAL_LOOKUP_HANDLER = 'onEditIndividualLookup';
const INDIVIDUAL_LOOKUP_HEADERS = [
  'Date', 'Action', 'Asset Tag', 'Serial Number', 'Model',
  'Location', 'Currently With'
];
const INDIVIDUAL_LOOKUP_DATA_COLS = INDIVIDUAL_LOOKUP_HEADERS.length;

function setupIndividualLookupSheet(ss) {
  let sheet = ss.getSheetByName(INDIVIDUAL_LOOKUP_SHEET);
  const isNew = !sheet;

  if (sheet) {
    if (sheet.getLastRow() > 2) {
      sheet.getRange(3, 1, sheet.getLastRow() - 2, INDIVIDUAL_LOOKUP_DATA_COLS).clearContent();
    }
    sheet.getRange(1, 4).clearContent();
  } else {
    sheet = ss.insertSheet(INDIVIDUAL_LOOKUP_SHEET);
    sheet.getRange(1, 1).setValue('Select User:').setFontWeight('bold');
    sheet.getRange(1, 3).setValue('Status:').setFontWeight('bold');
    sheet.getRange(2, 1, 1, INDIVIDUAL_LOOKUP_DATA_COLS)
      .setValues([INDIVIDUAL_LOOKUP_HEADERS]).setFontWeight('bold');
    sheet.setFrozenRows(2);
    // Column widths serve double-duty: row 1 labels/dropdown AND row 3+ data.
    sheet.setColumnWidth(1, 170); // Date (row 3+) / "Select User:" (row 1)
    sheet.setColumnWidth(2, 220); // Action / dropdown
    sheet.setColumnWidth(3, 110); // Asset Tag / "Status:"
    sheet.setColumnWidth(4, 200); // Serial Number / status text
    sheet.setColumnWidth(5, 260); // Model
    sheet.setColumnWidth(6, 220); // Location
    sheet.setColumnWidth(7, 180); // Currently With
    sheet.setTabColor('#f1663c');
  }

  // Column Z: sorted unique owner names, sourced from AssetData.
  // Used as the dropdown source. Hidden for a clean UI.
  sheet.getRange(1, 26).setFormula(
    '=IFERROR(SORT(UNIQUE(FILTER(AssetData!L2:L, AssetData!L2:L<>""))), "")'
  );
  sheet.hideColumns(26);

  // Data validation uses the spilled unique-names range.
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(sheet.getRange('Z1:Z'), true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(1, 2).setDataValidation(rule);

  // Placeholder message before a selection is made.
  if (isNew || !sheet.getRange(1, 2).getValue()) {
    sheet.getRange(3, 1).setValue('← Select a user in B1 to load their asset assignment history.');
  }

  installIndividualLookupTrigger(ss);
}

/**
 * Installable onEdit handler. Fires on every edit in the spreadsheet — must
 * guard tightly by sheet + cell before doing any work.
 */
function onEditIndividualLookup(e) {
  if (!e || !e.range) return;
  const range = e.range;
  if (range.getSheet().getName() !== INDIVIDUAL_LOOKUP_SHEET) return;
  if (range.getRow() !== 1 || range.getColumn() !== 2) return;

  const ss = range.getSheet().getParent();
  const sheet = range.getSheet();
  const selectedName = (e.value || range.getValue() || '').toString().trim();

  sheet.getRange(1, 4).setValue('Loading…');
  if (sheet.getLastRow() > 2) {
    sheet.getRange(3, 1, sheet.getLastRow() - 2, INDIVIDUAL_LOOKUP_DATA_COLS).clearContent();
  }

  if (!selectedName) {
    sheet.getRange(1, 4).clearContent();
    sheet.getRange(3, 1).setValue('← Select a user in B1 to load their assignment history.');
    return;
  }

  try {
    const ownerId = resolveOwnerIdByName(ss, selectedName);
    if (!ownerId) {
      sheet.getRange(1, 4).setValue('No OwnerId for that name in AssetData.');
      sheet.getRange(3, 1).setValue('That user could not be resolved. They may no longer own any asset.');
      return;
    }

    const response = getUserActivities(ownerId);
    const items = (response && response.Items) || [];

    // Build lookup maps once so parseActivityRow doesn't re-scan sheets.
    const assetMap = buildAssetTagMap(ss);
    const locationMap = buildLocationMap(ss);

    const rows = items
      .map(item => parseActivityRow(item, assetMap, locationMap))
      .filter(row => row !== null);

    if (rows.length === 0) {
      sheet.getRange(1, 4).setValue('No asset assignment history found.');
      sheet.getRange(3, 1).setValue('No asset assignment events returned for ' + selectedName + '.');
      return;
    }

    rows.sort((a, b) => (b[0] > a[0] ? 1 : b[0] < a[0] ? -1 : 0));
    sheet.getRange(3, 1, rows.length, INDIVIDUAL_LOOKUP_DATA_COLS).setValues(rows);

    const scanned = items.length;
    const suffix = scanned >= 500 ? ' (scanned 500 most-recent activities)' : '';
    sheet.getRange(1, 4).setValue(`${rows.length} asset event${rows.length === 1 ? '' : 's'}${suffix} — ${new Date().toLocaleString()}`);
  } catch (err) {
    sheet.getRange(1, 4).setValue('Error: ' + (err.message || err));
    sheet.getRange(3, 1).setValue('Failed to fetch assignment history. See Logs sheet for details.');
    try { logOperation('IndividualLookup', 'ERROR', (err.message || err).toString().substring(0, 200)); } catch (_) {}
  }
}

function resolveOwnerIdByName(ss, name) {
  const assetData = ss.getSheetByName('AssetData');
  if (!assetData || assetData.getLastRow() < 2) return null;
  // Columns K=OwnerId (11), L=OwnerName (12). Read both as a single range for speed.
  const values = assetData.getRange(2, 11, assetData.getLastRow() - 1, 2).getValues();
  for (let i = 0; i < values.length; i++) {
    if (values[i][1] === name && values[i][0]) return values[i][0];
  }
  return null;
}

function buildAssetTagMap(ss) {
  const assetData = ss.getSheetByName('AssetData');
  const map = {};
  if (!assetData || assetData.getLastRow() < 2) return map;
  // Columns A-L: AssetId, AssetTag, Name, SerialNumber, ModelName,
  // ManufacturerName, CategoryName, LocationId, LocationName, LocationType,
  // OwnerId, OwnerName.
  const rows = assetData.getRange(2, 1, assetData.getLastRow() - 1, 12).getValues();
  for (let i = 0; i < rows.length; i++) {
    const tag = rows[i][1];
    if (tag && !map[tag]) {
      map[tag] = {
        serial: rows[i][3],
        model: rows[i][4],
        location: rows[i][8],
        currentOwner: rows[i][11]
      };
    }
  }
  return map;
}

function buildLocationMap(ss) {
  const locSheet = ss.getSheetByName('Locations');
  const map = {};
  if (!locSheet || locSheet.getLastRow() < 2) return map;
  // Columns A=LocationId, B=Name.
  const rows = locSheet.getRange(2, 1, locSheet.getLastRow() - 1, 2).getValues();
  for (let i = 0; i < rows.length; i++) {
    if (rows[i][0]) map[rows[i][0]] = rows[i][1];
  }
  return map;
}

/**
 * Parse one UserActivity item. Returns a sheet row array, or null if the item
 * isn't an asset event. `Details` is a JSON-encoded string of entries shaped
 * like {p, o, c} — the "p" field identifies what changed (e.g. "Asset #TAG"
 * or "OwnerId" or "LocationId"), "o" is old value, "c" is new value.
 */
function parseActivityRow(item, assetMap, locationMap) {
  if (!item || !item.Details || item.Details.indexOf('Asset #') === -1) return null;

  let details;
  try { details = JSON.parse(item.Details); } catch (_) { return null; }
  if (!Array.isArray(details)) return null;

  const assetEntry = details.find(d => d && typeof d.p === 'string' && d.p.indexOf('Asset #') === 0);
  if (!assetEntry) return null;
  const tag = assetEntry.p.replace('Asset #', '').trim();

  const ownerEntry = details.find(d => d && d.p === 'OwnerId');
  let action = 'Updated';
  if (ownerEntry) {
    action = (ownerEntry.c === null || ownerEntry.c === '' || ownerEntry.c === undefined)
      ? 'Unassigned'
      : 'Assigned';
  }

  const locationEntry = details.find(d => d && d.p === 'LocationId');
  const historicalLocation = locationEntry && locationEntry.c ? locationMap[locationEntry.c] : '';

  const assetInfo = assetMap[tag] || {};
  const currentlyWith = assetInfo.currentOwner || '(Unassigned)';

  return [
    item.ActivityDate || '',
    action,
    tag,
    assetInfo.serial || '',
    assetInfo.model || '',
    historicalLocation || assetInfo.location || '',
    currentlyWith
  ];
}

function installIndividualLookupTrigger(ss) {
  const existing = ScriptApp.getProjectTriggers()
    .some(t => t.getHandlerFunction() === INDIVIDUAL_LOOKUP_HANDLER);
  if (existing) return;
  ScriptApp.newTrigger(INDIVIDUAL_LOOKUP_HANDLER)
    .forSpreadsheet(ss)
    .onEdit()
    .create();
  try { logOperation('IndividualLookup', 'SETUP', 'Installed onEdit trigger'); } catch (_) {}
}

function removeIndividualLookupTrigger() {
  let removed = 0;
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === INDIVIDUAL_LOOKUP_HANDLER) {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  return removed;
}

// =============================================================================
// CATEGORY REGENERATION FUNCTIONS
// Per-category regeneration: defaults always rebuild, optionals only if installed.
// =============================================================================

function regenerateFleetOperations(ss) {
  // Defaults (always)
  setupAssignmentOverviewSheet(ss);
  setupStatusOverviewSheet(ss);
  // Optional (only if installed)
  let count = 2;
  [['DeviceReadiness', setupDeviceReadinessSheet],
   ['SpareAssets', setupSpareAssetsSheet],
   ['LostStolenRate', setupLostStolenRateSheet],
   ['ModelFragmentation', setupModelFragmentationSheet],
   ['UnassignedInventory', setupUnassignedInventorySheet],
  ].forEach(([name, fn]) => { if (ss.getSheetByName(name)) { fn(ss); count++; } });
  return count;
}

function regenerateServiceReliability(ss) {
  // Defaults (always)
  setupServiceImpactSheet(ss);
  // Optional (only if installed)
  let count = 1;
  [['BreakRate', setupBreakRateSheet],
   ['HighTicketLocations', setupHighTicketLocationsSheet],
  ].forEach(([name, fn]) => { if (ss.getSheetByName(name)) { fn(ss); count++; } });
  return count;
}

function regenerateBudgetPlanning(ss) {
  // Defaults (always)
  setupBudgetPlanningSheet(ss);
  setupAgingAnalysisSheet(ss);
  // Optional (only if installed)
  let count = 2;
  [['ReplacementPlanning', setupReplacementPlanningSheet],
   ['ReplacementForecast', setupReplacementForecastSheet],
   ['WarrantyTimeline', setupWarrantyTimelineSheet],
   ['DeviceLifecycle', setupDeviceLifecycleSheet],
  ].forEach(([name, fn]) => { if (ss.getSheetByName(name)) { fn(ss); count++; } });
  return count;
}

function regenerateFleetComposition(ss) {
  // Defaults (always)
  setupFleetSummarySheet(ss);
  setupLocationSummarySheet(ss);
  setupModelBreakdownSheet(ss);
  // Optional (only if installed)
  let count = 3;
  [['LocationModelBreakdown', setupLocationModelBreakdownSheet],
   ['LocationModelFiltered', setupLocationModelFilteredSheet],
   ['CategoryBreakdown', setupCategoryBreakdownSheet],
   ['ManufacturerSummary', setupManufacturerSummarySheet],
  ].forEach(([name, fn]) => { if (ss.getSheetByName(name)) { fn(ss); count++; } });
  return count;
}

function regeneratePeople(ss) {
  // Optional only — no People defaults yet.
  let count = 0;
  [['IndividualLookup', setupIndividualLookupSheet],
  ].forEach(([name, fn]) => { if (ss.getSheetByName(name)) { fn(ss); count++; } });
  return count;
}

function regenerateAllDefault(ss) {
  // All 8 default analytics sheets
  setupFleetSummarySheet(ss);
  setupLocationSummarySheet(ss);
  setupModelBreakdownSheet(ss);
  setupAgingAnalysisSheet(ss);
  setupBudgetPlanningSheet(ss);
  setupServiceImpactSheet(ss);
  setupAssignmentOverviewSheet(ss);
  setupStatusOverviewSheet(ss);
  return 8;
}

function regenerateAllAnalytics(ss) {
  const defaultCount = regenerateAllDefault(ss);

  const optional = [
    ['DeviceReadiness', setupDeviceReadinessSheet],
    ['SpareAssets', setupSpareAssetsSheet],
    ['LostStolenRate', setupLostStolenRateSheet],
    ['ModelFragmentation', setupModelFragmentationSheet],
    ['UnassignedInventory', setupUnassignedInventorySheet],
    ['BreakRate', setupBreakRateSheet],
    ['HighTicketLocations', setupHighTicketLocationsSheet],
    ['ReplacementPlanning', setupReplacementPlanningSheet],
    ['ReplacementForecast', setupReplacementForecastSheet],
    ['WarrantyTimeline', setupWarrantyTimelineSheet],
    ['DeviceLifecycle', setupDeviceLifecycleSheet],
    ['LocationModelBreakdown', setupLocationModelBreakdownSheet],
    ['LocationModelFiltered', setupLocationModelFilteredSheet],
    ['CategoryBreakdown', setupCategoryBreakdownSheet],
    ['ManufacturerSummary', setupManufacturerSummarySheet],
    ['IndividualLookup', setupIndividualLookupSheet],
  ];

  let optCount = 0;
  optional.forEach(([name, fn]) => {
    if (ss.getSheetByName(name)) { fn(ss); optCount++; }
  });
  return defaultCount + optCount;
}
