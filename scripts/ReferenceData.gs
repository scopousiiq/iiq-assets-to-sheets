/**
 * iiQ Asset Reporting - Reference Data Loaders
 * Loads locations and asset status types.
 */

// =============================================================================
// AUTO-LOAD (called before first asset page)
// =============================================================================

/**
 * Load reference data if sheets are empty. Called automatically before
 * the first page of asset loading so users don't need a separate step.
 */
function ensureReferenceData(ss) {
  const locSheet = ss.getSheetByName('Locations');
  if (locSheet && locSheet.getLastRow() <= 1) {
    logOperation('ReferenceData', 'AUTO', 'Loading locations (first run)');
    loadLocations();
  }

  const statusSheet = ss.getSheetByName('StatusTypes');
  if (statusSheet && statusSheet.getLastRow() <= 1) {
    logOperation('ReferenceData', 'AUTO', 'Loading status types (first run)');
    loadStatusTypes();
  }

  // LocationEnrollment is NOT auto-loaded here — requires STUDENT_ROLE_ID
  // to be configured first. User runs it manually via menu after setup.
}

// =============================================================================
// LOCATIONS
// =============================================================================

/**
 * Load all locations into the Locations sheet.
 * Replaces all data each time (small dataset, no pagination needed).
 */
function loadLocations() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Locations');
  if (!sheet) throw new Error('Locations sheet not found. Run Setup first.');

  logOperation('Locations', 'START', 'Loading all locations');

  const locations = getAllLocations();
  if (!locations || locations.length === 0) {
    logOperation('Locations', 'COMPLETE', 'No locations returned');
    return;
  }

  // Clear existing data (keep header)
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).clearContent();
  }

  const rows = locations.map(loc => {
    const locType = loc.LocationType || {};
    const address = loc.Address || {};
    const addrParts = [address.Street1, address.City, address.State, address.Zip].filter(Boolean);

    return [
      loc.LocationId || '',
      loc.Name || '',
      loc.Abbreviation || '',
      locType.Name || loc.LocationTypeName || '',
      addrParts.join(', ')
    ];
  });

  sheet.getRange(2, 1, rows.length, 5).setValues(rows);
  logOperation('Locations', 'COMPLETE', `${rows.length} locations loaded`);
}

// =============================================================================
// ASSET STATUS TYPES
// =============================================================================

/**
 * Load all asset status types into the StatusTypes sheet.
 */
function loadStatusTypes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('StatusTypes');
  if (!sheet) throw new Error('StatusTypes sheet not found. Run Setup first.');

  logOperation('StatusTypes', 'START', 'Loading asset status types');

  const types = getAllAssetStatusTypes();
  if (!types || types.length === 0) {
    logOperation('StatusTypes', 'COMPLETE', 'No status types returned');
    return;
  }

  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).clearContent();
  }

  const rows = types
    .filter(t => t.IsActive !== false && t.IsDeleted !== true)
    .map(t => [
      t.AssetStatusTypeId || '',
      t.Name || '',
      t.IsRetired ? 'Yes' : 'No',
      t.SortOrder ?? ''
    ]);

  sheet.getRange(2, 1, rows.length, 4).setValues(rows);
  logOperation('StatusTypes', 'COMPLETE', `${rows.length} status types loaded`);
}

// =============================================================================
// LOCATION ENROLLMENT
// =============================================================================

/**
 * Load student enrollment data per location into LocationEnrollment sheet.
 * Uses 2N API calls (where N = number of locations):
 *   - N calls: total students per location (role + location filter)
 *   - N calls: students with assigned devices per location (+ hasassigneddevice facet)
 * Checkpoint resume: appends each row as counted, resumes from last row on next run.
 * Requires STUDENT_ROLE_ID in Config (use "View Available Roles" to find it).
 * Auto-creates the sheet if it doesn't exist.
 */
function loadLocationEnrollment() {
  const MAX_RUNTIME_MS = 5 * 60 * 1000;
  const startTime = Date.now();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = getConfig();

  // Require configured student role
  if (!config.studentRoleId) {
    logOperation('Enrollment', 'ERROR', 'STUDENT_ROLE_ID not configured. Use iiQ Assets > Load Reference Data > View Available Roles to find the correct ID.');
    throw new Error('STUDENT_ROLE_ID not configured in Config sheet.\n\nUse: iiQ Assets > Load Reference Data > View Available Roles\nto find the correct role ID, then paste it into the Config sheet.');
  }

  // Auto-create sheet if missing (safe for existing spreadsheets)
  let sheet = ss.getSheetByName('LocationEnrollment');
  if (!sheet) {
    setupLocationEnrollmentSheet(ss);
    sheet = ss.getSheetByName('LocationEnrollment');
  }

  // Read locations from the Locations sheet
  const locSheet = ss.getSheetByName('Locations');
  if (!locSheet || locSheet.getLastRow() <= 1) {
    logOperation('Enrollment', 'ERROR', 'Locations sheet empty — load locations first');
    throw new Error('Locations sheet is empty. Load reference data first.');
  }

  const locData = locSheet.getDataRange().getValues().slice(1); // skip header
  const totalLocations = locData.length;

  // Resume from checkpoint: count existing data rows
  const startIndex = Math.max(sheet.getLastRow() - 1, 0); // subtract header

  if (startIndex >= totalLocations) {
    logOperation('Enrollment', 'SKIP', 'All locations already counted');
    return 'already_complete';
  }

  logOperation('Enrollment', 'START',
    `Counting students at locations ${startIndex + 1}–${totalLocations} (role: ${config.studentRoleId})`);

  // 2N calls per location: total students + students with devices
  let counted = 0;
  for (let i = startIndex; i < totalLocations; i++) {
    if (Date.now() - startTime > MAX_RUNTIME_MS) {
      logOperation('Enrollment', 'PAUSED',
        `${startIndex + counted}/${totalLocations} locations done. Run again to continue.`);
      return 'paused';
    }

    const locId = locData[i][0];
    const locName = locData[i][1];
    const locType = locData[i][3]; // col D = LocationType
    if (!locId) continue;

    // Call 1: Total students at this location
    const totalStudents = getUserCount([
      { Facet: 'role', Id: config.studentRoleId },
      { Facet: 'location', Id: locId }
    ]);

    // Call 2: Students with assigned devices at this location
    const studentsWithDevices = getUserCount([
      { Facet: 'role', Id: config.studentRoleId },
      { Facet: 'location', Id: locId },
      { Facet: 'hasassigneddevice', Selected: true }
    ]);

    const coverage = totalStudents > 0 ? studentsWithDevices / totalStudents : 0;

    // Append row directly — each row is a checkpoint
    const writeRow = sheet.getLastRow() + 1;
    sheet.getRange(writeRow, 1, 1, 6).setValues([[locId, locName, locType, totalStudents, studentsWithDevices, coverage]]);
    counted++;

    if (counted % 10 === 0) {
      logOperation('Enrollment', 'PROGRESS', `${startIndex + counted}/${totalLocations} locations counted`);
    }
  }

  logOperation('Enrollment', 'COMPLETE',
    `All ${totalLocations} locations with student enrollment loaded`);
  return 'complete';
}
