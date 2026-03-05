/**
 * iiQ Asset Reporting - Reference Data Loaders
 * Loads locations, asset status types, and discovers custom fields.
 */

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
// CUSTOM FIELD DISCOVERY
// =============================================================================

/**
 * Discover available custom fields for assets and look for AUE-related fields.
 * Writes findings to the CustomFields sheet and auto-configures AUE if found.
 */
function discoverCustomFields() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('CustomFields');
  if (!sheet) throw new Error('CustomFields sheet not found. Run Setup first.');

  logOperation('CustomFields', 'START', 'Discovering asset custom fields');

  const fields = getAssetCustomFields();
  if (!fields || fields.length === 0) {
    logOperation('CustomFields', 'COMPLETE', 'No custom fields found for assets');
    return;
  }

  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).clearContent();
  }

  const editorTypes = { 0: 'Text', 1: 'Dropdown', 2: 'Date', 3: 'Number', 4: 'Checkbox', 5: 'Lookup', 6: 'RichText' };

  const rows = fields.map(f => [
    f.CustomFieldId || f.CustomFieldTypeId || '',
    f.Name || '',
    editorTypes[f.EditorType] || String(f.EditorType || ''),
    f.IsRequired ? 'Yes' : 'No',
    f.Description || ''
  ]);

  sheet.getRange(2, 1, rows.length, 5).setValues(rows);

  // Auto-detect AUE field (look for common names)
  const aueKeywords = ['aue', 'auto update', 'autoupdate', 'end of life', 'eol', 'chrome.*expir'];
  let aueField = null;
  for (const f of fields) {
    const name = (f.Name || '').toLowerCase();
    if (aueKeywords.some(kw => name.match(new RegExp(kw, 'i')))) {
      aueField = f;
      break;
    }
  }

  if (aueField) {
    const fieldId = aueField.CustomFieldId || aueField.CustomFieldTypeId || '';
    setConfigValue('AUE_CUSTOM_FIELD_ID', fieldId);
    setConfigValue('AUE_CUSTOM_FIELD_NAME', aueField.Name);
    logOperation('CustomFields', 'COMPLETE',
      `${rows.length} fields found. AUE field auto-detected: "${aueField.Name}" (${fieldId})`);
  } else {
    logOperation('CustomFields', 'COMPLETE',
      `${rows.length} fields found. No AUE field auto-detected. Set AUE_CUSTOM_FIELD_ID manually if needed.`);
  }
}
