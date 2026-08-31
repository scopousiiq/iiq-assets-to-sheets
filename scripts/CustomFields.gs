/**
 * iiQ Asset Reporting - Custom Fields
 *
 * Pulls up to ASSET_CUSTOM_FIELD_COUNT district-defined asset custom fields into
 * AssetData columns AM-AQ.
 *
 * Values come out of the CustomFieldValues array already present on every item
 * in the /v1.0/assets search response, so no per-asset request is added. The only
 * extra call is a single POST /v1.0/custom-fields/for/asset per run, and only
 * when a slot was given as a display name or a value needs an id-to-name map.
 *
 * The custom field columns sit AFTER the ARRAYFORMULA columns (AJ-AL) rather than
 * before them. Appending leaves every analytics formula and every dashboard
 * column offset valid, so an existing sheet upgrades by gaining five headers
 * instead of being rebuilt and reloaded.
 */

/**
 * iiQ EditorTypes enum (Spark.Shared/Enums.cs) — only the values a district is
 * likely to put on an asset custom field are named.
 */
const EDITOR_TYPE_LABELS = {
  0: 'None', 1: 'Text', 2: 'MultilineText', 3: 'RichText', 4: 'Number',
  5: 'NumberRange', 6: 'Date', 7: 'DateRange', 8: 'OnOff', 9: 'Select',
  10: 'MultiSelect', 11: 'Email', 12: 'Phone', 13: 'Address', 14: 'FileUpload',
  18: 'IPAddress', 21: 'IiqUser', 22: 'IiqLocation', 23: 'IiqAsset',
  29: 'IiqModel', 33: 'IiqTeam', 35: 'IiqRoom'
};

// =============================================================================
// SLOT RESOLUTION
// =============================================================================

/**
 * Resolve the configured custom field slots to CustomFieldTypeId UUIDs.
 *
 * A slot holds either a CustomFieldTypeId pasted from the CustomFields sheet —
 * the documented input, used as-is with no API call — or a field display name,
 * which costs one definitions call to look up. Names are accepted for
 * convenience but are not unique within a district: two field types can share a
 * display name, and a name lookup then picks one of them arbitrarily.
 *
 * Resolved ids are echoed into the CUSTOM_FIELD_n_ID Config rows for diagnostics
 * only, never read back as a cache. Caching them would mean a district that
 * edits a slot keeps pulling the previously resolved field forever.
 *
 * @param {Object} config - Config object from getConfig()
 * @returns {string[]} - ASSET_CUSTOM_FIELD_COUNT ids; '' for unset or unresolved
 */
function resolveAssetCustomFieldIds(config) {
  const slots = config.customFields || [];
  const ids = new Array(ASSET_CUSTOM_FIELD_COUNT).fill('');
  const needName = [];

  for (let i = 0; i < ASSET_CUSTOM_FIELD_COUNT; i++) {
    const slot = (slots[i] || '').trim();
    if (!slot) continue;
    if (looksLikeCustomFieldGuid_(slot)) {
      ids[i] = slot;
    } else {
      needName.push(i);
    }
  }

  if (needName.length > 0) {
    const nameToId = buildCustomFieldNameIndex_();
    if (nameToId) {
      needName.forEach(i => {
        const slot = slots[i].trim();
        const found = nameToId[slot.toLowerCase()];
        if (found) {
          ids[i] = found;
          logOperation('CustomFields', 'RESOLVED', `CUSTOM_FIELD_${i + 1} "${slot}" → ${found}`);
        } else {
          logOperation('CustomFields', 'WARNING',
            `CUSTOM_FIELD_${i + 1} "${slot}" is not an asset custom field in this district`);
        }
      });
    }
  }

  writeResolvedCustomFieldIds_(slots, ids);
  return ids;
}

/**
 * Lowercased display name → CustomFieldTypeId, or null if the definitions call
 * failed. First definition wins for a duplicated name; the ambiguity is why the
 * id is the documented input.
 */
function buildCustomFieldNameIndex_() {
  let definitions;
  try {
    definitions = getAssetCustomFieldDefinitions();
  } catch (e) {
    logOperation('CustomFields', 'ERROR', 'Could not fetch asset custom field definitions: ' + e.message);
    return null;
  }

  const nameToId = {};
  definitions.forEach(def => {
    const name = def.CustomFieldType && def.CustomFieldType.Name;
    if (!name || !def.CustomFieldTypeId) return;
    const key = String(name).trim().toLowerCase();
    if (!(key in nameToId)) nameToId[key] = def.CustomFieldTypeId;
  });
  return nameToId;
}

/**
 * Echo resolution results into the CUSTOM_FIELD_n_ID Config rows so a district
 * can see at a glance whether each slot landed on a real field. A configured
 * slot that did not resolve reads NOT_FOUND.
 */
function writeResolvedCustomFieldIds_(slots, ids) {
  for (let i = 0; i < ASSET_CUSTOM_FIELD_COUNT; i++) {
    const slot = (slots[i] || '').trim();
    setConfigValue('CUSTOM_FIELD_' + (i + 1) + '_ID', slot ? (ids[i] || 'NOT_FOUND') : '');
  }
}

/**
 * Does this slot value look like a CustomFieldTypeId rather than a field name?
 */
function looksLikeCustomFieldGuid_(value) {
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(String(value || '').trim());
}

/**
 * Per-slot status lines for the Verify Configuration dialog. Empty for a
 * district with no slots configured, so that dialog is unchanged for them.
 *
 * Resolution alone is not enough of a check here: a slot holding a well-formed
 * but wrong GUID resolves to itself and would report OK. Verify Configuration is
 * exactly where a mistyped id should surface, so each id is checked against the
 * district's actual field list.
 *
 * @param {Object} config - Config object from getConfig()
 * @returns {string[]} - Lines to append to the dialog
 */
function describeCustomFieldSlots(config) {
  const slots = config.customFields || [];
  const configured = [];
  for (let i = 0; i < ASSET_CUSTOM_FIELD_COUNT; i++) {
    const value = (slots[i] || '').trim();
    if (value) configured.push({ index: i, value: value });
  }
  if (configured.length === 0) return [];

  const ids = resolveAssetCustomFieldIds(config);

  let namesById = {};
  let haveNames = true;
  try {
    getAssetCustomFieldDefinitions().forEach(def => {
      if (!def.CustomFieldTypeId) return;
      namesById[def.CustomFieldTypeId] = (def.CustomFieldType && def.CustomFieldType.Name) || def.CustomFieldTypeId;
    });
  } catch (e) {
    haveNames = false; // report resolution only rather than failing the whole check
  }

  const lines = ['', `Custom fields (${configured.length} configured):`];
  configured.forEach(entry => {
    const id = ids[entry.index];
    let status;
    if (!id) {
      status = `"${entry.value}" not found in this district`;
    } else if (!haveNames) {
      status = 'resolved';
    } else if (namesById[id]) {
      status = 'OK — ' + namesById[id];
    } else {
      status = `id ${id} is not an asset custom field in this district`;
    }
    lines.push(`  • CUSTOM_FIELD_${entry.index + 1}: ${status}`);
  });
  return lines;
}

// =============================================================================
// PER-RUN CONTEXT
// =============================================================================

/**
 * Build everything a loading run needs to fill the custom field columns, once
 * per run rather than once per asset.
 *
 * @param {Object} config - Config object from getConfig()
 * @returns {Object|null} - { ids, lookupMaps } or null if no slot is configured
 */
function buildAssetCustomFieldContext(config) {
  const slots = config.customFields || [];
  if (!slots.some(s => s && String(s).trim())) return null;

  const ids = resolveAssetCustomFieldIds(config);
  if (!ids.some(id => id)) return null;

  return { ids: ids, lookupMaps: buildAssetCustomFieldLookupMaps_(ids) };
}

/**
 * One raw-value → display-name map per configured slot.
 *
 * Only two editor types store ids that need translating: Select/MultiSelect
 * (9/10), whose options are already on the definition as an Options JSON blob,
 * and IiqLocation (22), which needs the location directory. Everything else
 * stores a value that is already readable, so it gets an empty map.
 *
 * @param {string[]} ids - Resolved CustomFieldTypeIds ('' entries yield {})
 * @returns {Object[]} - Lookup objects, parallel to ids
 */
function buildAssetCustomFieldLookupMaps_(ids) {
  const blank = ids.map(() => ({}));
  if (!ids.some(id => id)) return blank;

  let definitions;
  try {
    definitions = getAssetCustomFieldDefinitions();
  } catch (e) {
    logOperation('CustomFields', 'WARNING', 'Could not fetch definitions for lookup maps: ' + e.message);
    return blank;
  }

  const defById = {};
  definitions.forEach(def => {
    if (def.CustomFieldTypeId) defById[def.CustomFieldTypeId] = def;
  });

  let locationMap = null; // lazy — at most one locations sweep per run

  return ids.map(id => {
    if (!id) return {};
    const def = defById[id];
    if (!def) return {};
    const fieldType = def.CustomFieldType || {};
    const editorType = def.EditorTypeId ?? fieldType.EditorType;

    if (editorType === 22) { // IiqLocation
      if (!locationMap) {
        locationMap = {};
        try {
          getAllLocations().forEach(loc => {
            if (loc.LocationId && loc.Name) locationMap[loc.LocationId] = loc.Name;
          });
          logOperation('CustomFields', 'INFO',
            `Built location lookup map with ${Object.keys(locationMap).length} entries`);
        } catch (e) {
          logOperation('CustomFields', 'WARNING', 'Could not fetch locations for lookup: ' + e.message);
        }
      }
      return locationMap;
    }

    if (editorType === 9 || editorType === 10) { // Select / MultiSelect
      return parseSelectOptions_(fieldType.Options || def.Options || '', id);
    }

    return {};
  });
}

/** Option id → option name from a definition's Options JSON. */
function parseSelectOptions_(optionsJson, id) {
  if (!optionsJson) return {};
  let options;
  try {
    options = JSON.parse(optionsJson);
  } catch (e) {
    return {};
  }
  if (!Array.isArray(options)) return {};

  const map = {};
  options.forEach(opt => {
    const optId = opt.Id || opt.CustomFieldOptionId || opt.OptionId;
    const name = opt.Name || opt.Label || opt.DisplayValue || opt.Text || opt.Value;
    if (optId && name) map[String(optId)] = String(name);
  });
  logOperation('CustomFields', 'INFO', `Built select lookup map for ${id} with ${Object.keys(map).length} options`);
  return map;
}

// =============================================================================
// VALUE EXTRACTION
// =============================================================================

/**
 * The custom field block for one asset — ASSET_CUSTOM_FIELD_COUNT values in slot
 * order, for AssetData columns AM-AQ.
 *
 * @param {Object} asset - Asset object from the search response
 * @param {Object} context - From buildAssetCustomFieldContext()
 * @returns {Array} - One value per slot
 */
function extractAssetCustomFieldRow(asset, context) {
  return context.ids.map((id, i) =>
    extractCustomFieldValue(asset, id, context.lookupMaps[i])
  );
}

/**
 * Read one custom field off an asset and render it as a cell value.
 *
 * Stored values range from plain scalars to JSON arrays of entity references,
 * so this unwraps whichever shape came back and joins multi-value fields with a
 * comma. lookupMap resolves raw ids to display names where the editor type
 * stores ids.
 *
 * @param {Object} asset - Asset object (may or may not have CustomFieldValues)
 * @param {string} customFieldTypeId - Resolved id of the field to read
 * @param {Object} lookupMap - Raw value → display name, may be empty
 * @returns {string} - Cell value, '' when the field is unset on this asset
 */
function extractCustomFieldValue(asset, customFieldTypeId, lookupMap) {
  if (!customFieldTypeId) return '';
  const values = asset.CustomFieldValues;
  if (!values || !values.length) return '';

  const match = values.find(cf => cf.CustomFieldTypeId === customFieldTypeId);
  if (!match || match.Value == null) return '';

  const raw = String(match.Value);
  const trimmed = raw.trim();
  if (!trimmed) return '';

  if (lookupMap && lookupMap[trimmed]) return lookupMap[trimmed];

  const firstChar = trimmed.charAt(0);
  if (firstChar !== '[' && firstChar !== '{') return raw;

  let parsed;
  try {
    parsed = JSON.parse(trimmed);
  } catch (e) {
    return raw;
  }

  if (Array.isArray(parsed)) {
    return parsed.map(e => formatCustomFieldEntry_(e, lookupMap)).filter(s => s !== '').join(', ');
  }
  return formatCustomFieldEntry_(parsed, lookupMap);
}

/**
 * Render one entry from a parsed custom field value. Entries are either
 * primitives (ids, numbers, booleans) or objects carrying a display field
 * alongside id fields — prefer the display field, fall back to a mapped id,
 * then to the raw id.
 */
function formatCustomFieldEntry_(entry, lookupMap) {
  if (entry == null) return '';

  if (typeof entry !== 'object') {
    const s = String(entry);
    return (lookupMap && lookupMap[s]) || s;
  }

  const displayKeys = ['Name', 'DisplayName', 'DisplayValue', 'Text', 'Label', 'Title', 'Caption', 'Value'];
  for (const key of displayKeys) {
    const v = entry[key];
    if (v != null && typeof v !== 'object' && String(v) !== '') return String(v);
  }

  const idKeys = ['AssetId', 'UserId', 'LocationId', 'LocationRoomId', 'ModelId', 'TicketId', 'CustomFieldOptionId', 'Id'];
  for (const key of idKeys) {
    const v = entry[key];
    if (v != null && String(v) !== '') {
      const s = String(v);
      return (lookupMap && lookupMap[s]) || s;
    }
  }
  return '';
}

// =============================================================================
// ASSETDATA COLUMN MIGRATION
// =============================================================================

/**
 * Make sure AssetData is wide enough for the custom field block and carries its
 * headers. Sheets created before custom field support stop at 38 columns, and
 * writing a fixed-width range past the grid edge throws rather than growing it.
 *
 * @param {Sheet} sheet - AssetData sheet
 */
function ensureAssetCustomFieldColumns(sheet) {
  if (!sheet) return;

  ensureAssetGridWidth_(sheet);

  const headerRange = sheet.getRange(1, ASSET_CUSTOM_FIELD_START_COL, 1, ASSET_CUSTOM_FIELD_COUNT);
  if (String(headerRange.getValues()[0][0]) === ASSET_CUSTOM_FIELD_HEADERS[0]) return;

  headerRange.setValues([ASSET_CUSTOM_FIELD_HEADERS]).setFontWeight('bold');
  logOperation('CustomFields', 'MIGRATED', 'Added CustomField1-' + ASSET_CUSTOM_FIELD_COUNT + ' columns to AssetData');
}

// =============================================================================
// CONFIG MIGRATION
// =============================================================================

/**
 * Non-destructive migration: add the custom field rows to a Config sheet created
 * before custom field support. Safe to call repeatedly — returns early once
 * CUSTOM_FIELD_1 exists.
 */
function migrateConfigForAssetCustomFields() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  for (const row of data) {
    if (row[0] === 'CUSTOM_FIELD_1') return; // Already migrated
  }

  const rows = buildCustomFieldConfigRows_();
  const start = sheet.getLastRow() + 1;
  sheet.getRange(start, 1, rows.length, 2).setValues(rows);
  annotateCustomFieldConfigCells_(sheet);
  resetConfigCache();

  logOperation('Config', 'MIGRATED', 'Added custom field configuration rows');
}

/**
 * The Config rows that back the custom field columns. Shared by first-time setup
 * and the migration so the two never drift apart.
 */
function buildCustomFieldConfigRows_() {
  const rows = [
    ['', ''],
    ['--- Asset Custom Fields (optional) ---', '']
  ];
  for (let n = 1; n <= ASSET_CUSTOM_FIELD_COUNT; n++) rows.push(['CUSTOM_FIELD_' + n, '']);
  rows.push(['', '']);
  rows.push(['--- Custom Field Resolution (auto-managed) ---', '']);
  for (let n = 1; n <= ASSET_CUSTOM_FIELD_COUNT; n++) rows.push(['CUSTOM_FIELD_' + n + '_ID', '']);
  return rows;
}

/**
 * Note the CUSTOM_FIELD_n cells with what to paste there. These take an id from
 * the CustomFields sheet, not a selection from a list — field names are not
 * unique within a district, and Sheets caps list validation well below the
 * number of fields a large district defines.
 */
function annotateCustomFieldConfigCells_(sheet) {
  if (!sheet) return;
  const note = 'Paste a CustomFieldTypeId from the CustomFields sheet (column B). ' +
    'A field name also works, but ids are unambiguous — a district can define two fields sharing a name.';

  const keys = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues();
  keys.forEach((row, i) => {
    if (/^CUSTOM_FIELD_\d+$/.test(String(row[0]).trim())) {
      sheet.getRange(i + 1, 2).setNote(note);
    }
  });
}

// =============================================================================
// CUSTOMFIELDS REFERENCE SHEET
// =============================================================================

/**
 * Create the CustomFields sheet. Districts copy a CustomFieldTypeId from here
 * into a CUSTOM_FIELD_n row on the Config sheet.
 */
function setupCustomFieldsSheet(ss) {
  const headers = [['Name', 'CustomFieldTypeId', 'EditorType']];
  const { sheet, isNew } = getOrCreateSheet(ss, 'CustomFields', headers, '#7b1fa2', {
    columnWidths: { 1: 260, 2: 300, 3: 130 }
  });

  if (isNew) {
    sheet.getRange('A1').setNote(
      'Asset Custom Fields Reference\n\n' +
      'Every custom field defined on assets in your district.\n\n' +
      'To pull one into AssetData:\n' +
      '1. Run iiQ Assets > Setup > Refresh Custom Fields to populate this sheet\n' +
      '2. Copy the CustomFieldTypeId (column B) of the field you want\n' +
      '3. Paste it into CUSTOM_FIELD_1-' + ASSET_CUSTOM_FIELD_COUNT + ' on the Config sheet\n' +
      '4. Run Full Reload to backfill the column for every asset\n\n' +
      'Values land in AssetData columns AM-AQ, in slot order.'
    );
  }

  return sheet;
}

/**
 * Populate the CustomFields sheet from the API. Also ensures the Config rows
 * exist, so this is the one action a district needs before configuring a slot.
 */
function refreshAssetCustomFields() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  migrateConfigForAssetCustomFields();

  // Fetching is separated from sheet-building so a formatting failure is not
  // reported as an API credentials problem.
  let definitions;
  try {
    definitions = getAssetCustomFieldDefinitions();
  } catch (e) {
    ui.alert('Error Fetching Custom Fields',
      'Could not reach the iiQ API: ' + e.message + '\n\n' +
      'Check that API_BASE_URL and BEARER_TOKEN are configured correctly on the Config sheet.',
      ui.ButtonSet.OK);
    return;
  }

  if (!definitions || definitions.length === 0) {
    ui.alert('No Custom Fields', 'No asset custom fields are defined in your district.', ui.ButtonSet.OK);
    return;
  }

  try {
    const sheet = setupCustomFieldsSheet(ss);
    const rows = buildAssetCustomFieldRows_(definitions);
    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, 3).setValues(rows);
    }
    annotateCustomFieldConfigCells_(ss.getSheetByName('Config'));

    logOperation('CustomFields', 'COMPLETE', `Listed ${rows.length} asset custom field(s)`);
    ui.alert('Custom Fields Refreshed',
      `Found ${rows.length} asset custom field(s).\n\n` +
      'To use one, copy its CustomFieldTypeId (column B) into CUSTOM_FIELD_1-' +
      ASSET_CUSTOM_FIELD_COUNT + ' on the Config sheet.\n\n' +
      'New slots populate on the next refresh for assets that change. Run Full Reload ' +
      'to backfill the column for the whole fleet.',
      ui.ButtonSet.OK);
  } catch (e) {
    logOperation('CustomFields', 'ERROR', 'Failed to build CustomFields sheet: ' + e.message);
    ui.alert('Error Building CustomFields Sheet',
      'The custom fields were fetched from iiQ successfully, but writing them to the sheet failed:\n\n' + e.message,
      ui.ButtonSet.OK);
  }
}

/**
 * Turn definitions into CustomFields sheet rows, sorted by name.
 *
 * Deduplicated by CustomFieldTypeId: /custom-fields/for/asset returns one item
 * per field-to-filter-set mapping rather than one per field, so a single field
 * can come back dozens of times with an identical name and id. Distinct field
 * types that happen to share a display name stay as separate rows — they are
 * genuinely different fields, and collapsing them would hide one.
 *
 * @param {Array} definitions - CustomFieldDetail objects
 * @returns {Array} - Rows of [name, customFieldTypeId, editorTypeLabel]
 */
function buildAssetCustomFieldRows_(definitions) {
  const rows = [];
  const seen = {};
  const nameCounts = {};
  let mappingCount = 0;

  (definitions || []).forEach(def => {
    const name = (def.CustomFieldType && def.CustomFieldType.Name) || '';
    if (!name) return; // Unnamed fields cannot be selected or displayed

    const id = def.CustomFieldTypeId || '';
    mappingCount++;
    if (id && seen[id]) return;
    if (id) seen[id] = true;

    const editorTypeId = def.EditorTypeId ?? (def.CustomFieldType && def.CustomFieldType.EditorType) ?? 0;
    const key = name.trim().toLowerCase();
    nameCounts[key] = (nameCounts[key] || 0) + 1;

    rows.push([name, id, EDITOR_TYPE_LABELS[editorTypeId] || 'Type ' + editorTypeId]);
  });

  if (mappingCount > rows.length) {
    logOperation('CustomFields', 'DEDUPED',
      `Collapsed ${mappingCount} field/filter-set mappings into ${rows.length} distinct field(s)`);
  }

  const ambiguous = Object.keys(nameCounts).filter(k => nameCounts[k] > 1).length;
  if (ambiguous > 0) {
    logOperation('CustomFields', 'WARNING',
      `${ambiguous} field name(s) are used by more than one field type — configure those by CustomFieldTypeId, not name`);
  }

  rows.sort((a, b) => String(a[0]).localeCompare(String(b[0])));
  return rows;
}
