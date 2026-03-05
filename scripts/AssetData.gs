/**
 * iiQ Asset Reporting - Asset Data Loader
 *
 * Two-phase loading:
 *   Phase 1: Bulk asset search (paginated, fast)
 *   Phase 2: Custom field enrichment (per-batch, for AUE date etc.)
 *
 * Column layout (25 columns, A-Y):
 *   A  AssetId            M  StatusName
 *   B  AssetTag           N  PurchasedDate
 *   C  Name               O  WarrantyExpDate
 *   D  SerialNumber       P  PurchasePrice
 *   E  ModelName          Q  CreatedDate
 *   F  ManufacturerName   R  ModifiedDate
 *   G  CategoryName       S  OpenTickets
 *   H  LocationId         T  AUEDate (custom field)
 *   I  LocationName       U  AUEStatus (formula)
 *   J  LocationType       V  AgeDays (formula)
 *   K  OwnerId            W  AgeYears (formula)
 *   L  OwnerName          X  WarrantyStatus (formula)
 *                         Y  ReplacementCycle (formula)
 */

const ASSET_HEADERS = [
  'AssetId', 'AssetTag', 'Name', 'SerialNumber',
  'ModelName', 'ManufacturerName', 'CategoryName',
  'LocationId', 'LocationName', 'LocationType',
  'OwnerId', 'OwnerName', 'StatusName',
  'PurchasedDate', 'WarrantyExpDate', 'PurchasePrice',
  'CreatedDate', 'ModifiedDate', 'OpenTickets',
  'AUEDate',
  'AUEStatus', 'AgeDays', 'AgeYears', 'WarrantyStatus', 'ReplacementCycle'
];
const ASSET_DATA_COLS = 20;  // Columns A-T (API data + AUE custom field)
const ASSET_TOTAL_COLS = ASSET_HEADERS.length; // 25 (includes formula columns)
const MAX_RUNTIME_MS = 5.5 * 60 * 1000;

// =============================================================================
// PHASE 1: BULK ASSET LOADING
// =============================================================================

/**
 * Load assets via paginated search. Resumes from last checkpoint.
 * @param {boolean} showUI - Whether to show UI alerts
 * @returns {string} - 'complete', 'paused', or 'already_complete'
 */
function loadAssetData(showUI) {
  const config = getConfig();
  if (config.assetComplete) return 'already_complete';

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('AssetData');
  if (!sheet) throw new Error('AssetData sheet not found. Run Setup first.');

  const startTime = Date.now();
  let currentPage = config.assetLastPage + 1;
  let totalPages = config.assetTotalPages;
  let totalRowsWritten = 0;

  logOperation('AssetLoad', 'START', `Resuming from page ${currentPage}`);
  cacheConfigRowPositions_();

  while (Date.now() - startTime < MAX_RUNTIME_MS) {
    const response = searchAssets([], currentPage, config.assetBatchSize);

    if (!response || !response.Items) {
      logOperation('AssetLoad', 'ERROR', `Empty response on page ${currentPage}`);
      break;
    }

    // Capture total pages on first response
    if (totalPages === -1 && response.Paging) {
      totalPages = Math.ceil(response.Paging.TotalRows / config.assetBatchSize) - 1;
      setConfigValue('ASSET_TOTAL_PAGES', String(totalPages));
    }

    // Extract and write rows
    const rows = response.Items.map(asset => extractAssetRow(asset));
    if (rows.length > 0) {
      const lastRow = Math.max(sheet.getLastRow(), 1);
      sheet.getRange(lastRow + 1, 1, rows.length, ASSET_DATA_COLS).setValues(rows);
      totalRowsWritten += rows.length;
    }

    // Checkpoint
    setConfigValue('ASSET_LAST_PAGE', String(currentPage));

    const displayPage = currentPage + 1;
    const displayTotal = totalPages + 1;
    logOperation('AssetLoad', 'BATCH', `Page ${displayPage}/${displayTotal} (${rows.length} assets, ${totalRowsWritten} total this run)`);

    // Check completion
    if (currentPage >= totalPages) {
      setConfigValue('ASSET_COMPLETE', 'TRUE');
      logOperation('AssetLoad', 'COMPLETE', `All ${displayTotal} pages loaded (${totalRowsWritten} assets this run)`);
      return 'complete';
    }

    currentPage++;
    Utilities.sleep(config.throttleMs);
  }

  // Timed out
  const displayPage = currentPage + 1;
  const displayTotal = totalPages >= 0 ? totalPages + 1 : '?';
  logOperation('AssetLoad', 'PAUSED', `Page ${displayPage}/${displayTotal} (${totalRowsWritten} assets this run). Will resume.`);
  return 'paused';
}

/**
 * Extract one row of asset data from an API response item.
 * Returns array of ASSET_DATA_COLS values (columns A-T).
 */
function extractAssetRow(asset) {
  const model = asset.Model || {};
  const location = asset.Location || {};
  const owner = asset.Owner || {};
  const status = asset.AssetStatus || asset.Status || {};

  return [
    asset.AssetId || '',
    asset.AssetTag || '',
    asset.Name || '',
    asset.SerialNumber || '',
    model.ModelName || model.Name || asset.ModelName || '',
    model.ManufacturerName || (model.Manufacturer ? model.Manufacturer.Name : '') || '',
    model.CategoryNameWithParent || model.CategoryName || '',
    location.LocationId || asset.LocationId || '',
    location.Name || asset.LocationName || '',
    location.LocationTypeName || '',
    owner.UserId || asset.OwnerId || '',
    owner.Name || asset.OwnerName || '',
    status.Name || asset.StatusName || '',
    formatDate(asset.PurchasedDate),
    formatDate(asset.WarrantyExpirationDate),
    asset.PurchasePrice ?? '',
    formatDate(asset.CreatedDate),
    formatDate(asset.ModifiedDate),
    asset.OpenTicketCount ?? asset.OpenTickets ?? '',
    '', // AUEDate - populated in Phase 2
  ];
}

function formatDate(val) {
  if (!val) return '';
  const d = new Date(val);
  if (isNaN(d.getTime())) return '';
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

// =============================================================================
// PHASE 2: CUSTOM FIELD ENRICHMENT (AUE Date)
// =============================================================================

/**
 * Enrich assets with custom field values (AUE date).
 * Processes in batches, resumes from last checkpoint.
 * @param {boolean} showUI
 * @returns {string} - 'complete', 'paused', 'skipped', or 'no_aue_field'
 */
function enrichAssetData(showUI) {
  const config = getConfig();

  if (!config.assetComplete) return 'skipped'; // Phase 1 not done
  if (config.enrichComplete) return 'already_complete';
  if (!config.aueFieldId) return 'no_aue_field';

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('AssetData');
  if (!sheet) return 'skipped';

  const startTime = Date.now();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 'complete';

  // Read all asset IDs from column A
  const assetIds = sheet.getRange(2, 1, lastRow - 1, 1).getValues().map(r => r[0]).filter(id => id);
  const startIdx = config.enrichLastIdx + 1;
  const BATCH_SIZE = 50;
  const AUE_COL = 20; // Column T

  logOperation('Enrich', 'START', `Processing from index ${startIdx}, ${assetIds.length} total assets`);
  cacheConfigRowPositions_();

  let idx = startIdx;
  while (idx < assetIds.length && Date.now() - startTime < MAX_RUNTIME_MS) {
    const batchEnd = Math.min(idx + BATCH_SIZE, assetIds.length);
    const batchIds = assetIds.slice(idx, batchEnd);

    try {
      const cfValues = getCustomFieldValuesForAssets(batchIds);
      const aueMap = buildAueMap(cfValues, config.aueFieldId);

      // Write AUE dates for this batch
      const aueValues = batchIds.map(id => [aueMap[id] || '']);
      const sheetRow = idx + 2; // +1 for header, +1 for 1-indexing
      sheet.getRange(sheetRow, AUE_COL, aueValues.length, 1).setValues(aueValues);
    } catch (e) {
      logOperation('Enrich', 'ERROR', `Batch at index ${idx}: ${e.message}`);
      // Continue — don't block on enrichment errors
    }

    idx = batchEnd;
    setConfigValue('ENRICH_LAST_IDX', String(idx - 1));

    Utilities.sleep(config.throttleMs);
  }

  if (idx >= assetIds.length) {
    setConfigValue('ENRICH_COMPLETE', 'TRUE');
    logOperation('Enrich', 'COMPLETE', `All ${assetIds.length} assets enriched`);
    return 'complete';
  }

  logOperation('Enrich', 'PAUSED', `At index ${idx}/${assetIds.length}. Will resume.`);
  return 'paused';
}

/**
 * Build a map of assetId -> AUE date from custom field values response
 */
function buildAueMap(cfValues, aueFieldId) {
  const map = {};
  if (!Array.isArray(cfValues)) return map;

  cfValues.forEach(item => {
    const entityId = item.EntityId || item.AssetId;
    if (!entityId) return;
    const fields = item.CustomFieldValues || item.Values || [];
    if (Array.isArray(fields)) {
      fields.forEach(f => {
        if (f.CustomFieldId === aueFieldId || f.CustomFieldTypeId === aueFieldId) {
          map[entityId] = formatDate(f.Value);
        }
      });
    }
  });
  return map;
}

// =============================================================================
// DATA MANAGEMENT
// =============================================================================

/**
 * Clear all asset data and reset progress for a full reload.
 */
function clearAssetDataAndReset() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('AssetData');
  if (sheet && sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, ASSET_TOTAL_COLS).clearContent();
  }

  resetConfigCache();
  setConfigValue('ASSET_LAST_PAGE', '-1');
  setConfigValue('ASSET_TOTAL_PAGES', '-1');
  setConfigValue('ASSET_COMPLETE', 'FALSE');
  setConfigValue('ENRICH_LAST_IDX', '-1');
  setConfigValue('ENRICH_COMPLETE', 'FALSE');

  logOperation('AssetData', 'RESET', 'All data cleared and progress reset');
}

/**
 * Get loading status summary
 */
function getLoadingStatus() {
  const config = getConfig();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('AssetData');
  const rowCount = sheet ? Math.max(sheet.getLastRow() - 1, 0) : 0;

  let phase1 = 'Not started';
  if (config.assetComplete) {
    phase1 = 'Complete';
  } else if (config.assetLastPage >= 0) {
    const displayPage = config.assetLastPage + 1;
    const displayTotal = config.assetTotalPages >= 0 ? config.assetTotalPages + 1 : '?';
    phase1 = `Page ${displayPage}/${displayTotal}`;
  }

  let phase2 = 'Not started';
  if (config.enrichComplete) {
    phase2 = 'Complete';
  } else if (config.enrichLastIdx >= 0) {
    phase2 = `Row ${config.enrichLastIdx + 1}/${rowCount}`;
  } else if (!config.aueFieldId) {
    phase2 = 'No AUE field configured';
  }

  return {
    rowCount: rowCount,
    phase1: phase1,
    phase2: phase2,
    assetComplete: config.assetComplete,
    enrichComplete: config.enrichComplete
  };
}
