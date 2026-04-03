/**
 * iiQ Asset Reporting - Asset Data Loader
 *
 * Loads assets via paginated search with checkpoint resume.
 * Supports incremental refresh via ModifiedDate filter.
 *
 * Column layout (28 columns, A-AB):
 *   A  AssetId            K  OwnerId
 *   B  AssetTag           L  OwnerName
 *   C  Name               M  StatusName
 *   D  SerialNumber       N  PurchasedDate
 *   E  ModelName          O  WarrantyExpDate
 *   F  ManufacturerName   P  PurchasePrice
 *   G  CategoryName       Q  CreatedDate
 *   H  LocationId         R  ModifiedDate
 *   I  LocationName       S  OwnerRoleName
 *   J  LocationType       T  OwnerGrade
 *                         U  OwnerLocationId
 *                         V  StorageLocationName
 *                         W  StorageUnitNumber
 *                         X  DeployedDate
 *                         Y  OpenTickets
 *                         Z  AgeDays (formula)
 *                         AA AgeYears (formula)
 *                         AB WarrantyStatus (formula)
 */

const ASSET_HEADERS = [
  'AssetId', 'AssetTag', 'Name', 'SerialNumber',
  'ModelName', 'ManufacturerName', 'CategoryName',
  'LocationId', 'LocationName', 'LocationType',
  'OwnerId', 'OwnerName', 'StatusName',
  'PurchasedDate', 'WarrantyExpDate', 'PurchasePrice',
  'CreatedDate', 'ModifiedDate',
  'OwnerRoleName', 'OwnerGrade', 'OwnerLocationId',
  'StorageLocationName', 'StorageUnitNumber', 'DeployedDate',
  'OpenTickets',
  'AgeDays', 'AgeYears', 'WarrantyStatus'
];
const ASSET_DATA_COLS = 25;  // Columns A-Y (API data)
const ASSET_TOTAL_COLS = ASSET_HEADERS.length; // 28 (includes formula columns)
const MAX_RUNTIME_MS = 5.5 * 60 * 1000;

// =============================================================================
// INITIAL LOAD: BULK ASSET LOADING (paginated with checkpoint resume)
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

  // Auto-load reference data on first run
  if (config.assetLastPage < 0) {
    ensureReferenceData(ss);
  }

  const startTime = Date.now();
  let currentPage = config.assetLastPage + 1;
  let totalPages = config.assetTotalPages;
  let totalRowsWritten = 0;

  logOperation('AssetLoad', 'START', `Resuming from page ${currentPage}`);
  cacheConfigRowPositions_();

  while (Date.now() - startTime < MAX_RUNTIME_MS) {
    const response = searchAssets([], currentPage, config.assetBatchSize, { field: 'AssetCreatedDate', direction: 'asc' });

    if (!response || !response.Items) {
      logOperation('AssetLoad', 'ERROR', `Empty response on page ${currentPage}`);
      break;
    }

    // Capture total pages on first response
    if (totalPages === -1 && response.Paging) {
      totalPages = Math.ceil(response.Paging.TotalRows / config.assetBatchSize) - 1;
      setConfigValue('ASSET_TOTAL_PAGES', String(totalPages));
    }

    // Checkpoint BEFORE writing data to prevent duplicates on resume.
    // If the write fails, we lose one page of data but the weekly reload catches it.
    // The alternative (checkpoint after write) causes duplicates if the script times out
    // between setValues and setConfigValue.
    setConfigValue('ASSET_LAST_PAGE', String(currentPage));

    // Extract and write rows
    const rows = response.Items.map(asset => extractAssetRow(asset));
    if (rows.length > 0) {
      const lastRow = Math.max(sheet.getLastRow(), 1);
      sheet.getRange(lastRow + 1, 1, rows.length, ASSET_DATA_COLS).setValues(rows);
      totalRowsWritten += rows.length;
    }

    const displayPage = currentPage + 1;
    const displayTotal = totalPages + 1;
    logOperation('AssetLoad', 'BATCH', `Page ${displayPage}/${displayTotal} (${rows.length} assets, ${totalRowsWritten} total this run)`);

    // Check completion
    if (currentPage >= totalPages) {
      setConfigValue('ASSET_COMPLETE', 'TRUE');
      setConfigValue('LAST_REFRESH_DATE', new Date().toISOString());
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

// =============================================================================
// INCREMENTAL REFRESH (ModifiedDate-based, in-place updates)
// =============================================================================

/**
 * Refresh assets modified since the last refresh date.
 * Updates existing rows in-place by AssetId, appends new assets.
 * @param {boolean} showUI - Whether to show UI alerts
 * @returns {Object|string} - {updated, added} counts, or status string
 */
function refreshAssetData(showUI) {
  const config = getConfig();
  if (!config.assetComplete) return 'initial_load_incomplete';

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('AssetData');
  if (!sheet) return 'no_sheet';

  const sinceDate = config.lastRefreshDate;
  if (!sinceDate) return 'no_refresh_date';

  const startTime = Date.now();
  const lastRow = sheet.getLastRow();

  // Build AssetId -> row index for in-place updates
  const assetIdValues = lastRow > 1
    ? sheet.getRange(2, 1, lastRow - 1, 1).getValues()
    : [];
  const idToRow = {};
  assetIdValues.forEach((row, i) => { if (row[0]) idToRow[row[0]] = i + 2; });

  // Format date for iiQ filter: date>=YYYY-MM-DD
  const sinceDateFormatted = sinceDate.substring(0, 10); // ISO -> YYYY-MM-DD
  const filters = [{ Facet: 'modifieddate', Value: 'date>=' + sinceDateFormatted }];

  logOperation('Refresh', 'START', `Fetching assets modified since ${sinceDateFormatted}`);
  cacheConfigRowPositions_();

  let page = 0;
  const updates = []; // [{row, data}]
  const newRows = [];

  while (Date.now() - startTime < MAX_RUNTIME_MS) {
    const response = searchAssets(filters, page, config.assetBatchSize, { field: 'AssetModifiedDate', direction: 'asc' });
    if (!response || !response.Items || response.Items.length === 0) break;

    response.Items.forEach(asset => {
      const row = extractAssetRow(asset);
      const assetId = asset.AssetId || '';
      const existingRow = idToRow[assetId];
      if (existingRow) {
        updates.push({ row: existingRow, data: row });
      } else if (assetId) {
        newRows.push(row);
        idToRow[assetId] = lastRow + newRows.length; // track for dedup within run
      }
    });

    const totalResults = response.Paging ? response.Paging.TotalRows : 0;
    logOperation('Refresh', 'BATCH', `Page ${page + 1}, ${response.Items.length} assets (${totalResults} total modified)`);

    if ((page + 1) * config.assetBatchSize >= totalResults) break;
    page++;
    Utilities.sleep(config.throttleMs);
  }

  // Write updates in-place
  updates.forEach(u => {
    sheet.getRange(u.row, 1, 1, ASSET_DATA_COLS).setValues([u.data]);
  });

  // Append new rows
  if (newRows.length > 0) {
    const appendStart = sheet.getLastRow() + 1;
    sheet.getRange(appendStart, 1, newRows.length, ASSET_DATA_COLS).setValues(newRows);
  }

  // Update refresh timestamp
  setConfigValue('LAST_REFRESH_DATE', new Date().toISOString());

  logOperation('Refresh', 'COMPLETE', `${updates.length} updated, ${newRows.length} new assets`);
  return { updated: updates.length, added: newRows.length };
}

// =============================================================================
// ROW EXTRACTION
// =============================================================================

/**
 * Extract one row of asset data from an API response item.
 * Returns array of ASSET_DATA_COLS values (columns A-Y).
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
    // Owner enrichment
    owner.RoleName || '',
    owner.Grade ?? '',
    owner.LocationId || '',
    // Storage
    asset.StorageLocationName || '',
    asset.StorageUnitNumber || '',
    formatDate(asset.DeployedDate),
    // Tickets
    asset.OpenTicketCount ?? asset.OpenTickets ?? '',
  ];
}

function formatDate(val) {
  if (!val) return '';
  const d = new Date(val);
  if (isNaN(d.getTime())) return '';
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
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
  setConfigValue('LAST_REFRESH_DATE', '');

  logOperation('AssetData', 'RESET', 'All data cleared and progress reset');
}

/**
 * Remove duplicate rows from AssetData by AssetId (column A).
 * Keeps the LAST occurrence of each AssetId (most recently written = most up-to-date).
 * Reads all data into memory, deduplicates, clears the sheet, and writes back in one batch.
 * Reapplies formulas after dedup.
 * @returns {number} Number of duplicate rows removed
 */
function deduplicateAssetData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('AssetData');
  if (!sheet) return 0;

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 0;

  // Read all data columns (API data only, not formula columns)
  const allData = sheet.getRange(2, 1, lastRow - 1, ASSET_DATA_COLS).getValues();

  // Walk backwards: keep last occurrence of each AssetId (most up-to-date)
  const seen = new Set();
  const uniqueRows = [];
  for (let i = allData.length - 1; i >= 0; i--) {
    const id = allData[i][0];
    if (!id || seen.has(id)) continue;
    seen.add(id);
    uniqueRows.push(allData[i]);
  }

  const removed = allData.length - uniqueRows.length;
  if (removed === 0) return 0;

  // Reverse to restore original order (we walked backwards)
  uniqueRows.reverse();

  // Clear all data rows and write back deduplicated set in one batch
  sheet.getRange(2, 1, lastRow - 1, ASSET_TOTAL_COLS).clearContent();
  sheet.getRange(2, 1, uniqueRows.length, ASSET_DATA_COLS).setValues(uniqueRows);

  // Reapply formulas to cover the deduplicated rows
  applyAssetFormulas();

  logOperation('Dedup', 'COMPLETE', `Removed ${removed} duplicate rows (${uniqueRows.length} unique remain)`);
  return removed;
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

  const lastRefresh = config.lastRefreshDate
    ? config.lastRefreshDate.substring(0, 19).replace('T', ' ')
    : 'Never';

  return {
    rowCount: rowCount,
    phase1: phase1,
    lastRefresh: lastRefresh,
    assetComplete: config.assetComplete
  };
}
