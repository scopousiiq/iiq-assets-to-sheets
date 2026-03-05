/**
 * iiQ Asset Reporting - API Client
 * HTTP requests with retry/exponential backoff for rate limiting.
 */

const MAX_RETRIES = 3;
const BACKOFF_BASE_MS = 2000;

function getThrottleMs() {
  try {
    const config = getConfig();
    return config.throttleMs || 1000;
  } catch (e) {
    return 1000;
  }
}

function makeApiRequest(endpoint, method, payload, retryCount) {
  const config = getConfig();
  retryCount = retryCount || 0;

  if (!config.baseUrl || !config.bearerToken) {
    throw new Error('API configuration missing. Check Config sheet.');
  }

  const url = config.baseUrl + endpoint;
  const options = {
    method: method || 'GET',
    headers: {
      'Authorization': 'Bearer ' + config.bearerToken,
      'Content-Type': 'application/json',
      'Accept': 'application/json'
    },
    muteHttpExceptions: true
  };

  if (config.siteId) options.headers['SiteId'] = config.siteId;
  if (payload) options.payload = JSON.stringify(payload);

  try {
    const startTime = Date.now();
    const response = UrlFetchApp.fetch(url, options);
    const elapsedMs = Date.now() - startTime;
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    if (responseCode >= 200 && responseCode < 300) {
      Utilities.sleep(Math.floor(getThrottleMs() * 0.5));
      return JSON.parse(responseText);
    } else if ((responseCode === 429 || responseCode === 503) && retryCount < MAX_RETRIES) {
      const backoffMs = BACKOFF_BASE_MS * Math.pow(2, retryCount);
      logOperation('API', 'RETRY', `${endpoint} - HTTP ${responseCode}, waiting ${backoffMs}ms (${retryCount + 1}/${MAX_RETRIES})`);
      Utilities.sleep(backoffMs);
      return makeApiRequest(endpoint, method, payload, retryCount + 1);
    } else {
      logOperation('API', 'ERROR', `${endpoint} - HTTP ${responseCode}: ${responseText.substring(0, 200)}`);
      throw new Error(`API Error ${responseCode}: ${responseText.substring(0, 200)}`);
    }
  } catch (e) {
    if (e.message && e.message.includes('API Error')) throw e;
    if (retryCount < MAX_RETRIES) {
      const backoffMs = BACKOFF_BASE_MS * Math.pow(2, retryCount);
      logOperation('API', 'NETWORK_ERROR', `${endpoint} - ${e.message} - retry in ${backoffMs}ms`);
      Utilities.sleep(backoffMs);
      return makeApiRequest(endpoint, method, payload, retryCount + 1);
    }
    throw e;
  }
}

/**
 * Search assets with filters
 * @param {Array} filters - Array of filter objects [{Facet, Id/Value}]
 * @param {number} page - Page index (0-based)
 * @param {number} pageSize - Records per page
 * @returns {Object} - API response with Items and Paging
 */
function searchAssets(filters, page, pageSize) {
  const config = getConfig();
  const size = pageSize || config.assetBatchSize;
  const endpoint = `/v1.0/assets?$p=${page || 0}&$s=${size}`;
  return makeApiRequest(endpoint, 'POST', { Filters: filters || [] });
}

/**
 * Get single asset with full detail
 * @param {string} assetId - Asset UUID
 * @returns {Object} - Full asset object
 */
function getAssetDetail(assetId) {
  return makeApiRequest(`/v1.0/assets/${assetId}`, 'GET');
}

/**
 * Get all locations
 * @returns {Array} - Array of location objects
 */
function getAllLocations() {
  const response = makeApiRequest('/v2.0/locations/all?$s=1000', 'GET');
  return response.Items || response || [];
}

/**
 * Get all asset status types
 * @returns {Array} - Array of status type objects
 */
function getAllAssetStatusTypes() {
  const response = makeApiRequest('/v1.0/assets/status/types?$s=100', 'GET');
  return response.Items || response || [];
}

/**
 * Get custom field definitions for assets
 * @returns {Array} - Array of custom field definitions
 */
function getAssetCustomFields() {
  const response = makeApiRequest('/v1.0/custom-fields/for/asset', 'POST', {});
  return response.Items || response || [];
}

/**
 * Get custom field values for a batch of assets
 * @param {Array} assetIds - Array of asset UUID strings
 * @returns {Object} - Map of assetId -> custom field values
 */
function getCustomFieldValuesForAssets(assetIds) {
  const response = makeApiRequest('/v1.0/custom-fields/values/for/assets', 'POST', assetIds);
  return response.Items || response || [];
}
