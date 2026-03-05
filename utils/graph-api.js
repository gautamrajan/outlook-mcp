/**
 * Microsoft Graph API helper functions
 */
const https = require('https');
const config = require('../config');
const mockData = require('./mock-data');

let _ensureAuthenticated = null;
function getEnsureAuthenticated() {
  if (!_ensureAuthenticated) {
    const { ensureAuthenticated } = require('../auth');
    _ensureAuthenticated = ensureAuthenticated;
  }
  return _ensureAuthenticated;
}

function _buildUrl(path, queryParams) {
  if (path.startsWith('http://') || path.startsWith('https://')) {
    console.error(`Using full URL from nextLink: ${path}`);
    return path;
  }

  const encodedPath = path.split('/')
    .map(segment => encodeURIComponent(segment))
    .join('/');

  let queryString = '';
  if (Object.keys(queryParams).length > 0) {
    const filter = queryParams.$filter;
    if (filter) {
      delete queryParams.$filter;
    }

    const params = new URLSearchParams();
    for (const [key, value] of Object.entries(queryParams)) {
      params.append(key, value);
    }

    queryString = params.toString();

    if (filter) {
      queryString += (queryString ? '&' : '') + `$filter=${encodeURIComponent(filter)}`;
    }

    if (queryString) {
      queryString = '?' + queryString;
    }
    console.error(`Query string: ${queryString}`);
  }

  const finalUrl = `${config.GRAPH_API_ENDPOINT}${encodedPath}${queryString}`;
  console.error(`Full URL: ${finalUrl}`);
  return finalUrl;
}

function _callGraphAPIOnce(accessToken, method, finalUrl, data, extraHeaders = {}) {
  return new Promise((resolve, reject) => {
    const req = https.request(finalUrl, {
      method,
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Content-Type': 'application/json',
        ...extraHeaders
      }
    }, (res) => {
      let responseData = '';
      res.on('data', (chunk) => { responseData += chunk; });
      res.on('end', () => {
        if (res.statusCode >= 200 && res.statusCode < 300) {
          try {
            resolve(JSON.parse(responseData || '{}'));
          } catch (error) {
            reject(new Error(`Error parsing API response: ${error.message}`));
          }
        } else if (res.statusCode === 401) {
          reject(new Error('UNAUTHORIZED'));
        } else {
          reject(new Error(`API call failed with status ${res.statusCode}: ${responseData}`));
        }
      });
    });

    req.on('error', (error) => {
      reject(new Error(`Network error during API call: ${error.message}`));
    });

    if (data && (method === 'POST' || method === 'PATCH' || method === 'PUT')) {
      req.write(JSON.stringify(data));
    }
    req.end();
  });
}

/**
 * Makes a request to the Microsoft Graph API.
 * Automatically retries once on 401 by refreshing the access token.
 */
async function callGraphAPI(accessToken, method, path, data = null, queryParams = {}, extraHeaders = {}) {
  if (config.USE_TEST_MODE && accessToken.startsWith('test_access_token_')) {
    console.error(`TEST MODE: Simulating ${method} ${path} API call`);
    return mockData.simulateGraphAPIResponse(method, path, data, queryParams);
  }

  console.error(`Making real API call: ${method} ${path}`);
  const finalUrl = _buildUrl(path, queryParams);

  try {
    return await _callGraphAPIOnce(accessToken, method, finalUrl, data, extraHeaders);
  } catch (error) {
    if (error.message !== 'UNAUTHORIZED') throw error;

    console.error('Got 401 — forcing token refresh and retrying.');
    try {
      const newToken = await getEnsureAuthenticated()({ forceRefresh: true });
      return await _callGraphAPIOnce(newToken, method, finalUrl, data, extraHeaders);
    } catch (retryError) {
      throw retryError;
    }
  }
}

/**
 * Calls Graph API with pagination support to retrieve all results up to maxCount.
 * Each page request benefits from the 401 retry in callGraphAPI.
 */
async function callGraphAPIPaginated(accessToken, method, path, queryParams = {}, maxCount = 0) {
  if (method !== 'GET') {
    throw new Error('Pagination only supports GET requests');
  }

  const allItems = [];
  let nextLink = null;
  let currentUrl = path;
  let currentParams = queryParams;

  do {
    const response = await callGraphAPI(accessToken, method, currentUrl, null, currentParams);

    if (response.value && Array.isArray(response.value)) {
      allItems.push(...response.value);
      console.error(`Pagination: Retrieved ${response.value.length} items, total so far: ${allItems.length}`);
    }

    if (maxCount > 0 && allItems.length >= maxCount) {
      console.error(`Pagination: Reached max count of ${maxCount}, stopping`);
      break;
    }

    nextLink = response['@odata.nextLink'];
    if (nextLink) {
      currentUrl = nextLink;
      currentParams = {};
      console.error(`Pagination: Following nextLink, ${allItems.length} items so far`);
    }
  } while (nextLink);

  const finalItems = maxCount > 0 ? allItems.slice(0, maxCount) : allItems;
  console.error(`Pagination complete: Retrieved ${finalItems.length} total items`);

  return {
    value: finalItems,
    '@odata.count': finalItems.length
  };
}

module.exports = {
  callGraphAPI,
  callGraphAPIPaginated
};
