/**
 * Improved search emails functionality
 */
const ENABLE_RECENT_EMAILS_FALLBACK = false;

const config = require('../config');
const { callGraphAPI, callGraphAPIPaginated } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { resolveFolderPath } = require('./folder-utils');
const { resolveIanaTimezone, formatEmailDate } = require('../utils/date-helpers');

/**
 * Sanitize a value for use inside a KQL expression.
 * Strips characters that break KQL parsing (single quotes, backslashes, colons
 * and double quotes inside values).
 * @param {string} value - Raw search value
 * @returns {string} - Sanitized value safe for KQL
 */
function sanitizeKqlValue(value) {
  if (!value) return value;

  // Remove backslashes and single quotes (KQL has no escape mechanism for these)
  let sanitized = value.replace(/[\\']/g, '');

  // Remove colons that would be misinterpreted as KQL field operators
  sanitized = sanitized.replace(/:/g, '');

  // Remove interior double quotes to avoid breaking the KQL string wrapper
  sanitized = sanitized.replace(/"/g, '');

  // Trim whitespace that may be left over
  sanitized = sanitized.trim();

  return sanitized;
}

/**
 * Split a sanitized KQL value into non-empty tokens.
 * @param {string} value - Raw search value
 * @returns {string[]} Sanitized tokens
 */
function tokenizeKqlValue(value) {
  const sanitized = sanitizeKqlValue(value);
  if (!sanitized) return [];
  return sanitized.split(/\s+/).filter(Boolean);
}

/**
 * Build KQL terms for a bare query value.
 * @param {string} value - Raw query value
 * @param {boolean} exactPhrase - Whether to preserve phrase semantics
 * @returns {string[]} KQL terms
 */
function buildBareQueryTerms(value, exactPhrase = false) {
  const tokens = tokenizeKqlValue(value);
  if (tokens.length === 0) return [];
  if (exactPhrase && tokens.length > 1) {
    return [`"${tokens.join(' ')}"`];
  }
  return tokens;
}

/**
 * Build KQL terms for a field-qualified value.
 * @param {string} field - KQL field name
 * @param {string} value - Raw field value
 * @param {boolean} exactPhrase - Whether to preserve phrase semantics
 * @returns {string[]} KQL terms
 */
function buildFieldQueryTerms(field, value, exactPhrase = false) {
  const tokens = tokenizeKqlValue(value);
  if (tokens.length === 0) return [];
  if (exactPhrase && tokens.length > 1) {
    return [`${field}:"${tokens.join(' ')}"`];
  }
  return tokens.map(token => `${field}:${token}`);
}

/**
 * Build text KQL terms from the supported search inputs.
 * @param {object} searchTerms - Search terms (query, from, to, subject)
 * @param {object} matchOptions - Exact phrase options
 * @returns {string[]} KQL text terms
 */
function buildTextKqlTerms(searchTerms, matchOptions = {}) {
  const kqlTerms = [];

  kqlTerms.push(...buildBareQueryTerms(searchTerms.query, matchOptions.queryExactPhrase === true));
  kqlTerms.push(...buildFieldQueryTerms('subject', searchTerms.subject, matchOptions.subjectExactPhrase === true));
  kqlTerms.push(...buildFieldQueryTerms('from', searchTerms.from, matchOptions.fromExactPhrase === true));
  kqlTerms.push(...buildFieldQueryTerms('to', searchTerms.to, matchOptions.toExactPhrase === true));

  return kqlTerms;
}

/**
 * Build a KQL fragment string for boolean filter terms.
 * Returns terms like "hasAttachments:true isRead:false" or empty string.
 * @param {object} filterTerms - { hasAttachments, unreadOnly }
 * @returns {string} KQL fragment (may be empty)
 */
function buildKqlBooleans(filterTerms) {
  const parts = [];
  if (filterTerms.hasAttachments === true) {
    parts.push('hasAttachments:true');
  }
  // Note: isRead is NOT a supported KQL $search property in Graph API.
  // unreadOnly is enforced via post-filtering in progressiveSearch().
  return parts.join(' ');
}

/**
 * Post-filter results to enforce unreadOnly when $search prevents use of $filter.
 * @param {object} response - Graph API response with .value array
 * @param {boolean} unreadOnly - Whether to filter to unread only
 * @returns {object} - Filtered response
 */
function applyPostFilters(response, filterTerms) {
  if (!response.value || !filterTerms.unreadOnly) return response;
  const filtered = response.value.filter(email => !email.isRead);
  return { ...response, value: filtered };
}

/**
 * Search emails handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleSearchEmails(args) {
  const folder = args.folder || "inbox";
  const requestedCount = args.count || 10;
  const query = args.query || '';
  const from = args.from || '';
  const to = args.to || '';
  const subject = args.subject || '';
  const hasAttachments = args.hasAttachments;
  const unreadOnly = args.unreadOnly;
  const matchOptions = {
    queryExactPhrase: args.queryExactPhrase === true,
    fromExactPhrase: args.fromExactPhrase === true,
    toExactPhrase: args.toExactPhrase === true,
    subjectExactPhrase: args.subjectExactPhrase === true,
  };
  
  try {
    // Get access token
    const accessToken = await ensureAuthenticated();
    
    // Resolve the folder path
    const endpoint = await resolveFolderPath(accessToken, folder);
    console.error(`Using endpoint: ${endpoint} for folder: ${folder}`);
    
    // Execute progressive search with pagination
    const response = await progressiveSearch(
      endpoint, 
      accessToken, 
      { query, from, to, subject },
      { hasAttachments, unreadOnly },
      requestedCount,
      matchOptions
    );
    
    return formatSearchResults(response);
  } catch (error) {
    // Handle authentication errors
    if (error.message === 'Authentication required') {
      return {
        content: [{ 
          type: "text", 
          text: "Authentication required. Please use the 'authenticate' tool first."
        }]
      };
    }
    
    // General error response
    return {
      content: [{ 
        type: "text", 
        text: `Error searching emails: ${error.message}`
      }]
    };
  }
}

/**
 * Execute a search with progressively simpler fallback strategies
 * @param {string} endpoint - API endpoint
 * @param {string} accessToken - Access token
 * @param {object} searchTerms - Search terms (query, from, to, subject)
 * @param {object} filterTerms - Filter terms (hasAttachments, unreadOnly)
 * @param {number} maxCount - Maximum number of results to retrieve
 * @returns {Promise<object>} - Search results
 */
async function progressiveSearch(endpoint, accessToken, searchTerms, filterTerms, maxCount, matchOptions = {}) {
  // Track search strategies attempted
  const searchAttempts = [];

  // When unreadOnly is combined with $search, we need to over-fetch since
  // isRead is not a supported KQL property and must be post-filtered.
  const hasRawSearchTerms = Object.values(searchTerms).some(v => v);
  const hasEffectiveKqlTerms = buildTextKqlTerms(searchTerms, matchOptions).length > 0;
  const hasBooleanFilters = filterTerms.hasAttachments === true || filterTerms.unreadOnly === true;
  const needsPostFilter = filterTerms.unreadOnly === true && hasEffectiveKqlTerms;
  const fetchCount = needsPostFilter ? Math.min(50, maxCount * 3) : Math.min(50, maxCount);

  // Avoid treating a fully sanitized-away search string as an unfiltered mailbox query.
  if (hasRawSearchTerms && !hasEffectiveKqlTerms && !hasBooleanFilters) {
    return {
      value: [],
      _searchInfo: {
        attemptsCount: 0,
        strategies: [],
        originalTerms: searchTerms,
        filterTerms: filterTerms,
        warning: 'Search terms were removed during sanitization. Please refine your query.'
      }
    };
  }
  
  // 1. Try combined search (most specific)
  try {
    const params = buildSearchParams(searchTerms, filterTerms, fetchCount, matchOptions);
    console.error("Attempting combined search with params:", params);
    searchAttempts.push("combined-search");

    const fetchMax = needsPostFilter ? fetchCount : maxCount;
    const response = await callGraphAPIPaginated(accessToken, 'GET', endpoint, params, fetchMax);
    const filtered = applyPostFilters(response, filterTerms);
    if (filtered.value && filtered.value.length > 0) {
      filtered.value = filtered.value.slice(0, maxCount);
      console.error(`Combined search successful: found ${filtered.value.length} results`);
      return filtered;
    }
  } catch (error) {
    console.error(`Combined search failed: ${error.message}`);
  }
  
  // 2. Try each search term individually, starting with most specific
  const searchPriority = ['subject', 'from', 'to', 'query'];
  
  for (const term of searchPriority) {
    if (searchTerms[term]) {
      try {
        console.error(`Attempting search with only ${term}: "${searchTerms[term]}"`);
        const singleTermSearchTerms = { query: '', from: '', to: '', subject: '' };
        singleTermSearchTerms[term] = searchTerms[term];

        const simplifiedParams = buildSearchParams(
          singleTermSearchTerms,
          filterTerms,
          fetchCount,
          matchOptions
        );

        if (!simplifiedParams.$search) {
          continue;
        }

        searchAttempts.push(`single-term-${term}`);

        const singleFetchMax = needsPostFilter ? fetchCount : maxCount;
        const response = await callGraphAPIPaginated(accessToken, 'GET', endpoint, simplifiedParams, singleFetchMax);
        const filtered = applyPostFilters(response, filterTerms);
        if (filtered.value && filtered.value.length > 0) {
          filtered.value = filtered.value.slice(0, maxCount);
          console.error(`Search with ${term} successful: found ${filtered.value.length} results`);
          filtered._searchInfo = {
            attemptsCount: searchAttempts.length,
            strategies: [...searchAttempts],
            originalTerms: searchTerms,
            filterTerms: filterTerms
          };
          return filtered;
        }
      } catch (error) {
        console.error(`Search with ${term} failed: ${error.message}`);
      }
    }
  }
  
  // 3. Try with only boolean filters (using $filter since there's no $search)
  if (hasBooleanFilters) {
    try {
      console.error("Attempting search with only boolean filters");
      searchAttempts.push("boolean-filters-only");

      const filterOnlyParams = {
        $top: Math.min(50, maxCount),
        $select: config.EMAIL_SELECT_FIELDS,
        $orderby: 'receivedDateTime desc'
      };

      // No $search here, so $filter is safe
      addBooleanFilters(filterOnlyParams, filterTerms);

      const response = await callGraphAPIPaginated(accessToken, 'GET', endpoint, filterOnlyParams, maxCount);
      console.error(`Boolean filter search found ${response.value?.length || 0} results`);

      // Warn that search terms were dropped (Bug 2: no silent data loss)
      if (hasRawSearchTerms) {
        response._searchInfo = {
          attemptsCount: searchAttempts.length,
          strategies: [...searchAttempts],
          originalTerms: searchTerms,
          filterTerms: filterTerms,
          warning: 'Search terms were dropped; results are filtered by boolean criteria only.'
        };
      }
      return response;
    } catch (error) {
      console.error(`Boolean filter search failed: ${error.message}`);
    }
  }
  
  if (ENABLE_RECENT_EMAILS_FALLBACK) {
    // 4. Final fallback: just get recent emails with pagination
    console.error("All search strategies failed, falling back to recent emails");
    searchAttempts.push("recent-emails");
    
    const basicParams = {
      $top: Math.min(50, maxCount),
      $select: config.EMAIL_SELECT_FIELDS,
      $orderby: 'receivedDateTime desc'
    };
    
    const response = await callGraphAPIPaginated(accessToken, 'GET', endpoint, basicParams, maxCount);
    console.error(`Fallback to recent emails found ${response.value?.length || 0} results`);
    
    // Add a note to the response about the search attempts
    response._searchInfo = {
      attemptsCount: searchAttempts.length,
      strategies: searchAttempts,
      originalTerms: searchTerms,
      filterTerms: filterTerms
    };
    
    return response;
  }

  return {
    value: [],
    _searchInfo: {
      attemptsCount: searchAttempts.length,
      strategies: searchAttempts,
      originalTerms: searchTerms,
      filterTerms: filterTerms,
      warning: 'All search strategies failed. No results returned (recent-emails fallback is disabled).'
    }
  };
}

/**
 * Build search parameters from search terms and filter terms
 * @param {object} searchTerms - Search terms (query, from, to, subject)
 * @param {object} filterTerms - Filter terms (hasAttachments, unreadOnly)
 * @param {number} count - Maximum number of results
 * @param {object} matchOptions - Exact phrase options
 * @returns {object} - Query parameters
 */
function buildSearchParams(searchTerms, filterTerms, count, matchOptions = {}) {
  const params = {
    $top: count,
    $select: config.EMAIL_SELECT_FIELDS
  };

  const kqlTerms = buildTextKqlTerms(searchTerms, matchOptions);

  // When we have KQL search terms, fold boolean filters into the KQL string
  // (Graph API does NOT allow combining $search with $filter)
  if (kqlTerms.length > 0) {
    const boolKql = buildKqlBooleans(filterTerms);
    if (boolKql) {
      kqlTerms.push(boolKql);
    }
    const joined = kqlTerms.join(' ');
    // If any term already contains inner double quotes (from exact-phrase
    // construction like subject:"Q4 Report"), the expression is valid KQL
    // as-is and must NOT be wrapped in outer quotes — KQL has no escape
    // mechanism for nested quotes.
    params.$search = joined.includes('"') ? joined : `"${joined}"`;
  } else {
    // No KQL search terms — use $filter for booleans
    addBooleanFilters(params, filterTerms);
  }

  return params;
}

/**
 * Add boolean filters to query parameters
 * @param {object} params - Query parameters
 * @param {object} filterTerms - Filter terms (hasAttachments, unreadOnly)
 */
function addBooleanFilters(params, filterTerms) {
  const filterConditions = [];
  
  if (filterTerms.hasAttachments === true) {
    filterConditions.push('hasAttachments eq true');
  }
  
  if (filterTerms.unreadOnly === true) {
    filterConditions.push('isRead eq false');
  }
  
  // Add $filter parameter if we have any filter conditions
  if (filterConditions.length > 0) {
    params.$filter = filterConditions.join(' and ');
  }
}

/**
 * Format search results into a readable text format
 * @param {object} response - The API response object
 * @returns {object} - MCP response object
 */
function formatSearchResults(response) {
  if (!response.value || response.value.length === 0) {
    let text = 'No emails found matching your search criteria.';
    if (response._searchInfo) {
      const attemptedStrategies = response._searchInfo.strategies?.join(', ');
      if (attemptedStrategies) {
        text += ` Strategies attempted: ${attemptedStrategies}.`;
      }
      if (response._searchInfo.warning) {
        text += ` ${response._searchInfo.warning}`;
      }
    }

    return {
      content: [{ 
        type: "text", 
        text
      }]
    };
  }
  
  // Format results
  const ianaTz = resolveIanaTimezone(config.DEFAULT_TIMEZONE);
  const emailList = response.value.map((email, index) => {
    const sender = email.from?.emailAddress || { name: 'Unknown', address: 'unknown' };
    const date = formatEmailDate(email.receivedDateTime, ianaTz);
    const readStatus = email.isRead ? '' : '[UNREAD] ';

    const preview = email.bodyPreview ? `\nPreview: ${email.bodyPreview}` : '';
    return `${index + 1}. ${readStatus}${date} - From: ${sender.name} (${sender.address})\nSubject: ${email.subject}${preview}\nID: ${email.id}\n`;
  }).join("\n");
  
  // Add search strategy info if available
  let additionalInfo = '';
  if (response._searchInfo) {
    additionalInfo = `\n(Search used ${response._searchInfo.strategies[response._searchInfo.strategies.length - 1]} strategy)`;
  }
  
  return {
    content: [{ 
      type: "text", 
      text: `Found ${response.value.length} emails matching your search criteria:${additionalInfo}\n\n${emailList}`
    }]
  };
}

module.exports = handleSearchEmails;

// Export internals for testing
module.exports._internal = {
  ENABLE_RECENT_EMAILS_FALLBACK,
  sanitizeKqlValue,
  tokenizeKqlValue,
  buildBareQueryTerms,
  buildFieldQueryTerms,
  buildTextKqlTerms,
  buildKqlBooleans,
  applyPostFilters,
  buildSearchParams,
  addBooleanFilters,
  progressiveSearch,
  formatSearchResults,
};
