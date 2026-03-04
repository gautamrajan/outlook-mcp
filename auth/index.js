/**
 * Authentication module for Outlook MCP server
 *
 * Supports two modes:
 *   - Local mode: single-user device code flow via TokenStorage (existing behavior)
 *   - Hosted mode: multi-user OBO flow via PerUserTokenStorage + exchangeOBO
 *
 * The mode is determined automatically by checking whether a per-request
 * user context exists (set by the HTTP transport layer via AsyncLocalStorage).
 */
const TokenStorage = require('./token-storage');
const PerUserTokenStorage = require('./per-user-token-storage');
const { getUserContext, isHostedMode } = require('./request-context');
const { exchangeOBO } = require('./obo-exchange');
const config = require('../config');
const { authTools } = require('./tools');

// Singleton TokenStorage instance for local (single-user) mode
const tokenStorage = new TokenStorage({
  tokenStorePath: config.AUTH_CONFIG.tokenStorePath,
  clientId: config.AUTH_CONFIG.clientId,
  clientSecret: config.AUTH_CONFIG.clientSecret,
  tokenEndpoint: config.AUTH_CONFIG.tokenEndpoint,
  redirectUri: config.AUTH_CONFIG.redirectUri,
  scopes: config.AUTH_CONFIG.scopes
});

// Singleton per-user token storage for hosted (multi-user) mode
const perUserTokenStorage = new PerUserTokenStorage();

/**
 * Ensures the user is authenticated and returns an access token.
 *
 * In hosted mode (user context present via AsyncLocalStorage):
 *   - Checks per-user cache for a valid Graph token
 *   - If missing/expired, exchanges the Entra JWT for a Graph token via OBO
 *   - Stores the result in per-user cache
 *
 * In local mode (no user context):
 *   - Uses the existing singleton TokenStorage flow (completely unchanged)
 *
 * @param {{ forceRefresh?: boolean }} options
 * @returns {Promise<string>} - Access token
 * @throws {Error} - If authentication fails
 */
async function ensureAuthenticated({ forceRefresh = false } = {}) {
  if (isHostedMode()) {
    return _ensureAuthenticatedHosted(forceRefresh);
  }

  return _ensureAuthenticatedLocal(forceRefresh);
}

/**
 * Local mode authentication — existing single-user device code flow.
 * @param {boolean} forceRefresh
 * @returns {Promise<string>}
 */
async function _ensureAuthenticatedLocal(forceRefresh) {
  if (forceRefresh) {
    tokenStorage.invalidateAccessToken();
  }

  const accessToken = await tokenStorage.getValidAccessToken();
  if (!accessToken) {
    throw new Error('Authentication required');
  }

  return accessToken;
}

/**
 * Hosted mode authentication — multi-user OBO flow.
 * @param {boolean} forceRefresh
 * @returns {Promise<string>}
 */
async function _ensureAuthenticatedHosted(forceRefresh) {
  const { userId, entraToken } = getUserContext();

  if (forceRefresh) {
    perUserTokenStorage.invalidateUser(userId);
  }

  // Check per-user cache first
  const cachedToken = perUserTokenStorage.getTokenForUser(userId);
  if (cachedToken) {
    return cachedToken;
  }

  // Cache miss or expired — exchange via OBO
  const oboConfig = {
    clientId: config.AUTH_CONFIG.clientId,
    clientSecret: config.AUTH_CONFIG.clientSecret,
    tenantId: config.AUTH_CONFIG.tenantId,
    scopes: config.AUTH_CONFIG.scopes.filter(s => s !== 'offline_access'),
  };

  const tokenData = await exchangeOBO(entraToken, oboConfig);
  perUserTokenStorage.setTokenForUser(userId, tokenData);

  return tokenData.access_token;
}

module.exports = {
  tokenStorage,
  authTools,
  ensureAuthenticated
};
