/**
 * Authentication module for Outlook MCP server
 *
 * Supports two modes:
 *   - Local mode: single-user device code flow via TokenStorage (existing behavior)
 *   - Hosted mode: multi-user direct auth via PerUserTokenStorage + silent refresh
 *
 * The mode is determined automatically by checking whether a per-request
 * user context exists (set by the HTTP transport layer via AsyncLocalStorage).
 */
const TokenStorage = require('./token-storage');
const PerUserTokenStorage = require('./per-user-token-storage');
const { isHostedMode, getUserContext } = require('./request-context');
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

// Hosted-mode per-user token storage (initialized lazily or injected by HTTP server)
let hostedTokenStorage = null;

/**
 * Returns the hosted-mode PerUserTokenStorage singleton, creating it lazily
 * if one has not been injected via setHostedTokenStorage().
 * @returns {PerUserTokenStorage}
 */
function getHostedTokenStorage() {
  if (!hostedTokenStorage) {
    hostedTokenStorage = new PerUserTokenStorage({
      filePath: config.AUTH_CONFIG.hostedTokenStorePath || config.AUTH_CONFIG.tokenStorePath,
      encryptionKey: process.env.TOKEN_ENCRYPTION_KEY,
    });
  }
  return hostedTokenStorage;
}

/**
 * Allows the HTTP server layer to inject a pre-configured PerUserTokenStorage
 * instance (e.g., one that has already been loaded from disk).
 * @param {PerUserTokenStorage} storage
 */
function setHostedTokenStorage(storage) {
  hostedTokenStorage = storage;
}

/**
 * Ensures the user is authenticated and returns an access token.
 *
 * In hosted mode (user context present via AsyncLocalStorage):
 *   - Looks up the user's Graph token from PerUserTokenStorage
 *   - Silently refreshes expired tokens using the stored refresh token
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
 * Hosted mode authentication — per-user token lookup with silent refresh.
 * @param {boolean} forceRefresh
 * @returns {Promise<string>}
 */
async function _ensureAuthenticatedHosted(forceRefresh) {
  const ctx = getUserContext();
  if (!ctx || !ctx.userId) {
    throw new Error('No user context — session middleware should have set this');
  }

  const { userId } = ctx;
  const storage = getHostedTokenStorage();

  // Check if we have a valid (non-expired) access token
  if (!forceRefresh) {
    const accessToken = storage.getTokenForUser(userId);
    if (accessToken) {
      return accessToken;
    }
  }

  // Access token expired or force refresh — try refreshing
  const refreshToken = storage.getRefreshToken(userId);
  if (!refreshToken) {
    const err = new Error('Authentication required');
    err.code = 'AUTH_REQUIRED';
    throw err;
  }

  // Concurrent refresh protection: per-user mutex
  return _refreshWithLock(userId, refreshToken, storage);
}

// ── Per-user refresh mutex ──────────────────────────────────────────────

const _refreshLocks = new Map(); // userId -> Promise

/**
 * Ensures only one token refresh is in-flight per user at any time.
 * Additional callers wait for the first refresh to complete, then check
 * whether the new token is usable.
 *
 * @param {string} userId
 * @param {string} refreshToken
 * @param {PerUserTokenStorage} storage
 * @returns {Promise<string>}
 */
async function _refreshWithLock(userId, refreshToken, storage) {
  // If a refresh is already in-flight for this user, wait for it
  if (_refreshLocks.has(userId)) {
    await _refreshLocks.get(userId);
    // After waiting, check if the token is now valid
    const accessToken = storage.getTokenForUser(userId);
    if (accessToken) return accessToken;
    // If still expired, fall through to refresh again
  }

  // Start a new refresh
  const refreshPromise = _doRefresh(userId, refreshToken, storage);
  _refreshLocks.set(userId, refreshPromise);

  try {
    const result = await refreshPromise;
    return result;
  } finally {
    _refreshLocks.delete(userId);
  }
}

/**
 * Performs the actual token refresh against the Entra token endpoint.
 *
 * @param {string} userId
 * @param {string} refreshToken
 * @param {PerUserTokenStorage} storage
 * @returns {Promise<string>} - New access token
 */
async function _doRefresh(userId, refreshToken, storage) {
  const tokenEndpoint = config.AUTH_CONFIG.tokenEndpoint;

  const body = new URLSearchParams({
    client_id: config.AUTH_CONFIG.clientId,
    client_secret: config.AUTH_CONFIG.clientSecret,
    grant_type: 'refresh_token',
    refresh_token: refreshToken,
    scope: config.AUTH_CONFIG.scopes.filter(s => s !== 'offline_access').join(' '),
  });

  const response = await fetch(tokenEndpoint, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: body.toString(),
  });

  if (!response.ok) {
    const err = new Error('Token refresh failed — re-authentication required');
    err.code = 'AUTH_REQUIRED';
    throw err;
  }

  const data = await response.json();

  // Store the new tokens
  const userInfo = storage.getUserInfo(userId);
  await storage.setTokensForUser(userId, {
    accessToken: data.access_token,
    refreshToken: data.refresh_token || refreshToken, // Entra may or may not rotate
    expiresIn: data.expires_in,
    scopes: data.scope || config.AUTH_CONFIG.scopes.join(' '),
    email: userInfo?.email || null,
    name: userInfo?.name || null,
  });

  return data.access_token;
}

module.exports = {
  tokenStorage,
  authTools,
  ensureAuthenticated,
  getHostedTokenStorage,
  setHostedTokenStorage,
};
