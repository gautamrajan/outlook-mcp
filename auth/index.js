/**
 * Authentication module for Outlook MCP server
 */
const TokenStorage = require('./token-storage');
const config = require('../config');
const { authTools } = require('./tools');

// Singleton TokenStorage instance with config from the main config file
const tokenStorage = new TokenStorage({
  tokenStorePath: config.AUTH_CONFIG.tokenStorePath,
  clientId: config.AUTH_CONFIG.clientId,
  clientSecret: config.AUTH_CONFIG.clientSecret,
  tokenEndpoint: config.AUTH_CONFIG.tokenEndpoint,
  redirectUri: config.AUTH_CONFIG.redirectUri,
  scopes: config.AUTH_CONFIG.scopes
});

/**
 * Ensures the user is authenticated and returns an access token.
 * @param {{ forceRefresh?: boolean }} options
 * @returns {Promise<string>} - Access token
 * @throws {Error} - If authentication fails
 */
async function ensureAuthenticated({ forceRefresh = false } = {}) {
  if (forceRefresh) {
    tokenStorage.invalidateAccessToken();
  }

  const accessToken = await tokenStorage.getValidAccessToken();
  if (!accessToken) {
    throw new Error('Authentication required');
  }

  return accessToken;
}

module.exports = {
  tokenStorage,
  authTools,
  ensureAuthenticated
};
