/**
 * On-Behalf-Of (OBO) token exchange for Microsoft Graph API.
 *
 * Exchanges a validated Entra access token (audience = our API) for a
 * Microsoft Graph access token via the OAuth 2.0 OBO flow.
 *
 * Includes an in-memory cache keyed by user OID to avoid redundant exchanges.
 */

const config = require('../config');

// In-memory token cache: Map<oid, { accessToken: string, expiresAt: number }>
const tokenCache = new Map();

// Buffer (in ms) before actual expiry at which we consider a token stale
const EXPIRY_BUFFER_MS = 5 * 60 * 1000; // 5 minutes

/**
 * Exchanges an Entra JWT for a Microsoft Graph access token via the OBO flow.
 * @param {string} entraToken - The incoming Entra access token (raw JWT string)
 * @returns {Promise<{ access_token: string, expires_in: number }>} The OBO token response
 * @throws {Error} If the token exchange fails
 */
async function exchangeOboToken(entraToken) {
  const { clientId, clientSecret, tenantId } = config.AUTH_CONFIG;
  const { oboScopes } = config.CONNECTOR_AUTH;

  const tokenEndpoint = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;

  const body = new URLSearchParams({
    grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
    client_id: clientId,
    client_secret: clientSecret,
    assertion: entraToken,
    scope: oboScopes,
    requested_token_use: 'on_behalf_of',
  });

  const response = await fetch(tokenEndpoint, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: body.toString(),
  });

  const data = await response.json();

  if (!response.ok) {
    const errorCode = data.error || 'unknown_error';
    const errorDescription = data.error_description || 'No error description provided';
    throw new Error(`OBO token exchange failed: [${errorCode}] ${errorDescription}`);
  }

  return data;
}

/**
 * Returns a Microsoft Graph access token for the given user, using the cache
 * when possible and falling back to an OBO exchange.
 *
 * @param {string} entraToken - The incoming Entra access token (raw JWT string)
 * @param {string} userOid - The user's object ID (oid claim from the Entra JWT)
 * @returns {Promise<string>} A Microsoft Graph access token
 * @throws {Error} If the OBO exchange fails
 */
async function getGraphToken(entraToken, userOid) {
  // Check cache
  const cached = tokenCache.get(userOid);
  if (cached && Date.now() < cached.expiresAt - EXPIRY_BUFFER_MS) {
    return cached.accessToken;
  }

  // Cache miss or expired — perform OBO exchange
  const data = await exchangeOboToken(entraToken);

  // Cache the new token
  const expiresAt = Date.now() + (data.expires_in * 1000);
  tokenCache.set(userOid, {
    accessToken: data.access_token,
    expiresAt,
  });

  return data.access_token;
}

/**
 * Clears the entire OBO token cache.
 */
function clearTokenCache() {
  tokenCache.clear();
}

/**
 * Clears the cached OBO token for a specific user.
 * @param {string} oid - The user's object ID
 */
function clearUserToken(oid) {
  tokenCache.delete(oid);
}

module.exports = {
  getGraphToken,
  clearTokenCache,
  clearUserToken,
};
