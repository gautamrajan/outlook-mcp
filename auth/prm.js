/**
 * Protected Resource Metadata (RFC 9728) for Claude custom connectors.
 *
 * Exposes a well-known endpoint that tells clients how to authenticate:
 *   GET /.well-known/oauth-protected-resource
 *
 * Also exports a helper to build the WWW-Authenticate challenge header
 * value used by auth middleware when returning 401 responses.
 */

const config = require('../config');
const { getServerBaseUrl, getRequestBaseUrl, getConfiguredServerBaseUrl } = require('./hosted-config');

const { tenantId } = config.AUTH_CONFIG;
const { apiAppId, apiScope } = config.CONNECTOR_AUTH;

/**
 * Build the full scope string, e.g. "api://xxx/mcp.access".
 */
function getFullScope() {
  return `${apiAppId}/${apiScope}`;
}

/**
 * Returns the externally reachable base URL used for PRM and auth challenges.
 * Hosted deployments should provide a canonical configured URL; otherwise we
 * fall back to the request host for local development.
 *
 * @param {import('express').Request} req
 * @returns {string|null}
 */
function getBaseUrl(req) {
  return getServerBaseUrl(config, req);
}

/**
 * Express handler for GET /.well-known/oauth-protected-resource.
 *
 * Returns the Protected Resource Metadata JSON document per RFC 9728.
 */
function prmHandler(req, res) {
  const baseUrl = getBaseUrl(req);
  if (!baseUrl) {
    return res.status(500).json({
      error: 'server_base_url_unavailable',
      message: 'Server base URL is not configured',
    });
  }

  res.set('Content-Type', 'application/json');
  res.json({
    resource: apiAppId,
    resource_name: 'MRC Outlook Assistant',
    authorization_servers: [
      `https://login.microsoftonline.com/${tenantId}/v2.0`,
    ],
    scopes_supported: [
      getFullScope(),
    ],
    bearer_methods_supported: ['header'],
  });
}

/**
 * Build the WWW-Authenticate header value for 401 responses.
 *
 * Format:
 *   Bearer realm="mcp", resource_metadata="<url>/.well-known/oauth-protected-resource", scope="<full_scope>"
 *
 * @param {import('express').Request} req  Used to derive the server base URL.
 * @returns {string}
 */
function buildWwwAuthenticateChallenge(req) {
  const baseUrl = getBaseUrl(req);
  const fullScope = getFullScope();
  if (!baseUrl) {
    throw new Error('Server base URL is not configured');
  }
  return `Bearer realm="mcp", resource_metadata="${baseUrl}/.well-known/oauth-protected-resource", scope="${fullScope}"`;
}

module.exports = {
  prmHandler,
  buildWwwAuthenticateChallenge,
  // Exported for testing
  getBaseUrl,
  getConfiguredServerBaseUrl: () => getConfiguredServerBaseUrl(config),
  getRequestBaseUrl,
  getFullScope,
};
