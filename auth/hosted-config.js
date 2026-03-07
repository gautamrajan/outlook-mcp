/**
 * Helpers for hosted-mode configuration and canonical public URLs.
 */

function tryGetOrigin(candidate) {
  if (!candidate) return null;
  try {
    return new URL(candidate).origin;
  } catch {
    return null;
  }
}

/**
 * Returns the configured externally reachable base URL for hosted mode.
 *
 * Preference order:
 *   1. PUBLIC_BASE_URL
 *   2. HOSTED_REDIRECT_URI origin
 *
 * @param {object} appConfig
 * @returns {string|null}
 */
function getConfiguredServerBaseUrl(appConfig) {
  const candidates = [
    appConfig?.HOSTED?.publicBaseUrl,
    appConfig?.AUTH_CONFIG?.hostedRedirectUri,
  ];

  for (const candidate of candidates) {
    const origin = tryGetOrigin(candidate);
    if (origin) {
      return origin;
    }
  }

  return null;
}

/**
 * Returns a best-effort request-derived base URL for local/dev fallback.
 *
 * This intentionally ignores X-Forwarded-* headers. Hosted deployments should
 * set a canonical configured base URL instead of trusting client-supplied
 * forwarding headers.
 *
 * @param {import('express').Request} req
 * @returns {string|null}
 */
function getRequestBaseUrl(req) {
  if (!req || typeof req.get !== 'function') {
    return null;
  }

  const host = req.get('host');
  if (!host) {
    return null;
  }

  return `${req.protocol || 'http'}://${host}`;
}

/**
 * Returns the preferred server base URL for hosted metadata and discovery.
 *
 * @param {object} appConfig
 * @param {import('express').Request} [req]
 * @returns {string|null}
 */
function getServerBaseUrl(appConfig, req) {
  return getConfiguredServerBaseUrl(appConfig) || getRequestBaseUrl(req);
}

/**
 * Returns a list of missing/invalid hosted connector settings.
 *
 * @param {object} appConfig
 * @returns {string[]}
 */
function getHostedConnectorConfigErrors(appConfig) {
  const authConfig = appConfig?.AUTH_CONFIG || {};
  const connectorConfig = appConfig?.CONNECTOR_AUTH || {};
  const errors = [];

  if (!authConfig.clientId) {
    errors.push('OUTLOOK_CLIENT_ID / MS_CLIENT_ID');
  }
  if (!authConfig.clientSecret) {
    errors.push('OUTLOOK_CLIENT_SECRET / MS_CLIENT_SECRET');
  }
  if (!authConfig.tenantId || authConfig.tenantId === 'common') {
    errors.push('OUTLOOK_TENANT_ID / MS_TENANT_ID (must be a specific tenant ID)');
  }
  if (!connectorConfig.apiAppId) {
    errors.push('MCP_API_APP_ID');
  }
  if (!connectorConfig.apiScope) {
    errors.push('MCP_API_SCOPE');
  }
  if (!connectorConfig.oboScopes) {
    errors.push('OBO_SCOPES');
  }

  return errors;
}

module.exports = {
  getConfiguredServerBaseUrl,
  getRequestBaseUrl,
  getServerBaseUrl,
  getHostedConnectorConfigErrors,
};
