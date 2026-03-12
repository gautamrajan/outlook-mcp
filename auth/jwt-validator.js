/**
 * Entra JWT Validation
 *
 * Validates incoming Bearer tokens as Entra-issued JWTs using JWKS
 * key discovery and standard claim checks (issuer, audience, expiry, scope).
 */
const jwt = require('jsonwebtoken');
const jwksRsa = require('jwks-rsa');
const config = require('../config');

const tenantId = config.AUTH_CONFIG.tenantId;

// Singleton JWKS client — caches keys and rate-limits requests
const jwksClient = jwksRsa({
  jwksUri: `https://login.microsoftonline.com/${tenantId}/discovery/v2.0/keys`,
  cache: true,
  rateLimit: true,
});

/**
 * Key-retrieval callback for jsonwebtoken.verify().
 * Looks up the signing key by the token's `kid` header.
 *
 * @param {object} header - JWT header (contains kid, alg, etc.)
 * @param {function} callback - Node-style callback(err, signingKey)
 */
function getSigningKey(header, callback) {
  jwksClient.getSigningKey(header.kid, (err, key) => {
    if (err) return callback(err);
    callback(null, key.getPublicKey());
  });
}

/**
 * Validates an Entra-issued JWT.
 *
 * Checks:
 *   - Algorithm: RS256
 *   - Issuer: https://login.microsoftonline.com/<tenantId>/v2.0
 *   - Audience: CONNECTOR_AUTH.apiAppId from config
 *   - Expiry: token must not be expired
 *   - Scope: `scp` claim must include CONNECTOR_AUTH.apiScope
 *   - Identity: token must include a usable `oid` claim
 *
 * @param {string} token - Raw JWT string (no "Bearer " prefix)
 * @returns {Promise<object>} - Decoded token claims
 * @throws {Error} - Descriptive error if validation fails
 */
function validateEntraJwt(token) {
  return new Promise((resolve, reject) => {
    const options = {
      algorithms: ['RS256'],
      issuer: `https://login.microsoftonline.com/${tenantId}/v2.0`,
      // Accept both "api://xxx" (identifier URI) and "xxx" (bare app ID) as audience,
      // because Entra v2.0 tokens may use either depending on how the scope was requested.
      audience: [config.CONNECTOR_AUTH.apiAppId, config.AUTH_CONFIG.clientId],
    };

    jwt.verify(token, getSigningKey, options, (err, decoded) => {
      if (err) {
        return reject(new Error(`JWT validation failed: ${err.message}`));
      }

      // Verify required scope
      const requiredScope = config.CONNECTOR_AUTH.apiScope;
      const tokenScopes = decoded.scp ? decoded.scp.split(' ') : [];

      if (!tokenScopes.includes(requiredScope)) {
        return reject(
          new Error(`JWT missing required scope '${requiredScope}'. Token scopes: ${tokenScopes.join(', ') || '(none)'}`)
        );
      }

      if (typeof decoded.oid !== 'string' || decoded.oid.length === 0) {
        return reject(new Error('JWT missing required oid claim'));
      }

      resolve(decoded);
    });
  });
}

module.exports = {
  validateEntraJwt,
};
