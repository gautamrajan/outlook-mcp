/**
 * Entra ID (Azure AD) JWT validation middleware for Express.
 *
 * Validates bearer tokens by:
 * 1. Verifying the cryptographic signature against Entra's JWKS keys
 * 2. Checking claims (tenant, audience, expiry)
 *
 * Usage:
 *   const { createEntraMiddleware } = require('./auth/entra-middleware');
 *   app.use(createEntraMiddleware({ tenantId, clientId }));
 */

const crypto = require('crypto');
const https = require('https');

// ── JWKS Key Cache ──────────────────────────────────────────────────

const CACHE_TTL_MS = 24 * 60 * 60 * 1000; // 24 hours

let cachedKeys = null;    // Array of { kid, pem } objects
let cacheTimestamp = 0;   // ms timestamp of last fetch

/**
 * Reset the key cache. Exported for testing only.
 */
function _resetKeyCache() {
  cachedKeys = null;
  cacheTimestamp = 0;
}

// ── Default JWKS Fetcher ────────────────────────────────────────────

/**
 * Fetch JWKS from Entra's discovery endpoint.
 *
 * @param {string} tenantId - Azure AD tenant ID
 * @returns {Promise<Object>} Parsed JWKS response with { keys: [...] }
 */
function defaultJWKSFetcher(tenantId) {
  const url = `https://login.microsoftonline.com/${tenantId}/discovery/v2.0/keys`;

  return new Promise((resolve, reject) => {
    https.get(url, (res) => {
      let data = '';
      res.on('data', (chunk) => { data += chunk; });
      res.on('end', () => {
        try {
          resolve(JSON.parse(data));
        } catch (err) {
          reject(new Error(`Failed to parse JWKS response: ${err.message}`));
        }
      });
      res.on('error', reject);
    }).on('error', reject);
  });
}

// ── Key Parsing ─────────────────────────────────────────────────────

/**
 * Parse a JWKS response into an array of { kid, pem } objects.
 *
 * @param {Object} jwksResponse - Parsed JSON from the JWKS endpoint
 * @returns {Array<{kid: string, pem: string}>}
 */
function parseJWKSKeys(jwksResponse) {
  if (!jwksResponse || !Array.isArray(jwksResponse.keys)) {
    return [];
  }

  return jwksResponse.keys
    .filter((key) => key.kty === 'RSA' && key.n && key.e)
    .map((key) => {
      const publicKey = crypto.createPublicKey({ key, format: 'jwk' });
      const pem = publicKey.export({ type: 'spki', format: 'pem' });
      return { kid: key.kid, pem };
    });
}

// ── Middleware Factory ───────────────────────────────────────────────

/**
 * Create Express middleware that validates Entra ID bearer tokens.
 *
 * @param {Object} config
 * @param {string} config.tenantId - Expected Azure AD tenant ID (tid claim).
 * @param {string} config.clientId - App registration client ID (aud claim).
 * @param {Function} [config.jwksFetcher] - Optional custom JWKS fetcher for testing.
 *   Receives (tenantId) and returns Promise<{ keys: [...] }>.
 * @returns {Function} Express async middleware (req, res, next)
 */
function createEntraMiddleware(config) {
  const { tenantId, clientId, jwksFetcher } = config;
  const fetchJWKS = jwksFetcher || ((tid) => defaultJWKSFetcher(tid));

  /**
   * Fetch and cache JWKS keys. Returns the cached keys if still fresh.
   * @param {boolean} forceRefresh - Bypass cache TTL and refetch
   * @returns {Promise<Array<{kid: string, pem: string}>>}
   */
  async function getKeys(forceRefresh = false) {
    const now = Date.now();
    const cacheExpired = (now - cacheTimestamp) >= CACHE_TTL_MS;

    if (cachedKeys && !forceRefresh && !cacheExpired) {
      return cachedKeys;
    }

    const jwksResponse = await fetchJWKS(tenantId);
    cachedKeys = parseJWKSKeys(jwksResponse);
    cacheTimestamp = Date.now();
    return cachedKeys;
  }

  /**
   * Look up a signing key by kid. If not found, refreshes the cache once
   * (keys may have rotated) and tries again.
   *
   * @param {string} kid - Key ID from JWT header
   * @returns {Promise<{kid: string, pem: string}|null>}
   */
  async function findKey(kid) {
    let keys = await getKeys();
    let key = keys.find((k) => k.kid === kid);
    if (key) return key;

    // Key not found — try refreshing (rotation may have occurred)
    keys = await getKeys(true);
    key = keys.find((k) => k.kid === kid);
    return key || null;
  }

  return async function entraMiddleware(req, res, next) {
    // ── 1. Extract token from Authorization header ───────────────
    const authHeader = req.headers.authorization;
    if (!authHeader || !authHeader.startsWith('Bearer ')) {
      return res.status(401).json({ error: 'Missing or invalid authorization header' });
    }

    const rawToken = authHeader.slice(7); // strip "Bearer "
    if (!rawToken) {
      return res.status(401).json({ error: 'Missing or invalid authorization header' });
    }

    // ── 2. Split and validate JWT structure ──────────────────────
    const parts = rawToken.split('.');
    if (parts.length !== 3) {
      return res.status(401).json({ error: 'Invalid token format' });
    }

    const [headerB64, payloadB64, signatureB64] = parts;

    // ── 3. Decode header to get kid ──────────────────────────────
    let header;
    try {
      header = JSON.parse(Buffer.from(headerB64, 'base64url').toString('utf8'));
    } catch {
      return res.status(401).json({ error: 'Invalid token format' });
    }

    if (!header.kid) {
      return res.status(401).json({ error: 'Invalid token: missing key ID' });
    }

    // ── 4. Fetch signing key and verify signature ────────────────
    let key;
    try {
      key = await findKey(header.kid);
    } catch {
      return res.status(401).json({ error: 'Invalid token: unable to verify signature' });
    }

    if (!key) {
      return res.status(401).json({ error: 'Invalid token: signing key not found' });
    }

    try {
      const verifier = crypto.createVerify('RSA-SHA256');
      verifier.update(`${headerB64}.${payloadB64}`);
      const isValid = verifier.verify(key.pem, Buffer.from(signatureB64, 'base64url'));

      if (!isValid) {
        return res.status(401).json({ error: 'Invalid token: signature verification failed' });
      }
    } catch {
      return res.status(401).json({ error: 'Invalid token: signature verification failed' });
    }

    // ── 5. Decode payload and validate claims ────────────────────
    let claims;
    try {
      claims = JSON.parse(Buffer.from(payloadB64, 'base64url').toString('utf8'));
    } catch {
      return res.status(401).json({ error: 'Invalid token format' });
    }

    // Tenant
    if (claims.tid !== tenantId) {
      return res.status(401).json({ error: 'Invalid token: wrong tenant' });
    }

    // Expiry
    const nowSeconds = Math.floor(Date.now() / 1000);
    if (!claims.exp || claims.exp <= nowSeconds) {
      return res.status(401).json({ error: 'Invalid token: expired' });
    }

    // Audience — accept either the raw clientId or the api:// prefixed form
    if (claims.aud !== clientId && claims.aud !== `api://${clientId}`) {
      return res.status(401).json({ error: 'Invalid token: wrong audience' });
    }

    // ── 6. Attach user identity to request ───────────────────────
    req.user = {
      id: claims.oid,
      email: claims.preferred_username,
      name: claims.name,
      token: rawToken,
    };

    next();
  };
}

module.exports = { createEntraMiddleware, _resetKeyCache };
