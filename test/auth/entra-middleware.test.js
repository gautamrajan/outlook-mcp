const crypto = require('crypto');
const { createEntraMiddleware, _resetKeyCache } = require('../../auth/entra-middleware');

// ── Test Key Infrastructure ─────────────────────────────────────────

const TEST_KID = 'test-key-id-001';
const UNKNOWN_KID = 'unknown-key-id-999';

const TEST_CONFIG = {
  tenantId: 'test-tenant-id-abc123',
  clientId: 'test-client-id-xyz789',
};

// Generate a real RSA key pair for signing/verifying in tests.
const { publicKey, privateKey } = crypto.generateKeyPairSync('rsa', {
  modulusLength: 2048,
  publicKeyEncoding: { type: 'spki', format: 'pem' },
  privateKeyEncoding: { type: 'pkcs8', format: 'pem' },
});

// Export the public key as JWK so we can build a mock JWKS response.
const publicKeyObject = crypto.createPublicKey(publicKey);
const jwk = publicKeyObject.export({ format: 'jwk' });

/**
 * Build a mock JWKS response containing the test public key.
 */
function mockJWKSResponse(kid = TEST_KID) {
  return {
    keys: [
      {
        kty: jwk.kty,
        use: 'sig',
        kid: kid,
        n: jwk.n,
        e: jwk.e,
        alg: 'RS256',
      },
    ],
  };
}

/**
 * Create a properly signed JWT with the given header overrides and payload.
 * Uses the test RSA private key for signing.
 */
function createSignedJWT(payload, headerOverrides = {}) {
  const header = {
    alg: 'RS256',
    typ: 'JWT',
    kid: TEST_KID,
    ...headerOverrides,
  };

  const headerB64 = Buffer.from(JSON.stringify(header)).toString('base64url');
  const payloadB64 = Buffer.from(JSON.stringify(payload)).toString('base64url');

  const signer = crypto.createSign('RSA-SHA256');
  signer.update(`${headerB64}.${payloadB64}`);
  const signature = signer.sign(privateKey);
  const signatureB64 = signature.toString('base64url');

  return `${headerB64}.${payloadB64}.${signatureB64}`;
}

/** Standard valid claims — override individual fields per test. */
function validClaims(overrides = {}) {
  return {
    tid: TEST_CONFIG.tenantId,
    aud: TEST_CONFIG.clientId,
    exp: Math.floor(Date.now() / 1000) + 3600, // 1 hour from now
    oid: 'user-object-id-001',
    preferred_username: 'gau@mrc.com',
    name: 'Gau Rajan',
    ...overrides,
  };
}

/**
 * Create a mock JWKS fetcher that returns the test JWKS response.
 * Tracks call count so tests can verify caching behaviour.
 */
function createMockFetcher(jwksResponse) {
  const fetcher = jest.fn().mockResolvedValue(jwksResponse || mockJWKSResponse());
  return fetcher;
}

function mockReqResNext(token) {
  const req = { headers: {} };
  if (token !== undefined) {
    req.headers.authorization = token === null ? undefined : token;
  }
  const res = { status: jest.fn().mockReturnThis(), json: jest.fn() };
  const next = jest.fn();
  return { req, res, next };
}

// ── Tests ────────────────────────────────────────────────────────────

describe('createEntraMiddleware', () => {
  let mockFetcher;

  beforeEach(() => {
    _resetKeyCache();
    mockFetcher = createMockFetcher();
  });

  function createMiddleware(configOverrides = {}) {
    return createEntraMiddleware({
      ...TEST_CONFIG,
      jwksFetcher: mockFetcher,
      ...configOverrides,
    });
  }

  // ── Existing claim validation tests (updated to use signed tokens) ──

  test('should extract user identity from a valid signed token and call next()', async () => {
    const claims = validClaims();
    const token = createSignedJWT(claims);
    const { req, res, next } = mockReqResNext(`Bearer ${token}`);
    const middleware = createMiddleware();

    await middleware(req, res, next);

    expect(next).toHaveBeenCalled();
    expect(res.status).not.toHaveBeenCalled();
    expect(req.user).toBeDefined();
    expect(req.user.id).toBe(claims.oid);
    expect(req.user.email).toBe(claims.preferred_username);
    expect(req.user.name).toBe(claims.name);
    expect(req.user.token).toBe(token);
  });

  test('should return 401 when Authorization header is missing', async () => {
    const { req, res, next } = mockReqResNext(undefined);
    delete req.headers.authorization;
    const middleware = createMiddleware();

    await middleware(req, res, next);

    expect(next).not.toHaveBeenCalled();
    expect(res.status).toHaveBeenCalledWith(401);
    expect(res.json).toHaveBeenCalledWith({
      error: 'Missing or invalid authorization header',
    });
  });

  test('should return 401 when Authorization header has no Bearer prefix', async () => {
    const token = createSignedJWT(validClaims());
    const { req, res, next } = mockReqResNext(`Basic ${token}`);
    const middleware = createMiddleware();

    await middleware(req, res, next);

    expect(next).not.toHaveBeenCalled();
    expect(res.status).toHaveBeenCalledWith(401);
    expect(res.json).toHaveBeenCalledWith({
      error: 'Missing or invalid authorization header',
    });
  });

  test('should return 401 when token is empty after Bearer prefix', async () => {
    const { req, res, next } = mockReqResNext('Bearer ');
    const middleware = createMiddleware();

    await middleware(req, res, next);

    expect(next).not.toHaveBeenCalled();
    expect(res.status).toHaveBeenCalledWith(401);
    expect(res.json).toHaveBeenCalledWith({
      error: 'Missing or invalid authorization header',
    });
  });

  test('should return 401 when JWT does not have three dot-separated parts', async () => {
    const { req, res, next } = mockReqResNext('Bearer not.a-valid-jwt');
    const middleware = createMiddleware();

    await middleware(req, res, next);

    expect(next).not.toHaveBeenCalled();
    expect(res.status).toHaveBeenCalledWith(401);
    expect(res.json).toHaveBeenCalledWith({
      error: 'Invalid token format',
    });
  });

  test('should return 401 when tenant ID does not match', async () => {
    const token = createSignedJWT(validClaims({ tid: 'wrong-tenant-id' }));
    const { req, res, next } = mockReqResNext(`Bearer ${token}`);
    const middleware = createMiddleware();

    await middleware(req, res, next);

    expect(next).not.toHaveBeenCalled();
    expect(res.status).toHaveBeenCalledWith(401);
    expect(res.json).toHaveBeenCalledWith({
      error: 'Invalid token: wrong tenant',
    });
  });

  test('should return 401 when token is expired', async () => {
    const token = createSignedJWT(validClaims({ exp: Math.floor(Date.now() / 1000) - 60 }));
    const { req, res, next } = mockReqResNext(`Bearer ${token}`);
    const middleware = createMiddleware();

    await middleware(req, res, next);

    expect(next).not.toHaveBeenCalled();
    expect(res.status).toHaveBeenCalledWith(401);
    expect(res.json).toHaveBeenCalledWith({
      error: 'Invalid token: expired',
    });
  });

  test('should return 401 when audience does not match', async () => {
    const token = createSignedJWT(validClaims({ aud: 'wrong-audience-id' }));
    const { req, res, next } = mockReqResNext(`Bearer ${token}`);
    const middleware = createMiddleware();

    await middleware(req, res, next);

    expect(next).not.toHaveBeenCalled();
    expect(res.status).toHaveBeenCalledWith(401);
    expect(res.json).toHaveBeenCalledWith({
      error: 'Invalid token: wrong audience',
    });
  });

  test('should accept audience in api://{clientId} format', async () => {
    const token = createSignedJWT(validClaims({ aud: `api://${TEST_CONFIG.clientId}` }));
    const { req, res, next } = mockReqResNext(`Bearer ${token}`);
    const middleware = createMiddleware();

    await middleware(req, res, next);

    expect(next).toHaveBeenCalled();
    expect(res.status).not.toHaveBeenCalled();
    expect(req.user).toBeDefined();
  });

  test('should populate req.user with id, email, name, and token', async () => {
    const claims = validClaims({
      oid: 'specific-oid-123',
      preferred_username: 'test@example.com',
      name: 'Test User',
    });
    const token = createSignedJWT(claims);
    const { req, res, next } = mockReqResNext(`Bearer ${token}`);
    const middleware = createMiddleware();

    await middleware(req, res, next);

    expect(req.user).toEqual({
      id: 'specific-oid-123',
      email: 'test@example.com',
      name: 'Test User',
      token: token,
    });
  });

  test('should handle multiple sequential requests independently (stateless)', async () => {
    const claimsA = validClaims({ oid: 'user-A', preferred_username: 'a@mrc.com', name: 'User A' });
    const claimsB = validClaims({ oid: 'user-B', preferred_username: 'b@mrc.com', name: 'User B' });

    const tokenA = createSignedJWT(claimsA);
    const tokenB = createSignedJWT(claimsB);

    const middleware = createMiddleware();

    // First request
    const r1 = mockReqResNext(`Bearer ${tokenA}`);
    await middleware(r1.req, r1.res, r1.next);

    expect(r1.next).toHaveBeenCalled();
    expect(r1.req.user.id).toBe('user-A');
    expect(r1.req.user.email).toBe('a@mrc.com');

    // Second request
    const r2 = mockReqResNext(`Bearer ${tokenB}`);
    await middleware(r2.req, r2.res, r2.next);

    expect(r2.next).toHaveBeenCalled();
    expect(r2.req.user.id).toBe('user-B');
    expect(r2.req.user.email).toBe('b@mrc.com');

    // Ensure first request's user object was not mutated
    expect(r1.req.user.id).toBe('user-A');
  });

  // ── Signature verification tests ──────────────────────────────────

  describe('signature verification', () => {
    test('should reject a token with a tampered payload', async () => {
      const claims = validClaims();
      const token = createSignedJWT(claims);

      // Tamper with the payload (change oid) but keep the original signature.
      // The signature was computed over the original payload, so it won't match.
      const parts = token.split('.');
      const tamperedPayload = Buffer.from(JSON.stringify({ ...claims, oid: 'tampered-oid' })).toString('base64url');
      const tamperedToken = `${parts[0]}.${tamperedPayload}.${parts[2]}`;

      const { req, res, next } = mockReqResNext(`Bearer ${tamperedToken}`);
      const middleware = createMiddleware();

      await middleware(req, res, next);

      expect(next).not.toHaveBeenCalled();
      expect(res.status).toHaveBeenCalledWith(401);
      expect(res.json).toHaveBeenCalledWith({
        error: 'Invalid token: signature verification failed',
      });
    });

    test('should reject a token with a completely fake signature', async () => {
      const header = Buffer.from(JSON.stringify({ alg: 'RS256', typ: 'JWT', kid: TEST_KID })).toString('base64url');
      const payload = Buffer.from(JSON.stringify(validClaims())).toString('base64url');
      const fakeToken = `${header}.${payload}.fakesignature`;

      const { req, res, next } = mockReqResNext(`Bearer ${fakeToken}`);
      const middleware = createMiddleware();

      await middleware(req, res, next);

      expect(next).not.toHaveBeenCalled();
      expect(res.status).toHaveBeenCalledWith(401);
      expect(res.json).toHaveBeenCalledWith({
        error: 'Invalid token: signature verification failed',
      });
    });

    test('should reject a token signed with a different key', async () => {
      // Generate a second, different key pair
      const { privateKey: otherPrivate } = crypto.generateKeyPairSync('rsa', {
        modulusLength: 2048,
        publicKeyEncoding: { type: 'spki', format: 'pem' },
        privateKeyEncoding: { type: 'pkcs8', format: 'pem' },
      });

      const claims = validClaims();
      const header = {
        alg: 'RS256',
        typ: 'JWT',
        kid: TEST_KID, // same kid, but signed with wrong key
      };

      const headerB64 = Buffer.from(JSON.stringify(header)).toString('base64url');
      const payloadB64 = Buffer.from(JSON.stringify(claims)).toString('base64url');

      const signer = crypto.createSign('RSA-SHA256');
      signer.update(`${headerB64}.${payloadB64}`);
      const signature = signer.sign(otherPrivate);
      const signatureB64 = signature.toString('base64url');

      const wrongKeyToken = `${headerB64}.${payloadB64}.${signatureB64}`;

      const { req, res, next } = mockReqResNext(`Bearer ${wrongKeyToken}`);
      const middleware = createMiddleware();

      await middleware(req, res, next);

      expect(next).not.toHaveBeenCalled();
      expect(res.status).toHaveBeenCalledWith(401);
      expect(res.json).toHaveBeenCalledWith({
        error: 'Invalid token: signature verification failed',
      });
    });
  });

  // ── Key ID (kid) handling tests ───────────────────────────────────

  describe('kid handling', () => {
    test('should reject token with unknown kid after refreshing cache', async () => {
      const claims = validClaims();
      const token = createSignedJWT(claims, { kid: UNKNOWN_KID });

      const { req, res, next } = mockReqResNext(`Bearer ${token}`);
      const middleware = createMiddleware();

      await middleware(req, res, next);

      expect(next).not.toHaveBeenCalled();
      expect(res.status).toHaveBeenCalledWith(401);
      expect(res.json).toHaveBeenCalledWith({
        error: 'Invalid token: signing key not found',
      });

      // Should have fetched twice: initial fetch + refresh attempt
      expect(mockFetcher).toHaveBeenCalledTimes(2);
    });

    test('should find key after refresh when kid is added to JWKS', async () => {
      const newKid = 'rotated-key-id-002';
      const claims = validClaims();
      const token = createSignedJWT(claims, { kid: newKid });

      // First call returns JWKS without the new kid, second call includes it
      const jwksWithoutNewKey = mockJWKSResponse(TEST_KID);
      const jwksWithNewKey = {
        keys: [
          ...mockJWKSResponse(TEST_KID).keys,
          { ...mockJWKSResponse(newKid).keys[0] },
        ],
      };

      mockFetcher
        .mockResolvedValueOnce(jwksWithoutNewKey)
        .mockResolvedValueOnce(jwksWithNewKey);

      const { req, res, next } = mockReqResNext(`Bearer ${token}`);
      const middleware = createMiddleware();

      await middleware(req, res, next);

      expect(next).toHaveBeenCalled();
      expect(res.status).not.toHaveBeenCalled();
      expect(mockFetcher).toHaveBeenCalledTimes(2);
    });

    test('should reject token with missing kid in header', async () => {
      const claims = validClaims();
      // Create a token with no kid in the header
      const header = { alg: 'RS256', typ: 'JWT' }; // no kid
      const headerB64 = Buffer.from(JSON.stringify(header)).toString('base64url');
      const payloadB64 = Buffer.from(JSON.stringify(claims)).toString('base64url');

      const signer = crypto.createSign('RSA-SHA256');
      signer.update(`${headerB64}.${payloadB64}`);
      const signature = signer.sign(privateKey);
      const signatureB64 = signature.toString('base64url');

      const token = `${headerB64}.${payloadB64}.${signatureB64}`;

      const { req, res, next } = mockReqResNext(`Bearer ${token}`);
      const middleware = createMiddleware();

      await middleware(req, res, next);

      expect(next).not.toHaveBeenCalled();
      expect(res.status).toHaveBeenCalledWith(401);
      expect(res.json).toHaveBeenCalledWith({
        error: 'Invalid token: missing key ID',
      });
    });
  });

  // ── JWKS caching tests ────────────────────────────────────────────

  describe('JWKS caching', () => {
    test('should cache JWKS keys and not refetch on second request', async () => {
      const middleware = createMiddleware();

      // First request — triggers JWKS fetch
      const token1 = createSignedJWT(validClaims());
      const r1 = mockReqResNext(`Bearer ${token1}`);
      await middleware(r1.req, r1.res, r1.next);
      expect(r1.next).toHaveBeenCalled();

      // Second request — should use cached keys
      const token2 = createSignedJWT(validClaims({ oid: 'user-2' }));
      const r2 = mockReqResNext(`Bearer ${token2}`);
      await middleware(r2.req, r2.res, r2.next);
      expect(r2.next).toHaveBeenCalled();

      // Only one fetch call — cache was used
      expect(mockFetcher).toHaveBeenCalledTimes(1);
    });

    test('should refetch JWKS after cache expires (24 hours)', async () => {
      const middleware = createMiddleware();

      // First request — triggers JWKS fetch
      const token1 = createSignedJWT(validClaims());
      const r1 = mockReqResNext(`Bearer ${token1}`);
      await middleware(r1.req, r1.res, r1.next);
      expect(r1.next).toHaveBeenCalled();
      expect(mockFetcher).toHaveBeenCalledTimes(1);

      // Advance time by 25 hours
      const realNow = Date.now;
      Date.now = () => realNow() + 25 * 60 * 60 * 1000;

      try {
        // Second request — cache should be stale, triggers refetch
        const token2 = createSignedJWT(validClaims({ oid: 'user-2' }));
        const r2 = mockReqResNext(`Bearer ${token2}`);
        await middleware(r2.req, r2.res, r2.next);
        expect(r2.next).toHaveBeenCalled();

        // Two fetches: initial + refresh after expiry
        expect(mockFetcher).toHaveBeenCalledTimes(2);
      } finally {
        Date.now = realNow;
      }
    });

    test('should not refetch JWKS before cache expires', async () => {
      const middleware = createMiddleware();

      // First request
      const token1 = createSignedJWT(validClaims());
      const r1 = mockReqResNext(`Bearer ${token1}`);
      await middleware(r1.req, r1.res, r1.next);
      expect(mockFetcher).toHaveBeenCalledTimes(1);

      // Advance time by 23 hours (still within 24-hour window)
      const realNow = Date.now;
      Date.now = () => realNow() + 23 * 60 * 60 * 1000;

      try {
        const token2 = createSignedJWT(validClaims({ oid: 'user-2' }));
        const r2 = mockReqResNext(`Bearer ${token2}`);
        await middleware(r2.req, r2.res, r2.next);
        expect(r2.next).toHaveBeenCalled();

        // Still only one fetch — cache hasn't expired
        expect(mockFetcher).toHaveBeenCalledTimes(1);
      } finally {
        Date.now = realNow;
      }
    });
  });

  // ── JWKS fetch error handling ─────────────────────────────────────

  describe('JWKS fetch errors', () => {
    test('should return 401 when JWKS fetch fails', async () => {
      const failingFetcher = jest.fn().mockRejectedValue(new Error('Network error'));
      const middleware = createMiddleware({ jwksFetcher: failingFetcher });

      const token = createSignedJWT(validClaims());
      const { req, res, next } = mockReqResNext(`Bearer ${token}`);

      await middleware(req, res, next);

      expect(next).not.toHaveBeenCalled();
      expect(res.status).toHaveBeenCalledWith(401);
      expect(res.json).toHaveBeenCalledWith({
        error: 'Invalid token: unable to verify signature',
      });
    });
  });

  // ── _resetKeyCache tests ──────────────────────────────────────────

  describe('_resetKeyCache', () => {
    test('should clear cached keys so next request refetches', async () => {
      const middleware = createMiddleware();

      // First request — triggers fetch
      const token1 = createSignedJWT(validClaims());
      const r1 = mockReqResNext(`Bearer ${token1}`);
      await middleware(r1.req, r1.res, r1.next);
      expect(mockFetcher).toHaveBeenCalledTimes(1);

      // Reset cache
      _resetKeyCache();

      // Second request — should fetch again because cache was cleared
      const token2 = createSignedJWT(validClaims({ oid: 'user-2' }));
      const r2 = mockReqResNext(`Bearer ${token2}`);
      await middleware(r2.req, r2.res, r2.next);
      expect(r2.next).toHaveBeenCalled();
      expect(mockFetcher).toHaveBeenCalledTimes(2);
    });
  });
});
