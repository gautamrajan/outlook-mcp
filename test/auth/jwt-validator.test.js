const crypto = require('crypto');
const jwt = require('jsonwebtoken');

// Generate an RSA key pair for signing test tokens
const { publicKey, privateKey } = crypto.generateKeyPairSync('rsa', {
  modulusLength: 2048,
  publicKeyEncoding: { type: 'spki', format: 'pem' },
  privateKeyEncoding: { type: 'pkcs8', format: 'pem' },
});

// Mock config values
jest.mock('../../config', () => ({
  AUTH_CONFIG: {
    tenantId: 'test-tenant-id',
  },
  CONNECTOR_AUTH: {
    apiAppId: 'api://test-app-id',
    apiScope: 'mcp.access',
  },
}));

// Mock jwks-rsa — the factory must not reference out-of-scope variables,
// so we use a mockGetSigningKey function that we configure in beforeAll.
const mockGetSigningKey = jest.fn();
jest.mock('jwks-rsa', () => {
  return jest.fn(() => ({
    getSigningKey: mockGetSigningKey,
  }));
});

// Configure the mock to return our test public key
beforeAll(() => {
  mockGetSigningKey.mockImplementation((_kid, callback) => {
    callback(null, { getPublicKey: () => publicKey });
  });
});

const { validateEntraJwt } = require('../../auth/jwt-validator');

/**
 * Helper: creates a signed JWT with the given payload overrides.
 */
function createTestToken(overrides = {}) {
  const jwtOptions = overrides._jwtOptions || {};
  const payload = { ...overrides };
  delete payload._jwtOptions;

  const defaults = {
    oid: 'user-oid-123',
    sub: 'user-sub-456',
    preferred_username: 'testuser@example.com',
    name: 'Test User',
    scp: 'mcp.access',
  };

  const finalPayload = { ...defaults, ...payload };

  const options = {
    algorithm: 'RS256',
    issuer: 'https://login.microsoftonline.com/test-tenant-id/v2.0',
    audience: 'api://test-app-id',
    expiresIn: '1h',
    keyid: 'test-kid',
    ...jwtOptions,
  };

  return jwt.sign(finalPayload, privateKey, options);
}

describe('validateEntraJwt', () => {
  test('should accept a valid token with correct claims', async () => {
    const token = createTestToken();
    const claims = await validateEntraJwt(token);

    expect(claims.oid).toBe('user-oid-123');
    expect(claims.sub).toBe('user-sub-456');
    expect(claims.preferred_username).toBe('testuser@example.com');
    expect(claims.name).toBe('Test User');
    expect(claims.scp).toBe('mcp.access');
  });

  test('should reject a token with invalid signature', async () => {
    const token = createTestToken();
    // Tamper with the signature portion (last segment)
    const parts = token.split('.');
    parts[2] = parts[2].split('').reverse().join('');
    const tampered = parts.join('.');

    await expect(validateEntraJwt(tampered)).rejects.toThrow('JWT validation failed');
  });

  test('should reject an expired token', async () => {
    const token = createTestToken({
      _jwtOptions: { expiresIn: '-10s' },
    });

    await expect(validateEntraJwt(token)).rejects.toThrow('JWT validation failed');
  });

  test('should reject a token with wrong audience', async () => {
    const token = createTestToken({
      _jwtOptions: { audience: 'api://wrong-app-id' },
    });

    await expect(validateEntraJwt(token)).rejects.toThrow('JWT validation failed');
  });

  test('should reject a token with wrong issuer', async () => {
    const token = createTestToken({
      _jwtOptions: { issuer: 'https://login.microsoftonline.com/wrong-tenant/v2.0' },
    });

    await expect(validateEntraJwt(token)).rejects.toThrow('JWT validation failed');
  });

  test('should reject a token missing the required scope', async () => {
    const token = createTestToken({ scp: 'some.other.scope' });

    await expect(validateEntraJwt(token)).rejects.toThrow("JWT missing required scope 'mcp.access'");
  });

  test('should reject a token with no scp claim at all', async () => {
    // Build a token without scp by explicitly setting it to undefined
    const payload = {
      oid: 'user-oid-123',
      sub: 'user-sub-456',
      preferred_username: 'testuser@example.com',
      name: 'Test User',
    };

    const token = jwt.sign(payload, privateKey, {
      algorithm: 'RS256',
      issuer: 'https://login.microsoftonline.com/test-tenant-id/v2.0',
      audience: 'api://test-app-id',
      expiresIn: '1h',
      keyid: 'test-kid',
    });

    await expect(validateEntraJwt(token)).rejects.toThrow("JWT missing required scope 'mcp.access'");
  });

  test('should reject a token with no oid claim', async () => {
    const payload = {
      sub: 'user-sub-456',
      preferred_username: 'testuser@example.com',
      name: 'Test User',
      scp: 'mcp.access',
    };

    const token = jwt.sign(payload, privateKey, {
      algorithm: 'RS256',
      issuer: 'https://login.microsoftonline.com/test-tenant-id/v2.0',
      audience: 'api://test-app-id',
      expiresIn: '1h',
      keyid: 'test-kid',
    });

    await expect(validateEntraJwt(token)).rejects.toThrow('JWT missing required oid claim');
  });

  test('should accept a token that has the required scope among multiple scopes', async () => {
    const token = createTestToken({ scp: 'other.scope mcp.access another.scope' });
    const claims = await validateEntraJwt(token);

    expect(claims.scp).toBe('other.scope mcp.access another.scope');
  });
});
