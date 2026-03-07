/**
 * Integration tests for the connector (Entra JWT + OBO) auth flow.
 *
 * End-to-end test of the connector auth path through the Express app:
 *   JWT middleware validates Entra token → session middleware skips →
 *   MCP handler sets connector user context → OBO exchange is available.
 *
 * Strategy:
 *   - Generate a real RSA key pair for test JWT signing
 *   - Mock `jwks-rsa` to return the test public key
 *   - Mock `global.fetch` to intercept the OBO token exchange
 *   - Mock MCP SDK classes so we test the HTTP/auth wiring, not SDK internals
 *   - Use supertest against the Express app from createHttpApp()
 */

const crypto = require('crypto');
const jwt = require('jsonwebtoken');
const supertest = require('supertest');

// ── RSA key pair for signing test JWTs ────────────────────────────────

const { publicKey, privateKey } = crypto.generateKeyPairSync('rsa', {
  modulusLength: 2048,
  publicKeyEncoding: { type: 'spki', format: 'pem' },
  privateKeyEncoding: { type: 'pkcs8', format: 'pem' },
});

// ── Mocks (must come before any require of modules under test) ────────

// Mock jwks-rsa — must be mocked before jwt-validator requires it
const mockGetSigningKey = jest.fn((_kid, callback) => {
  callback(null, { getPublicKey: () => publicKey });
});
jest.mock('jwks-rsa', () => {
  return jest.fn(() => ({
    getSigningKey: mockGetSigningKey,
  }));
});

// Mock config
jest.mock('../../config', () => ({
  SERVER_NAME: 'test-outlook-assistant',
  SERVER_VERSION: '1.0.0-test',
  AUTH_CONFIG: {
    clientId: 'test-client-id',
    clientSecret: 'test-client-secret',
    tenantId: 'test-tenant-id',
    tokenEndpoint: 'https://login.microsoftonline.com/test-tenant-id/oauth2/v2.0/token',
    redirectUri: 'http://localhost:3333/auth/callback',
    hostedRedirectUri: '',
    scopes: ['offline_access', 'Mail.Read', 'User.Read'],
    tokenStorePath: '/tmp/test-tokens.json',
    hostedTokenStorePath: '/tmp/test-hosted-tokens.json',
  },
  CONNECTOR_AUTH: {
    apiAppId: 'api://test-app-id',
    apiScope: 'mcp.access',
    oboScopes: 'Mail.Read Mail.Send User.Read',
  },
}));

// Mock MCP SDK — Server and StreamableHTTPServerTransport
const mockHandleRequest = jest.fn().mockImplementation((_req, res) => {
  if (!res.headersSent) {
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ jsonrpc: '2.0', result: {} }));
  }
});
jest.mock('@modelcontextprotocol/sdk/server/index.js', () => ({
  Server: jest.fn().mockImplementation(() => ({
    connect: jest.fn().mockResolvedValue(undefined),
    close: jest.fn().mockResolvedValue(undefined),
    fallbackRequestHandler: null,
  })),
}));
jest.mock('@modelcontextprotocol/sdk/server/streamableHttp.js', () => ({
  StreamableHTTPServerTransport: jest.fn().mockImplementation(() => ({
    handleRequest: mockHandleRequest,
  })),
}));

// ── Imports (after all mocks) ─────────────────────────────────────────

const { createHttpApp } = require('../../transport/http-server');
const { getUserContext } = require('../../auth/request-context');
const SessionStore = require('../../auth/session-store');
const config = require('../../config');

// ── Helpers ───────────────────────────────────────────────────────────

const originalFetch = global.fetch;

/**
 * Creates a SessionStore with file persistence disabled.
 */
function createInMemorySessionStore() {
  const store = new SessionStore({});
  store.filePath = null;
  store.saveToFile = async () => {};
  return store;
}

/**
 * Creates a signed Entra JWT with the test RSA key pair.
 */
function createTestJwt(overrides = {}) {
  const jwtOptions = overrides._jwtOptions || {};
  const payload = { ...overrides };
  delete payload._jwtOptions;

  const defaults = {
    oid: 'test-oid-123',
    sub: 'test-sub-456',
    preferred_username: 'testuser@example.com',
    name: 'Test User',
    scp: 'mcp.access',
  };

  const finalPayload = { ...defaults, ...payload };

  return jwt.sign(finalPayload, privateKey, {
    algorithm: 'RS256',
    issuer: `https://login.microsoftonline.com/${config.AUTH_CONFIG.tenantId}/v2.0`,
    audience: config.CONNECTOR_AUTH.apiAppId,
    expiresIn: '1h',
    keyid: 'test-kid-123',
    ...jwtOptions,
  });
}

/**
 * Build a standard JSON-RPC initialize request body.
 */
function initializeBody() {
  return {
    jsonrpc: '2.0',
    id: 1,
    method: 'initialize',
    params: {
      protocolVersion: '2024-11-05',
      capabilities: {},
      clientInfo: { name: 'test-client', version: '0.1.0' },
    },
  };
}

// ── Tests ─────────────────────────────────────────────────────────────

describe('Connector Auth Flow — Integration', () => {
  let sessionStore;

  beforeEach(() => {
    jest.clearAllMocks();
    sessionStore = createInMemorySessionStore();

    // Default: mock fetch for OBO exchange
    global.fetch = jest.fn((url) => {
      if (typeof url === 'string' && url.includes('/oauth2/v2.0/token')) {
        return Promise.resolve({
          ok: true,
          json: () => Promise.resolve({
            access_token: 'mock-graph-token',
            expires_in: 3600,
          }),
        });
      }
      return Promise.reject(new Error(`Unexpected fetch: ${url}`));
    });
  });

  afterEach(() => {
    global.fetch = originalFetch;
  });

  // ═══════════════════════════════════════════════════════════════════════
  // Full connector flow
  // ═══════════════════════════════════════════════════════════════════════

  test('1. Valid Entra JWT → request succeeds via connector path', async () => {
    const app = createHttpApp({ sessionStore });
    const token = createTestJwt();

    let capturedContext = null;

    mockHandleRequest.mockImplementationOnce((_req, res) => {
      capturedContext = getUserContext();
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ jsonrpc: '2.0', result: {} }));
    });

    const res = await supertest(app)
      .post('/mcp')
      .set('Authorization', `Bearer ${token}`)
      .set('Content-Type', 'application/json')
      .send(initializeBody());

    expect(res.status).toBe(200);
    expect(mockHandleRequest).toHaveBeenCalled();

    // Verify user context was set with connector auth method
    expect(capturedContext).not.toBeNull();
    expect(capturedContext.authMethod).toBe('connector');
    expect(capturedContext.userId).toBe('test-oid-123');
    expect(capturedContext.entraToken).toBe(token);
  });

  test('1b. Connector path sets req.entraUser, skipping session validation', async () => {
    const app = createHttpApp({ sessionStore });
    const token = createTestJwt();

    // Session store should NOT be queried for connector auth
    const validateSpy = jest.spyOn(sessionStore, 'validateSession');

    const res = await supertest(app)
      .post('/mcp')
      .set('Authorization', `Bearer ${token}`)
      .set('Content-Type', 'application/json')
      .send(initializeBody());

    expect(res.status).toBe(200);
    // The session middleware should have been skipped because req.entraUser was set
    expect(validateSpy).not.toHaveBeenCalled();

    validateSpy.mockRestore();
  });

  // ═══════════════════════════════════════════════════════════════════════
  // Invalid JWT → 401
  // ═══════════════════════════════════════════════════════════════════════

  test('2. Invalid JWT (not a valid session either) → 401 with WWW-Authenticate', async () => {
    const app = createHttpApp({ sessionStore });

    const res = await supertest(app)
      .post('/mcp')
      .set('Authorization', 'Bearer not-a-valid-jwt-or-session')
      .set('Content-Type', 'application/json')
      .send(initializeBody());

    expect(res.status).toBe(401);
    expect(res.body.error).toBe('auth_required');

    // Should include WWW-Authenticate header with resource_metadata
    const wwwAuth = res.headers['www-authenticate'];
    expect(wwwAuth).toBeDefined();
    expect(wwwAuth).toContain('Bearer');
    expect(wwwAuth).toContain('resource_metadata=');
    expect(wwwAuth).toContain('.well-known/oauth-protected-resource');

    expect(mockHandleRequest).not.toHaveBeenCalled();
  });

  // ═══════════════════════════════════════════════════════════════════════
  // Expired JWT → 401
  // ═══════════════════════════════════════════════════════════════════════

  test('3. Expired Entra JWT → 401 with WWW-Authenticate', async () => {
    const app = createHttpApp({ sessionStore });
    const expiredToken = createTestJwt({
      _jwtOptions: { expiresIn: '-10s' },
    });

    const res = await supertest(app)
      .post('/mcp')
      .set('Authorization', `Bearer ${expiredToken}`)
      .set('Content-Type', 'application/json')
      .send(initializeBody());

    expect(res.status).toBe(401);
    expect(res.body.error).toBe('auth_required');

    const wwwAuth = res.headers['www-authenticate'];
    expect(wwwAuth).toBeDefined();
    expect(wwwAuth).toContain('Bearer');

    expect(mockHandleRequest).not.toHaveBeenCalled();
  });

  // ═══════════════════════════════════════════════════════════════════════
  // Wrong audience JWT → 401
  // ═══════════════════════════════════════════════════════════════════════

  test('3b. JWT with wrong audience → 401', async () => {
    const app = createHttpApp({ sessionStore });
    const wrongAudToken = createTestJwt({
      _jwtOptions: { audience: 'api://wrong-app-id' },
    });

    const res = await supertest(app)
      .post('/mcp')
      .set('Authorization', `Bearer ${wrongAudToken}`)
      .set('Content-Type', 'application/json')
      .send(initializeBody());

    expect(res.status).toBe(401);
    expect(res.body.error).toBe('auth_required');
    expect(mockHandleRequest).not.toHaveBeenCalled();
  });

  // ═══════════════════════════════════════════════════════════════════════
  // Missing required scope → 401
  // ═══════════════════════════════════════════════════════════════════════

  test('3c. JWT missing required scope → 401', async () => {
    const app = createHttpApp({ sessionStore });
    const noScopeToken = createTestJwt({ scp: 'some.other.scope' });

    const res = await supertest(app)
      .post('/mcp')
      .set('Authorization', `Bearer ${noScopeToken}`)
      .set('Content-Type', 'application/json')
      .send(initializeBody());

    expect(res.status).toBe(401);
    expect(res.body.error).toBe('auth_required');
    expect(mockHandleRequest).not.toHaveBeenCalled();
  });

  test('3d. JWT missing oid → 401 and does not reach connector context', async () => {
    const app = createHttpApp({ sessionStore });
    const tokenWithoutOid = jwt.sign(
      {
        sub: 'test-sub-456',
        preferred_username: 'testuser@example.com',
        name: 'Test User',
        scp: 'mcp.access',
      },
      privateKey,
      {
        algorithm: 'RS256',
        issuer: `https://login.microsoftonline.com/${config.AUTH_CONFIG.tenantId}/v2.0`,
        audience: config.CONNECTOR_AUTH.apiAppId,
        expiresIn: '1h',
        keyid: 'test-kid-123',
      }
    );

    const res = await supertest(app)
      .post('/mcp')
      .set('Authorization', `Bearer ${tokenWithoutOid}`)
      .set('Content-Type', 'application/json')
      .send(initializeBody());

    expect(res.status).toBe(401);
    expect(res.body.error).toBe('auth_required');
    expect(mockHandleRequest).not.toHaveBeenCalled();
  });

  // ═══════════════════════════════════════════════════════════════════════
  // No Authorization header → 401
  // ═══════════════════════════════════════════════════════════════════════

  test('4. No Authorization header → 401 with WWW-Authenticate', async () => {
    const app = createHttpApp({ sessionStore });

    const res = await supertest(app)
      .post('/mcp')
      .set('Content-Type', 'application/json')
      .send(initializeBody());

    expect(res.status).toBe(401);
    expect(res.body.error).toBe('auth_required');

    const wwwAuth = res.headers['www-authenticate'];
    expect(wwwAuth).toBeDefined();
    expect(wwwAuth).toContain('resource_metadata=');

    expect(mockHandleRequest).not.toHaveBeenCalled();
  });

  // ═══════════════════════════════════════════════════════════════════════
  // User context fields are correct
  // ═══════════════════════════════════════════════════════════════════════

  test('5. Connector context includes oid, authMethod, and entraToken', async () => {
    const app = createHttpApp({ sessionStore });
    const oid = 'custom-oid-abc-789';
    const token = createTestJwt({ oid });

    let capturedContext = null;

    mockHandleRequest.mockImplementationOnce((_req, res) => {
      capturedContext = getUserContext();
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ jsonrpc: '2.0', result: {} }));
    });

    await supertest(app)
      .post('/mcp')
      .set('Authorization', `Bearer ${token}`)
      .set('Content-Type', 'application/json')
      .send(initializeBody());

    expect(capturedContext).toEqual({
      userId: oid,
      authMethod: 'connector',
      entraToken: token,
    });
  });
});
