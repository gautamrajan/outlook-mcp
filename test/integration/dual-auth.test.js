/**
 * Integration tests for dual-auth coexistence.
 *
 * Verifies that session-based auth (legacy) and connector-based auth (Entra JWT)
 * both work independently and can coexist within the same Express app instance.
 *
 * Strategy:
 *   - Generate RSA key pair for test JWT signing
 *   - Mock `jwks-rsa` to return test public key
 *   - Use real SessionStore (in-memory, no file persistence)
 *   - Mock MCP SDK classes so we test auth wiring, not SDK internals
 *   - supertest against createHttpApp()
 */

const crypto = require('crypto');
const jwt = require('jsonwebtoken');
const supertest = require('supertest');

// ── RSA key pair for signing test JWTs ────────────────────────────────

const { publicKey: mockPublicKey, privateKey } = crypto.generateKeyPairSync('rsa', {
  modulusLength: 2048,
  publicKeyEncoding: { type: 'spki', format: 'pem' },
  privateKeyEncoding: { type: 'pkcs8', format: 'pem' },
});

// ── Mocks (must come before any require of modules under test) ────────

jest.mock('jwks-rsa', () => {
  return jest.fn(() => ({
    getSigningKey: jest.fn((_kid, callback) => {
      callback(null, { getPublicKey: () => mockPublicKey });
    }),
  }));
});

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

function createInMemorySessionStore() {
  const store = new SessionStore({});
  store.filePath = null;
  store.saveToFile = async () => {};
  return store;
}

function createTestJwt(overrides = {}) {
  const jwtOptions = overrides._jwtOptions || {};
  const payload = { ...overrides };
  delete payload._jwtOptions;

  const defaults = {
    oid: 'connector-oid-456',
    sub: 'connector-sub-789',
    preferred_username: 'connector@example.com',
    name: 'Connector User',
    scp: 'mcp.access',
  };

  return jwt.sign({ ...defaults, ...payload }, privateKey, {
    algorithm: 'RS256',
    issuer: `https://login.microsoftonline.com/${config.AUTH_CONFIG.tenantId}/v2.0`,
    audience: config.CONNECTOR_AUTH.apiAppId,
    expiresIn: '1h',
    keyid: 'test-kid-dual',
    ...jwtOptions,
  });
}

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

describe('Dual Auth — Integration', () => {
  let sessionStore;
  let app;

  beforeEach(() => {
    jest.clearAllMocks();
    sessionStore = createInMemorySessionStore();
    app = createHttpApp({ sessionStore });

    // Mock fetch for OBO (connector path may trigger it indirectly)
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
  // Session auth still works
  // ═══════════════════════════════════════════════════════════════════════

  test('1. Session auth works — valid session token → 200', async () => {
    const sessionToken = await sessionStore.createSession('session-user-abc');

    let capturedContext = null;

    mockHandleRequest.mockImplementationOnce((_req, res) => {
      capturedContext = getUserContext();
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ jsonrpc: '2.0', result: {} }));
    });

    const res = await supertest(app)
      .post('/mcp')
      .set('Authorization', `Bearer ${sessionToken}`)
      .set('Content-Type', 'application/json')
      .send(initializeBody());

    expect(res.status).toBe(200);
    expect(mockHandleRequest).toHaveBeenCalled();

    // Verify session auth context
    expect(capturedContext).not.toBeNull();
    expect(capturedContext.authMethod).toBe('session');
    expect(capturedContext.userId).toBe('session-user-abc');
    expect(capturedContext.sessionToken).toBe(sessionToken);
  });

  // ═══════════════════════════════════════════════════════════════════════
  // Connector auth works alongside session auth
  // ═══════════════════════════════════════════════════════════════════════

  test('2. Connector auth works alongside — valid Entra JWT → 200', async () => {
    const entraJwt = createTestJwt();

    let capturedContext = null;

    mockHandleRequest.mockImplementationOnce((_req, res) => {
      capturedContext = getUserContext();
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ jsonrpc: '2.0', result: {} }));
    });

    const res = await supertest(app)
      .post('/mcp')
      .set('Authorization', `Bearer ${entraJwt}`)
      .set('Content-Type', 'application/json')
      .send(initializeBody());

    expect(res.status).toBe(200);
    expect(mockHandleRequest).toHaveBeenCalled();

    // Verify connector auth context
    expect(capturedContext).not.toBeNull();
    expect(capturedContext.authMethod).toBe('connector');
    expect(capturedContext.userId).toBe('connector-oid-456');
    expect(capturedContext.entraToken).toBe(entraJwt);
  });

  // ═══════════════════════════════════════════════════════════════════════
  // Sequential requests with different auth methods
  // ═══════════════════════════════════════════════════════════════════════

  test('3. Sequential requests — session then connector → both succeed independently', async () => {
    const sessionToken = await sessionStore.createSession('session-user-xyz');
    const entraJwt = createTestJwt({ oid: 'connector-oid-xyz' });

    // --- Request 1: session auth ---
    let sessionContext = null;
    mockHandleRequest.mockImplementationOnce((_req, res) => {
      sessionContext = getUserContext();
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ jsonrpc: '2.0', result: {} }));
    });

    const res1 = await supertest(app)
      .post('/mcp')
      .set('Authorization', `Bearer ${sessionToken}`)
      .set('Content-Type', 'application/json')
      .send(initializeBody());

    expect(res1.status).toBe(200);
    expect(sessionContext.authMethod).toBe('session');
    expect(sessionContext.userId).toBe('session-user-xyz');

    // --- Request 2: connector auth ---
    let connectorContext = null;
    mockHandleRequest.mockImplementationOnce((_req, res) => {
      connectorContext = getUserContext();
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ jsonrpc: '2.0', result: {} }));
    });

    const res2 = await supertest(app)
      .post('/mcp')
      .set('Authorization', `Bearer ${entraJwt}`)
      .set('Content-Type', 'application/json')
      .send(initializeBody());

    expect(res2.status).toBe(200);
    expect(connectorContext.authMethod).toBe('connector');
    expect(connectorContext.userId).toBe('connector-oid-xyz');

    // Verify they are independent contexts
    expect(sessionContext.userId).not.toBe(connectorContext.userId);
    expect(sessionContext.authMethod).not.toBe(connectorContext.authMethod);
  });

  // ═══════════════════════════════════════════════════════════════════════
  // Connector then session — reverse order also works
  // ═══════════════════════════════════════════════════════════════════════

  test('4. Sequential requests — connector then session → both succeed', async () => {
    const entraJwt = createTestJwt({ oid: 'connector-first-oid' });
    const sessionToken = await sessionStore.createSession('session-second-user');

    // --- Request 1: connector auth ---
    let ctx1 = null;
    mockHandleRequest.mockImplementationOnce((_req, res) => {
      ctx1 = getUserContext();
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ jsonrpc: '2.0', result: {} }));
    });

    const res1 = await supertest(app)
      .post('/mcp')
      .set('Authorization', `Bearer ${entraJwt}`)
      .set('Content-Type', 'application/json')
      .send(initializeBody());

    expect(res1.status).toBe(200);
    expect(ctx1.authMethod).toBe('connector');

    // --- Request 2: session auth ---
    let ctx2 = null;
    mockHandleRequest.mockImplementationOnce((_req, res) => {
      ctx2 = getUserContext();
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ jsonrpc: '2.0', result: {} }));
    });

    const res2 = await supertest(app)
      .post('/mcp')
      .set('Authorization', `Bearer ${sessionToken}`)
      .set('Content-Type', 'application/json')
      .send(initializeBody());

    expect(res2.status).toBe(200);
    expect(ctx2.authMethod).toBe('session');
    expect(ctx2.userId).toBe('session-second-user');
  });

  // ═══════════════════════════════════════════════════════════════════════
  // Invalid tokens from both paths still → 401
  // ═══════════════════════════════════════════════════════════════════════

  test('5. Invalid token (neither valid JWT nor session) → 401', async () => {
    const res = await supertest(app)
      .post('/mcp')
      .set('Authorization', 'Bearer garbage-token-12345')
      .set('Content-Type', 'application/json')
      .send(initializeBody());

    expect(res.status).toBe(401);
    expect(res.body.error).toBe('auth_required');
    expect(mockHandleRequest).not.toHaveBeenCalled();
  });

  // ═══════════════════════════════════════════════════════════════════════
  // Connector user with session user on same app — isolation
  // ═══════════════════════════════════════════════════════════════════════

  test('6. Different users via different auth methods — contexts are isolated', async () => {
    const sessionToken1 = await sessionStore.createSession('user-session-A');
    const sessionToken2 = await sessionStore.createSession('user-session-B');
    const entraJwt = createTestJwt({ oid: 'user-connector-C' });

    const contexts = [];

    // All three requests
    for (const authToken of [sessionToken1, entraJwt, sessionToken2]) {
      mockHandleRequest.mockImplementationOnce((_req, res) => {
        contexts.push(getUserContext());
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ jsonrpc: '2.0', result: {} }));
      });

      const res = await supertest(app)
        .post('/mcp')
        .set('Authorization', `Bearer ${authToken}`)
        .set('Content-Type', 'application/json')
        .send(initializeBody());

      expect(res.status).toBe(200);
    }

    expect(contexts).toHaveLength(3);
    expect(contexts[0].userId).toBe('user-session-A');
    expect(contexts[0].authMethod).toBe('session');
    expect(contexts[1].userId).toBe('user-connector-C');
    expect(contexts[1].authMethod).toBe('connector');
    expect(contexts[2].userId).toBe('user-session-B');
    expect(contexts[2].authMethod).toBe('session');
  });
});
