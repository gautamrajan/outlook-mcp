/**
 * Integration tests for the full hosted auth flow.
 *
 * These tests wire together REAL instances of PerUserTokenStorage, SessionStore,
 * createAuthRoutes, createHttpApp, and ensureAuthenticated — verifying the
 * end-to-end hosted multi-user auth flow without hitting real Entra/Graph
 * endpoints (fetch is mocked for external calls).
 *
 * Strategy:
 *   - PerUserTokenStorage and SessionStore are real in-memory instances (no filePath)
 *   - supertest for HTTP-level assertions
 *   - global.fetch mocked for Entra token endpoint + Graph /me calls
 *   - AsyncLocalStorage requestContext used to set user context for
 *     ensureAuthenticated tests
 */

const crypto = require('crypto');
const supertest = require('supertest');
const express = require('express');

const PerUserTokenStorage = require('../../auth/per-user-token-storage');
const SessionStore = require('../../auth/session-store');
const { createAuthRoutes, _pendingAuth } = require('../../auth/auth-routes');

// ── Helpers ──────────────────────────────────────────────────────────────

/** Save and restore global.fetch across tests. */
const originalFetch = global.fetch;

/**
 * Build a mock config object matching what auth-routes and auth/index expect.
 * Prefixed with "mock" so Jest allows it inside jest.mock() factories.
 */
function mockBuildTestConfig() {
  return {
    AUTH_CONFIG: {
      clientId: 'integ-client-id',
      clientSecret: 'integ-client-secret',
      tenantId: 'integ-tenant-id',
      tokenEndpoint: 'https://login.microsoftonline.com/integ-tenant-id/oauth2/v2.0/token',
      redirectUri: 'http://localhost:3333/auth/callback',
      hostedRedirectUri: 'https://outlook-mcp.example.com/auth/callback',
      scopes: ['offline_access', 'Mail.Read', 'User.Read'],
      tokenStorePath: '/tmp/integ-test-tokens.json',
      hostedTokenStorePath: '/tmp/integ-test-hosted-tokens.json',
    },
    SERVER_NAME: 'outlook-assistant-integ',
    SERVER_VERSION: '1.0.0-test',
  };
}

/**
 * Creates a mock fetch function for Entra token + Graph /me endpoints.
 */
function createMockFetch({
  tokenResponse = {
    access_token: 'integ-access-token',
    refresh_token: 'integ-refresh-token',
    expires_in: 3600,
    scope: 'offline_access Mail.Read User.Read',
  },
  tokenStatus = 200,
  meResponse = {
    id: 'user-oid-integ-1',
    mail: 'integuser@example.com',
    displayName: 'Integration User',
  },
  meStatus = 200,
} = {}) {
  return jest.fn().mockImplementation((url) => {
    if (typeof url === 'string' && url.includes('oauth2/v2.0/token')) {
      return Promise.resolve({
        ok: tokenStatus >= 200 && tokenStatus < 300,
        status: tokenStatus,
        text: () => Promise.resolve(JSON.stringify(tokenResponse)),
        json: () => Promise.resolve(tokenResponse),
      });
    }
    if (typeof url === 'string' && url.includes('graph.microsoft.com')) {
      return Promise.resolve({
        ok: meStatus >= 200 && meStatus < 300,
        status: meStatus,
        text: () => Promise.resolve(JSON.stringify(meResponse)),
        json: () => Promise.resolve(meResponse),
      });
    }
    return Promise.reject(new Error(`Unexpected fetch URL: ${url}`));
  });
}

/**
 * Creates a SessionStore with file persistence disabled.
 * SessionStore's constructor defaults filePath to a real path on disk,
 * so we neutralize it after construction.
 */
function createInMemorySessionStore() {
  const store = new SessionStore({});
  store.filePath = null;
  // Override saveToFile to be a no-op for in-memory operation
  store.saveToFile = async () => {};
  return store;
}

// ── Test suite ───────────────────────────────────────────────────────────

describe('Hosted Auth Flow — Integration', () => {
  afterEach(() => {
    _pendingAuth.clear();
    global.fetch = originalFetch;
  });

  // ════════════════════════════════════════════════════════════════════════
  // HTTP layer — session middleware on /mcp
  // ════════════════════════════════════════════════════════════════════════

  describe('HTTP layer', () => {
    let createHttpApp;

    beforeEach(() => {
      jest.resetModules();

      // Mock MCP SDK
      jest.mock('@modelcontextprotocol/sdk/server/index.js', () => ({
        Server: jest.fn().mockImplementation(() => ({
          connect: jest.fn().mockResolvedValue(undefined),
          close: jest.fn().mockResolvedValue(undefined),
          fallbackRequestHandler: null,
        })),
      }));

      jest.mock('@modelcontextprotocol/sdk/server/streamableHttp.js', () => ({
        StreamableHTTPServerTransport: jest.fn().mockImplementation(() => ({
          handleRequest: jest.fn().mockImplementation((_req, res) => {
            if (!res.headersSent) {
              res.writeHead(200, { 'Content-Type': 'application/json' });
              res.end(JSON.stringify({ jsonrpc: '2.0', result: {} }));
            }
          }),
        })),
      }));

      jest.mock('../../config', () => mockBuildTestConfig());

      createHttpApp = require('../../transport/http-server').createHttpApp;
    });

    afterEach(() => {
      jest.restoreAllMocks();
    });

    test('1. Unauthenticated request → 401 with auth URL', async () => {
      const sessionStore = createInMemorySessionStore();
      const app = createHttpApp({ sessionStore });

      const res = await supertest(app)
        .post('/mcp')
        .set('Content-Type', 'application/json')
        .send({ jsonrpc: '2.0', id: 1, method: 'initialize', params: {} });

      expect(res.status).toBe(401);
      expect(res.body).toEqual(
        expect.objectContaining({
          error: 'auth_required',
          authUrl: '/auth/login',
        }),
      );
      expect(res.body.message).toBeDefined();
    });

    test('2. Valid session token → request proceeds (200)', async () => {
      const sessionStore = createInMemorySessionStore();
      const sessionToken = await sessionStore.createSession('user-integ-abc');
      const app = createHttpApp({ sessionStore });

      const res = await supertest(app)
        .post('/mcp')
        .set('Authorization', `Bearer ${sessionToken}`)
        .set('Content-Type', 'application/json')
        .send({ jsonrpc: '2.0', id: 1, method: 'initialize', params: {} });

      expect(res.status).toBe(200);
    });

    test('3. Invalid/expired session token → 401', async () => {
      const sessionStore = createInMemorySessionStore();
      const app = createHttpApp({ sessionStore });

      const res = await supertest(app)
        .post('/mcp')
        .set('Authorization', 'Bearer totally-bogus-token')
        .set('Content-Type', 'application/json')
        .send({ jsonrpc: '2.0', id: 1, method: 'initialize', params: {} });

      expect(res.status).toBe(401);
      expect(res.body.error).toBe('auth_required');
    });

    test('3b. Expired session → 401', async () => {
      const sessionStore = createInMemorySessionStore();
      // Create a session that expires immediately (0 days)
      const token = await sessionStore.createSession('user-expired', { expiresInDays: 0 });
      const app = createHttpApp({ sessionStore });

      const res = await supertest(app)
        .post('/mcp')
        .set('Authorization', `Bearer ${token}`)
        .set('Content-Type', 'application/json')
        .send({ jsonrpc: '2.0', id: 1, method: 'initialize', params: {} });

      expect(res.status).toBe(401);
      expect(res.body.error).toBe('auth_required');
    });
  });

  // ════════════════════════════════════════════════════════════════════════
  // Browser auth flow — /auth/login and /auth/callback
  // ════════════════════════════════════════════════════════════════════════

  describe('Browser auth flow', () => {
    let tokenStorage;
    let sessionStore;
    let mockFetch;
    let app;

    beforeEach(() => {
      tokenStorage = new PerUserTokenStorage(); // in-memory, no filePath
      sessionStore = createInMemorySessionStore();
      mockFetch = createMockFetch();

      app = express();
      const router = createAuthRoutes({
        tokenStorage,
        sessionStore,
        config: mockBuildTestConfig(),
        fetch: mockFetch,
      });
      app.use('/auth', router);
    });

    test('4. GET /auth/login → 302 redirect to Entra with correct params', async () => {
      const res = await supertest(app)
        .get('/auth/login')
        .expect(302);

      const location = res.headers.location;
      expect(location).toContain(
        'https://login.microsoftonline.com/integ-tenant-id/oauth2/v2.0/authorize',
      );

      const url = new URL(location);
      const params = url.searchParams;

      expect(params.get('client_id')).toBe('integ-client-id');
      expect(params.get('response_type')).toBe('code');
      expect(params.get('redirect_uri')).toBe('https://outlook-mcp.example.com/auth/callback');
      expect(params.get('scope')).toBe('offline_access Mail.Read User.Read');
      expect(params.get('code_challenge_method')).toBe('S256');

      // PKCE: state and challenge must be present
      const state = params.get('state');
      const challenge = params.get('code_challenge');
      expect(state).toBeTruthy();
      expect(challenge).toBeTruthy();

      // Verify the challenge matches the stored verifier
      const pending = _pendingAuth.get(state);
      expect(pending).toBeDefined();
      const expectedChallenge = crypto
        .createHash('sha256')
        .update(pending.codeVerifier)
        .digest('base64url');
      expect(challenge).toBe(expectedChallenge);
    });

    test('5. Full /auth/callback → tokens stored, session created, HTML has token', async () => {
      // Step 1: /auth/login to get state
      const loginRes = await supertest(app).get('/auth/login').expect(302);
      const loginUrl = new URL(loginRes.headers.location);
      const state = loginUrl.searchParams.get('state');

      // Step 2: /auth/callback with code + state
      const callbackRes = await supertest(app)
        .get('/auth/callback')
        .query({ code: 'integ-auth-code', state })
        .expect(200);

      // Verify tokens were stored in the REAL PerUserTokenStorage
      const storedToken = tokenStorage.getTokenForUser('user-oid-integ-1');
      expect(storedToken).toBe('integ-access-token');

      const refreshToken = tokenStorage.getRefreshToken('user-oid-integ-1');
      expect(refreshToken).toBe('integ-refresh-token');

      const userInfo = tokenStorage.getUserInfo('user-oid-integ-1');
      expect(userInfo.email).toBe('integuser@example.com');
      expect(userInfo.name).toBe('Integration User');

      // Verify session was created in the REAL SessionStore
      expect(sessionStore.getSessionCountForUser('user-oid-integ-1')).toBe(1);

      // Verify HTML response contains the session token
      const html = callbackRes.text;
      expect(html).toContain('Authentication Successful');
      expect(html).toContain('Integration User');

      // Verify the session resolves in SessionStore
      const activeSessions = sessionStore.getActiveSessions();
      expect(activeSessions.length).toBe(1);
      expect(activeSessions[0].userId).toBe('user-oid-integ-1');
    });

    test('5b. Callback stores tokens that are retrievable and non-expired', async () => {
      const loginRes = await supertest(app).get('/auth/login').expect(302);
      const state = new URL(loginRes.headers.location).searchParams.get('state');

      await supertest(app)
        .get('/auth/callback')
        .query({ code: 'integ-auth-code', state })
        .expect(200);

      // Token should NOT be expired (was just stored with 3600s expiry)
      expect(tokenStorage.isTokenExpired('user-oid-integ-1')).toBe(false);
    });
  });

  // ════════════════════════════════════════════════════════════════════════
  // ensureAuthenticated hosted path
  // ════════════════════════════════════════════════════════════════════════

  describe('ensureAuthenticated hosted path', () => {
    // After jest.resetModules(), both auth/index.js and our test must share
    // the SAME request-context module instance so that requestContext.run()
    // in the test is visible to isHostedMode() inside auth/index.js.

    let ensureAuthenticated;
    let setHostedTokenStorage;
    let perUserStorage;
    let requestContext; // re-required after resetModules

    beforeEach(() => {
      jest.resetModules();
      global.fetch = jest.fn();

      jest.mock('../../auth/tools', () => ({
        authTools: [{ name: 'mock-tool' }],
      }));

      jest.mock('../../config', () => mockBuildTestConfig());

      // Mock TokenStorage (local mode) so it doesn't try to read files
      jest.mock('../../auth/token-storage', () => {
        return jest.fn().mockImplementation(() => ({
          getValidAccessToken: jest.fn().mockResolvedValue(null),
          invalidateAccessToken: jest.fn(),
        }));
      });

      // Real per-user storage (in-memory)
      perUserStorage = new PerUserTokenStorage();

      // Re-require request-context so we get the SAME instance as auth/index.js
      requestContext = require('../../auth/request-context').requestContext;

      // Require auth module fresh (it will use the same request-context)
      const authModule = require('../../auth/index');
      ensureAuthenticated = authModule.ensureAuthenticated;
      setHostedTokenStorage = authModule.setHostedTokenStorage;

      // Inject our in-memory storage
      setHostedTokenStorage(perUserStorage);
    });

    afterEach(() => {
      global.fetch = originalFetch;
    });

    test('6. Valid token → returns access token without network calls', async () => {
      await perUserStorage.setTokensForUser('user-A', {
        accessToken: 'valid-access-A',
        refreshToken: 'refresh-A',
        expiresIn: 3600,
        scopes: 'Mail.Read User.Read',
        email: 'a@example.com',
        name: 'User A',
      });

      let result;
      await requestContext.run({ userId: 'user-A' }, async () => {
        result = await ensureAuthenticated();
      });

      expect(result).toBe('valid-access-A');
      expect(global.fetch).not.toHaveBeenCalled();
    });

    test('7. Expired token + valid refresh → silent refresh returns new token', async () => {
      await perUserStorage.setTokensForUser('user-B', {
        accessToken: 'expired-access-B',
        refreshToken: 'valid-refresh-B',
        expiresIn: 0, // immediately expired due to 5-min buffer
        scopes: 'Mail.Read User.Read',
        email: 'b@example.com',
        name: 'User B',
      });

      // Mock fetch to return new tokens from Entra
      global.fetch = jest.fn().mockResolvedValue({
        ok: true,
        status: 200,
        json: () => Promise.resolve({
          access_token: 'refreshed-access-B',
          refresh_token: 'rotated-refresh-B',
          expires_in: 3600,
          scope: 'Mail.Read User.Read',
        }),
      });

      let result;
      await requestContext.run({ userId: 'user-B' }, async () => {
        result = await ensureAuthenticated();
      });

      expect(result).toBe('refreshed-access-B');

      // Verify fetch was called to the token endpoint
      expect(global.fetch).toHaveBeenCalledTimes(1);
      const [fetchUrl, fetchOpts] = global.fetch.mock.calls[0];
      expect(fetchUrl).toContain('oauth2/v2.0/token');
      expect(fetchOpts.method).toBe('POST');
      expect(fetchOpts.body).toContain('grant_type=refresh_token');
      expect(fetchOpts.body).toContain('refresh_token=valid-refresh-B');

      // Verify new tokens are persisted in storage
      const newToken = perUserStorage.getTokenForUser('user-B');
      expect(newToken).toBe('refreshed-access-B');
      const newRefresh = perUserStorage.getRefreshToken('user-B');
      expect(newRefresh).toBe('rotated-refresh-B');
    });

    test('8. Expired token + no refresh → AUTH_REQUIRED error', async () => {
      await perUserStorage.setTokensForUser('user-C', {
        accessToken: 'expired-access-C',
        refreshToken: null,
        expiresIn: 0,
        scopes: 'Mail.Read User.Read',
        email: 'c@example.com',
        name: 'User C',
      });

      let error;
      await requestContext.run({ userId: 'user-C' }, async () => {
        try {
          await ensureAuthenticated();
        } catch (e) {
          error = e;
        }
      });

      expect(error).toBeDefined();
      expect(error.code).toBe('AUTH_REQUIRED');
      expect(error.message).toContain('Authentication required');
    });

    test('9. Two users → independent token stores return correct tokens', async () => {
      await perUserStorage.setTokensForUser('user-X', {
        accessToken: 'token-for-X',
        refreshToken: 'refresh-X',
        expiresIn: 3600,
        scopes: 'Mail.Read',
        email: 'x@example.com',
        name: 'User X',
      });

      await perUserStorage.setTokensForUser('user-Y', {
        accessToken: 'token-for-Y',
        refreshToken: 'refresh-Y',
        expiresIn: 3600,
        scopes: 'Mail.Read',
        email: 'y@example.com',
        name: 'User Y',
      });

      let resultX, resultY;

      await requestContext.run({ userId: 'user-X' }, async () => {
        resultX = await ensureAuthenticated();
      });

      await requestContext.run({ userId: 'user-Y' }, async () => {
        resultY = await ensureAuthenticated();
      });

      expect(resultX).toBe('token-for-X');
      expect(resultY).toBe('token-for-Y');
      expect(resultX).not.toBe(resultY);
      expect(global.fetch).not.toHaveBeenCalled();
    });

    test('7b. Refresh failure → AUTH_REQUIRED error', async () => {
      await perUserStorage.setTokensForUser('user-D', {
        accessToken: 'expired-access-D',
        refreshToken: 'bad-refresh-D',
        expiresIn: 0,
        scopes: 'Mail.Read',
        email: 'd@example.com',
        name: 'User D',
      });

      global.fetch = jest.fn().mockResolvedValue({
        ok: false,
        status: 400,
        json: () => Promise.resolve({ error: 'invalid_grant' }),
      });

      let error;
      await requestContext.run({ userId: 'user-D' }, async () => {
        try {
          await ensureAuthenticated();
        } catch (e) {
          error = e;
        }
      });

      expect(error).toBeDefined();
      expect(error.code).toBe('AUTH_REQUIRED');
    });
  });

  // ════════════════════════════════════════════════════════════════════════
  // Local mode isolation
  // ════════════════════════════════════════════════════════════════════════

  describe('Local mode isolation', () => {
    test('10. Without hosted context, ensureAuthenticated uses local TokenStorage', async () => {
      jest.resetModules();

      const mockGetValidAccessToken = jest.fn().mockResolvedValue('local-access-token');

      jest.mock('../../auth/tools', () => ({
        authTools: [{ name: 'mock-tool' }],
      }));

      jest.mock('../../config', () => mockBuildTestConfig());

      jest.mock('../../auth/token-storage', () => {
        return jest.fn().mockImplementation(() => ({
          getValidAccessToken: mockGetValidAccessToken,
          invalidateAccessToken: jest.fn(),
        }));
      });

      const { ensureAuthenticated } = require('../../auth/index');

      // Call WITHOUT requestContext.run → isHostedMode() returns false
      const result = await ensureAuthenticated();

      expect(result).toBe('local-access-token');
      expect(mockGetValidAccessToken).toHaveBeenCalledTimes(1);
    });

    test('10b. Local mode: no token → throws Authentication required', async () => {
      jest.resetModules();

      const mockGetValidAccessToken = jest.fn().mockResolvedValue(null);

      jest.mock('../../auth/tools', () => ({
        authTools: [{ name: 'mock-tool' }],
      }));

      jest.mock('../../config', () => mockBuildTestConfig());

      jest.mock('../../auth/token-storage', () => {
        return jest.fn().mockImplementation(() => ({
          getValidAccessToken: mockGetValidAccessToken,
          invalidateAccessToken: jest.fn(),
        }));
      });

      const { ensureAuthenticated } = require('../../auth/index');

      await expect(ensureAuthenticated()).rejects.toThrow('Authentication required');
    });
  });

  // ════════════════════════════════════════════════════════════════════════
  // End-to-end: browser auth → session → MCP request
  // ════════════════════════════════════════════════════════════════════════

  describe('End-to-end: auth flow → MCP request', () => {
    test('user authenticates via browser then uses session token for MCP', async () => {
      jest.resetModules();

      // Mock MCP SDK
      jest.mock('@modelcontextprotocol/sdk/server/index.js', () => ({
        Server: jest.fn().mockImplementation(() => ({
          connect: jest.fn().mockResolvedValue(undefined),
          close: jest.fn().mockResolvedValue(undefined),
          fallbackRequestHandler: null,
        })),
      }));

      jest.mock('@modelcontextprotocol/sdk/server/streamableHttp.js', () => ({
        StreamableHTTPServerTransport: jest.fn().mockImplementation(() => ({
          handleRequest: jest.fn().mockImplementation((_req, res) => {
            if (!res.headersSent) {
              res.writeHead(200, { 'Content-Type': 'application/json' });
              res.end(JSON.stringify({ jsonrpc: '2.0', result: { ok: true } }));
            }
          }),
        })),
      }));

      jest.mock('../../config', () => mockBuildTestConfig());

      const { createHttpApp } = require('../../transport/http-server');
      const { createAuthRoutes: freshCreateAuthRoutes, _pendingAuth: freshPending } =
        require('../../auth/auth-routes');

      const tokenStorage = new PerUserTokenStorage();
      const sessionStore = createInMemorySessionStore();
      const mockFetch = createMockFetch();

      // Build combined app: auth routes + MCP routes with session middleware
      const app = createHttpApp({ sessionStore });
      const authRouter = freshCreateAuthRoutes({
        tokenStorage,
        sessionStore,
        config: mockBuildTestConfig(),
        fetch: mockFetch,
      });
      app.use('/auth', authRouter);

      // Step 1: Browser visits /auth/login
      const loginRes = await supertest(app).get('/auth/login').expect(302);
      const state = new URL(loginRes.headers.location).searchParams.get('state');

      // Step 2: Browser returns to /auth/callback
      await supertest(app)
        .get('/auth/callback')
        .query({ code: 'e2e-auth-code', state })
        .expect(200);

      // Get the full session token from the session store internals
      const allEntries = Array.from(sessionStore._sessions.entries());
      expect(allEntries.length).toBe(1);
      const [fullSessionToken] = allEntries[0];

      // Step 3: Use session token to make an MCP request
      const mcpRes = await supertest(app)
        .post('/mcp')
        .set('Authorization', `Bearer ${fullSessionToken}`)
        .set('Content-Type', 'application/json')
        .send({ jsonrpc: '2.0', id: 1, method: 'initialize', params: {} });

      expect(mcpRes.status).toBe(200);

      // Step 4: Verify an unauthenticated request still fails
      const unauthRes = await supertest(app)
        .post('/mcp')
        .set('Content-Type', 'application/json')
        .send({ jsonrpc: '2.0', id: 1, method: 'initialize', params: {} });

      expect(unauthRes.status).toBe(401);

      // Cleanup
      freshPending.clear();
    });
  });
});
