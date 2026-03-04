const crypto = require('crypto');
const express = require('express');
const supertest = require('supertest');
const {
  createAuthRoutes,
  generatePKCE,
  _pendingAuth,
  _consumedAuthStates,
} = require('../../auth/auth-routes');

// ── Mock factories ──────────────────────────────────────────────────────

function createMockTokenStorage() {
  return {
    setTokensForUser: jest.fn().mockResolvedValue(undefined),
    getTokenForUser: jest.fn(),
    getRefreshToken: jest.fn(),
    isTokenExpired: jest.fn(),
    getUserInfo: jest.fn(),
  };
}

function createMockSessionStore() {
  return {
    createSession: jest.fn().mockResolvedValue('mock-session-token-uuid'),
    validateSession: jest.fn(),
    revokeSession: jest.fn(),
  };
}

function createMockConfig() {
  return {
    AUTH_CONFIG: {
      clientId: 'test-client-id',
      clientSecret: 'test-client-secret',
      tenantId: 'test-tenant-id',
      redirectUri: 'http://localhost:3333/auth/callback',
      hostedRedirectUri: 'https://outlook-mcp.example.com/auth/callback',
      scopes: ['offline_access', 'Mail.Read', 'User.Read'],
    },
    PORT: 3000,
  };
}

/**
 * Creates a mock fetch function that responds to the Entra token endpoint
 * and the Graph /me endpoint.
 */
function createMockFetch({
  tokenResponse = {
    access_token: 'mock-access-token',
    refresh_token: 'mock-refresh-token',
    expires_in: 3600,
    scope: 'offline_access Mail.Read User.Read',
  },
  tokenStatus = 200,
  meResponse = {
    id: 'user-oid-123',
    mail: 'user@example.com',
    displayName: 'Test User',
  },
  meStatus = 200,
} = {}) {
  return jest.fn().mockImplementation((url, opts) => {
    if (url.includes('oauth2/v2.0/token')) {
      return Promise.resolve({
        ok: tokenStatus >= 200 && tokenStatus < 300,
        status: tokenStatus,
        text: () => Promise.resolve(JSON.stringify(tokenResponse)),
        json: () => Promise.resolve(tokenResponse),
      });
    }
    if (url.includes('graph.microsoft.com')) {
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
 * Creates an Express app wired with the auth routes for testing.
 */
function createTestApp(overrides = {}) {
  const tokenStorage = overrides.tokenStorage || createMockTokenStorage();
  const sessionStore = overrides.sessionStore || createMockSessionStore();
  const config = overrides.config || createMockConfig();
  const fetch = overrides.fetch || createMockFetch();

  const app = express();
  const router = createAuthRoutes({ tokenStorage, sessionStore, config, fetch });
  app.use('/auth', router);

  return { app, tokenStorage, sessionStore, config, fetch };
}

// ── Cleanup ─────────────────────────────────────────────────────────────

afterEach(() => {
  // Clear the shared pendingAuth map between tests
  _pendingAuth.clear();
  _consumedAuthStates.clear();
});

// ── PKCE generation ─────────────────────────────────────────────────────

describe('generatePKCE', () => {
  test('returns verifier and challenge as base64url strings', () => {
    const { verifier, challenge } = generatePKCE();

    expect(typeof verifier).toBe('string');
    expect(typeof challenge).toBe('string');
    // base64url: only alphanumeric, hyphen, underscore (no padding = in our case)
    expect(verifier).toMatch(/^[A-Za-z0-9_-]+$/);
    expect(challenge).toMatch(/^[A-Za-z0-9_-]+$/);
  });

  test('challenge is the SHA-256 of the verifier in base64url', () => {
    const { verifier, challenge } = generatePKCE();
    const expected = crypto.createHash('sha256').update(verifier).digest('base64url');
    expect(challenge).toBe(expected);
  });

  test('generates unique verifiers on successive calls', () => {
    const a = generatePKCE();
    const b = generatePKCE();
    expect(a.verifier).not.toBe(b.verifier);
  });
});

// ── GET /auth/login ─────────────────────────────────────────────────────

describe('GET /auth/login', () => {
  test('redirects to Entra authorize endpoint', async () => {
    const { app } = createTestApp();

    const res = await supertest(app)
      .get('/auth/login')
      .expect(302);

    const location = res.headers.location;
    expect(location).toContain('https://login.microsoftonline.com/test-tenant-id/oauth2/v2.0/authorize');
  });

  test('includes required OAuth parameters in redirect URL', async () => {
    const { app } = createTestApp();

    const res = await supertest(app)
      .get('/auth/login')
      .expect(302);

    const location = res.headers.location;
    const url = new URL(location);
    const params = url.searchParams;

    expect(params.get('client_id')).toBe('test-client-id');
    expect(params.get('response_type')).toBe('code');
    expect(params.get('redirect_uri')).toBe('https://outlook-mcp.example.com/auth/callback');
    expect(params.get('scope')).toBe('offline_access Mail.Read User.Read');
    expect(params.get('code_challenge_method')).toBe('S256');
    expect(params.get('state')).toBeTruthy();
    expect(params.get('code_challenge')).toBeTruthy();
  });

  test('uses redirectUri as fallback when hostedRedirectUri is not set', async () => {
    const config = createMockConfig();
    delete config.AUTH_CONFIG.hostedRedirectUri;
    const { app } = createTestApp({ config });

    const res = await supertest(app)
      .get('/auth/login')
      .expect(302);

    const url = new URL(res.headers.location);
    expect(url.searchParams.get('redirect_uri')).toBe('http://localhost:3333/auth/callback');
  });

  test('stores PKCE verifier in pending auth map', async () => {
    const { app } = createTestApp();

    const res = await supertest(app)
      .get('/auth/login')
      .expect(302);

    const url = new URL(res.headers.location);
    const state = url.searchParams.get('state');

    expect(_pendingAuth.has(state)).toBe(true);
    const entry = _pendingAuth.get(state);
    expect(entry.codeVerifier).toBeTruthy();
    expect(entry.expiresAt).toBeGreaterThan(Date.now());
  });

  test('PKCE challenge in URL matches stored verifier', async () => {
    const { app } = createTestApp();

    const res = await supertest(app)
      .get('/auth/login')
      .expect(302);

    const url = new URL(res.headers.location);
    const state = url.searchParams.get('state');
    const challenge = url.searchParams.get('code_challenge');

    const entry = _pendingAuth.get(state);
    const expectedChallenge = crypto
      .createHash('sha256')
      .update(entry.codeVerifier)
      .digest('base64url');

    expect(challenge).toBe(expectedChallenge);
  });
});

// ── GET /auth/callback — success path ───────────────────────────────────

describe('GET /auth/callback — success', () => {
  /**
   * Helper: perform a full login + callback flow.
   * Returns the supertest response from the callback.
   */
  async function performAuthFlow(overrides = {}) {
    const testApp = createTestApp(overrides);
    const { app } = testApp;

    // Step 1: Hit /auth/login to set up pending state
    const loginRes = await supertest(app).get('/auth/login').expect(302);
    const url = new URL(loginRes.headers.location);
    const state = url.searchParams.get('state');

    // Step 2: Hit /auth/callback with the state and a mock code
    const callbackRes = await supertest(app)
      .get('/auth/callback')
      .query({ code: 'mock-auth-code', state });

    return { ...testApp, callbackRes, state };
  }

  test('returns 200 on successful auth flow', async () => {
    const { callbackRes } = await performAuthFlow();
    expect(callbackRes.status).toBe(200);
  });

  test('exchanges code for tokens via Entra token endpoint', async () => {
    const mockFetch = createMockFetch();
    const { fetch } = await performAuthFlow({ fetch: mockFetch });

    // First call should be to the token endpoint
    const tokenCall = fetch.mock.calls.find(([url]) => url.includes('oauth2/v2.0/token'));
    expect(tokenCall).toBeDefined();
    expect(tokenCall[1].method).toBe('POST');

    // Verify the body includes required params
    const body = tokenCall[1].body;
    expect(body).toContain('client_id=test-client-id');
    expect(body).toContain('grant_type=authorization_code');
    expect(body).toContain('code=mock-auth-code');
    expect(body).toContain('code_verifier=');
  });

  test('fetches user profile from Graph /me', async () => {
    const mockFetch = createMockFetch();
    const { fetch } = await performAuthFlow({ fetch: mockFetch });

    const meCall = fetch.mock.calls.find(([url]) => url.includes('graph.microsoft.com'));
    expect(meCall).toBeDefined();
    expect(meCall[1].headers.Authorization).toBe('Bearer mock-access-token');
  });

  test('stores tokens via tokenStorage.setTokensForUser', async () => {
    const { tokenStorage } = await performAuthFlow();

    expect(tokenStorage.setTokensForUser).toHaveBeenCalledTimes(1);
    expect(tokenStorage.setTokensForUser).toHaveBeenCalledWith('user-oid-123', {
      accessToken: 'mock-access-token',
      refreshToken: 'mock-refresh-token',
      expiresIn: 3600,
      scopes: 'offline_access Mail.Read User.Read',
      email: 'user@example.com',
      name: 'Test User',
    });
  });

  test('creates a session via sessionStore.createSession', async () => {
    const { sessionStore } = await performAuthFlow();

    expect(sessionStore.createSession).toHaveBeenCalledTimes(1);
    expect(sessionStore.createSession).toHaveBeenCalledWith('user-oid-123');
  });

  test('returns HTML containing the session token', async () => {
    const { callbackRes } = await performAuthFlow();

    expect(callbackRes.text).toContain('mock-session-token-uuid');
  });

  test('returns HTML containing the MCP config snippet', async () => {
    const { callbackRes } = await performAuthFlow();

    // The config is rendered inside HTML, so quotes are entity-escaped
    expect(callbackRes.text).toContain('&quot;mcpServers&quot;');
    expect(callbackRes.text).toContain('/mcp');
    expect(callbackRes.text).toContain('Bearer mock-session-token-uuid');
    // The raw JSON is also embedded in the <script> for clipboard copy
    expect(callbackRes.text).toContain('copyConfig');
  });

  test('returns HTML containing user display name', async () => {
    const { callbackRes } = await performAuthFlow();

    expect(callbackRes.text).toContain('Test User');
  });

  test('uses trusted configured base URL, not Host header, in rendered config', async () => {
    const testApp = createTestApp();
    const { app } = testApp;

    const loginRes = await supertest(app).get('/auth/login').expect(302);
    const url = new URL(loginRes.headers.location);
    const state = url.searchParams.get('state');

    const callbackRes = await supertest(app)
      .get('/auth/callback')
      .set('Host', 'attacker.example.com')
      .query({ code: 'mock-auth-code', state })
      .expect(200);

    expect(callbackRes.text).toContain('https://outlook-mcp.example.com/mcp');
    expect(callbackRes.text).not.toContain('attacker.example.com');
  });

  test('state is single-use — second callback with same state returns 400', async () => {
    const testApp = createTestApp();
    const { app } = testApp;

    // Login to get state
    const loginRes = await supertest(app).get('/auth/login').expect(302);
    const url = new URL(loginRes.headers.location);
    const state = url.searchParams.get('state');

    // First callback succeeds
    await supertest(app)
      .get('/auth/callback')
      .query({ code: 'mock-auth-code', state })
      .expect(200);

    // Second callback with same state fails
    await supertest(app)
      .get('/auth/callback')
      .query({ code: 'mock-auth-code', state })
      .expect(400);
  });

  test('callback succeeds when pending map entry is missing but signed state is valid', async () => {
    const mockFetch = createMockFetch();
    const testApp = createTestApp({ fetch: mockFetch });
    const { app, fetch } = testApp;

    const loginRes = await supertest(app).get('/auth/login').expect(302);
    const loginUrl = new URL(loginRes.headers.location);
    const state = loginUrl.searchParams.get('state');
    const challenge = loginUrl.searchParams.get('code_challenge');

    // Simulate restart/another instance by clearing in-memory pending entries.
    _pendingAuth.clear();

    await supertest(app)
      .get('/auth/callback')
      .query({ code: 'mock-auth-code', state })
      .expect(200);

    const tokenCall = fetch.mock.calls.find(([url]) => url.includes('oauth2/v2.0/token'));
    const tokenBody = new URLSearchParams(tokenCall[1].body);
    const codeVerifier = tokenBody.get('code_verifier');
    const expectedChallenge = crypto.createHash('sha256').update(codeVerifier).digest('base64url');

    expect(expectedChallenge).toBe(challenge);
  });
});

// ── GET /auth/callback — error paths ────────────────────────────────────

describe('GET /auth/callback — errors', () => {
  test('returns 400 when state is missing', async () => {
    const { app } = createTestApp();

    const res = await supertest(app)
      .get('/auth/callback')
      .query({ code: 'some-code' })
      .expect(400);

    expect(res.text).toContain('Missing state parameter');
  });

  test('returns 400 when code is missing', async () => {
    const { app } = createTestApp();

    const res = await supertest(app)
      .get('/auth/callback')
      .query({ state: 'some-state' })
      .expect(400);

    expect(res.text).toContain('Missing authorization code');
  });

  test('returns 400 when state does not match any pending auth', async () => {
    const { app } = createTestApp();

    const res = await supertest(app)
      .get('/auth/callback')
      .query({ code: 'some-code', state: 'unknown-state' })
      .expect(400);

    expect(res.text).toContain('Invalid or expired state');
  });

  test('returns 400 when state has expired', async () => {
    const { app } = createTestApp();

    // Manually insert an expired pending entry
    const expiredState = 'expired-state-value';
    _pendingAuth.set(expiredState, {
      codeVerifier: 'some-verifier',
      expiresAt: Date.now() - 1000, // expired 1 second ago
    });

    const res = await supertest(app)
      .get('/auth/callback')
      .query({ code: 'some-code', state: expiredState })
      .expect(400);

    expect(res.text).toContain('Invalid or expired state');

    // Verify it was cleaned up
    expect(_pendingAuth.has(expiredState)).toBe(false);
  });

  test('returns 500 when token exchange fails', async () => {
    const mockFetch = createMockFetch({ tokenStatus: 400 });
    const { app } = createTestApp({ fetch: mockFetch });

    // Set up valid pending state
    const loginRes = await supertest(app).get('/auth/login').expect(302);
    const url = new URL(loginRes.headers.location);
    const state = url.searchParams.get('state');

    const res = await supertest(app)
      .get('/auth/callback')
      .query({ code: 'bad-code', state })
      .expect(500);

    expect(res.text).toContain('Failed to exchange');
  });

  test('returns 500 when Graph /me request fails', async () => {
    const mockFetch = createMockFetch({ meStatus: 401 });
    const { app } = createTestApp({ fetch: mockFetch });

    const loginRes = await supertest(app).get('/auth/login').expect(302);
    const url = new URL(loginRes.headers.location);
    const state = url.searchParams.get('state');

    const res = await supertest(app)
      .get('/auth/callback')
      .query({ code: 'auth-code', state })
      .expect(500);

    expect(res.text).toContain('Failed to retrieve user profile');
  });
});

// ── Pending auth state cleanup ──────────────────────────────────────────

describe('pending auth state cleanup', () => {
  test('expired entries are removed on next /auth/login call', async () => {
    const { app } = createTestApp();

    // Insert some expired entries manually
    _pendingAuth.set('old-state-1', { codeVerifier: 'v1', expiresAt: Date.now() - 60000 });
    _pendingAuth.set('old-state-2', { codeVerifier: 'v2', expiresAt: Date.now() - 30000 });

    // A valid non-expired entry
    _pendingAuth.set('fresh-state', { codeVerifier: 'v3', expiresAt: Date.now() + 600000 });

    expect(_pendingAuth.size).toBe(3);

    // Hit /auth/login which triggers cleanup
    await supertest(app).get('/auth/login').expect(302);

    // The expired entries should be gone; fresh + newly created should remain
    expect(_pendingAuth.has('old-state-1')).toBe(false);
    expect(_pendingAuth.has('old-state-2')).toBe(false);
    expect(_pendingAuth.has('fresh-state')).toBe(true);
    // Plus the new one from this login request
    expect(_pendingAuth.size).toBe(2);
  });
});
