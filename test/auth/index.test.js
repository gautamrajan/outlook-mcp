/**
 * Tests for auth/index.js — ensureAuthenticated()
 *
 * Covers both local mode (existing singleton TokenStorage flow)
 * and hosted mode (per-user token lookup with silent refresh).
 *
 * Because auth/index.js creates singletons at require time, we use
 * jest.resetModules() in beforeEach and re-require everything fresh.
 * This ensures mock constructors are wired up before the module runs.
 */

const PerUserTokenStorage = require('../../auth/per-user-token-storage');

let ensureAuthenticated;
let tokenStorage;
let authTools;
let getHostedTokenStorage;
let setHostedTokenStorage;
let requestContext;

// Mock instances (recreated each test)
let mockTokenStorageInstance;

// Save original fetch so we can restore it
const originalFetch = global.fetch;

beforeEach(() => {
  jest.resetModules();
  jest.clearAllMocks();

  // Reset global fetch mock
  global.fetch = jest.fn();

  // Static mocks that don't need instance tracking
  jest.mock('../../auth/tools', () => ({
    authTools: [{ name: 'mock-tool' }],
  }));

  jest.mock('../../config', () => ({
    AUTH_CONFIG: {
      tokenStorePath: '/tmp/test-tokens.json',
      clientId: 'test-client-id',
      clientSecret: 'test-client-secret',
      tenantId: 'test-tenant-id',
      tokenEndpoint: 'https://login.microsoftonline.com/test-tenant-id/oauth2/v2.0/token',
      redirectUri: 'http://localhost:3333/auth/callback',
      scopes: ['offline_access', 'Mail.Read', 'Mail.ReadWrite', 'User.Read', 'Calendars.Read'],
    },
    CONNECTOR_AUTH: {
      oboScopes: 'Mail.Read User.Read',
    },
  }));

  // Set up TokenStorage mock constructor
  mockTokenStorageInstance = {
    getValidAccessToken: jest.fn(),
    invalidateAccessToken: jest.fn(),
  };
  jest.mock('../../auth/token-storage', () => {
    return jest.fn().mockImplementation(() => mockTokenStorageInstance);
  });

  // DO NOT mock request-context — it must be the real AsyncLocalStorage
  // so that requestContext.run() in tests sets context visible to auth/index.js

  // Now require everything fresh — auth/index.js will use the mocked constructors
  // and the SAME request-context instance we get here
  const authModule = require('../../auth/index');
  ensureAuthenticated = authModule.ensureAuthenticated;
  tokenStorage = authModule.tokenStorage;
  authTools = authModule.authTools;
  getHostedTokenStorage = authModule.getHostedTokenStorage;
  setHostedTokenStorage = authModule.setHostedTokenStorage;

  // Get the same request-context instance that auth/index.js is using
  const rc = require('../../auth/request-context');
  requestContext = rc.requestContext;
});

afterEach(() => {
  global.fetch = originalFetch;
});

describe('auth/index.js', () => {
  // ── Exports ──────────────────────────────────────────────────────────

  describe('module exports', () => {
    test('should export tokenStorage singleton', () => {
      expect(tokenStorage).toBeDefined();
      expect(tokenStorage).toBe(mockTokenStorageInstance);
    });

    test('should export authTools', () => {
      expect(authTools).toBeDefined();
      expect(Array.isArray(authTools)).toBe(true);
    });

    test('should export ensureAuthenticated function', () => {
      expect(ensureAuthenticated).toBeDefined();
      expect(typeof ensureAuthenticated).toBe('function');
    });

    test('should export getHostedTokenStorage function', () => {
      expect(getHostedTokenStorage).toBeDefined();
      expect(typeof getHostedTokenStorage).toBe('function');
    });

    test('should export setHostedTokenStorage function', () => {
      expect(setHostedTokenStorage).toBeDefined();
      expect(typeof setHostedTokenStorage).toBe('function');
    });
  });

  // ── Local Mode (no user context) ────────────────────────────────────

  describe('local mode (no user context)', () => {
    test('should return access token from TokenStorage', async () => {
      mockTokenStorageInstance.getValidAccessToken.mockResolvedValue('local-access-token');

      const token = await ensureAuthenticated();

      expect(token).toBe('local-access-token');
      expect(mockTokenStorageInstance.getValidAccessToken).toHaveBeenCalledTimes(1);
    });

    test('should throw "Authentication required" when TokenStorage returns null', async () => {
      mockTokenStorageInstance.getValidAccessToken.mockResolvedValue(null);

      await expect(ensureAuthenticated()).rejects.toThrow('Authentication required');
    });

    test('should call invalidateAccessToken when forceRefresh is true', async () => {
      mockTokenStorageInstance.getValidAccessToken.mockResolvedValue('refreshed-token');

      const token = await ensureAuthenticated({ forceRefresh: true });

      expect(mockTokenStorageInstance.invalidateAccessToken).toHaveBeenCalledTimes(1);
      expect(token).toBe('refreshed-token');
    });

    test('should NOT call invalidateAccessToken when forceRefresh is false', async () => {
      mockTokenStorageInstance.getValidAccessToken.mockResolvedValue('normal-token');

      await ensureAuthenticated({ forceRefresh: false });

      expect(mockTokenStorageInstance.invalidateAccessToken).not.toHaveBeenCalled();
    });

    test('should NOT call invalidateAccessToken when forceRefresh is not provided', async () => {
      mockTokenStorageInstance.getValidAccessToken.mockResolvedValue('normal-token');

      await ensureAuthenticated();

      expect(mockTokenStorageInstance.invalidateAccessToken).not.toHaveBeenCalled();
    });

    test('should NOT interact with hosted mode path in local mode', async () => {
      mockTokenStorageInstance.getValidAccessToken.mockResolvedValue('local-token');

      const token = await ensureAuthenticated();

      expect(token).toBe('local-token');
      // fetch should never be called for local mode
      expect(global.fetch).not.toHaveBeenCalled();
    });
  });

  // ── Hosted Mode (user context present) ──────────────────────────────

  describe('hosted mode (user context present)', () => {
    const TEST_USER_ID = 'test-user-123';
    const testUserCtx = { userId: TEST_USER_ID, entraToken: 'entra-jwt-token' };

    /** @type {PerUserTokenStorage} */
    let hostedStorage;

    /**
     * Helper: pre-populate the hosted storage with token data for the test user.
     */
    async function seedTokens({ accessToken, refreshToken, expiresIn, email, name } = {}) {
      await hostedStorage.setTokensForUser(TEST_USER_ID, {
        accessToken: accessToken || 'stored-access-token',
        refreshToken: refreshToken || 'stored-refresh-token',
        expiresIn: expiresIn != null ? expiresIn : 3600,
        scopes: 'Mail.Read Mail.ReadWrite User.Read Calendars.Read',
        email: email || 'test@example.com',
        name: name || 'Test User',
      });
    }

    /**
     * Helper: mock a successful token refresh response from Entra.
     */
    function mockSuccessfulRefresh(overrides = {}) {
      global.fetch.mockResolvedValueOnce({
        ok: true,
        json: async () => ({
          access_token: overrides.access_token || 'new-access-token',
          refresh_token: overrides.refresh_token || 'new-refresh-token',
          expires_in: overrides.expires_in || 3600,
          scope: overrides.scope || 'Mail.Read Mail.ReadWrite User.Read Calendars.Read',
        }),
      });
    }

    /**
     * Helper: mock a failed token refresh response from Entra.
     */
    function mockFailedRefresh(status = 400) {
      global.fetch.mockResolvedValueOnce({
        ok: false,
        status,
        json: async () => ({ error: 'invalid_grant', error_description: 'Refresh token revoked' }),
      });
    }

    beforeEach(() => {
      // Create a fresh in-memory PerUserTokenStorage (no file path = no disk I/O)
      hostedStorage = new PerUserTokenStorage();

      // Inject it into auth/index.js so _ensureAuthenticatedHosted uses it
      setHostedTokenStorage(hostedStorage);
    });

    // ── 1. Valid stored token ──────────────────────────────────────────

    test('should return valid stored token directly without network call', async () => {
      await seedTokens();

      const token = await requestContext.run(testUserCtx, async () => {
        return ensureAuthenticated();
      });

      expect(token).toBe('stored-access-token');
      expect(global.fetch).not.toHaveBeenCalled();
    });

    // ── 2. Expired token with valid refresh ────────────────────────────

    test('should refresh expired token and return new access token', async () => {
      // Seed with already-expired token (expiresIn = 0 means already expired)
      await seedTokens({ expiresIn: 0 });

      mockSuccessfulRefresh({ access_token: 'refreshed-access-token' });

      const token = await requestContext.run(testUserCtx, async () => {
        return ensureAuthenticated();
      });

      expect(token).toBe('refreshed-access-token');
      expect(global.fetch).toHaveBeenCalledTimes(1);

      // Verify the fetch was called with correct parameters
      const [url, opts] = global.fetch.mock.calls[0];
      expect(url).toBe('https://login.microsoftonline.com/test-tenant-id/oauth2/v2.0/token');
      expect(opts.method).toBe('POST');
      expect(opts.body).toContain('grant_type=refresh_token');
      expect(opts.body).toContain('client_id=test-client-id');
      expect(opts.body).toContain('client_secret=test-client-secret');
      expect(opts.body).toContain('refresh_token=stored-refresh-token');
      // offline_access should be excluded from scope
      expect(opts.body).not.toContain('offline_access');
    });

    test('should store new tokens after successful refresh', async () => {
      await seedTokens({ expiresIn: 0 });

      mockSuccessfulRefresh({
        access_token: 'refreshed-access-token',
        refresh_token: 'rotated-refresh-token',
      });

      await requestContext.run(testUserCtx, async () => {
        return ensureAuthenticated();
      });

      // Verify the new tokens are stored
      const storedToken = hostedStorage.getTokenForUser(TEST_USER_ID);
      expect(storedToken).toBe('refreshed-access-token');

      const storedRefresh = hostedStorage.getRefreshToken(TEST_USER_ID);
      expect(storedRefresh).toBe('rotated-refresh-token');
    });

    test('should preserve original refresh token if Entra does not rotate it', async () => {
      await seedTokens({ expiresIn: 0, refreshToken: 'original-refresh' });

      // Entra response without refresh_token field
      global.fetch.mockResolvedValueOnce({
        ok: true,
        json: async () => ({
          access_token: 'new-access',
          expires_in: 3600,
          scope: 'Mail.Read',
          // no refresh_token — Entra did not rotate
        }),
      });

      await requestContext.run(testUserCtx, async () => {
        return ensureAuthenticated();
      });

      const storedRefresh = hostedStorage.getRefreshToken(TEST_USER_ID);
      expect(storedRefresh).toBe('original-refresh');
    });

    test('should preserve user info (email, name) after refresh', async () => {
      await seedTokens({ expiresIn: 0, email: 'user@corp.com', name: 'Jane Doe' });

      mockSuccessfulRefresh();

      await requestContext.run(testUserCtx, async () => {
        return ensureAuthenticated();
      });

      const info = hostedStorage.getUserInfo(TEST_USER_ID);
      expect(info.email).toBe('user@corp.com');
      expect(info.name).toBe('Jane Doe');
    });

    // ── 3. Expired token with failed refresh ───────────────────────────

    test('should throw AUTH_REQUIRED when refresh fails', async () => {
      await seedTokens({ expiresIn: 0 });

      mockFailedRefresh(400);

      let caughtErr;
      try {
        await requestContext.run(testUserCtx, async () => {
          return ensureAuthenticated();
        });
      } catch (err) {
        caughtErr = err;
      }

      expect(caughtErr).toBeDefined();
      expect(caughtErr.message).toMatch(/Token refresh failed/);
      expect(caughtErr.code).toBe('AUTH_REQUIRED');
    });

    // ── 4. No tokens at all for user ───────────────────────────────────

    test('should throw AUTH_REQUIRED when no tokens exist for user', async () => {
      // Don't seed any tokens — storage is empty

      const promise = requestContext.run(testUserCtx, async () => {
        return ensureAuthenticated();
      });

      await expect(promise).rejects.toThrow('Authentication required');

      try {
        await requestContext.run(testUserCtx, async () => {
          return ensureAuthenticated();
        });
      } catch (err) {
        expect(err.code).toBe('AUTH_REQUIRED');
      }
    });

    // ── 5. No user context ─────────────────────────────────────────────

    test('should fall back to local mode when context has no userId', async () => {
      const badCtx = { userId: null, entraToken: 'jwt' };

      await expect(
        requestContext.run(badCtx, async () => {
          return ensureAuthenticated();
        })
      ).rejects.toThrow('Authentication required');
    });

    test('should fail closed for malformed connector context instead of using local token storage', async () => {
      mockTokenStorageInstance.getValidAccessToken.mockResolvedValue('local-token-that-must-not-be-used');

      const malformedConnectorCtx = {
        authMethod: 'connector',
        entraToken: 'entra-jwt-token',
      };

      let caughtErr;
      try {
        await requestContext.run(malformedConnectorCtx, async () => {
          return ensureAuthenticated();
        });
      } catch (err) {
        caughtErr = err;
      }

      expect(caughtErr).toBeDefined();
      expect(caughtErr.code).toBe('AUTH_REQUIRED');
      expect(caughtErr.message).toMatch(/missing userId/i);
      expect(mockTokenStorageInstance.getValidAccessToken).not.toHaveBeenCalled();
    });

    // ── 6. forceRefresh=true ───────────────────────────────────────────

    test('should skip cached token and refresh when forceRefresh is true', async () => {
      // Seed with a perfectly valid (non-expired) token
      await seedTokens({ expiresIn: 3600 });

      mockSuccessfulRefresh({ access_token: 'force-refreshed-token' });

      const token = await requestContext.run(testUserCtx, async () => {
        return ensureAuthenticated({ forceRefresh: true });
      });

      // Should NOT have returned the cached token
      expect(token).toBe('force-refreshed-token');
      expect(global.fetch).toHaveBeenCalledTimes(1);
    });

    // ── 7. Concurrent refresh ──────────────────────────────────────────

    test('should only make one refresh request for concurrent calls by same user', async () => {
      await seedTokens({ expiresIn: 0 });

      // Make fetch take a moment so both calls overlap
      global.fetch.mockImplementation(() =>
        new Promise(resolve =>
          setTimeout(() => resolve({
            ok: true,
            json: async () => ({
              access_token: 'concurrent-token',
              refresh_token: 'concurrent-refresh',
              expires_in: 3600,
              scope: 'Mail.Read',
            }),
          }), 50)
        )
      );

      const results = await requestContext.run(testUserCtx, async () => {
        return Promise.all([
          ensureAuthenticated(),
          ensureAuthenticated(),
        ]);
      });

      // Both should get the token
      expect(results[0]).toBe('concurrent-token');
      expect(results[1]).toBe('concurrent-token');

      // Only one fetch call should have been made
      expect(global.fetch).toHaveBeenCalledTimes(1);
    });

    // ── 8. Local mode completely unaffected ─────────────────────────────

    test('should NOT interact with the singleton TokenStorage in hosted mode', async () => {
      await seedTokens();

      await requestContext.run(testUserCtx, async () => {
        await ensureAuthenticated();
      });

      expect(mockTokenStorageInstance.getValidAccessToken).not.toHaveBeenCalled();
      expect(mockTokenStorageInstance.invalidateAccessToken).not.toHaveBeenCalled();
    });

    // ── setHostedTokenStorage / getHostedTokenStorage ──────────────────

    test('setHostedTokenStorage should allow injecting a custom storage instance', () => {
      const customStorage = new PerUserTokenStorage();
      setHostedTokenStorage(customStorage);

      expect(getHostedTokenStorage()).toBe(customStorage);
    });

    test('getHostedTokenStorage should create a default instance when none injected', () => {
      // Reset hosted storage to null
      setHostedTokenStorage(null);

      // Lazy initialization should create a new PerUserTokenStorage
      const storage = getHostedTokenStorage();
      expect(storage).toBeDefined();
      // Verify it quacks like a PerUserTokenStorage (instanceof won't work across
      // jest.resetModules() boundaries because two copies of the class exist)
      expect(typeof storage.getTokenForUser).toBe('function');
      expect(typeof storage.getRefreshToken).toBe('function');
      expect(typeof storage.setTokensForUser).toBe('function');
    });
  });
});
