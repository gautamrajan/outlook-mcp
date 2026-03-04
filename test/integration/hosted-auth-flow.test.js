/**
 * Integration tests for the dual-mode authentication flow.
 *
 * Tests the full auth flow end-to-end using:
 *   - REAL: auth/request-context.js (AsyncLocalStorage)
 *   - REAL: auth/per-user-token-storage.js (in-memory cache)
 *   - REAL: auth/index.js (ensureAuthenticated — the system under test)
 *   - MOCKED: auth/obo-exchange.js (exchangeOBO — external HTTP calls)
 *   - MOCKED: auth/token-storage.js (TokenStorage — file I/O)
 *
 * Because auth/index.js creates singletons at require time, we use
 * jest.resetModules() in beforeEach and re-require everything fresh.
 */

let ensureAuthenticated;
let requestContext;
let perUserTokenStorage; // the real singleton instance from auth/index.js
let mockExchangeOBO;
let mockTokenStorageInstance;

beforeEach(() => {
  jest.resetModules();
  jest.clearAllMocks();

  // ── Mocks ──────────────────────────────────────────────────────────

  // Mock config
  jest.mock('../../config', () => ({
    AUTH_CONFIG: {
      tokenStorePath: '/tmp/test-tokens.json',
      clientId: 'integration-client-id',
      clientSecret: 'integration-client-secret',
      tenantId: 'integration-tenant-id',
      tokenEndpoint: 'https://login.microsoftonline.com/integration-tenant-id/oauth2/v2.0/token',
      redirectUri: 'http://localhost:3333/auth/callback',
      scopes: ['offline_access', 'Mail.Read', 'Mail.ReadWrite', 'User.Read', 'Calendars.Read'],
    },
  }));

  // Mock auth tools (avoid pulling in embedded-server, etc.)
  jest.mock('../../auth/tools', () => ({
    authTools: [{ name: 'mock-tool' }],
  }));

  // Mock TokenStorage — local mode uses this
  mockTokenStorageInstance = {
    getValidAccessToken: jest.fn(),
    invalidateAccessToken: jest.fn(),
  };
  jest.mock('../../auth/token-storage', () => {
    return jest.fn().mockImplementation(() => mockTokenStorageInstance);
  });

  // Mock OBO exchange — hosted mode uses this
  mockExchangeOBO = jest.fn();
  jest.mock('../../auth/obo-exchange', () => ({
    exchangeOBO: mockExchangeOBO,
  }));

  // DO NOT mock request-context or per-user-token-storage — use the real ones

  // ── Require fresh ──────────────────────────────────────────────────

  const authModule = require('../../auth/index');
  ensureAuthenticated = authModule.ensureAuthenticated;

  // Get the real request-context instance (same one auth/index.js uses)
  const rc = require('../../auth/request-context');
  requestContext = rc.requestContext;

  // Get the real PerUserTokenStorage singleton.
  // auth/index.js creates it internally; we can observe its effects
  // through ensureAuthenticated's behaviour.
  // We also get a direct reference via the module's internal import
  // to verify caching behaviour.
  const PerUserTokenStorage = require('../../auth/per-user-token-storage');
  // The singleton is created inside auth/index.js, so we access it indirectly.
  // For assertion purposes, we'll rely on observing OBO call counts.
});

// ── Helpers ──────────────────────────────────────────────────────────

function makeOBOResponse(suffix = '') {
  return {
    access_token: `graph-token${suffix}`,
    refresh_token: `graph-refresh${suffix}`,
    expires_in: 3600,
    scope: 'Mail.Read Mail.ReadWrite User.Read Calendars.Read',
    token_type: 'Bearer',
  };
}

// ── Hosted Mode (inside requestContext.run) ──────────────────────────

describe('Hosted mode — end-to-end auth flow', () => {
  const userCtx = { userId: 'user-hosted-001', entraToken: 'entra-jwt-hosted-001' };

  test('first call triggers OBO exchange and returns Graph token', async () => {
    mockExchangeOBO.mockResolvedValue(makeOBOResponse());

    const token = await requestContext.run(userCtx, async () => {
      return ensureAuthenticated();
    });

    expect(token).toBe('graph-token');
    expect(mockExchangeOBO).toHaveBeenCalledTimes(1);
    expect(mockExchangeOBO).toHaveBeenCalledWith(
      'entra-jwt-hosted-001',
      expect.objectContaining({
        clientId: 'integration-client-id',
        clientSecret: 'integration-client-secret',
        tenantId: 'integration-tenant-id',
      })
    );
  });

  test('second call returns cached token without calling OBO again', async () => {
    mockExchangeOBO.mockResolvedValue(makeOBOResponse());

    // First call — populates cache
    await requestContext.run(userCtx, async () => {
      return ensureAuthenticated();
    });

    // Second call — should use cache
    const token = await requestContext.run(userCtx, async () => {
      return ensureAuthenticated();
    });

    expect(token).toBe('graph-token');
    // OBO should only have been called once (during the first call)
    expect(mockExchangeOBO).toHaveBeenCalledTimes(1);
  });

  test('forceRefresh invalidates cache and triggers fresh OBO exchange', async () => {
    mockExchangeOBO
      .mockResolvedValueOnce(makeOBOResponse('-v1'))
      .mockResolvedValueOnce(makeOBOResponse('-v2'));

    // First call — populates cache
    const token1 = await requestContext.run(userCtx, async () => {
      return ensureAuthenticated();
    });
    expect(token1).toBe('graph-token-v1');

    // Force refresh — should invalidate and call OBO again
    const token2 = await requestContext.run(userCtx, async () => {
      return ensureAuthenticated({ forceRefresh: true });
    });
    expect(token2).toBe('graph-token-v2');
    expect(mockExchangeOBO).toHaveBeenCalledTimes(2);
  });

  test('two different users in concurrent contexts get independent tokens', async () => {
    const userA = { userId: 'user-A', entraToken: 'entra-A' };
    const userB = { userId: 'user-B', entraToken: 'entra-B' };

    mockExchangeOBO
      .mockImplementation(async (entraToken) => {
        // Return different tokens based on the input entra token
        if (entraToken === 'entra-A') {
          return makeOBOResponse('-A');
        }
        return makeOBOResponse('-B');
      });

    // Run both in parallel
    const [tokenA, tokenB] = await Promise.all([
      requestContext.run(userA, async () => {
        // Yield to allow interleaving
        await new Promise((resolve) => setImmediate(resolve));
        return ensureAuthenticated();
      }),
      requestContext.run(userB, async () => {
        await new Promise((resolve) => setImmediate(resolve));
        return ensureAuthenticated();
      }),
    ]);

    expect(tokenA).toBe('graph-token-A');
    expect(tokenB).toBe('graph-token-B');
    expect(mockExchangeOBO).toHaveBeenCalledTimes(2);

    // Verify each OBO call used the correct entra token
    const oboCallTokens = mockExchangeOBO.mock.calls.map((c) => c[0]);
    expect(oboCallTokens).toContain('entra-A');
    expect(oboCallTokens).toContain('entra-B');
  });

  test('user A cached token is not returned to user B', async () => {
    const userA = { userId: 'user-A', entraToken: 'entra-A' };
    const userB = { userId: 'user-B', entraToken: 'entra-B' };

    mockExchangeOBO
      .mockResolvedValueOnce(makeOBOResponse('-A'))
      .mockResolvedValueOnce(makeOBOResponse('-B'));

    // Populate cache for user A
    await requestContext.run(userA, () => ensureAuthenticated());

    // User B should trigger a new OBO (not get user A's cached token)
    const tokenB = await requestContext.run(userB, () => ensureAuthenticated());

    expect(tokenB).toBe('graph-token-B');
    expect(mockExchangeOBO).toHaveBeenCalledTimes(2);
  });

  test('OBO exchange failure propagates as error', async () => {
    mockExchangeOBO.mockRejectedValue(
      new Error('OBO exchange failed: invalid or expired user token')
    );

    await expect(
      requestContext.run(userCtx, () => ensureAuthenticated())
    ).rejects.toThrow('OBO exchange failed: invalid or expired user token');
  });

  test('hosted mode does NOT interact with singleton TokenStorage', async () => {
    mockExchangeOBO.mockResolvedValue(makeOBOResponse());

    await requestContext.run(userCtx, () => ensureAuthenticated());

    expect(mockTokenStorageInstance.getValidAccessToken).not.toHaveBeenCalled();
    expect(mockTokenStorageInstance.invalidateAccessToken).not.toHaveBeenCalled();
  });

  test('offline_access is filtered from OBO scopes', async () => {
    mockExchangeOBO.mockResolvedValue(makeOBOResponse());

    await requestContext.run(userCtx, () => ensureAuthenticated());

    const oboConfig = mockExchangeOBO.mock.calls[0][1];
    expect(oboConfig.scopes).not.toContain('offline_access');
    expect(oboConfig.scopes).toContain('Mail.Read');
    expect(oboConfig.scopes).toContain('User.Read');
  });
});

// ── Local Mode (outside requestContext.run) ──────────────────────────

describe('Local mode — existing single-user flow', () => {
  test('uses TokenStorage to get access token', async () => {
    mockTokenStorageInstance.getValidAccessToken.mockResolvedValue('local-access-token');

    const token = await ensureAuthenticated();

    expect(token).toBe('local-access-token');
    expect(mockTokenStorageInstance.getValidAccessToken).toHaveBeenCalledTimes(1);
  });

  test('throws when TokenStorage returns null', async () => {
    mockTokenStorageInstance.getValidAccessToken.mockResolvedValue(null);

    await expect(ensureAuthenticated()).rejects.toThrow('Authentication required');
  });

  test('forceRefresh invalidates then re-fetches via TokenStorage', async () => {
    mockTokenStorageInstance.getValidAccessToken.mockResolvedValue('refreshed-local-token');

    const token = await ensureAuthenticated({ forceRefresh: true });

    expect(mockTokenStorageInstance.invalidateAccessToken).toHaveBeenCalledTimes(1);
    expect(token).toBe('refreshed-local-token');
  });

  test('does NOT interact with per-user storage or OBO in local mode', async () => {
    mockTokenStorageInstance.getValidAccessToken.mockResolvedValue('local-token');

    await ensureAuthenticated();

    expect(mockExchangeOBO).not.toHaveBeenCalled();
  });

  test('does NOT touch OBO even with forceRefresh in local mode', async () => {
    mockTokenStorageInstance.getValidAccessToken.mockResolvedValue('local-token');

    await ensureAuthenticated({ forceRefresh: true });

    expect(mockExchangeOBO).not.toHaveBeenCalled();
  });
});

// ── Mode isolation ───────────────────────────────────────────────────

describe('Mode isolation', () => {
  test('local call followed by hosted call uses correct path each time', async () => {
    const userCtx = { userId: 'user-switch', entraToken: 'entra-switch' };

    // Local call
    mockTokenStorageInstance.getValidAccessToken.mockResolvedValue('local-token');
    const localToken = await ensureAuthenticated();
    expect(localToken).toBe('local-token');

    // Hosted call
    mockExchangeOBO.mockResolvedValue(makeOBOResponse('-hosted'));
    const hostedToken = await requestContext.run(userCtx, () => ensureAuthenticated());
    expect(hostedToken).toBe('graph-token-hosted');

    // Verify both paths were used
    expect(mockTokenStorageInstance.getValidAccessToken).toHaveBeenCalledTimes(1);
    expect(mockExchangeOBO).toHaveBeenCalledTimes(1);
  });

  test('hosted call followed by local call uses correct path each time', async () => {
    const userCtx = { userId: 'user-switch-2', entraToken: 'entra-switch-2' };

    // Hosted call
    mockExchangeOBO.mockResolvedValue(makeOBOResponse('-hosted'));
    const hostedToken = await requestContext.run(userCtx, () => ensureAuthenticated());
    expect(hostedToken).toBe('graph-token-hosted');

    // Local call
    mockTokenStorageInstance.getValidAccessToken.mockResolvedValue('local-token');
    const localToken = await ensureAuthenticated();
    expect(localToken).toBe('local-token');

    // Verify both paths were used
    expect(mockExchangeOBO).toHaveBeenCalledTimes(1);
    expect(mockTokenStorageInstance.getValidAccessToken).toHaveBeenCalledTimes(1);
  });
});
