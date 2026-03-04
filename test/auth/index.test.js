/**
 * Tests for auth/index.js — ensureAuthenticated()
 *
 * Covers both local mode (existing singleton TokenStorage flow)
 * and hosted mode (OBO + per-user token storage).
 *
 * Because auth/index.js creates singletons at require time, we use
 * jest.resetModules() in beforeEach and re-require everything fresh.
 * This ensures mock constructors are wired up before the module runs.
 */

let ensureAuthenticated;
let tokenStorage;
let authTools;
let requestContext;
let exchangeOBO;

// Mock instances (recreated each test)
let mockTokenStorageInstance;
let mockPerUserTokenStorageInstance;

beforeEach(() => {
  jest.resetModules();
  jest.clearAllMocks();

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
  }));

  // Set up TokenStorage mock constructor
  mockTokenStorageInstance = {
    getValidAccessToken: jest.fn(),
    invalidateAccessToken: jest.fn(),
  };
  jest.mock('../../auth/token-storage', () => {
    return jest.fn().mockImplementation(() => mockTokenStorageInstance);
  });

  // Set up PerUserTokenStorage mock constructor
  mockPerUserTokenStorageInstance = {
    getTokenForUser: jest.fn(),
    setTokenForUser: jest.fn(),
    invalidateUser: jest.fn(),
  };
  jest.mock('../../auth/per-user-token-storage', () => {
    return jest.fn().mockImplementation(() => mockPerUserTokenStorageInstance);
  });

  // Mock OBO exchange
  jest.mock('../../auth/obo-exchange', () => ({
    exchangeOBO: jest.fn(),
  }));

  // DO NOT mock request-context — it must be the real AsyncLocalStorage
  // so that requestContext.run() in tests sets context visible to auth/index.js

  // Now require everything fresh — auth/index.js will use the mocked constructors
  // and the SAME request-context instance we get here
  const authModule = require('../../auth/index');
  ensureAuthenticated = authModule.ensureAuthenticated;
  tokenStorage = authModule.tokenStorage;
  authTools = authModule.authTools;

  // Get the same request-context instance that auth/index.js is using
  const rc = require('../../auth/request-context');
  requestContext = rc.requestContext;

  // Get the mocked exchangeOBO
  const obo = require('../../auth/obo-exchange');
  exchangeOBO = obo.exchangeOBO;
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

    test('should NOT interact with per-user token storage in local mode', async () => {
      mockTokenStorageInstance.getValidAccessToken.mockResolvedValue('local-token');

      await ensureAuthenticated();

      expect(mockPerUserTokenStorageInstance.getTokenForUser).not.toHaveBeenCalled();
      expect(mockPerUserTokenStorageInstance.setTokenForUser).not.toHaveBeenCalled();
    });

    test('should NOT call exchangeOBO in local mode', async () => {
      mockTokenStorageInstance.getValidAccessToken.mockResolvedValue('local-token');

      await ensureAuthenticated();

      expect(exchangeOBO).not.toHaveBeenCalled();
    });
  });

  // ── Hosted Mode (user context present) ──────────────────────────────

  describe('hosted mode (user context present)', () => {
    const testUserCtx = { userId: 'test-user-123', entraToken: 'entra-jwt-token' };

    const mockOBOResponse = {
      access_token: 'graph-token-from-obo',
      refresh_token: 'graph-refresh-from-obo',
      expires_in: 3600,
      scope: 'Mail.Read Mail.ReadWrite User.Read Calendars.Read',
      token_type: 'Bearer',
    };

    test('should return cached Graph token from per-user storage if valid', async () => {
      mockPerUserTokenStorageInstance.getTokenForUser.mockReturnValue('cached-graph-token');

      const token = await requestContext.run(testUserCtx, async () => {
        return ensureAuthenticated();
      });

      expect(token).toBe('cached-graph-token');
      expect(mockPerUserTokenStorageInstance.getTokenForUser).toHaveBeenCalledWith('test-user-123');
      expect(exchangeOBO).not.toHaveBeenCalled();
    });

    test('should call OBO exchange when per-user token is null (missing)', async () => {
      mockPerUserTokenStorageInstance.getTokenForUser.mockReturnValue(null);
      exchangeOBO.mockResolvedValue(mockOBOResponse);

      const token = await requestContext.run(testUserCtx, async () => {
        return ensureAuthenticated();
      });

      expect(token).toBe('graph-token-from-obo');
      expect(exchangeOBO).toHaveBeenCalledTimes(1);
      expect(exchangeOBO).toHaveBeenCalledWith('entra-jwt-token', expect.objectContaining({
        clientId: 'test-client-id',
        clientSecret: 'test-client-secret',
        tenantId: 'test-tenant-id',
      }));
    });

    test('should filter offline_access from OBO scopes', async () => {
      mockPerUserTokenStorageInstance.getTokenForUser.mockReturnValue(null);
      exchangeOBO.mockResolvedValue(mockOBOResponse);

      await requestContext.run(testUserCtx, async () => {
        return ensureAuthenticated();
      });

      const oboConfig = exchangeOBO.mock.calls[0][1];
      expect(oboConfig.scopes).not.toContain('offline_access');
      expect(oboConfig.scopes).toContain('Mail.Read');
      expect(oboConfig.scopes).toContain('Mail.ReadWrite');
      expect(oboConfig.scopes).toContain('User.Read');
      expect(oboConfig.scopes).toContain('Calendars.Read');
    });

    test('should store OBO result in per-user storage after exchange', async () => {
      mockPerUserTokenStorageInstance.getTokenForUser.mockReturnValue(null);
      exchangeOBO.mockResolvedValue(mockOBOResponse);

      await requestContext.run(testUserCtx, async () => {
        return ensureAuthenticated();
      });

      expect(mockPerUserTokenStorageInstance.setTokenForUser).toHaveBeenCalledWith(
        'test-user-123',
        mockOBOResponse
      );
    });

    test('should invalidate per-user token with forceRefresh, then do OBO exchange', async () => {
      mockPerUserTokenStorageInstance.getTokenForUser.mockReturnValue(null);
      exchangeOBO.mockResolvedValue(mockOBOResponse);

      const token = await requestContext.run(testUserCtx, async () => {
        return ensureAuthenticated({ forceRefresh: true });
      });

      expect(mockPerUserTokenStorageInstance.invalidateUser).toHaveBeenCalledWith('test-user-123');
      expect(token).toBe('graph-token-from-obo');
    });

    test('should propagate OBO exchange failure as error', async () => {
      mockPerUserTokenStorageInstance.getTokenForUser.mockReturnValue(null);
      exchangeOBO.mockRejectedValue(new Error('OBO exchange failed: invalid or expired user token'));

      await expect(
        requestContext.run(testUserCtx, async () => {
          return ensureAuthenticated();
        })
      ).rejects.toThrow('OBO exchange failed: invalid or expired user token');
    });

    test('should NOT interact with the singleton TokenStorage in hosted mode', async () => {
      mockPerUserTokenStorageInstance.getTokenForUser.mockReturnValue('cached-token');

      await requestContext.run(testUserCtx, async () => {
        return ensureAuthenticated();
      });

      expect(mockTokenStorageInstance.getValidAccessToken).not.toHaveBeenCalled();
      expect(mockTokenStorageInstance.invalidateAccessToken).not.toHaveBeenCalled();
    });

    test('should NOT interact with singleton TokenStorage even with forceRefresh in hosted mode', async () => {
      mockPerUserTokenStorageInstance.getTokenForUser.mockReturnValue(null);
      exchangeOBO.mockResolvedValue(mockOBOResponse);

      await requestContext.run(testUserCtx, async () => {
        return ensureAuthenticated({ forceRefresh: true });
      });

      expect(mockTokenStorageInstance.getValidAccessToken).not.toHaveBeenCalled();
      expect(mockTokenStorageInstance.invalidateAccessToken).not.toHaveBeenCalled();
    });
  });
});
