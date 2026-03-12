const { getGraphToken, clearTokenCache, clearUserToken } = require('../../auth/obo-exchange');

// Mock config
jest.mock('../../config', () => ({
  AUTH_CONFIG: {
    clientId: 'test-client-id',
    clientSecret: 'test-client-secret',
    tenantId: 'test-tenant-id',
  },
  CONNECTOR_AUTH: {
    oboScopes: 'Mail.Read Mail.Send Calendars.ReadWrite User.Read',
  },
}));

// Helper to build a mock fetch response
function mockFetchResponse(data, ok = true, status = 200) {
  return Promise.resolve({
    ok,
    status,
    json: () => Promise.resolve(data),
  });
}

describe('OBO Token Exchange', () => {
  const originalFetch = global.fetch;

  beforeEach(() => {
    clearTokenCache();
    global.fetch = jest.fn();
  });

  afterAll(() => {
    global.fetch = originalFetch;
  });

  describe('exchangeOboToken (via getGraphToken)', () => {
    it('should return a Graph access token on successful OBO exchange', async () => {
      global.fetch.mockReturnValue(mockFetchResponse({
        access_token: 'graph-access-token-123',
        expires_in: 3600,
        token_type: 'Bearer',
      }));

      const token = await getGraphToken('entra-jwt-abc', 'user-oid-1');

      expect(token).toBe('graph-access-token-123');
      expect(global.fetch).toHaveBeenCalledTimes(1);
    });

    it('should throw a descriptive error when the OBO exchange fails', async () => {
      global.fetch.mockReturnValue(mockFetchResponse({
        error: 'invalid_grant',
        error_description: 'AADSTS65001: The user has not consented to the required scopes.',
      }, false, 400));

      await expect(getGraphToken('bad-entra-jwt', 'user-oid-fail'))
        .rejects
        .toThrow('OBO token exchange failed: [invalid_grant] AADSTS65001: The user has not consented to the required scopes.');
    });

    it('should throw with fallback messages when error response lacks details', async () => {
      global.fetch.mockReturnValue(mockFetchResponse({}, false, 500));

      await expect(getGraphToken('bad-jwt', 'user-oid-no-details'))
        .rejects
        .toThrow('OBO token exchange failed: [unknown_error] No error description provided');
    });

    it('should send the correct request body with all required OBO parameters', async () => {
      global.fetch.mockReturnValue(mockFetchResponse({
        access_token: 'graph-token',
        expires_in: 3600,
      }));

      await getGraphToken('my-entra-token', 'user-oid-params');

      expect(global.fetch).toHaveBeenCalledTimes(1);

      const [url, options] = global.fetch.mock.calls[0];

      // Verify endpoint
      expect(url).toBe('https://login.microsoftonline.com/test-tenant-id/oauth2/v2.0/token');

      // Verify headers
      expect(options.method).toBe('POST');
      expect(options.headers['Content-Type']).toBe('application/x-www-form-urlencoded');

      // Parse the URL-encoded body
      const params = new URLSearchParams(options.body);
      expect(params.get('grant_type')).toBe('urn:ietf:params:oauth:grant-type:jwt-bearer');
      expect(params.get('client_id')).toBe('test-client-id');
      expect(params.get('client_secret')).toBe('test-client-secret');
      expect(params.get('assertion')).toBe('my-entra-token');
      expect(params.get('scope')).toBe('Mail.Read Mail.Send Calendars.ReadWrite User.Read');
      expect(params.get('requested_token_use')).toBe('on_behalf_of');
    });
  });

  describe('Token caching', () => {
    it('should return a cached token on the second call without making another HTTP request', async () => {
      global.fetch.mockReturnValue(mockFetchResponse({
        access_token: 'cached-graph-token',
        expires_in: 3600,
      }));

      const token1 = await getGraphToken('entra-jwt', 'user-oid-cache');
      const token2 = await getGraphToken('entra-jwt', 'user-oid-cache');

      expect(token1).toBe('cached-graph-token');
      expect(token2).toBe('cached-graph-token');
      expect(global.fetch).toHaveBeenCalledTimes(1);
    });

    it('should make a new HTTP request when the cached token has expired', async () => {
      // First call: token that expires in 1 second (effectively expired with 5-min buffer)
      global.fetch
        .mockReturnValueOnce(mockFetchResponse({
          access_token: 'expired-token',
          expires_in: 1, // 1 second — well within the 5-minute buffer
        }))
        .mockReturnValueOnce(mockFetchResponse({
          access_token: 'fresh-token',
          expires_in: 3600,
        }));

      const token1 = await getGraphToken('entra-jwt', 'user-oid-expiry');
      // The token was cached but expires_in=1s means it's already past the 5-min buffer
      const token2 = await getGraphToken('entra-jwt', 'user-oid-expiry');

      // First call returns the token from the first exchange (it's still cached at that point)
      expect(token1).toBe('expired-token');
      // Second call should trigger a new exchange because cached token is within the buffer
      expect(token2).toBe('fresh-token');
      expect(global.fetch).toHaveBeenCalledTimes(2);
    });

    it('should cache tokens separately per user OID', async () => {
      global.fetch
        .mockReturnValueOnce(mockFetchResponse({
          access_token: 'token-for-user-a',
          expires_in: 3600,
        }))
        .mockReturnValueOnce(mockFetchResponse({
          access_token: 'token-for-user-b',
          expires_in: 3600,
        }));

      const tokenA = await getGraphToken('entra-jwt-a', 'user-a');
      const tokenB = await getGraphToken('entra-jwt-b', 'user-b');

      expect(tokenA).toBe('token-for-user-a');
      expect(tokenB).toBe('token-for-user-b');
      expect(global.fetch).toHaveBeenCalledTimes(2);

      // Subsequent calls for each user should be cached
      const tokenA2 = await getGraphToken('entra-jwt-a', 'user-a');
      const tokenB2 = await getGraphToken('entra-jwt-b', 'user-b');

      expect(tokenA2).toBe('token-for-user-a');
      expect(tokenB2).toBe('token-for-user-b');
      expect(global.fetch).toHaveBeenCalledTimes(2); // No additional calls
    });
  });

  describe('Cache clearing', () => {
    it('clearTokenCache() should clear all cached tokens', async () => {
      global.fetch
        .mockReturnValueOnce(mockFetchResponse({
          access_token: 'token-1',
          expires_in: 3600,
        }))
        .mockReturnValueOnce(mockFetchResponse({
          access_token: 'token-after-clear',
          expires_in: 3600,
        }));

      await getGraphToken('entra-jwt', 'user-oid-clear');
      expect(global.fetch).toHaveBeenCalledTimes(1);

      clearTokenCache();

      const token = await getGraphToken('entra-jwt', 'user-oid-clear');
      expect(token).toBe('token-after-clear');
      expect(global.fetch).toHaveBeenCalledTimes(2);
    });

    it('clearUserToken() should clear only the specified user token', async () => {
      global.fetch
        .mockReturnValueOnce(mockFetchResponse({
          access_token: 'token-user-x',
          expires_in: 3600,
        }))
        .mockReturnValueOnce(mockFetchResponse({
          access_token: 'token-user-y',
          expires_in: 3600,
        }))
        .mockReturnValueOnce(mockFetchResponse({
          access_token: 'token-user-x-refreshed',
          expires_in: 3600,
        }));

      await getGraphToken('jwt-x', 'user-x');
      await getGraphToken('jwt-y', 'user-y');
      expect(global.fetch).toHaveBeenCalledTimes(2);

      // Clear only user-x
      clearUserToken('user-x');

      // user-y should still be cached
      const tokenY = await getGraphToken('jwt-y', 'user-y');
      expect(tokenY).toBe('token-user-y');
      expect(global.fetch).toHaveBeenCalledTimes(2); // No new call for user-y

      // user-x should require a new exchange
      const tokenX = await getGraphToken('jwt-x', 'user-x');
      expect(tokenX).toBe('token-user-x-refreshed');
      expect(global.fetch).toHaveBeenCalledTimes(3);
    });
  });
});
