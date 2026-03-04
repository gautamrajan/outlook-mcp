const https = require('https');
const querystring = require('querystring');
const { exchangeOBO } = require('../../auth/obo-exchange');

jest.mock('https');

const validConfig = {
  clientId: 'test-client-id',
  clientSecret: 'test-client-secret',
  tenantId: 'test-tenant-id',
  scopes: ['Mail.ReadWrite', 'Calendars.ReadWrite', 'User.Read'],
};

const mockUserToken = 'eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiJ9.mock-user-token';

const mockSuccessResponse = {
  access_token: 'graph-access-token-xyz',
  refresh_token: 'graph-refresh-token-abc',
  expires_in: 3600,
  scope: 'Mail.ReadWrite Calendars.ReadWrite User.Read',
  token_type: 'Bearer',
};

describe('exchangeOBO', () => {
  let mockRequest;

  beforeEach(() => {
    jest.clearAllMocks();
    mockRequest = {
      on: jest.fn((event, cb) => {
        if (event === 'error') mockRequest.errorHandler = cb;
        return mockRequest;
      }),
      write: jest.fn(),
      end: jest.fn(),
    };
    https.request.mockImplementation((url, options, callback) => {
      mockRequest.callback = callback;
      return mockRequest;
    });
  });

  // Helper to simulate an HTTPS response
  function simulateResponse(statusCode, body) {
    const responseBody = typeof body === 'string' ? body : JSON.stringify(body);
    const mockRes = {
      statusCode,
      on: (event, cb) => {
        if (event === 'data') cb(Buffer.from(responseBody));
        if (event === 'end') cb();
      },
    };
    mockRequest.callback(mockRes);
  }

  describe('successful exchange', () => {
    it('should return token data on successful OBO exchange', async () => {
      const exchangePromise = exchangeOBO(mockUserToken, validConfig);
      simulateResponse(200, mockSuccessResponse);

      const result = await exchangePromise;

      expect(result).toEqual(mockSuccessResponse);
      expect(result.access_token).toBe('graph-access-token-xyz');
      expect(result.refresh_token).toBe('graph-refresh-token-abc');
      expect(result.expires_in).toBe(3600);
      expect(result.scope).toBe('Mail.ReadWrite Calendars.ReadWrite User.Read');
      expect(result.token_type).toBe('Bearer');
    });
  });

  describe('parameter validation', () => {
    it('should throw on missing userAccessToken', async () => {
      await expect(exchangeOBO(null, validConfig))
        .rejects.toThrow('OBO exchange failed: userAccessToken is required');
    });

    it('should throw on undefined userAccessToken', async () => {
      await expect(exchangeOBO(undefined, validConfig))
        .rejects.toThrow('OBO exchange failed: userAccessToken is required');
    });

    it('should throw on empty string userAccessToken', async () => {
      await expect(exchangeOBO('', validConfig))
        .rejects.toThrow('OBO exchange failed: userAccessToken is required');
    });

    it('should throw on missing config', async () => {
      await expect(exchangeOBO(mockUserToken, null))
        .rejects.toThrow('OBO exchange failed: config is required');
    });

    it('should throw on undefined config', async () => {
      await expect(exchangeOBO(mockUserToken, undefined))
        .rejects.toThrow('OBO exchange failed: config is required');
    });

    it('should throw on missing config.clientId', async () => {
      const badConfig = { ...validConfig, clientId: undefined };
      await expect(exchangeOBO(mockUserToken, badConfig))
        .rejects.toThrow('OBO exchange failed: config.clientId is required');
    });

    it('should throw on missing config.clientSecret', async () => {
      const badConfig = { ...validConfig, clientSecret: undefined };
      await expect(exchangeOBO(mockUserToken, badConfig))
        .rejects.toThrow('OBO exchange failed: config.clientSecret is required');
    });

    it('should throw on missing config.tenantId', async () => {
      const badConfig = { ...validConfig, tenantId: undefined };
      await expect(exchangeOBO(mockUserToken, badConfig))
        .rejects.toThrow('OBO exchange failed: config.tenantId is required');
    });

    it('should throw on missing config.scopes', async () => {
      const badConfig = { ...validConfig, scopes: undefined };
      await expect(exchangeOBO(mockUserToken, badConfig))
        .rejects.toThrow('OBO exchange failed: config.scopes is required');
    });
  });

  describe('Entra error responses', () => {
    it('should handle invalid_grant error', async () => {
      const exchangePromise = exchangeOBO(mockUserToken, validConfig);
      simulateResponse(400, {
        error: 'invalid_grant',
        error_description: 'AADSTS65001: The user or administrator has not consented.',
      });

      await expect(exchangePromise)
        .rejects.toThrow('OBO exchange failed: invalid or expired user token');
    });

    it('should handle consent_required error', async () => {
      const exchangePromise = exchangeOBO(mockUserToken, validConfig);
      simulateResponse(400, {
        error: 'consent_required',
        error_description: 'AADSTS65001: Consent required.',
      });

      await expect(exchangePromise)
        .rejects.toThrow('OBO exchange failed: admin consent required for Graph permissions');
    });

    it('should handle interaction_required error', async () => {
      const exchangePromise = exchangeOBO(mockUserToken, validConfig);
      simulateResponse(400, {
        error: 'interaction_required',
        error_description: 'AADSTS50079: User interaction required.',
      });

      await expect(exchangePromise)
        .rejects.toThrow('OBO exchange failed: admin consent required for Graph permissions');
    });

    it('should handle unknown Entra error codes with raw details', async () => {
      const exchangePromise = exchangeOBO(mockUserToken, validConfig);
      simulateResponse(400, {
        error: 'server_error',
        error_description: 'Something unexpected happened.',
      });

      await expect(exchangePromise)
        .rejects.toThrow('OBO exchange failed: server_error - Something unexpected happened.');
    });

    it('should handle error response without error_description', async () => {
      const exchangePromise = exchangeOBO(mockUserToken, validConfig);
      simulateResponse(400, {
        error: 'server_error',
      });

      await expect(exchangePromise)
        .rejects.toThrow('OBO exchange failed: server_error');
    });
  });

  describe('network errors', () => {
    it('should handle request error event', async () => {
      const exchangePromise = exchangeOBO(mockUserToken, validConfig);
      mockRequest.errorHandler(new Error('ECONNREFUSED'));

      await expect(exchangePromise)
        .rejects.toThrow('OBO exchange failed: network error - ECONNREFUSED');
    });

    it('should handle DNS resolution failure', async () => {
      const exchangePromise = exchangeOBO(mockUserToken, validConfig);
      mockRequest.errorHandler(new Error('getaddrinfo ENOTFOUND login.microsoftonline.com'));

      await expect(exchangePromise)
        .rejects.toThrow('OBO exchange failed: network error - getaddrinfo ENOTFOUND login.microsoftonline.com');
    });
  });

  describe('request timeout', () => {
    it('should set a 30-second timeout on the request', async () => {
      const exchangePromise = exchangeOBO(mockUserToken, validConfig);
      simulateResponse(200, mockSuccessResponse);
      await exchangePromise;

      const requestOptions = https.request.mock.calls[0][1];
      expect(requestOptions.timeout).toBe(30000);
    });

    it('should destroy the request when timeout fires', async () => {
      mockRequest.destroy = jest.fn();

      const exchangePromise = exchangeOBO(mockUserToken, validConfig);

      // Find and invoke the timeout handler
      const timeoutCall = mockRequest.on.mock.calls.find(([event]) => event === 'timeout');
      expect(timeoutCall).toBeDefined();
      timeoutCall[1](); // invoke the timeout callback

      await expect(exchangePromise)
        .rejects.toThrow('OBO exchange failed: request timed out');

      expect(mockRequest.destroy).toHaveBeenCalled();
    });

    it('should reject with a descriptive timeout error message', async () => {
      mockRequest.destroy = jest.fn();

      const exchangePromise = exchangeOBO(mockUserToken, validConfig);

      const timeoutCall = mockRequest.on.mock.calls.find(([event]) => event === 'timeout');
      timeoutCall[1]();

      await expect(exchangePromise)
        .rejects.toThrow(/OBO exchange failed: request timed out/);
    });
  });

  describe('malformed responses', () => {
    it('should handle non-JSON response from Entra', async () => {
      const exchangePromise = exchangeOBO(mockUserToken, validConfig);
      simulateResponse(500, '<html>Internal Server Error</html>');

      await expect(exchangePromise)
        .rejects.toThrow('OBO exchange failed: invalid response from token endpoint');
    });
  });

  describe('request formatting', () => {
    it('should correctly URL-encode the request body', async () => {
      const exchangePromise = exchangeOBO(mockUserToken, validConfig);
      simulateResponse(200, mockSuccessResponse);
      await exchangePromise;

      const writtenBody = mockRequest.write.mock.calls[0][0];
      const parsed = querystring.parse(writtenBody);

      expect(parsed.grant_type).toBe('urn:ietf:params:oauth:grant-type:jwt-bearer');
      expect(parsed.client_id).toBe('test-client-id');
      expect(parsed.client_secret).toBe('test-client-secret');
      expect(parsed.assertion).toBe(mockUserToken);
      expect(parsed.scope).toBe('Mail.ReadWrite Calendars.ReadWrite User.Read');
      expect(parsed.requested_token_use).toBe('on_behalf_of');
    });

    it('should send correct Content-Type header', async () => {
      const exchangePromise = exchangeOBO(mockUserToken, validConfig);
      simulateResponse(200, mockSuccessResponse);
      await exchangePromise;

      const requestOptions = https.request.mock.calls[0][1];
      expect(requestOptions.headers['Content-Type']).toBe('application/x-www-form-urlencoded');
    });

    it('should send Content-Length header', async () => {
      const exchangePromise = exchangeOBO(mockUserToken, validConfig);
      simulateResponse(200, mockSuccessResponse);
      await exchangePromise;

      const requestOptions = https.request.mock.calls[0][1];
      expect(requestOptions.headers['Content-Length']).toBeDefined();
      expect(typeof requestOptions.headers['Content-Length']).toBe('number');
    });

    it('should POST to the correct tenant-specific endpoint', async () => {
      const exchangePromise = exchangeOBO(mockUserToken, validConfig);
      simulateResponse(200, mockSuccessResponse);
      await exchangePromise;

      const requestUrl = https.request.mock.calls[0][0];
      expect(requestUrl).toBe('https://login.microsoftonline.com/test-tenant-id/oauth2/v2.0/token');
    });

    it('should use POST method', async () => {
      const exchangePromise = exchangeOBO(mockUserToken, validConfig);
      simulateResponse(200, mockSuccessResponse);
      await exchangePromise;

      const requestOptions = https.request.mock.calls[0][1];
      expect(requestOptions.method).toBe('POST');
    });

    it('should use different tenant ID when config changes', async () => {
      const otherConfig = { ...validConfig, tenantId: 'other-tenant-uuid' };
      const exchangePromise = exchangeOBO(mockUserToken, otherConfig);
      simulateResponse(200, mockSuccessResponse);
      await exchangePromise;

      const requestUrl = https.request.mock.calls[0][0];
      expect(requestUrl).toBe('https://login.microsoftonline.com/other-tenant-uuid/oauth2/v2.0/token');
    });
  });
});
