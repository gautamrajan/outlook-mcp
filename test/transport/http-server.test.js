/**
 * Tests for transport/http-server.js — HTTP transport for the MCP server.
 *
 * Strategy:
 *   - Mock the MCP SDK (Server, StreamableHTTPServerTransport) so we are
 *     testing OUR wiring, not the SDK internals.
 *   - Use supertest to make real HTTP requests against the Express app.
 *   - Validate that request context (AsyncLocalStorage) is set correctly.
 *   - Test session-token middleware rejects unauthenticated requests.
 */
const http = require('http');
const request = require('supertest');

// ── Mocks ────────────────────────────────────────────────────────────

// Mock the MCP SDK Server class
const mockServerConnect = jest.fn().mockResolvedValue(undefined);
const mockServerClose = jest.fn().mockResolvedValue(undefined);
const MockServer = jest.fn().mockImplementation(() => ({
  connect: mockServerConnect,
  close: mockServerClose,
  fallbackRequestHandler: null,
}));
jest.mock('@modelcontextprotocol/sdk/server/index.js', () => ({
  Server: MockServer,
}));

// Mock StreamableHTTPServerTransport
const mockHandleRequest = jest.fn().mockImplementation((_req, res) => {
  // Default: send a 200 so supertest gets a response
  if (!res.headersSent) {
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ jsonrpc: '2.0', result: {} }));
  }
});
const MockTransport = jest.fn().mockImplementation(() => ({
  handleRequest: mockHandleRequest,
}));
jest.mock('@modelcontextprotocol/sdk/server/streamableHttp.js', () => ({
  StreamableHTTPServerTransport: MockTransport,
}));

// Mock config
jest.mock('../../config', () => ({
  SERVER_NAME: 'test-outlook-assistant',
  SERVER_VERSION: '1.0.0-test',
  AUTH_CONFIG: {
    tenantId: 'test-tenant-id',
    clientId: 'test-client-id',
  },
}));

// Spy on request context to verify it's set during handling
const { requestContext, getUserContext } = require('../../auth/request-context');

// ── Import under test (after all mocks) ──────────────────────────────
const { createHttpApp, startHttpServer, createSessionMiddleware } = require('../../transport/http-server');

// ── Helpers ──────────────────────────────────────────────────────────

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

/**
 * Create a minimal mock SessionStore for testing.
 */
function createMockSessionStore() {
  const sessions = new Map();
  return {
    _sessions: sessions,
    validateSession: jest.fn((token) => {
      return sessions.get(token) || null;
    }),
    /** Helper to seed a valid session for testing. */
    _addSession(token, userId) {
      const now = new Date();
      const expiresAt = new Date(now.getTime() + 30 * 24 * 60 * 60 * 1000);
      sessions.set(token, {
        userId,
        createdAt: now.toISOString(),
        expiresAt: expiresAt.toISOString(),
      });
    },
  };
}

// ── Tests ────────────────────────────────────────────────────────────

describe('HTTP Transport Server', () => {
  let app;

  beforeEach(() => {
    jest.clearAllMocks();
    app = createHttpApp();
  });

  // ── Startup ──────────────────────────────────────────────────────

  describe('createHttpApp()', () => {
    test('returns an Express application', () => {
      expect(app).toBeDefined();
      expect(typeof app.listen).toBe('function');
    });

    test('accepts options object with sessionStore', () => {
      const mockStore = createMockSessionStore();
      const authedApp = createHttpApp({ sessionStore: mockStore });
      expect(authedApp).toBeDefined();
      expect(typeof authedApp.listen).toBe('function');
    });
  });

  describe('startHttpServer()', () => {
    let server;

    afterEach((done) => {
      if (server && server.listening) {
        server.close(done);
      } else {
        done();
      }
    });

    test('starts listening on the configured port and returns the http.Server', (done) => {
      const originalPort = process.env.PORT;
      process.env.PORT = '0'; // random free port

      server = startHttpServer();
      expect(server).toBeInstanceOf(http.Server);

      server.on('listening', () => {
        const addr = server.address();
        expect(addr.port).toBeGreaterThan(0);
        process.env.PORT = originalPort;
        done();
      });
    });

    test('defaults to port 3000 when PORT env is unset', (done) => {
      const originalPort = process.env.PORT;
      delete process.env.PORT;

      // We can't actually bind 3000 in tests (it may be in use),
      // so we just verify the function doesn't throw and returns a server.
      // We'll close it immediately.
      server = startHttpServer();
      expect(server).toBeInstanceOf(http.Server);

      // Give it a tick to attempt listening, then close
      server.on('listening', () => {
        process.env.PORT = originalPort;
        done();
      });

      server.on('error', (err) => {
        // Port 3000 might be in use — that's fine for this test
        process.env.PORT = originalPort;
        if (err.code === 'EADDRINUSE') {
          done(); // expected in CI/dev
        } else {
          done(err);
        }
      });
    });

    test('passes sessionStore through to createHttpApp', (done) => {
      const originalPort = process.env.PORT;
      process.env.PORT = '0';

      const mockStore = createMockSessionStore();
      server = startHttpServer({ sessionStore: mockStore });
      expect(server).toBeInstanceOf(http.Server);

      server.on('listening', () => {
        process.env.PORT = originalPort;
        done();
      });
    });
  });

  // ── Request routing ──────────────────────────────────────────────

  describe('MCP route handling', () => {
    test('POST /mcp creates a transport and calls handleRequest', async () => {
      await request(app)
        .post('/mcp')
        .set('Content-Type', 'application/json')
        .send(initializeBody());

      // A transport should have been created in stateless mode
      expect(MockTransport).toHaveBeenCalledWith(
        expect.objectContaining({ sessionIdGenerator: undefined })
      );

      // handleRequest should have been called
      expect(mockHandleRequest).toHaveBeenCalled();
    });

    test('GET /mcp creates a transport and calls handleRequest', async () => {
      await request(app).get('/mcp');

      expect(MockTransport).toHaveBeenCalledWith(
        expect.objectContaining({ sessionIdGenerator: undefined })
      );
      expect(mockHandleRequest).toHaveBeenCalled();
    });

    test('DELETE /mcp creates a transport and calls handleRequest', async () => {
      await request(app).delete('/mcp');

      expect(MockTransport).toHaveBeenCalledWith(
        expect.objectContaining({ sessionIdGenerator: undefined })
      );
      expect(mockHandleRequest).toHaveBeenCalled();
    });

    test('creates a new MCP Server per request with tools capability', async () => {
      await request(app)
        .post('/mcp')
        .send(initializeBody());

      expect(MockServer).toHaveBeenCalledWith(
        expect.objectContaining({
          name: 'test-outlook-assistant',
          version: '1.0.0-test',
        }),
        expect.objectContaining({
          capabilities: expect.objectContaining({
            tools: expect.any(Object),
          }),
        })
      );
    });

    test('connects the MCP Server to the transport', async () => {
      await request(app)
        .post('/mcp')
        .send(initializeBody());

      expect(mockServerConnect).toHaveBeenCalled();
    });

    test('each request gets its own Server and Transport instance', async () => {
      await request(app)
        .post('/mcp')
        .send(initializeBody());

      await request(app)
        .post('/mcp')
        .send(initializeBody());

      // Two requests = two Server instances + two Transport instances
      expect(MockServer).toHaveBeenCalledTimes(2);
      expect(MockTransport).toHaveBeenCalledTimes(2);
    });

    test('non-MCP routes return 404', async () => {
      const res = await request(app).get('/not-a-real-route');
      expect(res.status).toBe(404);
    });
  });

  // ── Request context (AsyncLocalStorage) ──────────────────────────

  describe('User context in AsyncLocalStorage', () => {
    test('wraps request handling in AsyncLocalStorage context', async () => {
      let capturedContext = null;

      // Override handleRequest to capture the async context
      mockHandleRequest.mockImplementationOnce((_req, res) => {
        capturedContext = getUserContext();
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ jsonrpc: '2.0', result: {} }));
      });

      await request(app)
        .post('/mcp')
        .send(initializeBody());

      // Context should exist (user fields are undefined without auth middleware)
      expect(capturedContext).not.toBeNull();
      expect(capturedContext).toHaveProperty('userId');
      expect(capturedContext).toHaveProperty('sessionToken');
    });

    test('populates userId and sessionToken from middleware when session is active', async () => {
      const mockStore = createMockSessionStore();
      const validToken = 'test-valid-token-123';
      mockStore._addSession(validToken, 'user-abc');
      const authedApp = createHttpApp({ sessionStore: mockStore });

      let capturedContext = null;

      mockHandleRequest.mockImplementationOnce((_req, res) => {
        capturedContext = getUserContext();
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ jsonrpc: '2.0', result: {} }));
      });

      await request(authedApp)
        .post('/mcp')
        .set('Authorization', `Bearer ${validToken}`)
        .send(initializeBody());

      expect(capturedContext).not.toBeNull();
      expect(capturedContext.userId).toBe('user-abc');
      expect(capturedContext.sessionToken).toBe(validToken);
    });
  });

  // ── Session middleware ─────────────────────────────────────────────

  describe('Session token middleware', () => {
    let mockStore;
    let authedApp;
    const validToken = 'valid-session-token-uuid';
    const testUserId = 'user-oid-12345';

    beforeEach(() => {
      jest.clearAllMocks();
      mockStore = createMockSessionStore();
      mockStore._addSession(validToken, testUserId);
      authedApp = createHttpApp({ sessionStore: mockStore });
    });

    test('rejects request without Authorization header with 401', async () => {
      const res = await request(authedApp)
        .post('/mcp')
        .set('Content-Type', 'application/json')
        .send(initializeBody());

      expect(res.status).toBe(401);
      expect(res.body).toEqual({
        error: 'auth_required',
        message: 'Session expired or missing. Authenticate at: /auth/login',
        authUrl: '/auth/login',
      });

      // Should NOT have reached the MCP handler
      expect(mockHandleRequest).not.toHaveBeenCalled();
    });

    test('rejects request with non-Bearer Authorization header with 401', async () => {
      const res = await request(authedApp)
        .post('/mcp')
        .set('Authorization', 'Basic dXNlcjpwYXNz')
        .send(initializeBody());

      expect(res.status).toBe(401);
      expect(res.body.error).toBe('auth_required');
      expect(mockHandleRequest).not.toHaveBeenCalled();
    });

    test('rejects request with invalid/unknown token with 401', async () => {
      const res = await request(authedApp)
        .post('/mcp')
        .set('Authorization', 'Bearer totally-bogus-token')
        .send(initializeBody());

      expect(res.status).toBe(401);
      expect(res.body).toEqual({
        error: 'auth_required',
        message: 'Session expired or invalid. Re-authenticate at: /auth/login',
        authUrl: '/auth/login',
      });
      expect(mockStore.validateSession).toHaveBeenCalledWith('totally-bogus-token');
      expect(mockHandleRequest).not.toHaveBeenCalled();
    });

    test('allows request with valid token and sets req.user', async () => {
      let capturedReqUser = null;

      mockHandleRequest.mockImplementationOnce((req, res) => {
        capturedReqUser = req.user;
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ jsonrpc: '2.0', result: {} }));
      });

      const res = await request(authedApp)
        .post('/mcp')
        .set('Authorization', `Bearer ${validToken}`)
        .send(initializeBody());

      expect(res.status).toBe(200);
      expect(mockStore.validateSession).toHaveBeenCalledWith(validToken);
      expect(mockHandleRequest).toHaveBeenCalled();
      expect(capturedReqUser).toEqual({
        id: testUserId,
        sessionToken: validToken,
      });
    });

    test('401 response body has correct JSON shape with error, message, and authUrl', async () => {
      const res = await request(authedApp)
        .post('/mcp')
        .send(initializeBody());

      expect(res.status).toBe(401);
      expect(res.body).toHaveProperty('error');
      expect(res.body).toHaveProperty('message');
      expect(res.body).toHaveProperty('authUrl');
      expect(typeof res.body.error).toBe('string');
      expect(typeof res.body.message).toBe('string');
      expect(typeof res.body.authUrl).toBe('string');
    });

    test('middleware is not applied when sessionStore is not provided', async () => {
      // The default `app` (no sessionStore) should allow unauthenticated access
      const res = await request(app)
        .post('/mcp')
        .send(initializeBody());

      // Should reach the MCP handler (200 from mock) — no 401
      expect(res.status).toBe(200);
      expect(mockHandleRequest).toHaveBeenCalled();
    });
  });

  // ── createSessionMiddleware export ─────────────────────────────────

  describe('createSessionMiddleware()', () => {
    test('is exported and returns a function', () => {
      expect(typeof createSessionMiddleware).toBe('function');
      const mockStore = createMockSessionStore();
      const middleware = createSessionMiddleware(mockStore);
      expect(typeof middleware).toBe('function');
    });
  });

  // ── SIGTERM ──────────────────────────────────────────────────────

  describe('SIGTERM handling', () => {
    test('SIGTERM listener is registered and does not crash the process', () => {
      // The process should have a SIGTERM handler registered.
      // We cannot easily test that the http server closes on SIGTERM
      // without risking the test process, so we verify the listener exists.
      const sigTermListeners = process.listeners('SIGTERM');
      expect(sigTermListeners.length).toBeGreaterThan(0);
    });
  });

  // ── Body parsing ─────────────────────────────────────────────────

  describe('Body parsing', () => {
    test('does NOT add express.json() middleware — raw body passed to transport', async () => {
      // The transport should receive the raw request, not pre-parsed JSON.
      // We verify by checking that handleRequest receives req (not req.body as parsed object).
      let receivedReq = null;

      mockHandleRequest.mockImplementationOnce((req, res) => {
        receivedReq = req;
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ jsonrpc: '2.0', result: {} }));
      });

      await request(app)
        .post('/mcp')
        .set('Content-Type', 'application/json')
        .send(JSON.stringify(initializeBody()));

      // The request object should NOT have a pre-parsed .body from express.json()
      // Since we don't apply express.json(), req.body should be undefined
      expect(receivedReq).not.toBeNull();
      expect(receivedReq.body).toBeUndefined();
    });
  });

  // ── Fallback request handler ─────────────────────────────────────

  describe('Fallback request handler', () => {
    test('each per-request Server gets a fallbackRequestHandler', async () => {
      await request(app)
        .post('/mcp')
        .send(initializeBody());

      // The MockServer instance should have had fallbackRequestHandler set
      const serverInstance = MockServer.mock.results[0].value;
      expect(serverInstance.fallbackRequestHandler).toBeDefined();
      expect(typeof serverInstance.fallbackRequestHandler).toBe('function');
    });
  });
});
