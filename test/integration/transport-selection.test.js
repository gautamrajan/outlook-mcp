/**
 * Integration tests for index.js transport selection.
 *
 * Verifies that the MCP_TRANSPORT environment variable correctly selects
 * between stdio (default) and HTTP transport modes.
 *
 * Strategy:
 *   - Mock the MCP SDK (Server, StdioServerTransport) and transport/http-server.js
 *     so that requiring index.js does not start real servers.
 *   - Use jest.resetModules() between tests to re-evaluate index.js with
 *     different env var values.
 *   - Verify which transport was initialised based on mock call counts.
 */

// Keep a reference to the original env so we can restore it
const originalEnv = { ...process.env };

// ── Mock factories (recreated per test via resetModules) ─────────────

let mockServerConnect;
let mockStdioTransportInstance;
let MockServer;
let MockStdioServerTransport;
let mockStartHttpServer;

function setupMocks() {
  // MCP Server mock
  mockServerConnect = jest.fn().mockResolvedValue(undefined);
  MockServer = jest.fn().mockImplementation(() => ({
    connect: mockServerConnect,
    fallbackRequestHandler: null,
  }));
  jest.mock('@modelcontextprotocol/sdk/server/index.js', () => ({
    Server: MockServer,
  }));

  // StdioServerTransport mock
  mockStdioTransportInstance = { type: 'stdio-mock' };
  MockStdioServerTransport = jest.fn().mockImplementation(() => mockStdioTransportInstance);
  jest.mock('@modelcontextprotocol/sdk/server/stdio.js', () => ({
    StdioServerTransport: MockStdioServerTransport,
  }));

  // HTTP server mock
  mockStartHttpServer = jest.fn();
  jest.mock('../../transport/http-server', () => ({
    startHttpServer: mockStartHttpServer,
  }));

  // Mock config to avoid needing real env vars
  jest.mock('../../config', () => ({
    SERVER_NAME: 'test-outlook-assistant',
    SERVER_VERSION: '1.0.0-test',
    USE_TEST_MODE: false,
    HOSTED: {
      enabled: true,
      tokenEncryptionKey: 'test-encryption-key',
      tokenStorePath: '/tmp/test-hosted-tokens.json',
      sessionStorePath: '/tmp/test-sessions.json',
      hostedRedirectUri: '',
    },
    AUTH_CONFIG: {
      tokenStorePath: '/tmp/test-tokens.json',
      clientId: 'test-client-id',
      clientSecret: 'test-client-secret',
      tenantId: 'test-tenant-id',
      tokenEndpoint: 'https://login.microsoftonline.com/test-tenant-id/oauth2/v2.0/token',
      redirectUri: 'http://localhost:3333/auth/callback',
      scopes: ['offline_access', 'Mail.Read'],
    },
  }));

  // Mock PerUserTokenStorage and SessionStore (used by HTTP branch in index.js)
  jest.mock('../../auth/per-user-token-storage', () => {
    return jest.fn().mockImplementation(() => ({
      loadFromFile: jest.fn().mockResolvedValue(undefined),
    }));
  });
  jest.mock('../../auth/session-store', () => {
    return jest.fn().mockImplementation(() => ({
      loadFromFile: jest.fn().mockResolvedValue(undefined),
    }));
  });

  // Mock auth, calendar, email, folder, rules modules to avoid side effects
  jest.mock('../../auth', () => ({
    authTools: [{ name: 'mock-auth-tool' }],
    setHostedTokenStorage: jest.fn(),
  }));
  jest.mock('../../calendar', () => ({
    calendarTools: [{ name: 'mock-calendar-tool' }],
  }));
  jest.mock('../../email', () => ({
    emailTools: [{ name: 'mock-email-tool' }],
  }));
  jest.mock('../../folder', () => ({
    folderTools: [{ name: 'mock-folder-tool' }],
  }));
  jest.mock('../../rules', () => ({
    rulesTools: [{ name: 'mock-rules-tool' }],
  }));

  // Mock dotenv to no-op
  jest.mock('dotenv', () => ({
    config: jest.fn(),
  }));
}

// ── Tests ────────────────────────────────────────────────────────────

describe('Transport selection (index.js)', () => {
  beforeEach(() => {
    jest.resetModules();
    jest.clearAllMocks();
    // Clean env
    delete process.env.MCP_TRANSPORT;
  });

  afterAll(() => {
    // Restore original env
    process.env = { ...originalEnv };
  });

  test('should use stdio transport when MCP_TRANSPORT is unset (default)', () => {
    delete process.env.MCP_TRANSPORT;
    setupMocks();

    require('../../index');

    // Stdio transport should have been instantiated and connected
    expect(MockStdioServerTransport).toHaveBeenCalledTimes(1);
    expect(mockServerConnect).toHaveBeenCalledWith(mockStdioTransportInstance);

    // HTTP server should NOT have been started
    expect(mockStartHttpServer).not.toHaveBeenCalled();
  });

  test('should use stdio transport when MCP_TRANSPORT=stdio', () => {
    process.env.MCP_TRANSPORT = 'stdio';
    setupMocks();

    require('../../index');

    expect(MockStdioServerTransport).toHaveBeenCalledTimes(1);
    expect(mockServerConnect).toHaveBeenCalledWith(mockStdioTransportInstance);
    expect(mockStartHttpServer).not.toHaveBeenCalled();
  });

  test('should start HTTP server when MCP_TRANSPORT=http', async () => {
    process.env.MCP_TRANSPORT = 'http';
    setupMocks();

    require('../../index');

    // Startup is async (Promise.all for store loading), so flush microtasks
    await new Promise(setImmediate);

    expect(mockStartHttpServer).toHaveBeenCalledTimes(1);

    // Stdio transport should NOT have been instantiated
    expect(MockStdioServerTransport).not.toHaveBeenCalled();
  });

  test('should be case-insensitive: MCP_TRANSPORT=HTTP works', async () => {
    process.env.MCP_TRANSPORT = 'HTTP';
    setupMocks();

    require('../../index');

    await new Promise(setImmediate);

    expect(mockStartHttpServer).toHaveBeenCalledTimes(1);
    expect(MockStdioServerTransport).not.toHaveBeenCalled();
  });

  test('should be case-insensitive: MCP_TRANSPORT=Http works', async () => {
    process.env.MCP_TRANSPORT = 'Http';
    setupMocks();

    require('../../index');

    await new Promise(setImmediate);

    expect(mockStartHttpServer).toHaveBeenCalledTimes(1);
    expect(MockStdioServerTransport).not.toHaveBeenCalled();
  });

  test('should be case-insensitive: MCP_TRANSPORT=STDIO works', () => {
    process.env.MCP_TRANSPORT = 'STDIO';
    setupMocks();

    require('../../index');

    expect(MockStdioServerTransport).toHaveBeenCalledTimes(1);
    expect(mockStartHttpServer).not.toHaveBeenCalled();
  });

  test('should treat unknown transport values as stdio (default)', () => {
    process.env.MCP_TRANSPORT = 'websocket';
    setupMocks();

    require('../../index');

    // Falls through to the else branch (stdio)
    expect(MockStdioServerTransport).toHaveBeenCalledTimes(1);
    expect(mockStartHttpServer).not.toHaveBeenCalled();
  });

  // ── Verify stdio mode starts without errors ─────────────────────

  test('stdio mode registers a SIGTERM handler', () => {
    delete process.env.MCP_TRANSPORT;
    setupMocks();

    const listenersBefore = process.listeners('SIGTERM').length;

    require('../../index');

    const listenersAfter = process.listeners('SIGTERM').length;
    expect(listenersAfter).toBeGreaterThan(listenersBefore);
  });

  test('stdio mode creates an MCP Server with correct server info', () => {
    delete process.env.MCP_TRANSPORT;
    setupMocks();

    require('../../index');

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
});
