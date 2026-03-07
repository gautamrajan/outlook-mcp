/**
 * Integration tests for PRM discovery endpoint and 401 challenges.
 *
 * Validates:
 *   - GET /.well-known/oauth-protected-resource returns valid RFC 9728 JSON
 *   - PRM includes the correct tenant ID in authorization_servers
 *   - PRM includes the correct scope (apiAppId/apiScope)
 *   - PRM is publicly accessible (no auth required)
 *   - 401 responses include WWW-Authenticate header with resource_metadata URL
 *
 * Strategy:
 *   - Mock `jwks-rsa` to prevent ESM import errors from jose
 *   - Mock MCP SDK classes
 *   - Use real createHttpApp() with a session store to test 401 behavior
 *   - supertest for HTTP assertions
 */

const supertest = require('supertest');

// ── Mocks (must come before any require of modules under test) ────────

jest.mock('jwks-rsa', () => {
  return jest.fn(() => ({
    getSigningKey: jest.fn(),
  }));
});

jest.mock('../../config', () => ({
  SERVER_NAME: 'test-outlook-assistant',
  SERVER_VERSION: '1.0.0-test',
  HOSTED: {
    publicBaseUrl: 'https://outlook-mcp.example.com',
  },
  AUTH_CONFIG: {
    clientId: 'test-client-id',
    clientSecret: 'test-client-secret',
    tenantId: 'test-tenant-abc-123',
    tokenEndpoint: 'https://login.microsoftonline.com/test-tenant-abc-123/oauth2/v2.0/token',
    redirectUri: 'http://localhost:3333/auth/callback',
    hostedRedirectUri: 'https://outlook-mcp.example.com/auth/callback',
    scopes: ['offline_access', 'Mail.Read', 'User.Read'],
    tokenStorePath: '/tmp/test-tokens.json',
    hostedTokenStorePath: '/tmp/test-hosted-tokens.json',
  },
  CONNECTOR_AUTH: {
    apiAppId: 'api://test-app-id-000',
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
const SessionStore = require('../../auth/session-store');
const config = require('../../config');

// ── Helpers ───────────────────────────────────────────────────────────

function createInMemorySessionStore() {
  const store = new SessionStore({});
  store.filePath = null;
  store.saveToFile = async () => {};
  return store;
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

describe('PRM Discovery & 401 Challenges — Integration', () => {
  let sessionStore;
  let app;

  beforeEach(() => {
    jest.clearAllMocks();
    sessionStore = createInMemorySessionStore();
    app = createHttpApp({ sessionStore });
  });

  // ═══════════════════════════════════════════════════════════════════════
  // PRM endpoint returns valid JSON
  // ═══════════════════════════════════════════════════════════════════════

  test('1. PRM endpoint returns 200 with valid JSON structure', async () => {
    const res = await supertest(app)
      .get('/.well-known/oauth-protected-resource');

    expect(res.status).toBe(200);
    expect(res.headers['content-type']).toMatch(/application\/json/);

    // Verify all required RFC 9728 fields
    expect(res.body).toHaveProperty('resource');
    expect(res.body).toHaveProperty('authorization_servers');
    expect(res.body).toHaveProperty('scopes_supported');
    expect(res.body).toHaveProperty('bearer_methods_supported');

    // resource should be a URL pointing to /mcp
    expect(res.body.resource).toBe('https://outlook-mcp.example.com/mcp');

    // authorization_servers should be an array
    expect(Array.isArray(res.body.authorization_servers)).toBe(true);
    expect(res.body.authorization_servers.length).toBe(1);

    // scopes_supported should be an array
    expect(Array.isArray(res.body.scopes_supported)).toBe(true);
    expect(res.body.scopes_supported.length).toBeGreaterThan(0);

    // bearer_methods_supported should include 'header'
    expect(res.body.bearer_methods_supported).toEqual(['header']);
  });

  // ═══════════════════════════════════════════════════════════════════════
  // PRM includes correct tenant
  // ═══════════════════════════════════════════════════════════════════════

  test('2. PRM authorization_servers URL contains the configured tenant ID', async () => {
    const res = await supertest(app)
      .get('/.well-known/oauth-protected-resource');

    expect(res.status).toBe(200);

    const authServer = res.body.authorization_servers[0];
    expect(authServer).toBe(
      `https://login.microsoftonline.com/${config.AUTH_CONFIG.tenantId}/v2.0`
    );
    expect(authServer).toContain('test-tenant-abc-123');
  });

  // ═══════════════════════════════════════════════════════════════════════
  // PRM includes correct scope
  // ═══════════════════════════════════════════════════════════════════════

  test('3. PRM scopes_supported contains apiAppId/apiScope', async () => {
    const res = await supertest(app)
      .get('/.well-known/oauth-protected-resource');

    expect(res.status).toBe(200);

    const expectedScope = `${config.CONNECTOR_AUTH.apiAppId}/${config.CONNECTOR_AUTH.apiScope}`;
    expect(res.body.scopes_supported).toContain(expectedScope);
    expect(res.body.scopes_supported).toContain('api://test-app-id-000/mcp.access');
  });

  // ═══════════════════════════════════════════════════════════════════════
  // 401 includes WWW-Authenticate header
  // ═══════════════════════════════════════════════════════════════════════

  test('4. Unauthenticated POST /mcp → 401 with WWW-Authenticate containing resource_metadata URL', async () => {
    const res = await supertest(app)
      .post('/mcp')
      .set('Content-Type', 'application/json')
      .send(initializeBody());

    expect(res.status).toBe(401);
    expect(res.body.error).toBe('auth_required');

    // Verify WWW-Authenticate header
    const wwwAuth = res.headers['www-authenticate'];
    expect(wwwAuth).toBeDefined();

    // Should start with Bearer
    expect(wwwAuth).toMatch(/^Bearer/);

    // Should contain realm
    expect(wwwAuth).toContain('realm="mcp"');

    // Should contain resource_metadata URL pointing to the PRM endpoint
    expect(wwwAuth).toContain('resource_metadata=');
    expect(wwwAuth).toContain('.well-known/oauth-protected-resource');

    // Should contain scope
    expect(wwwAuth).toContain('scope=');
    expect(wwwAuth).toContain('api://test-app-id-000/mcp.access');
  });

  test('4b. Invalid Bearer token → 401 with WWW-Authenticate header', async () => {
    const res = await supertest(app)
      .post('/mcp')
      .set('Authorization', 'Bearer invalid-token-xyz')
      .set('Content-Type', 'application/json')
      .send(initializeBody());

    expect(res.status).toBe(401);

    const wwwAuth = res.headers['www-authenticate'];
    expect(wwwAuth).toBeDefined();
    expect(wwwAuth).toContain('resource_metadata=');
  });

  // ═══════════════════════════════════════════════════════════════════════
  // PRM is publicly accessible
  // ═══════════════════════════════════════════════════════════════════════

  test('5. PRM endpoint requires no authentication', async () => {
    // No Authorization header → should still get 200 for PRM
    const res = await supertest(app)
      .get('/.well-known/oauth-protected-resource');

    expect(res.status).toBe(200);
    expect(res.body).toHaveProperty('resource');
    expect(res.body).toHaveProperty('authorization_servers');
  });

  test('5b. PRM is accessible even when /mcp requires auth', async () => {
    // Verify /mcp returns 401 (auth is required)
    const mcpRes = await supertest(app)
      .post('/mcp')
      .set('Content-Type', 'application/json')
      .send(initializeBody());

    expect(mcpRes.status).toBe(401);

    // But PRM should still be 200
    const prmRes = await supertest(app)
      .get('/.well-known/oauth-protected-resource');

    expect(prmRes.status).toBe(200);
    expect(prmRes.body.authorization_servers).toBeDefined();
  });

  // ═══════════════════════════════════════════════════════════════════════
  // PRM includes resource_name
  // ═══════════════════════════════════════════════════════════════════════

  test('6. PRM includes resource_name field', async () => {
    const res = await supertest(app)
      .get('/.well-known/oauth-protected-resource');

    expect(res.status).toBe(200);
    expect(res.body).toHaveProperty('resource_name');
    expect(typeof res.body.resource_name).toBe('string');
    expect(res.body.resource_name.length).toBeGreaterThan(0);
  });

  // ═══════════════════════════════════════════════════════════════════════
  // WWW-Authenticate challenge is consistent with PRM
  // ═══════════════════════════════════════════════════════════════════════

  test('7. WWW-Authenticate resource_metadata URL matches PRM endpoint', async () => {
    // Get the 401 response
    const authRes = await supertest(app)
      .post('/mcp')
      .set('Content-Type', 'application/json')
      .send(initializeBody());

    expect(authRes.status).toBe(401);

    const wwwAuth = authRes.headers['www-authenticate'];

    // Extract the resource_metadata URL from the header
    const match = wwwAuth.match(/resource_metadata="([^"]+)"/);
    expect(match).not.toBeNull();
    const metadataUrl = match[1];
    expect(metadataUrl).toContain('/.well-known/oauth-protected-resource');

    // The scope in WWW-Authenticate should match scopes_supported in PRM
    const prmRes = await supertest(app)
      .get('/.well-known/oauth-protected-resource');

    const scopeMatch = wwwAuth.match(/scope="([^"]+)"/);
    expect(scopeMatch).not.toBeNull();
    expect(prmRes.body.scopes_supported).toContain(scopeMatch[1]);
  });

  test('8. Canonical public base URL ignores spoofed Host and forwarded headers', async () => {
    const prmRes = await supertest(app)
      .get('/.well-known/oauth-protected-resource')
      .set('Host', 'attacker.example.com')
      .set('X-Forwarded-Host', 'attacker-proxy.example.com')
      .set('X-Forwarded-Proto', 'http');

    expect(prmRes.status).toBe(200);
    expect(prmRes.body.resource).toBe('https://outlook-mcp.example.com/mcp');

    const authRes = await supertest(app)
      .post('/mcp')
      .set('Host', 'attacker.example.com')
      .set('X-Forwarded-Host', 'attacker-proxy.example.com')
      .set('X-Forwarded-Proto', 'http')
      .set('Content-Type', 'application/json')
      .send(initializeBody());

    expect(authRes.status).toBe(401);
    expect(authRes.headers['www-authenticate']).toContain(
      'resource_metadata="https://outlook-mcp.example.com/.well-known/oauth-protected-resource"'
    );
  });
});
