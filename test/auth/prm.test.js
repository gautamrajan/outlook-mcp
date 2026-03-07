/**
 * Tests for auth/prm.js — Protected Resource Metadata (RFC 9728).
 *
 * Validates:
 *   - The PRM endpoint returns correct JSON structure
 *   - Authorization server URL includes the tenant ID
 *   - Scopes are correctly constructed from config values
 *   - Content-Type header is application/json
 *   - The WWW-Authenticate challenge helper produces correct output
 */
const request = require('supertest');
const express = require('express');

// ── Mocks ────────────────────────────────────────────────────────────

jest.mock('../../config', () => ({
  HOSTED: {
    publicBaseUrl: 'https://outlook-mcp.example.com',
  },
  AUTH_CONFIG: {
    tenantId: 'test-tenant-abc-123',
    hostedRedirectUri: 'https://outlook-mcp.example.com/auth/callback',
  },
  CONNECTOR_AUTH: {
    apiAppId: 'api://test-app-id-000',
    apiScope: 'mcp.access',
  },
}));

// ── Import under test (after mocks) ──────────────────────────────────

const {
  prmHandler,
  buildWwwAuthenticateChallenge,
  getBaseUrl,
  getConfiguredServerBaseUrl,
  getRequestBaseUrl,
  getFullScope,
} = require('../../auth/prm');

// ── Helpers ──────────────────────────────────────────────────────────

/**
 * Build a minimal Express app with the PRM route for supertest.
 */
function createTestApp() {
  const app = express();
  app.get('/.well-known/oauth-protected-resource', prmHandler);
  return app;
}

/**
 * Build a fake Express request object for unit-testing helpers.
 */
function fakeReq(overrides = {}) {
  const headers = {
    host: 'example.com',
    ...overrides.headers,
  };
  return {
    protocol: overrides.protocol || 'https',
    get: (name) => headers[name.toLowerCase()] || undefined,
    ...overrides,
  };
}

// ── Tests ────────────────────────────────────────────────────────────

describe('Protected Resource Metadata (auth/prm.js)', () => {
  let app;

  beforeAll(() => {
    app = createTestApp();
  });

  // ── PRM endpoint ─────────────────────────────────────────────────

  describe('GET /.well-known/oauth-protected-resource', () => {
    test('returns 200 with correct JSON structure', async () => {
      const res = await request(app)
        .get('/.well-known/oauth-protected-resource');

      expect(res.status).toBe(200);
      expect(res.body).toHaveProperty('resource');
      expect(res.body).toHaveProperty('resource_name', 'MRC Outlook Assistant');
      expect(res.body).toHaveProperty('authorization_servers');
      expect(res.body).toHaveProperty('scopes_supported');
      expect(res.body).toHaveProperty('bearer_methods_supported');
      expect(res.body.bearer_methods_supported).toEqual(['header']);
    });

    test('includes correct authorization_servers URL with tenant ID', async () => {
      const res = await request(app)
        .get('/.well-known/oauth-protected-resource');

      expect(res.body.authorization_servers).toEqual([
        'https://login.microsoftonline.com/test-tenant-abc-123/v2.0',
      ]);
    });

    test('includes correct scopes_supported from config', async () => {
      const res = await request(app)
        .get('/.well-known/oauth-protected-resource');

      expect(res.body.scopes_supported).toEqual([
        'api://test-app-id-000/mcp.access',
      ]);
    });

    test('has Content-Type application/json', async () => {
      const res = await request(app)
        .get('/.well-known/oauth-protected-resource');

      expect(res.headers['content-type']).toMatch(/application\/json/);
    });

    test('resource field points to /mcp on the server', async () => {
      const res = await request(app)
        .get('/.well-known/oauth-protected-resource');

      // supertest uses 127.0.0.1 — just verify it ends with /mcp
      expect(res.body.resource).toMatch(/\/mcp$/);
    });
  });

  // ── getBaseUrl helper ────────────────────────────────────────────

  describe('getBaseUrl()', () => {
    test('prefers configured canonical base URL over request headers', () => {
      const req = fakeReq({
        protocol: 'http',
        headers: {
          host: 'localhost:3000',
          'x-forwarded-proto': 'https',
          'x-forwarded-host': 'attacker.example.com',
        },
      });

      expect(getBaseUrl(req)).toBe('https://outlook-mcp.example.com');
    });

    test('configured server base URL is derived from config', () => {
      expect(getConfiguredServerBaseUrl()).toBe('https://outlook-mcp.example.com');
    });

    test('request base URL uses req.protocol and Host header only', () => {
      const req = fakeReq({
        protocol: 'http',
        headers: {
          host: 'localhost:3000',
          'x-forwarded-proto': 'https',
          'x-forwarded-host': 'attacker.example.com',
        },
      });

      expect(getRequestBaseUrl(req)).toBe('http://localhost:3000');
    });
  });

  // ── getFullScope helper ──────────────────────────────────────────

  describe('getFullScope()', () => {
    test('constructs scope as apiAppId/apiScope', () => {
      expect(getFullScope()).toBe('api://test-app-id-000/mcp.access');
    });
  });

  // ── WWW-Authenticate challenge ───────────────────────────────────

  describe('buildWwwAuthenticateChallenge()', () => {
    test('generates correct WWW-Authenticate header value', () => {
      const req = fakeReq({
        protocol: 'https',
        headers: { host: 'myserver.example.com' },
      });

      const challenge = buildWwwAuthenticateChallenge(req);

      expect(challenge).toBe(
        'Bearer realm="mcp", ' +
        'resource_metadata="https://outlook-mcp.example.com/.well-known/oauth-protected-resource", ' +
        'scope="api://test-app-id-000/mcp.access"'
      );
    });

    test('ignores forwarded headers when a canonical base URL is configured', () => {
      const req = fakeReq({
        protocol: 'http',
        headers: {
          host: 'localhost:3000',
          'x-forwarded-proto': 'https',
          'x-forwarded-host': 'prod.example.com',
        },
      });

      const challenge = buildWwwAuthenticateChallenge(req);

      expect(challenge).toContain('resource_metadata="https://outlook-mcp.example.com/.well-known/oauth-protected-resource"');
      expect(challenge).toContain('scope="api://test-app-id-000/mcp.access"');
      expect(challenge).toMatch(/^Bearer realm="mcp"/);
    });

    test('includes all three required fields: realm, resource_metadata, scope', () => {
      const req = fakeReq();
      const challenge = buildWwwAuthenticateChallenge(req);

      expect(challenge).toMatch(/^Bearer realm="mcp"/);
      expect(challenge).toMatch(/resource_metadata="[^"]+"/);
      expect(challenge).toMatch(/scope="[^"]+"/);
    });
  });
});
