/**
 * Tests for config.js
 *
 * Since config.js reads environment variables at require time,
 * we use jest.resetModules() and set env vars before re-requiring.
 */

describe('config', () => {
  const originalEnv = process.env;

  beforeEach(() => {
    // Shallow clone so we can mutate without affecting originals
    process.env = { ...originalEnv };
    // Clear the module cache so config.js re-reads env vars on next require
    jest.resetModules();
  });

  afterAll(() => {
    process.env = originalEnv;
  });

  // ── Existing config values still present ────────────────────────

  test('should export SERVER_NAME', () => {
    const config = require('../config');
    expect(config.SERVER_NAME).toBe('outlook-assistant');
  });

  test('should export SERVER_VERSION', () => {
    const config = require('../config');
    expect(config.SERVER_VERSION).toBe('1.0.0');
  });

  test('should export AUTH_CONFIG with expected keys', () => {
    const config = require('../config');
    expect(config.AUTH_CONFIG).toBeDefined();
    expect(config.AUTH_CONFIG).toHaveProperty('clientId');
    expect(config.AUTH_CONFIG).toHaveProperty('clientSecret');
    expect(config.AUTH_CONFIG).toHaveProperty('tenantId');
    expect(config.AUTH_CONFIG).toHaveProperty('tokenEndpoint');
    expect(config.AUTH_CONFIG).toHaveProperty('redirectUri');
    expect(config.AUTH_CONFIG).toHaveProperty('scopes');
    expect(config.AUTH_CONFIG).toHaveProperty('tokenStorePath');
    expect(config.AUTH_CONFIG).toHaveProperty('authServerUrl');
  });

  test('should export GRAPH_API_ENDPOINT', () => {
    const config = require('../config');
    expect(config.GRAPH_API_ENDPOINT).toBe('https://graph.microsoft.com/v1.0/');
  });

  test('should export EMAIL_SELECT_FIELDS', () => {
    const config = require('../config');
    expect(config.EMAIL_SELECT_FIELDS).toBeDefined();
    expect(typeof config.EMAIL_SELECT_FIELDS).toBe('string');
  });

  test('should export CALENDAR_SELECT_FIELDS', () => {
    const config = require('../config');
    expect(config.CALENDAR_SELECT_FIELDS).toBeDefined();
    expect(typeof config.CALENDAR_SELECT_FIELDS).toBe('string');
  });

  test('should export DEFAULT_PAGE_SIZE and MAX_RESULT_COUNT', () => {
    const config = require('../config');
    expect(config.DEFAULT_PAGE_SIZE).toBe(25);
    expect(config.MAX_RESULT_COUNT).toBe(50);
  });

  test('should export attachment download limits', () => {
    const config = require('../config');
    expect(config.ATTACHMENT_DOWNLOAD_TTL_MS).toBe(5 * 60 * 1000);
    expect(config.MAX_ATTACHMENT_DOWNLOAD_BYTES).toBe(25 * 1024 * 1024);
  });

  test('should export DEFAULT_TIMEZONE', () => {
    const config = require('../config');
    expect(config.DEFAULT_TIMEZONE).toBeDefined();
    expect(typeof config.DEFAULT_TIMEZONE).toBe('string');
  });

  // ── MCP_TRANSPORT ───────────────────────────────────────────────

  test('MCP_TRANSPORT should default to "stdio"', () => {
    delete process.env.MCP_TRANSPORT;
    const config = require('../config');
    expect(config.MCP_TRANSPORT).toBe('stdio');
  });

  test('MCP_TRANSPORT should reflect env var when set', () => {
    process.env.MCP_TRANSPORT = 'http';
    const config = require('../config');
    expect(config.MCP_TRANSPORT).toBe('http');
  });

  // ── PORT ────────────────────────────────────────────────────────

  test('PORT should default to 3000', () => {
    delete process.env.PORT;
    const config = require('../config');
    expect(config.PORT).toBe(3000);
  });

  test('PORT should parse env var as integer', () => {
    process.env.PORT = '8080';
    const config = require('../config');
    expect(config.PORT).toBe(8080);
  });

  // ── HOSTED ─────────────────────────────────────────────────────

  test('HOSTED.enabled should be false when MCP_TRANSPORT is not "http"', () => {
    delete process.env.MCP_TRANSPORT;
    const config = require('../config');
    expect(config.HOSTED.enabled).toBe(false);
  });

  test('HOSTED.enabled should be false when MCP_TRANSPORT is "stdio"', () => {
    process.env.MCP_TRANSPORT = 'stdio';
    const config = require('../config');
    expect(config.HOSTED.enabled).toBe(false);
  });

  test('HOSTED.enabled should be true when MCP_TRANSPORT is "http"', () => {
    process.env.MCP_TRANSPORT = 'http';
    const config = require('../config');
    expect(config.HOSTED.enabled).toBe(true);
  });

  test('HOSTED.enabled should be true when MCP_TRANSPORT is "HTTP" (case-insensitive)', () => {
    process.env.MCP_TRANSPORT = 'HTTP';
    const config = require('../config');
    expect(config.HOSTED.enabled).toBe(true);
  });

  test('HOSTED should have the expected shape', () => {
    const config = require('../config');
    expect(config.HOSTED).toHaveProperty('enabled');
    expect(config.HOSTED).toHaveProperty('tokenEncryptionKey');
    expect(config.HOSTED).toHaveProperty('tokenStorePath');
    expect(config.HOSTED).toHaveProperty('sessionStorePath');
    expect(config.HOSTED).toHaveProperty('publicBaseUrl');
    expect(config.HOSTED).toHaveProperty('hostedRedirectUri');
    expect(config.HOSTED).toHaveProperty('sessionExpirationDays');
  });

  test('HOSTED.tokenEncryptionKey should default to empty string', () => {
    delete process.env.TOKEN_ENCRYPTION_KEY;
    const config = require('../config');
    expect(config.HOSTED.tokenEncryptionKey).toBe('');
  });

  test('HOSTED.tokenEncryptionKey should reflect env var', () => {
    process.env.TOKEN_ENCRYPTION_KEY = 'my-secret-key';
    const config = require('../config');
    expect(config.HOSTED.tokenEncryptionKey).toBe('my-secret-key');
  });

  test('HOSTED.tokenStorePath should default to home dir path', () => {
    delete process.env.TOKEN_STORE_PATH;
    const config = require('../config');
    expect(config.HOSTED.tokenStorePath).toMatch(/\.outlook-mcp-hosted-tokens\.json$/);
  });

  test('HOSTED.tokenStorePath should reflect env var', () => {
    process.env.TOKEN_STORE_PATH = '/custom/path/tokens.json';
    const config = require('../config');
    expect(config.HOSTED.tokenStorePath).toBe('/custom/path/tokens.json');
  });

  test('HOSTED.sessionStorePath should default to home dir path', () => {
    delete process.env.SESSION_STORE_PATH;
    const config = require('../config');
    expect(config.HOSTED.sessionStorePath).toMatch(/\.outlook-mcp-sessions\.json$/);
  });

  test('HOSTED.sessionStorePath should reflect env var', () => {
    process.env.SESSION_STORE_PATH = '/custom/path/sessions.json';
    const config = require('../config');
    expect(config.HOSTED.sessionStorePath).toBe('/custom/path/sessions.json');
  });

  test('HOSTED.hostedRedirectUri should default to empty string', () => {
    delete process.env.HOSTED_REDIRECT_URI;
    const config = require('../config');
    expect(config.HOSTED.hostedRedirectUri).toBe('');
  });

  test('HOSTED.hostedRedirectUri should reflect env var', () => {
    process.env.HOSTED_REDIRECT_URI = 'https://myserver.com/auth/callback';
    const config = require('../config');
    expect(config.HOSTED.hostedRedirectUri).toBe('https://myserver.com/auth/callback');
  });

  test('HOSTED.publicBaseUrl should default to empty string', () => {
    delete process.env.PUBLIC_BASE_URL;
    const config = require('../config');
    expect(config.HOSTED.publicBaseUrl).toBe('');
  });

  test('HOSTED.publicBaseUrl should reflect env var', () => {
    process.env.PUBLIC_BASE_URL = 'https://public.example.com';
    const config = require('../config');
    expect(config.HOSTED.publicBaseUrl).toBe('https://public.example.com');
  });

  test('HOSTED.sessionExpirationDays should default to 14', () => {
    delete process.env.HOSTED_SESSION_EXPIRATION_DAYS;
    const config = require('../config');
    expect(config.HOSTED.sessionExpirationDays).toBe(14);
  });

  test('HOSTED.sessionExpirationDays should parse positive integer env var', () => {
    process.env.HOSTED_SESSION_EXPIRATION_DAYS = '30';
    const config = require('../config');
    expect(config.HOSTED.sessionExpirationDays).toBe(30);
  });

  test('HOSTED.sessionExpirationDays should fall back to 14 for invalid values', () => {
    process.env.HOSTED_SESSION_EXPIRATION_DAYS = '0';
    const config = require('../config');
    expect(config.HOSTED.sessionExpirationDays).toBe(14);
  });

  // ── AUTH_CONFIG hosted fields ─────────────────────────────────

  test('AUTH_CONFIG.hostedRedirectUri should default to empty string', () => {
    delete process.env.HOSTED_REDIRECT_URI;
    const config = require('../config');
    expect(config.AUTH_CONFIG.hostedRedirectUri).toBe('');
  });

  test('AUTH_CONFIG.hostedRedirectUri should reflect env var', () => {
    process.env.HOSTED_REDIRECT_URI = 'https://myserver.com/auth/callback';
    const config = require('../config');
    expect(config.AUTH_CONFIG.hostedRedirectUri).toBe('https://myserver.com/auth/callback');
  });

  test('AUTH_CONFIG.hostedTokenStorePath should default to home dir path', () => {
    delete process.env.TOKEN_STORE_PATH;
    const config = require('../config');
    expect(config.AUTH_CONFIG.hostedTokenStorePath).toMatch(/\.outlook-mcp-hosted-tokens\.json$/);
  });

  test('AUTH_CONFIG.hostedTokenStorePath should reflect env var', () => {
    process.env.TOKEN_STORE_PATH = '/custom/path/tokens.json';
    const config = require('../config');
    expect(config.AUTH_CONFIG.hostedTokenStorePath).toBe('/custom/path/tokens.json');
  });
});
