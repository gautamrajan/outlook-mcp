/**
 * Per-request context using AsyncLocalStorage.
 *
 * Allows tool handlers (and anything else on the call stack) to retrieve the
 * authenticated user's identity without threading it through every function.
 *
 * The context shape depends on the authentication method:
 *
 *   Session-based auth:
 *     { userId: string, authMethod: 'session', sessionToken: string, serverBaseUrl?: string }
 *
 *   Connector (Entra JWT) auth:
 *     { userId: string, authMethod: 'connector', entraToken: string, serverBaseUrl?: string }
 *
 * Usage:
 *   const { requestContext, getUserContext } = require('./request-context');
 *
 *   // Wrap a request handler:
 *   requestContext.run({ userId: 'oid-123', authMethod: 'connector', entraToken: 'jwt...', serverBaseUrl: 'https://mcp.example.com' }, () => {
 *     // ... inside here, getUserContext() returns the context
 *   });
 *
 *   // Read context from anywhere on the call stack:
 *   const ctx = getUserContext();
 *   // { userId, authMethod, entraToken, serverBaseUrl } or { userId, authMethod, sessionToken, serverBaseUrl } or null
 */

const { AsyncLocalStorage } = require('node:async_hooks');

const requestContext = new AsyncLocalStorage();

/**
 * Returns the user context for the current request, or null if none is set.
 * @returns {{ userId: string, authMethod: 'session', sessionToken: string, serverBaseUrl?: string } | { userId: string, authMethod: 'connector', entraToken: string, serverBaseUrl?: string } | null}
 */
function getUserContext() {
  return requestContext.getStore() || null;
}

/**
 * Returns true if the server is operating in hosted multi-user mode
 * (i.e., a user context is present on the current async call stack).
 * @returns {boolean}
 */
function isHostedMode() {
  const ctx = getUserContext();
  return !!(ctx && typeof ctx.userId === 'string' && ctx.userId.length > 0);
}

/**
 * Returns the authentication method for the current request, or null if
 * no user context is set.
 * @returns {'session' | 'connector' | null}
 */
function getAuthMethod() {
  const ctx = getUserContext();
  return ctx?.authMethod || null;
}

module.exports = { requestContext, getUserContext, isHostedMode, getAuthMethod };
