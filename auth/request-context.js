/**
 * Per-request context using AsyncLocalStorage.
 *
 * Allows tool handlers (and anything else on the call stack) to retrieve the
 * authenticated user's identity without threading it through every function.
 *
 * Usage:
 *   const { requestContext, getUserContext } = require('./request-context');
 *
 *   // Wrap a request handler:
 *   requestContext.run({ userId: 'oid-123', entraToken: 'jwt...' }, () => {
 *     // ... inside here, getUserContext() returns the context
 *   });
 *
 *   // Read context from anywhere on the call stack:
 *   const ctx = getUserContext(); // { userId, entraToken } or null
 */

const { AsyncLocalStorage } = require('node:async_hooks');

const requestContext = new AsyncLocalStorage();

/**
 * Returns the user context for the current request, or null if none is set.
 * @returns {{ userId: string, entraToken: string } | null}
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
  return getUserContext() !== null;
}

module.exports = { requestContext, getUserContext, isHostedMode };
