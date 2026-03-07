/**
 * HTTP transport for the Outlook MCP Server.
 *
 * Wraps the MCP SDK's StreamableHTTPServerTransport in an Express app that:
 *   1. Authenticates requests via session tokens
 *   2. Sets per-request user context via AsyncLocalStorage
 *   3. Creates a per-request MCP Server + Transport (stateless mode)
 *
 * Usage:
 *   const { startHttpServer } = require('./transport/http-server');
 *   startHttpServer({ sessionStore });
 */

const express = require('express');
const http = require('http');
const { Server } = require('@modelcontextprotocol/sdk/server/index.js');
const { StreamableHTTPServerTransport } = require('@modelcontextprotocol/sdk/server/streamableHttp.js');
const { requestContext } = require('../auth/request-context');
const { createAuthRoutes } = require('../auth/auth-routes');
const { prmHandler, oauthMetadataHandler, buildWwwAuthenticateChallenge } = require('../auth/prm');
const jwtMiddleware = require('../auth/jwt-middleware');
const config = require('../config');

// Import module tools
const { authTools } = require('../auth');
const { calendarTools } = require('../calendar');
const { emailTools } = require('../email');
const { folderTools } = require('../folder');
const { rulesTools } = require('../rules');

// Combine all tools (same set used by stdio transport)
const TOOLS = [
  ...authTools,
  ...calendarTools,
  ...emailTools,
  ...folderTools,
  ...rulesTools,
];

function hasValidEntraPrincipal(req) {
  return !!(req.entraUser && typeof req.entraUser.oid === 'string' && req.entraUser.oid.length > 0);
}

/**
 * Create Express middleware that validates session tokens on incoming requests.
 *
 * Connector-aware: if the upstream JWT middleware already validated an Entra JWT
 * (setting req.entraUser), session validation is skipped entirely.
 *
 * Otherwise, extracts the token from the `Authorization: Bearer <token>` header,
 * validates it against the SessionStore, and populates `req.user` with the
 * session identity.  Returns 401 with a WWW-Authenticate challenge header when
 * authentication fails.
 *
 * @param {import('../auth/session-store')} sessionStore
 * @returns {import('express').RequestHandler}
 */
function createSessionMiddleware(sessionStore) {
  return (req, res, next) => {
    // Connector path: JWT middleware already validated the Entra token
    if (hasValidEntraPrincipal(req)) {
      return next();
    }

    // Session path: validate the Bearer token as a session token
    const authHeader = req.headers.authorization;
    if (!authHeader || !authHeader.startsWith('Bearer ')) {
      res.set('WWW-Authenticate', buildWwwAuthenticateChallenge(req));
      return res.status(401).json({
        error: 'auth_required',
        message: 'Session expired or missing. Authenticate at: /auth/login',
        authUrl: '/auth/login',
      });
    }

    const token = authHeader.slice(7); // Remove 'Bearer '
    const session = sessionStore.validateSession(token);

    if (!session) {
      res.set('WWW-Authenticate', buildWwwAuthenticateChallenge(req));
      return res.status(401).json({
        error: 'auth_required',
        message: 'Session expired or invalid. Re-authenticate at: /auth/login',
        authUrl: '/auth/login',
      });
    }

    // Set user identity on request for downstream use
    req.user = {
      id: session.userId,
      sessionToken: token,
    };

    next();
  };
}

/**
 * Build the fallbackRequestHandler that is attached to each per-request
 * MCP Server instance. This is the same handler logic used by the stdio
 * transport in index.js.
 */
function createFallbackRequestHandler() {
  return async (request) => {
    try {
      const { method, params, id } = request;
      console.error(`REQUEST: ${method} [${id}]`);

      // Initialize handler
      if (method === 'initialize') {
        console.error(`INITIALIZE REQUEST: ID [${id}]`);
        return {
          protocolVersion: '2024-11-05',
          capabilities: {
            tools: TOOLS.reduce((acc, tool) => {
              acc[tool.name] = {};
              return acc;
            }, {}),
          },
          serverInfo: { name: config.SERVER_NAME, version: config.SERVER_VERSION },
        };
      }

      // Tools list handler
      if (method === 'tools/list') {
        console.error(`TOOLS LIST REQUEST: ID [${id}]`);
        console.error(`TOOLS COUNT: ${TOOLS.length}`);
        return {
          tools: TOOLS.map((tool) => ({
            name: tool.name,
            description: tool.description,
            inputSchema: tool.inputSchema,
          })),
        };
      }

      // Required empty responses for other capabilities
      if (method === 'resources/list') return { resources: [] };
      if (method === 'prompts/list') return { prompts: [] };

      // Tool call handler
      if (method === 'tools/call') {
        try {
          const { name, arguments: args = {} } = params || {};
          console.error(`TOOL CALL: ${name}`);

          const tool = TOOLS.find((t) => t.name === name);
          if (tool && tool.handler) {
            return await tool.handler(args);
          }

          return {
            error: { code: -32601, message: `Tool not found: ${name}` },
          };
        } catch (error) {
          console.error(`Error in tools/call:`, error);
          return {
            error: { code: -32603, message: `Error processing tool call: ${error.message}` },
          };
        }
      }

      return {
        error: { code: -32601, message: `Method not found: ${method}` },
      };
    } catch (error) {
      console.error(`Error in fallbackRequestHandler:`, error);
      return {
        error: { code: -32603, message: `Error processing request: ${error.message}` },
      };
    }
  };
}

/**
 * Create and return the Express app (without starting it).
 * Exported separately for testing with supertest.
 *
 * @param {object} [opts]
 * @param {import('../auth/session-store')} [opts.sessionStore]  When provided,
 *   session-token auth middleware is applied to the `/mcp` route. Omit for
 *   unauthenticated usage (e.g. tests that don't need auth).
 * @param {import('../auth/per-user-token-storage')} [opts.tokenStorage]  When provided
 *   alongside sessionStore, browser auth routes are mounted at `/auth`.
 * @returns {express.Application}
 */
function createHttpApp({ sessionStore, tokenStorage } = {}) {
  const app = express();

  // ── Discovery endpoints — public, no auth ──────────────────────────
  app.get('/.well-known/oauth-protected-resource', prmHandler);
  app.get('/.well-known/oauth-authorization-server', oauthMetadataHandler);

  // ── Browser auth routes (login + callback) ────────────────────────
  if (sessionStore && tokenStorage) {
    const authRouter = createAuthRoutes({
      tokenStorage,
      sessionStore,
      config,
    });
    app.use('/auth', authRouter);
  }

  // ── JWT auth middleware (connector path) ────────────────────────
  app.use('/mcp', jwtMiddleware);

  // ── Session-token auth middleware (optional) ───────────────────────
  if (sessionStore) {
    app.use('/mcp', createSessionMiddleware(sessionStore));
  }

  // ── MCP request handler ──────────────────────────────────────────
  // POST, GET, DELETE all go through the same handler.
  // NOTE: We do NOT use express.json() — the StreamableHTTPServerTransport
  // handles body parsing internally.
  app.all('/mcp', async (req, res) => {
    const userCtx = hasValidEntraPrincipal(req)
      ? {
          userId: req.entraUser.oid,
          authMethod: 'connector',
          entraToken: req.entraToken,
        }
      : req.user?.id
        ? {
          userId: req.user?.id,
          authMethod: 'session',
          sessionToken: req.user?.sessionToken,
        }
        : null;

    await requestContext.run(userCtx, async () => {
      // Per-request transport (stateless mode)
      const transport = new StreamableHTTPServerTransport({
        sessionIdGenerator: undefined,
      });

      // Per-request MCP Server
      const server = new Server(
        { name: config.SERVER_NAME, version: config.SERVER_VERSION },
        {
          capabilities: {
            tools: TOOLS.reduce((acc, tool) => {
              acc[tool.name] = {};
              return acc;
            }, {}),
          },
        }
      );

      server.fallbackRequestHandler = createFallbackRequestHandler();

      await server.connect(transport);
      await transport.handleRequest(req, res, req.body);
    });
  });

  return app;
}

/**
 * Start the HTTP server.
 *
 * @param {object} [opts]
 * @param {import('../auth/session-store')} [opts.sessionStore]  Passed through to createHttpApp.
 * @param {import('../auth/per-user-token-storage')} [opts.tokenStorage]  Passed through to createHttpApp.
 * @returns {http.Server} The Node http.Server instance (useful for tests / graceful shutdown).
 */
function startHttpServer({ sessionStore, tokenStorage } = {}) {
  const app = createHttpApp({ sessionStore, tokenStorage });
  const port = process.env.PORT || 3000;

  const httpServer = app.listen(port, () => {
    console.error(`${config.SERVER_NAME} HTTP transport listening on port ${port}`);
  });

  // Graceful shutdown
  process.on('SIGTERM', () => {
    console.error('SIGTERM received — closing HTTP server');
    httpServer.close(() => {
      console.error('HTTP server closed');
    });
  });

  return httpServer;
}

module.exports = { createHttpApp, startHttpServer, createSessionMiddleware };
