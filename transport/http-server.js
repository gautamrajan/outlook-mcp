/**
 * HTTP transport for the Outlook MCP Server.
 *
 * Wraps the MCP SDK's StreamableHTTPServerTransport in an Express app that:
 *   1. Validates Entra ID bearer tokens on every /mcp request
 *   2. Sets per-request user context via AsyncLocalStorage
 *   3. Creates a per-request MCP Server + Transport (stateless mode)
 *
 * Usage:
 *   const { startHttpServer } = require('./transport/http-server');
 *   startHttpServer();
 */

const express = require('express');
const http = require('http');
const { Server } = require('@modelcontextprotocol/sdk/server/index.js');
const { StreamableHTTPServerTransport } = require('@modelcontextprotocol/sdk/server/streamableHttp.js');
const { createEntraMiddleware } = require('../auth/entra-middleware');
const { requestContext } = require('../auth/request-context');
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
 * @returns {express.Application}
 */
function createHttpApp() {
  const app = express();

  // ── Entra JWT middleware on the MCP route ─────────────────────────
  const entraMiddleware = createEntraMiddleware({
    tenantId: config.AUTH_CONFIG.tenantId,
    clientId: config.AUTH_CONFIG.clientId,
  });

  // Apply auth to all MCP methods
  app.use('/mcp', entraMiddleware);

  // ── MCP request handler ──────────────────────────────────────────
  // POST, GET, DELETE all go through the same handler.
  // NOTE: We do NOT use express.json() — the StreamableHTTPServerTransport
  // handles body parsing internally.
  app.all('/mcp', async (req, res) => {
    // Wrap in AsyncLocalStorage so tool handlers can access user identity
    const userCtx = {
      userId: req.user.id,
      entraToken: req.user.token,
    };

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
 * @returns {http.Server} The Node http.Server instance (useful for tests / graceful shutdown).
 */
function startHttpServer() {
  const app = createHttpApp();
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

module.exports = { createHttpApp, startHttpServer };
