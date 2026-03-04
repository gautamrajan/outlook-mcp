const http = require('http');
const querystring = require('querystring');
const config = require('../config');

const AUTH_TIMEOUT_MS = 5 * 60 * 1000;
const DEFAULT_PORT = 3333;

let _server = null;
let _state = 'idle'; // idle | waiting | authenticated | error
let _authCompleteResolve = null;
let _authCompletePromise = null;
let _timeoutHandle = null;
let _activePort = null;

function _getTokenStorage() {
  const { tokenStorage } = require('./index');
  return tokenStorage;
}

function _buildAuthorizationUrl(clientId) {
  const authParams = {
    client_id: clientId,
    response_type: 'code',
    redirect_uri: config.AUTH_CONFIG.redirectUri,
    scope: config.AUTH_CONFIG.scopes.join(' '),
    response_mode: 'query',
    state: Date.now().toString()
  };

  const tenantId = config.AUTH_CONFIG.tenantId || 'common';
  return `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize?${querystring.stringify(authParams)}`;
}

function _handleRequest(req, res) {
  const parsedUrl = new URL(req.url, `http://localhost`);
  const pathname = parsedUrl.pathname;

  console.error(`[embedded-auth] Request: ${pathname}`);

  if (pathname === '/auth') {
    _handleAuthRedirect(parsedUrl, res);
  } else if (pathname === '/auth/callback') {
    _handleCallback(parsedUrl, res);
  } else if (pathname === '/') {
    res.writeHead(200, { 'Content-Type': 'text/html' });
    res.end(`<html><body style="font-family:sans-serif;text-align:center;margin-top:60px;">
      <h2>Outlook MCP Auth Server</h2>
      <p>This server is handling authentication. Use the <code>authenticate</code> tool in Claude to start.</p>
    </body></html>`);
  } else {
    res.writeHead(404, { 'Content-Type': 'text/plain' });
    res.end('Not Found');
  }
}

function _handleAuthRedirect(parsedUrl, res) {
  const { clientId, clientSecret } = config.AUTH_CONFIG;
  if (!clientId || !clientSecret) {
    res.writeHead(500, { 'Content-Type': 'text/html' });
    res.end(`<html><body style="font-family:sans-serif;text-align:center;margin-top:60px;">
      <h1 style="color:#d9534f;">Configuration Error</h1>
      <p>Client ID or Client Secret is not set. Check your environment variables.</p>
    </body></html>`);
    return;
  }

  const queryClientId = parsedUrl.searchParams.get('client_id') || clientId;
  const authUrl = _buildAuthorizationUrl(queryClientId);
  console.error(`[embedded-auth] Redirecting to Microsoft login`);
  res.writeHead(302, { 'Location': authUrl });
  res.end();
}

async function _handleCallback(parsedUrl, res) {
  const error = parsedUrl.searchParams.get('error');
  const errorDescription = parsedUrl.searchParams.get('error_description');
  const code = parsedUrl.searchParams.get('code');

  if (error) {
    console.error(`[embedded-auth] Auth error: ${error} - ${errorDescription}`);
    _state = 'error';
    res.writeHead(400, { 'Content-Type': 'text/html' });
    res.end(`<html><body style="font-family:sans-serif;text-align:center;margin-top:60px;">
      <h1 style="color:#d9534f;">Authentication Error</h1>
      <p><strong>${error}</strong></p>
      <p>${errorDescription || ''}</p>
      <p>Close this window and try again in Claude.</p>
    </body></html>`);
    _resolveAuth(false);
    return;
  }

  if (!code) {
    res.writeHead(400, { 'Content-Type': 'text/html' });
    res.end(`<html><body style="font-family:sans-serif;text-align:center;margin-top:60px;">
      <h1 style="color:#d9534f;">Missing Authorization Code</h1>
      <p>No authorization code was provided. Close this window and try again.</p>
    </body></html>`);
    return;
  }

  try {
    console.error('[embedded-auth] Exchanging authorization code for tokens...');
    const tokenStorage = _getTokenStorage();
    await tokenStorage.exchangeCodeForTokens(code);

    _state = 'authenticated';
    console.error('[embedded-auth] Authentication successful, tokens saved.');

    res.writeHead(200, { 'Content-Type': 'text/html' });
    res.end(`<html><body style="font-family:sans-serif;text-align:center;margin-top:60px;">
      <h1 style="color:#5cb85c;">Authentication Successful</h1>
      <p>Tokens have been saved. You can close this window and return to Claude.</p>
    </body></html>`);

    _resolveAuth(true);
    setTimeout(() => stopAuthServer(), 1000);
  } catch (err) {
    console.error(`[embedded-auth] Token exchange failed: ${err.message}`);
    _state = 'error';
    res.writeHead(500, { 'Content-Type': 'text/html' });
    res.end(`<html><body style="font-family:sans-serif;text-align:center;margin-top:60px;">
      <h1 style="color:#d9534f;">Token Exchange Failed</h1>
      <p>${err.message}</p>
      <p>Close this window and try again in Claude.</p>
    </body></html>`);
    _resolveAuth(false);
  }
}

function _resolveAuth(success) {
  if (_authCompleteResolve) {
    _authCompleteResolve(success);
    _authCompleteResolve = null;
  }
  if (_timeoutHandle) {
    clearTimeout(_timeoutHandle);
    _timeoutHandle = null;
  }
}

/**
 * Start the embedded auth HTTP server.
 * Returns the base URL (e.g. http://localhost:3333).
 * If the server is already running, returns the existing URL.
 */
function startAuthServer(port = DEFAULT_PORT) {
  if (_server && _server.listening) {
    console.error(`[embedded-auth] Server already running on port ${_activePort}`);
    _state = 'waiting';
    return Promise.resolve(`http://localhost:${_activePort}`);
  }

  _state = 'waiting';
  _authCompletePromise = new Promise((resolve) => {
    _authCompleteResolve = resolve;
  });

  return new Promise((resolve, reject) => {
    _server = http.createServer(_handleRequest);

    _server.on('error', (err) => {
      if (err.code === 'EADDRINUSE') {
        console.error(`[embedded-auth] Port ${port} in use, trying ${port + 1}`);
        _server.close();
        resolve(startAuthServer(port + 1));
      } else {
        _state = 'error';
        reject(err);
      }
    });

    _server.listen(port, () => {
      _activePort = port;
      console.error(`[embedded-auth] Auth server started on port ${port}`);

      _timeoutHandle = setTimeout(() => {
        console.error('[embedded-auth] Auth timeout reached, stopping server.');
        _state = 'idle';
        _resolveAuth(false);
        stopAuthServer();
      }, AUTH_TIMEOUT_MS);

      resolve(`http://localhost:${port}`);
    });
  });
}

/**
 * Stop the embedded auth HTTP server.
 */
function stopAuthServer() {
  if (_timeoutHandle) {
    clearTimeout(_timeoutHandle);
    _timeoutHandle = null;
  }
  if (_server) {
    _server.close(() => {
      console.error('[embedded-auth] Auth server stopped.');
    });
    _server = null;
    _activePort = null;
  }
  if (_state === 'waiting') {
    _state = 'idle';
  }
}

/**
 * Get current server state.
 * @returns {{ state: string, port: number|null, authComplete: Promise<boolean>|null }}
 */
function getServerStatus() {
  return {
    state: _state,
    port: _activePort,
    running: !!(_server && _server.listening),
    authComplete: _authCompletePromise
  };
}

module.exports = { startAuthServer, stopAuthServer, getServerStatus };
