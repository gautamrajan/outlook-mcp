/**
 * Browser-based auth flow routes for hosted mode.
 *
 * Provides /auth/login and /auth/callback endpoints that implement the
 * OAuth 2.0 authorization code flow with PKCE against Microsoft Entra ID.
 *
 * Users visit /auth/login once in a browser, authenticate with Microsoft,
 * and receive a session token they can use with the MCP server.
 *
 * Dependencies are injected for testability.
 */

const crypto = require('crypto');
const express = require('express');

// ── PKCE helpers ────────────────────────────────────────────────────────

/**
 * Generates a PKCE code verifier and S256 challenge.
 * @returns {{ verifier: string, challenge: string }}
 */
function generatePKCE() {
  const verifier = crypto.randomBytes(32).toString('base64url');
  const challenge = crypto.createHash('sha256').update(verifier).digest('base64url');
  return { verifier, challenge };
}

// ── Pending auth state (in-memory, per-process) ─────────────────────────

const PENDING_AUTH_TTL_MS = 10 * 60 * 1000; // 10 minutes

/**
 * In-memory store for pending authorization requests.
 * Keyed by the random `state` parameter, each entry holds the PKCE
 * code verifier and an expiry timestamp.
 *
 * @type {Map<string, { codeVerifier: string, expiresAt: number }>}
 */
const pendingAuth = new Map();

/**
 * Lazily removes expired entries from the pending auth map.
 * Called on each /auth/login request to keep memory bounded.
 */
function cleanupExpiredPendingAuth() {
  const now = Date.now();
  for (const [state, entry] of pendingAuth) {
    if (entry.expiresAt <= now) {
      pendingAuth.delete(state);
    }
  }
}

// ── Route factory ───────────────────────────────────────────────────────

/**
 * Creates an Express Router with /auth/login and /auth/callback routes.
 *
 * @param {object} deps
 * @param {import('./per-user-token-storage')} deps.tokenStorage
 * @param {import('./session-store')} deps.sessionStore
 * @param {object} deps.config — the config object from config.js
 * @param {Function} [deps.fetch] — optional fetch override for testing
 * @returns {import('express').Router}
 */
function createAuthRoutes({ tokenStorage, sessionStore, config, fetch: fetchFn }) {
  const router = express.Router();
  const _fetch = fetchFn || globalThis.fetch;

  const authConfig = config.AUTH_CONFIG;
  const tenantId = authConfig.tenantId;
  const clientId = authConfig.clientId;
  const clientSecret = authConfig.clientSecret;
  const scopes = authConfig.scopes;
  const redirectUri = authConfig.hostedRedirectUri || authConfig.redirectUri;

  // ── GET /login ─────────────────────────────────────────────────────

  router.get('/login', (req, res) => {
    // Lazy cleanup of expired pending auth entries
    cleanupExpiredPendingAuth();

    // Generate PKCE pair
    const { verifier, challenge } = generatePKCE();

    // Generate random state for CSRF protection
    const state = crypto.randomBytes(16).toString('hex');

    // Store pending auth entry with TTL
    pendingAuth.set(state, {
      codeVerifier: verifier,
      expiresAt: Date.now() + PENDING_AUTH_TTL_MS,
    });

    // Build the Entra authorize URL
    const params = new URLSearchParams({
      client_id: clientId,
      response_type: 'code',
      redirect_uri: redirectUri,
      scope: scopes.join(' '),
      state,
      code_challenge: challenge,
      code_challenge_method: 'S256',
    });

    const authorizeUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize?${params.toString()}`;

    res.redirect(authorizeUrl);
  });

  // ── GET /callback ──────────────────────────────────────────────────

  router.get('/callback', async (req, res) => {
    const { code, state } = req.query;

    // Validate required params
    if (!code) {
      return res.status(400).send('Missing authorization code');
    }
    if (!state) {
      return res.status(400).send('Missing state parameter');
    }

    // Validate state against pending auth map (CSRF check)
    const pending = pendingAuth.get(state);
    if (!pending) {
      return res.status(400).send('Invalid or expired state parameter');
    }

    // Check if the pending entry has expired
    if (pending.expiresAt <= Date.now()) {
      pendingAuth.delete(state);
      return res.status(400).send('Invalid or expired state parameter');
    }

    // Single-use: delete the entry
    const { codeVerifier } = pending;
    pendingAuth.delete(state);

    try {
      // Exchange auth code for tokens
      const tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
      const tokenBody = new URLSearchParams({
        client_id: clientId,
        client_secret: clientSecret,
        grant_type: 'authorization_code',
        code,
        redirect_uri: redirectUri,
        code_verifier: codeVerifier,
      });

      const tokenResponse = await _fetch(tokenUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: tokenBody.toString(),
      });

      if (!tokenResponse.ok) {
        const errorText = await tokenResponse.text();
        console.error('Token exchange failed:', tokenResponse.status, errorText);
        return res.status(500).send('Failed to exchange authorization code for tokens');
      }

      const tokenData = await tokenResponse.json();
      const { access_token, refresh_token, expires_in, scope } = tokenData;

      // Fetch user profile from Microsoft Graph
      const meResponse = await _fetch('https://graph.microsoft.com/v1.0/me', {
        headers: { Authorization: `Bearer ${access_token}` },
      });

      if (!meResponse.ok) {
        console.error('Graph /me request failed:', meResponse.status);
        return res.status(500).send('Failed to retrieve user profile');
      }

      const userProfile = await meResponse.json();
      const userId = userProfile.id;
      const userEmail = userProfile.mail || userProfile.userPrincipalName || null;
      const userName = userProfile.displayName || null;

      // Store tokens
      await tokenStorage.setTokensForUser(userId, {
        accessToken: access_token,
        refreshToken: refresh_token,
        expiresIn: expires_in,
        scopes: scope,
        email: userEmail,
        name: userName,
      });

      // Create session
      const sessionToken = await sessionStore.createSession(userId);

      // Determine server base URL for the config snippet
      const serverBase = `${req.protocol}://${req.get('host')}`;

      // Return success HTML
      res.status(200).send(renderSuccessPage({
        userName,
        userEmail,
        sessionToken,
        serverBase,
      }));

    } catch (err) {
      console.error('Auth callback error:', err);
      res.status(500).send('Internal server error during authentication');
    }
  });

  return router;
}

// ── Success page HTML ───────────────────────────────────────────────────

/**
 * Renders a minimal HTML success page with the session token and config snippet.
 *
 * @param {object} opts
 * @param {string|null} opts.userName
 * @param {string|null} opts.userEmail
 * @param {string} opts.sessionToken
 * @param {string} opts.serverBase
 * @returns {string}
 */
function renderSuccessPage({ userName, userEmail, sessionToken, serverBase }) {
  const displayName = userName || userEmail || 'User';
  const configSnippet = JSON.stringify({
    mcpServers: {
      outlook: {
        url: `${serverBase}/mcp`,
        headers: {
          Authorization: `Bearer ${sessionToken}`,
        },
      },
    },
  }, null, 2);

  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Outlook MCP - Authentication Successful</title>
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
      background: #f5f5f5;
      color: #333;
      display: flex;
      justify-content: center;
      align-items: center;
      min-height: 100vh;
      padding: 20px;
    }
    .card {
      background: #fff;
      border-radius: 12px;
      box-shadow: 0 2px 12px rgba(0,0,0,0.1);
      max-width: 640px;
      width: 100%;
      padding: 40px;
    }
    .success-icon {
      font-size: 48px;
      margin-bottom: 16px;
    }
    h1 {
      font-size: 24px;
      margin-bottom: 8px;
      color: #1a1a1a;
    }
    .subtitle {
      color: #666;
      margin-bottom: 32px;
    }
    .section {
      margin-bottom: 24px;
    }
    .section-label {
      font-size: 12px;
      font-weight: 600;
      text-transform: uppercase;
      letter-spacing: 0.5px;
      color: #888;
      margin-bottom: 8px;
    }
    .token-box {
      background: #f8f8f8;
      border: 1px solid #e0e0e0;
      border-radius: 8px;
      padding: 12px 16px;
      font-family: 'SF Mono', Monaco, 'Cascadia Code', monospace;
      font-size: 13px;
      word-break: break-all;
      position: relative;
    }
    .config-box {
      background: #1e1e1e;
      color: #d4d4d4;
      border-radius: 8px;
      padding: 16px;
      font-family: 'SF Mono', Monaco, 'Cascadia Code', monospace;
      font-size: 12px;
      white-space: pre;
      overflow-x: auto;
      position: relative;
    }
    .copy-btn {
      position: absolute;
      top: 8px;
      right: 8px;
      background: #0078d4;
      color: #fff;
      border: none;
      border-radius: 6px;
      padding: 6px 12px;
      font-size: 12px;
      cursor: pointer;
      transition: background 0.2s;
    }
    .copy-btn:hover { background: #106ebe; }
    .copy-btn.copied { background: #107c10; }
    .note {
      font-size: 13px;
      color: #888;
      margin-top: 24px;
      line-height: 1.5;
    }
  </style>
</head>
<body>
  <div class="card">
    <div class="success-icon">&#10003;</div>
    <h1>Authentication Successful</h1>
    <p class="subtitle">Welcome, ${escapeHtml(displayName)}. Your Outlook MCP session is ready.</p>

    <div class="section">
      <div class="section-label">Session Token</div>
      <div class="token-box" id="token-box">
        ${escapeHtml(sessionToken)}
        <button class="copy-btn" onclick="copyText('${escapeHtml(sessionToken)}', this)">Copy</button>
      </div>
    </div>

    <div class="section">
      <div class="section-label">Claude Code Configuration</div>
      <div class="config-box" id="config-box">
${escapeHtml(configSnippet)}
        <button class="copy-btn" onclick="copyConfig(this)">Copy</button>
      </div>
    </div>

    <p class="note">
      Add the configuration above to your Claude Code MCP settings.
      This session token is valid for 30 days. You can close this page.
    </p>
  </div>
  <script>
    function copyText(text, btn) {
      navigator.clipboard.writeText(text).then(function() {
        btn.textContent = 'Copied!';
        btn.classList.add('copied');
        setTimeout(function() {
          btn.textContent = 'Copy';
          btn.classList.remove('copied');
        }, 2000);
      });
    }
    function copyConfig(btn) {
      var config = ${JSON.stringify(configSnippet)};
      copyText(config, btn);
    }
  </script>
</body>
</html>`;
}

/**
 * Escapes HTML entities to prevent XSS in rendered pages.
 * @param {string} str
 * @returns {string}
 */
function escapeHtml(str) {
  if (!str) return '';
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

module.exports = { createAuthRoutes, generatePKCE, _pendingAuth: pendingAuth };
