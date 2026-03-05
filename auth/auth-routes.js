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
const STATE_TOKEN_VERSION = 1;

/**
 * In-memory store for pending authorization requests.
 * Keyed by the random `state` parameter, each entry holds the PKCE
 * code verifier and an expiry timestamp.
 *
 * @type {Map<string, { codeVerifier: string, expiresAt: number }>}
 */
const pendingAuth = new Map();
const consumedAuthStates = new Map();

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

function cleanupConsumedAuthStates() {
  const now = Date.now();
  for (const [stateHash, expiresAt] of consumedAuthStates) {
    if (expiresAt <= now) {
      consumedAuthStates.delete(stateHash);
    }
  }
}

function getStateSigningSecret(authConfig) {
  return (
    process.env.AUTH_STATE_SECRET ||
    process.env.TOKEN_ENCRYPTION_KEY ||
    authConfig.clientSecret ||
    null
  );
}

function createSignedState({ codeVerifier, expiresAt, secret }) {
  const payload = JSON.stringify({
    v: STATE_TOKEN_VERSION,
    cv: codeVerifier,
    exp: expiresAt,
    n: crypto.randomBytes(12).toString('base64url'),
  });
  const payloadB64 = Buffer.from(payload, 'utf8').toString('base64url');
  const sig = crypto.createHmac('sha256', secret).update(payloadB64).digest('base64url');
  return `${payloadB64}.${sig}`;
}

function parseSignedState(state, secret) {
  if (!state || typeof state !== 'string') {
    return null;
  }

  const parts = state.split('.');
  if (parts.length !== 2) {
    return null;
  }

  const [payloadB64, sig] = parts;
  if (!payloadB64 || !sig) {
    return null;
  }

  const expectedSig = crypto.createHmac('sha256', secret).update(payloadB64).digest('base64url');
  const actualSigBuf = Buffer.from(sig);
  const expectedSigBuf = Buffer.from(expectedSig);
  if (actualSigBuf.length !== expectedSigBuf.length) {
    return null;
  }
  if (!crypto.timingSafeEqual(actualSigBuf, expectedSigBuf)) {
    return null;
  }

  try {
    const payloadJson = Buffer.from(payloadB64, 'base64url').toString('utf8');
    const payload = JSON.parse(payloadJson);

    if (payload?.v !== STATE_TOKEN_VERSION) {
      return null;
    }
    if (typeof payload.cv !== 'string' || !payload.cv) {
      return null;
    }
    if (!Number.isFinite(payload.exp)) {
      return null;
    }
    if (payload.exp <= Date.now()) {
      return null;
    }

    return {
      codeVerifier: payload.cv,
      expiresAt: payload.exp,
    };
  } catch {
    return null;
  }
}

function markStateConsumed(state, expiresAt) {
  const stateHash = crypto.createHash('sha256').update(state).digest('hex');
  consumedAuthStates.set(stateHash, expiresAt);
}

function isStateConsumed(state) {
  const stateHash = crypto.createHash('sha256').update(state).digest('hex');
  const expiresAt = consumedAuthStates.get(stateHash);
  if (!expiresAt) {
    return false;
  }
  if (expiresAt <= Date.now()) {
    consumedAuthStates.delete(stateHash);
    return false;
  }
  return true;
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
  const stateSigningSecret = getStateSigningSecret(authConfig);

  // ── GET /login ─────────────────────────────────────────────────────

  router.get('/login', (req, res) => {
    // Lazy cleanup of expired pending auth entries
    cleanupExpiredPendingAuth();
    cleanupConsumedAuthStates();

    // Generate PKCE pair
    const { verifier, challenge } = generatePKCE();
    const expiresAt = Date.now() + PENDING_AUTH_TTL_MS;
    if (!stateSigningSecret) {
      return res.status(500).send('Auth state signing secret is not configured');
    }
    const state = createSignedState({
      codeVerifier: verifier,
      expiresAt,
      secret: stateSigningSecret,
    });

    // Store pending auth entry with TTL
    pendingAuth.set(state, {
      codeVerifier: verifier,
      expiresAt,
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
    if (typeof code !== 'string' || typeof state !== 'string') {
      return res.status(400).send('Invalid callback parameters');
    }
    if (!stateSigningSecret) {
      return res.status(500).send('Auth state signing secret is not configured');
    }
    cleanupConsumedAuthStates();
    if (isStateConsumed(state)) {
      return res.status(400).send('Invalid or expired state parameter');
    }

    // Validate state against pending auth map first (same-process fast path).
    const pending = pendingAuth.get(state);
    let codeVerifier;
    let stateExpiresAt;
    if (pending) {
      if (pending.expiresAt <= Date.now()) {
        pendingAuth.delete(state);
        return res.status(400).send('Invalid or expired state parameter');
      }
      ({ codeVerifier } = pending);
      stateExpiresAt = pending.expiresAt;
      pendingAuth.delete(state);
    } else {
      // Restart / multi-instance fallback: verify signed state payload.
      const parsed = parseSignedState(state, stateSigningSecret);
      if (!parsed) {
        return res.status(400).send('Invalid or expired state parameter');
      }
      codeVerifier = parsed.codeVerifier;
      stateExpiresAt = parsed.expiresAt;
    }
    markStateConsumed(state, stateExpiresAt);

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

      // Determine trusted server base URL for the config snippet
      const serverBase = getTrustedServerBaseUrl(authConfig, config.PORT);

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
  const claudeCodeConfig = JSON.stringify({
    mcpServers: {
      outlook: {
        url: `${serverBase}/mcp`,
        headers: {
          Authorization: `Bearer ${sessionToken}`,
        },
      },
    },
  }, null, 2);

  const claudeDesktopConfig = JSON.stringify({
    mcpServers: {
      outlook: {
        command: "npx",
        args: [
          "mcp-remote",
          `${serverBase}/mcp`,
          "--header",
          `Authorization:\${AUTH_HEADER}`,
        ],
        env: {
          AUTH_HEADER: `Bearer ${sessionToken}`,
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
      <div class="section-label">Claude Desktop Configuration</div>
      <div class="config-box" id="desktop-config-box">
${escapeHtml(claudeDesktopConfig)}
        <button class="copy-btn" onclick="copyConfig('desktop', this)">Copy</button>
      </div>
    </div>

    <div class="section">
      <div class="section-label">Claude Code Configuration</div>
      <div class="config-box" id="code-config-box">
${escapeHtml(claudeCodeConfig)}
        <button class="copy-btn" onclick="copyConfig('code', this)">Copy</button>
      </div>
    </div>

    <p class="note">
      Add one of the configurations above to your Claude settings.
      Claude Desktop requires <code>mcp-remote</code> (installed automatically via npx).
      You can close this page.
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
    function copyConfig(type, btn) {
      var config = type === 'desktop' ? ${JSON.stringify(claudeDesktopConfig)} : ${JSON.stringify(claudeCodeConfig)};
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

module.exports = {
  createAuthRoutes,
  generatePKCE,
  _pendingAuth: pendingAuth,
  _consumedAuthStates: consumedAuthStates,
};

function getTrustedServerBaseUrl(authConfig, fallbackPort = 3000) {
  const candidates = [
    process.env.PUBLIC_BASE_URL,
    authConfig.hostedRedirectUri,
    authConfig.redirectUri,
  ];

  for (const candidate of candidates) {
    if (!candidate) continue;
    try {
      return new URL(candidate).origin;
    } catch {
      // Ignore malformed values and continue.
    }
  }

  return `http://localhost:${fallbackPort || 3000}`;
}
