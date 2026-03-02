const fs = require('fs').promises;
const path = require('path');
const https = require('https');
const querystring = require('querystring');

const REFRESH_TIMEOUT_MS = 30000;
const MAX_REFRESH_RETRIES = 2;
const RETRY_DELAY_MS = 1000;

class TokenStorage {
  constructor(config) {
    this.config = {
      tokenStorePath: path.join(process.env.HOME || process.env.USERPROFILE, '.outlook-mcp-tokens.json'),
      clientId: process.env.MS_CLIENT_ID,
      clientSecret: process.env.MS_CLIENT_SECRET,
      redirectUri: process.env.MS_REDIRECT_URI || 'http://localhost:3333/auth/callback',
      scopes: (process.env.MS_SCOPES || 'offline_access User.Read Mail.Read').split(' '),
      tokenEndpoint: process.env.MS_TOKEN_ENDPOINT || `https://login.microsoftonline.com/${process.env.MS_TENANT_ID || 'common'}/oauth2/v2.0/token`,
      refreshTokenBuffer: 5 * 60 * 1000,
      ...config
    };
    this.tokens = null;
    this._loadPromise = null;
    this._refreshPromise = null;

    if (!this.config.clientId || !this.config.clientSecret) {
      console.warn("TokenStorage: MS_CLIENT_ID or MS_CLIENT_SECRET is not configured. Token operations might fail.");
    }
  }

  _isAuthError(error) {
    const msg = (error.message || '').toLowerCase();
    return msg.includes('invalid_grant') ||
           msg.includes('invalid_client') ||
           msg.includes('unauthorized_client') ||
           msg.includes('interaction_required') ||
           msg.includes('consent_required');
  }

  async _loadTokensFromFile() {
    try {
      const tokenData = await fs.readFile(this.config.tokenStorePath, 'utf8');
      this.tokens = JSON.parse(tokenData);
      console.error('Tokens loaded from file.');
      return this.tokens;
    } catch (error) {
      if (error.code === 'ENOENT') {
        console.error('Token file not found. No tokens loaded.');
      } else {
        console.error('Error loading token cache:', error);
      }
      this.tokens = null;
      return null;
    }
  }

  async _saveTokensToFile() {
    if (!this.tokens) {
      console.warn('No tokens to save.');
      return false;
    }
    try {
      await fs.writeFile(this.config.tokenStorePath, JSON.stringify(this.tokens, null, 2));
      console.error('Tokens saved successfully.');
      // return true; // No longer returning boolean, will throw on error.
    } catch (error) {
      console.error('Error saving token cache:', error);
      throw error; // Propagate the error
    }
  }

  async getTokens() {
    if (this.tokens) {
      return this.tokens;
    }
    if (!this._loadPromise) {
        this._loadPromise = this._loadTokensFromFile().finally(() => {
            this._loadPromise = null; // Reset promise once completed
        });
    }
    return this._loadPromise;
  }

  getExpiryTime() {
    return this.tokens && this.tokens.expires_at ? this.tokens.expires_at : 0;
  }

  isTokenExpired() {
    if (!this.tokens || !this.tokens.expires_at) {
      return true; // No token or no expiry means it's effectively expired or invalid
    }
    // Check if current time is past expiry time, considering a buffer
    return Date.now() >= (this.tokens.expires_at - this.config.refreshTokenBuffer);
  }

  invalidateAccessToken() {
    if (this.tokens) {
      this.tokens.expires_at = 0;
    }
  }

  async getValidAccessToken() {
    await this.getTokens();

    if (!this.tokens || !this.tokens.access_token) {
      console.error('No access token available.');
      return null;
    }

    if (this.isTokenExpired()) {
      console.error('Access token expired or nearing expiration. Attempting refresh.');
      if (this.tokens.refresh_token) {
        try {
          return await this.refreshAccessToken();
        } catch (refreshError) {
          console.error('Failed to refresh access token:', refreshError);
          if (this._isAuthError(refreshError)) {
            console.error('Definitive auth error — clearing in-memory tokens (token file retained).');
            this.tokens = null;
          }
          return null;
        }
      } else {
        console.warn('No refresh token available. Cannot refresh access token.');
        return null;
      }
    }
    return this.tokens.access_token;
  }

  async refreshAccessToken() {
    if (!this.tokens || !this.tokens.refresh_token) {
      throw new Error('No refresh token available to refresh the access token.');
    }

    if (this._refreshPromise) {
      console.error("Refresh already in progress, returning existing promise.");
      return this._refreshPromise;
    }

    this._refreshPromise = this._refreshWithRetry().finally(() => {
      this._refreshPromise = null;
    });

    return this._refreshPromise;
  }

  async _refreshWithRetry() {
    let lastError;
    for (let attempt = 0; attempt < MAX_REFRESH_RETRIES; attempt++) {
      try {
        return await this._doRefresh();
      } catch (error) {
        lastError = error;
        if (this._isAuthError(error)) {
          throw error;
        }
        console.error(`Refresh attempt ${attempt + 1} failed (transient):`, error.message);
        if (attempt < MAX_REFRESH_RETRIES - 1) {
          await new Promise(r => setTimeout(r, RETRY_DELAY_MS * (attempt + 1)));
        }
      }
    }
    throw lastError;
  }

  _doRefresh() {
    const postData = querystring.stringify({
      client_id: this.config.clientId,
      client_secret: this.config.clientSecret,
      grant_type: 'refresh_token',
      refresh_token: this.tokens.refresh_token,
      scope: this.config.scopes.join(' ')
    });

    return new Promise((resolve, reject) => {
      const req = https.request(this.config.tokenEndpoint, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded',
          'Content-Length': Buffer.byteLength(postData)
        },
        timeout: REFRESH_TIMEOUT_MS
      }, (res) => {
        let data = '';
        res.on('data', (chunk) => data += chunk);
        res.on('end', async () => {
          try {
            const responseBody = JSON.parse(data);
            if (res.statusCode >= 200 && res.statusCode < 300) {
              this.tokens.access_token = responseBody.access_token;
              if (responseBody.refresh_token) {
                this.tokens.refresh_token = responseBody.refresh_token;
              }
              this.tokens.expires_in = responseBody.expires_in;
              this.tokens.expires_at = Date.now() + (responseBody.expires_in * 1000);
              try {
                await this._saveTokensToFile();
              } catch (saveError) {
                console.error('Failed to persist refreshed tokens (using in-memory):', saveError);
              }
              console.error('Access token refreshed successfully.');
              resolve(this.tokens.access_token);
            } else {
              const errorCode = responseBody.error || '';
              const errorDesc = responseBody.error_description || `status ${res.statusCode}`;
              const errMsg = errorCode ? `${errorCode}: ${errorDesc}` : errorDesc;
              console.error('Error refreshing token:', responseBody);
              reject(new Error(errMsg));
            }
          } catch (e) {
            console.error('Error processing refresh response:', e);
            reject(e);
          }
        });
      });

      req.on('timeout', () => {
        req.destroy();
        reject(new Error('Token refresh request timed out'));
      });
      req.on('error', (error) => {
        console.error('HTTP error during token refresh:', error);
        reject(error);
      });
      req.write(postData);
      req.end();
    });
  }


  async exchangeCodeForTokens(authCode) {
    if (!this.config.clientId || !this.config.clientSecret) {
        throw new Error("Client ID or Client Secret is not configured. Cannot exchange code for tokens.");
    }
    console.error('Exchanging authorization code for tokens...');
    const postData = querystring.stringify({
      client_id: this.config.clientId,
      client_secret: this.config.clientSecret,
      grant_type: 'authorization_code',
      code: authCode,
      redirect_uri: this.config.redirectUri,
      scope: this.config.scopes.join(' ')
    });

    const requestOptions = {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        'Content-Length': Buffer.byteLength(postData)
      }
    };

    return new Promise((resolve, reject) => {
      const req = https.request(this.config.tokenEndpoint, requestOptions, (res) => {
        let data = '';
        res.on('data', (chunk) => data += chunk);
        res.on('end', async () => {
          try {
            const responseBody = JSON.parse(data);
            if (res.statusCode >= 200 && res.statusCode < 300) {
              this.tokens = {
                access_token: responseBody.access_token,
                refresh_token: responseBody.refresh_token,
                expires_in: responseBody.expires_in,
                expires_at: Date.now() + (responseBody.expires_in * 1000),
                scope: responseBody.scope,
                token_type: responseBody.token_type
              };
              try {
                await this._saveTokensToFile();
                console.error('Tokens exchanged and saved successfully.');
                resolve(this.tokens);
              } catch (saveError) {
                console.error('Failed to save exchanged tokens:', saveError);
                // Similar to refresh, tokens are in memory but not persisted.
                // Rejecting to indicate the operation wasn't fully successful.
                reject(new Error(`Tokens exchanged but failed to save: ${saveError.message}`));
              }
            } else {
              console.error('Error exchanging code for tokens:', responseBody);
              reject(new Error(responseBody.error_description || `Token exchange failed with status ${res.statusCode}`));
            }
          } catch (e) { // Catch any error during parsing or saving
            console.error('Error processing token exchange response or saving tokens:', e, "Raw data:", data);
            reject(new Error(`Error processing token response: ${e.message}. Response data: ${data}`));
          }
        });
      });
      req.on('error', (error) => {
        console.error('HTTP error during code exchange:', error);
        reject(error);
      });
      req.write(postData);
      req.end();
    });
  }

  // Utility to clear tokens, e.g., for logout or forcing re-auth
  async clearTokens() {
    this.tokens = null;
    try {
      await fs.unlink(this.config.tokenStorePath);
      console.error('Token file deleted successfully.');
    } catch (error) {
      if (error.code === 'ENOENT') {
        console.error('Token file not found, nothing to delete.');
      } else {
        console.error('Error deleting token file:', error);
      }
    }
  }
}

module.exports = TokenStorage;
// Adding a newline at the end of the file as requested by Gemini Code Assist
