/**
 * On-Behalf-Of (OBO) Token Exchange
 *
 * Exchanges a user's Entra bearer token for a Microsoft Graph API token
 * using the OAuth 2.0 On-Behalf-Of flow.
 *
 * @see https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-on-behalf-of-flow
 */
const https = require('https');
const querystring = require('querystring');

/**
 * Exchanges a user's Entra access token for a Graph API token via the OBO flow.
 *
 * @param {string} userAccessToken - The JWT bearer token from the user's Entra authentication
 * @param {object} config - Configuration object with { clientId, clientSecret, tenantId, scopes }
 * @returns {Promise<{access_token: string, refresh_token: string, expires_in: number, scope: string, token_type: string}>}
 */
async function exchangeOBO(userAccessToken, config) {
  // Validate parameters
  if (!userAccessToken) {
    throw new Error('OBO exchange failed: userAccessToken is required');
  }
  if (!config) {
    throw new Error('OBO exchange failed: config is required');
  }
  if (!config.clientId) {
    throw new Error('OBO exchange failed: config.clientId is required');
  }
  if (!config.clientSecret) {
    throw new Error('OBO exchange failed: config.clientSecret is required');
  }
  if (!config.tenantId) {
    throw new Error('OBO exchange failed: config.tenantId is required');
  }
  if (!config.scopes) {
    throw new Error('OBO exchange failed: config.scopes is required');
  }

  const tokenEndpoint = `https://login.microsoftonline.com/${config.tenantId}/oauth2/v2.0/token`;

  const postData = querystring.stringify({
    grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
    client_id: config.clientId,
    client_secret: config.clientSecret,
    assertion: userAccessToken,
    scope: config.scopes.join(' '),
    requested_token_use: 'on_behalf_of',
  });

  return new Promise((resolve, reject) => {
    const req = https.request(tokenEndpoint, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        'Content-Length': Buffer.byteLength(postData),
      },
      timeout: 30000,
    }, (res) => {
      let data = '';
      res.on('data', (chunk) => data += chunk);
      res.on('end', () => {
        let responseBody;
        try {
          responseBody = JSON.parse(data);
        } catch (e) {
          reject(new Error('OBO exchange failed: invalid response from token endpoint'));
          return;
        }

        if (res.statusCode >= 200 && res.statusCode < 300) {
          resolve(responseBody);
        } else {
          const errorCode = responseBody.error || '';
          const errorDesc = responseBody.error_description || '';

          if (errorCode === 'invalid_grant') {
            reject(new Error('OBO exchange failed: invalid or expired user token'));
          } else if (errorCode === 'interaction_required' || errorCode === 'consent_required') {
            reject(new Error('OBO exchange failed: admin consent required for Graph permissions'));
          } else {
            const detail = errorDesc ? `${errorCode} - ${errorDesc}` : errorCode;
            reject(new Error(`OBO exchange failed: ${detail}`));
          }
        }
      });
    });

    req.on('timeout', () => {
      req.destroy();
      reject(new Error('OBO exchange failed: request timed out'));
    });

    req.on('error', (error) => {
      reject(new Error(`OBO exchange failed: network error - ${error.message}`));
    });

    req.write(postData);
    req.end();
  });
}

module.exports = { exchangeOBO };
