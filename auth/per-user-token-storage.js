/**
 * Per-user token storage for multi-user hosted mode.
 *
 * In-memory Map keyed by userId. No file persistence — hosted server
 * tokens are ephemeral and refresh via OBO (On-Behalf-Of) flow.
 *
 * Mirrors the expiry semantics of the single-user TokenStorage:
 *   - 5-minute buffer before actual expiry
 *   - expires_at calculated from expires_in at storage time
 */

const TOKEN_EXPIRY_BUFFER_MS = 5 * 60 * 1000; // 5 minutes

class PerUserTokenStorage {
  constructor() {
    /** @type {Map<string, {access_token: string, refresh_token: string, expires_at: number, scope: string}>} */
    this._tokens = new Map();
  }

  /**
   * Validates that a userId is present and usable.
   * @param {*} userId
   * @throws {Error} if userId is null or undefined
   */
  _validateUserId(userId) {
    if (userId == null) {
      throw new Error('userId must not be null or undefined');
    }
  }

  /**
   * Returns the raw stored data for a user. Exposed for test access only.
   * @param {string} userId
   * @returns {object|undefined}
   */
  _getUserData(userId) {
    return this._tokens.get(userId);
  }

  /**
   * Returns a valid access token string for the given user, or null if
   * the user has no token or the token is expired / within the 5-minute buffer.
   *
   * @param {string} userId
   * @returns {string|null}
   */
  getTokenForUser(userId) {
    this._validateUserId(userId);

    const data = this._tokens.get(userId);
    if (!data) {
      return null;
    }

    if (this._isExpired(data)) {
      return null;
    }

    return data.access_token;
  }

  /**
   * Stores token data for a user. Automatically calculates expires_at
   * from the provided expires_in (seconds).
   *
   * @param {string} userId
   * @param {{access_token: string, refresh_token: string, expires_in: number, scope: string}} tokenData
   */
  setTokenForUser(userId, tokenData) {
    this._validateUserId(userId);

    this._tokens.set(userId, {
      access_token: tokenData.access_token,
      refresh_token: tokenData.refresh_token,
      expires_at: Date.now() + (tokenData.expires_in * 1000),
      scope: tokenData.scope,
    });
  }

  /**
   * Returns true if the user's token is expired or within the 5-minute
   * expiry buffer. Also returns true if the user has no stored token.
   *
   * @param {string} userId
   * @returns {boolean}
   */
  isTokenExpired(userId) {
    this._validateUserId(userId);

    const data = this._tokens.get(userId);
    if (!data || !data.expires_at) {
      return true;
    }

    return this._isExpired(data);
  }

  /**
   * Forces the user's token to be treated as expired on next access
   * by setting expires_at to 0, matching the single-user invalidation pattern.
   *
   * @param {string} userId
   */
  invalidateUser(userId) {
    this._validateUserId(userId);

    const data = this._tokens.get(userId);
    if (data) {
      data.expires_at = 0;
    }
  }

  /**
   * Removes all token data for a user.
   *
   * @param {string} userId
   */
  removeUser(userId) {
    this._validateUserId(userId);
    this._tokens.delete(userId);
  }

  /**
   * Returns the number of users with stored tokens.
   *
   * @returns {number}
   */
  getActiveUserCount() {
    return this._tokens.size;
  }

  /**
   * Internal expiry check against the 5-minute buffer.
   * Matches TokenStorage.isTokenExpired() semantics:
   *   Date.now() >= (expires_at - buffer)
   *
   * @param {{expires_at: number}} data
   * @returns {boolean}
   */
  _isExpired(data) {
    return Date.now() >= (data.expires_at - TOKEN_EXPIRY_BUFFER_MS);
  }
}

module.exports = PerUserTokenStorage;
