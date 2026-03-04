/**
 * Per-user token storage for multi-user hosted mode.
 *
 * In-memory Map keyed by userId, backed by an encrypted JSON file on disk.
 * The Map serves as a cache; all mutations are persisted automatically.
 *
 * Encryption: AES-256-GCM via node:crypto. When no encryptionKey is
 * provided (local dev / testing), tokens are stored as plain JSON.
 *
 * Expiry semantics:
 *   - 5-minute buffer before actual expiry
 *   - expiresAt (ms timestamp) calculated from expiresIn at storage time
 */

const fs = require('fs');
const fsp = require('fs/promises');
const path = require('path');
const crypto = require('crypto');

const TOKEN_EXPIRY_BUFFER_MS = 5 * 60 * 1000; // 5 minutes

class PerUserTokenStorage {
  /**
   * @param {object} [options]
   * @param {string} [options.filePath] — where to persist tokens on disk
   * @param {string} [options.encryptionKey] — AES-256 passphrase; required in hosted mode
   */
  constructor({ filePath, encryptionKey } = {}) {
    /** @type {Map<string, {accessToken: string, refreshToken: string, expiresAt: number, scopes: string, email: string|null, name: string|null}>} */
    this._tokens = new Map();

    /** @type {string|null} */
    this._filePath = filePath || null;

    /** @type {Buffer|null} 32-byte derived key */
    this._key = encryptionKey
      ? crypto.createHash('sha256').update(encryptionKey).digest()
      : null;
  }

  // ── Validation ──────────────────────────────────────────────────────

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

  // ── Internal helpers ────────────────────────────────────────────────

  /**
   * Returns the raw stored data for a user. Exposed for test access only.
   * @param {string} userId
   * @returns {object|undefined}
   */
  _getUserData(userId) {
    return this._tokens.get(userId);
  }

  /**
   * Internal expiry check against the 5-minute buffer.
   * @param {{expiresAt: number}} data
   * @returns {boolean}
   */
  _isExpired(data) {
    return Date.now() >= (data.expiresAt - TOKEN_EXPIRY_BUFFER_MS);
  }

  // ── Public read methods ─────────────────────────────────────────────

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
    if (!data) return null;
    if (this._isExpired(data)) return null;

    return data.accessToken;
  }

  /**
   * Returns the stored refresh token string for a user, or null.
   *
   * @param {string} userId
   * @returns {string|null}
   */
  getRefreshToken(userId) {
    this._validateUserId(userId);

    const data = this._tokens.get(userId);
    if (!data) return null;

    return data.refreshToken || null;
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
    if (!data || !data.expiresAt) return true;

    return this._isExpired(data);
  }

  /**
   * Returns { email, name } for the user, or null if unknown.
   *
   * @param {string} userId
   * @returns {{ email: string|null, name: string|null }|null}
   */
  getUserInfo(userId) {
    this._validateUserId(userId);

    const data = this._tokens.get(userId);
    if (!data) return null;

    return { email: data.email, name: data.name };
  }

  /**
   * Returns an array of all stored users with basic info.
   *
   * @returns {Array<{ userId: string, email: string|null, name: string|null, hasValidToken: boolean }>}
   */
  getAllUsers() {
    const users = [];
    for (const [userId, data] of this._tokens) {
      users.push({
        userId,
        email: data.email,
        name: data.name,
        hasValidToken: !this._isExpired(data),
      });
    }
    return users;
  }

  /**
   * Returns the number of users with stored tokens.
   *
   * @returns {number}
   */
  getActiveUserCount() {
    return this._tokens.size;
  }

  // ── Public mutation methods ─────────────────────────────────────────

  /**
   * Stores token data for a user. Calculates expiresAt from expiresIn.
   * Persists to disk automatically.
   *
   * @param {string} userId
   * @param {{ accessToken: string, refreshToken: string, expiresIn: number, scopes: string, email?: string|null, name?: string|null }} tokenData
   * @returns {Promise<void>}
   */
  async setTokensForUser(userId, tokenData) {
    this._validateUserId(userId);

    this._tokens.set(userId, {
      accessToken: tokenData.accessToken,
      refreshToken: tokenData.refreshToken,
      expiresAt: Date.now() + (tokenData.expiresIn * 1000),
      scopes: tokenData.scopes,
      email: tokenData.email ?? null,
      name: tokenData.name ?? null,
    });

    await this.saveToFile();
  }

  /**
   * Forces the user's token to be treated as expired on next access
   * by setting expiresAt to 0. Persists to disk automatically.
   *
   * @param {string} userId
   * @returns {Promise<void>}
   */
  async invalidateUser(userId) {
    this._validateUserId(userId);

    const data = this._tokens.get(userId);
    if (data) {
      data.expiresAt = 0;
      await this.saveToFile();
    }
  }

  /**
   * Removes all token data for a user. Persists to disk automatically.
   *
   * @param {string} userId
   * @returns {Promise<void>}
   */
  async removeUser(userId) {
    this._validateUserId(userId);

    const deleted = this._tokens.delete(userId);
    if (deleted) {
      await this.saveToFile();
    }
  }

  // ── File persistence ────────────────────────────────────────────────

  /**
   * Loads tokens from the persisted file into the in-memory Map.
   * If the file does not exist, the Map is left empty (not an error).
   *
   * @returns {Promise<void>}
   */
  async loadFromFile() {
    if (!this._filePath) return;

    let raw;
    try {
      raw = await fsp.readFile(this._filePath, 'utf8');
    } catch (err) {
      if (err.code === 'ENOENT') return; // file doesn't exist yet — fine
      throw err;
    }

    const parsed = JSON.parse(raw);
    let data;

    if (parsed.encrypted) {
      if (!this._key) {
        throw new Error('Token file is encrypted but no encryptionKey was provided');
      }
      data = this._decrypt(parsed);
    } else {
      data = parsed;
    }

    this._tokens.clear();
    for (const [userId, tokenData] of Object.entries(data)) {
      this._tokens.set(userId, tokenData);
    }
  }

  /**
   * Serializes the in-memory Map to JSON, optionally encrypts, and writes
   * to disk using atomic rename (write .tmp -> fsync -> rename).
   *
   * @returns {Promise<void>}
   */
  async saveToFile() {
    if (!this._filePath) return;

    // Ensure directory exists
    const dir = path.dirname(this._filePath);
    await fsp.mkdir(dir, { recursive: true });

    // Serialize Map to plain object
    const plain = {};
    for (const [userId, tokenData] of this._tokens) {
      plain[userId] = tokenData;
    }

    let content;
    if (this._key) {
      content = JSON.stringify(this._encrypt(plain));
    } else {
      content = JSON.stringify(plain, null, 2);
    }

    // Atomic write: write tmp -> fsync -> rename
    const tmpPath = this._filePath + '.tmp';
    const fd = await fsp.open(tmpPath, 'w');
    try {
      await fd.writeFile(content, 'utf8');
      await fd.sync();
    } finally {
      await fd.close();
    }
    await fsp.rename(tmpPath, this._filePath);
  }

  // ── Encryption helpers ──────────────────────────────────────────────

  /**
   * Encrypts a plain object using AES-256-GCM.
   *
   * @param {object} data
   * @returns {{ encrypted: true, iv: string, authTag: string, data: string }}
   */
  _encrypt(data) {
    const iv = crypto.randomBytes(12);
    const cipher = crypto.createCipheriv('aes-256-gcm', this._key, iv);

    const jsonStr = JSON.stringify(data);
    let encrypted = cipher.update(jsonStr, 'utf8', 'hex');
    encrypted += cipher.final('hex');

    return {
      encrypted: true,
      iv: iv.toString('hex'),
      authTag: cipher.getAuthTag().toString('hex'),
      data: encrypted,
    };
  }

  /**
   * Decrypts an encrypted envelope back to a plain object.
   *
   * @param {{ iv: string, authTag: string, data: string }} envelope
   * @returns {object}
   */
  _decrypt(envelope) {
    const iv = Buffer.from(envelope.iv, 'hex');
    const authTag = Buffer.from(envelope.authTag, 'hex');
    const decipher = crypto.createDecipheriv('aes-256-gcm', this._key, iv);
    decipher.setAuthTag(authTag);

    let decrypted = decipher.update(envelope.data, 'hex', 'utf8');
    decrypted += decipher.final('utf8');

    return JSON.parse(decrypted);
  }
}

module.exports = PerUserTokenStorage;
