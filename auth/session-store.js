/**
 * Session token store for multi-user hosted mode.
 *
 * Maps opaque session tokens (sent in the Authorization header) to user IDs.
 * Sessions do not expire by default and are persisted to disk using the
 * same encrypted-JSON pattern as the rest of the auth layer.
 *
 * Storage format (encrypted):
 *   { encrypted: true, iv: <hex>, authTag: <hex>, data: <hex ciphertext> }
 *
 * Storage format (unencrypted):
 *   Plain JSON object keyed by session token.
 */

const crypto = require('node:crypto');
const fsPromises = require('node:fs/promises');
const path = require('node:path');
const os = require('node:os');

const DEFAULT_FILE_PATH = path.join(
  process.env.HOME || process.env.USERPROFILE || os.homedir() || '/tmp',
  '.outlook-mcp-sessions.json'
);

const ALGORITHM = 'aes-256-gcm';
const IV_BYTES = 12;

class SessionStore {
  /**
   * @param {object} [opts]
   * @param {string} [opts.filePath]       Where to persist sessions on disk.
   * @param {string} [opts.encryptionKey]  Optional AES-256 passphrase (hashed to 32 bytes).
   */
  constructor({ filePath, encryptionKey } = {}) {
    this.filePath = filePath || DEFAULT_FILE_PATH;
    this._encryptionKey = encryptionKey
      ? crypto.createHash('sha256').update(encryptionKey).digest()
      : null;

    /** @type {Map<string, {userId: string, createdAt: string, expiresAt: string}>} */
    this._sessions = new Map();

    // Serialize writes to avoid temp-file races under concurrent save operations.
    this._saveChain = Promise.resolve();
  }

  // ── Public API ──────────────────────────────────────────────────────

  /**
   * Creates a new session for the given user and persists immediately.
   *
   * @param {string} userId
   * @param {object} [opts]
   * @param {number} [opts.expiresInDays=0] 0 means no expiry.
   * @returns {Promise<string>} The generated session token.
   */
  async createSession(userId, { expiresInDays = 0 } = {}) {
    const token = crypto.randomUUID();
    const now = new Date();
    const expiresAt = expiresInDays > 0
      ? new Date(now.getTime() + expiresInDays * 24 * 60 * 60 * 1000)
      : null;

    this._sessions.set(token, {
      userId,
      createdAt: now.toISOString(),
      expiresAt: expiresAt ? expiresAt.toISOString() : null,
    });

    await this.saveToFile();
    return token;
  }

  /**
   * Validates a session token.
   *
   * @param {string} token
   * @returns {{ userId: string, createdAt: string, expiresAt: string|null } | null}
   */
  validateSession(token) {
    if (token == null) return null;

    const session = this._sessions.get(token);
    if (!session) return null;

    // Lazy expiry cleanup (skip if no expiry set)
    if (session.expiresAt && new Date(session.expiresAt) <= new Date()) {
      this._sessions.delete(token);
      // Fire-and-forget save; callers don't need to wait for cleanup persistence
      this.saveToFile().catch(() => {});
      return null;
    }

    return { ...session };
  }

  /**
   * Revokes a single session token.
   *
   * @param {string} token
   * @returns {Promise<boolean>} true if the token existed and was removed.
   */
  async revokeSession(token) {
    const deleted = this._sessions.delete(token);
    if (deleted) await this.saveToFile();
    return deleted;
  }

  /**
   * Revokes every session belonging to a given user.
   *
   * @param {string} userId
   * @returns {Promise<number>} Number of sessions removed.
   */
  async revokeAllForUser(userId) {
    let count = 0;
    for (const [token, session] of this._sessions) {
      if (session.userId === userId) {
        this._sessions.delete(token);
        count++;
      }
    }
    if (count > 0) await this.saveToFile();
    return count;
  }

  /**
   * Returns all active (non-expired) sessions with truncated tokens for admin display.
   *
   * @returns {Array<{token: string, userId: string, createdAt: string, expiresAt: string}>}
   */
  getActiveSessions() {
    const now = new Date();
    const result = [];
    for (const [token, session] of this._sessions) {
      if (!session.expiresAt || new Date(session.expiresAt) > now) {
        result.push({
          token: token.slice(0, 8) + '...',
          userId: session.userId,
          createdAt: session.createdAt,
          expiresAt: session.expiresAt,
        });
      }
    }
    return result;
  }

  /**
   * Returns the number of active (non-expired) sessions for a user.
   *
   * @param {string} userId
   * @returns {number}
   */
  getSessionCountForUser(userId) {
    const now = new Date();
    let count = 0;
    for (const session of this._sessions.values()) {
      if (session.userId === userId && (!session.expiresAt || new Date(session.expiresAt) > now)) {
        count++;
      }
    }
    return count;
  }

  // ── Persistence ─────────────────────────────────────────────────────

  /**
   * Loads sessions from the persisted file on disk.
   * If the file does not exist the store is left empty (no error).
   */
  async loadFromFile() {
    let raw;
    try {
      raw = await fsPromises.readFile(this.filePath, 'utf8');
    } catch (err) {
      if (err.code === 'ENOENT') {
        // File doesn't exist yet — start with an empty store.
        this._sessions = new Map();
        return;
      }
      throw err;
    }

    const parsed = JSON.parse(raw);
    let plain;

    if (parsed.encrypted) {
      if (!this._encryptionKey) {
        throw new Error('Session file is encrypted but no encryptionKey was provided');
      }
      plain = this._decrypt(parsed);
    } else {
      plain = parsed;
    }

    this._sessions = new Map(Object.entries(plain));
  }

  /**
   * Persists all sessions to disk.  Uses atomic write (tmp + fsync + rename)
   * to avoid corruption on crash.
   */
  async saveToFile() {
    return this._enqueueSave(async () => {
      const plain = Object.fromEntries(this._sessions);
      let payload;

      if (this._encryptionKey) {
        payload = JSON.stringify(this._encrypt(JSON.stringify(plain)));
      } else {
        payload = JSON.stringify(plain, null, 2);
      }

      const tmpPath = this._buildTmpPath();

      // Ensure the directory exists
      const dir = path.dirname(this.filePath);
      await fsPromises.mkdir(dir, { recursive: true });

      try {
        // Write to temp file with user-only permissions
        const fd = await fsPromises.open(tmpPath, 'w', 0o600);
        try {
          await fd.writeFile(payload, 'utf8');
          await fd.sync();
        } finally {
          await fd.close();
        }

        // Atomic rename
        await fsPromises.rename(tmpPath, this.filePath);
      } catch (err) {
        await fsPromises.unlink(tmpPath).catch(() => {});
        throw err;
      }
    });
  }

  _enqueueSave(saveFn) {
    const next = this._saveChain.then(saveFn, saveFn);
    this._saveChain = next.catch(() => {});
    return next;
  }

  _buildTmpPath() {
    return `${this.filePath}.${process.pid}.${Date.now()}.${crypto.randomUUID()}.tmp`;
  }

  // ── Encryption helpers ──────────────────────────────────────────────

  /**
   * Encrypts a plaintext string with AES-256-GCM.
   * @param {string} plaintext
   * @returns {{encrypted: true, iv: string, authTag: string, data: string}}
   */
  _encrypt(plaintext) {
    const iv = crypto.randomBytes(IV_BYTES);
    const cipher = crypto.createCipheriv(ALGORITHM, this._encryptionKey, iv);
    const encrypted = Buffer.concat([cipher.update(plaintext, 'utf8'), cipher.final()]);
    const authTag = cipher.getAuthTag();
    return {
      encrypted: true,
      iv: iv.toString('hex'),
      authTag: authTag.toString('hex'),
      data: encrypted.toString('hex'),
    };
  }

  /**
   * Decrypts an encrypted payload back to a parsed JSON object.
   * @param {{iv: string, authTag: string, data: string}} envelope
   * @returns {object}
   */
  _decrypt(envelope) {
    const iv = Buffer.from(envelope.iv, 'hex');
    const authTag = Buffer.from(envelope.authTag, 'hex');
    const ciphertext = Buffer.from(envelope.data, 'hex');
    const decipher = crypto.createDecipheriv(ALGORITHM, this._encryptionKey, iv);
    decipher.setAuthTag(authTag);
    const decrypted = Buffer.concat([decipher.update(ciphertext), decipher.final()]);
    return JSON.parse(decrypted.toString('utf8'));
  }
}

module.exports = SessionStore;
