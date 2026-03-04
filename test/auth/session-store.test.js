const path = require('path');
const fs = require('fs').promises;
const os = require('os');
const SessionStore = require('../../auth/session-store');

// UUID v4 regex (crypto.randomUUID output)
const UUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/;

/**
 * Helper: create a temp file path that won't collide across parallel tests.
 */
function tmpSessionPath(suffix = '') {
  return path.join(os.tmpdir(), `session-store-test-${Date.now()}-${Math.random().toString(36).slice(2)}${suffix}.json`);
}

/**
 * Helper: silently remove a file if it exists.
 */
async function rm(filePath) {
  try { await fs.unlink(filePath); } catch { /* ignore */ }
}

// Collect every temp path created so afterAll can clean up stragglers.
const tempFiles = [];

afterAll(async () => {
  for (const f of tempFiles) {
    await rm(f);
    await rm(f + '.tmp');
  }
});

// ── Core functionality ────────────────────────────────────────────────

describe('SessionStore', () => {
  let filePath;
  let store;

  beforeEach(() => {
    filePath = tmpSessionPath();
    tempFiles.push(filePath);
    store = new SessionStore({ filePath });
  });

  afterEach(async () => {
    await rm(filePath);
    await rm(filePath + '.tmp');
  });

  // ── createSession ───────────────────────────────────────────────────

  describe('createSession', () => {
    test('returns a valid UUID token', async () => {
      const token = await store.createSession('user-1');
      expect(token).toMatch(UUID_RE);
    });

    test('creates distinct tokens for successive calls', async () => {
      const t1 = await store.createSession('user-1');
      const t2 = await store.createSession('user-1');
      expect(t1).not.toBe(t2);
    });

    test('persists session to file automatically', async () => {
      await store.createSession('user-1');
      const raw = await fs.readFile(filePath, 'utf8');
      const data = JSON.parse(raw);
      expect(Object.keys(data)).toHaveLength(1);
    });
  });

  // ── validateSession ─────────────────────────────────────────────────

  describe('validateSession', () => {
    test('returns correct userId for a valid token', async () => {
      const token = await store.createSession('user-42');
      const result = store.validateSession(token);
      expect(result).not.toBeNull();
      expect(result.userId).toBe('user-42');
      expect(result.createdAt).toBeDefined();
      expect(result.expiresAt).toBeDefined();
    });

    test('returns null for an unknown token', () => {
      const result = store.validateSession('not-a-real-token');
      expect(result).toBeNull();
    });

    test('returns null for an expired token', async () => {
      const token = await store.createSession('user-1', { expiresInDays: 0 });

      // Manually backdate the expiry to guarantee it's in the past.
      const session = store._sessions.get(token);
      session.expiresAt = new Date(Date.now() - 1000).toISOString();

      const result = store.validateSession(token);
      expect(result).toBeNull();
    });

    test('removes expired token from internal map on validate', async () => {
      const token = await store.createSession('user-1');

      // Force expiry into the past
      const session = store._sessions.get(token);
      session.expiresAt = new Date(Date.now() - 1000).toISOString();

      store.validateSession(token);
      expect(store._sessions.has(token)).toBe(false);
    });

    test('returns null for null token (does not throw)', () => {
      expect(store.validateSession(null)).toBeNull();
    });

    test('returns null for undefined token (does not throw)', () => {
      expect(store.validateSession(undefined)).toBeNull();
    });

    test('returned object is a copy, not the internal reference', async () => {
      const token = await store.createSession('user-1');
      const result = store.validateSession(token);
      result.userId = 'tampered';

      const fresh = store.validateSession(token);
      expect(fresh.userId).toBe('user-1');
    });
  });

  // ── revokeSession ───────────────────────────────────────────────────

  describe('revokeSession', () => {
    test('makes a previously valid token invalid', async () => {
      const token = await store.createSession('user-1');
      expect(store.validateSession(token)).not.toBeNull();

      await store.revokeSession(token);
      expect(store.validateSession(token)).toBeNull();
    });

    test('returns true when token existed', async () => {
      const token = await store.createSession('user-1');
      const result = await store.revokeSession(token);
      expect(result).toBe(true);
    });

    test('returns false for unknown token', async () => {
      const result = await store.revokeSession('nope');
      expect(result).toBe(false);
    });
  });

  // ── revokeAllForUser ────────────────────────────────────────────────

  describe('revokeAllForUser', () => {
    test('removes all sessions for the target user', async () => {
      await store.createSession('user-1');
      await store.createSession('user-1');
      await store.createSession('user-1');

      const removed = await store.revokeAllForUser('user-1');
      expect(removed).toBe(3);
      expect(store.getSessionCountForUser('user-1')).toBe(0);
    });

    test('does not remove sessions belonging to other users', async () => {
      const kept = await store.createSession('user-2');
      await store.createSession('user-1');
      await store.createSession('user-1');

      await store.revokeAllForUser('user-1');

      expect(store.validateSession(kept)).not.toBeNull();
      expect(store.getSessionCountForUser('user-2')).toBe(1);
    });

    test('returns 0 when user has no sessions', async () => {
      const removed = await store.revokeAllForUser('ghost');
      expect(removed).toBe(0);
    });
  });

  // ── getActiveSessions ───────────────────────────────────────────────

  describe('getActiveSessions', () => {
    test('returns truncated tokens (first 8 chars + "...")', async () => {
      const token = await store.createSession('user-1');
      const sessions = store.getActiveSessions();

      expect(sessions).toHaveLength(1);
      expect(sessions[0].token).toBe(token.slice(0, 8) + '...');
      expect(sessions[0].userId).toBe('user-1');
      expect(sessions[0].createdAt).toBeDefined();
      expect(sessions[0].expiresAt).toBeDefined();
    });

    test('excludes expired sessions', async () => {
      const token = await store.createSession('user-1');
      const session = store._sessions.get(token);
      session.expiresAt = new Date(Date.now() - 1000).toISOString();

      const sessions = store.getActiveSessions();
      expect(sessions).toHaveLength(0);
    });

    test('returns empty array when store is empty', () => {
      expect(store.getActiveSessions()).toEqual([]);
    });
  });

  // ── getSessionCountForUser ──────────────────────────────────────────

  describe('getSessionCountForUser', () => {
    test('returns correct count', async () => {
      expect(store.getSessionCountForUser('user-1')).toBe(0);

      await store.createSession('user-1');
      expect(store.getSessionCountForUser('user-1')).toBe(1);

      await store.createSession('user-1');
      expect(store.getSessionCountForUser('user-1')).toBe(2);
    });

    test('does not count sessions from other users', async () => {
      await store.createSession('user-1');
      await store.createSession('user-2');
      await store.createSession('user-2');

      expect(store.getSessionCountForUser('user-1')).toBe(1);
      expect(store.getSessionCountForUser('user-2')).toBe(2);
    });

    test('does not count expired sessions', async () => {
      const token = await store.createSession('user-1');
      await store.createSession('user-1');

      // Expire the first one
      const session = store._sessions.get(token);
      session.expiresAt = new Date(Date.now() - 1000).toISOString();

      expect(store.getSessionCountForUser('user-1')).toBe(1);
    });
  });

  // ── Multiple sessions per user ──────────────────────────────────────

  describe('multiple sessions per user', () => {
    test('each session has its own independent token', async () => {
      const t1 = await store.createSession('user-1');
      const t2 = await store.createSession('user-1');

      expect(store.validateSession(t1)).not.toBeNull();
      expect(store.validateSession(t2)).not.toBeNull();
      expect(store.validateSession(t1).userId).toBe('user-1');
      expect(store.validateSession(t2).userId).toBe('user-1');
    });

    test('revoking one session does not affect the other', async () => {
      const t1 = await store.createSession('user-1');
      const t2 = await store.createSession('user-1');

      await store.revokeSession(t1);

      expect(store.validateSession(t1)).toBeNull();
      expect(store.validateSession(t2)).not.toBeNull();
    });
  });

  // ── File persistence round-trip ─────────────────────────────────────

  describe('file persistence', () => {
    test('save then load round-trip preserves sessions', async () => {
      const token = await store.createSession('user-persist');

      // Load into a fresh store from the same file
      const store2 = new SessionStore({ filePath });
      await store2.loadFromFile();

      const result = store2.validateSession(token);
      expect(result).not.toBeNull();
      expect(result.userId).toBe('user-persist');
    });

    test('loading from a non-existent file yields an empty store (no error)', async () => {
      const store2 = new SessionStore({ filePath: tmpSessionPath('-nonexistent') });
      await expect(store2.loadFromFile()).resolves.not.toThrow();
      expect(store2.getActiveSessions()).toEqual([]);
    });

    test('multiple sessions survive round-trip', async () => {
      const t1 = await store.createSession('alice');
      const t2 = await store.createSession('bob');

      const store2 = new SessionStore({ filePath });
      await store2.loadFromFile();

      expect(store2.validateSession(t1).userId).toBe('alice');
      expect(store2.validateSession(t2).userId).toBe('bob');
    });
  });

  // ── Encryption round-trip ───────────────────────────────────────────

  describe('encryption', () => {
    const ENC_KEY = 'test-encryption-key-do-not-use-in-production';

    test('encrypted save then load round-trip preserves sessions', async () => {
      const encPath = tmpSessionPath('-enc');
      tempFiles.push(encPath);

      const encStore = new SessionStore({ filePath: encPath, encryptionKey: ENC_KEY });
      const token = await encStore.createSession('secure-user');

      // Verify the file on disk looks encrypted
      const raw = JSON.parse(await fs.readFile(encPath, 'utf8'));
      expect(raw.encrypted).toBe(true);
      expect(raw.iv).toBeDefined();
      expect(raw.authTag).toBeDefined();
      expect(raw.data).toBeDefined();

      // Load into a fresh store with the same key
      const encStore2 = new SessionStore({ filePath: encPath, encryptionKey: ENC_KEY });
      await encStore2.loadFromFile();

      const result = encStore2.validateSession(token);
      expect(result).not.toBeNull();
      expect(result.userId).toBe('secure-user');

      await rm(encPath);
    });

    test('loading encrypted file without key throws', async () => {
      const encPath = tmpSessionPath('-enc-nokey');
      tempFiles.push(encPath);

      const encStore = new SessionStore({ filePath: encPath, encryptionKey: ENC_KEY });
      await encStore.createSession('user-1');

      const plainStore = new SessionStore({ filePath: encPath });
      await expect(plainStore.loadFromFile()).rejects.toThrow(/encryptionKey/);

      await rm(encPath);
    });

    test('loading encrypted file with wrong key throws', async () => {
      const encPath = tmpSessionPath('-enc-wrongkey');
      tempFiles.push(encPath);

      const encStore = new SessionStore({ filePath: encPath, encryptionKey: ENC_KEY });
      await encStore.createSession('user-1');

      const wrongStore = new SessionStore({ filePath: encPath, encryptionKey: 'wrong-key' });
      await expect(wrongStore.loadFromFile()).rejects.toThrow();

      await rm(encPath);
    });
  });

  // ── Edge cases ──────────────────────────────────────────────────────

  describe('edge cases', () => {
    test('custom expiresInDays is respected', async () => {
      const token = await store.createSession('user-1', { expiresInDays: 7 });
      const session = store.validateSession(token);

      const expectedExpiry = Date.now() + 7 * 24 * 60 * 60 * 1000;
      const actualExpiry = new Date(session.expiresAt).getTime();

      // Allow 5 seconds of drift
      expect(Math.abs(actualExpiry - expectedExpiry)).toBeLessThan(5000);
    });

    test('session created with 0 days expires immediately', async () => {
      const token = await store.createSession('user-1', { expiresInDays: 0 });

      // The token was created at "now" with 0 days added, so expiresAt === createdAt.
      // Give it a moment and it should be expired.
      const session = store._sessions.get(token);
      session.expiresAt = new Date(Date.now() - 1).toISOString();

      expect(store.validateSession(token)).toBeNull();
    });

    test('atomic write cleans up .tmp file on success', async () => {
      await store.createSession('user-1');

      // After save, .tmp should not exist
      let tmpExists = true;
      try {
        await fs.access(filePath + '.tmp');
      } catch {
        tmpExists = false;
      }
      expect(tmpExists).toBe(false);
    });
  });
});
