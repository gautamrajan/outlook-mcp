const PerUserTokenStorage = require('../../auth/per-user-token-storage');
const fs = require('fs');
const fsp = require('fs/promises');
const path = require('path');
const os = require('os');

// ── Helpers ─────────────────────────────────────────────────────────

function makeTokenData(expiresIn = 3600) {
  return {
    accessToken: 'test-access-token',
    refreshToken: 'test-refresh-token',
    expiresIn,
    scopes: 'Mail.Read User.Read',
    email: 'test@example.com',
    name: 'Test User',
  };
}

/**
 * Creates a temp directory and returns { dir, filePath, cleanup }.
 */
async function makeTempStore() {
  const dir = await fsp.mkdtemp(path.join(os.tmpdir(), 'per-user-token-test-'));
  const filePath = path.join(dir, 'tokens.json');
  const cleanup = async () => {
    await fsp.rm(dir, { recursive: true, force: true });
  };
  return { dir, filePath, cleanup };
}

// ══════════════════════════════════════════════════════════════════════
// In-memory behaviour (no file persistence)
// ══════════════════════════════════════════════════════════════════════

describe('PerUserTokenStorage — in-memory', () => {
  let storage;

  beforeEach(() => {
    storage = new PerUserTokenStorage();
  });

  // Test 1: Store and retrieve a token for a user
  test('should store and retrieve a token for a user', async () => {
    await storage.setTokensForUser('user-1', makeTokenData());

    const token = storage.getTokenForUser('user-1');
    expect(token).toBe('test-access-token');
  });

  // Test 2: Return null for unknown user
  test('should return null for an unknown user', () => {
    const token = storage.getTokenForUser('nonexistent-user');
    expect(token).toBeNull();
  });

  // Test 3: Token expiry detection (expired token returns null)
  test('should return null for an expired token', async () => {
    await storage.setTokensForUser('user-1', makeTokenData());

    // Manually set expiresAt to the past
    const stored = storage._getUserData('user-1');
    stored.expiresAt = Date.now() - 1000;

    const token = storage.getTokenForUser('user-1');
    expect(token).toBeNull();
  });

  // Test 4: Token within 5-minute buffer is treated as expired
  test('should treat a token within the 5-minute buffer as expired', async () => {
    await storage.setTokensForUser('user-1', makeTokenData());

    // Set expiresAt to 4 minutes from now (within the 5-minute buffer)
    const stored = storage._getUserData('user-1');
    stored.expiresAt = Date.now() + (4 * 60 * 1000);

    const token = storage.getTokenForUser('user-1');
    expect(token).toBeNull();
    expect(storage.isTokenExpired('user-1')).toBe(true);
  });

  // Test 5: Token that hasn't expired yet returns valid accessToken
  test('should return a valid accessToken for a non-expired token', async () => {
    await storage.setTokensForUser('user-1', makeTokenData(3600));

    const token = storage.getTokenForUser('user-1');
    expect(token).toBe('test-access-token');
    expect(storage.isTokenExpired('user-1')).toBe(false);
  });

  // Test 6: Invalidate user forces token to be treated as expired
  test('should treat an invalidated user token as expired', async () => {
    await storage.setTokensForUser('user-1', makeTokenData());
    await storage.invalidateUser('user-1');

    expect(storage.isTokenExpired('user-1')).toBe(true);
    expect(storage.getTokenForUser('user-1')).toBeNull();
  });

  // Test 7: Remove user deletes all data
  test('should remove all data for a user', async () => {
    await storage.setTokensForUser('user-1', makeTokenData());
    await storage.removeUser('user-1');

    expect(storage.getTokenForUser('user-1')).toBeNull();
    expect(storage.getActiveUserCount()).toBe(0);
  });

  // Test 8: Multiple users can store tokens independently
  test('should store tokens for multiple users independently', async () => {
    await storage.setTokensForUser('user-1', {
      accessToken: 'token-alpha',
      refreshToken: 'refresh-alpha',
      expiresIn: 3600,
      scopes: 'Mail.Read',
    });

    await storage.setTokensForUser('user-2', {
      accessToken: 'token-beta',
      refreshToken: 'refresh-beta',
      expiresIn: 3600,
      scopes: 'Mail.Read',
    });

    expect(storage.getTokenForUser('user-1')).toBe('token-alpha');
    expect(storage.getTokenForUser('user-2')).toBe('token-beta');

    // Invalidating one user should not affect the other
    await storage.invalidateUser('user-1');
    expect(storage.getTokenForUser('user-1')).toBeNull();
    expect(storage.getTokenForUser('user-2')).toBe('token-beta');
  });

  // Test 9: getActiveUserCount reflects actual stored users
  test('should return correct active user count', async () => {
    expect(storage.getActiveUserCount()).toBe(0);

    await storage.setTokensForUser('user-1', makeTokenData());
    expect(storage.getActiveUserCount()).toBe(1);

    await storage.setTokensForUser('user-2', makeTokenData());
    expect(storage.getActiveUserCount()).toBe(2);

    await storage.removeUser('user-1');
    expect(storage.getActiveUserCount()).toBe(1);

    await storage.removeUser('user-2');
    expect(storage.getActiveUserCount()).toBe(0);
  });

  // Test 10: Throw on null/undefined userId
  test('should throw on null or undefined userId', async () => {
    const tokenData = makeTokenData();

    expect(() => storage.getTokenForUser(null)).toThrow();
    expect(() => storage.getTokenForUser(undefined)).toThrow();
    await expect(storage.setTokensForUser(null, tokenData)).rejects.toThrow();
    await expect(storage.setTokensForUser(undefined, tokenData)).rejects.toThrow();
    expect(() => storage.isTokenExpired(null)).toThrow();
    expect(() => storage.isTokenExpired(undefined)).toThrow();
    await expect(storage.invalidateUser(null)).rejects.toThrow();
    await expect(storage.invalidateUser(undefined)).rejects.toThrow();
    await expect(storage.removeUser(null)).rejects.toThrow();
    await expect(storage.removeUser(undefined)).rejects.toThrow();
  });

  // Test: setTokensForUser correctly calculates expiresAt
  test('should calculate expiresAt from expiresIn', async () => {
    const before = Date.now();
    await storage.setTokensForUser('user-1', makeTokenData(7200)); // 2 hours
    const after = Date.now();

    const stored = storage._getUserData('user-1');
    expect(stored.expiresAt).toBeGreaterThanOrEqual(before + 7200 * 1000);
    expect(stored.expiresAt).toBeLessThanOrEqual(after + 7200 * 1000);
  });

  // ── getRefreshToken ───────────────────────────────────────────────

  describe('getRefreshToken', () => {
    test('returns refresh token for a known user', async () => {
      await storage.setTokensForUser('user-1', makeTokenData());
      expect(storage.getRefreshToken('user-1')).toBe('test-refresh-token');
    });

    test('returns null for an unknown user', () => {
      expect(storage.getRefreshToken('nonexistent')).toBeNull();
    });

    test('throws on null/undefined userId', () => {
      expect(() => storage.getRefreshToken(null)).toThrow();
      expect(() => storage.getRefreshToken(undefined)).toThrow();
    });
  });

  // ── getUserInfo ───────────────────────────────────────────────────

  describe('getUserInfo', () => {
    test('returns email and name for a known user', async () => {
      await storage.setTokensForUser('user-1', makeTokenData());
      const info = storage.getUserInfo('user-1');
      expect(info).toEqual({ email: 'test@example.com', name: 'Test User' });
    });

    test('returns null for an unknown user', () => {
      expect(storage.getUserInfo('nonexistent')).toBeNull();
    });

    test('returns null email/name when not provided', async () => {
      await storage.setTokensForUser('user-1', {
        accessToken: 'tok',
        refreshToken: 'ref',
        expiresIn: 3600,
        scopes: 'Mail.Read',
      });
      const info = storage.getUserInfo('user-1');
      expect(info).toEqual({ email: null, name: null });
    });

    test('throws on null/undefined userId', () => {
      expect(() => storage.getUserInfo(null)).toThrow();
      expect(() => storage.getUserInfo(undefined)).toThrow();
    });
  });

  // ── getAllUsers ────────────────────────────────────────────────────

  describe('getAllUsers', () => {
    test('returns empty array with no users', () => {
      expect(storage.getAllUsers()).toEqual([]);
    });

    test('returns all users with info and validity status', async () => {
      await storage.setTokensForUser('user-1', makeTokenData());
      await storage.setTokensForUser('user-2', {
        accessToken: 'tok2',
        refreshToken: 'ref2',
        expiresIn: 3600,
        scopes: 'Mail.Read',
        email: 'two@example.com',
        name: 'User Two',
      });

      // Expire user-1
      await storage.invalidateUser('user-1');

      const users = storage.getAllUsers();
      expect(users).toHaveLength(2);

      const u1 = users.find((u) => u.userId === 'user-1');
      const u2 = users.find((u) => u.userId === 'user-2');

      expect(u1.hasValidToken).toBe(false);
      expect(u2.hasValidToken).toBe(true);
      expect(u2.email).toBe('two@example.com');
      expect(u2.name).toBe('User Two');
    });
  });

  // ── Edge cases ────────────────────────────────────────────────────

  describe('edge cases', () => {
    test('overwriting a token for an existing user replaces the old data', async () => {
      await storage.setTokensForUser('user-1', {
        accessToken: 'old-token',
        refreshToken: 'old-refresh',
        expiresIn: 3600,
        scopes: 'Mail.Read',
      });
      await storage.setTokensForUser('user-1', {
        accessToken: 'new-token',
        refreshToken: 'new-refresh',
        expiresIn: 3600,
        scopes: 'Mail.Read',
      });

      expect(storage.getTokenForUser('user-1')).toBe('new-token');
      expect(storage.getActiveUserCount()).toBe(1);
    });

    test('isTokenExpired returns true for a user with no stored token', () => {
      expect(storage.isTokenExpired('nonexistent')).toBe(true);
    });

    test('invalidateUser is a no-op for unknown users', async () => {
      await expect(storage.invalidateUser('ghost')).resolves.toBeUndefined();
    });

    test('removeUser is a no-op for unknown users', async () => {
      await expect(storage.removeUser('ghost')).resolves.toBeUndefined();
      expect(storage.getActiveUserCount()).toBe(0);
    });

    test('token at exactly the 5-minute boundary is treated as expired', async () => {
      await storage.setTokensForUser('user-1', makeTokenData());

      const data = storage._getUserData('user-1');
      data.expiresAt = Date.now() + (5 * 60 * 1000);

      expect(storage.isTokenExpired('user-1')).toBe(true);
      expect(storage.getTokenForUser('user-1')).toBeNull();
    });

    test('token just outside the 5-minute buffer is still valid', async () => {
      await storage.setTokensForUser('user-1', makeTokenData());

      const data = storage._getUserData('user-1');
      data.expiresAt = Date.now() + (5 * 60 * 1000) + 2000;

      expect(storage.isTokenExpired('user-1')).toBe(false);
      expect(storage.getTokenForUser('user-1')).toBe('test-access-token');
    });

    test('stores scopes from token data', async () => {
      await storage.setTokensForUser('user-1', {
        accessToken: 'token',
        refreshToken: 'refresh',
        expiresIn: 3600,
        scopes: 'Mail.ReadWrite Calendars.Read',
      });

      const data = storage._getUserData('user-1');
      expect(data.scopes).toBe('Mail.ReadWrite Calendars.Read');
    });

    test('removing one user does not affect other users', async () => {
      await storage.setTokensForUser('user-1', makeTokenData());
      await storage.setTokensForUser('user-2', makeTokenData());

      await storage.removeUser('user-1');

      expect(storage.getTokenForUser('user-1')).toBeNull();
      expect(storage.getTokenForUser('user-2')).toBe('test-access-token');
      expect(storage.getActiveUserCount()).toBe(1);
    });
  });
});

// ══════════════════════════════════════════════════════════════════════
// File persistence — plain JSON (no encryption)
// ══════════════════════════════════════════════════════════════════════

describe('PerUserTokenStorage — file persistence (plain)', () => {
  let tmpCtx;

  afterEach(async () => {
    if (tmpCtx) await tmpCtx.cleanup();
  });

  test('saveToFile writes and loadFromFile reads back correctly', async () => {
    tmpCtx = await makeTempStore();
    const storage = new PerUserTokenStorage({ filePath: tmpCtx.filePath });

    await storage.setTokensForUser('user-1', makeTokenData());
    await storage.setTokensForUser('user-2', {
      accessToken: 'tok2',
      refreshToken: 'ref2',
      expiresIn: 1800,
      scopes: 'Calendars.Read',
      email: 'u2@example.com',
      name: 'User 2',
    });

    // Create a fresh instance and load
    const storage2 = new PerUserTokenStorage({ filePath: tmpCtx.filePath });
    await storage2.loadFromFile();

    expect(storage2.getActiveUserCount()).toBe(2);
    expect(storage2.getTokenForUser('user-1')).toBe('test-access-token');
    expect(storage2.getTokenForUser('user-2')).toBe('tok2');
    expect(storage2.getUserInfo('user-2')).toEqual({ email: 'u2@example.com', name: 'User 2' });
  });

  test('loadFromFile with non-existent file leaves Map empty', async () => {
    tmpCtx = await makeTempStore();
    const storage = new PerUserTokenStorage({ filePath: path.join(tmpCtx.dir, 'does-not-exist.json') });

    await storage.loadFromFile();
    expect(storage.getActiveUserCount()).toBe(0);
  });

  test('mutations auto-persist to file', async () => {
    tmpCtx = await makeTempStore();
    const storage = new PerUserTokenStorage({ filePath: tmpCtx.filePath });

    await storage.setTokensForUser('user-1', makeTokenData());

    // File should exist now
    const stat = await fsp.stat(tmpCtx.filePath);
    expect(stat.isFile()).toBe(true);

    // Remove user — file should be updated
    await storage.removeUser('user-1');

    const storage2 = new PerUserTokenStorage({ filePath: tmpCtx.filePath });
    await storage2.loadFromFile();
    expect(storage2.getActiveUserCount()).toBe(0);
  });

  test('invalidateUser persists to file', async () => {
    tmpCtx = await makeTempStore();
    const storage = new PerUserTokenStorage({ filePath: tmpCtx.filePath });

    await storage.setTokensForUser('user-1', makeTokenData());
    await storage.invalidateUser('user-1');

    const storage2 = new PerUserTokenStorage({ filePath: tmpCtx.filePath });
    await storage2.loadFromFile();
    expect(storage2.isTokenExpired('user-1')).toBe(true);
  });

  test('atomic write uses .tmp file', async () => {
    tmpCtx = await makeTempStore();
    const storage = new PerUserTokenStorage({ filePath: tmpCtx.filePath });

    // After a save completes, the .tmp file should be gone (renamed to the real file)
    await storage.setTokensForUser('user-1', makeTokenData());

    const tmpFile = tmpCtx.filePath + '.tmp';
    let tmpExists = false;
    try {
      await fsp.access(tmpFile);
      tmpExists = true;
    } catch {
      tmpExists = false;
    }
    expect(tmpExists).toBe(false);

    // But the real file should exist
    const realExists = fs.existsSync(tmpCtx.filePath);
    expect(realExists).toBe(true);
  });

  test('saveToFile is a no-op when no filePath is set', async () => {
    const storage = new PerUserTokenStorage();
    // Should not throw
    await storage.saveToFile();
  });

  test('loadFromFile is a no-op when no filePath is set', async () => {
    const storage = new PerUserTokenStorage();
    await storage.loadFromFile();
    expect(storage.getActiveUserCount()).toBe(0);
  });

  test('plain JSON file is human-readable', async () => {
    tmpCtx = await makeTempStore();
    const storage = new PerUserTokenStorage({ filePath: tmpCtx.filePath });
    await storage.setTokensForUser('user-1', makeTokenData());

    const raw = await fsp.readFile(tmpCtx.filePath, 'utf8');
    const parsed = JSON.parse(raw);

    // Should be a plain object, not encrypted
    expect(parsed.encrypted).toBeUndefined();
    expect(parsed['user-1']).toBeDefined();
    expect(parsed['user-1'].accessToken).toBe('test-access-token');
  });
});

// ══════════════════════════════════════════════════════════════════════
// File persistence — encrypted
// ══════════════════════════════════════════════════════════════════════

describe('PerUserTokenStorage — file persistence (encrypted)', () => {
  const TEST_KEY = 'my-super-secret-encryption-key-for-testing';
  let tmpCtx;

  afterEach(async () => {
    if (tmpCtx) await tmpCtx.cleanup();
  });

  test('encryption round-trip: save encrypted, load encrypted, data matches', async () => {
    tmpCtx = await makeTempStore();
    const storage = new PerUserTokenStorage({ filePath: tmpCtx.filePath, encryptionKey: TEST_KEY });

    await storage.setTokensForUser('user-1', makeTokenData());
    await storage.setTokensForUser('user-2', {
      accessToken: 'tok2',
      refreshToken: 'ref2',
      expiresIn: 1800,
      scopes: 'Calendars.Read',
      email: 'enc@example.com',
      name: 'Encrypted User',
    });

    // Verify file is encrypted
    const raw = await fsp.readFile(tmpCtx.filePath, 'utf8');
    const parsed = JSON.parse(raw);
    expect(parsed.encrypted).toBe(true);
    expect(parsed.iv).toBeDefined();
    expect(parsed.authTag).toBeDefined();
    expect(parsed.data).toBeDefined();

    // Load into fresh instance
    const storage2 = new PerUserTokenStorage({ filePath: tmpCtx.filePath, encryptionKey: TEST_KEY });
    await storage2.loadFromFile();

    expect(storage2.getActiveUserCount()).toBe(2);
    expect(storage2.getTokenForUser('user-1')).toBe('test-access-token');
    expect(storage2.getRefreshToken('user-2')).toBe('ref2');
    expect(storage2.getUserInfo('user-2')).toEqual({ email: 'enc@example.com', name: 'Encrypted User' });
  });

  test('loading encrypted file without key throws', async () => {
    tmpCtx = await makeTempStore();
    const storage = new PerUserTokenStorage({ filePath: tmpCtx.filePath, encryptionKey: TEST_KEY });
    await storage.setTokensForUser('user-1', makeTokenData());

    // Try to load without the key
    const storage2 = new PerUserTokenStorage({ filePath: tmpCtx.filePath });
    await expect(storage2.loadFromFile()).rejects.toThrow('encrypted but no encryptionKey');
  });

  test('loading encrypted file with wrong key throws', async () => {
    tmpCtx = await makeTempStore();
    const storage = new PerUserTokenStorage({ filePath: tmpCtx.filePath, encryptionKey: TEST_KEY });
    await storage.setTokensForUser('user-1', makeTokenData());

    const storage2 = new PerUserTokenStorage({ filePath: tmpCtx.filePath, encryptionKey: 'wrong-key' });
    await expect(storage2.loadFromFile()).rejects.toThrow();
  });

  test('encrypted file does not contain plaintext tokens', async () => {
    tmpCtx = await makeTempStore();
    const storage = new PerUserTokenStorage({ filePath: tmpCtx.filePath, encryptionKey: TEST_KEY });
    await storage.setTokensForUser('user-1', makeTokenData());

    const raw = await fsp.readFile(tmpCtx.filePath, 'utf8');
    expect(raw).not.toContain('test-access-token');
    expect(raw).not.toContain('test-refresh-token');
  });
});
