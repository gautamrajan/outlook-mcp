const PerUserTokenStorage = require('../../auth/per-user-token-storage');

describe('PerUserTokenStorage', () => {
  let storage;

  beforeEach(() => {
    storage = new PerUserTokenStorage();
  });

  function makeTokenData(expiresIn = 3600) {
    return {
      access_token: 'test-access-token',
      refresh_token: 'test-refresh-token',
      expires_in: expiresIn,
      scope: 'Mail.Read User.Read',
    };
  }

  // Test 1: Store and retrieve a token for a user
  test('should store and retrieve a token for a user', () => {
    const tokenData = makeTokenData();
    storage.setTokenForUser('user-1', tokenData);

    const token = storage.getTokenForUser('user-1');
    expect(token).toBe('test-access-token');
  });

  // Test 2: Return null for unknown user
  test('should return null for an unknown user', () => {
    const token = storage.getTokenForUser('nonexistent-user');
    expect(token).toBeNull();
  });

  // Test 3: Token expiry detection (expired token returns null from getTokenForUser)
  test('should return null for an expired token', () => {
    const tokenData = makeTokenData();
    storage.setTokenForUser('user-1', tokenData);

    // Manually set expires_at to the past
    const stored = storage._getUserData('user-1');
    stored.expires_at = Date.now() - 1000;

    const token = storage.getTokenForUser('user-1');
    expect(token).toBeNull();
  });

  // Test 4: Token within 5-minute buffer is treated as expired
  test('should treat a token within the 5-minute buffer as expired', () => {
    const tokenData = makeTokenData();
    storage.setTokenForUser('user-1', tokenData);

    // Set expires_at to 4 minutes from now (within the 5-minute buffer)
    const stored = storage._getUserData('user-1');
    stored.expires_at = Date.now() + (4 * 60 * 1000);

    const token = storage.getTokenForUser('user-1');
    expect(token).toBeNull();
    expect(storage.isTokenExpired('user-1')).toBe(true);
  });

  // Test 5: Token that hasn't expired yet returns valid access_token
  test('should return a valid access_token for a non-expired token', () => {
    const tokenData = makeTokenData(3600); // 1 hour
    storage.setTokenForUser('user-1', tokenData);

    const token = storage.getTokenForUser('user-1');
    expect(token).toBe('test-access-token');
    expect(storage.isTokenExpired('user-1')).toBe(false);
  });

  // Test 6: Invalidate user forces token to be treated as expired
  test('should treat an invalidated user token as expired', () => {
    const tokenData = makeTokenData();
    storage.setTokenForUser('user-1', tokenData);

    storage.invalidateUser('user-1');

    expect(storage.isTokenExpired('user-1')).toBe(true);
    expect(storage.getTokenForUser('user-1')).toBeNull();
  });

  // Test 7: Remove user deletes all data
  test('should remove all data for a user', () => {
    const tokenData = makeTokenData();
    storage.setTokenForUser('user-1', tokenData);

    storage.removeUser('user-1');

    expect(storage.getTokenForUser('user-1')).toBeNull();
    expect(storage.getActiveUserCount()).toBe(0);
  });

  // Test 8: Multiple users can store tokens independently
  test('should store tokens for multiple users independently', () => {
    storage.setTokenForUser('user-1', {
      access_token: 'token-alpha',
      refresh_token: 'refresh-alpha',
      expires_in: 3600,
      scope: 'Mail.Read',
    });

    storage.setTokenForUser('user-2', {
      access_token: 'token-beta',
      refresh_token: 'refresh-beta',
      expires_in: 3600,
      scope: 'Mail.Read',
    });

    expect(storage.getTokenForUser('user-1')).toBe('token-alpha');
    expect(storage.getTokenForUser('user-2')).toBe('token-beta');

    // Invalidating one user should not affect the other
    storage.invalidateUser('user-1');
    expect(storage.getTokenForUser('user-1')).toBeNull();
    expect(storage.getTokenForUser('user-2')).toBe('token-beta');
  });

  // Test 9: getActiveUserCount reflects actual stored users
  test('should return correct active user count', () => {
    expect(storage.getActiveUserCount()).toBe(0);

    storage.setTokenForUser('user-1', makeTokenData());
    expect(storage.getActiveUserCount()).toBe(1);

    storage.setTokenForUser('user-2', makeTokenData());
    expect(storage.getActiveUserCount()).toBe(2);

    storage.removeUser('user-1');
    expect(storage.getActiveUserCount()).toBe(1);

    storage.removeUser('user-2');
    expect(storage.getActiveUserCount()).toBe(0);
  });

  // Test 10: Throw on null/undefined userId
  test('should throw on null or undefined userId', () => {
    const tokenData = makeTokenData();

    expect(() => storage.getTokenForUser(null)).toThrow();
    expect(() => storage.getTokenForUser(undefined)).toThrow();
    expect(() => storage.setTokenForUser(null, tokenData)).toThrow();
    expect(() => storage.setTokenForUser(undefined, tokenData)).toThrow();
    expect(() => storage.isTokenExpired(null)).toThrow();
    expect(() => storage.isTokenExpired(undefined)).toThrow();
    expect(() => storage.invalidateUser(null)).toThrow();
    expect(() => storage.invalidateUser(undefined)).toThrow();
    expect(() => storage.removeUser(null)).toThrow();
    expect(() => storage.removeUser(undefined)).toThrow();
  });

  // Additional: setTokenForUser correctly calculates expires_at
  test('should calculate expires_at from expires_in', () => {
    const before = Date.now();
    storage.setTokenForUser('user-1', makeTokenData(7200)); // 2 hours
    const after = Date.now();

    const stored = storage._getUserData('user-1');
    // expires_at should be approximately now + 7200 * 1000
    expect(stored.expires_at).toBeGreaterThanOrEqual(before + 7200 * 1000);
    expect(stored.expires_at).toBeLessThanOrEqual(after + 7200 * 1000);
  });

  // ── Edge cases ─────────────────────────────────────────────────────
  describe('edge cases', () => {
    test('overwriting a token for an existing user replaces the old data', () => {
      storage.setTokenForUser('user-1', {
        access_token: 'old-token',
        refresh_token: 'old-refresh',
        expires_in: 3600,
        scope: 'Mail.Read',
      });
      storage.setTokenForUser('user-1', {
        access_token: 'new-token',
        refresh_token: 'new-refresh',
        expires_in: 3600,
        scope: 'Mail.Read',
      });

      expect(storage.getTokenForUser('user-1')).toBe('new-token');
      expect(storage.getActiveUserCount()).toBe(1);
    });

    test('isTokenExpired returns true for a user with no stored token', () => {
      expect(storage.isTokenExpired('nonexistent')).toBe(true);
    });

    test('invalidateUser is a no-op for unknown users', () => {
      expect(() => storage.invalidateUser('ghost')).not.toThrow();
    });

    test('removeUser is a no-op for unknown users', () => {
      expect(() => storage.removeUser('ghost')).not.toThrow();
      expect(storage.getActiveUserCount()).toBe(0);
    });

    test('token at exactly the 5-minute boundary is treated as expired', () => {
      storage.setTokenForUser('user-1', makeTokenData());

      const data = storage._getUserData('user-1');
      // Set expires_at to exactly 5 minutes from now
      data.expires_at = Date.now() + (5 * 60 * 1000);

      // At the boundary: Date.now() >= (expires_at - 5min buffer) => true
      expect(storage.isTokenExpired('user-1')).toBe(true);
      expect(storage.getTokenForUser('user-1')).toBeNull();
    });

    test('token just outside the 5-minute buffer is still valid', () => {
      storage.setTokenForUser('user-1', makeTokenData());

      const data = storage._getUserData('user-1');
      // Set expires_at to 5 minutes + 2 seconds from now (safely outside buffer)
      data.expires_at = Date.now() + (5 * 60 * 1000) + 2000;

      expect(storage.isTokenExpired('user-1')).toBe(false);
      expect(storage.getTokenForUser('user-1')).toBe('test-access-token');
    });

    test('stores scope from token data', () => {
      storage.setTokenForUser('user-1', {
        access_token: 'token',
        refresh_token: 'refresh',
        expires_in: 3600,
        scope: 'Mail.ReadWrite Calendars.Read',
      });

      const data = storage._getUserData('user-1');
      expect(data.scope).toBe('Mail.ReadWrite Calendars.Read');
    });

    test('removing one user does not affect other users', () => {
      storage.setTokenForUser('user-1', makeTokenData());
      storage.setTokenForUser('user-2', makeTokenData());

      storage.removeUser('user-1');

      expect(storage.getTokenForUser('user-1')).toBeNull();
      expect(storage.getTokenForUser('user-2')).toBe('test-access-token');
      expect(storage.getActiveUserCount()).toBe(1);
    });
  });
});
