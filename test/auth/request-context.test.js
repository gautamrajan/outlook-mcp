const { AsyncLocalStorage } = require('node:async_hooks');
const { requestContext, getUserContext, isHostedMode } = require('../../auth/request-context');

describe('request-context', () => {
  // Test 1: requestContext is an AsyncLocalStorage instance
  test('requestContext should be an AsyncLocalStorage instance', () => {
    expect(requestContext).toBeInstanceOf(AsyncLocalStorage);
  });

  // Test 2: getUserContext() returns null when not in a run context
  test('getUserContext() should return null outside of requestContext.run()', () => {
    expect(getUserContext()).toBeNull();
  });

  // Test 3: getUserContext() returns the store when inside requestContext.run()
  test('getUserContext() should return the store inside requestContext.run()', async () => {
    const userCtx = { userId: 'user-abc', entraToken: 'entra-token-xyz' };

    await requestContext.run(userCtx, async () => {
      const result = getUserContext();
      expect(result).toBe(userCtx);
      expect(result.userId).toBe('user-abc');
      expect(result.entraToken).toBe('entra-token-xyz');
    });
  });

  // Test 4: isHostedMode() returns false outside of run context
  test('isHostedMode() should return false outside of requestContext.run()', () => {
    expect(isHostedMode()).toBe(false);
  });

  // Test 5: isHostedMode() returns true inside run context
  test('isHostedMode() should return true inside requestContext.run()', async () => {
    const userCtx = { userId: 'user-abc', entraToken: 'entra-token-xyz' };

    await requestContext.run(userCtx, async () => {
      expect(isHostedMode()).toBe(true);
    });
  });

  test('isHostedMode() should return false for context without userId', async () => {
    const partialCtx = { entraToken: 'entra-token-xyz' };

    await requestContext.run(partialCtx, async () => {
      expect(isHostedMode()).toBe(false);
    });
  });

  // Test 6: Nested runs properly scope context
  test('nested runs should properly scope context', async () => {
    const outerCtx = { userId: 'outer-user', entraToken: 'outer-token' };
    const innerCtx = { userId: 'inner-user', entraToken: 'inner-token' };

    await requestContext.run(outerCtx, async () => {
      expect(getUserContext()).toBe(outerCtx);
      expect(getUserContext().userId).toBe('outer-user');

      await requestContext.run(innerCtx, async () => {
        expect(getUserContext()).toBe(innerCtx);
        expect(getUserContext().userId).toBe('inner-user');
      });

      // After inner run completes, outer context is restored
      expect(getUserContext()).toBe(outerCtx);
      expect(getUserContext().userId).toBe('outer-user');
    });
  });

  // Test 7: Context is isolated between concurrent async operations
  test('context should be isolated between concurrent async operations', async () => {
    const ctxA = { userId: 'user-A', entraToken: 'token-A' };
    const ctxB = { userId: 'user-B', entraToken: 'token-B' };

    const results = [];

    const promiseA = requestContext.run(ctxA, async () => {
      // Yield to let B start
      await new Promise(resolve => setImmediate(resolve));
      results.push({ op: 'A', context: getUserContext() });
    });

    const promiseB = requestContext.run(ctxB, async () => {
      // Yield to let A continue
      await new Promise(resolve => setImmediate(resolve));
      results.push({ op: 'B', context: getUserContext() });
    });

    await Promise.all([promiseA, promiseB]);

    const resultA = results.find(r => r.op === 'A');
    const resultB = results.find(r => r.op === 'B');

    expect(resultA.context).toBe(ctxA);
    expect(resultA.context.userId).toBe('user-A');
    expect(resultB.context).toBe(ctxB);
    expect(resultB.context.userId).toBe('user-B');
  });

  // Test 8: getUserContext() returns null again after run completes
  test('getUserContext() should return null after requestContext.run() completes', async () => {
    const userCtx = { userId: 'temp-user', entraToken: 'temp-token' };

    await requestContext.run(userCtx, async () => {
      expect(getUserContext()).toBe(userCtx);
    });

    expect(getUserContext()).toBeNull();
  });
});
