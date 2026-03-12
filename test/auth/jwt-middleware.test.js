// Mock jwks-rsa to prevent jose ESM import from blowing up in Jest
jest.mock('jwks-rsa', () => jest.fn(() => ({ getSigningKey: jest.fn() })));

// Mock jwt-validator BEFORE it gets required by jwt-middleware
jest.mock('../../auth/jwt-validator');

const jwtMiddleware = require('../../auth/jwt-middleware');
const { validateEntraJwt } = require('../../auth/jwt-validator');

/**
 * Helper: creates a minimal mock Express request.
 */
function mockReq(headers = {}) {
  return { headers };
}

/**
 * Helper: creates a minimal mock Express response (unused by this middleware).
 */
function mockRes() {
  return {};
}

describe('jwtMiddleware', () => {
  let next;

  beforeEach(() => {
    jest.clearAllMocks();
    next = jest.fn();
  });

  test('should call next() with no enrichment when Authorization header is absent', async () => {
    const req = mockReq();
    const res = mockRes();

    await jwtMiddleware(req, res, next);

    expect(next).toHaveBeenCalledTimes(1);
    expect(next).toHaveBeenCalledWith(); // no error argument
    expect(req.entraUser).toBeUndefined();
    expect(req.entraToken).toBeUndefined();
    expect(validateEntraJwt).not.toHaveBeenCalled();
  });

  test('should call next() with no enrichment for non-Bearer Authorization header', async () => {
    const req = mockReq({ authorization: 'Basic dXNlcjpwYXNz' });
    const res = mockRes();

    await jwtMiddleware(req, res, next);

    expect(next).toHaveBeenCalledTimes(1);
    expect(next).toHaveBeenCalledWith();
    expect(req.entraUser).toBeUndefined();
    expect(req.entraToken).toBeUndefined();
    expect(validateEntraJwt).not.toHaveBeenCalled();
  });

  test('should set req.entraUser and req.entraToken for a valid Bearer JWT', async () => {
    const mockClaims = {
      oid: 'oid-123',
      sub: 'sub-456',
      preferred_username: 'user@example.com',
      name: 'Test User',
      scp: 'mcp.access',
    };
    validateEntraJwt.mockResolvedValue(mockClaims);

    const rawToken = 'valid.jwt.token';
    const req = mockReq({ authorization: `Bearer ${rawToken}` });
    const res = mockRes();

    await jwtMiddleware(req, res, next);

    expect(validateEntraJwt).toHaveBeenCalledWith(rawToken);
    expect(req.entraUser).toEqual({
      oid: 'oid-123',
      sub: 'sub-456',
      preferred_username: 'user@example.com',
      name: 'Test User',
      scp: 'mcp.access',
    });
    expect(req.entraToken).toBe(rawToken);
    expect(next).toHaveBeenCalledTimes(1);
    expect(next).toHaveBeenCalledWith(); // no error
  });

  test('should call next() without error when JWT validation fails', async () => {
    validateEntraJwt.mockRejectedValue(new Error('JWT validation failed: invalid signature'));

    const req = mockReq({ authorization: 'Bearer invalid.jwt.token' });
    const res = mockRes();

    await jwtMiddleware(req, res, next);

    expect(validateEntraJwt).toHaveBeenCalledWith('invalid.jwt.token');
    expect(req.entraUser).toBeUndefined();
    expect(req.entraToken).toBeUndefined();
    expect(next).toHaveBeenCalledTimes(1);
    expect(next).toHaveBeenCalledWith(); // no error -- pass through
  });

  test('should only extract claims specified in the middleware (oid, sub, preferred_username, name, scp)', async () => {
    const mockClaims = {
      oid: 'oid-789',
      sub: 'sub-012',
      preferred_username: 'admin@example.com',
      name: 'Admin User',
      scp: 'mcp.access admin.all',
      tid: 'tenant-id-should-not-appear',
      extra_claim: 'should-not-appear',
    };
    validateEntraJwt.mockResolvedValue(mockClaims);

    const req = mockReq({ authorization: 'Bearer some.jwt.token' });
    const res = mockRes();

    await jwtMiddleware(req, res, next);

    // Only the specified fields should be on entraUser
    expect(req.entraUser).toEqual({
      oid: 'oid-789',
      sub: 'sub-012',
      preferred_username: 'admin@example.com',
      name: 'Admin User',
      scp: 'mcp.access admin.all',
    });
    expect(req.entraUser.tid).toBeUndefined();
    expect(req.entraUser.extra_claim).toBeUndefined();
  });
});
