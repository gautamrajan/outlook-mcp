/**
 * Express middleware for Entra JWT authentication
 *
 * Optimistic middleware: attempts to validate the Authorization header as an
 * Entra JWT.  On success it enriches the request with decoded claims; on
 * failure it silently passes through so downstream auth (session-based) can
 * take over.  Never returns 401 — that decision belongs to the dual-path
 * auth layer in Wave 2.
 */
const { validateEntraJwt } = require('./jwt-validator');

/**
 * Extracts and validates a Bearer JWT from the request.
 *
 * Behaviour:
 *   - No Authorization header or non-Bearer scheme → next() (no enrichment)
 *   - Valid JWT → sets req.entraUser and req.entraToken, then next()
 *   - Invalid JWT → next() without error (pass-through)
 *
 * @param {import('express').Request} req
 * @param {import('express').Response} res
 * @param {import('express').NextFunction} next
 */
async function jwtMiddleware(req, res, next) {
  const authHeader = req.headers.authorization;

  // No header or not a Bearer token — let session middleware handle it
  if (!authHeader || !authHeader.startsWith('Bearer ')) {
    return next();
  }

  const token = authHeader.slice(7); // Strip "Bearer "

  try {
    const claims = await validateEntraJwt(token);

    req.entraUser = {
      oid: claims.oid,
      sub: claims.sub,
      preferred_username: claims.preferred_username,
      name: claims.name,
      scp: claims.scp,
    };
    req.entraToken = token;
  } catch (_err) {
    // Validation failed — this may be a session token, not a JWT.
    // Silently pass through; downstream auth decides whether to 401.
  }

  next();
}

module.exports = jwtMiddleware;
