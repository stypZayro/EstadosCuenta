// middlewares.js
const jwt = require('jsonwebtoken');
const { z } = require('zod');

const JWT_SECRET = process.env.JWT_SECRET || 'Zayroserver2025##';
const JWT_ISSUER = process.env.JWT_ISSUER || 'zayrocom';
const JWT_AUDIENCE = process.env.JWT_AUDIENCE || 'apizayrocom';

// -------------------- Zod: validador genérico --------------------
function validate(schema, where = 'query') {
  return (req, res, next) => {
    const data = req[where];
    const parsed = schema.safeParse(data);
    if (!parsed.success) {
      return res.status(400).json({
        error: 'validation',
        where,
        details: parsed.error.issues.map(i => ({
          path: i.path.join('.'),
          msg: i.message
        }))
      });
    }
    req[where] = parsed.data; // normalizado
    next();
  };
}

// -------------------- Auth: Bearer JWT real --------------------
function authBearer(req, res, next) {
  const h = req.get('Authorization') || '';
  // Soporta "Bearer <token>" con espacios y case-insensitive
  const m = h.match(/^Bearer\s+(.+)$/i);
  if (!m) return res.status(401).json({ error: 'No token' });

  try {
    const payload = jwt.verify(m[1], JWT_SECRET, {
      algorithms: ['HS256'],
      issuer: JWT_ISSUER,
      audience: JWT_AUDIENCE,
      clockTolerance: 5 // segundos de tolerancia de reloj
    });

    // payload típico: { sub, role, scope, iat, exp, jti? }
    req.user = {
      id: payload.sub,
      role: payload.role || 'user',
      scope: payload.scope || '' // ej: "read:sica write:sica"
    };

    // (Opcional) Si manejas lista de revocados por jti, checar aquí

    next();
  } catch (e) {
    return res.status(401).json({ error: 'Token inválido o expirado' });
  }
}

// -------------------- Autorización por scope/rol --------------------
function requireScope(scopeNeeded) {
  return (req, res, next) => {
    const scopes = (req.user?.scope || '').split(/\s+/).filter(Boolean);
    if (!scopes.includes(scopeNeeded)) {
      return res.status(403).json({ error: 'Forbidden' });
    }
    next();
  };
}

function requireRole(roleNeeded) {
  return (req, res, next) => {
    const role = req.user?.role || 'user';
    if (role !== roleNeeded) {
      return res.status(403).json({ error: 'Forbidden' });
    }
    next();
  };
}

// -------------------- Error handler global --------------------
function errorHandler(err, req, res, next) { // eslint-disable-line
  console.error('[ERR]', err);
  if (res.headersSent) return;
  res.status(500).json({ error: 'internal_error' });
}

module.exports = { z, validate, authBearer, requireScope, requireRole, errorHandler };
