const crypto = require('crypto');

/**
 * Checks x-admin-password header using timing-safe comparison.
 * Returns true if authorized (or if no ADMIN_PASSWORD is set).
 * Returns false if unauthorized.
 */
function checkAdminAuth(req) {
  const adminPassword = process.env.ADMIN_PASSWORD;
  if (!adminPassword) return true; // no password configured → open (dev mode)

  const provided = req.headers['x-admin-password'];
  if (!provided) return false;

  // Use constant-length buffers for timing-safe comparison
  const a = Buffer.alloc(64);
  const b = Buffer.alloc(64);
  Buffer.from(provided).copy(a);
  Buffer.from(adminPassword).copy(b);

  return crypto.timingSafeEqual(a, b) && provided.length === adminPassword.length;
}

module.exports = { checkAdminAuth };
