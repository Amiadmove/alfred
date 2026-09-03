const crypto = require('crypto');

function checkAdminAuth(req) {
  const adminPassword = process.env.ADMIN_PASSWORD;
  if (!adminPassword) return true;

  const provided = req.headers['x-admin-password'];
  if (!provided) return false;

  const a = Buffer.alloc(64);
  const b = Buffer.alloc(64);
  Buffer.from(provided).copy(a);
  Buffer.from(adminPassword).copy(b);

  return crypto.timingSafeEqual(a, b) && provided.length === adminPassword.length;
}

function verifySessionToken(token) {
  if (!token || typeof token !== 'string') return null;
  const secret = process.env.OTP_SECRET;
  if (!secret) return null;
  const dot = token.lastIndexOf('.');
  if (dot === -1) return null;
  const payload = token.substring(0, dot);
  const mac = token.substring(dot + 1);
  const expectedMac = crypto.createHmac('sha256', secret).update(payload).digest('hex');
  const macBuf = Buffer.from(mac, 'hex');
  const expectedBuf = Buffer.from(expectedMac, 'hex');
  if (macBuf.length !== expectedBuf.length) return null;
  try {
    if (!crypto.timingSafeEqual(macBuf, expectedBuf)) return null;
  } catch { return null; }
  try {
    const data = JSON.parse(Buffer.from(payload, 'base64url').toString());
    if (data.type !== 'session') return null;
    if (Date.now() > data.exp) return null;
    return data;
  } catch { return null; }
}

module.exports = { checkAdminAuth, verifySessionToken };
