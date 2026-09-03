const crypto = require('crypto');

// Verifies HMAC-signed OTP challenge (stateless, Vercel-safe)
// Then issues a stateless session token (also HMAC-signed)

const SESSION_TTL_MS = 2 * 60 * 60 * 1000; // 2 hours

function getSecret() {
  const s = process.env.OTP_SECRET;
  if (!s) throw new Error('OTP_SECRET env var not set');
  return s;
}

// Verify challenge and return { email, otp, exp } or null
function parseChallenge(challenge) {
  if (!challenge || typeof challenge !== 'string') return null;
  const dot = challenge.lastIndexOf('.');
  if (dot === -1) return null;
  const payload = challenge.substring(0, dot);
  const mac = challenge.substring(dot + 1);

  const expectedMac = crypto.createHmac('sha256', getSecret()).update(payload).digest('hex');
  const macBuf = Buffer.from(mac, 'hex');
  const expectedBuf = Buffer.from(expectedMac, 'hex');
  if (macBuf.length !== expectedBuf.length) return null;
  if (!crypto.timingSafeEqual(macBuf, expectedBuf)) return null;

  try {
    return JSON.parse(Buffer.from(payload, 'base64url').toString());
  } catch {
    return null;
  }
}

// Build a signed session token: base64url(payload).hmac
function buildSessionToken(email, exp) {
  const payload = Buffer.from(JSON.stringify({ email, exp, type: 'session' })).toString('base64url');
  const mac = crypto.createHmac('sha256', getSecret()).update(payload).digest('hex');
  return `${payload}.${mac}`;
}

// Export for use by submit handlers
function verifySessionToken(token) {
  if (!token || typeof token !== 'string') return null;
  const dot = token.lastIndexOf('.');
  if (dot === -1) return null;
  const payload = token.substring(0, dot);
  const mac = token.substring(dot + 1);

  const expectedMac = crypto.createHmac('sha256', getSecret()).update(payload).digest('hex');
  const macBuf = Buffer.from(mac, 'hex');
  const expectedBuf = Buffer.from(expectedMac, 'hex');
  if (macBuf.length !== expectedBuf.length) return null;
  try {
    if (!crypto.timingSafeEqual(macBuf, expectedBuf)) return null;
  } catch {
    return null;
  }

  try {
    const data = JSON.parse(Buffer.from(payload, 'base64url').toString());
    if (data.type !== 'session') return null;
    if (Date.now() > data.exp) return null;
    return data;
  } catch {
    return null;
  }
}

module.exports = async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({ ok: false, error: 'Method not allowed' });

  const { email, otp, challenge } = req.body || {};
  if (!email || !otp || !challenge) {
    return res.status(400).json({ ok: false, error: 'Email, OTP, and challenge required' });
  }

  const normalizedEmail = email.toLowerCase().trim();

  let parsed;
  try {
    parsed = parseChallenge(challenge);
  } catch (err) {
    return res.status(400).json({ ok: false, error: 'Invalid challenge token' });
  }

  if (!parsed) {
    return res.status(400).json({ ok: false, error: 'Invalid or tampered verification token' });
  }

  if (Date.now() > parsed.exp) {
    return res.status(400).json({ ok: false, error: 'Verification code has expired. Please request a new one.' });
  }

  if (parsed.email !== normalizedEmail) {
    return res.status(400).json({ ok: false, error: 'Email mismatch' });
  }

  // Constant-time OTP comparison
  const expected = Buffer.from(String(parsed.otp));
  const provided = Buffer.from(String(otp).trim());
  if (expected.length !== provided.length || !crypto.timingSafeEqual(expected, provided)) {
    return res.status(400).json({ ok: false, error: 'Incorrect code. Please try again.' });
  }

  // Issue session token
  const sessionExp = Date.now() + SESSION_TTL_MS;
  const token = buildSessionToken(normalizedEmail, sessionExp);

  res.status(200).json({ ok: true, token });
};

module.exports.verifySessionToken = verifySessionToken;
