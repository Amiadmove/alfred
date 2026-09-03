const crypto = require('crypto');

// Stateless OTP: the "challenge" returned to the client is an HMAC-signed token
// that encodes email + OTP + expiry. No server-side state needed.
// Vercel serverless-safe.

const OTP_TTL_MS = 10 * 60 * 1000; // 10 minutes

function getSecret() {
  const s = process.env.OTP_SECRET;
  if (!s) throw new Error('OTP_SECRET env var not set');
  return s;
}

function generateOtp() {
  return String(crypto.randomInt(100000, 999999));
}

// challenge = base64url(payload) . hmac
// payload = JSON { email, otp, exp }
function buildChallenge(email, otp, exp) {
  const payload = Buffer.from(JSON.stringify({ email, otp, exp })).toString('base64url');
  const mac = crypto.createHmac('sha256', getSecret()).update(payload).digest('hex');
  return `${payload}.${mac}`;
}

async function sendOtpEmail(email, otp) {
  if (process.env.RESEND_API_KEY) {
    const res = await fetch('https://api.resend.com/emails', {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${process.env.RESEND_API_KEY}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        from: 'Move Onboarding <onboarding@resend.dev>',
        to: [email],
        subject: 'Your Move Onboarding verification code',
        html: `<!DOCTYPE html><html><head><meta charset="UTF-8"/></head><body style="font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;background:#F7F9FC;padding:40px 16px;margin:0;"><div style="max-width:480px;margin:0 auto;background:white;border-radius:20px;padding:40px;border:1px solid #E2E8F0;"><p style="font-size:15px;font-weight:700;color:#3689FB;margin:0 0 32px;letter-spacing:-0.3px;">Move</p><h1 style="font-size:22px;font-weight:700;color:#152656;margin:0 0 16px;">Your verification code</h1><p style="color:#64748B;font-size:14px;line-height:1.7;margin:0 0 24px;">Enter the code below to access your onboarding questionnaire. This code expires in <strong style="color:#152656;">10 minutes</strong>.</p><div style="background:#F7F9FC;border-radius:16px;padding:28px;text-align:center;margin-bottom:28px;border:1px solid #E2E8F0;"><span style="font-size:40px;font-weight:800;letter-spacing:12px;color:#152656;font-family:monospace;">${otp}</span></div><p style="color:#94A3B8;font-size:12px;line-height:1.6;margin:0;">If you didn't request this code, you can safely ignore this email.<br/><br/><strong style="color:#152656;">Move</strong> · The commerce infrastructure for AI travel</p></div></body></html>`,
      }),
    });
    if (!res.ok) {
      const text = await res.text();
      throw new Error(`Email send failed: ${res.status} ${text}`);
    }
    return;
  }

  if (process.env.SMTP_USER && process.env.SMTP_PASS) {
    const nodemailer = require('nodemailer');
    const transporter = nodemailer.createTransport({
      host: process.env.SMTP_HOST || 'smtp.gmail.com',
      port: Number(process.env.SMTP_PORT) || 587,
      secure: false,
      auth: { user: process.env.SMTP_USER, pass: process.env.SMTP_PASS },
    });
    await transporter.sendMail({
      from: `"Move Onboarding" <${process.env.SMTP_USER}>`,
      to: email,
      subject: 'Your Move Onboarding verification code',
      text: `Your verification code is: ${otp}\n\nThis code expires in 10 minutes.\n\nMove · The commerce infrastructure for AI travel`,
    });
    return;
  }

  throw new Error('No email transport configured. Set RESEND_API_KEY or SMTP_USER/SMTP_PASS.');
}

module.exports = async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({ ok: false, error: 'Method not allowed' });

  const { email } = req.body || {};
  if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    return res.status(400).json({ ok: false, error: 'Valid email address required' });
  }

  const normalizedEmail = email.toLowerCase().trim();
  const otp = generateOtp();
  const exp = Date.now() + OTP_TTL_MS;

  try {
    const challenge = buildChallenge(normalizedEmail, otp, exp);
    await sendOtpEmail(normalizedEmail, otp);
    // Return challenge token to client — they must send it back when verifying
    res.status(200).json({ ok: true, challenge });
  } catch (err) {
    console.error('OTP send error:', err.message);
    res.status(500).json({ ok: false, error: 'Failed to send verification email. Please try again.' });
  }
};

module.exports.buildChallenge = buildChallenge;
module.exports.getSecret = getSecret;
