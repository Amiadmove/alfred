const { del } = require('@vercel/blob');

module.exports = async function handler(req, res) {
  if (req.method !== 'DELETE') return res.status(405).end();

  const adminPassword = process.env.ADMIN_PASSWORD;
  if (adminPassword && req.headers['x-admin-password'] !== adminPassword) return res.status(401).end();

  const { url } = req.query;
  if (!url) return res.status(400).json({ error: 'Missing url' });

  try {
    const token = process.env.BLOB_READ_WRITE_TOKEN;
    await del(url, { token });
    res.json({ ok: true });
  } catch (err) {
    console.error(err);
    res.status(500).json({ ok: false, error: err.message });
  }
};
