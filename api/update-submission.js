const { put } = require('@vercel/blob');

module.exports = async function handler(req, res) {
  if (req.method !== 'PUT') return res.status(405).end();

  const adminPassword = process.env.ADMIN_PASSWORD;
  if (adminPassword) {
    const auth = req.headers['x-admin-password'];
    if (auth !== adminPassword) return res.status(401).end();
  }

  const { id } = req.query;
  if (!id) return res.status(400).json({ error: 'Missing id' });

  try {
    const token = process.env.BLOB_READ_WRITE_TOKEN;
    await put(`submissions/${id}.json`, JSON.stringify(req.body), {
      access: 'private',
      contentType: 'application/json',
      token,
      allowOverwrite: true,
    });
    res.json({ ok: true });
  } catch (err) {
    console.error(err);
    res.status(500).json({ ok: false, error: err.message });
  }
};
